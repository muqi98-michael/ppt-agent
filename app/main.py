from __future__ import annotations

import os
from pathlib import Path
from tempfile import TemporaryDirectory
from urllib.parse import quote
from uuid import uuid4

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import HTMLResponse, Response
from fastapi.templating import Jinja2Templates
from dotenv import load_dotenv, set_key

from app.ppt_merge import merge_with_template
from app.ppt_research import append_research_slides
from app.research import get_deepseek_logs, research_industry_and_customer

ROOT_DIR = Path(__file__).resolve().parents[1]
ENV_PATH = ROOT_DIR / ".env"
load_dotenv(dotenv_path=ENV_PATH)

app = FastAPI(title="PPT Merge Agent", version="1.0.0")
templates = Jinja2Templates(directory=str(Path(__file__).parent / "templates"))


@app.get("/", response_class=HTMLResponse)
async def home(request: Request):
    return templates.TemplateResponse("home.html", {"request": request, "active": "new_chat"})


@app.get("/merge", response_class=HTMLResponse)
async def merge_page(request: Request):
    return templates.TemplateResponse("merge.html", {"request": request, "active": "merge"})


@app.get("/search-fill", response_class=HTMLResponse)
async def search_fill_page(request: Request):
    return templates.TemplateResponse("search_fill.html", {"request": request, "active": "search_fill"})


@app.get("/settings", response_class=HTMLResponse)
async def settings_page(request: Request):
    return templates.TemplateResponse("settings.html", {"request": request, "active": "settings"})


@app.get("/api/deepseek-logs")
async def api_deepseek_logs(limit: int = 20):
    return {"items": get_deepseek_logs(limit=limit)}


@app.get("/api/settings")
async def api_get_settings():
    key = os.getenv("DEEPSEEK_API_KEY", "").strip()
    return {
        "has_api_key": bool(key),
        "masked_api_key": _mask_key(key),
        "deepseek_base_url": os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com"),
        "deepseek_model": os.getenv("DEEPSEEK_MODEL", "deepseek-chat"),
        "deepseek_timeout_seconds": os.getenv("DEEPSEEK_TIMEOUT_SECONDS", "90"),
    }


@app.post("/api/settings")
async def api_save_settings(
    deepseek_api_key: str = Form(""),
    deepseek_base_url: str = Form("https://api.deepseek.com"),
    deepseek_model: str = Form("deepseek-chat"),
    deepseek_timeout_seconds: str = Form("90"),
):
    try:
        timeout_value = float(deepseek_timeout_seconds)
        if timeout_value <= 0:
            raise ValueError()
    except ValueError as exc:
        raise HTTPException(status_code=400, detail="超时时间必须是大于 0 的数字") from exc

    if not ENV_PATH.exists():
        ENV_PATH.write_text("", encoding="utf-8")

    set_key(str(ENV_PATH), "DEEPSEEK_BASE_URL", deepseek_base_url.strip() or "https://api.deepseek.com")
    set_key(str(ENV_PATH), "DEEPSEEK_MODEL", deepseek_model.strip() or "deepseek-chat")
    set_key(str(ENV_PATH), "DEEPSEEK_TIMEOUT_SECONDS", str(timeout_value))
    if deepseek_api_key.strip():
        set_key(str(ENV_PATH), "DEEPSEEK_API_KEY", deepseek_api_key.strip())

    load_dotenv(dotenv_path=ENV_PATH, override=True)
    key = os.getenv("DEEPSEEK_API_KEY", "").strip()
    return {"ok": True, "has_api_key": bool(key), "masked_api_key": _mask_key(key)}


@app.post("/api/settings/test")
async def api_test_settings_connection(
    deepseek_base_url: str = Form(""),
    deepseek_model: str = Form(""),
    deepseek_timeout_seconds: str = Form("90"),
):
    import httpx

    api_key = os.getenv("DEEPSEEK_API_KEY", "").strip()
    if not api_key:
        raise HTTPException(status_code=400, detail="未配置 DEEPSEEK_API_KEY，请先保存 API Key。")

    base_url = (deepseek_base_url or os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com")).strip()
    model = (deepseek_model or os.getenv("DEEPSEEK_MODEL", "deepseek-chat")).strip()
    try:
        timeout_s = float(deepseek_timeout_seconds or os.getenv("DEEPSEEK_TIMEOUT_SECONDS", "90"))
        if timeout_s <= 0:
            raise ValueError()
    except ValueError as exc:
        raise HTTPException(status_code=400, detail="超时时间必须是大于 0 的数字") from exc

    body = {
        "model": model,
        "messages": [
            {"role": "system", "content": "You are a concise assistant."},
            {"role": "user", "content": "请回复: ok"},
        ],
        "temperature": 0,
        "max_tokens": 16,
    }
    try:
        with httpx.Client(timeout=timeout_s) as client:
            resp = client.post(
                f"{base_url.rstrip('/')}/chat/completions",
                headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
                json=body,
            )
            resp.raise_for_status()
            data = resp.json()
    except httpx.HTTPStatusError as exc:
        status_code = exc.response.status_code if exc.response else 0
        raw_text = (exc.response.text or "").strip()[:400] if exc.response else str(exc)
        detail_parts = [f"HTTP {status_code}"]
        try:
            if exc.response and "json" in (exc.response.headers.get("content-type") or ""):
                err_json = exc.response.json()
                err_obj = err_json.get("error") or err_json
                if isinstance(err_obj, dict):
                    code = err_obj.get("code") or (err_obj.get("error") or {}).get("code")
                    msg = err_obj.get("message") or (err_obj.get("error") or {}).get("message")
                    if code:
                        detail_parts.append(f"code: {code}")
                    if msg:
                        detail_parts.append(f"message: {msg}")
                elif isinstance(err_json.get("error"), str):
                    detail_parts.append(f"error: {err_json['error']}")
        except Exception:
            pass
        if raw_text:
            detail_parts.append(f"body: {raw_text}")
        raise HTTPException(status_code=400, detail=" | ".join(detail_parts)) from exc
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"连接失败：{str(exc)[:300]}") from exc

    answer = ""
    try:
        answer = data["choices"][0]["message"]["content"]
    except Exception:
        answer = ""
    return {"ok": True, "message": "连接成功", "model": model, "preview": (answer or "").strip()[:80]}


@app.post("/merge")
async def merge_ppt(
    template: UploadFile = File(..., description="Template PPTX"),
    sources: list[UploadFile] = File(..., description="Source PPTX files"),
):
    if not _is_pptx(template.filename):
        raise HTTPException(status_code=400, detail="模板文件必须是 .pptx")
    if not sources:
        raise HTTPException(status_code=400, detail="请至少上传一个内容 PPT 文件")
    if any(not _is_pptx(source.filename) for source in sources):
        raise HTTPException(status_code=400, detail="内容文件必须全部是 .pptx")

    with TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir)
        template_path = tmp_path / "template.pptx"
        source_paths: list[Path] = []

        await _save_upload(template, template_path)

        for index, source in enumerate(sources, start=1):
            source_path = tmp_path / f"source_{index}.pptx"
            await _save_upload(source, source_path)
            source_paths.append(source_path)

        output_path = tmp_path / f"merged_{uuid4().hex[:8]}.pptx"
        report = merge_with_template(template_path, source_paths, output_path)
        data = output_path.read_bytes()

    headers = {
        "Content-Disposition": f'attachment; filename="merged_{report.imported_slides}_slides.pptx"',
        "X-Merge-Imported-Slides": str(report.imported_slides),
        "X-Merge-Adjusted-Slides": str(report.layout_adjusted_slides),
    }
    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers=headers,
    )


@app.post("/search-fill")
async def search_fill(
    industry: str = Form(...),
    customer: str = Form(...),
    model: str = Form(""),
    template: UploadFile = File(..., description="Template PPTX"),
):
    if not industry.strip():
        raise HTTPException(status_code=400, detail="请输入行业")
    if not customer.strip():
        raise HTTPException(status_code=400, detail="请输入客户")
    if not _is_pptx(template.filename):
        raise HTTPException(status_code=400, detail="模板文件必须是 .pptx")

    with TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir)
        template_path = tmp_path / "template.pptx"
        await _save_upload(template, template_path)

        try:
            research = research_industry_and_customer(
                industry.strip(),
                customer.strip(),
                model_override=model,
            )
        except ValueError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        output_path = tmp_path / f"search_fill_{uuid4().hex[:8]}.pptx"
        append_research_slides(template_path, output_path, research)
        data = output_path.read_bytes()

    safe_name = f"search_fill_{industry.strip()}_{customer.strip()}.pptx"
    headers = {
        "Content-Disposition": f"attachment; filename=\"search_fill_output.pptx\"; filename*=UTF-8''{quote(safe_name)}"
    }
    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers=headers,
    )


def _is_pptx(filename: str | None) -> bool:
    return bool(filename and filename.lower().endswith(".pptx"))


async def _save_upload(upload_file: UploadFile, destination: Path) -> None:
    payload = await upload_file.read()
    destination.write_bytes(payload)


def _mask_key(key: str) -> str:
    if not key:
        return ""
    if len(key) <= 8:
        return "*" * len(key)
    return f"{key[:4]}{'*' * (len(key) - 8)}{key[-4:]}"
