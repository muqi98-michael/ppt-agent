from __future__ import annotations

import base64
import os
import re
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Optional
from urllib.parse import quote
from uuid import uuid4

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import HTMLResponse, Response
from fastapi.templating import Jinja2Templates
from dotenv import load_dotenv, set_key

from app.ppt_merge import merge_with_template
from app.ppt_import import process_ppt_import
from app.ppt_research import append_research_slides
from app.research import (
    ResearchResult,
    SectionResult,
    get_deepseek_logs,
    parse_visit_requirements,
    research_industry_and_customer,
)
from app.storage import (
    delete_chapter,
    delete_job,
    delete_unified_template,
    get_active_unified_template_blob,
    get_active_unified_template_meta,
    get_chapter_file_blob,
    get_db_info,
    get_job_detail,
    get_job_zip_blob,
    get_unified_template_blob_by_id,
    init_db,
    list_unified_templates,
    list_session_records,
    list_import_jobs,
    set_active_unified_template,
    save_unified_template,
    save_session_record,
    save_import_result,
    search_top_chapter_ppts,
)

ROOT_DIR = Path(__file__).resolve().parents[1]
ENV_PATH = ROOT_DIR / ".env"
load_dotenv(dotenv_path=ENV_PATH)

app = FastAPI(title="PPT Merge Agent", version="1.0.0")
templates = Jinja2Templates(directory=str(Path(__file__).parent / "templates"))
init_db()


@app.get("/", response_class=HTMLResponse)
async def home(request: Request):
    return templates.TemplateResponse("home.html", {"request": request, "active": "new_chat"})


@app.get("/merge", response_class=HTMLResponse)
async def merge_page(request: Request):
    return templates.TemplateResponse("merge.html", {"request": request, "active": "merge"})


@app.get("/history-sessions", response_class=HTMLResponse)
async def history_sessions_page(request: Request):
    return templates.TemplateResponse("history_sessions.html", {"request": request, "active": "history_sessions"})


@app.get("/search-fill", response_class=HTMLResponse)
async def search_fill_page(request: Request):
    return templates.TemplateResponse("search_fill.html", {"request": request, "active": "search_fill"})


@app.get("/ppt-import", response_class=HTMLResponse)
async def ppt_import_page(request: Request):
    return templates.TemplateResponse("ppt_import.html", {"request": request, "active": "ppt_import"})


@app.get("/settings", response_class=HTMLResponse)
async def settings_page(request: Request):
    return templates.TemplateResponse("settings.html", {"request": request, "active": "settings"})


@app.get("/practice-library", response_class=HTMLResponse)
async def practice_library_page(request: Request):
    return templates.TemplateResponse("practice_library.html", {"request": request, "active": "practice_library"})


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


@app.get("/api/db-info")
async def api_db_info():
    return get_db_info()


@app.get("/api/history-sessions")
async def api_history_sessions(limit: int = 100):
    return {"items": list_session_records(limit=limit)}


@app.post("/api/history-sessions")
async def api_create_history_session(
    raw_query: str = Form(""),
    generated_prompt: str = Form(...),
    industry: str = Form(...),
    customer: str = Form(...),
    duration: str = Form(...),
    product_name: str = Form(...),
    visit_role: str = Form(...),
    business_domains: str = Form(""),
):
    domains = [x.strip() for x in business_domains.split(",") if x.strip()]
    rec = save_session_record(
        raw_query=raw_query,
        generated_prompt=generated_prompt,
        industry=industry,
        customer=customer,
        duration=duration,
        product_name=product_name,
        visit_role=visit_role,
        business_domains=domains,
    )
    return {"ok": True, "item": rec}


@app.post("/api/new-chat/parse")
async def api_new_chat_parse(
    query: str = Form(...),
    model: str = Form(""),
):
    item = parse_visit_requirements(query=query, model_override=model)
    return {"ok": True, "item": item}


@app.get("/api/practice-library/jobs")
async def api_practice_jobs(limit: int = 100):
    return {"items": list_import_jobs(limit=limit)}


@app.get("/api/practice-library/unified-template")
async def api_get_unified_template():
    item = get_active_unified_template_meta()
    return {"item": item}


@app.post("/api/practice-library/unified-template")
async def api_upload_unified_template(template: UploadFile = File(..., description="Unified PPT template")):
    if not _is_pptx(template.filename):
        raise HTTPException(status_code=400, detail="模板文件必须是 .pptx")
    payload = await template.read()
    if not payload:
        raise HTTPException(status_code=400, detail="模板文件内容为空")
    item = save_unified_template(filename=template.filename or "统一模板.pptx", ppt_bytes=payload)
    return {"ok": True, "item": item}


@app.get("/api/practice-library/unified-templates")
async def api_list_unified_templates(limit: int = 200):
    return {"items": list_unified_templates(limit=limit)}


@app.post("/api/practice-library/unified-template/{template_id}/activate")
async def api_activate_unified_template(template_id: int):
    item = set_active_unified_template(template_id)
    if not item:
        raise HTTPException(status_code=404, detail="模板不存在")
    return {"ok": True, "item": item}


@app.delete("/api/practice-library/unified-template/{template_id}")
async def api_delete_unified_template(template_id: int):
    result = delete_unified_template(template_id)
    if not result:
        raise HTTPException(status_code=404, detail="模板不存在")
    return {"ok": True, **result}


@app.get("/api/practice-library/unified-template/download")
async def api_download_unified_template(template_id: int = 0):
    result = get_unified_template_blob_by_id(template_id) if template_id > 0 else None
    if not result:
        active = get_active_unified_template_blob()
        if active:
            _, filename, payload = active
            result = (filename, payload)
    if not result:
        raise HTTPException(status_code=404, detail="尚未导入统一模板")
    filename, payload = result
    return Response(
        content=payload,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={
            "Content-Disposition": f"attachment; filename=\"unified_template.pptx\"; filename*=UTF-8''{quote(filename)}"
        },
    )


@app.get("/api/practice-library/jobs/{job_id}")
async def api_practice_job_detail(job_id: int):
    data = get_job_detail(job_id)
    if not data:
        raise HTTPException(status_code=404, detail="记录不存在")
    return data


@app.get("/api/practice-library/jobs/{job_id}/zip")
async def api_practice_job_zip(job_id: int):
    result = get_job_zip_blob(job_id)
    if not result:
        raise HTTPException(status_code=404, detail="记录不存在")
    filename, payload = result
    return Response(
        content=payload,
        media_type="application/zip",
        headers={
            "Content-Disposition": f"attachment; filename=\"download.zip\"; filename*=UTF-8''{quote(filename)}"
        },
    )


@app.get("/api/practice-library/chapters/{chapter_id}/download/{file_type}")
async def api_practice_chapter_download(chapter_id: int, file_type: str):
    result = get_chapter_file_blob(chapter_id, file_type)
    if not result:
        raise HTTPException(status_code=404, detail="附件不存在")
    filename, payload, media_type = result
    return Response(
        content=payload,
        media_type=media_type,
        headers={
            "Content-Disposition": f"attachment; filename=\"download\"; filename*=UTF-8''{quote(filename)}"
        },
    )


@app.delete("/api/practice-library/jobs/{job_id}")
async def api_practice_delete_job(job_id: int):
    ok = delete_job(job_id)
    if not ok:
        raise HTTPException(status_code=404, detail="记录不存在")
    return {"ok": True, "deleted_job_id": job_id}


@app.delete("/api/practice-library/chapters/{chapter_id}")
async def api_practice_delete_chapter(chapter_id: int):
    result = delete_chapter(chapter_id)
    if not result:
        raise HTTPException(status_code=404, detail="章节不存在")
    return {"ok": True, **result}


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
    template: Optional[UploadFile] = File(None, description="Template PPTX (optional)"),
    sources: list[UploadFile] = File(..., description="Source PPTX files"),
):
    if not sources:
        raise HTTPException(status_code=400, detail="请至少上传一个内容 PPT 文件")
    if any(not _is_pptx(source.filename) for source in sources):
        raise HTTPException(status_code=400, detail="内容文件必须全部是 .pptx")

    with TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir)
        template_path, template_source = await _resolve_template_path(
            upload_template=template,
            tmp_path=tmp_path,
        )
        source_paths: list[Path] = []

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
        "X-Template-Source": template_source,
    }
    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers=headers,
    )


@app.post("/ppt-import")
async def ppt_import(files: list[UploadFile] = File(..., description="PPTX files")):
    if not files:
        raise HTTPException(status_code=400, detail="请至少上传一个 .pptx 文件")
    if any(not _is_pptx(f.filename) for f in files):
        raise HTTPException(status_code=400, detail="仅支持 .pptx 文件")

    items: list[dict] = []
    for idx, file in enumerate(files, start=1):
        with TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir) / f"upload_{idx}.pptx"
            await _save_upload(file, tmp_path)
            try:
                chapters, zip_bytes = process_ppt_import(tmp_path)
            except Exception as exc:
                raise HTTPException(
                    status_code=400,
                    detail=f"文件《{file.filename or f'第{idx}个文件'}》处理失败: {str(exc)[:200]}",
                ) from exc

        zip_base64 = base64.b64encode(zip_bytes).decode("utf-8") if zip_bytes else ""
        chapters_payload = [
            {
                "title": c.title,
                "content": c.content,
                "summary": c.summary,
                "slide_count": len(c.slide_indices),
                "ppt_base64": c.ppt_base64,
            }
            for c in chapters
        ]
        saved = save_import_result(
            source_filename=file.filename or f"upload_{idx}.pptx",
            chapters=chapters_payload,
            zip_bytes=zip_bytes,
        )
        db_info = get_db_info()
        db_info.update({"saved_job_id": saved["job_id"], "saved_chapters": saved["chapter_count"]})
        items.append(
            {
                "source_filename": file.filename or f"upload_{idx}.pptx",
                "chapters": chapters_payload,
                "zip_base64": zip_base64,
                "db_info": db_info,
            }
        )

    # 兼容旧前端：当仅单文件时保留原字段
    if len(items) == 1:
        one = items[0]
        return {"items": items, "chapters": one["chapters"], "zip_base64": one["zip_base64"], "db_info": one["db_info"]}
    return {"items": items}


@app.post("/search-fill")
async def search_fill(
    industry: str = Form(...),
    customer: str = Form(...),
    model: str = Form(""),
    template: Optional[UploadFile] = File(None, description="Template PPTX (optional)"),
):
    if not industry.strip():
        raise HTTPException(status_code=400, detail="请输入行业")
    if not customer.strip():
        raise HTTPException(status_code=400, detail="请输入客户")

    with TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir)
        template_path, template_source = await _resolve_template_path(
            upload_template=template,
            tmp_path=tmp_path,
        )

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
        "Content-Disposition": f"attachment; filename=\"search_fill_output.pptx\"; filename*=UTF-8''{quote(safe_name)}",
        "X-Template-Source": template_source,
    }
    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers=headers,
    )


@app.post("/api/generate-visit-ppt")
async def api_generate_visit_ppt(
    industry: str = Form(...),
    customer: str = Form(...),
    product_name: str = Form(...),
    business_domains: str = Form(""),
    visit_role: str = Form(...),
    model: str = Form(""),
    match_ids: str = Form(""),
):
    if not industry.strip():
        raise HTTPException(status_code=400, detail="请输入行业")
    if not customer.strip():
        raise HTTPException(status_code=400, detail="请输入客户")
    if not product_name.strip():
        raise HTTPException(status_code=400, detail="请输入产品")
    if not visit_role.strip():
        raise HTTPException(status_code=400, detail="请输入拜访角色")

    domains = [x.strip() for x in business_domains.split(",") if x.strip()]
    selected_ids: list[int] = []
    if match_ids.strip():
        for raw in match_ids.split(","):
            raw = raw.strip()
            if not raw:
                continue
            try:
                selected_ids.append(int(raw))
            except ValueError:
                continue

    with TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir)
        base_template_path, template_source = await _resolve_template_path(upload_template=None, tmp_path=tmp_path)

        try:
            research = research_industry_and_customer(
                industry.strip(),
                customer.strip(),
                model_override=model,
            )
        except ValueError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        research = _compact_research_result(research)

        researched_path = tmp_path / "visit_research_base.pptx"
        append_research_slides(base_template_path, researched_path, research)

        matches = _resolve_visit_matches(
            selected_ids=selected_ids,
            product_name=product_name.strip(),
            business_domains=domains,
            visit_role=visit_role.strip(),
        )
        source_paths: list[Path] = []
        for idx, item in enumerate(matches, start=1):
            source_path = tmp_path / f"matched_{idx}.pptx"
            source_path.write_bytes(item["ppt_blob"])
            source_paths.append(source_path)

        output_path = tmp_path / f"visit_ppt_{uuid4().hex[:8]}.pptx"
        report = merge_with_template(researched_path, source_paths, output_path)
        payload = output_path.read_bytes()

    safe_name = _safe_filename(f"拜访方案_{industry}_{customer}_{product_name}.pptx")
    return Response(
        content=payload,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={
            "Content-Disposition": f"attachment; filename=\"visit_plan.pptx\"; filename*=UTF-8''{quote(safe_name)}",
            "X-Template-Source": template_source,
            "X-Matched-PPT-Count": str(len(matches)),
            "X-Merge-Imported-Slides": str(report.imported_slides),
        },
    )


@app.post("/api/generate-visit-ppt/preview")
async def api_generate_visit_ppt_preview(
    product_name: str = Form(...),
    business_domains: str = Form(""),
    visit_role: str = Form(...),
):
    domains = [x.strip() for x in business_domains.split(",") if x.strip()]
    matches = search_top_chapter_ppts(
        product_name=product_name.strip(),
        business_domains=domains,
        visit_role=visit_role.strip(),
        limit=3,
    )
    items = [
        {
            "chapter_id": int(item["chapter_id"]),
            "title": item["title"],
            "source_filename": item["source_filename"],
            "score": int(item["score"]),
        }
        for item in matches
    ]
    return {"ok": True, "items": items}


def _is_pptx(filename: str | None) -> bool:
    return bool(filename and filename.lower().endswith(".pptx"))


async def _save_upload(upload_file: UploadFile, destination: Path) -> None:
    payload = await upload_file.read()
    destination.write_bytes(payload)


async def _resolve_template_path(*, upload_template: Optional[UploadFile], tmp_path: Path) -> tuple[Path, str]:
    """
    模板优先级：1) 本次上传模板 2) 统一模板库中的最新模板。
    """
    if upload_template and upload_template.filename:
        if not _is_pptx(upload_template.filename):
            raise HTTPException(status_code=400, detail="模板文件必须是 .pptx")
        upload_path = tmp_path / "template_upload.pptx"
        await _save_upload(upload_template, upload_path)
        if upload_path.stat().st_size <= 0:
            raise HTTPException(status_code=400, detail="上传模板内容为空")
        return upload_path, "upload"

    active = get_active_unified_template_blob()
    if not active:
        raise HTTPException(status_code=400, detail="未上传临时模板，且未配置统一模板。请先到方案库管理导入统一模板。")
    _, _, payload = active
    if not payload:
        raise HTTPException(status_code=400, detail="统一模板内容为空，请重新导入统一模板。")
    db_path = tmp_path / "template_unified.pptx"
    db_path.write_bytes(payload)
    return db_path, "unified_db"


def _mask_key(key: str) -> str:
    if not key:
        return ""
    if len(key) <= 8:
        return "*" * len(key)
    return f"{key[:4]}{'*' * (len(key) - 8)}{key[-4:]}"


def _compact_research_result(research: ResearchResult) -> ResearchResult:
    def _compact_sections(sections: list[SectionResult]) -> list[SectionResult]:
        out: list[SectionResult] = []
        for sec in sections[:3]:
            bullets = []
            for bullet in sec.bullets[:3]:
                text = " ".join(str(bullet).split())
                bullets.append(text[:90])
            out.append(SectionResult(title=str(sec.title)[:40], bullets=bullets, sources=list(sec.sources[:2])))
        return out

    return ResearchResult(
        industry=research.industry,
        customer=research.customer,
        industry_sections=_compact_sections(research.industry_sections),
        customer_sections=_compact_sections(research.customer_sections),
    )


def _safe_filename(name: str) -> str:
    cleaned = re.sub(r'[\\/:*?"<>|]+', "_", name or "").strip()
    return cleaned or "visit_plan.pptx"


def _resolve_visit_matches(
    *,
    selected_ids: list[int],
    product_name: str,
    business_domains: list[str],
    visit_role: str,
) -> list[dict[str, object]]:
    if selected_ids:
        chosen: list[dict[str, object]] = []
        for cid in selected_ids[:3]:
            file_row = get_chapter_file_blob(cid, "ppt")
            if not file_row:
                continue
            filename, payload, _ = file_row
            chosen.append(
                {
                    "chapter_id": cid,
                    "title": filename,
                    "ppt_filename": filename,
                    "ppt_blob": payload,
                    "source_filename": "",
                    "score": 0,
                }
            )
        if chosen:
            return chosen
    return search_top_chapter_ppts(
        product_name=product_name,
        business_domains=business_domains,
        visit_role=visit_role,
        limit=3,
    )
