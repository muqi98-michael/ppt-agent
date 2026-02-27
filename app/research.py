from __future__ import annotations

import json
import os
import time
from collections import deque
from dataclasses import dataclass
from datetime import datetime
from typing import Any, Iterable

import httpx


@dataclass
class SectionResult:
    title: str
    bullets: list[str]
    sources: list[str]


@dataclass
class ResearchResult:
    industry: str
    customer: str
    industry_sections: list[SectionResult]
    customer_sections: list[SectionResult]


_DEEPSEEK_LOGS: deque[dict[str, Any]] = deque(maxlen=100)


def research_industry_and_customer(
    industry: str,
    customer: str,
    model_override: str | None = None,
) -> ResearchResult:
    started_at = time.perf_counter()
    model_used = model_override.strip() if model_override and model_override.strip() else os.getenv("DEEPSEEK_MODEL", "deepseek-chat")
    try:
        payload, model_used = _call_deepseek(industry=industry, customer=customer, model_override=model_override)
    except Exception as exc:
        _append_log(
            industry=industry,
            customer=customer,
            model=model_used,
            status="error",
            duration_ms=int((time.perf_counter() - started_at) * 1000),
            summary="调用失败",
            error=str(exc),
        )
        raise

    industry_sections = _parse_sections(payload.get("industry_sections"), fallback_prefix=industry)
    customer_sections = _parse_sections(payload.get("customer_sections"), fallback_prefix=customer)

    result = ResearchResult(
        industry=industry,
        customer=customer,
        industry_sections=industry_sections,
        customer_sections=customer_sections,
    )
    _append_log(
        industry=industry,
        customer=customer,
        model=model_used,
        status="success",
        duration_ms=int((time.perf_counter() - started_at) * 1000),
        summary=f"行业:{industry_sections[0].title} | 客户:{customer_sections[0].title}",
        error="",
    )
    return result


def _call_deepseek(
    industry: str,
    customer: str,
    model_override: str | None = None,
) -> tuple[dict[str, Any], str]:
    api_key = os.getenv("DEEPSEEK_API_KEY", "").strip()
    if not api_key:
        raise ValueError("未配置 DEEPSEEK_API_KEY，请先在环境变量或 .env 文件中设置。")

    base_url = os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com")
    model = model_override.strip() if model_override and model_override.strip() else os.getenv("DEEPSEEK_MODEL", "deepseek-chat")
    timeout_s = float(os.getenv("DEEPSEEK_TIMEOUT_SECONDS", "90"))

    system_prompt = (
        "你是一名企业数字化咨询顾问。请先进行联网检索，再返回结构化结果。"
        "输出必须是合法 JSON，不要输出 markdown 代码块。"
    )
    user_prompt = f"""
请围绕以下信息进行公开网络检索，并输出简明结果：
- 行业：{industry}
- 客户：{customer}

请按以下 JSON 结构返回：
{{
  "industry_sections": [
    {{"title": "{industry} 最新趋势", "bullets": ["...","...","..."], "sources": ["https://...","https://..."]}},
    {{"title": "{industry} 行业痛点", "bullets": ["...","...","..."], "sources": ["https://..."]}},
    {{"title": "{industry} IT规划", "bullets": ["...","...","..."], "sources": ["https://..."]}}
  ],
  "customer_sections": [
    {{"title": "{customer} 官网与新闻", "bullets": ["...","...","..."], "sources": ["https://..."]}},
    {{"title": "{customer} 业务规划与战略", "bullets": ["...","...","..."], "sources": ["https://..."]}},
    {{"title": "{customer} 痛点与IT需求", "bullets": ["...","...","..."], "sources": ["https://..."]}}
  ]
}}

要求：
1) 每个 bullets 3-5 条，单条不超过 150 字，保持完整句子。
2) 每个 section 给 1-3 个 sources，优先官网和权威媒体。
3) 内容用中文，避免空话，尽量具体。
"""

    req_body = {
        "model": model,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        "temperature": 0.3,
    }

    with httpx.Client(timeout=timeout_s) as client:
        resp = client.post(
            f"{base_url.rstrip('/')}/chat/completions",
            headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
            json=req_body,
        )
        resp.raise_for_status()
        data = resp.json()

    content = data["choices"][0]["message"]["content"]
    return _extract_json_object(content), model


def _parse_sections(raw_sections: Any, fallback_prefix: str) -> list[SectionResult]:
    if not isinstance(raw_sections, list):
        return _default_sections(fallback_prefix)

    parsed: list[SectionResult] = []
    for item in raw_sections[:3]:
        if not isinstance(item, dict):
            continue
        title = _normalize_text(str(item.get("title") or "").strip()) or f"{fallback_prefix} 信息分析"
        bullets_raw = item.get("bullets")
        sources_raw = item.get("sources")
        bullets = _to_clean_list(bullets_raw, max_items=5, max_chars=200)
        sources = _to_clean_urls(sources_raw, max_items=3)
        if not bullets:
            bullets = ["暂无充分公开数据，建议补充企业材料后再次生成。"]
        parsed.append(SectionResult(title=title, bullets=bullets, sources=sources))

    if len(parsed) < 3:
        defaults = _default_sections(fallback_prefix)
        parsed.extend(defaults[len(parsed) :])
    return parsed[:3]


def _extract_json_object(content: str) -> dict[str, Any]:
    text = content.strip()
    if text.startswith("```"):
        text = text.strip("`")
        text = text.replace("json\n", "", 1).strip()
    try:
        data = json.loads(text)
        if isinstance(data, dict):
            return data
    except json.JSONDecodeError:
        pass

    start = text.find("{")
    end = text.rfind("}")
    if start != -1 and end != -1 and end > start:
        snippet = text[start : end + 1]
        data = json.loads(snippet)
        if isinstance(data, dict):
            return data

    raise ValueError("DeepSeek 返回结果不是合法 JSON，请稍后重试。")


def _normalize_text(text: str) -> str:
    return " ".join(text.replace("\n", " ").split())


def _to_clean_list(value: Any, max_items: int, max_chars: int) -> list[str]:
    if not isinstance(value, list):
        return []
    out: list[str] = []
    for item in value:
        text = _normalize_text(str(item))
        if not text:
            continue
        if len(text) > max_chars:
            text = text[:max_chars] + "..."
        out.append(text)
        if len(out) >= max_items:
            break
    return out


def _to_clean_urls(value: Any, max_items: int) -> list[str]:
    if not isinstance(value, list):
        return []
    cleaned = [str(x).strip() for x in value if str(x).strip().startswith("http")]
    return _dedupe_keep_order(cleaned)[:max_items]


def _dedupe_keep_order(items: Iterable[str]) -> list[str]:
    seen = set()
    out: list[str] = []
    for item in items:
        if item in seen:
            continue
        seen.add(item)
        out.append(item)
    return out


def _default_sections(prefix: str) -> list[SectionResult]:
    return [
        SectionResult(
            title=f"{prefix} 最新趋势",
            bullets=["暂无充分公开数据，建议补充企业材料后再次生成。"],
            sources=[],
        ),
        SectionResult(
            title=f"{prefix} 核心痛点",
            bullets=["暂无充分公开数据，建议补充企业材料后再次生成。"],
            sources=[],
        ),
        SectionResult(
            title=f"{prefix} IT规划/需求",
            bullets=["暂无充分公开数据，建议补充企业材料后再次生成。"],
            sources=[],
        ),
    ]


def _append_log(
    *,
    industry: str,
    customer: str,
    model: str,
    status: str,
    duration_ms: int,
    summary: str,
    error: str,
) -> None:
    _DEEPSEEK_LOGS.appendleft(
        {
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "industry": industry,
            "customer": customer,
            "model": model,
            "status": status,
            "duration_ms": duration_ms,
            "summary": summary,
            "error": error[:200],
        }
    )


def get_deepseek_logs(limit: int = 20) -> list[dict[str, Any]]:
    size = max(1, min(limit, 100))
    return list(_DEEPSEEK_LOGS)[:size]


def summarize_chapter_contents(
    chapters: list[dict[str, str]],
    model_override: str | None = None,
) -> list[str]:
    """调用 DeepSeek 对章节内容做通顺、简约的提炼总结。

    输入: [{"title": "...", "content": "..."}, ...]
    输出: 与输入同长度的摘要列表
    """
    if not chapters:
        return []

    api_key = os.getenv("DEEPSEEK_API_KEY", "").strip()
    if not api_key:
        return [_fallback_summary(item.get("content", "")) for item in chapters]

    base_url = os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com")
    model = model_override.strip() if model_override and model_override.strip() else os.getenv("DEEPSEEK_MODEL", "deepseek-chat")
    timeout_s = float(os.getenv("DEEPSEEK_TIMEOUT_SECONDS", "90"))

    compact_items = []
    for idx, item in enumerate(chapters, start=1):
        compact_items.append(
            {
                "index": idx,
                "title": str(item.get("title", "")).strip()[:120],
                "content": str(item.get("content", "")).strip()[:6000],
            }
        )

    system_prompt = (
        "你是企业级文档编审助手。请对章节内容做通顺、简约、无空话的中文提炼。"
        "输出必须是合法 JSON，不要输出 markdown 代码块。"
    )
    user_prompt = f"""
请根据以下章节内容，生成每章摘要。

输入 JSON:
{json.dumps(compact_items, ensure_ascii=False)}

请输出 JSON:
{{
  "summaries": [
    {{"index": 1, "summary": "..." }}
  ]
}}

要求：
1) 每章摘要 2-4 句，总长度 60-150 字；
2) 语言自然、通顺、简约，保留关键信息，不要口号式空话；
3) 严禁编造输入中没有的事实；
4) summaries 顺序与输入一致。
"""

    req_body = {
        "model": model,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        "temperature": 0.2,
    }

    try:
        with httpx.Client(timeout=timeout_s) as client:
            resp = client.post(
                f"{base_url.rstrip('/')}/chat/completions",
                headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
                json=req_body,
            )
            resp.raise_for_status()
            data = resp.json()
        content = data["choices"][0]["message"]["content"]
        payload = _extract_json_object(content)
        summaries_raw = payload.get("summaries")
        if not isinstance(summaries_raw, list):
            raise ValueError("summaries 字段缺失")

        summaries: list[str] = []
        for i, item in enumerate(summaries_raw[: len(chapters)]):
            if not isinstance(item, dict):
                summaries.append(_fallback_summary(chapters[i].get("content", "")))
                continue
            text = _normalize_text(str(item.get("summary") or ""))
            if not text:
                text = _fallback_summary(chapters[i].get("content", ""))
            summaries.append(text[:220])

        while len(summaries) < len(chapters):
            summaries.append(_fallback_summary(chapters[len(summaries)].get("content", "")))
        return summaries
    except Exception:
        return [_fallback_summary(item.get("content", "")) for item in chapters]


def _fallback_summary(content: str) -> str:
    text = _normalize_text(content or "")
    if not text:
        return "本章内容较少，建议结合原始页面进行补充。"
    if len(text) <= 140:
        return text
    return text[:140] + "..."


def parse_visit_requirements(query: str, model_override: str | None = None) -> dict[str, Any]:
    text = _normalize_text(query or "")
    if not text:
        return {
            "industry": "",
            "customer": "",
            "duration": "",
            "product_name": "",
            "visit_role": "",
            "business_domains": [],
        }

    industries = ["装备制造", "电子高科技", "汽车零部件", "生命科学", "食品消费", "流程制造", "现代服务", "日化日用品", "现代农牧业", "餐饮行业"]
    durations = ["15分钟", "30分钟"]
    products = ["金蝶AI星空", "金蝶AI星瀚", "金蝶AI HR"]
    roles = ["老板", "供应链负责人", "财务负责人", "生产制造负责人", "IT负责人"]
    domains = ["财务管理", "供应链管理", "采购管理", "服务管理", "研发管理", "生产管理", "资产管理", "人力资源管理"]

    api_key = os.getenv("DEEPSEEK_API_KEY", "").strip()
    if api_key:
        try:
            base_url = os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com")
            model = model_override.strip() if model_override and model_override.strip() else os.getenv("DEEPSEEK_MODEL", "deepseek-chat")
            timeout_s = float(os.getenv("DEEPSEEK_TIMEOUT_SECONDS", "90"))
            system_prompt = "你是企业拜访PPT需求解析助手。只输出 JSON，不要输出 markdown。"
            user_prompt = f"""
请从用户输入中抽取字段，返回 JSON：
{{
  "industry": "",
  "customer": "",
  "duration": "",
  "product_name": "",
  "visit_role": "",
  "business_domains": []
}}

限制：
- customer 仅保留企业名，不得重复词组，不要包含“行业/分钟/产品/角色”等无关词。
- business_domains 仅从候选中选择，可多选。

候选：
industry: {industries}
duration: {durations}
product_name: {products}
visit_role: {roles}
business_domains: {domains}

输入：{text}
"""
            req_body = {
                "model": model,
                "messages": [
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt},
                ],
                "temperature": 0.1,
            }
            with httpx.Client(timeout=timeout_s) as client:
                resp = client.post(
                    f"{base_url.rstrip('/')}/chat/completions",
                    headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
                    json=req_body,
                )
                resp.raise_for_status()
                data = resp.json()
            content = data["choices"][0]["message"]["content"]
            parsed = _extract_json_object(content)
            out = {
                "industry": _pick_option(str(parsed.get("industry") or ""), industries, {"电子高科": "电子高科技"}),
                "customer": _sanitize_customer(str(parsed.get("customer") or "")),
                "duration": _pick_option(str(parsed.get("duration") or ""), durations, {}),
                "product_name": _pick_option(str(parsed.get("product_name") or ""), products, {}),
                "visit_role": _pick_option(
                    str(parsed.get("visit_role") or ""),
                    roles,
                    {"CFO": "财务负责人", "cfo": "财务负责人", "CEO": "老板", "ceo": "老板", "CIO": "IT负责人", "cio": "IT负责人"},
                ),
                "business_domains": _pick_multi_options(parsed.get("business_domains"), domains),
            }
            if out["customer"]:
                return out
        except Exception:
            pass

    return _fallback_parse_visit_requirements(text)


def _fallback_parse_visit_requirements(text: str) -> dict[str, Any]:
    industries = ["装备制造", "电子高科技", "汽车零部件", "生命科学", "食品消费", "流程制造", "现代服务", "日化日用品", "现代农牧业", "餐饮行业"]
    products = ["金蝶AI星空", "金蝶AI星瀚", "金蝶AI HR"]
    roles = ["老板", "供应链负责人", "财务负责人", "生产制造负责人", "IT负责人"]
    domains = ["财务管理", "供应链管理", "采购管理", "服务管理", "研发管理", "生产管理", "资产管理", "人力资源管理"]
    role_alias = {"CFO": "财务负责人", "cfo": "财务负责人", "CEO": "老板", "ceo": "老板", "CIO": "IT负责人", "cio": "IT负责人"}

    industry = _pick_option(text, industries, {"电子高科": "电子高科技"})
    duration = "30分钟" if "30分钟" in text else "15分钟" if "15分钟" in text else ""
    product_name = _pick_option(text, products, {})
    visit_role = _pick_option(text, roles, role_alias)
    business_domains = [d for d in domains if d in text]

    customer = ""
    m = __import__("re").search(r"拜访\s*([^，。,.\s]{2,40}?)(?:企业|公司)", text)
    if m and m.group(1):
        customer = m.group(1)
    if not customer:
        m2 = __import__("re").search(r"客户(?:是|为)\s*([^，。,.\s]{2,40})", text)
        if m2 and m2.group(1):
            cand = m2.group(1)
            if cand not in roles and cand not in role_alias:
                customer = cand
    customer = _sanitize_customer(customer)
    return {
        "industry": industry,
        "customer": customer,
        "duration": duration,
        "product_name": product_name,
        "visit_role": visit_role,
        "business_domains": business_domains,
    }


def _pick_option(text: str, options: list[str], alias: dict[str, str]) -> str:
    for key, val in alias.items():
        if key and key in text:
            return val
    for op in options:
        if op in text:
            return op
    return ""


def _pick_multi_options(value: Any, options: list[str]) -> list[str]:
    if isinstance(value, list):
        picked = [str(x).strip() for x in value if str(x).strip() in options]
        return _dedupe_keep_order(picked)
    text = str(value or "")
    picked = [op for op in options if op in text]
    return _dedupe_keep_order(picked)


def _sanitize_customer(raw: str) -> str:
    text = _normalize_text(raw or "")
    if not text:
        return ""
    text = text.replace("企业", "").replace("公司", "")
    text = __import__("re").sub(r"(行业|分钟|产品|角色|业务域).*$", "", text).strip()
    # 抑制重复短语（例如连续重复“电子高科技行业的”）
    for n in range(2, 9):
        if len(text) < n * 3:
            continue
        unit = text[:n]
        while text.startswith(unit * 2):
            text = text[n:]
    return text[:40]
