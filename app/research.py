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
