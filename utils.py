from __future__ import annotations

import html
import json
import re
from pathlib import Path
from typing import Any

CONFIG_PATH = Path.home() / ".docling_review_config.json"


def canonical_doc_key(file_name: str) -> str:
    """Build a normalized grouping key so related files are matched together."""
    name = Path(file_name).name
    lowered = name.lower()
    stripped = name
    known_suffixes = (
        ".docling.html",
        ".docling.htm",
        ".docling.md",
        ".docling.markdown",
        ".docling.json",
        ".html",
        ".htm",
        ".md",
        ".markdown",
        ".json",
        ".pdf",
        ".pptx",
    )
    for suffix in known_suffixes:
        if lowered.endswith(suffix):
            stripped = name[: -len(suffix)]
            break
    base = Path(stripped).stem
    base = re.sub(r"(?i)([_\-. ]?(docling|output|export|parsed|result))$", "", base).strip(" _-.")
    return (base or Path(name).stem).lower()


def short_snippet(text: str, max_len: int = 160) -> str:
    compact = " ".join(text.split())
    if len(compact) <= max_len:
        return compact
    return compact[: max_len - 1].rstrip() + "..."


def load_text_from_bytes(data: bytes) -> str:
    for enc in ("utf-8", "utf-8-sig", "cp1250", "latin-1"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace")


def load_config() -> dict[str, Any]:
    if not CONFIG_PATH.exists():
        return {}
    try:
        return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}


def save_config(config: dict[str, Any]) -> None:
    try:
        CONFIG_PATH.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")
    except OSError:
        pass


def load_recent_folders(limit: int = 8) -> list[str]:
    config = load_config()
    recent = config.get("recent_folders", [])
    if not isinstance(recent, list):
        return []
    result: list[str] = []
    for item in recent:
        if isinstance(item, str) and item and item not in result:
            result.append(item)
    return result[:limit]


def push_recent_folder(path: str, limit: int = 8) -> list[str]:
    path = path.strip()
    if not path:
        return load_recent_folders(limit=limit)
    recent = [p for p in load_recent_folders(limit=limit * 2) if p != path]
    updated = [path, *recent][:limit]
    config = load_config()
    config["recent_folders"] = updated
    save_config(config)
    return updated


def highlight_context(haystack: str, needle: str, context_chars: int = 120) -> str:
    """Return a safe HTML snippet with one highlighted match, or empty string."""
    if not haystack or not needle:
        return ""
    index = haystack.lower().find(needle.lower())
    if index < 0:
        return ""
    start = max(0, index - context_chars)
    end = min(len(haystack), index + len(needle) + context_chars)
    before = html.escape(haystack[start:index])
    hit = html.escape(haystack[index : index + len(needle)])
    after = html.escape(haystack[index + len(needle) : end])
    prefix = "..." if start > 0 else ""
    suffix = "..." if end < len(haystack) else ""
    return f"{prefix}{before}<mark>{hit}</mark>{after}{suffix}"


def simple_text_search(text: str, query: str, max_hits: int = 8) -> list[str]:
    if not text or not query:
        return []
    hits: list[str] = []
    lowered = text.lower()
    q = query.lower()
    cursor = 0
    while len(hits) < max_hits:
        idx = lowered.find(q, cursor)
        if idx < 0:
            break
        snippet = highlight_context(text, text[idx : idx + len(query)], context_chars=90)
        if snippet:
            hits.append(snippet)
        cursor = idx + len(query)
    return hits

