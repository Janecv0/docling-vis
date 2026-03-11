from __future__ import annotations

import json
from dataclasses import dataclass
from typing import Any, Iterable

import pandas as pd

from utils import short_snippet

ID_KEYS = ("id", "block_id", "uuid", "uid", "self_ref", "ref")
LABEL_KEYS = ("label", "type", "kind", "category", "class", "name", "role")
TEXT_KEYS = (
    "text",
    "content",
    "value",
    "markdown",
    "md",
    "caption",
    "title",
    "raw_text",
    "normalized_text",
)
PAGE_KEYS = ("page", "page_no", "page_num", "page_number", "page_index", "page_idx")
BBOX_KEYS = ("bbox", "bounding_box", "box", "rect", "coordinates", "coords")


@dataclass
class BlockRecord:
    block_id: str
    label: str | None
    page: str | None
    text: str
    snippet: str
    path: str
    bbox: Any
    raw: dict[str, Any]

    def to_row(self) -> dict[str, Any]:
        return {
            "block_id": self.block_id,
            "label": self.label or "",
            "page": self.page or "",
            "snippet": self.snippet,
            "path": self.path,
        }


def _is_scalar(value: Any) -> bool:
    return isinstance(value, (str, int, float, bool))


def _first_scalar(node: dict[str, Any], keys: Iterable[str]) -> Any | None:
    for key in keys:
        value = node.get(key)
        if value is None:
            continue
        if _is_scalar(value):
            text = str(value).strip()
            if text:
                return text
    return None


def _value_to_text(value: Any, max_parts: int = 40) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, (int, float, bool)):
        return str(value)
    if isinstance(value, list):
        chunks: list[str] = []
        for item in value[:max_parts]:
            text = _value_to_text(item, max_parts=max_parts)
            if text:
                chunks.append(text)
        return "\n".join(chunks).strip()
    if isinstance(value, dict):
        for key in ("text", "content", "value", "title", "caption"):
            if key in value:
                text = _value_to_text(value.get(key), max_parts=max_parts)
                if text:
                    return text
        chunks = []
        for item in value.values():
            text = _value_to_text(item, max_parts=10)
            if text:
                chunks.append(text)
            if len(chunks) >= 8:
                break
        return " ".join(chunks).strip()
    return ""


def _deep_find_first(node: Any, keys: Iterable[str], max_depth: int = 4, _depth: int = 0) -> Any | None:
    if _depth > max_depth or node is None:
        return None
    if isinstance(node, dict):
        value = _first_scalar(node, keys)
        if value is not None:
            return value
        for child in node.values():
            found = _deep_find_first(child, keys=keys, max_depth=max_depth, _depth=_depth + 1)
            if found is not None:
                return found
    elif isinstance(node, list):
        for child in node[:40]:
            found = _deep_find_first(child, keys=keys, max_depth=max_depth, _depth=_depth + 1)
            if found is not None:
                return found
    return None


def _extract_text(node: dict[str, Any]) -> str:
    for key in TEXT_KEYS:
        if key in node:
            text = _value_to_text(node.get(key))
            if text:
                return text
    for key in ("items", "children", "spans", "lines", "paragraphs", "cells", "words"):
        if key in node:
            text = _value_to_text(node.get(key))
            if text:
                return text
    return ""


def _extract_page(node: dict[str, Any]) -> str | None:
    direct = _first_scalar(node, PAGE_KEYS)
    if direct is not None:
        return str(direct)
    for prov_key in ("prov", "provenance", "source", "location", "meta", "metadata"):
        if prov_key in node:
            page = _deep_find_first(node.get(prov_key), keys=PAGE_KEYS, max_depth=5)
            if page is not None:
                return str(page)
    return None


def _extract_bbox(node: dict[str, Any]) -> Any:
    for key in BBOX_KEYS:
        if key in node:
            return node.get(key)
    for prov_key in ("prov", "provenance", "location", "meta", "metadata"):
        if prov_key in node and isinstance(node[prov_key], dict):
            nested = _extract_bbox(node[prov_key])
            if nested is not None:
                return nested
    return None


def _looks_interesting(block_id: Any, label: Any, text: str, page: Any, bbox: Any) -> bool:
    score = int(bool(block_id)) + int(bool(label)) + int(bool(text)) + int(bool(page)) + int(bool(bbox))
    if text and (label or page or block_id or bbox):
        return True
    return score >= 3


def _build_block(node: dict[str, Any], path: str, index: int) -> BlockRecord | None:
    block_id = _first_scalar(node, ID_KEYS)
    label = _first_scalar(node, LABEL_KEYS)
    text = _extract_text(node)
    page = _extract_page(node)
    bbox = _extract_bbox(node)
    if not _looks_interesting(block_id, label, text, page, bbox):
        return None
    final_id = str(block_id) if block_id else f"block_{index + 1:04d}"
    fallback_text = text or json.dumps(node, ensure_ascii=False)[:300]
    return BlockRecord(
        block_id=final_id,
        label=str(label) if label else None,
        page=page,
        text=text,
        snippet=short_snippet(fallback_text),
        path=path,
        bbox=bbox,
        raw=node,
    )


def _collect_text_leaves(node: Any, limit: int = 120) -> list[str]:
    leaves: list[str] = []

    def visit(value: Any) -> None:
        if len(leaves) >= limit:
            return
        if isinstance(value, str):
            stripped = value.strip()
            if len(stripped) >= 20:
                leaves.append(stripped)
            return
        if isinstance(value, dict):
            for child in value.values():
                visit(child)
            return
        if isinstance(value, list):
            for child in value:
                visit(child)

    visit(node)
    return leaves


def extract_blocks(doc_json: Any, max_blocks: int = 5000) -> list[BlockRecord]:
    if doc_json is None:
        return []
    blocks: list[BlockRecord] = []
    seen: set[tuple[str, str, str]] = set()

    def visit(node: Any, path: str) -> None:
        if len(blocks) >= max_blocks:
            return
        if isinstance(node, dict):
            block = _build_block(node, path=path, index=len(blocks))
            if block is not None:
                signature = (block.block_id, block.path, block.snippet)
                if signature not in seen:
                    blocks.append(block)
                    seen.add(signature)
            for key, value in node.items():
                visit(value, f"{path}/{key}")
            return
        if isinstance(node, list):
            for idx, item in enumerate(node):
                visit(item, f"{path}[{idx}]")

    visit(doc_json, "$")

    if blocks:
        return blocks

    # Last-resort fallback: expose text leaves for inspection when schema is unknown.
    for idx, text in enumerate(_collect_text_leaves(doc_json, limit=200)):
        blocks.append(
            BlockRecord(
                block_id=f"text_{idx + 1:04d}",
                label="text_leaf",
                page=None,
                text=text,
                snippet=short_snippet(text),
                path="$",
                bbox=None,
                raw={"text": text},
            )
        )
    return blocks


def blocks_to_dataframe(blocks: list[BlockRecord]) -> pd.DataFrame:
    if not blocks:
        return pd.DataFrame(columns=["block_id", "label", "page", "snippet", "path"])
    rows = [block.to_row() for block in blocks]
    return pd.DataFrame(rows)


def filter_blocks(
    blocks: list[BlockRecord],
    query: str = "",
    labels: list[str] | None = None,
    pages: list[str] | None = None,
) -> list[BlockRecord]:
    filtered = blocks
    if labels:
        label_set = set(labels)
        filtered = [block for block in filtered if (block.label or "") in label_set]
    if pages:
        page_set = set(pages)
        filtered = [block for block in filtered if (block.page or "") in page_set]
    if query.strip():
        q = query.strip().lower()
        filtered = [
            block
            for block in filtered
            if q in block.block_id.lower()
            or q in (block.label or "").lower()
            or q in (block.text or "").lower()
            or q in block.snippet.lower()
        ]
    return filtered

