from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Sequence

from utils import canonical_doc_key, load_text_from_bytes


@dataclass
class AssetRef:
    kind: str
    name: str
    path: Path | None = None
    data: bytes | None = None

    def read_bytes(self) -> bytes | None:
        if self.data is not None:
            return self.data
        if self.path is None or not self.path.exists():
            return None
        try:
            return self.path.read_bytes()
        except OSError:
            return None

    def read_text(self) -> str | None:
        payload = self.read_bytes()
        if payload is None:
            return None
        return load_text_from_bytes(payload)


@dataclass
class DocumentAssets:
    key: str
    display_name: str
    assets: dict[str, AssetRef] = field(default_factory=dict)

    def set_asset(self, asset: AssetRef) -> None:
        self.assets[asset.kind] = asset

    def get(self, kind: str) -> AssetRef | None:
        return self.assets.get(kind)

    def available_kinds(self) -> list[str]:
        return sorted(self.assets.keys())


def detect_kind(file_name: str) -> str | None:
    lowered = file_name.lower()
    if lowered.endswith(".pdf"):
        return "pdf"
    if lowered.endswith(".pptx"):
        return "pptx"
    if lowered.endswith(".html") or lowered.endswith(".htm"):
        return "html"
    if lowered.endswith(".md") or lowered.endswith(".markdown"):
        return "md"
    if lowered.endswith(".json"):
        return "json"
    return None


def _upsert_asset(
    documents: dict[str, DocumentAssets],
    key: str,
    display_name: str,
    asset: AssetRef,
) -> None:
    if key not in documents:
        documents[key] = DocumentAssets(key=key, display_name=display_name)
    documents[key].set_asset(asset)


def build_docs_from_folder(folder: Path) -> tuple[dict[str, DocumentAssets], list[str]]:
    docs: dict[str, DocumentAssets] = {}
    warnings: list[str] = []
    if not folder.exists():
        return docs, [f"Folder does not exist: {folder}"]
    if not folder.is_dir():
        return docs, [f"Path is not a folder: {folder}"]

    try:
        files = [path for path in folder.iterdir() if path.is_file()]
    except OSError as exc:
        return docs, [f"Could not list folder: {exc}"]

    for path in files:
        kind = detect_kind(path.name)
        if kind is None:
            continue
        key = canonical_doc_key(path.name)
        _upsert_asset(
            docs,
            key=key,
            display_name=key,
            asset=AssetRef(kind=kind, name=path.name, path=path),
        )

    if not docs:
        warnings.append("No supported files found (.pdf, .pptx, .html, .md, .json).")
    return docs, warnings


def build_docs_from_uploads(uploaded_files: Sequence[Any]) -> dict[str, DocumentAssets]:
    docs: dict[str, DocumentAssets] = {}
    for uploaded in uploaded_files:
        name = getattr(uploaded, "name", "")
        kind = detect_kind(name)
        if kind is None:
            continue
        key = canonical_doc_key(name)
        raw = uploaded.getvalue()
        _upsert_asset(
            docs,
            key=key,
            display_name=key,
            asset=AssetRef(kind=kind, name=name, data=raw),
        )
    return docs


def merge_document_sets(*sources: dict[str, DocumentAssets]) -> dict[str, DocumentAssets]:
    merged: dict[str, DocumentAssets] = {}
    for source in sources:
        for key, doc in source.items():
            if key not in merged:
                merged[key] = DocumentAssets(key=doc.key, display_name=doc.display_name)
            for kind, asset in doc.assets.items():
                merged[key].assets[kind] = asset
    return merged


def load_json_asset(asset: AssetRef | None) -> tuple[Any | None, str | None]:
    if asset is None:
        return None, "Missing JSON file."
    payload = asset.read_text()
    if payload is None:
        return None, "Could not read JSON content."
    try:
        return json.loads(payload), None
    except json.JSONDecodeError as exc:
        return None, f"Malformed JSON: {exc}"


def load_sample_document(sample_json_path: Path) -> dict[str, DocumentAssets]:
    if not sample_json_path.exists():
        return {}
    doc = DocumentAssets(key="sample_doc", display_name="sample_doc")
    doc.set_asset(
        AssetRef(
            kind="json",
            name=sample_json_path.name,
            path=sample_json_path,
        )
    )
    return {"sample_doc": doc}

