from __future__ import annotations

import json
import os
import re
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any

import streamlit as st

from loaders import (
    DocumentAssets,
    build_docs_from_folder,
    build_docs_from_uploads,
    load_json_asset,
    load_sample_document,
    merge_document_sets,
)
from parsers import BlockRecord, blocks_to_dataframe, extract_blocks, filter_blocks
from ui_components import (
    apply_theme_css,
    get_pdf_page_count,
    html_to_text,
    render_asset_status_panel,
    render_html_preview,
    render_json_preview,
    render_markdown_preview,
    render_pdf_preview,
    render_pptx_preview,
    trigger_parent_center_scroll,
)
from utils import highlight_context, load_recent_folders, push_recent_folder, simple_text_search

st.set_page_config(page_title="Docling Local Review", layout="wide")


def init_state() -> None:
    defaults: dict[str, Any] = {
        "folder_docs": {},
        "folder_warnings": [],
        "folder_path": "",
        "sample_docs": {},
        "recent_folders": load_recent_folders(),
        "selected_doc_key": "",
        "theme_mode": "System",
        "view_mode": "HTML",
        "pane_split": 50,
        "sync_slides": True,
        "pdf_page_selector": 1,
        "pptx_slide_selector": 1,
        "pptx_scroll_ready": False,
        "source_mode_selector": "PPTX",
        "search_focus_text": "",
        "pending_nav_slide": None,
        "uploader_nonce": 0,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def load_folder(path_str: str) -> None:
    cleaned = path_str.strip()
    if not cleaned:
        st.session_state["folder_docs"] = {}
        st.session_state["folder_warnings"] = []
        return
    folder = Path(cleaned).expanduser()
    docs, warnings = build_docs_from_folder(folder)
    st.session_state["folder_docs"] = docs
    st.session_state["folder_warnings"] = warnings
    st.session_state["folder_path"] = str(folder)
    if docs:
        st.session_state["recent_folders"] = push_recent_folder(str(folder))


def clear_loaded_documents() -> None:
    st.session_state["folder_docs"] = {}
    st.session_state["folder_warnings"] = []
    st.session_state["sample_docs"] = {}
    st.session_state["folder_path"] = ""
    st.session_state["selected_doc_key"] = ""
    st.session_state["pending_nav_slide"] = None
    st.session_state["search_focus_text"] = ""
    st.session_state["search_query"] = ""
    st.session_state["uploader_nonce"] = int(st.session_state.get("uploader_nonce", 0)) + 1


def doc_view_options(doc: DocumentAssets) -> list[str]:
    options: list[str] = []
    if doc.get("html"):
        options.append("HTML")
    if doc.get("md"):
        options.append("Markdown")
    if doc.get("json"):
        options.append("JSON")
    return options


def _extract_int(value: str | None) -> int | None:
    if not value:
        return None
    match = re.search(r"\d+", str(value))
    if not match:
        return None
    try:
        return int(match.group(0))
    except ValueError:
        return None


def _find_synced_blocks(
    blocks: list[BlockRecord],
    source_index: int | None,
) -> tuple[list[BlockRecord], int | None]:
    if source_index is None or not blocks:
        return [], None
    for target in (source_index, source_index - 1, source_index + 1):
        if target <= 0:
            continue
        subset = [block for block in blocks if _extract_int(block.page) == target]
        if subset:
            return subset, target
    return [], None


def _normalize_text(value: str) -> str:
    lowered = value.lower()
    lowered = re.sub(r"[^\w\s]", " ", lowered)
    lowered = re.sub(r"\s+", " ", lowered)
    return lowered.strip()


def _tokens(value: str, min_len: int = 4) -> set[str]:
    return {token for token in _normalize_text(value).split() if len(token) >= min_len}


def _slide_title_from_text(slide_text: str) -> str:
    for line in slide_text.splitlines():
        candidate = line.strip()
        if len(candidate) >= 4:
            return candidate[:220]
    return ""


def _find_title_synced_blocks(
    blocks: list[BlockRecord],
    slide_title: str,
) -> tuple[list[BlockRecord], float]:
    title_norm = _normalize_text(slide_title)
    if len(title_norm) < 4:
        return [], 0.0

    scored: list[tuple[float, BlockRecord]] = []
    for block in blocks:
        candidate = (block.text or block.snippet or "").strip()
        if not candidate:
            continue
        candidate_norm = _normalize_text(candidate)
        if not candidate_norm:
            continue

        contains_bonus = 0.0
        if title_norm in candidate_norm:
            contains_bonus = 0.35
        ratio = SequenceMatcher(None, title_norm, candidate_norm[: max(80, len(title_norm) * 2)]).ratio()
        score = max(ratio + contains_bonus, ratio)
        if score >= 0.38:
            # Prefer heading/title-like blocks if label exists.
            label = (block.label or "").lower()
            if any(token in label for token in ("title", "heading", "header", "section")):
                score += 0.15
            scored.append((score, block))

    if not scored:
        return [], 0.0

    scored.sort(key=lambda item: item[0], reverse=True)
    top_score, top_block = scored[0]
    top_page = _extract_int(top_block.page)
    if top_page is not None:
        same_page = [block for block in blocks if _extract_int(block.page) == top_page]
        if same_page:
            return same_page[:80], top_score
    return [top_block], top_score


def _find_text_synced_blocks(
    blocks: list[BlockRecord],
    slide_text: str,
) -> tuple[list[BlockRecord], float]:
    normalized_slide = _normalize_text(slide_text)
    if len(normalized_slide) < 20:
        return [], 0.0

    slide_tokens = _tokens(slide_text)
    if not slide_tokens:
        return [], 0.0

    scored: list[tuple[float, BlockRecord]] = []
    for block in blocks:
        candidate = (block.text or block.snippet or "").strip()
        if len(candidate) < 12:
            continue
        candidate_tokens = _tokens(candidate)
        if not candidate_tokens:
            continue

        token_overlap = len(slide_tokens & candidate_tokens) / max(1, len(slide_tokens))
        sequence_ratio = SequenceMatcher(None, normalized_slide[:900], _normalize_text(candidate)[:900]).ratio()
        score = max(token_overlap, sequence_ratio)
        if score >= 0.14:
            scored.append((score, block))

    if not scored:
        return [], 0.0

    scored.sort(key=lambda item: item[0], reverse=True)
    top_score, top_block = scored[0]
    top_page = _extract_int(top_block.page)
    if top_page is not None:
        same_page = [block for block in blocks if _extract_int(block.page) == top_page]
        if same_page:
            return same_page[:80], top_score
    return [top_block], top_score


def main() -> None:
    init_state()
    is_railway = bool(
        os.getenv("RAILWAY_PROJECT_ID")
        or os.getenv("RAILWAY_ENVIRONMENT")
        or os.getenv("RAILWAY_ENVIRONMENT_NAME")
    )

    st.title("Docling Local Review App")
    st.caption("Local-first document QA/review for source files and Docling exports (also deployable on Railway).")
    if is_railway:
        st.info(
            "Running on Railway: folder paths refer to the server filesystem. "
            "Use file uploads for documents from your own machine."
        )

    top = st.container()
    with top:
        row1_col1, row1_col2, row1_col3 = st.columns([2.8, 2.2, 1.0])
        with row1_col1:
            upload_key = f"upload_files_{int(st.session_state.get('uploader_nonce', 0))}"
            uploaded_files = st.file_uploader(
                "Upload files (.pdf, .pptx, .html, .md, .json)",
                type=["pdf", "pptx", "html", "htm", "md", "markdown", "json"],
                accept_multiple_files=True,
                help="You can upload source file(s) and Docling outputs together.",
                key=upload_key,
            )
            if st.button("Remove all files", width="content", help="Clear uploads and loaded folder/sample docs."):
                clear_loaded_documents()
                st.rerun()
        with row1_col2:
            st.session_state["folder_path"] = st.text_input(
                "Local folder path",
                value=st.session_state.get("folder_path", ""),
                placeholder=r"C:\path\to\folder",
            )
            folder_btn_col, sample_btn_col = st.columns(2)
            with folder_btn_col:
                if st.button("Load folder", width="stretch"):
                    load_folder(st.session_state["folder_path"])
            with sample_btn_col:
                if st.button("Load sample JSON", width="stretch"):
                    sample_path = Path("sample_data") / "sample_docling.json"
                    st.session_state["sample_docs"] = load_sample_document(sample_path)

            recent = st.selectbox(
                "Recent folders",
                options=[""] + st.session_state.get("recent_folders", []),
                index=0,
                help="Select a recent path and click Load folder to scan it.",
            )
            if recent:
                st.session_state["folder_path"] = recent
        with row1_col3:
            st.session_state["theme_mode"] = st.selectbox(
                "Theme",
                options=["System", "Light", "Dark"],
                index=["System", "Light", "Dark"].index(st.session_state.get("theme_mode", "System")),
            )
            st.session_state["pane_split"] = st.slider(
                "Pane split %",
                min_value=35,
                max_value=65,
                value=int(st.session_state.get("pane_split", 50)),
                step=5,
            )

    apply_theme_css(st.session_state["theme_mode"])

    upload_docs = build_docs_from_uploads(uploaded_files or [])
    docs = merge_document_sets(st.session_state["folder_docs"], st.session_state["sample_docs"], upload_docs)

    controls = st.container()
    with controls:
        ctl1, ctl2, ctl3 = st.columns([1.8, 1.0, 1.3])
        doc_keys = sorted(docs.keys())
        with ctl1:
            if not doc_keys:
                st.selectbox("Document", options=["(no documents loaded)"], index=0, disabled=True)
                selected_doc_key = ""
            else:
                current_selected = st.session_state.get("selected_doc_key", "")
                default_index = doc_keys.index(current_selected) if current_selected in doc_keys else 0
                selected_doc_key = st.selectbox(
                    "Document",
                    options=doc_keys,
                    index=default_index,
                    key="selected_doc_key",
                )
        with ctl2:
            st.session_state["view_mode"] = st.selectbox(
                "Right view",
                options=["HTML", "Markdown", "JSON"],
                index=["HTML", "Markdown", "JSON"].index(st.session_state.get("view_mode", "HTML")),
            )
        with ctl3:
            st.toggle(
                "Synchronise pages/slides",
                key="sync_slides",
                help="When enabled, the right pane follows the selected PDF page or PPTX slide using Docling JSON page mapping.",
            )

    if not docs:
        st.info(
            "No files loaded yet. Upload files above or provide a local folder path with matching source/docling exports."
        )
        st.stop()

    current_doc = docs[selected_doc_key]
    json_obj, json_error = load_json_asset(current_doc.get("json"))
    blocks = extract_blocks(json_obj) if json_obj is not None else []

    with st.expander("Status panel", expanded=True):
        st.write(f"Document key: `{current_doc.key}`")
        render_asset_status_panel(current_doc)
        for warning in st.session_state.get("folder_warnings", []):
            st.warning(warning)
        if json_error and current_doc.get("json"):
            st.warning(json_error)

    with st.expander("Search panel", expanded=False):
        search_query = st.text_input(
            "Search across selected output",
            value=st.session_state.get("search_query", ""),
            key="search_query",
            placeholder="Type text to find...",
        )
        search_matches_container = st.container()
        search_block_hits: list[BlockRecord] = []
        if search_query.strip() and blocks:
            search_block_hits = filter_blocks(blocks, query=search_query.strip())[:12]

    left_weight = int(st.session_state["pane_split"])
    left_col, right_col = st.columns([left_weight, 100 - left_weight], gap="large")

    right_plain_text = ""
    source_mode_selected: str | None = None
    source_focus_kind = ""
    source_focus_index: int | None = None
    source_slide_text = ""
    slide_changed = False
    preview_height = 620

    with left_col:
        st.subheader("Source preview")
        st.markdown('<div id="source-preview-anchor"></div>', unsafe_allow_html=True)
        source_modes: list[str] = []
        if current_doc.get("pdf"):
            source_modes.append("PDF")
        if current_doc.get("pptx"):
            source_modes.append("PPTX")
        pending_nav_slide = st.session_state.get("pending_nav_slide")
        if pending_nav_slide is not None and "PPTX" in source_modes:
            try:
                pending_value = max(1, int(pending_nav_slide))
                st.session_state["source_mode_selector"] = "PPTX"
                st.session_state["pptx_slide_selector"] = pending_value
            except (TypeError, ValueError):
                pass
            st.session_state["pending_nav_slide"] = None
        if not source_modes:
            st.info("No source file loaded for this document. Add a PDF or PPTX.")
        else:
            if st.session_state.get("source_mode_selector") not in source_modes:
                st.session_state["source_mode_selector"] = source_modes[0]
            source_mode = str(st.session_state.get("source_mode_selector", source_modes[0]))
            source_mode_selected = source_mode
            if source_mode == "PDF":
                pdf_asset = current_doc.get("pdf")
                page_count = get_pdf_page_count(pdf_asset) if pdf_asset else None
                current_page = max(1, int(st.session_state.get("pdf_page_selector", 1)))
                if page_count:
                    current_page = min(current_page, page_count)
                st.session_state["pdf_page_selector"] = current_page

                render_pdf_preview(pdf_asset, height=preview_height, page=current_page)  # type: ignore[arg-type]

                nav_prev_col, nav_mid_col, nav_next_col = st.columns([1, 2, 1], gap="small")
                with nav_prev_col:
                    prev_clicked = st.button("←", key="pdf_prev_page", width="stretch")
                with nav_mid_col:
                    total_text = f" / {page_count}" if page_count else ""
                    st.markdown(f"**Page {current_page}{total_text}**")
                with nav_next_col:
                    next_clicked = st.button("→", key="pdf_next_page", width="stretch")

                target_page = current_page
                if prev_clicked and current_page > 1:
                    target_page -= 1
                if next_clicked and (page_count is None or current_page < page_count):
                    target_page += 1
                if target_page != current_page:
                    st.session_state["pdf_page_selector"] = target_page
                    st.rerun()

                source_focus_kind = "page"
                source_focus_index = current_page
            else:
                source_focus_kind = "slide"
                source_focus_index, source_slide_text, slide_changed = render_pptx_preview(current_doc.get("pptx"))  # type: ignore[arg-type]
                if slide_changed:
                    if st.session_state.get("pptx_scroll_ready", False):
                        trigger_parent_center_scroll("source-preview-anchor")
                    else:
                        st.session_state["pptx_scroll_ready"] = True
            st.radio("Source mode", options=source_modes, horizontal=True, key="source_mode_selector")

    sync_slides_enabled = bool(st.session_state.get("sync_slides", True))
    apply_sync = sync_slides_enabled and source_mode_selected in {"PPTX", "PDF"}

    synced_blocks: list[BlockRecord] = []
    synced_page_value: int | None = None
    sync_method = ""
    text_match_score = 0.0
    title_match_score = 0.0
    if apply_sync:
        synced_blocks, synced_page_value = _find_synced_blocks(blocks, source_focus_index)
        if synced_blocks:
            sync_method = "page"
        elif source_slide_text.strip():
            slide_title = _slide_title_from_text(source_slide_text)
            if slide_title:
                synced_blocks, title_match_score = _find_title_synced_blocks(blocks, slide_title)
                if synced_blocks:
                    sync_method = "title"
            if not synced_blocks:
                synced_blocks, text_match_score = _find_text_synced_blocks(blocks, source_slide_text)
                if synced_blocks:
                    sync_method = "text"
    synced_block = synced_blocks[0] if synced_blocks else None
    focus_text = ""
    if synced_block is not None:
        focus_text = (synced_block.text or synced_block.snippet).strip()
    elif apply_sync and source_mode_selected == "PPTX":
        # Fallback when JSON page/block mapping is weak: use slide title/text directly.
        slide_title = _slide_title_from_text(source_slide_text)
        focus_text = slide_title or source_slide_text.strip()
    search_focus_text = (st.session_state.get("search_focus_text") or "").strip()
    if search_focus_text:
        focus_text = search_focus_text

    with right_col:
        st.subheader("Docling output preview")
        
        selected_view = st.session_state["view_mode"]
        available_views = doc_view_options(current_doc)
        if selected_view not in available_views and available_views:
            st.warning(f"{selected_view} view is unavailable; showing {available_views[0]} instead.")
            selected_view = available_views[0]

        if not available_views:
            st.info("No Docling outputs loaded (.html, .md, .json).")
        elif selected_view == "HTML":
            html_asset = current_doc.get("html")
            if html_asset:
                raw_html = render_html_preview(html_asset, height=preview_height, focus_text=focus_text)
                right_plain_text = html_to_text(raw_html)
        elif selected_view == "Markdown":
            md_asset = current_doc.get("md")
            if md_asset:
                right_plain_text = render_markdown_preview(md_asset, height=preview_height, focus_text=focus_text)
        elif selected_view == "JSON":
            if json_obj is not None:
                right_plain_text = render_json_preview(json_obj, height=preview_height, focus_text=focus_text)
            else:
                st.warning(json_error or "JSON is unavailable.")

    with search_matches_container:
        if search_query.strip() and search_block_hits:
            st.caption(f"Search matches: {len(search_block_hits)}")
            for idx, hit in enumerate(search_block_hits, start=1):
                label = (
                    f"{idx}. p{hit.page or '-'} | {hit.label or 'block'} | {hit.snippet}"
                )
                if st.button(label, key=f"search_hit_{idx}", width="stretch"):
                    target_page = _extract_int(hit.page)
                    if target_page is not None:
                        if source_mode_selected == "PDF" and current_doc.get("pdf"):
                            st.session_state["source_mode_selector"] = "PDF"
                            st.session_state["pdf_page_selector"] = max(1, target_page)
                        elif current_doc.get("pptx"):
                            st.session_state["source_mode_selector"] = "PPTX"
                            st.session_state["pending_nav_slide"] = target_page
                        elif current_doc.get("pdf"):
                            st.session_state["source_mode_selector"] = "PDF"
                            st.session_state["pdf_page_selector"] = max(1, target_page)
                    if current_doc.get("md"):
                        st.session_state["view_mode"] = "Markdown"
                    st.session_state["search_focus_text"] = (hit.text or hit.snippet or "").strip()
                    st.rerun()
        elif search_query.strip() and right_plain_text:
            matches = simple_text_search(right_plain_text, search_query.strip(), max_hits=8)
            st.caption(f"Search matches: {len(matches)}")
            if not matches:
                st.write("No matches found in the currently visible output.")
            for idx, snippet in enumerate(matches, start=1):
                st.markdown(f"{idx}. {snippet}", unsafe_allow_html=True)
        elif search_query.strip():
            st.caption("Search matches: 0")
            st.write("No searchable content in the current right view.")

    if search_focus_text:
        st.session_state["search_focus_text"] = ""

    if apply_sync and source_focus_index is not None and blocks:
        if synced_block is not None:
            if sync_method == "page":
                st.caption(
                    f"Synced to {source_focus_kind} {source_focus_index} via block "
                    f"`{synced_block.block_id}` (page={synced_page_value or synced_block.page or '-'})"
                )
            elif sync_method == "title":
                st.caption(
                    f"Synced to {source_focus_kind} {source_focus_index} by slide title match "
                    f"(score={title_match_score:.2f}) using block `{synced_block.block_id}` "
                    f"(page={synced_block.page or '-'})"
                )
            elif sync_method == "text":
                st.caption(
                    f"Synced to {source_focus_kind} {source_focus_index} by text similarity "
                    f"(score={text_match_score:.2f}) using block `{synced_block.block_id}` "
                    f"(page={synced_block.page or '-'})"
                )
        else:
            if source_mode_selected == "PPTX":
                st.caption(
                    f"No JSON block matched {source_focus_kind} {source_focus_index}; "
                    "using slide title/text fallback for right-side focus."
                )
            else:
                st.caption(
                    f"No JSON block matched {source_focus_kind} {source_focus_index}; "
                    "right-side focus is unavailable for this selection."
                )

    with st.expander("Metadata and block inspection", expanded=False):
        meta_col, detail_col = st.columns([1.6, 1.0], gap="large")

        selected_block = None
        with meta_col:
            if current_doc.get("json") is None:
                st.info("No JSON loaded. Block extraction and metadata table require Docling JSON.")
            elif json_error:
                st.error(json_error)
            elif not blocks:
                st.info("No blocks detected in JSON. You can still inspect the raw JSON in right pane.")
            else:
                labels = sorted({block.label for block in blocks if block.label})
                pages = sorted({block.page for block in blocks if block.page})
                f1, f2, f3 = st.columns([2.0, 1.5, 1.5])
                with f1:
                    block_query = st.text_input("Filter blocks", placeholder="id, label, text...")
                with f2:
                    label_filter = st.multiselect("Label/type", options=labels)
                with f3:
                    page_filter = st.multiselect("Page", options=pages)

                filtered = filter_blocks(blocks, query=block_query, labels=label_filter, pages=page_filter)
                st.caption(f"Showing {len(filtered)} of {len(blocks)} extracted blocks.")
                st.dataframe(blocks_to_dataframe(filtered), hide_index=True, width="stretch", height=280)

                if filtered:
                    option_ids = list(range(len(filtered)))
                    selected_idx = st.selectbox(
                        "Select block",
                        options=option_ids,
                        format_func=lambda idx: (
                            f"{filtered[idx].block_id} | p{filtered[idx].page or '-'} | {filtered[idx].snippet}"
                        ),
                    )
                    selected_block = filtered[selected_idx]

        with detail_col:
            st.markdown('<div class="review-card">', unsafe_allow_html=True)
            st.markdown("**Detail panel**")
            if selected_block is None:
                st.write("Select a block to inspect metadata and focused snippet.")
            else:
                st.json(
                    {
                        "block_id": selected_block.block_id,
                        "label": selected_block.label,
                        "page": selected_block.page,
                        "path": selected_block.path,
                        "bbox": selected_block.bbox,
                        "snippet": selected_block.snippet,
                    }
                )
                st.download_button(
                    "Export selected block JSON",
                    data=json.dumps(selected_block.raw, ensure_ascii=False, indent=2),
                    file_name=f"{selected_block.block_id}.json",
                    mime="application/json",
                    width="stretch",
                )
                st.markdown("**Raw block JSON**")
                st.json(selected_block.raw, expanded=False)

                needle = (selected_block.text or selected_block.snippet or selected_block.block_id).strip()
                if needle and right_plain_text:
                    focused = highlight_context(right_plain_text, needle[:120], context_chars=160)
                    if not focused and len(needle) > 24:
                        focused = highlight_context(right_plain_text, needle[:24], context_chars=160)
                    if focused:
                        st.markdown("**Focused text snippet (fallback highlight)**")
                        st.markdown(focused, unsafe_allow_html=True)
                    else:
                        st.info("No direct text match found in the currently visible right-pane content.")
                elif not right_plain_text:
                    st.info("Focused snippet appears after loading HTML/Markdown/JSON content on the right.")
            st.markdown("</div>", unsafe_allow_html=True)


if __name__ == "__main__":
    main()
