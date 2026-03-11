from __future__ import annotations

import json
import os
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
from parsers import blocks_to_dataframe, extract_blocks, filter_blocks
from ui_components import (
    apply_theme_css,
    html_to_text,
    render_asset_status_panel,
    render_html_preview,
    render_json_preview,
    render_markdown_preview,
    render_pdf_preview,
    render_pptx_preview,
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


def doc_view_options(doc: DocumentAssets) -> list[str]:
    options: list[str] = []
    if doc.get("html"):
        options.append("HTML")
    if doc.get("md"):
        options.append("Markdown")
    if doc.get("json"):
        options.append("JSON")
    return options


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
            uploaded_files = st.file_uploader(
                "Upload files (.pdf, .pptx, .html, .md, .json)",
                type=["pdf", "pptx", "html", "htm", "md", "markdown", "json"],
                accept_multiple_files=True,
                help="You can upload source file(s) and Docling outputs together.",
            )
        with row1_col2:
            st.session_state["folder_path"] = st.text_input(
                "Local folder path",
                value=st.session_state.get("folder_path", ""),
                placeholder=r"C:\path\to\folder",
            )
            folder_btn_col, sample_btn_col = st.columns(2)
            with folder_btn_col:
                if st.button("Load folder", use_container_width=True):
                    load_folder(st.session_state["folder_path"])
            with sample_btn_col:
                if st.button("Load sample JSON", use_container_width=True):
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
        ctl1, ctl2, ctl3 = st.columns([1.8, 1.1, 2.1])
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
            search_query = st.text_input("Search across selected output", value="", placeholder="Type text to find...")

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

    left_weight = int(st.session_state["pane_split"])
    left_col, right_col = st.columns([left_weight, 100 - left_weight], gap="large")

    right_plain_text = ""

    with left_col:
        st.subheader("Source preview")
        source_modes: list[str] = []
        if current_doc.get("pdf"):
            source_modes.append("PDF")
        if current_doc.get("pptx"):
            source_modes.append("PPTX")
        if not source_modes:
            st.info("No source file loaded for this document. Add a PDF or PPTX.")
        else:
            source_mode = st.radio("Source mode", options=source_modes, horizontal=True)
            if source_mode == "PDF":
                render_pdf_preview(current_doc.get("pdf"))  # type: ignore[arg-type]
            else:
                render_pptx_preview(current_doc.get("pptx"))  # type: ignore[arg-type]

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
                raw_html = render_html_preview(html_asset)
                right_plain_text = html_to_text(raw_html)
        elif selected_view == "Markdown":
            md_asset = current_doc.get("md")
            if md_asset:
                right_plain_text = render_markdown_preview(md_asset)
        elif selected_view == "JSON":
            if json_obj is not None:
                render_json_preview(json_obj)
                right_plain_text = json.dumps(json_obj, ensure_ascii=False, indent=2)
            else:
                st.warning(json_error or "JSON is unavailable.")

        if search_query.strip() and right_plain_text:
            matches = simple_text_search(right_plain_text, search_query.strip(), max_hits=6)
            with st.expander(f"Search matches ({len(matches)})", expanded=bool(matches)):
                if not matches:
                    st.write("No matches found in the currently visible output.")
                for idx, snippet in enumerate(matches, start=1):
                    st.markdown(f"{idx}. {snippet}", unsafe_allow_html=True)

    st.subheader("Metadata and block inspection")
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
            st.dataframe(blocks_to_dataframe(filtered), hide_index=True, use_container_width=True, height=280)

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
                use_container_width=True,
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


