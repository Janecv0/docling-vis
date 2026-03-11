from __future__ import annotations

import base64
import io
import textwrap
from typing import Any

import bleach
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from bs4 import BeautifulSoup
from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation

from loaders import AssetRef, DocumentAssets
from utils import load_text_from_bytes

ALLOWED_HTML_TAGS = [
    "a",
    "abbr",
    "b",
    "blockquote",
    "br",
    "code",
    "div",
    "em",
    "h1",
    "h2",
    "h3",
    "h4",
    "h5",
    "h6",
    "hr",
    "i",
    "img",
    "li",
    "mark",
    "ol",
    "p",
    "pre",
    "span",
    "strong",
    "table",
    "tbody",
    "td",
    "th",
    "thead",
    "tr",
    "u",
    "ul",
]
ALLOWED_HTML_ATTRS = {
    "a": ["href", "title", "target"],
    "img": ["src", "alt", "title"],
    "span": ["class"],
    "div": ["class"],
    "p": ["class"],
    "code": ["class"],
    "table": ["class"],
}


def apply_theme_css(theme: str) -> None:
    if theme == "Dark":
        css = """
        <style>
        .review-card {
            border: 1px solid #3f3f46;
            border-radius: 10px;
            padding: 0.8rem;
            background: #141418;
        }
        .review-muted { color: #a1a1aa; }
        mark { background-color: #854d0e; color: #fff7ed; padding: 0.05rem 0.2rem; border-radius: 4px; }
        </style>
        """
    elif theme == "Light":
        css = """
        <style>
        .review-card {
            border: 1px solid #d4d4d8;
            border-radius: 10px;
            padding: 0.8rem;
            background: #ffffff;
        }
        .review-muted { color: #52525b; }
        mark { background-color: #fef08a; color: #111827; padding: 0.05rem 0.2rem; border-radius: 4px; }
        </style>
        """
    else:
        css = """
        <style>
        .review-card {
            border: 1px solid color-mix(in srgb, currentColor 30%, transparent);
            border-radius: 10px;
            padding: 0.8rem;
            background: color-mix(in srgb, currentColor 5%, transparent);
        }
        .review-muted { opacity: 0.8; }
        mark { padding: 0.05rem 0.2rem; border-radius: 4px; }
        </style>
        """
    st.markdown(css, unsafe_allow_html=True)


def render_asset_status_panel(doc: DocumentAssets) -> None:
    rows = []
    for kind in ("pdf", "pptx", "html", "md", "json"):
        asset = doc.get(kind)
        rows.append(
            {
                "asset": kind.upper(),
                "status": "Loaded" if asset else "Missing",
                "name": asset.name if asset else "",
            }
        )
    st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)


def render_pdf_preview(asset: AssetRef, height: int = 760) -> None:
    data = asset.read_bytes()
    if not data:
        st.warning("PDF data is missing or unreadable.")
        return
    encoded = base64.b64encode(data).decode("utf-8")
    iframe_html = (
        f'<iframe src="data:application/pdf;base64,{encoded}" '
        f'width="100%" height="{height}" type="application/pdf"></iframe>'
    )
    components.html(iframe_html, height=height + 10, scrolling=True)


def _extract_slide_text(slide: Any) -> tuple[str, list[str]]:
    title = ""
    lines: list[str] = []
    for shape in slide.shapes:
        if not hasattr(shape, "has_text_frame") or not shape.has_text_frame:
            continue
        try:
            text = (shape.text or "").strip()
        except Exception:
            text = ""
        if not text:
            continue
        if not title:
            title = text.splitlines()[0]
        lines.append(text)
    return title, lines


def _slide_card_image(slide_no: int, title: str, lines: list[str]) -> Image.Image:
    img = Image.new("RGB", (1280, 720), color=(248, 250, 252))
    draw = ImageDraw.Draw(img)
    font = ImageFont.load_default()
    draw.rectangle((22, 22, 1258, 698), outline=(30, 64, 175), width=3)
    draw.text((40, 36), f"Slide {slide_no}", fill=(30, 41, 59), font=font)
    draw.text((40, 72), title or "(No title detected)", fill=(15, 23, 42), font=font)
    y = 110
    wrapped = []
    for line in lines[:12]:
        wrapped.extend(textwrap.wrap(line, width=95) or [""])
    for line in wrapped[:24]:
        draw.text((48, y), f"- {line}", fill=(51, 65, 85), font=font)
        y += 22
        if y > 680:
            break
    return img


def render_pptx_preview(asset: AssetRef) -> None:
    data = asset.read_bytes()
    if not data:
        st.warning("PPTX data is missing or unreadable.")
        return
    try:
        presentation = Presentation(io.BytesIO(data))
    except Exception as exc:
        st.error(f"Could not parse PPTX: {exc}")
        return

    slide_count = len(presentation.slides)
    if slide_count == 0:
        st.info("PPTX has no slides.")
        return

    st.caption(
        "PPTX preview mode uses a Python-only slide card rendering fallback based on extracted text; "
        "it is not full visual fidelity."
    )
    selected_slide = st.number_input(
        "Slide",
        min_value=1,
        max_value=slide_count,
        value=1,
        step=1,
        key="pptx_slide_selector",
    )
    slide = presentation.slides[selected_slide - 1]
    title, lines = _extract_slide_text(slide)
    st.image(_slide_card_image(selected_slide, title, lines), use_container_width=True)
    with st.expander("Extracted slide text", expanded=False):
        if lines:
            st.text("\n\n".join(lines))
        else:
            st.write("No text extracted from this slide.")


def sanitize_html(raw_html: str) -> str:
    return bleach.clean(
        raw_html,
        tags=ALLOWED_HTML_TAGS,
        attributes=ALLOWED_HTML_ATTRS,
        protocols=["http", "https", "mailto", "data"],
        strip=True,
    )


def html_to_text(raw_html: str) -> str:
    soup = BeautifulSoup(raw_html, "html.parser")
    return soup.get_text("\n")


def render_html_preview(asset: AssetRef, height: int = 760) -> str:
    raw_bytes = asset.read_bytes()
    if not raw_bytes:
        st.warning("HTML file is missing or unreadable.")
        return ""
    raw_html = load_text_from_bytes(raw_bytes)
    safe_html = sanitize_html(raw_html)
    components.html(safe_html, height=height, scrolling=True)
    return raw_html


def render_markdown_preview(asset: AssetRef) -> str:
    raw_bytes = asset.read_bytes()
    if not raw_bytes:
        st.warning("Markdown file is missing or unreadable.")
        return ""
    markdown_text = load_text_from_bytes(raw_bytes)
    st.markdown(markdown_text)
    return markdown_text


def render_json_preview(json_obj: Any) -> None:
    st.json(json_obj, expanded=False)

