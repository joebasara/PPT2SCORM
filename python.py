import base64
import html
import io
import math
import os
import re
import zipfile
from pathlib import Path
from typing import Optional

import streamlit as st
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.action import PP_ACTION
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE, MSO_SHAPE_TYPE


st.set_page_config(page_title="PPT to HTML + SCORM", layout="wide")
st.title("PPT to HTML + SCORM")


# =========================
# Utility helpers
# =========================
EMU_PER_INCH = 914400
PX_PER_INCH = 96


def emu_to_px(emu: int) -> float:
    return emu * PX_PER_INCH / EMU_PER_INCH


def sanitize_filename(name: str) -> str:
    name = re.sub(r"[^\w\-. ]+", "_", name).strip()
    return name or "package"


def css_escape(s: str) -> str:
    return s.replace("\\", "\\\\").replace('"', '\\"')


def html_text(s: str) -> str:
    return html.escape(s).replace("\n", "<br>")


def rgb_to_css(rgb) -> Optional[str]:
    try:
        if rgb is None:
            return None
        if isinstance(rgb, RGBColor):
            return f"#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}"
        if hasattr(rgb, "__iter__"):
            vals = list(rgb)
            if len(vals) >= 3:
                return f"#{vals[0]:02x}{vals[1]:02x}{vals[2]:02x}"
    except Exception:
        pass
    return None


def get_solid_fill_css(shape) -> Optional[str]:
    try:
        fill = shape.fill
        if fill is None:
            return None
        if fill.type == 1:  # solid
            fore = fill.fore_color
            color = rgb_to_css(getattr(fore, "rgb", None))
            if color:
                return color
    except Exception:
        pass
    return None


def get_line_css(shape) -> tuple[Optional[str], Optional[float]]:
    color = None
    width = None
    try:
        line = shape.line
        if line is not None:
            color = rgb_to_css(getattr(line.color, "rgb", None))
            if getattr(line, "width", None):
                width = emu_to_px(line.width)
    except Exception:
        pass
    return color, width


def get_rotation_deg(shape) -> float:
    try:
        return float(shape.rotation or 0)
    except Exception:
        return 0.0


def blob_to_data_uri(blob: bytes, ext: str) -> str:
    ext = (ext or "").lower().strip(".")
    mime_map = {
        "png": "image/png",
        "jpg": "image/jpeg",
        "jpeg": "image/jpeg",
        "gif": "image/gif",
        "bmp": "image/bmp",
        "svg": "image/svg+xml",
        "tif": "image/tiff",
        "tiff": "image/tiff",
        "webp": "image/webp",
        "wmf": "image/wmf",
        "emf": "image/emf",
    }
    mime = mime_map.get(ext, "application/octet-stream")
    b64 = base64.b64encode(blob).decode("utf-8")
    return f"data:{mime};base64,{b64}"


def get_shape_action_target(shape, slide_index: int, total_slides: int, slide_id_to_index: dict[int, int]) -> Optional[int]:
    try:
        action = shape.click_action.action

        if action == PP_ACTION.NAMED_SLIDE:
            target = shape.click_action.target_slide
            if target is not None:
                return slide_id_to_index.get(target.slide_id)

        if action == PP_ACTION.FIRST_SLIDE:
            return 1
        if action == PP_ACTION.LAST_SLIDE:
            return total_slides
        if action == PP_ACTION.NEXT_SLIDE:
            return min(total_slides, slide_index + 1)
        if action == PP_ACTION.PREVIOUS_SLIDE:
            return max(1, slide_index - 1)
    except Exception:
        return None

    return None


def build_slide_id_to_index(prs: Presentation) -> dict[int, int]:
    out = {}
    for i, slide in enumerate(prs.slides, start=1):
        out[slide.slide_id] = i
    return out


def is_supported_autoshape(shape) -> bool:
    try:
        if shape.shape_type != MSO_SHAPE_TYPE.AUTO_SHAPE:
            return False

        supported = {
            MSO_AUTO_SHAPE_TYPE.RECTANGLE,
            MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
            MSO_AUTO_SHAPE_TYPE.OVAL,
            MSO_AUTO_SHAPE_TYPE.DIAMOND,
            MSO_AUTO_SHAPE_TYPE.ISOSCELES_TRIANGLE,
            MSO_AUTO_SHAPE_TYPE.RIGHT_TRIANGLE,
            MSO_AUTO_SHAPE_TYPE.HEXAGON,
            MSO_AUTO_SHAPE_TYPE.PENTAGON,
            MSO_AUTO_SHAPE_TYPE.OCTAGON,
            MSO_AUTO_SHAPE_TYPE.PARALLELOGRAM,
            MSO_AUTO_SHAPE_TYPE.TRAPEZOID,
        }
        return shape.auto_shape_type in supported
    except Exception:
        return False


def autoshape_clip_path(shape) -> Optional[str]:
    try:
        t = shape.auto_shape_type
        mapping = {
            MSO_AUTO_SHAPE_TYPE.RECTANGLE: None,
            MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE: None,
            MSO_AUTO_SHAPE_TYPE.OVAL: "ellipse(50% 50% at 50% 50%)",
            MSO_AUTO_SHAPE_TYPE.DIAMOND: "polygon(50% 0%, 100% 50%, 50% 100%, 0% 50%)",
            MSO_AUTO_SHAPE_TYPE.ISOSCELES_TRIANGLE: "polygon(50% 0%, 100% 100%, 0% 100%)",
            MSO_AUTO_SHAPE_TYPE.RIGHT_TRIANGLE: "polygon(0% 0%, 100% 100%, 0% 100%)",
            MSO_AUTO_SHAPE_TYPE.HEXAGON: "polygon(25% 0%, 75% 0%, 100% 50%, 75% 100%, 25% 100%, 0% 50%)",
            MSO_AUTO_SHAPE_TYPE.PENTAGON: "polygon(50% 0%, 100% 38%, 82% 100%, 18% 100%, 0% 38%)",
            MSO_AUTO_SHAPE_TYPE.OCTAGON: "polygon(30% 0%, 70% 0%, 100% 30%, 100% 70%, 70% 100%, 30% 100%, 0% 70%, 0% 30%)",
            MSO_AUTO_SHAPE_TYPE.PARALLELOGRAM: "polygon(18% 0%, 100% 0%, 82% 100%, 0% 100%)",
            MSO_AUTO_SHAPE_TYPE.TRAPEZOID: "polygon(20% 0%, 80% 0%, 100% 100%, 0% 100%)",
        }
        return mapping.get(t)
    except Exception:
        return None


# =========================
# Rendering helpers
# =========================
def render_text_frame_html(shape) -> str:
    tf = shape.text_frame
    parts = []

    margin_left = emu_to_px(getattr(tf, "margin_left", 0))
    margin_right = emu_to_px(getattr(tf, "margin_right", 0))
    margin_top = emu_to_px(getattr(tf, "margin_top", 0))
    margin_bottom = emu_to_px(getattr(tf, "margin_bottom", 0))

    style_outer = [
        "position:absolute",
        "inset:0",
        f"padding:{margin_top:.2f}px {margin_right:.2f}px {margin_bottom:.2f}px {margin_left:.2f}px",
        "overflow:hidden",
        "display:flex",
        "flex-direction:column",
        "justify-content:flex-start",
        "white-space:normal",
        "word-break:break-word",
    ]

    parts.append(f'<div style="{";".join(style_outer)}">')

    for para in tf.paragraphs:
        align_css = "left"
        try:
            # 1 left, 2 center, 3 right, 4 justify, etc.
            if para.alignment == 2:
                align_css = "center"
            elif para.alignment == 3:
                align_css = "right"
            elif para.alignment == 4:
                align_css = "justify"
        except Exception:
            pass

        para_html = []

        if not para.runs and para.text:
            para_html.append(html_text(para.text))

        for run in para.runs:
            text = html_text(run.text or "")
            font = run.font

            style = []
            size = None
            try:
                if font.size:
                    size = font.size.pt
            except Exception:
                size = None
            if size is None:
                size = 18
            style.append(f"font-size:{size:.2f}px")

            try:
                if font.name:
                    style.append(f'font-family:"{css_escape(font.name)}", Arial, sans-serif')
            except Exception:
                style.append('font-family:Arial, sans-serif')

            try:
                if font.bold:
                    style.append("font-weight:700")
            except Exception:
                pass

            try:
                if font.italic:
                    style.append("font-style:italic")
            except Exception:
                pass

            deco = []
            try:
                if font.underline:
                    deco.append("underline")
            except Exception:
                pass
            if deco:
                style.append(f"text-decoration:{' '.join(deco)}")

            try:
                color = rgb_to_css(font.color.rgb)
                if color:
                    style.append(f"color:{color}")
            except Exception:
                pass

            para_html.append(f'<span style="{";".join(style)}">{text}</span>')

        parts.append(
            f'<div style="margin:0; text-align:{align_css}; line-height:1.2;">{"".join(para_html) or "&nbsp;"}</div>'
        )

    parts.append("</div>")
    return "".join(parts)


def render_picture_html(shape) -> str:
    image = shape.image
    ext = getattr(image, "ext", "png")
    data_uri = blob_to_data_uri(image.blob, ext)

    crop_left = getattr(shape, "crop_left", 0.0) or 0.0
    crop_right = getattr(shape, "crop_right", 0.0) or 0.0
    crop_top = getattr(shape, "crop_top", 0.0) or 0.0
    crop_bottom = getattr(shape, "crop_bottom", 0.0) or 0.0

    # Approximation of PPT crop using enlarged image with negative offsets
    visible_w = max(0.001, 1.0 - crop_left - crop_right)
    visible_h = max(0.001, 1.0 - crop_top - crop_bottom)

    img_w_pct = 100.0 / visible_w
    img_h_pct = 100.0 / visible_h
    left_pct = -(crop_left / visible_w) * 100.0
    top_pct = -(crop_top / visible_h) * 100.0

    return (
        '<div style="position:absolute; inset:0; overflow:hidden;">'
        f'<img src="{data_uri}" alt="" '
        f'style="position:absolute; left:{left_pct:.4f}%; top:{top_pct:.4f}%; width:{img_w_pct:.4f}%; height:{img_h_pct:.4f}%; object-fit:fill; user-select:none; -webkit-user-drag:none;">'
        "</div>"
    )


def render_autoshape_html(shape) -> str:
    fill = get_solid_fill_css(shape)
    line_color, line_width = get_line_css(shape)
    clip = autoshape_clip_path(shape)

    style = [
        "position:absolute",
        "inset:0",
    ]

    if fill:
        style.append(f"background:{fill}")
    else:
        style.append("background:transparent")

    if line_color and line_width and line_width > 0:
        style.append(f"border:{line_width:.2f}px solid {line_color}")
    else:
        style.append("border:none")

    if clip:
        style.append(f"clip-path:{clip}")

    # Rounded rectangle approximation
    try:
        if shape.auto_shape_type == MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE:
            style.append("border-radius:12px")
    except Exception:
        pass

    return f'<div style="{";".join(style)}"></div>'


def render_shape_box(shape, inner_html: str, click_target: Optional[int]) -> str:
    x = emu_to_px(shape.left)
    y = emu_to_px(shape.top)
    w = emu_to_px(shape.width)
    h = emu_to_px(shape.height)
    rot = get_rotation_deg(shape)

    style = [
        "position:absolute",
        f"left:{x:.2f}px",
        f"top:{y:.2f}px",
        f"width:{w:.2f}px",
        f"height:{h:.2f}px",
        f"transform:rotate({rot:.4f}deg)",
        "transform-origin:center center",
        "box-sizing:border-box",
    ]

    clickable_overlay = ""
    if click_target:
        clickable_overlay = (
            f'<button class="hotspot-btn" type="button" aria-label="Go to slide {click_target}" '
            f'onclick="goToSlide({click_target})" '
            'style="position:absolute; inset:0; background:transparent; border:none; cursor:pointer; padding:0; margin:0;"></button>'
        )

    return f'<div style="{";".join(style)}">{inner_html}{clickable_overlay}</div>'


# =========================
# PPT parsing
# =========================
def parse_ppt(prs: Presentation):
    slide_id_to_index = build_slide_id_to_index(prs)
    total_slides = len(prs.slides)

    slides_out = []
    warnings = []

    for slide_index, slide in enumerate(prs.slides, start=1):
        slide_items = []

        for shape in slide.shapes:
            try:
                click_target = get_shape_action_target(shape, slide_index, total_slides, slide_id_to_index)
                stype = shape.shape_type

                if stype == MSO_SHAPE_TYPE.PICTURE:
                    html_block = render_shape_box(shape, render_picture_html(shape), click_target)
                    slide_items.append(html_block)
                    continue

                if stype == MSO_SHAPE_TYPE.TEXT_BOX:
                    html_block = render_shape_box(shape, render_text_frame_html(shape), click_target)
                    slide_items.append(html_block)
                    continue

                if stype == MSO_SHAPE_TYPE.AUTO_SHAPE:
                    inner = []

                    if is_supported_autoshape(shape):
                        inner.append(render_autoshape_html(shape))
                    else:
                        warnings.append(
                            f"Slide {slide_index}: unsupported AutoShape '{getattr(shape, 'name', 'Unnamed')}'"
                        )
                        continue

                    if getattr(shape, "has_text_frame", False) and shape.text_frame and shape.text_frame.text:
                        inner.append(render_text_frame_html(shape))

                    html_block = render_shape_box(shape, "".join(inner), click_target)
                    slide_items.append(html_block)
                    continue

                if getattr(shape, "has_text_frame", False) and shape.text_frame and shape.text_frame.text:
                    html_block = render_shape_box(shape, render_text_frame_html(shape), click_target)
                    slide_items.append(html_block)
                    warnings.append(
                        f"Slide {slide_index}: approximated text-bearing shape '{getattr(shape, 'name', 'Unnamed')}'"
                    )
                    continue

                if stype == MSO_SHAPE_TYPE.GROUP:
                    warnings.append(f"Slide {slide_index}: unsupported grouped shape '{getattr(shape, 'name', 'Unnamed')}'")
                    continue

                if stype == MSO_SHAPE_TYPE.CHART:
                    warnings.append(f"Slide {slide_index}: unsupported chart '{getattr(shape, 'name', 'Unnamed')}'")
                    continue

                if stype == MSO_SHAPE_TYPE.TABLE:
                    warnings.append(f"Slide {slide_index}: unsupported table '{getattr(shape, 'name', 'Unnamed')}'")
                    continue

                if stype == MSO_SHAPE_TYPE.SMART_ART:
                    warnings.append(f"Slide {slide_index}: unsupported SmartArt '{getattr(shape, 'name', 'Unnamed')}'")
                    continue

                warnings.append(f"Slide {slide_index}: unsupported shape type '{stype}' in '{getattr(shape, 'name', 'Unnamed')}'")

            except Exception as e:
                warnings.append(f"Slide {slide_index}: failed to render '{getattr(shape, 'name', 'Unnamed')}' ({e})")

        slides_out.append(slide_items)

    return slides_out, warnings


# =========================
# HTML + SCORM generation
# =========================
def build_player_html(course_title: str, width_px: int, height_px: int, slides_items: list[list[str]]) -> str:
    slides_html = []
    for i, items in enumerate(slides_items, start=1):
        content = "".join(items)
        slides_html.append(
            f'<div class="slide" id="slide-{i}" data-slide-index="{i}" style="display:none;">{content}</div>'
        )

    slides_joined = "".join(slides_html)
    aspect = width_px / height_px

    return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{html.escape(course_title)}</title>
  <style>
    :root {{
      --bg: #111;
      --panel: #1b1b1b;
      --btn: #2a2a2a;
      --btn-hover: #383838;
      --text: #fff;
    }}

    * {{
      box-sizing: border-box;
    }}

    html, body {{
      width: 100%;
      height: 100%;
      margin: 0;
      padding: 0;
      background: var(--bg);
      color: var(--text);
      font-family: Arial, Helvetica, sans-serif;
      overflow: hidden;
    }}

    .app {{
      width: 100%;
      height: 100%;
      display: flex;
      flex-direction: column;
    }}

    .topbar {{
      flex: 0 0 auto;
      display: flex;
      justify-content: center;
      align-items: center;
      gap: 12px;
      padding: 14px 16px;
      background: rgba(0,0,0,0.4);
      border-bottom: 1px solid rgba(255,255,255,0.08);
      z-index: 10;
    }}

    .topbar button {{
      background: var(--btn);
      color: var(--text);
      border: 0;
      border-radius: 10px;
      padding: 12px 18px;
      font-size: 18px;
      cursor: pointer;
      min-width: 110px;
    }}

    .topbar button:hover {{
      background: var(--btn-hover);
    }}

    .topbar button:disabled {{
      opacity: 0.45;
      cursor: not-allowed;
    }}

    .counter {{
      min-width: 140px;
      text-align: center;
      font-size: 20px;
      font-weight: 700;
    }}

    .viewport {{
      flex: 1 1 auto;
      min-height: 0;
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 18px;
    }}

    .stage {{
      position: relative;
      width: min(calc(100vw - 36px), calc((100vh - 100px) * {aspect:.10f}));
      aspect-ratio: {width_px} / {height_px};
      max-height: calc(100vh - 100px);
      background: white;
      overflow: hidden;
      border-radius: 8px;
      box-shadow: 0 10px 28px rgba(0,0,0,0.35);
    }}

    .slide {{
      position: absolute;
      inset: 0;
      width: {width_px}px;
      height: {height_px}px;
      transform-origin: top left;
      background: white;
      overflow: hidden;
    }}

    .hotspot-btn:focus {{
      outline: 2px solid rgba(0,120,255,0.75);
      outline-offset: -2px;
    }}
  </style>
</head>
<body>
  <div class="app">
    <div class="topbar">
      <button id="prevBtn" type="button">◀ Prev</button>
      <div id="counter" class="counter">1 / 1</div>
      <button id="nextBtn" type="button">Next ▶</button>
    </div>

    <div class="viewport">
      <div id="stage" class="stage">
        {slides_joined}
      </div>
    </div>
  </div>

  <script>
    const TOTAL = {len(slides_items)};
    const BASE_W = {width_px};
    const BASE_H = {height_px};
    let current = 1;

    const stage = document.getElementById("stage");
    const counter = document.getElementById("counter");
    const prevBtn = document.getElementById("prevBtn");
    const nextBtn = document.getElementById("nextBtn");

    function updateScale() {{
      const w = stage.clientWidth;
      const h = stage.clientHeight;
      const scale = Math.min(w / BASE_W, h / BASE_H);

      document.querySelectorAll(".slide").forEach(slide => {{
        slide.style.transform = `scale(${{scale}})`;
      }});
    }}

    function render() {{
      current = Math.max(1, Math.min(TOTAL, current));
      document.querySelectorAll(".slide").forEach(slide => {{
        slide.style.display = "none";
      }});
      const active = document.getElementById(`slide-${{current}}`);
      if (active) active.style.display = "block";

      counter.textContent = `${{current}} / ${{TOTAL}}`;
      prevBtn.disabled = current === 1;
      nextBtn.disabled = current === TOTAL;

      updateScale();
    }}

    function goToSlide(n) {{
      current = n;
      render();
    }}

    prevBtn.addEventListener("click", function() {{
      if (current > 1) {{
        current -= 1;
        render();
      }}
    }});

    nextBtn.addEventListener("click", function() {{
      if (current < TOTAL) {{
        current += 1;
        render();
      }}
    }});

    document.addEventListener("keydown", function(e) {{
      if (e.key === "ArrowLeft" && current > 1) {{
        current -= 1;
        render();
      }} else if (e.key === "ArrowRight" && current < TOTAL) {{
        current += 1;
        render();
      }}
    }});

    window.addEventListener("resize", updateScale);

    render();
  </script>
</body>
</html>
"""


def build_scorm_api_js() -> str:
    return """var scormData = {
  "cmi.core.lesson_status": "not attempted",
  "cmi.core.score.raw": "",
  "cmi.core.student_name": "Learner",
  "cmi.core.student_id": "0001"
};

function LMSInitialize(param) { return "true"; }
function LMSFinish(param) { return "true"; }
function LMSGetValue(key) { return scormData[key] || ""; }
function LMSSetValue(key, value) { scormData[key] = value; return "true"; }
function LMSCommit(param) { return "true"; }
function LMSGetLastError() { return "0"; }
function LMSGetErrorString(code) { return "No error"; }
function LMSGetDiagnostic(code) { return "No diagnostic"; }

var API = {
  LMSInitialize: LMSInitialize,
  LMSFinish: LMSFinish,
  LMSGetValue: LMSGetValue,
  LMSSetValue: LMSSetValue,
  LMSCommit: LMSCommit,
  LMSGetLastError: LMSGetLastError,
  LMSGetErrorString: LMSGetErrorString,
  LMSGetDiagnostic: LMSGetDiagnostic
};
"""


def build_index_html(course_title: str) -> str:
    safe_title = html.escape(course_title)
    return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>{safe_title}</title>
  <script src="scorm_api.js"></script>
  <script>
    window.onload = function() {{
      try {{
        API.LMSInitialize("");
        API.LMSSetValue("cmi.core.lesson_status", "incomplete");
        API.LMSCommit("");
      }} catch (e) {{}}
    }};

    window.onbeforeunload = function() {{
      try {{
        API.LMSSetValue("cmi.core.lesson_status", "completed");
        API.LMSCommit("");
        API.LMSFinish("");
      }} catch (e) {{}}
    }};
  </script>
</head>
<frameset rows="100%" border="0">
  <frame src="player.html" name="content_frame" frameborder="0" />
</frameset>
</html>
"""


def build_manifest_xml(course_id: str, course_title: str) -> str:
    safe_id = re.sub(r"[^A-Za-z0-9_.-]", "_", course_id)
    safe_title = html.escape(course_title)

    return f"""<?xml version="1.0" encoding="UTF-8"?>
<manifest identifier="{safe_id}"
          version="1.0"
          xmlns="http://www.imsproject.org/xsd/imscp_rootv1p1p2"
          xmlns:adlcp="http://www.adlnet.org/xsd/adlcp_rootv1p2"
          xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
          xsi:schemaLocation="
            http://www.imsproject.org/xsd/imscp_rootv1p1p2 imscp_rootv1p1p2.xsd
            http://www.adlnet.org/xsd/adlcp_rootv1p2 adlcp_rootv1p2.xsd">
  <metadata>
    <schema>ADL SCORM</schema>
    <schemaversion>1.2</schemaversion>
  </metadata>
  <organizations default="ORG1">
    <organization identifier="ORG1">
      <title>{safe_title}</title>
      <item identifier="ITEM1" identifierref="RES1" isvisible="true">
        <title>{safe_title}</title>
      </item>
    </organization>
  </organizations>
  <resources>
    <resource identifier="RES1"
              type="webcontent"
              adlcp:scormtype="sco"
              href="index.html">
      <file href="index.html"/>
      <file href="player.html"/>
      <file href="scorm_api.js"/>
    </resource>
  </resources>
</manifest>
"""


def create_package(uploaded_file, course_title: str) -> tuple[bytes, list[str], str]:
    pptx_bytes = uploaded_file.getvalue()
    prs = Presentation(io.BytesIO(pptx_bytes))

    width_px = int(round(emu_to_px(prs.slide_width)))
    height_px = int(round(emu_to_px(prs.slide_height)))

    slides_items, warnings = parse_ppt(prs)

    player_html = build_player_html(course_title, width_px, height_px, slides_items)
    index_html = build_index_html(course_title)
    scorm_js = build_scorm_api_js()
    manifest_xml = build_manifest_xml("ppt_html_scorm_course", course_title)

    html_preview = player_html

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("player.html", player_html)
        zf.writestr("index.html", index_html)
        zf.writestr("scorm_api.js", scorm_js)
        zf.writestr("imsmanifest.xml", manifest_xml)

    buf.seek(0)
    return buf.read(), warnings, html_preview


# =========================
# UI
# =========================
uploaded = st.file_uploader("Upload PowerPoint (.pptx)", type=["pptx"])

col1, col2 = st.columns([1, 1])
with col1:
    package_name = st.text_input("Package name", value="ppt_html_scorm")
with col2:
    course_title = st.text_input("Course title", value="PPT Course")

st.markdown(
    """
This version is built to be **stable on Streamlit**.

It supports a controlled subset of PowerPoint features:
text boxes, pictures, simple shapes, basic fills/outlines, and internal slide links.
"""
)

if uploaded is not None:
    st.success(f"Loaded: {uploaded.name}")

    if st.button("Generate HTML + SCORM", type="primary"):
        try:
            zip_bytes, warnings, preview_html = create_package(
                uploaded,
                course_title.strip() or "PPT Course"
            )

            safe_name = sanitize_filename(package_name) + ".zip"

            st.success("Package created.")
            st.download_button(
                label="Download SCORM ZIP",
                data=zip_bytes,
                file_name=safe_name,
                mime="application/zip",
            )

            with st.expander("Preview rendered HTML", expanded=True):
                st.components.v1.html(preview_html, height=720, scrolling=False)

            if warnings:
                with st.expander("Unsupported or approximated items", expanded=False):
                    for w in warnings:
                        st.write(f"- {w}")
            else:
                st.info("No unsupported items detected.")

        except Exception as e:
            st.error(f"Failed to generate package: {e}")
