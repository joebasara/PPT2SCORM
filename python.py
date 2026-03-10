import html
import io
import json
import re
import uuid
import zipfile
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import streamlit as st
from pptx import Presentation
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE, MSO_SHAPE_TYPE

APP_TITLE = "PPTX to SCORM Publisher"


# =========================
# Helpers
# =========================

def safe_name(value: str) -> str:
    value = re.sub(r"[^A-Za-z0-9._-]+", "_", value.strip())
    return value.strip("._-") or "course"


def shape_bounds(shape) -> Tuple[int, int, int, int]:
    return int(shape.left), int(shape.top), int(shape.width), int(shape.height)


def emu_to_pct(value: int, total: int) -> float:
    if total == 0:
        return 0.0
    return (value / total) * 100.0


def color_to_hex(color_format) -> Optional[str]:
    try:
        rgb = color_format.rgb
        if rgb:
            return f"#{rgb}"
    except Exception:
        pass
    return None


def shape_fill_color(shape) -> Optional[str]:
    try:
        fill = shape.fill
        fore = fill.fore_color
        return color_to_hex(fore)
    except Exception:
        return None


def shape_line_color(shape) -> Optional[str]:
    try:
        line = shape.line
        color = line.color
        return color_to_hex(color)
    except Exception:
        return None


def shape_line_width_px(shape, slide_w_emu: int) -> float:
    try:
        width_emu = int(shape.line.width)
        return max(1.0, (width_emu / slide_w_emu) * 1280)
    except Exception:
        return 2.0


def extract_shape_external_link(shape) -> Optional[str]:
    try:
        click_action = getattr(shape, "click_action", None)
        if click_action is not None:
            hyperlink = getattr(click_action, "hyperlink", None)
            if hyperlink is not None:
                return getattr(hyperlink, "address", None)
    except Exception:
        pass
    return None


def detect_internal_link_target(slides, slide_idx_zero: int, shape) -> Optional[int]:
    try:
        click_action = getattr(shape, "click_action", None)
        if click_action is not None:
            target_slide = getattr(click_action, "target_slide", None)
            if target_slide is not None:
                for i, s in enumerate(slides, start=1):
                    if s == target_slide:
                        return i
    except Exception:
        pass

    try:
        hlink_click = shape.element.xpath(".//*[local-name()='hlinkClick']")
        if hlink_click:
            action = hlink_click[0].get("action", "").lower()
            if "firstslide" in action:
                return 1
            if "lastslide" in action:
                return len(slides)
            if "nextslide" in action:
                return min(len(slides), slide_idx_zero + 2)
            if "previousslide" in action:
                return max(1, slide_idx_zero)
    except Exception:
        pass

    return None


# =========================
# Text extraction
# =========================

def extract_text_element(shape) -> List[Dict[str, Any]]:
    text_frame = getattr(shape, "text_frame", None)
    if text_frame is None:
        return []

    x, y, w, h = shape_bounds(shape)
    parts = []

    for para in text_frame.paragraphs:
        run_parts = []
        for run in para.runs:
            txt = html.escape(run.text or "")
            if not txt:
                continue

            styles = []

            try:
                font = run.font
                if getattr(font, "bold", False):
                    styles.append("font-weight:700")
                if getattr(font, "italic", False):
                    styles.append("font-style:italic")
                if getattr(font, "underline", False):
                    styles.append("text-decoration:underline")
                if getattr(font, "size", None):
                    styles.append(f"font-size:{max(10, int(font.size.pt))}px")
                try:
                    fc = color_to_hex(font.color)
                    if fc:
                        styles.append(f"color:{fc}")
                except Exception:
                    pass
            except Exception:
                pass

            if styles:
                txt = f'<span style="{";".join(styles)}">{txt}</span>'

            url = None
            try:
                url = run.hyperlink.address
            except Exception:
                pass

            if url:
                txt = f'<a href="{html.escape(url)}" target="_blank" rel="noopener noreferrer">{txt}</a>'

            run_parts.append(txt)

        if run_parts:
            parts.append("<div>" + "".join(run_parts) + "</div>")

    if not parts:
        return []

    return [{
        "type": "text",
        "x": x,
        "y": y,
        "w": w,
        "h": h,
        "html": "".join(parts),
    }]


# =========================
# Shape to SVG extraction
# =========================

def extract_svg_shape(shape, slide_w: int, slide_h: int) -> Optional[Dict[str, Any]]:
    x, y, w, h = shape_bounds(shape)

    fill = shape_fill_color(shape) or "transparent"
    stroke = shape_line_color(shape) or "#000000"
    stroke_width = shape_line_width_px(shape, slide_w)

    shape_type = getattr(shape, "shape_type", None)

    # Straight lines / connectors
    if shape_type == MSO_SHAPE_TYPE.LINE:
        return {
            "kind": "line",
            "x1": x,
            "y1": y,
            "x2": x + w,
            "y2": y + h,
            "stroke": stroke,
            "stroke_width": stroke_width,
            "arrow_end": False,
        }

    auto_shape_type = getattr(shape, "auto_shape_type", None)

    if auto_shape_type == MSO_AUTO_SHAPE_TYPE.RECTANGLE:
        return {
            "kind": "rect",
            "x": x,
            "y": y,
            "w": w,
            "h": h,
            "fill": fill,
            "stroke": stroke,
            "stroke_width": stroke_width,
            "rx": 0,
            "ry": 0,
        }

    if auto_shape_type == MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE:
        return {
            "kind": "rect",
            "x": x,
            "y": y,
            "w": w,
            "h": h,
            "fill": fill,
            "stroke": stroke,
            "stroke_width": stroke_width,
            "rx": min(w, h) * 0.12,
            "ry": min(w, h) * 0.12,
        }

    if auto_shape_type == MSO_AUTO_SHAPE_TYPE.OVAL:
        return {
            "kind": "ellipse",
            "cx": x + w / 2,
            "cy": y + h / 2,
            "rx": w / 2,
            "ry": h / 2,
            "fill": fill,
            "stroke": stroke,
            "stroke_width": stroke_width,
        }

    # Basic arrows
    arrow_types = {
        MSO_AUTO_SHAPE_TYPE.RIGHT_ARROW,
        MSO_AUTO_SHAPE_TYPE.LEFT_ARROW,
        MSO_AUTO_SHAPE_TYPE.UP_ARROW,
        MSO_AUTO_SHAPE_TYPE.DOWN_ARROW,
        MSO_AUTO_SHAPE_TYPE.LEFT_RIGHT_ARROW,
    }
    if auto_shape_type in arrow_types:
        # Approximate arrow as polygon
        pts = []
        if auto_shape_type == MSO_AUTO_SHAPE_TYPE.RIGHT_ARROW:
            pts = [
                (x, y + h * 0.25),
                (x + w * 0.65, y + h * 0.25),
                (x + w * 0.65, y),
                (x + w, y + h * 0.5),
                (x + w * 0.65, y + h),
                (x + w * 0.65, y + h * 0.75),
                (x, y + h * 0.75),
            ]
        elif auto_shape_type == MSO_AUTO_SHAPE_TYPE.LEFT_ARROW:
            pts = [
                (x + w, y + h * 0.25),
                (x + w * 0.35, y + h * 0.25),
                (x + w * 0.35, y),
                (x, y + h * 0.5),
                (x + w * 0.35, y + h),
                (x + w * 0.35, y + h * 0.75),
                (x + w, y + h * 0.75),
            ]
        elif auto_shape_type == MSO_AUTO_SHAPE_TYPE.UP_ARROW:
            pts = [
                (x + w * 0.25, y + h),
                (x + w * 0.25, y + h * 0.35),
                (x, y + h * 0.35),
                (x + w * 0.5, y),
                (x + w, y + h * 0.35),
                (x + w * 0.75, y + h * 0.35),
                (x + w * 0.75, y + h),
            ]
        elif auto_shape_type == MSO_AUTO_SHAPE_TYPE.DOWN_ARROW:
            pts = [
                (x + w * 0.25, y),
                (x + w * 0.25, y + h * 0.65),
                (x, y + h * 0.65),
                (x + w * 0.5, y + h),
                (x + w, y + h * 0.65),
                (x + w * 0.75, y + h * 0.65),
                (x + w * 0.75, y),
            ]
        elif auto_shape_type == MSO_AUTO_SHAPE_TYPE.LEFT_RIGHT_ARROW:
            pts = [
                (x + w * 0.2, y + h * 0.25),
                (x + w * 0.8, y + h * 0.25),
                (x + w * 0.8, y),
                (x + w, y + h * 0.5),
                (x + w * 0.8, y + h),
                (x + w * 0.8, y + h * 0.75),
                (x + w * 0.2, y + h * 0.75),
                (x + w * 0.2, y + h),
                (x, y + h * 0.5),
                (x + w * 0.2, y),
            ]

        return {
            "kind": "polygon",
            "points": pts,
            "fill": fill,
            "stroke": stroke,
            "stroke_width": stroke_width,
        }

    return None


def svg_shape_to_markup(shape: Dict[str, Any], slide_w: int, slide_h: int) -> str:
    if shape["kind"] == "rect":
        return (
            f'<rect x="{emu_to_pct(shape["x"], slide_w)}%" '
            f'y="{emu_to_pct(shape["y"], slide_h)}%" '
            f'width="{emu_to_pct(shape["w"], slide_w)}%" '
            f'height="{emu_to_pct(shape["h"], slide_h)}%" '
            f'rx="{emu_to_pct(shape["rx"], slide_w)}%" '
            f'ry="{emu_to_pct(shape["ry"], slide_h)}%" '
            f'fill="{shape["fill"]}" '
            f'stroke="{shape["stroke"]}" '
            f'stroke-width="{shape["stroke_width"]}" />'
        )

    if shape["kind"] == "ellipse":
        return (
            f'<ellipse cx="{emu_to_pct(shape["cx"], slide_w)}%" '
            f'cy="{emu_to_pct(shape["cy"], slide_h)}%" '
            f'rx="{emu_to_pct(shape["rx"], slide_w)}%" '
            f'ry="{emu_to_pct(shape["ry"], slide_h)}%" '
            f'fill="{shape["fill"]}" '
            f'stroke="{shape["stroke"]}" '
            f'stroke-width="{shape["stroke_width"]}" />'
        )

    if shape["kind"] == "line":
        marker = ' marker-end="url(#arrowhead)"' if shape.get("arrow_end") else ""
        return (
            f'<line x1="{emu_to_pct(shape["x1"], slide_w)}%" '
            f'y1="{emu_to_pct(shape["y1"], slide_h)}%" '
            f'x2="{emu_to_pct(shape["x2"], slide_w)}%" '
            f'y2="{emu_to_pct(shape["y2"], slide_h)}%" '
            f'stroke="{shape["stroke"]}" '
            f'stroke-width="{shape["stroke_width"]}"{marker} />'
        )

    if shape["kind"] == "polygon":
        points = " ".join(
            f'{emu_to_pct(px, slide_w)},{emu_to_pct(py, slide_h)}'
            for px, py in shape["points"]
        )
        return (
            f'<polygon points="{points}" '
            f'fill="{shape["fill"]}" '
            f'stroke="{shape["stroke"]}" '
            f'stroke-width="{shape["stroke_width"]}" />'
        )

    return ""


# =========================
# Slide extraction
# =========================

def extract_course(prs: Presentation) -> Tuple[Dict[str, Any], Dict[str, bytes]]:
    media: Dict[str, bytes] = {}
    slides_data: List[Dict[str, Any]] = []

    for s_idx, slide in enumerate(prs.slides, start=1):
        text_elements: List[Dict[str, Any]] = []
        svg_shapes: List[Dict[str, Any]] = []
        hotspots: List[Dict[str, Any]] = []
        images: List[Dict[str, Any]] = []
        pic_no = 0

        for shape in slide.shapes:
            if getattr(shape, "has_text_frame", False):
                text_elements.extend(extract_text_element(shape))

            try:
                image = getattr(shape, "image", None)
                if image is not None:
                    pic_no += 1
                    ext = image.ext or "png"
                    filename = f"media/slide_{s_idx:03d}_img_{pic_no:02d}.{ext}"
                    media[filename] = image.blob
                    x, y, w, h = shape_bounds(shape)
                    images.append({
                        "type": "image",
                        "x": x,
                        "y": y,
                        "w": w,
                        "h": h,
                        "src": filename,
                    })
            except Exception:
                pass

            try:
                svg = extract_svg_shape(shape, int(prs.slide_width), int(prs.slide_height))
                if svg:
                    svg_shapes.append(svg)
            except Exception:
                pass

            try:
                x, y, w, h = shape_bounds(shape)

                external = extract_shape_external_link(shape)
                if external:
                    hotspots.append({
                        "x": x,
                        "y": y,
                        "w": w,
                        "h": h,
                        "kind": "external",
                        "url": external,
                    })

                target = detect_internal_link_target(prs.slides, s_idx - 1, shape)
                if target is not None:
                    hotspots.append({
                        "x": x,
                        "y": y,
                        "w": w,
                        "h": h,
                        "kind": "internal",
                        "target_slide": target,
                    })
            except Exception:
                pass

        slides_data.append({
            "index": s_idx,
            "svg_shapes": svg_shapes,
            "images": images,
            "text_elements": text_elements,
            "hotspots": hotspots,
        })

    course = {
        "slideWidthEmu": int(prs.slide_width),
        "slideHeightEmu": int(prs.slide_height),
        "slides": slides_data,
    }
    return course, media


# =========================
# SCORM / player generation
# =========================

def build_scorm_driver_js() -> str:
    return """
var scormAPI = null;
function findAPI(win){
  var tries = 0;
  while ((win.API == null) && (win.parent != null) && (win.parent != win)) {
    tries++;
    if (tries > 20) return null;
    win = win.parent;
  }
  return win.API;
}
function getAPI(){
  if (scormAPI == null) scormAPI = findAPI(window);
  return scormAPI;
}
function scormInitialize(){
  var api = getAPI();
  if (api) return api.LMSInitialize("");
  return false;
}
function scormTerminate(){
  var api = getAPI();
  if (api) return api.LMSFinish("");
  return false;
}
function scormGetValue(key){
  var api = getAPI();
  if (api) return api.LMSGetValue(key);
  return "";
}
function scormSetValue(key, value){
  var api = getAPI();
  if (api) return api.LMSSetValue(key, value);
  return false;
}
function scormCommit(){
  var api = getAPI();
  if (api) return api.LMSCommit("");
  return false;
}
"""


def build_player_html(title: str, course: Dict[str, Any]) -> str:
    data = json.dumps(course, ensure_ascii=False)

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>{html.escape(title)}</title>
<style>
body {{
  margin: 0;
  font-family: Arial, sans-serif;
  background: #111;
  color: #f1f1f1;
  display: flex;
  flex-direction: column;
  min-height: 100vh;
}}
header {{
  background: #1a1a1a;
  border-bottom: 1px solid #333;
  padding: 12px 16px;
  display: flex;
  justify-content: space-between;
  align-items: center;
  gap: 12px;
  flex-wrap: wrap;
}}
.controls {{
  display: flex;
  gap: 8px;
  flex-wrap: wrap;
}}
button, select {{
  background: #262626;
  color: #f1f1f1;
  border: 1px solid #444;
  border-radius: 8px;
  padding: 8px 12px;
}}
main {{
  flex: 1;
  display: grid;
  place-items: center;
  padding: 18px;
}}
.frame {{
  width: min(100%, 1280px);
  aspect-ratio: 16 / 9;
  position: relative;
  background: white;
  overflow: hidden;
  border-radius: 10px;
  box-shadow: 0 14px 40px rgba(0,0,0,.45);
}}
.svg-layer,
.image-layer,
.text-layer,
.hotspot-layer {{
  position: absolute;
  inset: 0;
}}
.svg-layer svg {{
  width: 100%;
  height: 100%;
  display: block;
}}
.el-text {{
  position: absolute;
  overflow: hidden;
  color: #111;
  line-height: 1.2;
  pointer-events: auto;
}}
.el-text a {{
  color: #0a58ca;
}}
.el-img {{
  position: absolute;
  object-fit: contain;
}}
.hotspot {{
  position: absolute;
  display: block;
  background: rgba(255,255,255,0);
  text-indent: -9999px;
  overflow: hidden;
}}
.footer {{
  background: #1a1a1a;
  border-top: 1px solid #333;
  padding: 10px 16px;
  color: #bbb;
  font-size: 13px;
}}
</style>
</head>
<body>
<header>
  <div>
    <div style="font-weight:700">{html.escape(title)}</div>
    <div id="counter" style="color:#bbb;font-size:14px">Slide 1</div>
  </div>
  <div class="controls">
    <button onclick="prevSlide()">Previous</button>
    <button onclick="nextSlide()">Next</button>
    <select id="jumpSelect" onchange="jumpSlide(this.value)"></select>
  </div>
</header>

<main>
  <div class="frame">
    <div class="svg-layer" id="svgLayer"></div>
    <div class="image-layer" id="imageLayer"></div>
    <div class="text-layer" id="textLayer"></div>
    <div class="hotspot-layer" id="hotspotLayer"></div>
  </div>
</main>

<div class="footer">SVG shapes + text layer + clickable hotspots.</div>

<script src="scormdriver.js"></script>
<script>
const course = {data};
let currentSlide = 1;

const svgLayer = document.getElementById("svgLayer");
const imageLayer = document.getElementById("imageLayer");
const textLayer = document.getElementById("textLayer");
const hotspotLayer = document.getElementById("hotspotLayer");
const counter = document.getElementById("counter");
const jumpSelect = document.getElementById("jumpSelect");

function pct(value, total) {{
  return (value / total) * 100;
}}

function populateJumpMenu() {{
  jumpSelect.innerHTML = "";
  for (const slide of course.slides) {{
    const opt = document.createElement("option");
    opt.value = slide.index;
    opt.textContent = `Slide ${{slide.index}}`;
    jumpSelect.appendChild(opt);
  }}
}}

function renderSVG(slide) {{
  let shapes = "";
  for (const raw of slide.svg_shapes || []) {{
    if (raw.kind === "rect") {{
      shapes += `<rect x="${{pct(raw.x, course.slideWidthEmu)}}%" y="${{pct(raw.y, course.slideHeightEmu)}}%" width="${{pct(raw.w, course.slideWidthEmu)}}%" height="${{pct(raw.h, course.slideHeightEmu)}}%" rx="${{pct(raw.rx || 0, course.slideWidthEmu)}}%" ry="${{pct(raw.ry || 0, course.slideHeightEmu)}}%" fill="${{raw.fill}}" stroke="${{raw.stroke}}" stroke-width="${{raw.stroke_width}}" />`;
    }} else if (raw.kind === "ellipse") {{
      shapes += `<ellipse cx="${{pct(raw.cx, course.slideWidthEmu)}}%" cy="${{pct(raw.cy, course.slideHeightEmu)}}%" rx="${{pct(raw.rx, course.slideWidthEmu)}}%" ry="${{pct(raw.ry, course.slideHeightEmu)}}%" fill="${{raw.fill}}" stroke="${{raw.stroke}}" stroke-width="${{raw.stroke_width}}" />`;
    }} else if (raw.kind === "line") {{
      shapes += `<line x1="${{pct(raw.x1, course.slideWidthEmu)}}%" y1="${{pct(raw.y1, course.slideHeightEmu)}}%" x2="${{pct(raw.x2, course.slideWidthEmu)}}%" y2="${{pct(raw.y2, course.slideHeightEmu)}}%" stroke="${{raw.stroke}}" stroke-width="${{raw.stroke_width}}" />`;
    }} else if (raw.kind === "polygon") {{
      const pts = (raw.points || []).map(p => `${{pct(p[0], course.slideWidthEmu)}},${{pct(p[1], course.slideHeightEmu)}}`).join(" ");
      shapes += `<polygon points="${{pts}}" fill="${{raw.fill}}" stroke="${{raw.stroke}}" stroke-width="${{raw.stroke_width}}" />`;
    }}
  }}

  svgLayer.innerHTML = `
    <svg viewBox="0 0 100 100" preserveAspectRatio="none">
      ${{shapes}}
    </svg>
  `;
}}

function renderImages(slide) {{
  imageLayer.innerHTML = "";
  for (const el of slide.images || []) {{
    const i = document.createElement("img");
    i.className = "el-img";
    i.src = el.src;
    i.style.left = pct(el.x, course.slideWidthEmu) + "%";
    i.style.top = pct(el.y, course.slideHeightEmu) + "%";
    i.style.width = pct(el.w, course.slideWidthEmu) + "%";
    i.style.height = pct(el.h, course.slideHeightEmu) + "%";
    imageLayer.appendChild(i);
  }}
}}

function renderText(slide) {{
  textLayer.innerHTML = "";
  for (const el of slide.text_elements || []) {{
    const d = document.createElement("div");
    d.className = "el-text";
    d.style.left = pct(el.x, course.slideWidthEmu) + "%";
    d.style.top = pct(el.y, course.slideHeightEmu) + "%";
    d.style.width = pct(el.w, course.slideWidthEmu) + "%";
    d.style.height = pct(el.h, course.slideHeightEmu) + "%";
    d.innerHTML = el.html;
    textLayer.appendChild(d);
  }}
}}

function renderHotspots(slide) {{
  hotspotLayer.innerHTML = "";
  for (const link of slide.hotspots || []) {{
    const a = document.createElement("a");
    a.className = "hotspot";
    a.style.left = pct(link.x, course.slideWidthEmu) + "%";
    a.style.top = pct(link.y, course.slideHeightEmu) + "%";
    a.style.width = pct(link.w, course.slideWidthEmu) + "%";
    a.style.height = pct(link.h, course.slideHeightEmu) + "%";

    if (link.kind === "external" && link.url) {{
      a.href = link.url;
      a.target = "_blank";
      a.rel = "noopener noreferrer";
      a.textContent = link.url;
    }} else if (link.kind === "internal" && link.target_slide) {{
      a.href = "#";
      a.textContent = "Go to slide " + link.target_slide;
      a.addEventListener("click", function(e) {{
        e.preventDefault();
        goToSlide(Number(link.target_slide));
      }});
    }} else {{
      continue;
    }}

    hotspotLayer.appendChild(a);
  }}
}}

function setScormState() {{
  if (!window.scormSetValue) return;
  try {{
    window.scormSetValue("cmi.core.lesson_location", String(currentSlide));
    const progress = Math.round((currentSlide / course.slides.length) * 100);
    window.scormSetValue("cmi.core.score.raw", String(progress));
    window.scormSetValue("cmi.core.lesson_status", currentSlide >= course.slides.length ? "completed" : "incomplete");
    window.scormCommit();
  }} catch (e) {{}}
}}

function goToSlide(num) {{
  if (num < 1 || num > course.slides.length) return;
  currentSlide = num;
  const slide = course.slides[currentSlide - 1];

  renderSVG(slide);
  renderImages(slide);
  renderText(slide);
  renderHotspots(slide);

  counter.textContent = `Slide ${{currentSlide}} of ${{course.slides.length}}`;
  jumpSelect.value = String(currentSlide);
  setScormState();
}}

function nextSlide() {{
  if (currentSlide < course.slides.length) goToSlide(currentSlide + 1);
}}

function prevSlide() {{
  if (currentSlide > 1) goToSlide(currentSlide - 1);
}}

function jumpSlide(value) {{
  const num = Number(value);
  if (!Number.isNaN(num)) goToSlide(num);
}}

document.addEventListener("keydown", function(e) {{
  if (e.key === "ArrowRight") nextSlide();
  if (e.key === "ArrowLeft") prevSlide();
}});

window.addEventListener("load", function() {{
  populateJumpMenu();
  if (window.scormInitialize) {{
    try {{
      window.scormInitialize();
      const saved = window.scormGetValue("cmi.core.lesson_location");
      const n = parseInt(saved || "1", 10);
      if (!Number.isNaN(n) && n >= 1 && n <= course.slides.length) currentSlide = n;
    }} catch (e) {{}}
  }}
  goToSlide(currentSlide);
}});

window.addEventListener("beforeunload", function() {{
  if (window.scormTerminate) {{
    try {{ window.scormTerminate(); }} catch (e) {{}}
  }}
}});
</script>
</body>
</html>"""


def build_manifest_xml(course_id: str, title: str, media_files: List[str]) -> str:
    media_xml = "\n".join([f'      <file href="{name}" />' for name in media_files])

    return f"""<?xml version="1.0" encoding="UTF-8"?>
<manifest identifier="{course_id}"
    version="1.0"
    xmlns="http://www.imsproject.org/xsd/imscp_rootv1p1p2"
    xmlns:adlcp="http://www.adlnet.org/xsd/adlcp_rootv1p2"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
    xsi:schemaLocation="http://www.imsproject.org/xsd/imscp_rootv1p1p2 imscp_rootv1p1p2.xsd
    http://www.adlnet.org/xsd/adlcp_rootv1p2 adlcp_rootv1p2.xsd">
  <metadata>
    <schema>ADL SCORM</schema>
    <schemaversion>1.2</schemaversion>
  </metadata>
  <organizations default="ORG1">
    <organization identifier="ORG1">
      <title>{html.escape(title)}</title>
      <item identifier="ITEM1" identifierref="RES1">
        <title>{html.escape(title)}</title>
      </item>
    </organization>
  </organizations>
  <resources>
    <resource identifier="RES1" type="webcontent" adlcp:scormtype="sco" href="index_lms.html">
      <file href="index_lms.html" />
      <file href="scormdriver.js" />
{media_xml}
    </resource>
  </resources>
</manifest>"""


def build_scorm_zip(pptx_bytes: bytes, pptx_name: str, course_title: str):
    prs = Presentation(io.BytesIO(pptx_bytes))
    course, media = extract_course(prs)

    mem = io.BytesIO()
    with zipfile.ZipFile(mem, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, blob in media.items():
            zf.writestr(name, blob)

        zf.writestr("index_lms.html", build_player_html(course_title, course))
        zf.writestr("scormdriver.js", build_scorm_driver_js())
        zf.writestr(
            "imsmanifest.xml",
            build_manifest_xml(
                f"{safe_name(course_title)}_{uuid.uuid4().hex[:8]}",
                course_title,
                sorted(media.keys()),
            ),
        )

    mem.seek(0)
    out_name = f"{safe_name(Path(pptx_name).stem)}_scorm12.zip"
    summary = {
        "slides": len(prs.slides),
        "media": len(media),
    }
    return mem.read(), out_name, summary


# =========================
# Streamlit UI
# =========================

st.set_page_config(page_title=APP_TITLE, layout="centered")
st.title(APP_TITLE)
st.caption("Pure-Python MVP: PPTX → HTML/SVG slide player → SCORM 1.2 ZIP")

with st.expander("What this version supports"):
    st.markdown(
        """
- Slide-by-slide navigation
- External hyperlinks in text and clickable shapes
- Simple internal slide jumps where detectable
- Embedded pictures
- Basic SVG rendering for common PowerPoint shapes
- SCORM 1.2 packaging for Blackboard testing

Still limited:
- Not full PowerPoint fidelity
- No animations or transitions
- SmartArt/charts/freeform diagrams may not render correctly
- Complex grouped objects may not be preserved
        """
    )

course_title = st.text_input("Course title", value="My SCORM Course")
uploaded = st.file_uploader("Upload PPTX", type=["pptx"])

if st.button("Publish SCORM", type="primary", use_container_width=True):
    if not uploaded:
        st.error("Please upload a .pptx file first.")
    else:
        try:
            zip_bytes, zip_name, summary = build_scorm_zip(
                uploaded.getvalue(),
                uploaded.name,
                course_title,
            )
            st.success("SCORM package created.")
            st.write(f"Slides detected: {summary['slides']}")
            st.write(f"Embedded pictures copied: {summary['media']}")
            st.download_button(
                "Download SCORM ZIP",
                data=zip_bytes,
                file_name=zip_name,
                mime="application/zip",
                use_container_width=True,
            )
        except Exception as e:
            st.error(f"Failed to publish SCORM package: {e}")

st.divider()
st.subheader("requirements.txt")
st.code("streamlit\npython-pptx\nlxml\nPillow\n")
