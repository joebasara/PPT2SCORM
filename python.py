import html
import io
import json
import math
import re
import uuid
import zipfile
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import streamlit as st
from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE, MSO_SHAPE_TYPE

APP_TITLE = "PPTX to SCORM Publisher"
CANVAS_W = 1600


# ============================================================
# Helpers
# ============================================================

def safe_name(value: str) -> str:
    value = re.sub(r"[^A-Za-z0-9._-]+", "_", value.strip())
    return value.strip("._-") or "course"


def get_auto_shape_type(shape):
    try:
        return shape.auto_shape_type
    except Exception:
        return None


def shape_bounds(shape, offset_x: int = 0, offset_y: int = 0) -> Tuple[int, int, int, int]:
    return (
        int(getattr(shape, "left", 0)) + offset_x,
        int(getattr(shape, "top", 0)) + offset_y,
        int(getattr(shape, "width", 0)),
        int(getattr(shape, "height", 0)),
    )


def color_to_rgb(color_obj, default=(0, 0, 0)):
    try:
        rgb = getattr(color_obj, "rgb", None)
        if rgb:
            s = str(rgb)
            return tuple(int(s[i:i+2], 16) for i in (0, 2, 4))
    except Exception:
        pass
    return default


def shape_fill_rgb(shape, default=(255, 255, 255, 0)):
    try:
        fill = shape.fill
        fc = fill.fore_color
        rgb = getattr(fc, "rgb", None)
        if rgb:
            s = str(rgb)
            return tuple(int(s[i:i+2], 16) for i in (0, 2, 4)) + (255,)
    except Exception:
        pass
    return default


def shape_line_rgb(shape, default=(0, 0, 0, 255)):
    try:
        line = shape.line
        c = line.color
        rgb = getattr(c, "rgb", None)
        if rgb:
            s = str(rgb)
            return tuple(int(s[i:i+2], 16) for i in (0, 2, 4)) + (255,)
    except Exception:
        pass
    return default


def shape_line_width_px(shape, scale: float) -> int:
    try:
        width_emu = float(shape.line.width)
        px = int(max(1, round((width_emu / 12700.0) * scale)))
        return px
    except Exception:
        return max(1, int(round(2 * scale)))


def emu_to_px(value: float, scale: float) -> int:
    return int(round(value * scale))


def fit_text_lines(draw: ImageDraw.ImageDraw, text: str, font, max_width: int) -> list[str]:
    words = text.split()
    if not words:
        return []

    lines = []
    current = words[0]
    for w in words[1:]:
        trial = current + " " + w
        if draw.textlength(trial, font=font) <= max_width:
            current = trial
        else:
            lines.append(current)
            current = w
    lines.append(current)
    return lines


def get_font(size_px: int, bold: bool = False):
    candidates = (
        ["arialbd.ttf", "Arial Bold.ttf", "DejaVuSans-Bold.ttf"]
        if bold
        else ["arial.ttf", "Arial.ttf", "DejaVuSans.ttf"]
    )
    for name in candidates:
        try:
            return ImageFont.truetype(name, size_px)
        except Exception:
            continue
    return ImageFont.load_default()


# ============================================================
# XML / image helpers
# ============================================================

def local_name(tag: str) -> str:
    return tag.split("}")[-1] if "}" in tag else tag


def find_descendant_by_localname(el, name: str):
    for node in el.iter():
        if local_name(node.tag) == name:
            return node
    return None


def find_descendants_by_localname(el, name: str):
    out = []
    for node in el.iter():
        if local_name(node.tag) == name:
            out.append(node)
    return out


def get_crop_rect_from_shape(shape) -> Tuple[float, float, float, float]:
    """
    Returns crop fractions (left, top, right, bottom) in range 0..1
    """
    try:
        src_rect = find_descendant_by_localname(shape.element, "srcRect")
        if src_rect is None:
            return 0.0, 0.0, 0.0, 0.0

        l = float(src_rect.get("l", "0")) / 100000.0
        t = float(src_rect.get("t", "0")) / 100000.0
        r = float(src_rect.get("r", "0")) / 100000.0
        b = float(src_rect.get("b", "0")) / 100000.0
        return l, t, r, b
    except Exception:
        return 0.0, 0.0, 0.0, 0.0


def apply_crop_to_image(img: Image.Image, crop_rect: Tuple[float, float, float, float]) -> Image.Image:
    l, t, r, b = crop_rect
    if max(l, t, r, b) <= 0:
        return img

    w, h = img.size
    left = int(round(w * l))
    top = int(round(h * t))
    right = int(round(w * (1 - r)))
    bottom = int(round(h * (1 - b)))

    left = max(0, min(left, w - 1))
    top = max(0, min(top, h - 1))
    right = max(left + 1, min(right, w))
    bottom = max(top + 1, min(bottom, h))

    return img.crop((left, top, right, bottom))


def get_image_blob_from_rid(shape_part, rid: str) -> Optional[bytes]:
    try:
        rel = shape_part.related_part(rid)
        blob = getattr(rel, "blob", None)
        if blob:
            return blob
    except Exception:
        pass
    return None


def get_shape_fill_image_blob(shape) -> Optional[bytes]:
    try:
        blip = find_descendant_by_localname(shape.element, "blip")
        if blip is None:
            return None

        rid = None
        for k, v in blip.attrib.items():
            if k.endswith("embed"):
                rid = v
                break
        if not rid:
            return None

        return get_image_blob_from_rid(shape.part, rid)
    except Exception:
        return None


def get_slide_background_image_blob(slide) -> Optional[bytes]:
    try:
        bg = slide.background
        fill = getattr(bg, "fill", None)
        _ = fill  # keep for sanity, not strictly needed
    except Exception:
        pass

    try:
        blip = find_descendant_by_localname(slide._element, "blip")
        if blip is None:
            return None

        rid = None
        for k, v in blip.attrib.items():
            if k.endswith("embed"):
                rid = v
                break
        if not rid:
            return None

        return get_image_blob_from_rid(slide.part, rid)
    except Exception:
        return None


# ============================================================
# Hyperlinks / hotspots
# ============================================================

def extract_shape_external_link(shape) -> Optional[str]:
    try:
        click_action = getattr(shape, "click_action", None)
        if click_action is not None:
            hyperlink = getattr(click_action, "hyperlink", None)
            if hyperlink is not None:
                address = getattr(hyperlink, "address", None)
                if address:
                    return address
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
        hlink_click = find_descendant_by_localname(shape.element, "hlinkClick")
        if hlink_click is not None:
            action = (hlink_click.get("action") or "").lower()
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


def hotspot_geometry_for_shape(shape, offset_x=0, offset_y=0) -> Dict[str, Any]:
    x, y, w, h = shape_bounds(shape, offset_x, offset_y)
    auto = get_auto_shape_type(shape)
    shape_type = getattr(shape, "shape_type", None)

    if shape_type == MSO_SHAPE_TYPE.LINE:
        return {
            "geom": "line",
            "x1": x,
            "y1": y,
            "x2": x + w,
            "y2": y + h,
            "x": x,
            "y": y,
            "w": w,
            "h": h,
        }

    if auto == MSO_AUTO_SHAPE_TYPE.OVAL:
        return {"geom": "ellipse", "x": x, "y": y, "w": w, "h": h}

    if auto in (MSO_AUTO_SHAPE_TYPE.RECTANGLE, MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE):
        return {"geom": "rect", "x": x, "y": y, "w": w, "h": h}

    return {"geom": "rect", "x": x, "y": y, "w": w, "h": h}


# ============================================================
# Drawing
# ============================================================

def draw_line(draw, x1, y1, x2, y2, fill, width):
    draw.line((x1, y1, x2, y2), fill=fill, width=width)


def draw_rect(draw, x, y, w, h, fill, outline, width, radius=0):
    x2, y2 = x + w, y + h
    if radius > 0:
        draw.rounded_rectangle((x, y, x2, y2), radius=radius, fill=fill, outline=outline, width=width)
    else:
        draw.rectangle((x, y, x2, y2), fill=fill, outline=outline, width=width)


def draw_ellipse(draw, x, y, w, h, fill, outline, width):
    draw.ellipse((x, y, x + w, y + h), fill=fill, outline=outline, width=width)


def arrow_polygon_points(auto, x, y, w, h):
    if auto == MSO_AUTO_SHAPE_TYPE.RIGHT_ARROW:
        return [
            (x, y + h * 0.25),
            (x + w * 0.65, y + h * 0.25),
            (x + w * 0.65, y),
            (x + w, y + h * 0.5),
            (x + w * 0.65, y + h),
            (x + w * 0.65, y + h * 0.75),
            (x, y + h * 0.75),
        ]
    if auto == MSO_AUTO_SHAPE_TYPE.LEFT_ARROW:
        return [
            (x + w, y + h * 0.25),
            (x + w * 0.35, y + h * 0.25),
            (x + w * 0.35, y),
            (x, y + h * 0.5),
            (x + w * 0.35, y + h),
            (x + w * 0.35, y + h * 0.75),
            (x + w, y + h * 0.75),
        ]
    if auto == MSO_AUTO_SHAPE_TYPE.UP_ARROW:
        return [
            (x + w * 0.25, y + h),
            (x + w * 0.25, y + h * 0.35),
            (x, y + h * 0.35),
            (x + w * 0.5, y),
            (x + w, y + h * 0.35),
            (x + w * 0.75, y + h * 0.35),
            (x + w * 0.75, y + h),
        ]
    if auto == MSO_AUTO_SHAPE_TYPE.DOWN_ARROW:
        return [
            (x + w * 0.25, y),
            (x + w * 0.25, y + h * 0.65),
            (x, y + h * 0.65),
            (x + w * 0.5, y + h),
            (x + w, y + h * 0.65),
            (x + w * 0.75, y + h * 0.65),
            (x + w * 0.75, y),
        ]
    return None


def draw_text_box(draw, shape, scale: float, offset_x=0, offset_y=0):
    text_frame = getattr(shape, "text_frame", None)
    if text_frame is None:
        return

    x_emu, y_emu, w_emu, h_emu = shape_bounds(shape, offset_x, offset_y)
    x = emu_to_px(x_emu, scale)
    y = emu_to_px(y_emu, scale)
    w = max(4, emu_to_px(w_emu, scale))
    h = max(4, emu_to_px(h_emu, scale))

    paragraphs = []
    for para in text_frame.paragraphs:
        txt = "".join(run.text or "" for run in para.runs).strip()
        if txt:
            paragraphs.append((para, txt))

    if not paragraphs:
        return

    first_run = None
    for para, _ in paragraphs:
        if para.runs:
            first_run = para.runs[0]
            break

    font_size = 22
    font_bold = False
    font_fill = (0, 0, 0, 255)

    if first_run:
        try:
            fs = getattr(first_run.font, "size", None)
            if fs is not None:
                font_size = max(10, int(fs.pt * scale))
        except Exception:
            pass
        try:
            font_bold = bool(getattr(first_run.font, "bold", False))
        except Exception:
            pass
        try:
            font_fill = color_to_rgb(first_run.font.color, default=(0, 0, 0)) + (255,)
        except Exception:
            pass

    font = get_font(font_size, bold=font_bold)
    line_gap = max(2, int(font_size * 0.25))
    all_lines = []
    line_aligns = []

    for para, txt in paragraphs:
        max_width = max(10, w - 8)
        lines = fit_text_lines(draw, txt, font, max_width)
        if not lines:
            lines = [txt]
        align = "left"
        try:
            if para.alignment is not None:
                name = str(para.alignment).lower()
                if "center" in name:
                    align = "center"
                elif "right" in name:
                    align = "right"
        except Exception:
            pass
        for line in lines:
            all_lines.append(line)
            line_aligns.append(align)

    bbox = draw.textbbox((0, 0), "Ag", font=font)
    line_h = max(1, bbox[3] - bbox[1])
    total_h = len(all_lines) * line_h + max(0, len(all_lines) - 1) * line_gap

    v_align = "top"
    try:
        name = str(text_frame.vertical_anchor).lower()
        if "middle" in name:
            v_align = "middle"
        elif "bottom" in name:
            v_align = "bottom"
    except Exception:
        pass

    if v_align == "middle":
        cy = y + (h - total_h) // 2
    elif v_align == "bottom":
        cy = y + h - total_h - 4
    else:
        cy = y + 4

    for line, align in zip(all_lines, line_aligns):
        tw = draw.textlength(line, font=font)
        if align == "center":
            tx = x + (w - tw) / 2
        elif align == "right":
            tx = x + w - tw - 4
        else:
            tx = x + 4
        draw.text((tx, cy), line, font=font, fill=font_fill)
        cy += line_h + line_gap


def draw_table(draw, shape, scale: float, offset_x=0, offset_y=0):
    try:
        table = shape.table
    except Exception:
        return

    x_emu, y_emu, w_emu, h_emu = shape_bounds(shape, offset_x, offset_y)
    x = emu_to_px(x_emu, scale)
    y = emu_to_px(y_emu, scale)
    w = max(4, emu_to_px(w_emu, scale))
    h = max(4, emu_to_px(h_emu, scale))

    try:
        col_widths = [emu_to_px(col.width, scale) for col in table.columns]
    except Exception:
        col_count = len(table.columns)
        col_widths = [w // max(1, col_count)] * max(1, col_count)

    try:
        row_heights = [emu_to_px(row.height, scale) for row in table.rows]
    except Exception:
        row_count = len(table.rows)
        row_heights = [h // max(1, row_count)] * max(1, row_count)

    cy = y
    for r_idx, row in enumerate(table.rows):
        rh = row_heights[r_idx] if r_idx < len(row_heights) else max(20, h // max(1, len(table.rows)))
        cx = x
        for c_idx, cell in enumerate(row.cells):
            cw = col_widths[c_idx] if c_idx < len(col_widths) else max(30, w // max(1, len(table.columns)))
            fill = (255, 255, 255, 0)
            try:
                fill = color_to_rgb(cell.fill.fore_color, default=(255, 255, 255)) + (255,)
            except Exception:
                pass

            draw.rectangle((cx, cy, cx + cw, cy + rh), fill=fill, outline=(85, 85, 85, 255), width=1)

            text = cell.text.strip()
            if text:
                font = get_font(max(11, int(16 * scale)))
                lines = fit_text_lines(draw, text, font, max(10, cw - 8))
                bbox = draw.textbbox((0, 0), "Ag", font=font)
                lh = max(1, bbox[3] - bbox[1])
                total_h = len(lines) * lh
                ty = cy + max(2, (rh - total_h) // 2)
                for line in lines:
                    tw = draw.textlength(line, font=font)
                    tx = cx + max(4, (cw - tw) / 2)
                    draw.text((tx, ty), line, font=font, fill=(0, 0, 0, 255))
                    ty += lh
            cx += cw
        cy += rh


def draw_shape_object(draw, shape, scale: float, offset_x=0, offset_y=0):
    x_emu, y_emu, w_emu, h_emu = shape_bounds(shape, offset_x, offset_y)
    x = emu_to_px(x_emu, scale)
    y = emu_to_px(y_emu, scale)
    w = max(1, emu_to_px(w_emu, scale))
    h = max(1, emu_to_px(h_emu, scale))

    fill = shape_fill_rgb(shape)
    outline = shape_line_rgb(shape)
    width = shape_line_width_px(shape, scale)
    auto = get_auto_shape_type(shape)
    shape_type = getattr(shape, "shape_type", None)

    if shape_type == MSO_SHAPE_TYPE.LINE:
        draw_line(draw, x, y, x + w, y + h, outline, width)
        return

    if auto == MSO_AUTO_SHAPE_TYPE.RECTANGLE:
        draw_rect(draw, x, y, w, h, fill, outline, width, radius=0)
        return

    if auto == MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE:
        draw_rect(draw, x, y, w, h, fill, outline, width, radius=max(4, min(w, h) // 8))
        return

    if auto == MSO_AUTO_SHAPE_TYPE.OVAL:
        draw_ellipse(draw, x, y, w, h, fill, outline, width)
        return

    pts = arrow_polygon_points(auto, x, y, w, h)
    if pts:
        draw.polygon(pts, fill=fill, outline=outline)
        return

    if auto is not None:
        draw_rect(draw, x, y, w, h, fill, outline, width, radius=0)


def paste_image_into_box(canvas: Image.Image, img: Image.Image, x: int, y: int, w: int, h: int):
    if w <= 0 or h <= 0:
        return
    img = img.resize((w, h))
    if img.mode != "RGBA":
        img = img.convert("RGBA")
    canvas.alpha_composite(img, (x, y))


def paste_picture(canvas: Image.Image, shape, scale: float, offset_x=0, offset_y=0):
    try:
        image = shape.image
        if image is None:
            return
        x_emu, y_emu, w_emu, h_emu = shape_bounds(shape, offset_x, offset_y)
        x = emu_to_px(x_emu, scale)
        y = emu_to_px(y_emu, scale)
        w = max(1, emu_to_px(w_emu, scale))
        h = max(1, emu_to_px(h_emu, scale))

        img = Image.open(io.BytesIO(image.blob)).convert("RGBA")
        img = apply_crop_to_image(img, get_crop_rect_from_shape(shape))
        paste_image_into_box(canvas, img, x, y, w, h)
    except Exception:
        pass


def paste_shape_fill_image(canvas: Image.Image, shape, scale: float, offset_x=0, offset_y=0):
    try:
        blob = get_shape_fill_image_blob(shape)
        if not blob:
            return False
        x_emu, y_emu, w_emu, h_emu = shape_bounds(shape, offset_x, offset_y)
        x = emu_to_px(x_emu, scale)
        y = emu_to_px(y_emu, scale)
        w = max(1, emu_to_px(w_emu, scale))
        h = max(1, emu_to_px(h_emu, scale))

        img = Image.open(io.BytesIO(blob)).convert("RGBA")
        paste_image_into_box(canvas, img, x, y, w, h)
        return True
    except Exception:
        return False


def paste_slide_background(canvas: Image.Image, slide, prs: Presentation):
    try:
        blob = get_slide_background_image_blob(slide)
        if not blob:
            return False
        img = Image.open(io.BytesIO(blob)).convert("RGBA")
        paste_image_into_box(canvas, img, 0, 0, canvas.width, canvas.height)
        return True
    except Exception:
        return False


def render_shape_recursive(canvas: Image.Image, draw: ImageDraw.ImageDraw, shape, scale: float, offset_x=0, offset_y=0):
    shape_type = getattr(shape, "shape_type", None)

    if shape_type == MSO_SHAPE_TYPE.GROUP:
        gx, gy, _, _ = shape_bounds(shape, offset_x, offset_y)
        try:
            for child in shape.shapes:
                render_shape_recursive(canvas, draw, child, scale, gx, gy)
        except Exception:
            pass
        return

    used_fill_img = False
    try:
        used_fill_img = paste_shape_fill_image(canvas, shape, scale, offset_x, offset_y)
    except Exception:
        used_fill_img = False

    try:
        if getattr(shape, "image", None) is not None:
            paste_picture(canvas, shape, scale, offset_x, offset_y)
    except Exception:
        pass

    try:
        draw_table(draw, shape, scale, offset_x, offset_y)
    except Exception:
        pass

    try:
        draw_shape_object(draw, shape, scale, offset_x, offset_y)
    except Exception:
        pass

    try:
        if getattr(shape, "has_text_frame", False):
            draw_text_box(draw, shape, scale, offset_x, offset_y)
    except Exception:
        pass


def render_slide_to_png(slide, prs: Presentation) -> bytes:
    slide_w = int(prs.slide_width)
    slide_h = int(prs.slide_height)
    scale = CANVAS_W / slide_w
    canvas_h = max(1, int(round(slide_h * scale)))

    canvas = Image.new("RGBA", (CANVAS_W, canvas_h), (255, 255, 255, 255))
    draw = ImageDraw.Draw(canvas)

    paste_slide_background(canvas, slide, prs)

    for shape in slide.shapes:
        render_shape_recursive(canvas, draw, shape, scale)

    out = io.BytesIO()
    canvas.save(out, format="PNG")
    return out.getvalue()


# ============================================================
# Extract course package data
# ============================================================

def collect_hotspots_recursive(shape, prs, slide_idx_zero: int, hotspots: List[Dict[str, Any]], offset_x=0, offset_y=0):
    shape_type = getattr(shape, "shape_type", None)

    if shape_type == MSO_SHAPE_TYPE.GROUP:
        gx, gy, _, _ = shape_bounds(shape, offset_x, offset_y)
        try:
            for child in shape.shapes:
                collect_hotspots_recursive(child, prs, slide_idx_zero, hotspots, gx, gy)
        except Exception:
            pass
        return

    try:
        external = extract_shape_external_link(shape)
        internal = detect_internal_link_target(prs.slides, slide_idx_zero, shape)
        if external or internal:
            geom = hotspot_geometry_for_shape(shape, offset_x, offset_y)
            if external:
                geom = {**geom, "kind": "external", "url": external}
            else:
                geom = {**geom, "kind": "internal", "target_slide": internal}
            hotspots.append(geom)
    except Exception:
        pass


def extract_course(prs: Presentation) -> Tuple[Dict[str, Any], Dict[str, bytes]]:
    media: Dict[str, bytes] = {}
    slides_out = []

    for s_idx, slide in enumerate(prs.slides, start=1):
        slide_img_name = f"slides/slide_{s_idx:03d}.png"
        media[slide_img_name] = render_slide_to_png(slide, prs)

        hotspots: List[Dict[str, Any]] = []
        for shape in slide.shapes:
            collect_hotspots_recursive(shape, prs, s_idx - 1, hotspots, 0, 0)

        slides_out.append({
            "index": s_idx,
            "image": slide_img_name,
            "hotspots": hotspots,
        })

    course = {
        "slideWidthEmu": int(prs.slide_width),
        "slideHeightEmu": int(prs.slide_height),
        "slides": slides_out,
    }
    return course, media


# ============================================================
# SCORM / player
# ============================================================

def build_scorm_driver_js() -> str:
    return """
var scormAPI = null;

function findAPI(win) {
  var tries = 0;
  while ((win.API == null) && (win.parent != null) && (win.parent != win)) {
    tries++;
    if (tries > 20) return null;
    win = win.parent;
  }
  return win.API;
}

function getAPI() {
  if (scormAPI == null) scormAPI = findAPI(window);
  return scormAPI;
}

function scormInitialize() {
  var api = getAPI();
  if (api) return api.LMSInitialize("");
  return false;
}

function scormTerminate() {
  var api = getAPI();
  if (api) return api.LMSFinish("");
  return false;
}

function scormGetValue(key) {
  var api = getAPI();
  if (api) return api.LMSGetValue(key);
  return "";
}

function scormSetValue(key, value) {
  var api = getAPI();
  if (api) return api.LMSSetValue(key, value);
  return false;
}

function scormCommit() {
  var api = getAPI();
  if (api) return api.LMSCommit("");
  return false;
}
"""


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


def build_player_html(title: str, course: Dict[str, Any], show_nav: bool) -> str:
    data = json.dumps(course, ensure_ascii=False)
    sidebar_style = "" if show_nav else "display:none;"
    bottombar_style = "display:flex;" if show_nav else "display:none;"
    app_padding_left = "130px" if show_nav else "0"

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>{html.escape(title)}</title>
<style>
* {{
  box-sizing: border-box;
}}
body {{
  margin: 0;
  font-family: Arial, sans-serif;
  background: #111;
  color: #f1f1f1;
  min-height: 100vh;
}}
.app {{
  display: flex;
  min-height: 100vh;
}}
.sidebar {{
  width: 130px;
  background: #1a1a1a;
  border-right: 1px solid #333;
  padding: 16px 10px;
  {sidebar_style}
}}
.sidebar label {{
  display: block;
  font-size: 13px;
  color: #bbb;
  margin-bottom: 8px;
}}
.sidebar select {{
  width: 100%;
  height: calc(100vh - 40px);
  min-height: 260px;
  background: #262626;
  color: #f1f1f1;
  border: 1px solid #444;
  border-radius: 8px;
  padding: 8px;
}}
.main {{
  flex: 1;
  display: flex;
  flex-direction: column;
}}
.header {{
  background: #1a1a1a;
  border-bottom: 1px solid #333;
  padding: 12px 16px;
}}
.header-title {{
  font-weight: 700;
}}
.header-sub {{
  color: #bbb;
  font-size: 14px;
  margin-top: 4px;
}}
.stage-wrap {{
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
.slide-image {{
  position: absolute;
  inset: 0;
  width: 100%;
  height: 100%;
  object-fit: contain;
}}
.hotspot-layer {{
  position: absolute;
  inset: 0;
}}
.hotspot {{
  position: absolute;
  display: block;
  background: rgba(0,0,0,0);
  pointer-events: auto;
}}
.hotspot.ellipse {{
  border-radius: 50%;
}}
.bottom-bar {{
  background: #1a1a1a;
  border-top: 1px solid #333;
  padding: 12px 16px;
  justify-content: center;
  gap: 12px;
  {bottombar_style}
}}
.bottom-bar button {{
  background: #262626;
  color: #f1f1f1;
  border: 1px solid #444;
  border-radius: 8px;
  padding: 10px 16px;
  cursor: pointer;
  min-width: 100px;
}}
.footer {{
  text-align: center;
  background: #1a1a1a;
  color: #bbb;
  font-size: 13px;
  padding: 8px 16px 14px;
}}
</style>
</head>
<body>
<div class="app">
  <aside class="sidebar">
    <label for="jumpSelect">Slides</label>
    <select id="jumpSelect" size="20" onchange="jumpSlide(this.value)"></select>
  </aside>

  <div class="main">
    <div class="header">
      <div class="header-title">{html.escape(title)}</div>
      <div class="header-sub" id="counter">Slide 1</div>
    </div>

    <div class="stage-wrap">
      <div class="frame">
        <img id="slideImage" class="slide-image" alt="Slide" />
        <div class="hotspot-layer" id="hotspotLayer"></div>
      </div>
    </div>

    <div class="bottom-bar">
      <button onclick="prevSlide()">Previous</button>
      <button onclick="nextSlide()">Next</button>
    </div>

    <div class="footer">Static slide image with HTML hotspot overlays.</div>
  </div>
</div>

<script src="scormdriver.js"></script>
<script>
const course = {data};
let currentSlide = 1;

const slideImage = document.getElementById("slideImage");
const hotspotLayer = document.getElementById("hotspotLayer");
const counter = document.getElementById("counter");
const jumpSelect = document.getElementById("jumpSelect");

function pct(value, total) {{
  return (value / total) * 100;
}}

function populateJumpMenu() {{
  if (!jumpSelect) return;
  jumpSelect.innerHTML = "";
  for (const slide of course.slides) {{
    const opt = document.createElement("option");
    opt.value = slide.index;
    opt.textContent = `Slide ${{slide.index}}`;
    jumpSelect.appendChild(opt);
  }}
}}

function createHotspot(link) {{
  const a = document.createElement("a");
  a.className = "hotspot";
  a.href = "#";

  if (link.kind === "external" && link.url) {{
    a.href = link.url;
    a.target = "_blank";
    a.rel = "noopener noreferrer";
  }} else if (link.kind === "internal" && link.target_slide) {{
    a.addEventListener("click", function(e) {{
      e.preventDefault();
      goToSlide(Number(link.target_slide));
    }});
  }} else {{
    return null;
  }}

  if (link.geom === "ellipse") {{
    a.classList.add("ellipse");
    a.style.left = pct(link.x, course.slideWidthEmu) + "%";
    a.style.top = pct(link.y, course.slideHeightEmu) + "%";
    a.style.width = pct(link.w, course.slideWidthEmu) + "%";
    a.style.height = pct(link.h, course.slideHeightEmu) + "%";
  }} else if (link.geom === "line") {{
    const x1p = pct(link.x1, course.slideWidthEmu);
    const y1p = pct(link.y1, course.slideHeightEmu);
    const x2p = pct(link.x2, course.slideWidthEmu);
    const y2p = pct(link.y2, course.slideHeightEmu);

    const dx = x2p - x1p;
    const dy = y2p - y1p;
    const len = Math.sqrt(dx*dx + dy*dy);
    const angle = Math.atan2(dy, dx) * 180 / Math.PI;

    a.style.left = x1p + "%";
    a.style.top = y1p + "%";
    a.style.width = len + "%";
    a.style.height = "2%";
    a.style.transformOrigin = "0 50%";
    a.style.transform = `translateY(-50%) rotate(${{angle}}deg)`;
  }} else {{
    a.style.left = pct(link.x, course.slideWidthEmu) + "%";
    a.style.top = pct(link.y, course.slideHeightEmu) + "%";
    a.style.width = pct(link.w, course.slideWidthEmu) + "%";
    a.style.height = pct(link.h, course.slideHeightEmu) + "%";
  }}

  return a;
}}

function renderHotspots(slide) {{
  hotspotLayer.innerHTML = "";
  for (const link of slide.hotspots || []) {{
    const el = createHotspot(link);
    if (el) hotspotLayer.appendChild(el);
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
  slideImage.src = slide.image;
  renderHotspots(slide);

  counter.textContent = `Slide ${{currentSlide}} of ${{course.slides.length}}`;
  if (jumpSelect) jumpSelect.value = String(currentSlide);
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


# ============================================================
# Build ZIP
# ============================================================

def build_scorm_zip(pptx_bytes: bytes, pptx_name: str, course_title: str, show_nav: bool):
    prs = Presentation(io.BytesIO(pptx_bytes))
    course, media = extract_course(prs)

    mem = io.BytesIO()
    with zipfile.ZipFile(mem, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, blob in media.items():
            zf.writestr(name, blob)

        zf.writestr("index_lms.html", build_player_html(course_title, course, show_nav=show_nav))
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
        "assets": len(media),
    }
    return mem.read(), out_name, summary


# ============================================================
# Streamlit UI
# ============================================================

st.set_page_config(page_title=APP_TITLE, layout="centered")
st.title(APP_TITLE)
st.caption("Pure-Python MVP: PPTX → static slide images + hotspots → SCORM 1.2 ZIP")

with st.expander("What this version supports"):
    st.markdown(
        """
- Static slide image rendering in Python
- Cropped normal images
- Slide background images
- Shape fill images where detectable
- Text rendered into the slide image
- Basic shapes: circles/ovals, rectangles, rounded rectangles, lines, simple arrows
- Basic tables
- External hyperlinks
- Internal slide-jump hotspots
- HTML hotspot overlays
- Optional built-in navigation
        """
    )

course_title = st.text_input("Course title", value="My SCORM Course")
uploaded = st.file_uploader("Upload PPTX", type=["pptx"])
show_nav = st.checkbox("Show built-in navigation", value=True)

if st.button("Publish SCORM", type="primary", use_container_width=True):
    if not uploaded:
        st.error("Please upload a .pptx file first.")
    else:
        try:
            zip_bytes, zip_name, summary = build_scorm_zip(
                uploaded.getvalue(),
                uploaded.name,
                course_title,
                show_nav,
            )
            st.success("SCORM package created.")
            st.write(f"Slides detected: {summary['slides']}")
            st.write(f"Generated assets: {summary['assets']}")
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
