import base64
import html
import io
import json
import math
import os
import re
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path

import streamlit as st
from PIL import Image
from pptx import Presentation
from pptx.enum.action import PP_ACTION
from pptx.enum.shapes import MSO_SHAPE_TYPE


# =========================
# App config
# =========================
st.set_page_config(page_title="PPT to SCORM", layout="wide")
st.title("PPT to SCORM Package Generator")


# =========================
# Helpers
# =========================
def sanitize_filename(name: str) -> str:
    name = re.sub(r"[^\w\-. ]+", "_", name).strip()
    return name or "package"


def px_from_emu(emu: int) -> float:
    # 1 inch = 914400 EMU, 96 px = 1 inch
    return emu * 96.0 / 914400.0


def shape_has_internal_jump(shape) -> tuple[bool, int | None]:
    """
    Returns (True, target_slide_index_1_based) if shape has an internal jump action.
    Otherwise returns (False, None).
    """
    try:
        click_action = shape.click_action
        action = click_action.action

        # Common internal jump types
        if action == PP_ACTION.NAMED_SLIDE:
            target = click_action.target_slide
            if target is not None:
                return True, target.slide_id

        if action == PP_ACTION.FIRST_SLIDE:
            return True, -1

        if action == PP_ACTION.LAST_SLIDE:
            return True, -2

        if action == PP_ACTION.NEXT_SLIDE:
            return True, -3

        if action == PP_ACTION.PREVIOUS_SLIDE:
            return True, -4

    except Exception:
        pass

    return False, None


def build_slide_id_to_index(prs: Presentation) -> dict[int, int]:
    mapping = {}
    for idx, slide in enumerate(prs.slides, start=1):
        mapping[slide.slide_id] = idx
    return mapping


def resolve_relative_target(current_idx: int, target_marker: int, total_slides: int) -> int | None:
    if target_marker == -1:
        return 1
    if target_marker == -2:
        return total_slides
    if target_marker == -3:
        return min(total_slides, current_idx + 1)
    if target_marker == -4:
        return max(1, current_idx - 1)
    return None


def extract_internal_hotspots(prs: Presentation) -> list[dict]:
    slide_id_to_index = build_slide_id_to_index(prs)
    total_slides = len(prs.slides)
    slides_data = []

    for slide_idx, slide in enumerate(prs.slides, start=1):
        hotspots = []

        for shape in slide.shapes:
            try:
                has_jump, target = shape_has_internal_jump(shape)
                if not has_jump:
                    continue

                target_index = None
                if target in slide_id_to_index:
                    target_index = slide_id_to_index[target]
                else:
                    target_index = resolve_relative_target(slide_idx, target, total_slides)

                if not target_index:
                    continue

                x = px_from_emu(shape.left)
                y = px_from_emu(shape.top)
                w = px_from_emu(shape.width)
                h = px_from_emu(shape.height)

                if w <= 1 or h <= 1:
                    continue

                hotspots.append(
                    {
                        "x": round(x, 2),
                        "y": round(y, 2),
                        "w": round(w, 2),
                        "h": round(h, 2),
                        "target": target_index,
                    }
                )

            except Exception:
                continue

        slides_data.append({"slide": slide_idx, "hotspots": hotspots})

    return slides_data


def find_libreoffice() -> str | None:
    candidates = [
        shutil.which("soffice"),
        shutil.which("libreoffice"),
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for c in candidates:
        if c and os.path.exists(c):
            return c
    return None


def export_slides_with_powerpoint_windows(pptx_path: str, out_dir: str) -> bool:
    """
    Best rendering on Windows if PowerPoint is installed.
    Exports each slide as PNG.
    """
    try:
        import win32com.client  # type: ignore
    except Exception:
        return False

    powerpoint = None
    presentation = None
    try:
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.Visible = 1
        presentation = powerpoint.Presentations.Open(os.path.abspath(pptx_path), WithWindow=False)
        presentation.Export(os.path.abspath(out_dir), "PNG")
        return True
    except Exception:
        return False
    finally:
        try:
            if presentation is not None:
                presentation.Close()
        except Exception:
            pass
        try:
            if powerpoint is not None:
                powerpoint.Quit()
        except Exception:
            pass


def export_slides_with_libreoffice(pptx_path: str, out_dir: str) -> bool:
    soffice = find_libreoffice()
    if not soffice:
        return False

    try:
        subprocess.run(
            [
                soffice,
                "--headless",
                "--convert-to",
                "png",
                "--outdir",
                out_dir,
                pptx_path,
            ],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
    except Exception:
        return False

    # LibreOffice often exports the whole deck as multiple slide PNGs into the folder.
    # Some systems may instead create a single file or use naming patterns.
    pngs = list(Path(out_dir).glob("*.png"))
    return len(pngs) > 0


def collect_exported_slide_images(export_dir: str, expected_count: int) -> list[Path]:
    """
    Tries to sort exported PNGs into slide order.
    """
    pngs = list(Path(export_dir).glob("*.png"))
    if not pngs:
        return []

    def slide_sort_key(path: Path):
        name = path.stem.lower()
        nums = re.findall(r"\d+", name)
        if nums:
            return int(nums[-1])
        return 10**9

    pngs = sorted(pngs, key=slide_sort_key)

    # If count matches, great
    if len(pngs) >= expected_count:
        return pngs[:expected_count]

    return pngs


def export_slide_images(pptx_path: str, work_dir: str, expected_count: int) -> list[Path]:
    """
    Priority:
    1. Windows PowerPoint COM export for best fidelity
    2. LibreOffice export
    """
    export_dir = os.path.join(work_dir, "slide_exports")
    os.makedirs(export_dir, exist_ok=True)

    ok = export_slides_with_powerpoint_windows(pptx_path, export_dir)
    if not ok:
        ok = export_slides_with_libreoffice(pptx_path, export_dir)

    if not ok:
        raise RuntimeError(
            "Could not export slide images. Install Microsoft PowerPoint on Windows "
            "or LibreOffice for headless export."
        )

    images = collect_exported_slide_images(export_dir, expected_count)
    if not images:
        raise RuntimeError("Slide image export finished, but no PNG slide images were found.")

    return images


def image_to_base64(path: Path) -> str:
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")


def write_text(path: str, content: str):
    with open(path, "w", encoding="utf-8") as f:
        f.write(content)


def build_player_html(
    title: str,
    slide_width_px: int,
    slide_height_px: int,
    image_files: list[str],
    hotspots_by_slide: list[dict],
) -> str:
    images_json = json.dumps(image_files)
    hotspots_json = json.dumps(hotspots_by_slide)

    safe_title = html.escape(title)

    return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{safe_title}</title>
  <style>
    :root {{
      --bg: #111;
      --panel: #1b1b1b;
      --text: #fff;
      --muted: #ccc;
      --btn: #2b2b2b;
      --btn-hover: #3a3a3a;
      --accent: #4da3ff;
    }}

    * {{
      box-sizing: border-box;
    }}

    html, body {{
      margin: 0;
      padding: 0;
      width: 100%;
      height: 100%;
      background: var(--bg);
      color: var(--text);
      font-family: Arial, Helvetica, sans-serif;
      overflow: hidden;
    }}

    .app {{
      display: flex;
      flex-direction: column;
      width: 100%;
      height: 100%;
    }}

    .topbar {{
      flex: 0 0 auto;
      display: flex;
      align-items: center;
      justify-content: center;
      gap: 12px;
      padding: 14px 18px;
      background: rgba(0,0,0,0.45);
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

    .slide-count {{
      font-size: 20px;
      font-weight: 700;
      letter-spacing: 0.02em;
      min-width: 150px;
      text-align: center;
    }}

    .stage-wrap {{
      flex: 1 1 auto;
      min-height: 0;
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 18px;
    }}

    .stage {{
      position: relative;
      width: min(calc(100vw - 36px), calc((100vh - 110px) * {slide_width_px / slide_height_px:.8f}));
      aspect-ratio: {slide_width_px} / {slide_height_px};
      max-height: calc(100vh - 110px);
      background: #000;
      overflow: hidden;
      border-radius: 8px;
      box-shadow: 0 10px 30px rgba(0,0,0,0.4);
    }}

    .slide-image {{
      position: absolute;
      inset: 0;
      width: 100%;
      height: 100%;
      object-fit: contain;
      display: block;
      user-select: none;
      -webkit-user-drag: none;
    }}

    .hotspot {{
      position: absolute;
      background: transparent;
      border: none;
      outline: none;
      cursor: pointer;
      padding: 0;
      margin: 0;
    }}

    .hotspot:focus {{
      outline: 2px solid rgba(77,163,255,0.7);
      outline-offset: -2px;
    }}
  </style>
</head>
<body>
  <div class="app">
    <div class="topbar">
      <button id="prevBtn" type="button">◀ Prev</button>
      <div id="slideCount" class="slide-count">1 / 1</div>
      <button id="nextBtn" type="button">Next ▶</button>
    </div>

    <div class="stage-wrap">
      <div id="stage" class="stage">
        <img id="slideImage" class="slide-image" alt="Slide">
      </div>
    </div>
  </div>

  <script>
    const images = {images_json};
    const hotspotsBySlide = {hotspots_json};
    const slideWidth = {slide_width_px};
    const slideHeight = {slide_height_px};

    let current = 0;

    const slideImage = document.getElementById("slideImage");
    const stage = document.getElementById("stage");
    const slideCount = document.getElementById("slideCount");
    const prevBtn = document.getElementById("prevBtn");
    const nextBtn = document.getElementById("nextBtn");

    function clamp(n, min, max) {{
      return Math.max(min, Math.min(max, n));
    }}

    function render() {{
      current = clamp(current, 0, images.length - 1);
      slideImage.src = images[current];
      slideCount.textContent = `${{current + 1}} / ${{images.length}}`;
      prevBtn.disabled = current === 0;
      nextBtn.disabled = current === images.length - 1;

      [...stage.querySelectorAll(".hotspot")].forEach(el => el.remove());

      const entry = hotspotsBySlide.find(s => s.slide === current + 1);
      if (!entry || !entry.hotspots) return;

      for (const h of entry.hotspots) {{
        const btn = document.createElement("button");
        btn.className = "hotspot";
        btn.type = "button";
        btn.setAttribute("aria-label", `Go to slide ${{h.target}}`);

        btn.style.left = `${{(h.x / slideWidth) * 100}}%`;
        btn.style.top = `${{(h.y / slideHeight) * 100}}%`;
        btn.style.width = `${{(h.w / slideWidth) * 100}}%`;
        btn.style.height = `${{(h.h / slideHeight) * 100}}%`;

        btn.addEventListener("click", () => {{
          current = h.target - 1;
          render();
        }});

        stage.appendChild(btn);
      }}
    }}

    prevBtn.addEventListener("click", () => {{
      if (current > 0) {{
        current -= 1;
        render();
      }}
    }});

    nextBtn.addEventListener("click", () => {{
      if (current < images.length - 1) {{
        current += 1;
        render();
      }}
    }});

    document.addEventListener("keydown", (e) => {{
      if (e.key === "ArrowLeft") {{
        if (current > 0) {{
          current -= 1;
          render();
        }}
      }} else if (e.key === "ArrowRight") {{
        if (current < images.length - 1) {{
          current += 1;
          render();
        }}
      }}
    }});

    render();
  </script>
</body>
</html>
"""


def build_scorm_api_js() -> str:
    return r"""
var scormData = {
  "cmi.core.lesson_status": "not attempted",
  "cmi.core.score.raw": "",
  "cmi.core.student_name": "Learner",
  "cmi.core.student_id": "0001"
};

function LMSInitialize(param) { return "true"; }
function LMSFinish(param) { return "true"; }
function LMSGetValue(key) {
  return scormData[key] || "";
}
function LMSSetValue(key, value) {
  scormData[key] = value;
  return "true";
}
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


def build_scorm_launcher_html(course_title: str) -> str:
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
    };
  </script>
</head>
<frameset rows="100%" border="0">
  <frame src="player.html" name="content_frame" frameborder="0" />
</frameset>
</html>
"""


def build_manifest_xml(course_id: str, course_title: str) -> str:
    safe_title = html.escape(course_title)
    safe_id = re.sub(r"[^A-Za-z0-9_.-]", "_", course_id)

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


def create_scorm_package(uploaded_file, course_title: str) -> bytes:
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir_path = Path(tmpdir)
        pptx_path = tmpdir_path / "input.pptx"

        with open(pptx_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        prs = Presentation(str(pptx_path))
        slide_width_px = int(round(px_from_emu(prs.slide_width)))
        slide_height_px = int(round(px_from_emu(prs.slide_height)))
        hotspots = extract_internal_hotspots(prs)

        exported_images = export_slide_images(str(pptx_path), tmpdir, len(prs.slides))

        package_dir = tmpdir_path / "scorm_package"
        assets_dir = package_dir / "assets"
        assets_dir.mkdir(parents=True, exist_ok=True)

        image_files = []
        for idx, src in enumerate(exported_images, start=1):
            ext = src.suffix.lower() or ".png"
            dst_name = f"slide_{idx:03d}{ext}"
            dst = assets_dir / dst_name
            shutil.copy2(src, dst)
            image_files.append(f"assets/{dst_name}")

        player_html = build_player_html(
            title=course_title,
            slide_width_px=slide_width_px,
            slide_height_px=slide_height_px,
            image_files=image_files,
            hotspots_by_slide=hotspots,
        )

        write_text(str(package_dir / "player.html"), player_html)
        write_text(str(package_dir / "scorm_api.js"), build_scorm_api_js())
        write_text(str(package_dir / "index.html"), build_scorm_launcher_html(course_title))
        write_text(str(package_dir / "imsmanifest.xml"), build_manifest_xml("ppt_to_scorm_course", course_title))

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for root, _, files in os.walk(package_dir):
                for file in files:
                    abs_path = os.path.join(root, file)
                    rel_path = os.path.relpath(abs_path, package_dir)
                    zf.write(abs_path, rel_path)

        zip_buffer.seek(0)
        return zip_buffer.read()


# =========================
# UI
# =========================
uploaded = st.file_uploader("Upload PowerPoint (.pptx)", type=["pptx"])
default_name = "ppt_scorm_package"

col1, col2 = st.columns([2, 1])
with col1:
    package_name = st.text_input("Package name", value=default_name)
with col2:
    course_title = st.text_input("Course title", value="PPT Course")

st.caption("Best fidelity is on Windows with Microsoft PowerPoint installed. LibreOffice is used as fallback.")

if uploaded is not None:
    st.success(f"Loaded: {uploaded.name}")

    if st.button("Publish SCORM package", type="primary"):
        try:
            output_bytes = create_scorm_package(uploaded, course_title=course_title.strip() or "PPT Course")
            safe_name = sanitize_filename(package_name) + ".zip"

            st.success("SCORM package created successfully.")
            st.download_button(
                label="Download SCORM ZIP",
                data=output_bytes,
                file_name=safe_name,
                mime="application/zip",
            )

        except Exception as e:
            st.error(f"Failed to publish SCORM package: {e}")
