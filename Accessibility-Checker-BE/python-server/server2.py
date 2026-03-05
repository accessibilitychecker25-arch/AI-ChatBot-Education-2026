import os
import time
import shutil
from typing import List, Optional
from pathlib import Path
import zipfile
import xml.etree.ElementTree as ET
import re
import json
from lxml import etree

import platform
import subprocess
import uuid
import win32com.client

from fastapi import FastAPI, File, UploadFile, HTTPException, Body, Request, Response
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse, PlainTextResponse
from fastapi.exceptions import RequestValidationError
from starlette.exceptions import HTTPException as StarletteHTTPException
import traceback

# ---------- CONFIG ----------
BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# ---------- APP SETUP ----------
app = FastAPI()

# Configure CORS (Angular frontend -> Python backend)
origins = [
    "http://localhost:4200",
    "http://localhost:3000",
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
    expose_headers=["Content-Disposition"],
)

@app.exception_handler(Exception)
async def debug_exception_handler(request: Request, exc: Exception):
    traceback.print_exc()
    return PlainTextResponse(str(exc), status_code=500)

@app.middleware("http")
async def access_log(request: Request, call_next):
    t0 = time.time()
    response = await call_next(request)
    ms = (time.time() - t0) * 1000
    print(f"[{request.method}] {request.url.path} -> {response.status_code} ({ms:.2f} ms)")
    return response

@app.get("/")
def health_check():
    return {"status": "running", "service": "PowerPoint Accessibility Backend"}

SOFFICE_PATH = r"C:\Program Files\LibreOffice\program\soffice.exe"

def is_windows() -> bool:
    return platform.system().lower().startswith("win")

def convert_legacy_ppt_to_pptx_powerpoint(src_path: Path, out_dir: Path) -> Path:

    out_dir.mkdir(parents=True, exist_ok=True)
    dst_path = out_dir / f"{src_path.stem}.pptx"

    # try:
    #     import win32com.client  # type: ignore
    # except Exception as e:
    #     raise RuntimeError(f"pywin32 not available: {e}")

    pp = win32com.client.Dispatch("PowerPoint.Application")
    pp.Visible = 1

    try:
        pres = pp.Presentations.Open(str(src_path), 1, 0, 0)  # ReadOnly=1, WithWindow=0
        try:
            pres.SaveAs(str(dst_path), 24)  # 24 = ppSaveAsOpenXMLPresentation (.pptx)
        finally:
            pres.Close()
    finally:
        pp.Quit()

    if not dst_path.exists():
        raise RuntimeError("PowerPoint conversion did not produce a .pptx file.")
    return dst_path

def convert_legacy_to_pptx(src_path: Path, out_dir: Path) -> Path:

    if is_windows():
        try:
            return convert_legacy_ppt_to_pptx_powerpoint(src_path, out_dir)
        except Exception as e:
            # fallback to LibreOffice if PowerPoint fails
            return convert_legacy_ppt_to_pptx_powerpoint(src_path, out_dir)
    else:
        return convert_legacy_ppt_to_pptx_powerpoint(src_path, out_dir)
    
@app.post("/upload")
async def upload_files(
    # Accept ANY of these common multipart field names:
    files: Optional[List[UploadFile]] = File(default=None),
    file: Optional[UploadFile] = File(default=None),
    pptxFile: Optional[UploadFile] = File(default=None),
    docxFile: Optional[UploadFile] = File(default=None),
):
    """
    Accepts PowerPoint files, analyzes them, and returns accessibility report.
    Compatible with FE sending: files[] OR file OR pptxFile OR docxFile.
    """

    # ---- normalize inputs into one list ----
    incoming: List[UploadFile] = []
    if files:
        incoming.extend(files)
    if file:
        incoming.append(file)
    if pptxFile:
        incoming.append(pptxFile)
    if docxFile:
        incoming.append(docxFile)

    if not incoming:
        raise HTTPException(
            status_code=400,
            detail="No file uploaded. Send multipart/form-data with one of: files, file, pptxFile, docxFile"
        )

    # if len(incoming) > 10:
    #     raise HTTPException(
    #         status_code=400,
    #         detail=f"Too many files. You uploaded {len(incoming)}, but the limit is 10."
    #     )

    # For now handle single file
    up = incoming[0]
    filename = up.filename or "unnamed.pptx"
    filename_lower = filename.lower()

    # allowed_ext = (".pptx", ".ppt", ".pps", ".potx")
    allowed_ext = (".pptx", ".ppt", ".pps", ".pot", ".potx", ".ppsx")

    if not filename_lower.endswith(allowed_ext):
        raise HTTPException(
            status_code=400,
            detail="Invalid file type. Please upload a PowerPoint file."
        )

    # Save upload
    try:
        file_location = UPLOAD_DIR / filename
        with file_location.open("wb") as buffer:
            shutil.copyfileobj(up.file, buffer)
            
        ext = Path(filename_lower).suffix
        converted_dir = UPLOAD_DIR / "converted"

        if ext in [".ppt", ".pps", ".pot"]:
            pptx_input = convert_legacy_ppt_to_pptx_powerpoint(file_location, converted_dir)
        else:
            pptx_input = file_location
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to save file: {str(e)}")

    # Stage output file (copy for now)
    # Stage output file (NOW: remediate alt text into the output pptx)
    try:
        base = Path(filename).stem
        if base.startswith("remediated-"):
            out_name = f"{base}.pptx"
        else:
            out_name = f"remediated-{base}.pptx"
        out_path = OUTPUT_DIR / out_name

        # fixed_count, fix_details = remediate_alt_text_pptx(file_location, out_path)
        fixed_count, fix_details = remediate_alt_text_pptx(pptx_input, out_path)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to remediate alt text: {str(e)}")

    # Analyze & respond
    # Analyze & respond (analyze the REMEDIATED file)
    try:
        report = analyze_powerpoint(out_path, out_name)
        report["summary"]["fixed"] += fixed_count
        report["details"]["autoFixedAltText"] = fix_details
        return JSONResponse(content={
            "fileName": filename,
            "suggestedFileName": out_name,
            "report": report
        })
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to analyze file: {str(e)}")

# @app.post("/upload")
# async def upload_files(
#     # Accept ANY of these common multipart field names:
#     files: Optional[List[UploadFile]] = File(default=None),
#     file: Optional[UploadFile] = File(default=None),
#     pptxFile: Optional[UploadFile] = File(default=None),
#     docxFile: Optional[UploadFile] = File(default=None),
# ):
#     """
#     Accepts PowerPoint files, analyzes them, and returns accessibility report.
#     Compatible with FE sending: files[] OR file OR pptxFile OR docxFile.
#     """

#     # ---- normalize inputs into one list ----
#     incoming: List[UploadFile] = []
#     if files:
#         incoming.extend(files)
#     if file:
#         incoming.append(file)
#     if pptxFile:
#         incoming.append(pptxFile)
#     if docxFile:
#         incoming.append(docxFile)

#     if not incoming:
#         raise HTTPException(
#             status_code=400,
#             detail="No file uploaded. Send multipart/form-data with one of: files, file, pptxFile, docxFile"
#         )

#     # if len(incoming) > 10:
#     #     raise HTTPException(
#     #         status_code=400,
#     #         detail=f"Too many files. You uploaded {len(incoming)}, but the limit is 10."
#     #     )

#     # For now handle single file (same as your current logic)
#     up = incoming[0]
#     filename = up.filename or "unnamed.pptx"
#     filename_lower = filename.lower()

#     allowed_ext = (".pptx", ".ppt", ".pps", ".pot", ".potx", ".ppsx")
#     if not filename_lower.endswith(allowed_ext):
#         raise HTTPException(status_code=400, detail="Invalid file type")

#     file_location = UPLOAD_DIR / filename
#     with file_location.open("wb") as buffer:
#         shutil.copyfileobj(up.file, buffer)

#     ext = Path(filename_lower).suffix

#     # Make a per-upload converted folder (prevents collisions)
#     converted_dir = UPLOAD_DIR / "converted" / uuid.uuid4().hex[:8]
#     converted_dir.mkdir(parents=True, exist_ok=True)

#     if ext in [".ppt", ".pps", ".pot"]:
#         # Windows: PowerPoint first; otherwise LibreOffice
#         pptx_input = convert_legacy_to_pptx(file_location, converted_dir)
#     else:
#         pptx_input = file_location  # already OpenXML

#     # Save upload
#     try:
#         file_location = UPLOAD_DIR / filename
#         with file_location.open("wb") as buffer:
#             shutil.copyfileobj(up.file, buffer)
            
#         ext = Path(filename_lower).suffix
#         converted_dir = UPLOAD_DIR / "converted"

#         if ext in [".ppt", ".pps", ".pot"]:
#             pptx_input = convert_legacy_ppt_to_pptx_powerpoint(file_location, converted_dir)
#         else:
#             pptx_input = file_location
#     except Exception as e:
#         raise HTTPException(status_code=500, detail=f"Failed to save file: {str(e)}")

#     # Stage output file (copy for now)
#     # Stage output file (NOW: remediate alt text into the output pptx)
#     try:
#         base = Path(filename).stem
#         if base.startswith("remediated-"):
#             out_name = f"{base}.pptx"
#         else:
#             out_name = f"remediated-{base}.pptx"
#         out_path = OUTPUT_DIR / out_name

#         # fixed_count, fix_details = remediate_alt_text_pptx(file_location, out_path)
#         fixed_count, fix_details = remediate_alt_text_pptx(pptx_input, out_path)
#     except Exception as e:
#         raise HTTPException(status_code=500, detail=f"Failed to remediate alt text: {str(e)}")

#     # Analyze & respond
#     # Analyze & respond (analyze the REMEDIATED file)
#     try:
#         report = analyze_powerpoint(out_path, out_name)
#         report["summary"]["fixed"] += fixed_count
#         report["details"]["autoFixedAltText"] = fix_details
#         return JSONResponse(content={
#             "fileName": filename,
#             "suggestedFileName": out_name,
#             "report": report
#         })
#     except Exception as e:
#         raise HTTPException(status_code=500, detail=f"Failed to analyze file: {str(e)}")

def analyze_powerpoint(file_path: Path, filename: str):
    """
    Analyze PowerPoint file for accessibility issues.
    Checks:
    1. Slide titles (missing or empty)
    2. Image alt text
    3. GIF detection
    4. Presentation title
    5. File naming
    6. Hidden slides
    7. List formatting issues
    """
    report = {
        "fileName": filename,
        "suggestedFileName": f"remediated-{Path(filename).stem}.pptx",
        "summary": {"fixed": 0, "flagged": 0},
        "details": {
            "titleNeedsFixing": False,
            "slidesMissingTitles": [],
            "imagesMissingOrBadAlt": [],
            "gifsDetected": [],
            "fileNameNeedsFixing": False,
            "hiddenSlidesDetected": [],
            "listFormattingIssues": [],
        }
    }

    try:
        # Open PPTX as ZIP
        with zipfile.ZipFile(file_path, 'r') as zip_file:
            # Check presentation title
            try:
                core_xml = zip_file.read('docProps/core.xml').decode('utf-8')
                if '<dc:title></dc:title>' in core_xml or '<dc:title/>' in core_xml:
                    report["details"]["titleNeedsFixing"] = True
                    report["summary"]["flagged"] += 1
            except:
                pass

            # Check filename
            if '_' in filename or filename.lower().startswith('presentation') or filename.lower().startswith('untitled'):
                report["details"]["fileNameNeedsFixing"] = True
                report["summary"]["flagged"] += 1

            # Get list of slides
            slides = [name for name in zip_file.namelist() if name.startswith('ppt/slides/slide') and name.endswith('.xml')]
            slides.sort()

            # Analyze each slide
            for i, slide_path in enumerate(slides):
                slide_number = i + 1
                slide_xml = zip_file.read(slide_path).decode('utf-8')

                # Check slide title
                title_check = check_slide_title(slide_xml, slide_number)
                if title_check["missing"]:
                    report["details"]["slidesMissingTitles"].append(title_check)
                    report["summary"]["flagged"] += 1

                # Check images
                image_issues = check_slide_images(slide_xml, slide_number)
                if image_issues:
                    report["details"]["imagesMissingOrBadAlt"].extend(image_issues)
                    report["summary"]["flagged"] += len(image_issues)

                # Check for list formatting issues
                list_issues = check_list_formatting(slide_xml, slide_number)
                if list_issues:
                    report["details"]["listFormattingIssues"].extend(list_issues)
                    report["summary"]["flagged"] += len(list_issues)

            # Check for GIFs
            gif_files = [name for name in zip_file.namelist() if name.startswith('ppt/media/') and name.lower().endswith('.gif')]
            if gif_files:
                report["details"]["gifsDetected"] = gif_files
                report["summary"]["flagged"] += len(gif_files)

    except Exception as e:
        print(f"Error analyzing PowerPoint: {e}")
        raise

    return report


def check_slide_title(slide_xml: str, slide_number: int):
    """Check if slide has a title."""
    # Look for title placeholder
    title_pattern = r'<p:ph[^>]*type="(title|ctrTitle)"[^>]*>'
    has_title_placeholder = re.search(title_pattern, slide_xml)
    
    if not has_title_placeholder:
        return {
            "missing": True,
            "slideNumber": slide_number,
            "message": f"Slide {slide_number} is missing a title"
        }
    
    # Check if title has text
    text_pattern = r'<a:t[^>]*>(.*?)</a:t>'
    text_matches = re.findall(text_pattern, slide_xml)
    
    if not any(text.strip() for text in text_matches):
        return {
            "missing": True,
            "slideNumber": slide_number,
            "message": f"Slide {slide_number} has an empty title"
        }
    
    return {"missing": False}


def check_list_formatting(slide_xml: str, slide_number: int):
    """Check for hyphenated paragraphs that should be lists."""
    issues = []
    
    # Find all text elements
    text_pattern = r'<a:t[^>]*>(.*?)</a:t>'
    text_matches = re.findall(text_pattern, slide_xml)
    
    for text in text_matches:
        # Check for hyphenated list patterns
        if re.match(r'^[\s]*[-–—•]\s+.+', text):
            issues.append({
                "slideNumber": slide_number,
                "location": f"Slide {slide_number}",
                "issue": f'Possible improperly formatted list: "{text[:50]}..."',
                "type": "listFormatting"
            })
    
    return issues


ALT_TEXT_MAX = 250

def check_slide_images(slide_xml: str, slide_number: int):
    issues = []

    pic_pattern = r'<p:pic[\s\S]*?</p:pic>'
    pic_matches = re.findall(pic_pattern, slide_xml)

    for pic_xml in pic_matches:
        cnvpr_pattern = r'<p:cNvPr([^>]*)/?>'
        m = re.search(cnvpr_pattern, pic_xml)
        attrs = m.group(1) if m else ""

        def get_attr(attr_name: str) -> str:
            am = re.search(rf'{attr_name}="([^"]*)"', attrs)
            return am.group(1) if am else ""

        shape_id = get_attr("id")
        shape_name = get_attr("name")
        alt_text = get_attr("descr")

        alt_text_clean = (alt_text or "").strip().lower()
        is_decorative = (alt_text_clean == "decorative")

        # --- RULES ---

        # 1. Missing alt text
        if not alt_text or alt_text.strip() == "":
            issues.append({
                "slideNumber": slide_number,
                "shapeId": shape_id,
                "shapeName": shape_name,
                "issue": "Image missing alt text",
                "type": "imageAltMissing"
            })

        # 2. Decorative images
        elif is_decorative:
            continue

        # 3. Too long alt text
        elif len(alt_text) > ALT_TEXT_MAX:
            issues.append({
                "slideNumber": slide_number,
                "shapeId": shape_id,
                "shapeName": shape_name,
                "issue": f"Alt text exceeds {ALT_TEXT_MAX} characters",
                "type": "imageAltTooLong",
                "length": len(alt_text),
                "max": ALT_TEXT_MAX
            })

        elif alt_text_clean in ["image", "picture", "photo"]:
            issues.append({
                "slideNumber": slide_number,
                "shapeId": shape_id,
                "shapeName": shape_name,
                "issue": "Alt text is too generic",
                "type": "imageAltTooGeneric"
            })

    return issues

def escape_xml_attr(s: str) -> str:
    return (s.replace("&", "&amp;")
             .replace('"', "&quot;")
             .replace("<", "&lt;")
             .replace(">", "&gt;"))

def choose_default_alt(shape_name: str, slide_number: int) -> str:
    """
    Heuristic:
    - If it looks decorative (name hints), set "decorative"
    - Otherwise set a non-generic placeholder
    """
    n = (shape_name or "").lower()
    decorative_hints = ["background", "bg", "decor", "decoration", "border", "divider", "logo", "icon", "watermark"]
    if any(h in n for h in decorative_hints):
        return "decorative"
    return f"Image on slide {slide_number}"

def remediate_slide_alt_text(slide_xml: str, slide_number: int):
    """
    Returns: (new_xml, fixed_count, fix_details)
    Fix rules:
      - Missing descr -> add descr (decorative or placeholder)
      - descr > 250 -> truncate
      - descr is generic image/picture/photo -> replace with placeholder
    """
    fixed = 0
    fix_details = []

    pic_pattern = r'<p:pic[\s\S]*?</p:pic>'
    pics = re.findall(pic_pattern, slide_xml)

    # If no pics, return unchanged
    if not pics:
        return slide_xml, 0, []

    new_xml = slide_xml

    for pic_xml in pics:
        # Extract cNvPr attrs
        cnvpr_pattern = r'<p:cNvPr([^>]*)/?>'
        m = re.search(cnvpr_pattern, pic_xml)
        attrs = m.group(1) if m else ""

        def get_attr(attr_name: str) -> str:
            am = re.search(rf'{attr_name}="([^"]*)"', attrs)
            return am.group(1) if am else ""

        shape_id = get_attr("id")
        shape_name = get_attr("name")
        alt_text = get_attr("descr")
        alt_clean = (alt_text or "").strip().lower()

        # Decide what to write (if needed)
        if not alt_text or alt_text.strip() == "":
            new_alt = choose_default_alt(shape_name, slide_number)
            fixed += 1
            fix_details.append({
                "slideNumber": slide_number,
                "shapeId": shape_id,
                "shapeName": shape_name,
                "fix": "addedAltText",
                "altText": new_alt
            })
            # update in the FULL slide XML by matching the cNvPr with this id
            new_xml = set_cnvpr_descr(new_xml, shape_id, new_alt)

        elif len(alt_text) > ALT_TEXT_MAX:
            new_alt = alt_text[:ALT_TEXT_MAX]
            fixed += 1
            fix_details.append({
                "slideNumber": slide_number,
                "shapeId": shape_id,
                "shapeName": shape_name,
                "fix": "truncatedAltText",
                "altText": new_alt
            })
            new_xml = set_cnvpr_descr(new_xml, shape_id, new_alt)

        elif alt_clean in ["image", "picture", "photo"]:
            new_alt = f"Image on slide {slide_number}"
            fixed += 1
            fix_details.append({
                "slideNumber": slide_number,
                "shapeId": shape_id,
                "shapeName": shape_name,
                "fix": "replacedGenericAltText",
                "altText": new_alt
            })
            new_xml = set_cnvpr_descr(new_xml, shape_id, new_alt)

    return new_xml, fixed, fix_details

def set_cnvpr_descr(full_slide_xml: str, shape_id: str, new_alt: str) -> str:
    """
    Sets/updates descr="..." on the <p:cNvPr ... id="{shape_id}" ...> element.
    Works for both self-closing (<p:cNvPr ... />) and normal (<p:cNvPr ...>).
    """
    if not shape_id:
        return full_slide_xml

    escaped = escape_xml_attr(new_alt)

    # 1) Replace existing descr if present
    pattern_has_descr = rf'(<p:cNvPr\b[^>]*\bid="{re.escape(shape_id)}"[^>]*\bdescr=")([^"]*)(")'
    if re.search(pattern_has_descr, full_slide_xml):
        return re.sub(pattern_has_descr, rf'\1{escaped}\3', full_slide_xml)

    # 2) Inject descr before the tag closes (handles .../> and ...>)
    pattern_inject = rf'(<p:cNvPr\b[^>]*\bid="{re.escape(shape_id)}"[^>]*?)(\s*/?>)'
    return re.sub(pattern_inject, rf'\1 descr="{escaped}"\2', full_slide_xml, count=1)

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"

def set_alt_text_in_slide_xml(slide_xml_bytes: bytes, slide_number: int):
    """
    Finds all picture cNvPr nodes and fixes their 'descr' safely.
    Returns: (new_xml_bytes, fixed_count, fix_details)
    """
    parser = etree.XMLParser(remove_blank_text=False, recover=False)
    root = etree.fromstring(slide_xml_bytes, parser=parser)

    ns = {"p": P_NS}

    fixed = 0
    fix_details = []

    # Pictures: p:pic -> p:nvPicPr -> p:cNvPr
    for cnvpr in root.xpath(".//p:pic/p:nvPicPr/p:cNvPr", namespaces=ns):
        shape_id = cnvpr.get("id") or ""
        shape_name = cnvpr.get("name") or ""
        descr = cnvpr.get("descr")  # can be None

        # Decide if we need a fix
        if descr is None or descr.strip() == "":
            new_alt = choose_default_alt(shape_name, slide_number)  # your existing function
            cnvpr.set("descr", new_alt)
            fixed += 1
            fix_details.append({
                "slideNumber": slide_number,
                "shapeId": shape_id,
                "shapeName": shape_name,
                "fix": "addedAltText",
                "altText": new_alt
            })

        elif len(descr) > ALT_TEXT_MAX:
            new_alt = descr[:ALT_TEXT_MAX]
            cnvpr.set("descr", new_alt)
            fixed += 1
            fix_details.append({
                "slideNumber": slide_number,
                "shapeId": shape_id,
                "shapeName": shape_name,
                "fix": "truncatedAltText",
                "altText": new_alt
            })

        else:
            d = descr.strip().lower()
            if d in ["image", "picture", "photo"]:
                new_alt = f"Image on slide {slide_number}"
                cnvpr.set("descr", new_alt)
                fixed += 1
                fix_details.append({
                    "slideNumber": slide_number,
                    "shapeId": shape_id,
                    "shapeName": shape_name,
                    "fix": "replacedGenericAltText",
                    "altText": new_alt
                })

    new_bytes = etree.tostring(
        root,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=None
    )
    return new_bytes, fixed, fix_details

def remediate_alt_text_pptx(src_pptx: Path, dst_pptx: Path):
    fixed_total = 0
    all_fix_details = []

    with zipfile.ZipFile(src_pptx, "r") as zin, zipfile.ZipFile(dst_pptx, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)

            m = re.match(r"ppt/slides/slide(\d+)\.xml$", item.filename)
            if m:
                slide_num = int(m.group(1))
                try:
                    new_data, fixed, details = set_alt_text_in_slide_xml(data, slide_num)
                    if fixed:
                        data = new_data
                        fixed_total += fixed
                        all_fix_details.extend(details)
                except Exception:
                    # If parsing fails, leave slide unchanged rather than corrupting the file
                    pass

            zout.writestr(item, data)

    return fixed_total, all_fix_details

@app.get("/download")
def download_latest_get():
    candidates = [p for p in OUTPUT_DIR.glob("*") if p.is_file()]
    if not candidates:
        raise HTTPException(status_code=404, detail="No files available to download yet.")
    file_path = max(candidates, key=lambda p: p.stat().st_mtime)

    return FileResponse(
        path=str(file_path),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename=file_path.name
    )

# @app.post("/download")
# async def download_latest(request: Request):
#     """
#     Accepts POST /download from whatever the frontend sends.
#     If body is JSON and contains {"filename": "..."} -> use it.
#     Otherwise -> return newest file from /output.
#     """
#     filename = None

#     body = await request.body()  # raw bytes

#     # Try to parse JSON only if it seems like JSON
#     if body:
#         try:
#             text = body.decode("utf-8")
#             data = json.loads(text)
#             if isinstance(data, dict):
#                 filename = data.get("filename")
#         except Exception:
#             # not JSON / not utf-8 -> ignore and fall back to newest output file
#             pass

#     if filename:
#         file_path = OUTPUT_DIR / filename
#         if not file_path.exists():
#             raise HTTPException(status_code=404, detail=f"File not found: {filename}")
#     else:
#         candidates = [p for p in OUTPUT_DIR.glob("*") if p.is_file()]
#         if not candidates:
#             raise HTTPException(status_code=404, detail="No files available to download yet.")
#         file_path = max(candidates, key=lambda p: p.stat().st_mtime)
#         filename = file_path.name

#     return FileResponse(
#         path=str(file_path),
#         media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
#         filename=filename
#     )
    
@app.post("/download")
async def download_latest_post(request: Request):
    filename = None

    body = await request.body()
    if body:
        try:
            data = json.loads(body.decode("utf-8"))
            if isinstance(data, dict):
                filename = data.get("filename") or data.get("fileName") or data.get("suggestedFileName")
        except Exception:
            pass  # ignore non-json

    if filename:
        file_path = OUTPUT_DIR / filename
        if not file_path.exists():
            raise HTTPException(status_code=404, detail=f"File not found: {filename}")
    else:
        candidates = [p for p in OUTPUT_DIR.glob("*") if p.is_file()]
        if not candidates:
            raise HTTPException(status_code=404, detail="No files available to download yet.")
        file_path = max(candidates, key=lambda p: p.stat().st_mtime)
        filename = file_path.name

    return FileResponse(
        path=str(file_path),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename=filename
    )

# ---------- RUN ----------
if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=5000)