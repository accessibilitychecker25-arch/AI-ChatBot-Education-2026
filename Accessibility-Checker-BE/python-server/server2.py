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

# Load environment variables (optional)
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass  # .env is optional

# Import FREE Local AI Vision - Only Option!
AI_AVAILABLE = False

try:
    from local_vision import generate_alt_text_free, get_vision_model
    local_model = get_vision_model()
    
    if local_model and local_model.is_enabled():
        AI_AVAILABLE = True
        print("✅ Local AI vision model loaded (BLIP - 100% FREE, No Costs)")
    else:
        print("⚠️  Local AI model not ready yet (will download on first use)")
except ImportError as e:
    print(f"⚠️  AI vision module not available: {e}")
    print("ℹ️  Will use placeholder alt text")

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
    files: Optional[List[UploadFile]] = File(default=None),
    file: Optional[UploadFile] = File(default=None),
    pptxFile: Optional[UploadFile] = File(default=None),
    docxFile: Optional[UploadFile] = File(default=None),
):
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

    if len(incoming) > 10:
        raise HTTPException(
            status_code=400,
            detail=f"Too many files. You uploaded {len(incoming)}, but the limit is 10."
        )

    results = []

    for up in incoming:
        try:
            filename = up.filename or "unnamed.pptx"
            filename_lower = filename.lower()
            allowed_ext = (".pptx", ".ppt", ".pps", ".pot", ".potx", ".ppsx")

            if not filename_lower.endswith(allowed_ext):
                results.append({
                    "fileName": filename,
                    "error": "Invalid file type. Please upload a PowerPoint file."
                })
                continue

            # save with unique name to avoid collisions
            unique_prefix = uuid.uuid4().hex[:8]
            saved_name = f"{unique_prefix}_{filename}"
            file_location = UPLOAD_DIR / saved_name

            with file_location.open("wb") as buffer:
                shutil.copyfileobj(up.file, buffer)

            ext = Path(filename_lower).suffix
            converted_dir = UPLOAD_DIR / "converted" / unique_prefix
            converted_dir.mkdir(parents=True, exist_ok=True)

            if ext in [".ppt", ".pps", ".pot"]:
                pptx_input = convert_legacy_to_pptx(file_location, converted_dir)
            else:
                pptx_input = file_location

            base = Path(filename).stem
            out_name = f"remediated-{base}.pptx"
            out_path = OUTPUT_DIR / f"{unique_prefix}_{out_name}"

            fixed_count, fix_details = remediate_alt_text_pptx(pptx_input, out_path)

            report = analyze_powerpoint(out_path, out_name)
            report["summary"]["fixed"] += fixed_count
            report["details"]["autoFixedAltText"] = fix_details

            results.append({
                "fileName": filename,
                # "suggestedFileName": f"{unique_prefix}_{out_name}",
                "suggestedFileName": out_name,
                "report": report
            })

        except Exception as e:
            results.append({
                "fileName": getattr(up, "filename", "unknown"),
                "error": str(e)
            })

    return JSONResponse(content={"files": results})

def get_slide_num(path: str) -> int:
    """
    Extract numeric slide number from path for sorting.
    """
    m = re.search(r"ppt/slides/slide(\d+)\.xml$", path)
    return int(m.group(1)) if m else 10**9

def analyze_powerpoint(file_path, filename):
    """Analyze PowerPoint file for accessibility issues."""
    report = {
        "fileName": filename,
        "summary": {
            "fixed": 0,
            "flagged": 0
        },
        "details": {
            "slidesMissingTitles": [],
            "imagesMissingOrBadAlt": [],
            "gifsDetected": [],
            "listFormattingIssues": [],
            "titleNeedsFixing": False,
            "fileNameNeedsFixing": False
        }
    }

    try:
        with zipfile.ZipFile(file_path, 'r') as zip_file:
            # ---- Title metadata check ----
            if 'docProps/core.xml' in zip_file.namelist():
                core_xml = zip_file.read('docProps/core.xml').decode('utf-8', errors='ignore')
                if '<dc:title/>' in core_xml or '<dc:title></dc:title>' in core_xml:
                    report["details"]["titleNeedsFixing"] = True
                    report["summary"]["flagged"] += 1

            # ---- File name check ----
            if "_" in filename or filename.lower().startswith("presentation") or filename.lower().startswith("untitled"):
                report["details"]["fileNameNeedsFixing"] = True
                report["summary"]["flagged"] += 1

            # ---- Collect slides in TRUE numeric order ----
            slides = [
                name for name in zip_file.namelist()
                if name.startswith("ppt/slides/slide") and name.endswith(".xml")
            ]
            slides = sorted(slides, key=get_slide_num)

            # ---- Analyze each slide in presentation order ----
            for slide_path in slides:
                slide_number = get_slide_num(slide_path)
                slide_xml = zip_file.read(slide_path).decode('utf-8', errors='ignore')

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

                # Check list formatting
                list_issues = check_list_formatting(slide_xml, slide_number)
                if list_issues:
                    report["details"]["listFormattingIssues"].extend(list_issues)
                    report["summary"]["flagged"] += len(list_issues)

            # ---- GIF check ----
            gif_files = [
                name for name in zip_file.namelist()
                if name.startswith("ppt/media/") and name.lower().endswith(".gif")
            ]
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
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

def extract_image_from_pptx_slide(
    pptx_path: Path,
    slide_number: int,
    rel_id: str
) -> Optional[bytes]:
    """
    Extract image data from PowerPoint using relationship ID
    
    Args:
        pptx_path: Path to the PowerPoint file
        slide_number: Slide number (1-indexed)
        rel_id: Relationship ID (e.g., 'rId2')
        
    Returns:
        Image bytes or None if not found
    """
    try:
        with zipfile.ZipFile(pptx_path, 'r') as zip_ref:
            # Get relationship file for this slide
            rels_path = f'ppt/slides/_rels/slide{slide_number}.xml.rels'
            
            if rels_path not in zip_ref.namelist():
                return None
            
            rels_xml = zip_ref.read(rels_path).decode('utf-8')
            
            # Find the target for this relationship ID
            # <Relationship Id="rId2" Target="../media/image1.png" />
            pattern = rf'<Relationship[^>]*Id="{re.escape(rel_id)}"[^>]*Target="([^"]*)"[^>]*/>'
            match = re.search(pattern, rels_xml)
            
            if not match:
                return None
            
            target = match.group(1)
            # Convert relative path to absolute in ZIP
            if target.startswith('../'):
                media_path = 'ppt/' + target[3:]
            else:
                media_path = target
            
            if media_path in zip_ref.namelist():
                return zip_ref.read(media_path)
                
    except Exception as e:
        print(f"Error extracting image {rel_id} from slide {slide_number}: {e}")
    
    return None

def get_image_rel_id_for_pic(pic_element, namespaces: dict) -> Optional[str]:
    """
    Extract the relationship ID for an image from a p:pic element
    
    Args:
        pic_element: The p:pic XML element
        namespaces: XML namespaces dict
        
    Returns:
        Relationship ID (e.g., 'rId2') or None
    """
    try:
        # Navigate: p:pic -> p:blipFill -> a:blip[@r:embed]
        blip = pic_element.find('.//a:blip[@r:embed]', namespaces)
        if blip is not None:
            return blip.get(f'{{{R_NS}}}embed')
    except Exception as e:
        print(f"Error getting rel ID from pic element: {e}")
    
    return None

def set_alt_text_in_slide_xml(
    slide_xml_bytes: bytes,
    slide_number: int,
    pptx_path: Optional[Path] = None
):
    """
    Finds all picture cNvPr nodes and fixes their 'descr' safely.
    Uses FREE local AI for intelligent alt text generation.
    
    Args:
        slide_xml_bytes: The slide XML as bytes
        slide_number: Slide number (1-indexed)
        pptx_path: Path to the PowerPoint file (needed for AI image extraction)
        
    Returns: (new_xml_bytes, fixed_count, fix_details)
    """
    parser = etree.XMLParser(remove_blank_text=False, recover=False)
    root = etree.fromstring(slide_xml_bytes, parser=parser)

    ns = {
        "p": P_NS,
        "a": A_NS,
        "r": R_NS
    }

    fixed = 0
    fix_details = []
    
    # Check if AI is available and enabled
    use_ai = AI_AVAILABLE and os.getenv("ENABLE_AI_ALT_TEXT", "true").lower() == "true"
    
    if use_ai:
        print(f"🤖 Using FREE local AI (BLIP) for slide {slide_number}")
    else:
        print(f"ℹ️  Using placeholder alt text for slide {slide_number}")

    # Pictures: p:pic -> p:nvPicPr -> p:cNvPr
    pic_elements = root.xpath(".//p:pic", namespaces=ns)
    
    for pic in pic_elements:
        cnvpr = pic.find(".//p:nvPicPr/p:cNvPr", namespaces=ns)
        if cnvpr is None:
            continue
            
        shape_id = cnvpr.get("id") or ""
        shape_name = cnvpr.get("name") or ""
        descr = cnvpr.get("descr")  # can be None
        
        # Get relationship ID for AI image extraction
        rel_id = get_image_rel_id_for_pic(pic, ns) if use_ai and pptx_path else None

        # Decide if we need a fix
        if descr is None or descr.strip() == "":
            new_alt = None
            
            # Try AI generation first
            if use_ai and pptx_path and rel_id:
                try:
                    image_data = extract_image_from_pptx_slide(pptx_path, slide_number, rel_id)
                    if image_data:
                        new_alt = generate_alt_text_free(
                            image_data,
                            shape_name=shape_name,
                            slide_number=slide_number,
                            max_length=ALT_TEXT_MAX
                        )
                        if new_alt:
                            print(f"  ✅ AI generated alt text for {shape_name}: '{new_alt[:50]}...'")
                except Exception as e:
                    print(f"  ⚠️  AI alt text generation failed for {shape_name}: {e}")
            
            # Fallback to placeholder if AI fails or is disabled
            if not new_alt:
                new_alt = choose_default_alt(shape_name, slide_number)
                
            cnvpr.set("descr", new_alt)
            fixed += 1
            fix_details.append({
                "slideNumber": slide_number,
                "shapeId": shape_id,
                "shapeName": shape_name,
                "fix": "addedAltText" if use_ai else "addedPlaceholderAltText",
                "altText": new_alt,
                "aiGenerated": use_ai and rel_id is not None
            })
        
        elif len(descr) > ALT_TEXT_MAX:
            new_alt = None

            if use_ai and pptx_path and rel_id:
                try:
                    image_data = extract_image_from_pptx_slide(pptx_path, slide_number, rel_id)
                    if image_data:
                        new_alt = generate_alt_text_free(
                            image_data,
                            shape_name=shape_name,
                            slide_number=slide_number,
                            max_length=ALT_TEXT_MAX
                        )
                except Exception as e:
                    print(f"AI alt text generation failed for long alt text on {shape_name}: {e}")

            if not new_alt:
                new_alt = descr[:ALT_TEXT_MAX]

            cnvpr.set("descr", new_alt)
            fixed += 1
            fix_details.append({
                "slideNumber": slide_number,
                "shapeId": shape_id,
                "shapeName": shape_name,
                "fix": "replacedLongAltText" if new_alt != descr[:ALT_TEXT_MAX] else "truncatedAltText",
                "altText": new_alt
            })

        else:
            # Check for generic descriptions that could be improved
            descr_lower = descr.lower()
            if descr_lower in ["image", "picture", "photo"]:
                new_alt = None
                
                # Try AI generation for generic descriptions
                if use_ai and pptx_path and rel_id:
                    try:
                        image_data = extract_image_from_pptx_slide(pptx_path, slide_number, rel_id)
                        if image_data:
                            new_alt = generate_alt_text_free(
                                image_data,
                                shape_name=shape_name,
                                slide_number=slide_number,
                                max_length=ALT_TEXT_MAX
                            )
                            if new_alt:
                                print(f"  ✅ AI replaced generic alt text for {shape_name}: '{new_alt[:50]}...'")
                    except Exception as e:
                        print(f"  ⚠️  AI alt text generation failed for {shape_name}: {e}")
                
                # Fallback to placeholder
                if not new_alt:
                    new_alt = f"Image on slide {slide_number}"
                    
                cnvpr.set("descr", new_alt)
                fixed += 1
                fix_details.append({
                    "slideNumber": slide_number,
                    "shapeId": shape_id,
                    "shapeName": shape_name,
                    "fix": "replacedGenericAltText",
                    "altText": new_alt,
                    "aiGenerated": use_ai and rel_id is not None
                })
    new_bytes = etree.tostring(
        root,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=None
    )
    return new_bytes, fixed, fix_details

def remediate_alt_text_pptx(src_pptx: Path, dst_pptx: Path):
    """
    Remediate alt text in PowerPoint file using AI-powered descriptions,
    while processing slides in true numeric presentation order.
    """
    fixed_total = 0
    all_fix_details = []

    print(f"\n🔧 Starting alt text remediation for: {src_pptx.name}")
    print(f"   AI Mode: {os.getenv('ENABLE_AI_ALT_TEXT', 'true')}")

    with zipfile.ZipFile(src_pptx, "r") as zin, zipfile.ZipFile(dst_pptx, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        # Build a lookup of all original zip entries
        info_by_name = {item.filename: item for item in zin.infolist()}

        # Separate slide XMLs from everything else
        slide_names = [
            name for name in info_by_name.keys()
            if re.match(r"ppt/slides/slide\d+\.xml$", name)
        ]
        slide_names = sorted(slide_names, key=get_slide_num)

        non_slide_names = [
            name for name in info_by_name.keys()
            if name not in slide_names
        ]

        # Write non-slide files first exactly as they are
        for name in non_slide_names:
            item = info_by_name[name]
            data = zin.read(name)
            zout.writestr(item, data)

        # Then write slides in true numeric order
        for name in slide_names:
            item = info_by_name[name]
            data = zin.read(name)

            slide_num = get_slide_num(name)
            try:
                new_data, fixed, details = set_alt_text_in_slide_xml(
                    data,
                    slide_num,
                    pptx_path=src_pptx
                )
                if fixed:
                    data = new_data
                    fixed_total += fixed
                    all_fix_details.extend(details)
            except Exception as e:
                print(f"  ⚠️ Error processing slide {slide_num}: {e}")

            zout.writestr(item, data)

    print(f"\n✅ Remediation complete: {fixed_total} images processed")
    ai_count = sum(1 for d in all_fix_details if d.get("aiGenerated", False))
    if ai_count > 0:
        print(f"   🤖 {ai_count} alt texts generated by FREE local AI (no cost)")

    return fixed_total, all_fix_details

@app.get("/download")
def download_all_files():
    candidates = [p for p in OUTPUT_DIR.glob("*") if p.is_file()]
    if not candidates:
        raise HTTPException(status_code=404, detail="No files available to download yet.")

    zip_name = f"remediated-files-{uuid.uuid4().hex[:8]}.zip"
    zip_path = OUTPUT_DIR / zip_name

    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for p in candidates:
            zf.write(p, arcname=p.name)

    return FileResponse(
        path=str(zip_path),
        media_type="application/zip",
        filename="remediated-files.zip"
    )

@app.post("/download")
async def download_selected_files(request: Request):
    body = await request.json()

    file_name = body.get("fileName") or body.get("filename") or body.get("suggestedFileName")
    files = body.get("files", [])

    # Case 1: single file download
    if file_name:
        file_path = OUTPUT_DIR / file_name

        if not file_path.exists():
            matches = list(OUTPUT_DIR.glob(f"*_{file_name}"))
            if matches:
                file_path = matches[0]
            else:
                raise HTTPException(status_code=404, detail=f"File not found: {file_name}")

        return FileResponse(
            path=str(file_path),
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            filename=file_name
        )

    # Case 2: multiple files -> zip
    if files:
        zip_name = f"remediated-files-{uuid.uuid4().hex[:8]}.zip"
        zip_path = OUTPUT_DIR / zip_name

        added_any = False
        with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for name in files:
                file_path = OUTPUT_DIR / name

                # if clean name not found, try prefixed stored file
                if not file_path.exists():
                    matches = list(OUTPUT_DIR.glob(f"*_{name}"))
                    if matches:
                        file_path = matches[0]
                    else:
                        continue

                zf.write(file_path, arcname=name)
                added_any = True

        if not added_any:
            raise HTTPException(status_code=404, detail="None of the requested files were found.")

        return FileResponse(
            path=str(zip_path),
            media_type="application/zip",
            filename="remediated-files.zip"
        )

    raise HTTPException(status_code=400, detail="No file name(s) provided.")

# ---------- RUN ----------
if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=5000)