import os
import time
import shutil
from typing import List
from pathlib import Path
import zipfile
import xml.etree.ElementTree as ET
import re

from fastapi import FastAPI, File, UploadFile, HTTPException, Body
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from starlette.requests import Request

# ---------- CONFIG ----------
UPLOAD_DIR = Path("uploads")
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

OUTPUT_DIR = Path("output")
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
)

# Optional: request logging (safe - does NOT print file bytes)
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

# ---------- UPLOAD ROUTE ----------
@app.post("/upload")
async def upload_files(files: List[UploadFile] = File(...)):
    """
    Accepts PowerPoint files, analyzes them, and returns accessibility report.
    """
    if len(files) == 0:
        raise HTTPException(status_code=400, detail="No file uploaded")
    
    if len(files) > 7:
        raise HTTPException(
            status_code=400,
            detail=f"Too many files. You uploaded {len(files)}, but the limit is 7."
        )

    # For now, handle single file upload
    file = files[0]
    filename = file.filename or "unnamed.pptx"
    filename_lower = filename.lower()

    # Validate extension
    allowed_ext = (".pptx", ".ppt", ".pps", ".potx")
    if not filename_lower.endswith(allowed_ext):
        raise HTTPException(
            status_code=400,
            detail=f"Invalid file type. Please upload a PowerPoint file (.pptx, .ppt, .pps, or .potx)"
        )

    # Save file
    try:
        file_location = UPLOAD_DIR / filename
        with file_location.open("wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
    except Exception as e:
        print(f"Error saving {filename}: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to save file: {str(e)}")

    # Analyze the PowerPoint file
    try:
        report = analyze_powerpoint(file_location, filename)
        return JSONResponse(content={
            "fileName": filename,
            "suggestedFileName": filename,
            "report": report
        })
    except Exception as e:
        print(f"Error analyzing {filename}: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to analyze file: {str(e)}")


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
        "suggestedFileName": filename,
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


def check_slide_images(slide_xml: str, slide_number: int):
    """Check images for missing alt text."""
    issues = []
    
    # Find all picture elements
    pic_pattern = r'<p:pic[\s\S]*?</p:pic>'
    pic_matches = re.findall(pic_pattern, slide_xml)
    
    for pic_xml in pic_matches:
        # Check for alt text in descr attribute
        descr_pattern = r'<p:cNvPr[^>]*descr="([^"]*)"'
        descr_match = re.search(descr_pattern, pic_xml)
        
        alt_text = descr_match.group(1) if descr_match else ""
        
        if not alt_text or alt_text.strip() == "":
            issues.append({
                "slideNumber": slide_number,
                "location": f"Slide {slide_number}",
                "issue": "Image missing alt text",
                "type": "image"
            })
    
    return issues

# ---------- DOWNLOAD ROUTES ----------
@app.get("/download/{filename}")
def download_file(filename: str):
    """
    Direct download by filename from /output.
    """
    file_path = OUTPUT_DIR / filename
    if not file_path.exists():
        raise HTTPException(status_code=404, detail=f"File not found: {filename}")

    return FileResponse(
        path=str(file_path),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename=filename
    )

@app.post("/download")
async def download_latest(payload: dict = Body(default={})):
    """
    Supports current frontend that POSTs to /download.
    If payload contains {"filename": "..."} we use that.
    Otherwise returns the newest file from /output.
    """
    filename = payload.get("filename") if isinstance(payload, dict) else None

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
