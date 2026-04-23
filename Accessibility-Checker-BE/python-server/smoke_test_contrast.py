
import io
import zipfile

from color_contrast import build_pptx_color_context, check_slide_color_contrast, remediate_slide_color_contrast, parse_xml_bytes, NS

SLIDE_XML = b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="6" name="Text 4"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr sz="1800">
                <a:solidFill><a:srgbClr val="C6C6C6"/></a:solidFill>
              </a:rPr>
              <a:t>Normal body text on a very light gray card.</a:t>
            </a:r>
            <a:r>
              <a:rPr sz="1800">
                <a:solidFill><a:srgbClr val="C6C6C6"/></a:solidFill>
              </a:rPr>
              <a:t>This should fail WCAG AA and be auto-darkened.</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>
'''

def make_minimal_pptx_bytes():
    bio = io.BytesIO()
    with zipfile.ZipFile(bio, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("ppt/slides/slide1.xml", SLIDE_XML)
    bio.seek(0)
    return bio.getvalue()


def test_bad_contrast_slide():
    pptx_bytes = make_minimal_pptx_bytes()
    with zipfile.ZipFile(io.BytesIO(pptx_bytes), "r") as zf:
        context = build_pptx_color_context(zf)
        issues = check_slide_color_contrast(zf.read("ppt/slides/slide1.xml"), 1, context)
        assert issues, "expected pre-fix contrast issues"
        merged_issue = next((i for i in issues if i.get("shapeId") == "6"), None)
        assert merged_issue is not None, "expected merged issue for shapeId 6"
        assert "auto-darkened" in merged_issue.get("text", ""), "expected merged text content"

        new_slide_xml, fixed, details = remediate_slide_color_contrast(zf.read("ppt/slides/slide1.xml"), 1, context)
        assert fixed == 2, "expected 2 run-level fixes"
        merged_fix = next((d for d in details if d.get("shapeId") == "6"), None)
        assert merged_fix is not None, "expected merged fix for shapeId 6"
        assert merged_fix.get("afterContrastRatio", 0) >= merged_fix.get("requiredRatio", 999), "expected final ratio to meet requirement"

    print("PASS: merged issue/fix output and final contrast ratio are present")


if __name__ == "__main__":
    test_bad_contrast_slide()
