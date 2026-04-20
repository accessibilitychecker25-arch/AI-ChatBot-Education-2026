from pathlib import Path
import io
import zipfile
import sys

sys.path.insert(0, str(Path(__file__).resolve().parent))

from color_contrast import build_pptx_color_context, check_slide_color_contrast, remediate_slide_color_contrast


def main():
    xml = b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="TextBox 1"/></p:nvSpPr>
        <p:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr sz="1200"><a:solidFill><a:srgbClr val="777777"/></a:solidFill></a:rPr>
              <a:t>Test text</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>'''

    with io.BytesIO() as b:
        with zipfile.ZipFile(b, 'w'):
            pass
        b.seek(0)
        with zipfile.ZipFile(b, 'r') as z:
            ctx = build_pptx_color_context(z)

    issues = check_slide_color_contrast(xml, 1, ctx)
    assert len(issues) == 1, f"expected 1 issue, got {len(issues)}"
    assert issues[0]["type"] == "colorContrast"

    new_xml, fixed, details = remediate_slide_color_contrast(xml, 1, ctx)
    assert fixed == 1, f"expected 1 fix, got {fixed}"
    assert details[0]["afterColor"] == "#767676", details[0]
    assert b'val="767676"' in new_xml, "expected remediated XML to contain new color"

    print("PASS: color contrast detection and remediation smoke test")


if __name__ == "__main__":
    main()
