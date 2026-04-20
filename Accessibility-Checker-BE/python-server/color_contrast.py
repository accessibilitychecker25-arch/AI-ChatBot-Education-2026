
from __future__ import annotations

import colorsys
import posixpath
from collections import OrderedDict
from typing import Dict, List, Optional, Tuple

from lxml import etree

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

NS = {"p": P_NS, "a": A_NS, "r": R_NS}
RELATIONSHIP_NS = {"rel": REL_NS}

DEFAULT_COLOR_MAP = {
    "bg1": "lt1",
    "tx1": "dk1",
    "bg2": "lt2",
    "tx2": "dk2",
    "accent1": "accent1",
    "accent2": "accent2",
    "accent3": "accent3",
    "accent4": "accent4",
    "accent5": "accent5",
    "accent6": "accent6",
    "hlink": "hlink",
    "folHlink": "folHlink",
}
DEFAULT_THEME_COLORS = {
    "dk1": "000000",
    "lt1": "FFFFFF",
    "dk2": "1F1F1F",
    "lt2": "EEECE1",
    "accent1": "4F81BD",
    "accent2": "C0504D",
    "accent3": "9BBB59",
    "accent4": "8064A2",
    "accent5": "4BACC6",
    "accent6": "F79646",
    "hlink": "0000FF",
    "folHlink": "800080",
}


def _parser() -> etree.XMLParser:
    return etree.XMLParser(remove_blank_text=False, recover=True)


def parse_xml_bytes(xml_bytes: bytes):
    return etree.fromstring(xml_bytes, parser=_parser())


def _local_name(element) -> str:
    return etree.QName(element).localname


def hex_to_rgb(hex_value: str) -> Tuple[int, int, int]:
    value = (hex_value or "").strip().replace("#", "")
    if len(value) == 3:
        value = "".join(ch * 2 for ch in value)
    if len(value) != 6:
        raise ValueError(f"Invalid hex color: {hex_value}")
    return tuple(int(value[i:i + 2], 16) for i in (0, 2, 4))


def rgb_to_hex(rgb: Tuple[int, int, int]) -> str:
    return "{:02X}{:02X}{:02X}".format(*rgb)


def clamp_channel(value: float) -> int:
    return max(0, min(255, int(round(value))))


def srgb_to_linear(channel: int) -> float:
    c = channel / 255.0
    return c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4


def relative_luminance(rgb: Tuple[int, int, int]) -> float:
    r, g, b = (srgb_to_linear(c) for c in rgb)
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def contrast_ratio(fg: Tuple[int, int, int], bg: Tuple[int, int, int]) -> float:
    l1 = relative_luminance(fg)
    l2 = relative_luminance(bg)
    lighter = max(l1, l2)
    darker = min(l1, l2)
    return (lighter + 0.05) / (darker + 0.05)


def is_large_text(font_size_pt: Optional[float], is_bold: bool) -> bool:
    if font_size_pt is None:
        return False
    if is_bold and font_size_pt >= 14:
        return True
    return font_size_pt >= 18


def required_contrast(font_size_pt: Optional[float], is_bold: bool) -> float:
    return 3.0 if is_large_text(font_size_pt, is_bold) else 4.5


def _join_zip_path(base_path: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    base_dir = posixpath.dirname(base_path)
    return posixpath.normpath(posixpath.join(base_dir, target))


def _resolve_relationship_target(zip_ref, source_part: str, rels_path: str, rel_type_suffix: str) -> Optional[str]:
    if rels_path not in zip_ref.namelist():
        return None
    root = parse_xml_bytes(zip_ref.read(rels_path))
    for rel in root.findall("rel:Relationship", namespaces=RELATIONSHIP_NS):
        rel_type = rel.get("Type", "")
        if rel_type.endswith(rel_type_suffix):
            target = rel.get("Target")
            if target:
                return _join_zip_path(source_part, target)
    return None


def _has_non_opaque_alpha(color_element) -> bool:
    for child in color_element:
        if _local_name(child) == "alpha":
            try:
                return int(child.get("val", "100000")) < 100000
            except Exception:
                return True
    return False


def _resolve_scheme_color_name(name: str, context: Dict) -> str:
    mapped = context["color_map"].get(name, name)
    return context["theme_colors"].get(mapped, context["theme_colors"].get(name, context["default_text"]))


def resolve_color_from_color_element(color_element, context: Dict) -> Tuple[Optional[str], Optional[str]]:
    if color_element is None:
        return None, None

    if _has_non_opaque_alpha(color_element):
        return None, "transparentColor"

    local = _local_name(color_element)
    if local == "srgbClr":
        return (color_element.get("val") or "").upper() or None, None
    if local == "sysClr":
        return (color_element.get("lastClr") or "").upper() or None, None
    if local == "schemeClr":
        val = color_element.get("val") or ""
        return _resolve_scheme_color_name(val, context), None
    if local == "prstClr":
        preset = color_element.get("val", "").lower()
        preset_map = {
            "white": "FFFFFF",
            "black": "000000",
            "gray": "808080",
            "grey": "808080",
            "red": "FF0000",
            "green": "008000",
            "blue": "0000FF",
            "yellow": "FFFF00",
        }
        return preset_map.get(preset), None
    return None, "unresolvedColorElement"


def resolve_color_from_fill_parent(parent, context: Dict) -> Tuple[Optional[str], Optional[str]]:
    if parent is None:
        return None, None

    solid_fill = parent.find("a:solidFill", namespaces=NS)
    if solid_fill is not None:
        for child in solid_fill:
            color, reason = resolve_color_from_color_element(child, context)
            if color or reason:
                return color, reason
        return None, "unresolvedSolidFill"

    if parent.find("a:blipFill", namespaces=NS) is not None:
        return None, "imageFill"
    if parent.find("a:gradFill", namespaces=NS) is not None:
        return None, "gradientFill"
    if parent.find("a:pattFill", namespaces=NS) is not None:
        return None, "patternFill"
    if parent.find("a:noFill", namespaces=NS) is not None:
        return None, "transparentFill"

    return None, None


def _extract_background_from_root(root, context: Dict) -> Tuple[Optional[str], Optional[str]]:
    bg_pr = root.find(".//p:cSld/p:bg/p:bgPr", namespaces=NS)
    if bg_pr is not None:
        color, reason = resolve_color_from_fill_parent(bg_pr, context)
        if color or reason:
            return color, reason

    bg_ref = root.find(".//p:cSld/p:bg/p:bgRef", namespaces=NS)
    if bg_ref is not None:
        for child in bg_ref:
            color, reason = resolve_color_from_color_element(child, context)
            if color or reason:
                return color, reason
        return None, "backgroundReference"

    return None, None


def _build_slide_background_map(zip_ref, context: Dict) -> Dict[str, Dict[str, Optional[str]]]:
    background_map: Dict[str, Dict[str, Optional[str]]] = {}
    slide_paths = sorted(
        [n for n in zip_ref.namelist() if n.startswith("ppt/slides/slide") and n.endswith(".xml")]
    )

    for slide_path in slide_paths:
        slide_root = parse_xml_bytes(zip_ref.read(slide_path))
        slide_color, slide_reason = _extract_background_from_root(slide_root, context)
        if slide_color or slide_reason:
            background_map[slide_path] = {"color": slide_color, "reason": slide_reason}
            continue

        rels_path = slide_path.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels"
        layout_path = _resolve_relationship_target(zip_ref, slide_path, rels_path, "/slideLayout")
        layout_color = layout_reason = None
        master_path = None

        if layout_path and layout_path in zip_ref.namelist():
            layout_root = parse_xml_bytes(zip_ref.read(layout_path))
            layout_color, layout_reason = _extract_background_from_root(layout_root, context)
            layout_rels_path = layout_path.replace("ppt/slideLayouts/", "ppt/slideLayouts/_rels/") + ".rels"
            master_path = _resolve_relationship_target(zip_ref, layout_path, layout_rels_path, "/slideMaster")

        master_color = master_reason = None
        if master_path and master_path in zip_ref.namelist():
            master_root = parse_xml_bytes(zip_ref.read(master_path))
            master_color, master_reason = _extract_background_from_root(master_root, context)

        final_color = slide_color or layout_color or master_color or "FFFFFF"
        final_reason = slide_reason or layout_reason or master_reason
        background_map[slide_path] = {"color": final_color, "reason": final_reason}

    return background_map


def build_pptx_color_context(zip_ref) -> Dict:
    theme_colors = dict(DEFAULT_THEME_COLORS)
    color_map = dict(DEFAULT_COLOR_MAP)

    try:
        if "ppt/theme/theme1.xml" in zip_ref.namelist():
            root = parse_xml_bytes(zip_ref.read("ppt/theme/theme1.xml"))
            clr_scheme = root.find(".//a:themeElements/a:clrScheme", namespaces=NS)
            if clr_scheme is not None:
                for child in clr_scheme:
                    local = etree.QName(child).localname
                    srgb = child.find("a:srgbClr", namespaces=NS)
                    sysclr = child.find("a:sysClr", namespaces=NS)
                    if srgb is not None and srgb.get("val"):
                        theme_colors[local] = srgb.get("val").upper()
                    elif sysclr is not None:
                        theme_colors[local] = (sysclr.get("lastClr") or "000000").upper()
    except Exception:
        pass

    try:
        masters = sorted(
            [n for n in zip_ref.namelist() if n.startswith("ppt/slideMasters/slideMaster") and n.endswith(".xml")]
        )
        for master_name in masters[:1]:
            root = parse_xml_bytes(zip_ref.read(master_name))
            clr_map = root.find(".//p:clrMap", namespaces=NS)
            if clr_map is not None:
                for key in list(DEFAULT_COLOR_MAP.keys()):
                    if clr_map.get(key):
                        color_map[key] = clr_map.get(key)
    except Exception:
        pass

    default_text_key = color_map.get("tx1", "dk1")
    default_text = theme_colors.get(default_text_key, theme_colors.get("dk1", "000000"))
    context = {
        "theme_colors": theme_colors,
        "color_map": color_map,
        "default_text": default_text,
    }
    context["slide_backgrounds"] = _build_slide_background_map(zip_ref, context)
    context["slide_path_map"] = {
        int(path.split("slide")[-1].split(".xml")[0]): path
        for path in context["slide_backgrounds"].keys()
        if "slide" in path
    }
    return context


def get_slide_background(slide_number: int, context: Dict) -> Tuple[Optional[str], Optional[str]]:
    slide_path = context.get("slide_path_map", {}).get(slide_number)
    info = context.get("slide_backgrounds", {}).get(slide_path or "", {})
    return info.get("color", "FFFFFF"), info.get("reason")


def describe_shape(shape) -> Tuple[str, str]:
    cnvpr = shape.find(".//p:cNvPr", namespaces=NS)
    shape_id = cnvpr.get("id") if cnvpr is not None and cnvpr.get("id") else ""
    shape_name = cnvpr.get("name") if cnvpr is not None and cnvpr.get("name") else ""
    return shape_id, shape_name


def get_text_style(text_node, context: Dict) -> Tuple[Optional[str], Optional[float], bool, Optional[str], object]:
    rpr = text_node.find("a:rPr", namespaces=NS)
    if rpr is None:
        rpr = text_node.find("a:fldPr", namespaces=NS)

    font_size_pt: Optional[float] = None
    is_bold = False
    color_hex: Optional[str] = None
    unresolved_reason: Optional[str] = None

    if rpr is not None:
        if rpr.get("sz"):
            try:
                font_size_pt = int(rpr.get("sz")) / 100.0
            except Exception:
                font_size_pt = None
        is_bold = rpr.get("b") in {"1", "true", "True"}
        color_hex, unresolved_reason = resolve_color_from_fill_parent(rpr, context)

    if color_hex is None and unresolved_reason is None:
        color_hex = context.get("default_text")

    return color_hex, font_size_pt, is_bold, unresolved_reason, rpr


def _iter_shape_ancestors(node):
    current = node.getparent()
    while current is not None:
        yield current
        current = current.getparent()


def get_shape_background(shape, slide_background_hex: Optional[str], slide_background_reason: Optional[str], context: Dict) -> Tuple[Optional[str], Optional[str]]:
    sppr = shape.find("p:spPr", namespaces=NS)
    if sppr is not None:
        color, reason = resolve_color_from_fill_parent(sppr, context)
        if color:
            return color, None
        if reason and reason not in {"transparentFill", None}:
            return None, reason
        if reason == "transparentFill":
            # try ancestor groups first, then slide background
            pass

    for ancestor in _iter_shape_ancestors(shape):
        if _local_name(ancestor) != "grpSp":
            continue
        grp_sppr = ancestor.find("p:grpSpPr", namespaces=NS)
        if grp_sppr is not None:
            color, reason = resolve_color_from_fill_parent(grp_sppr, context)
            if color:
                return color, None
            if reason and reason not in {"transparentFill", None}:
                return None, f"group{reason[:1].upper()}{reason[1:]}"

    return slide_background_hex, slide_background_reason


def _collect_run_text(paragraph, node) -> str:
    text_node = node.find("a:t", namespaces=NS)
    text = text_node.text if text_node is not None else ""
    return text if text and text.strip() else ""


def get_text_runs_for_shape(shape) -> List[Tuple[object, str, object]]:
    results: List[Tuple[object, str, object]] = []
    for paragraph in shape.findall(".//p:txBody/a:p", namespaces=NS):
        for node in paragraph:
            local = _local_name(node)
            if local in {"r", "fld"}:
                text = _collect_run_text(paragraph, node)
                if text:
                    results.append((node, text, paragraph))
    return results


def get_text_runs_for_table_cell(cell) -> List[Tuple[object, str, object]]:
    results: List[Tuple[object, str, object]] = []
    for paragraph in cell.findall(".//a:txBody/a:p", namespaces=NS):
        for node in paragraph:
            local = _local_name(node)
            if local in {"r", "fld"}:
                text = _collect_run_text(paragraph, node)
                if text:
                    results.append((node, text, paragraph))
    return results


def _manual_issue(
    slide_number: int,
    shape_id: str,
    shape_name: str,
    text: str,
    reason: str,
) -> Dict:
    return {
        "slideNumber": slide_number,
        "shapeId": shape_id,
        "shapeName": shape_name,
        "text": text[:160],
        "issue": "Manual review required for color contrast",
        "type": "colorContrastManualReview",
        "reason": reason,
    }


def _merge_issue_entries(items: List[Dict]) -> List[Dict]:
    merged: "OrderedDict[Tuple, Dict]" = OrderedDict()
    for item in items:
        if item.get("type") == "colorContrast":
            key = (
                item.get("slideNumber"),
                item.get("shapeId"),
                item.get("type"),
                item.get("foregroundColor"),
                item.get("backgroundColor"),
                item.get("requiredRatio"),
                item.get("fontSizePt"),
                item.get("isBold"),
            )
        elif item.get("type") == "colorContrastManualReview":
            key = (
                item.get("slideNumber"),
                item.get("shapeId"),
                item.get("type"),
                item.get("reason"),
            )
        else:
            key = tuple(sorted(item.items()))

        if key not in merged:
            merged[key] = dict(item)
            continue

        existing_text = merged[key].get("text", "")
        new_text = item.get("text", "")
        if new_text and new_text not in existing_text:
            merged[key]["text"] = (existing_text + " " + new_text).strip()[:160]
    return list(merged.values())


def _merge_fix_entries(items: List[Dict]) -> List[Dict]:
    merged: "OrderedDict[Tuple, Dict]" = OrderedDict()
    for item in items:
        key = (
            item.get("slideNumber"),
            item.get("shapeId"),
            item.get("fix"),
            item.get("beforeColor"),
            item.get("afterColor"),
            item.get("backgroundColor"),
            item.get("requiredRatio"),
            item.get("fontSizePt"),
            item.get("isBold"),
        )
        if key not in merged:
            merged[key] = dict(item)
            continue
        existing_text = merged[key].get("text", "")
        new_text = item.get("text", "")
        if new_text and new_text not in existing_text:
            merged[key]["text"] = (existing_text + " " + new_text).strip()[:160]
    return list(merged.values())


def _adjust_lightness(rgb: Tuple[int, int, int], new_l: float) -> Tuple[int, int, int]:
    r, g, b = (c / 255.0 for c in rgb)
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    nr, ng, nb = colorsys.hls_to_rgb(h, max(0.0, min(1.0, new_l)), s)
    return (clamp_channel(nr * 255), clamp_channel(ng * 255), clamp_channel(nb * 255))


def choose_accessible_text_color(
    foreground_rgb: Tuple[int, int, int],
    background_rgb: Tuple[int, int, int],
    required_ratio_value: float,
) -> Optional[Tuple[int, int, int]]:
    current_ratio = contrast_ratio(foreground_rgb, background_rgb)
    if current_ratio >= required_ratio_value:
        return foreground_rgb

    r, g, b = (c / 255.0 for c in foreground_rgb)
    _, lightness, _ = colorsys.rgb_to_hls(r, g, b)

    def search(direction: str) -> Optional[Tuple[float, Tuple[int, int, int]]]:
        low, high = (0.0, lightness) if direction == "darken" else (lightness, 1.0)
        candidate = None
        for _ in range(24):
            mid = (low + high) / 2.0
            test_rgb = _adjust_lightness(foreground_rgb, mid)
            ratio_value = contrast_ratio(test_rgb, background_rgb)
            if ratio_value >= required_ratio_value:
                candidate = (mid, test_rgb)
                if direction == "darken":
                    low = mid
                else:
                    high = mid
            else:
                if direction == "darken":
                    high = mid
                else:
                    low = mid
        return candidate

    candidates = []
    for direction in ("darken", "lighten"):
        result = search(direction)
        if result is not None:
            new_l, new_rgb = result
            candidates.append((abs(new_l - lightness), new_rgb))

    if not candidates:
        black_ratio = contrast_ratio((0, 0, 0), background_rgb)
        white_ratio = contrast_ratio((255, 255, 255), background_rgb)
        if black_ratio >= required_ratio_value or white_ratio >= required_ratio_value:
            return (0, 0, 0) if black_ratio >= white_ratio else (255, 255, 255)
        return None

    candidates.sort(key=lambda item: item[0])
    return candidates[0][1]


def _set_text_color(text_node, new_hex: str):
    rpr = text_node.find("a:rPr", namespaces=NS)
    if rpr is None:
        rpr = etree.Element(f"{{{A_NS}}}rPr")
        text_node.insert(0, rpr)

    for child in list(rpr):
        if _local_name(child) in {"solidFill", "gradFill", "blipFill", "pattFill", "noFill"}:
            rpr.remove(child)

    solid_fill = etree.Element(f"{{{A_NS}}}solidFill")
    srgb = etree.Element(f"{{{A_NS}}}srgbClr")
    srgb.set("val", new_hex.upper())
    solid_fill.append(srgb)
    rpr.insert(0, solid_fill)


def _analyze_runs(
    run_records: List[Tuple[object, str, object]],
    slide_number: int,
    shape_id: str,
    shape_name: str,
    background_hex: Optional[str],
    background_reason: Optional[str],
    context: Dict,
) -> List[Dict]:
    issues: List[Dict] = []
    if background_hex is None:
        preview = " ".join(text for _, text, _ in run_records)[:160]
        if preview:
            issues.append(_manual_issue(slide_number, shape_id, shape_name, preview, background_reason or "unresolvedBackground"))
        return issues

    background_rgb = hex_to_rgb(background_hex)
    for text_node, text, _ in run_records:
        foreground_hex, font_size_pt, is_bold, color_reason, _ = get_text_style(text_node, context)
        if foreground_hex is None:
            issues.append(_manual_issue(slide_number, shape_id, shape_name, text, color_reason or "unresolvedTextColor"))
            continue

        foreground_rgb = hex_to_rgb(foreground_hex)
        needed = required_contrast(font_size_pt, is_bold)
        ratio_value = contrast_ratio(foreground_rgb, background_rgb)
        if ratio_value < needed:
            issues.append({
                "slideNumber": slide_number,
                "shapeId": shape_id,
                "shapeName": shape_name,
                "text": text[:160],
                "issue": "Insufficient color contrast",
                "type": "colorContrast",
                "foregroundColor": f"#{foreground_hex.upper()}",
                "backgroundColor": f"#{background_hex.upper()}",
                "contrastRatio": round(ratio_value, 2),
                "requiredRatio": needed,
                "fontSizePt": round(font_size_pt, 2) if font_size_pt is not None else None,
                "isBold": is_bold,
            })
    return issues


def _remediate_runs(
    run_records: List[Tuple[object, str, object]],
    slide_number: int,
    shape_id: str,
    shape_name: str,
    background_hex: Optional[str],
    background_reason: Optional[str],
    context: Dict,
) -> Tuple[int, List[Dict]]:
    fixed = 0
    fix_details: List[Dict] = []
    if background_hex is None:
        return fixed, fix_details

    background_rgb = hex_to_rgb(background_hex)
    for text_node, text, _ in run_records:
        foreground_hex, font_size_pt, is_bold, color_reason, _ = get_text_style(text_node, context)
        if foreground_hex is None:
            continue

        foreground_rgb = hex_to_rgb(foreground_hex)
        needed = required_contrast(font_size_pt, is_bold)
        before_ratio = contrast_ratio(foreground_rgb, background_rgb)
        if before_ratio >= needed:
            continue

        new_rgb = choose_accessible_text_color(foreground_rgb, background_rgb, needed)
        if new_rgb is None:
            continue

        new_hex = rgb_to_hex(new_rgb)
        if new_hex.upper() == foreground_hex.upper():
            continue

        after_ratio = contrast_ratio(new_rgb, background_rgb)
        _set_text_color(text_node, new_hex)
        fixed += 1
        fix_details.append({
            "slideNumber": slide_number,
            "shapeId": shape_id,
            "shapeName": shape_name,
            "text": text[:160],
            "fix": "adjustedTextColorForContrast",
            "beforeColor": f"#{foreground_hex.upper()}",
            "afterColor": f"#{new_hex.upper()}",
            "backgroundColor": f"#{background_hex.upper()}",
            "beforeContrastRatio": round(before_ratio, 2),
            "afterContrastRatio": round(after_ratio, 2),
            "requiredRatio": needed,
            "fontSizePt": round(font_size_pt, 2) if font_size_pt is not None else None,
            "isBold": is_bold,
        })
    return fixed, fix_details


def check_slide_color_contrast(slide_xml_bytes: bytes, slide_number: int, context: Dict) -> List[Dict]:
    root = parse_xml_bytes(slide_xml_bytes)
    slide_background_hex, slide_background_reason = get_slide_background(slide_number, context)
    issues: List[Dict] = []

    for shape in root.xpath(".//p:sp[p:txBody]", namespaces=NS):
        shape_id, shape_name = describe_shape(shape)
        shape_background_hex, shape_background_reason = get_shape_background(
            shape,
            slide_background_hex,
            slide_background_reason,
            context,
        )
        issues.extend(
            _analyze_runs(
                get_text_runs_for_shape(shape),
                slide_number,
                shape_id,
                shape_name,
                shape_background_hex,
                shape_background_reason,
                context,
            )
        )

    for frame in root.xpath(".//p:graphicFrame[a:graphic/a:graphicData/a:tbl]", namespaces=NS):
        shape_id, shape_name = describe_shape(frame)
        tbl = frame.find(".//a:tbl", namespaces=NS)
        if tbl is None:
            continue
        for idx, cell in enumerate(tbl.findall(".//a:tr/a:tc", namespaces=NS), start=1):
            tc_pr = cell.find("a:tcPr", namespaces=NS)
            cell_color, cell_reason = resolve_color_from_fill_parent(tc_pr, context) if tc_pr is not None else (None, None)
            if cell_reason == "transparentFill" or (cell_color is None and cell_reason is None):
                cell_color, cell_reason = slide_background_hex, slide_background_reason
            issues.extend(
                _analyze_runs(
                    get_text_runs_for_table_cell(cell),
                    slide_number,
                    shape_id,
                    f"{shape_name} cell {idx}",
                    cell_color,
                    cell_reason,
                    context,
                )
            )

    return _merge_issue_entries(issues)


def remediate_slide_color_contrast(slide_xml_bytes: bytes, slide_number: int, context: Dict):
    root = parse_xml_bytes(slide_xml_bytes)
    slide_background_hex, slide_background_reason = get_slide_background(slide_number, context)
    fixed_total = 0
    fix_details: List[Dict] = []

    for shape in root.xpath(".//p:sp[p:txBody]", namespaces=NS):
        shape_id, shape_name = describe_shape(shape)
        shape_background_hex, shape_background_reason = get_shape_background(
            shape,
            slide_background_hex,
            slide_background_reason,
            context,
        )
        fixed, details = _remediate_runs(
            get_text_runs_for_shape(shape),
            slide_number,
            shape_id,
            shape_name,
            shape_background_hex,
            shape_background_reason,
            context,
        )
        fixed_total += fixed
        fix_details.extend(details)

    for frame in root.xpath(".//p:graphicFrame[a:graphic/a:graphicData/a:tbl]", namespaces=NS):
        shape_id, shape_name = describe_shape(frame)
        tbl = frame.find(".//a:tbl", namespaces=NS)
        if tbl is None:
            continue
        for idx, cell in enumerate(tbl.findall(".//a:tr/a:tc", namespaces=NS), start=1):
            tc_pr = cell.find("a:tcPr", namespaces=NS)
            cell_color, cell_reason = resolve_color_from_fill_parent(tc_pr, context) if tc_pr is not None else (None, None)
            if cell_reason == "transparentFill" or (cell_color is None and cell_reason is None):
                cell_color, cell_reason = slide_background_hex, slide_background_reason
            fixed, details = _remediate_runs(
                get_text_runs_for_table_cell(cell),
                slide_number,
                shape_id,
                f"{shape_name} cell {idx}",
                cell_color,
                cell_reason,
                context,
            )
            fixed_total += fixed
            fix_details.extend(details)

    new_bytes = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=None)
    return new_bytes, fixed_total, _merge_fix_entries(fix_details)
