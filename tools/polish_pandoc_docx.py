#!/usr/bin/env python3
"""Apply LaTeX-like Word styles to a Pandoc DOCX.

Pandoc's reference DOCX controls styles, but some files still benefit from a
post-pass because raw LaTeX often lands in generic Body Text paragraphs. This
script modifies only OOXML style/page metadata; it does not rewrite manuscript
text.
"""

from __future__ import annotations

import argparse
import re
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
NS = {"w": W_NS, "wp": WP_NS, "a": A_NS, "pic": PIC_NS}
ET.register_namespace("w", W_NS)
ET.register_namespace("r", R_NS)
ET.register_namespace("wp", WP_NS)
ET.register_namespace("a", A_NS)
ET.register_namespace("pic", PIC_NS)
ET.register_namespace("m", M_NS)


def w(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


def attr(name: str) -> str:
    return f"{{{W_NS}}}{name}"


def child(parent: ET.Element, tag: str) -> ET.Element | None:
    return parent.find(f"w:{tag}", NS)


def ensure_child(parent: ET.Element, tag: str, before: str | None = None) -> ET.Element:
    existing = child(parent, tag)
    if existing is not None:
        return existing

    node = ET.Element(w(tag))
    if before is None:
        parent.append(node)
        return node

    before_name = w(before)
    for index, current in enumerate(list(parent)):
        if current.tag == before_name:
            parent.insert(index, node)
            return node

    parent.append(node)
    return node


def remove_theme_font_attrs(rfonts: ET.Element) -> None:
    for key in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
        rfonts.attrib.pop(attr(key), None)


def set_font(rpr: ET.Element, font_name: str) -> None:
    rfonts = ensure_child(rpr, "rFonts", before="b")
    remove_theme_font_attrs(rfonts)
    for key in ("ascii", "hAnsi", "cs"):
        rfonts.set(attr(key), font_name)


def set_size(rpr: ET.Element, half_points: int) -> None:
    ensure_child(rpr, "sz").set(attr("val"), str(half_points))
    ensure_child(rpr, "szCs").set(attr("val"), str(half_points))


def set_bold(rpr: ET.Element, enabled: bool) -> None:
    for tag in ("b", "bCs"):
        node = child(rpr, tag)
        if enabled and node is None:
            ensure_child(rpr, tag)
        elif not enabled and node is not None:
            rpr.remove(node)


def set_italic(rpr: ET.Element, enabled: bool) -> None:
    for tag in ("i", "iCs"):
        node = child(rpr, tag)
        if enabled and node is None:
            ensure_child(rpr, tag)
        elif not enabled and node is not None:
            rpr.remove(node)


def set_color(rpr: ET.Element, value: str = "000000") -> None:
    color = ensure_child(rpr, "color")
    color.attrib.clear()
    color.set(attr("val"), value)


def set_spacing(ppr: ET.Element, before: int = 0, after: int = 120, line: int = 360) -> None:
    spacing = ensure_child(ppr, "spacing")
    spacing.set(attr("before"), str(before))
    spacing.set(attr("after"), str(after))
    spacing.set(attr("line"), str(line))
    spacing.set(attr("lineRule"), "auto")


def set_justification(ppr: ET.Element, value: str) -> None:
    ensure_child(ppr, "jc").set(attr("val"), value)


def find_style(root: ET.Element, style_id: str) -> ET.Element | None:
    for style in root.findall("w:style", NS):
        if style.get(attr("styleId")) == style_id:
            return style
    return None


def style_parts(style: ET.Element) -> tuple[ET.Element, ET.Element]:
    ppr = ensure_child(style, "pPr", before="rPr")
    rpr = ensure_child(style, "rPr")
    return ppr, rpr


def patch_paragraph_style(
    root: ET.Element,
    style_id: str,
    *,
    font: str,
    size: int,
    justify: str,
    before: int,
    after: int,
    line: int,
    bold: bool = False,
    italic: bool = False,
) -> None:
    style = find_style(root, style_id)
    if style is None:
        return

    ppr, rpr = style_parts(style)
    set_spacing(ppr, before=before, after=after, line=line)
    set_justification(ppr, justify)
    set_font(rpr, font)
    set_size(rpr, size)
    set_bold(rpr, bold)
    set_italic(rpr, italic)
    set_color(rpr)


def patch_character_style(
    root: ET.Element,
    style_id: str,
    *,
    font: str,
    size: int,
    bold: bool = False,
    italic: bool = False,
) -> None:
    style = find_style(root, style_id)
    if style is None:
        return

    rpr = ensure_child(style, "rPr")
    set_font(rpr, font)
    set_size(rpr, size)
    set_bold(rpr, bold)
    set_italic(rpr, italic)
    set_color(rpr)


def patch_styles(xml: bytes, font: str) -> bytes:
    root = ET.fromstring(xml)

    defaults = ensure_child(root, "docDefaults", before="latentStyles")
    rpr_default = ensure_child(ensure_child(defaults, "rPrDefault"), "rPr")
    set_font(rpr_default, font)
    set_size(rpr_default, 20)

    ppr_default = ensure_child(ensure_child(defaults, "pPrDefault"), "pPr")
    set_spacing(ppr_default, before=0, after=120, line=360)
    set_justification(ppr_default, "both")

    body_styles = [
        "Normal",
        "BodyText",
        "FirstParagraph",
        "Bibliography",
        "BlockText",
        "FootnoteText",
        "FootnoteBlockText",
    ]
    for style_id in body_styles:
        patch_paragraph_style(
            root,
            style_id,
            font=font,
            size=20,
            justify="both",
            before=0,
            after=120,
            line=360,
        )

    patch_paragraph_style(root, "Compact", font=font, size=20, justify="both", before=0, after=40, line=360)
    patch_paragraph_style(root, "Caption", font=font, size=18, justify="both", before=0, after=120, line=300)
    patch_paragraph_style(root, "ImageCaption", font=font, size=18, justify="both", before=0, after=120, line=300)
    patch_paragraph_style(root, "TableCaption", font=font, size=18, justify="both", before=0, after=120, line=300)
    patch_paragraph_style(root, "Figure", font=font, size=20, justify="center", before=0, after=120, line=360)

    patch_paragraph_style(root, "Title", font=font, size=32, justify="center", before=0, after=120, line=360, bold=True)
    patch_paragraph_style(root, "Author", font=font, size=20, justify="center", before=0, after=120, line=360)
    patch_paragraph_style(root, "Date", font=font, size=20, justify="center", before=0, after=120, line=360)

    patch_paragraph_style(root, "Heading1", font=font, size=28, justify="left", before=240, after=120, line=360, bold=True)
    patch_paragraph_style(root, "Heading2", font=font, size=24, justify="left", before=200, after=100, line=360, bold=True)
    patch_paragraph_style(root, "Heading3", font=font, size=22, justify="left", before=160, after=80, line=360, bold=True)

    patch_character_style(root, "BodyTextChar", font=font, size=20)
    patch_character_style(root, "TitleChar", font=font, size=32, bold=True)
    patch_character_style(root, "Heading1Char", font=font, size=28, bold=True)
    patch_character_style(root, "Heading2Char", font=font, size=24, bold=True)
    patch_character_style(root, "Heading3Char", font=font, size=22, bold=True)

    hyperlink = find_style(root, "Hyperlink")
    if hyperlink is not None:
        rpr = ensure_child(hyperlink, "rPr")
        set_font(rpr, font)
        set_size(rpr, 20)

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def paragraph_style(paragraph: ET.Element) -> str | None:
    ppr = child(paragraph, "pPr")
    if ppr is None:
        return None
    pstyle = child(ppr, "pStyle")
    if pstyle is None:
        return None
    return pstyle.get(attr("val"))


def has_drawing(paragraph: ET.Element) -> bool:
    return paragraph.find(".//w:drawing", NS) is not None


def has_manual_break(paragraph: ET.Element) -> bool:
    return paragraph.find(".//w:br", NS) is not None


def ensure_paragraph_ppr(paragraph: ET.Element) -> ET.Element:
    ppr = child(paragraph, "pPr")
    if ppr is not None:
        return ppr

    ppr = ET.Element(w("pPr"))
    paragraph.insert(0, ppr)
    return ppr


def paragraph_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.findall(".//w:t", NS))


def page_break_paragraph() -> ET.Element:
    paragraph = ET.Element(w("p"))
    run = ET.SubElement(paragraph, w("r"))
    ET.SubElement(run, w("br"), {attr("type"): "page"})
    return paragraph


def image_name_aliases(image_name: str) -> set[str]:
    path = Path(image_name)
    aliases = {image_name, path.name}

    if path.suffix.lower() in {".pdf", ".png"}:
        swapped = path.with_suffix(".png" if path.suffix.lower() == ".pdf" else ".pdf")
        aliases.add(str(swapped))
        aliases.add(swapped.name)

    return aliases


def insert_pagebreaks_before_targets(
    root: ET.Element,
    headings: set[str],
    image_names: set[str],
) -> None:
    if not headings and not image_names:
        return

    body = root.find("w:body", NS)
    if body is None:
        return

    index = 0
    while index < len(body):
        node = body[index]
        if node.tag != w("p"):
            index += 1
            continue

        style = paragraph_style(node)
        text = paragraph_text(node)
        should_insert = bool(style and style.startswith("Heading") and text in headings)

        if not should_insert and has_drawing(node):
            for drawing in node.findall(".//w:drawing", NS):
                image_name = drawing_image_name(drawing)
                if image_name is not None and image_name in image_names:
                    should_insert = True
                    break

        if should_insert:
            body.insert(index, page_break_paragraph())
            index += 2
        else:
            index += 1


def patch_section(section: ET.Element, line_numbers: bool) -> None:
    pg_sz = ensure_child(section, "pgSz")
    pg_sz.set(attr("w"), "12240")
    pg_sz.set(attr("h"), "15840")

    pg_mar = ensure_child(section, "pgMar")
    pg_mar.set(attr("top"), "1440")
    pg_mar.set(attr("right"), "1440")
    pg_mar.set(attr("bottom"), "1440")
    pg_mar.set(attr("left"), "1440")
    pg_mar.set(attr("header"), "720")
    pg_mar.set(attr("footer"), "720")
    pg_mar.set(attr("gutter"), "0")

    ln_num_type = child(section, "lnNumType")
    if line_numbers:
        if ln_num_type is None:
            ln_num_type = ensure_child(section, "lnNumType")
        ln_num_type.set(attr("countBy"), "1")
        ln_num_type.set(attr("start"), "1")
        ln_num_type.set(attr("restart"), "continuous")
    elif ln_num_type is not None:
        section.remove(ln_num_type)


def parse_includegraphics_limits(tex_path: Path | None) -> dict[str, float]:
    if tex_path is None or not tex_path.exists():
        return {}

    text = tex_path.read_text(encoding="utf-8")
    pattern = re.compile(r"\\includegraphics(?:\[(?P<opts>[^\]]*)\])?\{(?P<path>[^}]+)\}")
    limits: dict[str, float] = {}

    for match in pattern.finditer(text):
        options = match.group("opts") or ""
        image_path = match.group("path")
        height_match = re.search(r"height\s*=\s*([0-9.]+)\s*\\textheight", options)
        if height_match is None:
            continue

        limit = float(height_match.group(1))
        limits[image_path] = limit
        limits[Path(image_path).name] = limit

    return limits


def parse_graphicspath(tex_text: str, tex_dir: Path) -> list[Path]:
    paths = [tex_dir]
    match = re.search(r"\\graphicspath\{(?P<body>(?:\{[^}]+\})+)\}", tex_text)
    if match is None:
        return paths

    for part in re.findall(r"\{([^}]+)\}", match.group("body")):
        path = (tex_dir / part).resolve()
        paths.append(path)
    return paths


def parse_figure_scale(tex_text: str) -> float:
    match = re.search(r"\\newcommand\{\\FigureScale\}\{([0-9.]+)\}", tex_text)
    if match is None:
        return 1.0
    return float(match.group(1))


def parse_linenumbers_enabled(tex_path: Path | None) -> bool:
    if tex_path is None or not tex_path.exists():
        return False

    enabled = False
    for line in tex_path.read_text(encoding="utf-8").splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("%"):
            continue
        if re.search(r"\\nolinenumbers\b", stripped):
            enabled = False
        elif re.search(r"\\linenumbers\b", stripped):
            enabled = True

    return enabled


def pdf_page_size_points(pdf_path: Path) -> tuple[float, float] | None:
    try:
        proc = subprocess.run(
            ["pdfinfo", str(pdf_path)],
            check=True,
            text=True,
            capture_output=True,
        )
    except (OSError, subprocess.CalledProcessError):
        return None

    match = re.search(r"Page size:\s+([0-9.]+)\s+x\s+([0-9.]+)\s+pts", proc.stdout)
    if match is None:
        return None
    return float(match.group(1)), float(match.group(2))


def parse_latex_image_sizes(tex_path: Path | None) -> dict[str, tuple[int, int]]:
    if tex_path is None or not tex_path.exists():
        return {}

    tex_dir = tex_path.parent
    tex_text = tex_path.read_text(encoding="utf-8")
    figure_scale = parse_figure_scale(tex_text)
    search_paths = parse_graphicspath(tex_text, tex_dir)
    image_names = re.findall(r"\\figinclude\{([^}]+)\}", tex_text)
    image_names += re.findall(r"\\includegraphics(?:\[[^\]]*\])?\{([^}]+)\}", tex_text)

    sizes: dict[str, tuple[int, int]] = {}
    for image_name in image_names:
        image_path = None
        for search_path in search_paths:
            candidate = (search_path / image_name).resolve()
            if candidate.exists():
                image_path = candidate
                break

        if image_path is None or image_path.suffix.lower() != ".pdf":
            continue

        page_size = pdf_page_size_points(image_path)
        if page_size is None:
            continue

        width_pt, height_pt = page_size
        # PDF points are 1/72 inch; Word uses 914400 EMU per inch.
        cx = int(round(width_pt / 72.0 * 914400 * figure_scale))
        cy = int(round(height_pt / 72.0 * 914400 * figure_scale))
        sizes[image_name] = (cx, cy)
        sizes[Path(image_name).name] = (cx, cy)
        if image_path.suffix.lower() == ".pdf":
            png_name = str(Path(image_name).with_suffix(".png"))
            sizes[png_name] = (cx, cy)
            sizes[Path(png_name).name] = (cx, cy)

    return sizes


def latex_heading_text(line: str) -> str | None:
    match = re.search(r"\\(?:section|subsection|subsubsection)\*?\{([^}]+)\}", line)
    if match is None:
        return None
    return match.group(1).strip()


def latex_image_name(line: str) -> str | None:
    match = re.search(r"\\figinclude\{([^}]+)\}", line)
    if match is not None:
        return match.group(1).strip()

    match = re.search(r"\\includegraphics(?:\[[^\]]*\])?\{([^}]+)\}", line)
    if match is not None:
        return match.group(1).strip()

    return None


def parse_pagebreak_targets(tex_path: Path | None) -> tuple[set[str], set[str]]:
    if tex_path is None or not tex_path.exists():
        return set(), set()

    headings: set[str] = set()
    image_names: set[str] = set()
    pending_pagebreak = False
    for line in tex_path.read_text(encoding="utf-8").splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("%"):
            continue

        if re.fullmatch(r"\\(?:newpage|clearpage)\*?", stripped):
            pending_pagebreak = True
            continue

        if pending_pagebreak:
            heading = latex_heading_text(stripped)
            if heading is not None:
                headings.add(heading)

            image_name = latex_image_name(stripped)
            if image_name is not None:
                image_names.update(image_name_aliases(image_name))

            pending_pagebreak = False

    return headings, image_names


def drawing_image_name(drawing: ET.Element) -> str | None:
    for node in drawing.findall(".//pic:cNvPr", NS):
        descr = node.get("descr") or node.get("name")
        if descr:
            return descr
    return None


def set_drawing_size(drawing: ET.Element, cx: int, cy: int) -> None:
    extent = drawing.find(".//wp:extent", NS)
    shape_extent = drawing.find(".//a:xfrm/a:ext", NS)
    if extent is not None:
        extent.set("cx", str(cx))
        extent.set("cy", str(cy))
    if shape_extent is not None:
        shape_extent.set("cx", str(cx))
        shape_extent.set("cy", str(cy))


def resize_drawing(
    drawing: ET.Element,
    image_height_limits: dict[str, float],
    latex_image_sizes: dict[str, tuple[int, int]],
) -> None:
    image_name = drawing_image_name(drawing)
    if image_name is None:
        return

    latex_size = latex_image_sizes.get(image_name) or latex_image_sizes.get(Path(image_name).name)
    if latex_size is not None:
        set_drawing_size(drawing, latex_size[0], latex_size[1])
        return

    height_limit = image_height_limits.get(image_name) or image_height_limits.get(Path(image_name).name)
    if height_limit is None:
        return

    extent = drawing.find(".//wp:extent", NS)
    shape_extent = drawing.find(".//a:xfrm/a:ext", NS)
    if extent is None:
        return

    current_cx = int(extent.get("cx", "0"))
    current_cy = int(extent.get("cy", "0"))
    if current_cx <= 0 or current_cy <= 0:
        return

    # This pipeline fixes Word to letter paper with 1 inch margins. EMU:
    # 914400 per inch.
    text_width_emu = int(6.5 * 914400)
    text_height_emu = int(9.0 * 914400)
    max_height_emu = int(text_height_emu * height_limit)

    ratio = current_cy / current_cx
    new_cx = text_width_emu
    new_cy = int(round(new_cx * ratio))
    if new_cy > max_height_emu:
        new_cy = max_height_emu
        new_cx = int(round(new_cy / ratio))

    set_drawing_size(drawing, new_cx, new_cy)


def patch_document(
    xml: bytes,
    image_height_limits: dict[str, float],
    latex_image_sizes: dict[str, tuple[int, int]],
    pagebreak_headings: set[str],
    pagebreak_images: set[str],
    line_numbers: bool,
) -> bytes:
    root = ET.fromstring(xml)
    body_like_styles = {None, "Normal", "BodyText", "FirstParagraph", "Bibliography", "Compact"}

    insert_pagebreaks_before_targets(root, pagebreak_headings, pagebreak_images)

    for section in root.findall(".//w:sectPr", NS):
        patch_section(section, line_numbers)

    for drawing in root.findall(".//w:drawing", NS):
        resize_drawing(drawing, image_height_limits, latex_image_sizes)

    for paragraph in root.findall(".//w:p", NS):
        style = paragraph_style(paragraph)
        ppr = ensure_paragraph_ppr(paragraph)

        if has_drawing(paragraph):
            set_justification(ppr, "center")
        elif has_manual_break(paragraph) and style in body_like_styles:
            set_justification(ppr, "left")
        elif style in body_like_styles:
            set_justification(ppr, "both")

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def patch_font_table(xml: bytes, font: str) -> bytes:
    root = ET.fromstring(xml)
    for font_node in root.findall("w:font", NS):
        if font_node.get(attr("name")) == font:
            return ET.tostring(root, encoding="utf-8", xml_declaration=True)

    font_node = ET.Element(w("font"), {attr("name"): font})
    ET.SubElement(font_node, w("family"), {attr("val"): "roman"})
    ET.SubElement(font_node, w("pitch"), {attr("val"): "variable"})
    root.insert(0, font_node)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def patch_docx(input_path: Path, output_path: Path, font: str, tex_path: Path | None) -> None:
    image_height_limits = parse_includegraphics_limits(tex_path)
    latex_image_sizes = parse_latex_image_sizes(tex_path)
    pagebreak_headings, pagebreak_images = parse_pagebreak_targets(tex_path)
    line_numbers = parse_linenumbers_enabled(tex_path)

    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        staged = tmp_path / "staged.docx"
        shutil.copyfile(input_path, staged)

        replacements: dict[str, bytes] = {}
        with zipfile.ZipFile(staged, "r") as zin:
            for name in zin.namelist():
                data = zin.read(name)
                if name == "word/styles.xml":
                    data = patch_styles(data, font)
                elif name == "word/document.xml":
                    data = patch_document(
                        data,
                        image_height_limits,
                        latex_image_sizes,
                        pagebreak_headings,
                        pagebreak_images,
                        line_numbers,
                    )
                elif name == "word/fontTable.xml":
                    data = patch_font_table(data, font)
                replacements[name] = data

        output_path.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for name, data in replacements.items():
                zout.writestr(name, data)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("input", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--font", default="Latin Modern Roman")
    parser.add_argument("--tex", type=Path)
    args = parser.parse_args()

    patch_docx(args.input, args.output, args.font, args.tex)


if __name__ == "__main__":
    main()
