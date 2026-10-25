"""Reusable python-docx helpers for explicit CJK style fonts and three-line tables.

These helpers implement formatting, not a complete document compliance audit.
"""

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt


def set_style_fonts(style, east_asia, latin="Times New Roman", size_pt=None,
                    bold=None):
    """Set font slots and optional size/weight; omitted bold preserves weight."""
    style.font.name = latin
    fonts = style.element.get_or_add_rPr().get_or_add_rFonts()
    for name in list(fonts.attrib):
        if name.rsplit("}", 1)[-1].lower().endswith("theme"):
            del fonts.attrib[name]
    for slot, value in (("eastAsia", east_asia), ("ascii", latin),
                        ("hAnsi", latin), ("cs", latin)):
        fonts.set(qn("w:" + slot), value)
    if size_pt is not None:
        style.font.size = Pt(size_pt)
        properties = style.element.get_or_add_rPr()
        size_cs = properties.find(qn("w:szCs"))
        if size_cs is None:
            size_cs = OxmlElement("w:szCs")
            properties.insert_element_before(
                size_cs, "w:highlight", "w:u", "w:effect", "w:bdr", "w:shd",
                "w:fitText", "w:vertAlign", "w:rtl", "w:cs", "w:em", "w:lang",
                "w:eastAsianLayout", "w:specVanish", "w:oMath"
            )
        size_cs.set(qn("w:val"), str(round(float(size_pt) * 2)))
    if bold is not None:
        style.font.bold = bold
        style.font.cs_bold = bold


def configure_body_style(document, size_pt=12, style_name="Normal",
                         east_asia="宋体", latin="Times New Roman"):
    """Initialize existing body styles; return the chosen paragraph style.

Use the resolved profile size. This changes fonts and size only; paragraphs
must then be bound with apply_paragraph_style to clear direct overrides.
"""
    selected = document.styles[style_name]
    for name in dict.fromkeys(("Normal", "Body Text", style_name)):
        if name in document.styles:
            set_style_fonts(document.styles[name], east_asia, latin, size_pt)
    return selected


def apply_paragraph_style(paragraph, style, clear_bold_override=False):
    """Bind a style and remove direct font/size overrides; preserve emphasis.

Use clear_bold_override only when normalizing a heading's style-controlled weight.
Character styles and mathematical objects require separate review by the caller.
"""
    paragraph.style = style
    for properties in paragraph._p.xpath(".//w:r/w:rPr | ./w:pPr/w:rPr"):
        names = ("rFonts", "sz", "szCs")
        if clear_bold_override:
            names += ("b", "bCs")
        for name in names:
            for node in list(properties.findall(qn("w:" + name))):
                properties.remove(node)


def _replace_borders(properties, element_name):
    for node in list(properties.findall(qn("w:" + element_name))):
        properties.remove(node)
    borders = OxmlElement("w:" + element_name)
    if element_name == "tblBorders":
        properties.insert_element_before(
            borders, "w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook", "w:tblPrChange"
        )
    else:
        properties.insert_element_before(
            borders, "w:shd", "w:noWrap", "w:tcMar", "w:textDirection",
            "w:tcFitText", "w:vAlign", "w:hideMark", "w:tcPrChange"
        )
    for side in ("top", "left", "start", "bottom", "right", "end", "insideH", "insideV"):
        edge = OxmlElement("w:" + side)
        edge.set(qn("w:val"), "nil")
        borders.append(edge)
    return borders


def _draw(borders, side, size_pt):
    edge = borders.find(qn("w:" + side))
    edge.set(qn("w:val"), "single")
    edge.set(qn("w:sz"), str(round(size_pt * 8)))
    edge.set(qn("w:color"), "000000")
    edge.set(qn("w:space"), "0")


def _clear_shading(properties):
    for node in list(properties.findall(qn("w:shd"))):
        properties.remove(node)


def three_line_table(table, header_rows=1, outer_pt=1.5, inner_pt=0.75):
    """Format a nonmerged academic table; preserve data and existing dimensions.

Merged cells need individual handling and are rejected before any mutation.
"""
    count = len(table.rows)
    if not isinstance(header_rows, int) or not 1 <= header_rows < count:
        raise ValueError("Use at least one header row and one data row.")
    if outer_pt <= 0 or inner_pt <= 0:
        raise ValueError("Border widths must be positive.")
    if table._tbl.xpath("./w:tr/w:tc/w:tcPr/w:gridSpan | ./w:tr/w:tc/w:tcPr/w:vMerge | ./w:tr/w:tc/w:tcPr/w:hMerge"):
        raise ValueError("Merged table: handle borders individually; do not unmerge data.")

    table.style = "Normal Table"
    table_properties = table._tbl.tblPr
    _replace_borders(table_properties, "tblBorders")
    _clear_shading(table_properties)
    # Row-level table-property exceptions can reintroduce inherited borders.
    for exception in table._tbl.xpath("./w:tr/w:tblPrEx"):
        for name in ("tblBorders", "shd"):
            for node in list(exception.findall(qn("w:" + name))):
                exception.remove(node)

    for index, row in enumerate(table.rows):
        row_properties = row._tr.get_or_add_trPr()
        for node in list(row_properties.findall(qn("w:tblHeader"))):
            row_properties.remove(node)
        if index < header_rows:
            repeat = OxmlElement("w:tblHeader")
            row_properties.insert_element_before(
                repeat, "w:tblCellSpacing", "w:jc", "w:hidden", "w:ins", "w:del", "w:trPrChange"
            )
        for cell in row.cells:
            cell_properties = cell._tc.get_or_add_tcPr()
            borders = _replace_borders(cell_properties, "tcBorders")
            _clear_shading(cell_properties)
            if index == 0:
                _draw(borders, "top", outer_pt)
            if index == header_rows - 1:
                _draw(borders, "bottom", inner_pt)
            if index == count - 1:
                _draw(borders, "bottom", outer_pt)
