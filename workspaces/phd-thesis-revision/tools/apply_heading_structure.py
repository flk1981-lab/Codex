from __future__ import annotations

import re
import sys
from pathlib import Path

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt


H1_STYLE = "Codex Heading 1"
H2_STYLE = "Codex Heading 2"
H3_STYLE = "Codex Heading 3"

FRONT_BACK_MATTER_TITLES = {
    "学位论文原创性声明",
    "学位论文版权使用授权书",
    "目录",
    "参考文献",
    "文献综述",
    "文献综述参考文献",
    "致谢",
    "个人简历",
    "攻读学位期间获得学术成果情况",
    "学位论文答辩委员会组成和答辩决议",
}

CHAPTER_HEADING_RE = re.compile(r"^第[一二三四五六七八九十百零〇两]+章(?:\s|　|$)")
APPENDIX_HEADING_RE = re.compile(r"^附录(?:\d+|[一二三四五六七八九十百零〇两])?(?:\s|　|$)")
NUMBERED_SECTION_RE = re.compile(r"^(\d+)\.(\d+)(?:\.(\d+))?(?:\s|$)")


def set_east_asia_font(style, font_name: str) -> None:
    style.font.name = font_name
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.rFonts
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:eastAsia"), font_name)
    rfonts.set(qn("w:ascii"), font_name)
    rfonts.set(qn("w:hAnsi"), font_name)


def set_outline_level(style, level: int) -> None:
    ppr = style.element.get_or_add_pPr()
    outline = ppr.find(qn("w:outlineLvl"))
    if outline is None:
        outline = OxmlElement("w:outlineLvl")
        ppr.append(outline)
    outline.set(qn("w:val"), str(level))


def ensure_style(doc: Document, name: str, level: int, align: WD_ALIGN_PARAGRAPH):
    try:
        style = doc.styles[name]
    except KeyError:
        style = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
    style.base_style = doc.styles["Normal"]
    set_east_asia_font(style, "黑体")
    style.font.size = Pt(16)
    style.font.bold = True
    style.paragraph_format.alignment = align
    style.paragraph_format.space_before = Pt(12)
    style.paragraph_format.space_after = Pt(6)
    set_outline_level(style, level)
    return style


def classify_heading(text: str) -> str | None:
    t = text.strip().replace("\u3000", " ")
    if not t:
        return None

    if t in FRONT_BACK_MATTER_TITLES:
        return H1_STYLE
    if APPENDIX_HEADING_RE.match(t):
        return H1_STYLE
    if CHAPTER_HEADING_RE.match(t):
        return H1_STYLE

    match = NUMBERED_SECTION_RE.match(t)
    if match:
        return H3_STYLE if match.group(3) else H2_STYLE
    return None


def is_heading_style(style_name: str) -> bool:
    lowered = style_name.lower()
    return lowered.startswith("codex heading") or lowered.startswith("heading")


def main() -> int:
    if len(sys.argv) != 2:
        print("Usage: python apply_heading_structure.py <thesis.docx>")
        return 1

    path = Path(sys.argv[1])
    if not path.exists():
        print(f"File not found: {path}")
        return 1

    doc = Document(str(path))
    ensure_style(doc, H1_STYLE, 0, WD_ALIGN_PARAGRAPH.CENTER)
    ensure_style(doc, H2_STYLE, 1, WD_ALIGN_PARAGRAPH.LEFT)
    ensure_style(doc, H3_STYLE, 2, WD_ALIGN_PARAGRAPH.LEFT)

    updated = 0
    for p in doc.paragraphs:
        style_name = p.style.name
        if style_name.lower().startswith("toc"):
            continue

        new_style = classify_heading(p.text)
        if new_style:
            if style_name != new_style:
                p.style = doc.styles[new_style]
                updated += 1
        elif is_heading_style(style_name):
            p.style = doc.styles["Normal"]
            updated += 1

    doc.save(str(path))
    print(f"Updated heading structure for {updated} paragraphs in {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
