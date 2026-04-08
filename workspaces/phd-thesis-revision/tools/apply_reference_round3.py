from __future__ import annotations

import json
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph


ROOT = Path(__file__).resolve().parent.parent
FINAL_DIR = ROOT / "output" / "final"
REPORT_PATH = ROOT / "output" / "references" / "reference_round3_report.json"
STRATEGY_PATH = ROOT / "output" / "references" / "citation_strategy.md"


INLINE_CITATION_UPDATES = [
    (
        "本章通过动物感染模型和巨噬细胞感染模型，从体内外两个层面评价了SASP对宿主抗结核效应和炎症反应的影响。",
        "[30,49,92]",
    ),
    (
        "将本章结果与前述临床观察相联系，可以发现二者之间具有较好的方向一致性。",
        "[17,22,60]",
    ),
    (
        "但必须明确指出，本章目前提供的是功能现象与文献机制之间的方向一致性支持，而不是直接的分子机制终证。",
        "[30,56,98]",
    ),
    (
        "本章存在以下局限。第一，实验设计主要聚焦于宿主侧功能现象验证",
        "[30,56,98]",
    ),
    (
        "因此，本章的价值主要在于完成“功能支持性实验”这一关键环节",
        "[24,29,98]",
    ),
]


INSERT_AFTER_TARGETS = {
    "7.1.2 基础机制层面的主要发现": [
        "第五章最重要的贡献，不在于已经完成了SASP作用机制的终局性证明，而在于把临床观察中出现的积极信号推进到了“具有宿主侧功能支撑”的层面。动物和巨噬细胞实验共同提示，SASP能够同时表现出增强胞内清菌和减轻感染相关炎症反应的双重表型，这使其更符合宿主导向治疗而非单纯抗炎辅助用药的逻辑定位[24,29,30]。",
        "但这部分证据当前仍应被克制地理解为“功能支持 + 机制线索”，而不是“xCT机制闭环已被完全证实”。换言之，本论文已经把问题从“是否可能有效”推进到“为何可能有效”，但尚未完成对xCT表达动态、GSH/ROS变化、mycothiol氧化状态以及关键阻断/逆转实验的系统整合，因此机制层面的结论仍需保留明确边界[30,56,98]。",
    ],
    "7.4.2 机制研究深度层面的局限": [
        "从总讨论角度看，机制部分最大的局限并不是“完全没有方向”，而是“证据链尚未闭合”。当前结果已经能够支持SASP与xCT相关宿主调控方向具有较高一致性，但尚不足以区分xCT依赖与非xCT依赖效应，也不足以判定双通路模型中各环节对临床信号的相对贡献。因此，第五章和第七章最稳妥的写法应始终保持一致：强调功能支撑和转化价值，而不把现有结果表述为终证性机制研究[30,56,98]。",
    ],
    "7.5.1 主要结论": [
        "综合全文证据，当前最稳妥的总体判断是：SASP在耐药肺结核中已经形成“值得继续推进”的临床与实验联合证据基础，但尚未达到可以提出确定性临床推荐的程度。临床部分提供了前瞻性、真实世界、以pre-XDR-TB为重点人群的积极信号；基础实验部分则提供了宿主侧功能支撑，这两部分共同构成了本论文最核心的实质性结论[17,22,30]。",
        "进一步说，本论文真正向前推进的，不只是“某个老药是否可用”这个问题，而是把SASP研究放回了更大的宿主导向治疗框架中：未来评价这类策略时，必须同时考虑证据强度边界、机制桥接、宿主异质性以及长期功能结局，而不能只凭近期培养学结果作出过强结论。也正因为如此，第五章和第七章所强调的机制边界，以及第四章对长期毒性和长期功能代价的提醒，实际上共同构成了本论文结论可信度的重要部分[24,29,105-107]。",
    ],
}


def find_paragraph(doc: Document, needle: str):
    for paragraph in doc.paragraphs:
        if needle in paragraph.text:
            return paragraph
    return None


def append_citation(text: str, citation: str) -> str:
    if citation in text:
        return text
    for punct in ("。", "；", ".", "！", "？"):
        if text.endswith(punct):
            return f"{text[:-1]}{citation}{punct}"
    return f"{text}{citation}"


def insert_paragraph_after(paragraph: Paragraph, text: str) -> Paragraph:
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    new_para = Paragraph(new_p, paragraph._parent)
    new_para.text = text
    if paragraph.style is not None:
        new_para.style = paragraph.style
    return new_para


def delete_paragraph(paragraph: Paragraph) -> None:
    element = paragraph._element
    parent = element.getparent()
    if parent is not None:
        parent.remove(element)


def ensure_after(doc: Document, anchor_text: str, text: str) -> bool:
    if find_paragraph(doc, text) is not None:
        return False
    anchor = find_paragraph(doc, anchor_text)
    if anchor is None:
        raise ValueError(f"Anchor not found: {anchor_text}")
    insert_paragraph_after(anchor, text)
    return True


def delete_exact_text(doc: Document, text: str) -> None:
    paragraph = find_paragraph(doc, text)
    if paragraph is not None and paragraph.text == text:
        delete_paragraph(paragraph)


def append_strategy_note() -> None:
    note = """

## 7. 第 3 轮新增处理

### 主体讨论段加固

- 第五章讨论补强了 5 个关键位置的引文：
  - 实验结果总述
  - 与临床发现一致性
  - 机制终证边界
  - 局限性段
  - 方法学与转化价值段

### 第七章总讨论加固

- 在 `7.1.2`、`7.4.2`、`7.5.1` 后新增了 5 段口径更稳的总结性文字。
- 处理目标不是重复结果，而是把以下三点写得更清楚：
  - 第五章是“功能支持”，不是“机制终证”
  - 第七章结论必须与第五章证据边界一致
  - 长期毒性与长期功能结局应被纳入总讨论，而不是只当作附属安全性信息
"""
    text = STRATEGY_PATH.read_text(encoding="utf-8")
    if "## 7. 第 3 轮新增处理" not in text:
        STRATEGY_PATH.write_text(text.rstrip() + note + "\n", encoding="utf-8")


def main() -> int:
    report: list[dict[str, object]] = []

    for docx_path in sorted(FINAL_DIR.glob("*.docx")):
        doc = Document(str(docx_path))
        inline_updates: list[str] = []
        inserted_blocks: list[str] = []

        for needle, citation in INLINE_CITATION_UPDATES:
            paragraph = find_paragraph(doc, needle)
            if paragraph is None:
                raise ValueError(f"Paragraph not found in {docx_path.name}: {needle}")
            updated = append_citation(paragraph.text, citation)
            if updated != paragraph.text:
                paragraph.text = updated
                inline_updates.append(citation)

        for anchor, texts in INSERT_AFTER_TARGETS.items():
            for text in texts:
                delete_exact_text(doc, text)
            for text in reversed(texts):
                if ensure_after(doc, anchor, text):
                    inserted_blocks.append(anchor)

        doc.save(str(docx_path))
        report.append(
            {
                "file": str(docx_path),
                "inline_citation_updates": inline_updates,
                "inserted_blocks": inserted_blocks,
            }
        )

    append_strategy_note()
    REPORT_PATH.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
