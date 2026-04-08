#!/usr/bin/env python
from __future__ import annotations

import argparse
import csv
import re
import shutil
import sys
import zipfile
from collections import Counter
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = {"w": W_NS}


@dataclass
class Paragraph:
    text: str
    style_id: str | None
    style_name: str | None
    heading_level: int | None


def ensure_parent(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)


def normalize_space(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def count_words(text: str) -> int:
    cjk = len(re.findall(r"[\u4e00-\u9fff]", text))
    latin = len(re.findall(r"[A-Za-z0-9_]+", text))
    return cjk + latin


def parse_docx(docx_path: Path) -> list[Paragraph]:
    with zipfile.ZipFile(docx_path) as zf:
        document_xml = ET.fromstring(zf.read("word/document.xml"))
        styles_xml = ET.fromstring(zf.read("word/styles.xml"))

    styles = {}
    for style in styles_xml.findall("w:style", XML_NS):
        style_id = style.attrib.get(f"{{{W_NS}}}styleId")
        name_node = style.find("w:name", XML_NS)
        name = name_node.attrib.get(f"{{{W_NS}}}val") if name_node is not None else None
        styles[style_id] = name

    paragraphs: list[Paragraph] = []
    for para in document_xml.findall(".//w:body/w:p", XML_NS):
        style_id = None
        p_style = para.find("w:pPr/w:pStyle", XML_NS)
        if p_style is not None:
            style_id = p_style.attrib.get(f"{{{W_NS}}}val")
        style_name = styles.get(style_id)
        text = extract_paragraph_text(para)
        heading_level = detect_heading_level(style_id, style_name)
        paragraphs.append(
            Paragraph(
                text=text,
                style_id=style_id,
                style_name=style_name,
                heading_level=heading_level,
            )
        )
    return paragraphs


def extract_paragraph_text(para: ET.Element) -> str:
    parts: list[str] = []
    for node in para.iter():
        tag = strip_ns(node.tag)
        if tag == "t":
            parts.append(node.text or "")
        elif tag == "tab":
            parts.append("\t")
        elif tag in {"br", "cr"}:
            parts.append("\n")
    return normalize_space("".join(parts))


def strip_ns(tag: str) -> str:
    if "}" in tag:
        return tag.split("}", 1)[1]
    return tag


def detect_heading_level(style_id: str | None, style_name: str | None) -> int | None:
    candidates = [value for value in (style_id, style_name) if value]
    for value in candidates:
        lowered = value.lower()
        match = re.search(r"heading\s*([1-9])", lowered)
        if match:
            return int(match.group(1))
        match = re.search(r"标题\s*([1-9])", value)
        if match:
            return int(match.group(1))
    return None


def read_docx_text(docx_path: Path) -> str:
    paragraphs = parse_docx(docx_path)
    return "\n".join(p.text for p in paragraphs if p.text)


def cmd_env_check(args: argparse.Namespace) -> int:
    root = Path(args.root).resolve()
    print(f"Workspace: {root}")
    print(f"Python: {sys.executable}")
    print(f"Python version: {sys.version.split()[0]}")
    print(f"Pandoc: {shutil.which('pandoc') or 'NOT_FOUND'}")
    print(f"WPS path hint: {'available in GUI; CLI check skipped'}")
    return 0


def cmd_backup(args: argparse.Namespace) -> int:
    source = Path(args.docx).resolve()
    if not source.exists():
        raise FileNotFoundError(f"Document not found: {source}")
    backup_dir = Path(args.backup_dir).resolve()
    backup_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    target = backup_dir / f"{source.stem}.{timestamp}{source.suffix}"
    shutil.copy2(source, target)
    print(target)
    return 0


def cmd_docx_stats(args: argparse.Namespace) -> int:
    docx = Path(args.docx).resolve()
    paragraphs = parse_docx(docx)
    text = "\n".join(p.text for p in paragraphs if p.text)
    heading_counts = Counter(p.heading_level for p in paragraphs if p.heading_level)
    non_empty = [p for p in paragraphs if p.text]

    lines = [
        f"# 文档统计：{docx.name}",
        "",
        f"- 段落总数：{len(paragraphs)}",
        f"- 非空段落：{len(non_empty)}",
        f"- 估算字数：{count_words(text)}",
        f"- 一级标题数：{heading_counts.get(1, 0)}",
        f"- 二级标题数：{heading_counts.get(2, 0)}",
        f"- 三级标题数：{heading_counts.get(3, 0)}",
    ]
    for level in sorted(level for level in heading_counts if level and level > 3):
        lines.append(f"- {level}级标题数：{heading_counts[level]}")
    write_text_output(args.output, "\n".join(lines) + "\n")
    print(Path(args.output).resolve())
    return 0


def cmd_outline(args: argparse.Namespace) -> int:
    docx = Path(args.docx).resolve()
    paragraphs = parse_docx(docx)
    lines = [f"# 论文大纲：{docx.name}", ""]
    for index, para in enumerate(paragraphs, start=1):
        if para.heading_level and para.text:
            indent = "  " * (para.heading_level - 1)
            lines.append(f"{indent}- L{para.heading_level} | P{index} | {para.text}")
    write_text_output(args.output, "\n".join(lines) + "\n")
    print(Path(args.output).resolve())
    return 0


def cmd_split_docx(args: argparse.Namespace) -> int:
    docx = Path(args.docx).resolve()
    outdir = Path(args.outdir).resolve()
    outdir.mkdir(parents=True, exist_ok=True)

    paragraphs = parse_docx(docx)
    chapters: list[tuple[str, list[str]]] = []
    current_title = "front_matter"
    current_lines: list[str] = []

    for para in paragraphs:
        if para.heading_level == 1 and para.text:
            if current_lines:
                chapters.append((current_title, current_lines))
            current_title = para.text
            current_lines = [f"# {para.text}", ""]
            continue
        if para.heading_level and para.text:
            hashes = "#" * min(para.heading_level, 6)
            current_lines.append(f"{hashes} {para.text}")
            current_lines.append("")
        elif para.text:
            current_lines.append(para.text)
            current_lines.append("")

    if current_lines:
        chapters.append((current_title, current_lines))

    written = []
    for index, (title, lines) in enumerate(chapters, start=1):
        filename = f"{index:02d}-{slugify(title)}.md"
        target = outdir / filename
        target.write_text("\n".join(lines).strip() + "\n", encoding="utf-8")
        written.append(target)

    print(f"Written {len(written)} files to {outdir}")
    for path in written:
        print(path)
    return 0


def cmd_term_check(args: argparse.Namespace) -> int:
    docx = Path(args.docx).resolve()
    terms = load_terms(Path(args.terms).resolve())
    text = read_docx_text(docx)
    lines = [f"# 术语一致性检查：{docx.name}", ""]

    for term in terms:
        canonical = term["canonical"]
        variants = term["variants"]
        notes = term["notes"]
        canonical_count = text.count(canonical)
        variant_counts = [(variant, text.count(variant)) for variant in variants if variant]
        active_variants = [(variant, count) for variant, count in variant_counts if count > 0]

        lines.append(f"## {canonical}")
        lines.append(f"- 规范写法出现次数：{canonical_count}")
        if active_variants:
            variant_text = ", ".join(f"{variant}={count}" for variant, count in active_variants)
            lines.append(f"- 发现候选变体：{variant_text}")
        else:
            lines.append("- 发现候选变体：无")
        if notes:
            lines.append(f"- 备注：{notes}")
        if canonical_count > 0 and active_variants:
            lines.append("- 状态：需要统一")
        else:
            lines.append("- 状态：通过或未出现")
        lines.append("")

    write_text_output(args.output, "\n".join(lines).rstrip() + "\n")
    print(Path(args.output).resolve())
    return 0


def load_terms(path: Path) -> list[dict[str, object]]:
    rows: list[dict[str, object]] = []
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        for row in reader:
            canonical = (row.get("canonical") or "").strip()
            if not canonical:
                continue
            variants = [item.strip() for item in (row.get("variants") or "").split("|") if item.strip()]
            notes = (row.get("notes") or "").strip()
            rows.append({"canonical": canonical, "variants": variants, "notes": notes})
    return rows


def write_text_output(output: str, content: str) -> None:
    path = Path(output).resolve()
    ensure_parent(path)
    path.write_text(content, encoding="utf-8")


def slugify(text: str) -> str:
    text = normalize_space(text)
    text = re.sub(r"[\\/:*?\"<>|]", "-", text)
    text = text.replace(" ", "-")
    text = re.sub(r"-{2,}", "-", text).strip("-")
    if not text:
        return "untitled"
    return text[:80]


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Utilities for thesis revision workflow.")
    subparsers = parser.add_subparsers(dest="command", required=True)

    env_check = subparsers.add_parser("env-check")
    env_check.add_argument("--root", default=".")
    env_check.set_defaults(func=cmd_env_check)

    backup = subparsers.add_parser("backup")
    backup.add_argument("docx")
    backup.add_argument("--backup-dir", required=True)
    backup.set_defaults(func=cmd_backup)

    stats = subparsers.add_parser("docx-stats")
    stats.add_argument("docx")
    stats.add_argument("--output", required=True)
    stats.set_defaults(func=cmd_docx_stats)

    outline = subparsers.add_parser("outline")
    outline.add_argument("docx")
    outline.add_argument("--output", required=True)
    outline.set_defaults(func=cmd_outline)

    split_docx = subparsers.add_parser("split-docx")
    split_docx.add_argument("docx")
    split_docx.add_argument("--outdir", required=True)
    split_docx.set_defaults(func=cmd_split_docx)

    term_check = subparsers.add_parser("term-check")
    term_check.add_argument("docx")
    term_check.add_argument("--terms", required=True)
    term_check.add_argument("--output", required=True)
    term_check.set_defaults(func=cmd_term_check)

    return parser


def main(argv: Iterable[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(list(argv) if argv is not None else None)
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
