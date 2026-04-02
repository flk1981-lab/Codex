#!/usr/bin/env python3

import argparse
import json
import re
import smtplib
import sys
import time
import zipfile
from datetime import datetime
from email.message import EmailMessage
from html import escape
from pathlib import Path
from typing import List
from xml.sax.saxutils import escape as xml_escape
import socket


CONTENT_TYPES = {
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Convert a markdown TB report to DOCX and send it by SMTP.")
    parser.add_argument("--markdown-path", required=True)
    parser.add_argument("--report-title", required=True)
    parser.add_argument("--subject")
    parser.add_argument("--plain-text-summary", default="")
    parser.add_argument("--config-path", default=str(Path.home() / ".codex" / "secrets" / "tb-report-mail.json"))
    parser.add_argument("--output-root", default=str(Path.home() / "Codex" / "reports" / "tb-research"))
    parser.add_argument("--dry-run", action="store_true")
    return parser.parse_args()


def ensure_exists(path: Path, label: str) -> None:
    if not path.exists():
        raise FileNotFoundError(f"{label} not found: {path}")


def load_config(config_path: Path) -> dict:
    ensure_exists(config_path, "Mail config")
    with config_path.open("r", encoding="utf-8") as handle:
        config = json.load(handle)

    required = [
        "sender_email",
        "recipient_email",
        "smtp_host",
        "port",
        "use_ssl",
        "smtp_username",
        "smtp_password",
    ]
    missing = [key for key in required if key not in config or config[key] in ("", None)]
    if missing:
        raise ValueError(f"Mail config is missing required fields: {', '.join(missing)}")
    return config


def safe_title(text: str) -> str:
    cleaned = re.sub(r"[^0-9A-Za-z\u4e00-\u9fff._-]+", "_", text).strip("_")
    return cleaned or "tb_report"


def markdown_lines_to_blocks(markdown_text: str) -> List[dict]:
    blocks = []
    for raw_line in markdown_text.splitlines():
        line = raw_line.rstrip()
        if not line.strip():
            blocks.append({"type": "blank", "text": ""})
            continue

        heading_match = re.match(r"^(#{1,6})\s+(.*)$", line)
        if heading_match:
            level = len(heading_match.group(1))
            blocks.append({"type": "heading", "level": level, "text": heading_match.group(2).strip()})
            continue

        bullet_match = re.match(r"^\s*[-*]\s+(.*)$", line)
        if bullet_match:
            blocks.append({"type": "bullet", "text": bullet_match.group(1).strip()})
            continue

        ordered_match = re.match(r"^\s*\d+\.\s+(.*)$", line)
        if ordered_match:
            blocks.append({"type": "bullet", "text": ordered_match.group(1).strip()})
            continue

        blocks.append({"type": "paragraph", "text": line.strip()})

    return blocks


def paragraph_xml(text: str, style: str = None, bullet: bool = False) -> str:
    text = xml_escape(text)
    style_xml = f"<w:pStyle w:val=\"{style}\"/>" if style else ""
    num_xml = (
        "<w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr>"
        if bullet
        else ""
    )
    return (
        "<w:p>"
        f"<w:pPr>{style_xml}{num_xml}</w:pPr>"
        "<w:r><w:rPr/><w:t xml:space=\"preserve\">"
        f"{text}"
        "</w:t></w:r>"
        "</w:p>"
    )


def build_document_xml(title: str, markdown_text: str) -> str:
    blocks = [{"type": "heading", "level": 1, "text": title}] + markdown_lines_to_blocks(markdown_text)
    paragraphs = []
    for block in blocks:
        if block["type"] == "blank":
            paragraphs.append(paragraph_xml(""))
        elif block["type"] == "heading":
            style = f"Heading{min(block['level'], 3)}"
            paragraphs.append(paragraph_xml(block["text"], style=style))
        elif block["type"] == "bullet":
            paragraphs.append(paragraph_xml(block["text"], bullet=True))
        else:
            paragraphs.append(paragraph_xml(block["text"]))

    body = "".join(paragraphs)
    return (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<w:document xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\" "
        "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" "
        "xmlns:o=\"urn:schemas-microsoft-com:office:office\" "
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" "
        "xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" "
        "xmlns:v=\"urn:schemas-microsoft-com:vml\" "
        "xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\" "
        "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" "
        "xmlns:w10=\"urn:schemas-microsoft-com:office:word\" "
        "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" "
        "xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" "
        "xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" "
        "xmlns:wpi=\"http://schemas.microsoft.com/office/word/2010/wordprocessingInk\" "
        "xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\" "
        "xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" "
        "mc:Ignorable=\"w14 wp14\">"
        "<w:body>"
        f"{body}"
        "<w:sectPr><w:pgSz w:w=\"11906\" w:h=\"16838\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"708\" w:footer=\"708\" w:gutter=\"0\"/></w:sectPr>"
        "</w:body>"
        "</w:document>"
    )


def make_docx(markdown_path: Path, output_path: Path, title: str) -> None:
    markdown_text = markdown_path.read_text(encoding="utf-8")
    output_path.parent.mkdir(parents=True, exist_ok=True)

    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
"""
    root_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"""
    document_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>
"""
    styles_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:qFormat/>
    <w:rPr><w:b/><w:sz w:val="32"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:qFormat/>
    <w:rPr><w:b/><w:sz w:val="28"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:qFormat/>
    <w:rPr><w:b/><w:sz w:val="24"/></w:rPr>
  </w:style>
</w:styles>
"""
    numbering_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="•"/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
      <w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>
"""
    now = datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")
    core_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>{escape(title)}</dc:title>
  <dc:creator>Codex</dc:creator>
  <cp:lastModifiedBy>Codex</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>
</cp:coreProperties>
"""
    app_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Codex</Application>
</Properties>
"""
    document_xml = build_document_xml(title, markdown_text)

    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types)
        archive.writestr("_rels/.rels", root_rels)
        archive.writestr("docProps/core.xml", core_xml)
        archive.writestr("docProps/app.xml", app_xml)
        archive.writestr("word/document.xml", document_xml)
        archive.writestr("word/styles.xml", styles_xml)
        archive.writestr("word/numbering.xml", numbering_xml)
        archive.writestr("word/_rels/document.xml.rels", document_rels)


def send_email(config: dict, docx_path: Path, subject: str, report_title: str, plain_text_summary: str) -> None:
    message = EmailMessage()
    message["From"] = config["sender_email"]
    message["To"] = config["recipient_email"]
    message["Subject"] = subject or report_title
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    body = "\n".join(
        [
            "Attached is the latest TB learning report in Word format.",
            "",
            plain_text_summary or "This message was generated automatically by Codex.",
            "",
            f"Sent at: {now}",
        ]
    )
    message.set_content(body)

    docx_bytes = docx_path.read_bytes()
    message.add_attachment(
        docx_bytes,
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=docx_path.name,
    )

    smtp_host = config["smtp_host"]
    port = int(config["port"])
    username = config["smtp_username"]
    password = config["smtp_password"]
    use_ssl = bool(config["use_ssl"])

    last_error = None
    for attempt in range(1, 4):
        try:
            if use_ssl:
                with smtplib.SMTP_SSL(smtp_host, port, timeout=60) as smtp:
                    smtp.login(username, password)
                    smtp.send_message(message)
            else:
                with smtplib.SMTP(smtp_host, port, timeout=60) as smtp:
                    smtp.starttls()
                    smtp.login(username, password)
                    smtp.send_message(message)
            return
        except (socket.gaierror, TimeoutError, OSError, smtplib.SMTPException) as exc:
            last_error = exc
            if attempt == 3:
                break
            time.sleep(2 * attempt)

    raise RuntimeError(
        f"Failed to send email via {smtp_host}:{port} after 3 attempts: {last_error}"
    )


def main() -> int:
    args = parse_args()
    markdown_path = Path(args.markdown_path).expanduser().resolve()
    config_path = Path(args.config_path).expanduser().resolve()
    output_root = Path(args.output_root).expanduser().resolve()

    ensure_exists(markdown_path, "Markdown report")
    config = load_config(config_path)

    now = datetime.now()
    date_folder = now.strftime("%Y-%m")
    base_name = f"{now.strftime('%Y-%m-%d_%H%M%S')}_{safe_title(args.report_title)}"
    target_dir = output_root / date_folder
    target_dir.mkdir(parents=True, exist_ok=True)
    docx_path = target_dir / f"{base_name}.docx"

    make_docx(markdown_path, docx_path, args.report_title)
    if args.dry_run:
        print(f"Dry run complete. DOCX generated at {docx_path}")
        print(f"Configured recipient: {config['recipient_email']}")
        return 0

    send_email(config, docx_path, args.subject, args.report_title, args.plain_text_summary)

    print(f"Sent report email to {config['recipient_email']}")
    print(f"Attachment: {docx_path}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        raise
