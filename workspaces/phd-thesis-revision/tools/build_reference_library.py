from __future__ import annotations

import csv
import hashlib
import json
import ssl
import subprocess
import time
import urllib.parse
import urllib.error
import urllib.request
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parent.parent
CONFIG_PATH = ROOT / "config" / "reference_queries.json"
OUTPUT_DIR = ROOT / "output" / "references"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


USER_AGENT = "CodexReferenceBuilder/1.0"
INSECURE_SSL_CONTEXT = ssl.create_default_context()
INSECURE_SSL_CONTEXT.check_hostname = False
INSECURE_SSL_CONTEXT.verify_mode = ssl.CERT_NONE


def http_get(url: str, pause: float = 0.34) -> bytes:
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            data = resp.read()
    except urllib.error.URLError:
        try:
            with urllib.request.urlopen(req, timeout=30, context=INSECURE_SSL_CONTEXT) as resp:
                data = resp.read()
        except urllib.error.URLError:
            result = subprocess.run(
                [
                    "curl.exe",
                    "-L",
                    "-sS",
                    "-A",
                    USER_AGENT,
                    url,
                ],
                capture_output=True,
                check=True,
                text=False,
            )
            data = result.stdout
    time.sleep(pause)
    return data


def http_post(url: str, payload: bytes, content_type: str) -> bytes:
    req = urllib.request.Request(
        url,
        data=payload,
        headers={
            "User-Agent": USER_AGENT,
            "Content-Type": content_type,
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            return resp.read()
    except urllib.error.URLError:
        try:
            with urllib.request.urlopen(req, timeout=60, context=INSECURE_SSL_CONTEXT) as resp:
                return resp.read()
        except urllib.error.URLError:
            result = subprocess.run(
                [
                    "curl.exe",
                    "-sS",
                    "-X",
                    "POST",
                    "-H",
                    f"Content-Type: {content_type}",
                    "-A",
                    USER_AGENT,
                    "--data-binary",
                    "@-",
                    url,
                ],
                input=payload,
                capture_output=True,
                check=True,
                text=False,
            )
            return result.stdout


def normalize_space(text: str) -> str:
    return " ".join((text or "").replace("\n", " ").split())


def text_or_none(node: ET.Element | None, path: str) -> str:
    if node is None:
        return ""
    child = node.find(path)
    return normalize_space("".join(child.itertext())) if child is not None else ""


def safe_key(text: str) -> str:
    cleaned = "".join(ch for ch in text if ch.isalnum())
    return cleaned[:40] or hashlib.md5(text.encode("utf-8")).hexdigest()[:8]


@dataclass
class ReferenceItem:
    citation_key: str
    item_type: str
    title: str
    year: str = ""
    journal: str = ""
    volume: str = ""
    issue: str = ""
    pages: str = ""
    doi: str = ""
    pmid: str = ""
    abstract: str = ""
    publisher: str = ""
    place: str = ""
    url: str = ""
    authors: list[dict[str, str]] = field(default_factory=list)
    keywords: list[str] = field(default_factory=list)
    source_query: str = ""
    source_collection: str = ""

    def dedupe_id(self) -> str:
        if self.doi:
            return f"doi:{self.doi.lower()}"
        if self.pmid:
            return f"pmid:{self.pmid}"
        return f"title:{normalize_space(self.title).lower()}"

    def bibtex_type(self) -> str:
        return "article" if self.item_type == "article" else "techreport"

    def author_bibtex(self) -> str:
        names: list[str] = []
        for person in self.authors:
            if person.get("literal"):
                names.append(person["literal"])
            elif person.get("family"):
                given = person.get("given", "")
                if given:
                    names.append(f"{person['family']}, {given}")
                else:
                    names.append(person["family"])
        return " and ".join(names)

    def to_bibtex(self) -> str:
        fields: list[tuple[str, str]] = [
            ("title", self.title),
            ("author", self.author_bibtex()),
        ]
        if self.journal:
            fields.append(("journal", self.journal))
        if self.year:
            fields.append(("year", self.year))
        if self.volume:
            fields.append(("volume", self.volume))
        if self.issue:
            fields.append(("number", self.issue))
        if self.pages:
            fields.append(("pages", self.pages))
        if self.publisher:
            fields.append(("publisher", self.publisher))
        if self.place:
            fields.append(("address", self.place))
        if self.doi:
            fields.append(("doi", self.doi))
        if self.url:
            fields.append(("url", self.url))
        if self.pmid:
            fields.append(("pmid", self.pmid))
        if self.keywords:
            fields.append(("keywords", ", ".join(self.keywords)))

        body = ",\n".join(f"  {k} = {{{v}}}" for k, v in fields if v)
        return f"@{self.bibtex_type()}{{{self.citation_key},\n{body}\n}}\n"

    def to_csv_row(self) -> dict[str, str]:
        return {
            "citation_key": self.citation_key,
            "item_type": self.item_type,
            "title": self.title,
            "year": self.year,
            "journal": self.journal,
            "volume": self.volume,
            "issue": self.issue,
            "pages": self.pages,
            "doi": self.doi,
            "pmid": self.pmid,
            "keywords": "; ".join(self.keywords),
            "source_collection": self.source_collection,
            "source_query": self.source_query,
            "url": self.url,
        }


def parse_pubmed_article(article: ET.Element, collection_tag: str, query_label: str) -> ReferenceItem | None:
    medline = article.find("MedlineCitation")
    article_node = medline.find("Article") if medline is not None else None
    if article_node is None:
        return None

    title = text_or_none(article_node, "ArticleTitle")
    if not title:
        return None

    journal = text_or_none(article_node, "Journal/ISOAbbreviation") or text_or_none(article_node, "Journal/Title")
    journal_issue = article_node.find("Journal/JournalIssue")
    pub_year = ""
    if journal_issue is not None:
        pub_date = journal_issue.find("PubDate")
        if pub_date is not None:
            pub_year = text_or_none(pub_date, "Year")
            if not pub_year:
                medline_date = text_or_none(pub_date, "MedlineDate")
                pub_year = medline_date[:4]

    volume = text_or_none(journal_issue, "Volume") if journal_issue is not None else ""
    issue = text_or_none(journal_issue, "Issue") if journal_issue is not None else ""
    pages = text_or_none(article_node, "Pagination/MedlinePgn")

    abstract_parts = []
    abstract_node = article_node.find("Abstract")
    if abstract_node is not None:
        for child in abstract_node.findall("AbstractText"):
            label = child.attrib.get("Label", "")
            text = normalize_space("".join(child.itertext()))
            if label and text:
                abstract_parts.append(f"{label}: {text}")
            elif text:
                abstract_parts.append(text)
    abstract = " ".join(abstract_parts)

    pmid = text_or_none(medline, "PMID")
    doi = ""
    pubmed_data = article.find("PubmedData")
    if pubmed_data is not None:
        for article_id in pubmed_data.findall(".//ArticleId"):
            if article_id.attrib.get("IdType") == "doi":
                doi = normalize_space("".join(article_id.itertext()))
                break

    authors: list[dict[str, str]] = []
    for author in article_node.findall("AuthorList/Author"):
        collective = text_or_none(author, "CollectiveName")
        if collective:
            authors.append({"literal": collective})
            continue
        family = text_or_none(author, "LastName")
        given = text_or_none(author, "ForeName")
        if family:
            authors.append({"family": family, "given": given})

    lead = authors[0]["family"] if authors and authors[0].get("family") else "Ref"
    year = pub_year or "n.d."
    citation_key = safe_key(f"{lead}{year}{title}")

    url = f"https://pubmed.ncbi.nlm.nih.gov/{pmid}/" if pmid else ""
    return ReferenceItem(
        citation_key=citation_key,
        item_type="article",
        title=title,
        year=pub_year,
        journal=journal,
        volume=volume,
        issue=issue,
        pages=pages,
        doi=doi,
        pmid=pmid,
        abstract=abstract,
        url=url,
        authors=authors,
        keywords=[collection_tag, query_label],
        source_query=query_label,
        source_collection=collection_tag,
    )


def pubmed_search(term: str, retmax: int) -> list[str]:
    url = (
        "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi?"
        + urllib.parse.urlencode(
            {
                "db": "pubmed",
                "retmode": "json",
                "retmax": str(retmax),
                "sort": "relevance",
                "term": term,
            }
        )
    )
    payload = json.loads(http_get(url).decode("utf-8"))
    return payload.get("esearchresult", {}).get("idlist", [])


def pubmed_fetch(pmids: list[str], collection_tag: str, query_label: str) -> list[ReferenceItem]:
    if not pmids:
        return []
    url = (
        "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi?"
        + urllib.parse.urlencode(
            {
                "db": "pubmed",
                "retmode": "xml",
                "id": ",".join(pmids),
            }
        )
    )
    root = ET.fromstring(http_get(url))
    items: list[ReferenceItem] = []
    for article in root.findall("PubmedArticle"):
        item = parse_pubmed_article(article, collection_tag, query_label)
        if item is not None:
            items.append(item)
    return items


def manual_to_item(raw: dict[str, Any], collection_tag: str) -> ReferenceItem:
    return ReferenceItem(
        citation_key=raw["citation_key"],
        item_type=raw["item_type"],
        title=raw["title"],
        year=raw.get("year", ""),
        journal=raw.get("journal", ""),
        volume=raw.get("volume", ""),
        issue=raw.get("issue", ""),
        pages=raw.get("pages", ""),
        doi=raw.get("doi", ""),
        pmid=raw.get("pmid", ""),
        publisher=raw.get("publisher", ""),
        place=raw.get("place", ""),
        url=raw.get("url", ""),
        authors=raw.get("authors", []),
        keywords=raw.get("keywords", [collection_tag]),
        source_query="manual",
        source_collection=collection_tag,
    )


def dedupe(items: list[ReferenceItem]) -> list[ReferenceItem]:
    seen: dict[str, ReferenceItem] = {}
    for item in items:
        key = item.dedupe_id()
        if key not in seen:
            seen[key] = item
            continue

        existing = seen[key]
        existing_keywords = sorted(set(existing.keywords + item.keywords))
        existing.keywords = existing_keywords
        if not existing.abstract and item.abstract:
            existing.abstract = item.abstract
        if not existing.doi and item.doi:
            existing.doi = item.doi
        if not existing.url and item.url:
            existing.url = item.url
    return sorted(seen.values(), key=lambda x: (x.source_collection, x.year, x.title.lower()))


def write_outputs(items: list[ReferenceItem]) -> None:
    json_path = OUTPUT_DIR / "reference_library.json"
    bib_path = OUTPUT_DIR / "reference_library.bib"
    csv_path = OUTPUT_DIR / "reference_library.csv"
    review_bib_path = OUTPUT_DIR / "review_collection.bib"
    body_bib_path = OUTPUT_DIR / "body_collection.bib"

    json_path.write_text(
        json.dumps([item.__dict__ for item in items], ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    bib_path.write_text("".join(item.to_bibtex() for item in items), encoding="utf-8")
    review_bib_path.write_text(
        "".join(item.to_bibtex() for item in items if "thesis-review" in item.keywords),
        encoding="utf-8",
    )
    body_bib_path.write_text(
        "".join(item.to_bibtex() for item in items if "thesis-body" in item.keywords),
        encoding="utf-8",
    )

    fieldnames = list(items[0].to_csv_row().keys()) if items else []
    with csv_path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for item in items:
            writer.writerow(item.to_csv_row())


def zotero_import(bib_text: str) -> Any:
    payload = bib_text.encode("utf-8")
    raw = http_post(
        "http://127.0.0.1:23119/connector/import",
        payload=payload,
        content_type="application/x-bibtex",
    )
    return json.loads(raw.decode("utf-8"))


def main() -> None:
    config = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    items: list[ReferenceItem] = []

    for collection in config["collections"]:
        collection_tag = collection["tag"]
        for query in collection["queries"]:
            pmids = pubmed_search(query["term"], query["retmax"])
            items.extend(pubmed_fetch(pmids, collection_tag=collection_tag, query_label=query["label"]))
        for manual in collection.get("manual_items", []):
            items.append(manual_to_item(manual, collection_tag=collection_tag))

    items = dedupe(items)
    write_outputs(items)

    bib_text = (OUTPUT_DIR / "reference_library.bib").read_text(encoding="utf-8")
    imported = zotero_import(bib_text)
    summary = {
        "item_count": len(items),
        "review_count": sum(1 for item in items if "thesis-review" in item.keywords),
        "body_count": sum(1 for item in items if "thesis-body" in item.keywords),
        "zotero_imported_count": len(imported),
        "output_dir": str(OUTPUT_DIR),
    }
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
