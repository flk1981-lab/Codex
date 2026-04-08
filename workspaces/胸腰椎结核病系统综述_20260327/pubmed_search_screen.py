import csv
import json
import re
import sys
import time
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from pathlib import Path


BASE = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/"
EMAIL = "codex-local@example.com"
TOOL = "codex_systematic_review"

CORE_QUERY = (
    '("Tuberculosis, Spinal"[Mesh] OR spinal tubercul*[tiab] OR spine tubercul*[tiab] '
    'OR tuberculous spondylitis[tiab] OR tuberculous spondylodiscitis[tiab] '
    'OR Pott disease[tiab] OR Potts disease[tiab]) '
    'AND (thoracolumbar[tiab] OR thoraco-lumbar[tiab] OR dorsolumbar[tiab] '
    'OR thoracolumbar junction[tiab])'
)

MANUAL_INCLUDE_PMIDS = {
    "40322305",  # institutional experience + literature review, contains original patient series
    "25953095",  # retrospective pediatric clinical series
    "23812354",  # retrospective thoracolumbar kyphosis series
    "18173344",  # prospective observational study despite PubMed case report tagging
}

MANUAL_EXCLUDE_PMIDS = {
    "38988427",  # sociodemographic cross-sectional description, not diagnosis/treatment-focused
    "40496790",  # general spinal TB description, thoracolumbar data not the main analytical target
    "29198597",  # preoperative vascular anatomy, not diagnosis/treatment outcome study
    "17620817",  # abscess pathway description, not diagnosis/treatment effectiveness
    "13374916",  # abscess pathway description, duplicate historical concept paper
}


def http_get(url: str) -> bytes:
    req = urllib.request.Request(url, headers={"User-Agent": TOOL})
    with urllib.request.urlopen(req, timeout=60) as resp:
        return resp.read()


def esearch(query: str):
    params = {
        "db": "pubmed",
        "term": query,
        "retmode": "json",
        "retmax": "100000",
        "tool": TOOL,
        "email": EMAIL,
    }
    url = BASE + "esearch.fcgi?" + urllib.parse.urlencode(params)
    data = json.loads(http_get(url).decode("utf-8"))
    return data["esearchresult"]["idlist"], int(data["esearchresult"]["count"]), url


def chunked(seq, size):
    for i in range(0, len(seq), size):
        yield seq[i : i + size]


def normalize_space(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "")).strip()


def xml_text(node, path: str) -> str:
    target = node.find(path)
    return normalize_space("".join(target.itertext())) if target is not None else ""


def pub_types(article):
    pts = []
    for pt in article.findall(".//PublicationTypeList/PublicationType"):
        if pt.text:
            pts.append(normalize_space(pt.text))
    return "; ".join(sorted(set(pts)))


def authors(article):
    names = []
    for au in article.findall(".//AuthorList/Author"):
        last = xml_text(au, "LastName")
        fore = xml_text(au, "ForeName")
        coll = xml_text(au, "CollectiveName")
        if coll:
            names.append(coll)
        elif last:
            names.append((last + " " + fore).strip())
    return "; ".join(names[:8])


def abstract_text(article):
    parts = []
    for ab in article.findall(".//Abstract/AbstractText"):
        label = ab.attrib.get("Label", "").strip()
        text = normalize_space("".join(ab.itertext()))
        if text:
            parts.append((label + ": " if label else "") + text)
    return " ".join(parts)


def extract_records(pmids):
    rows = []
    for batch in chunked(pmids, 100):
        params = {
            "db": "pubmed",
            "id": ",".join(batch),
            "retmode": "xml",
            "tool": TOOL,
            "email": EMAIL,
        }
        url = BASE + "efetch.fcgi?" + urllib.parse.urlencode(params)
        xml_bytes = http_get(url)
        root = ET.fromstring(xml_bytes)
        for article in root.findall(".//PubmedArticle"):
            medline = article.find("MedlineCitation")
            pmid = xml_text(article, ".//PMID")
            art = article.find(".//Article")
            title = normalize_space("".join(art.find("ArticleTitle").itertext())) if art is not None and art.find("ArticleTitle") is not None else ""
            journal = xml_text(article, ".//Journal/Title")
            year = ""
            for path in [".//PubDate/Year", ".//ArticleDate/Year", ".//PubMedPubDate[@PubStatus='pubmed']/Year"]:
                year = xml_text(article, path)
                if year:
                    break
            doi = ""
            for el in article.findall(".//ArticleIdList/ArticleId"):
                if el.attrib.get("IdType") == "doi" and el.text:
                    doi = normalize_space(el.text)
                    break
            rows.append(
                {
                    "pmid": pmid,
                    "doi": doi,
                    "title": title,
                    "year": year,
                    "journal": journal,
                    "authors": authors(article),
                    "publication_types": pub_types(article),
                    "abstract": abstract_text(article),
                    "pubmed_url": f"https://pubmed.ncbi.nlm.nih.gov/{pmid}/",
                }
            )
        time.sleep(0.34)
    return rows


def first_pass_decision(row):
    title = (row["title"] or "").lower()
    abstract = (row["abstract"] or "").lower()
    publication_types = (row["publication_types"] or "").lower()
    blob = " ".join([title, abstract, publication_types])
    pmid = row["pmid"]

    if pmid in MANUAL_INCLUDE_PMIDS:
        return "include", "人工复核纳入"
    if pmid in MANUAL_EXCLUDE_PMIDS:
        return "exclude", "人工复核排除"

    if any(term in title for term in ["letter to the editor", "letter to editor", "comment on", "commentary", "editorial", "reply to"]):
        return "exclude", "来信/评论"
    if any(term in publication_types for term in ["review", "systematic review", "meta-analysis"]):
        return "exclude", "综述/非原始研究"
    if any(term in title for term in ["systematic review", "meta-analysis", "scientific mapping", "bibliometric"]):
        return "exclude", "综述/非原始研究"
    if any(term in publication_types for term in ["case reports"]):
        return "exclude", "病例报告"
    if "case report" in title:
        return "exclude", "病例报告"
    if "thoracolumbar fractures" in title:
        return "exclude", "非结核目标疾病"
    if any(term in title for term in ["sociodemographic patterns", "spinal tuberculosis in afghanistan", "position of the aorta relative to the spine", "pathways tracked by dorsolumbar tuberculous abscesses"]):
        return "exclude", "与综述问题不直接相关"
    if any(term in blob for term in ["cervical", "upper thoracic"]) and not any(
        term in blob for term in ["thoracolumbar", "dorsolumbar", "thoracolumbar junction"]
    ):
        return "exclude", "非胸腰段人群"
    if any(term in blob for term in ["animal", "rabbit", "rat model", "cadaver", "finite element", "biomechanical study"]):
        return "exclude", "非临床研究"
    if "tuberc" not in blob and "pott" not in blob:
        return "exclude", "非脊柱结核"
    if not any(term in blob for term in ["diagn", "imaging", "biopsy", "culture", "pcr", "xpert", "surgery", "operative", "treatment", "management", "fusion", "fixation", "debridement", "chemotherapy", "comparative study", "decompression", "stabilization"]):
        return "maybe", "摘要信息不足/待人工判断"
    if not row["abstract"]:
        if any(term in title for term in ["diagnosis", "treatment", "surgical", "surgery", "approach", "operative", "management", "fusion", "osteosynthesis", "comparative study", "decompression", "stabilization"]):
            return "include", "无摘要但标题相关"
        return "maybe", "无摘要"
    return "include", "符合首轮筛选"


def write_csv(path: Path, rows, fieldnames):
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


def main():
    outdir = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(".")
    outdir.mkdir(parents=True, exist_ok=True)

    pmids, count, search_url = esearch(CORE_QUERY)
    records = extract_records(pmids)

    # PubMed single-query export should already be unique by PMID; keep DOI/title dedup as safeguard.
    seen = set()
    deduped = []
    for row in records:
        key = row["pmid"] or row["doi"] or row["title"].lower()
        if key in seen:
            continue
        seen.add(key)
        deduped.append(row)

    screened = []
    for idx, row in enumerate(deduped, start=1):
        decision, reason = first_pass_decision(row)
        screened.append(
            {
                "record_id": f"PUBMED-{idx:04d}",
                "source": "PubMed",
                "pmid": row["pmid"],
                "doi": row["doi"],
                "year": row["year"],
                "title": row["title"],
                "journal": row["journal"],
                "authors": row["authors"],
                "publication_types": row["publication_types"],
                "title_abstract_decision": decision,
                "title_abstract_reason": reason,
                "abstract": row["abstract"],
                "pubmed_url": row["pubmed_url"],
            }
        )

    search_log = [
        {
            "source": "PubMed",
            "search_date": time.strftime("%Y-%m-%d"),
            "query": CORE_QUERY,
            "raw_count": count,
            "deduplicated_count": len(deduped),
            "search_url": search_url,
        }
    ]

    write_csv(
        outdir / "pubmed_search_results_dedup_screened.csv",
        screened,
        [
            "record_id",
            "source",
            "pmid",
            "doi",
            "year",
            "title",
            "journal",
            "authors",
            "publication_types",
            "title_abstract_decision",
            "title_abstract_reason",
            "abstract",
            "pubmed_url",
        ],
    )
    write_csv(
        outdir / "pubmed_search_log.csv",
        search_log,
        ["source", "search_date", "query", "raw_count", "deduplicated_count", "search_url"],
    )

    summary = {"raw_count": count, "deduplicated_count": len(deduped)}
    for label in ["include", "maybe", "exclude"]:
        summary[label] = sum(1 for row in screened if row["title_abstract_decision"] == label)
    (outdir / "pubmed_search_summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
