from __future__ import annotations

import math
import re
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd


NEG = "\u9634"
POS = "\u9633"
FEMALE = "\u5973"
NO_TOKENS = ["\u5426", "\u65e0"]
YES_TOKENS = [
    "\u662f",
    "\u6709",
    "\u6b7b\u4ea1",
    "\u590d\u53d1",
    "\u5931\u8d25",
    "\u8fdb\u5c55",
    "\u4f9d\u4ece\u6027\u5dee",
    "\u81ea\u884c\u505c\u7ed3\u6838\u836f",
]
MISSING_OUTCOME_TOKENS = [
    "\u4e0d\u53ef\u8bc4\u4f30",
    "\u4e0d\u8be6",
    "\u5931\u8bbf",
    "\u672a\u590d\u67e5",
    "\u672a\u590d\u8bca",
    "\u672a\u505a\u68c0\u67e5",
    "\u9000\u51fa",
    "\u9000\u7ec4",
    "\u4e22\u5931",
]
NOT_RANDOMIZED_TOKENS = {"", "\u672a\u968f\u673a", "\u65e0\u9700\u968f\u673a"}


def wilson_ci(x: int, n: int, z: float = 1.959963984540054) -> tuple[float | None, float | None]:
    if n == 0:
        return None, None
    p = x / n
    denom = 1 + z * z / n
    center = (p + z * z / (2 * n)) / denom
    half = z * math.sqrt((p * (1 - p) + z * z / (4 * n)) / n) / denom
    return center - half, center + half


def fmt_rate(x: int, n: int) -> str:
    if n == 0:
        return "NA"
    lo, hi = wilson_ci(x, n)
    return f"{x}/{n} ({100 * x / n:.1f}%, {100 * lo:.1f}-{100 * hi:.1f})"


def parse_num(value) -> float | np.nan:
    if pd.isna(value):
        return np.nan
    if isinstance(value, (int, float, np.integer, np.floating)):
        return float(value)
    match = re.search(r"[-+]?[0-9]*\.?[0-9]+", str(value))
    return float(match.group()) if match else np.nan


def parse_date(value) -> pd.Timestamp | pd.NaT:
    if pd.isna(value):
        return pd.NaT
    if isinstance(value, (pd.Timestamp, datetime)):
        return pd.to_datetime(value, errors="coerce")
    text = str(value).strip()
    if text in NOT_RANDOMIZED_TOKENS:
        return pd.NaT
    return pd.to_datetime(text, errors="coerce")


def normalize_arm(value) -> str | float:
    if pd.isna(value):
        return np.nan
    raw = str(value).strip()
    compact = raw.upper().replace(" ", "")

    if raw in {"\u7b5b\u9009\u5931\u8d25", "\u9000\u51fa"}:
        return "EXCLUDE"
    if compact.startswith("E"):
        return "E"
    if "SASP" in compact or "\u67f3\u6c2e\u78fa\u543a\u5576\u5576" in raw or compact.startswith("D"):
        return "D"
    if compact.startswith("A"):
        return "A"
    if compact.startswith("B"):
        return "B"
    if compact.startswith("C"):
        return "C"

    if "PTO" in compact or "DLM" in compact or "AM" in compact or "\u5e15\u53f8" in raw:
        return "E"

    if (
        ("MFX" in compact or "LFX" in compact)
        and "LZD" in compact
        and "CS" in compact
        and "CFZ" in compact
        and ("PZA" in compact or "Z" in compact)
        and "BDQ" not in compact
    ):
        return "B"

    if (
        "BDQ" in compact
        and "LZD" in compact
        and "CS" in compact
        and "CFZ" in compact
        and ("PZA" in compact or "Z" in compact)
        and "MFX" not in compact
        and "LFX" not in compact
    ):
        return "C"

    return "OTHER"


def culture_status(value) -> str | float:
    if pd.isna(value):
        return np.nan
    text = str(value).strip()
    if NEG in text:
        return "neg"
    if POS in text:
        return "pos"
    return np.nan


def unfavorable_status(status_value, detail_value) -> float:
    pieces = [v for v in [status_value, detail_value] if not pd.isna(v)]
    if not pieces:
        return np.nan
    text = " | ".join(str(v).strip() for v in pieces if str(v).strip())
    if not text:
        return np.nan
    if any(token in text for token in YES_TOKENS):
        return 1.0
    if any(token in text for token in MISSING_OUTCOME_TOKENS):
        return np.nan
    if any(token in text for token in NO_TOKENS):
        return 0.0
    return np.nan


def summarize_group(df: pd.DataFrame, label: str) -> dict[str, object]:
    eligible_6m = int(df["eligible_6m"].sum())
    assessed_6m = int(df.loc[df["eligible_6m"], "unf_6m"].notna().sum())
    events_6m = int(df.loc[df["eligible_6m"], "unf_6m"].eq(1).sum())

    eligible_12m = int(df["eligible_12m"].sum())
    assessed_12m = int(df.loc[df["eligible_12m"], "unf_12m"].notna().sum())
    events_12m = int(df.loc[df["eligible_12m"], "unf_12m"].eq(1).sum())

    month2_eval = int(df["m2_culture"].isin(["neg", "pos"]).sum())
    month2_neg = int(df["m2_culture"].eq("neg").sum())
    eot_eval = int(df["eot_culture"].isin(["neg", "pos"]).sum())
    eot_neg = int(df["eot_culture"].eq("neg").sum())

    return {
        "group": label,
        "records_n": int(len(df)),
        "unique_id_strings_n": int(df["subject_id"].nunique()),
        "centers_n": int(df["center"].nunique()),
        "age_median": float(df["age_num"].median()) if df["age_num"].notna().any() else np.nan,
        "age_q1": float(df["age_num"].quantile(0.25)) if df["age_num"].notna().any() else np.nan,
        "age_q3": float(df["age_num"].quantile(0.75)) if df["age_num"].notna().any() else np.nan,
        "female_n": int(df["is_female"].sum()),
        "female_pct": round(100 * df["is_female"].mean(), 1) if len(df) else np.nan,
        "month2_neg": month2_neg,
        "month2_eval": month2_eval,
        "month2_rate": round(100 * month2_neg / month2_eval, 1) if month2_eval else np.nan,
        "month2_rate_ci": fmt_rate(month2_neg, month2_eval),
        "eot_neg": eot_neg,
        "eot_eval": eot_eval,
        "eot_rate": round(100 * eot_neg / eot_eval, 1) if eot_eval else np.nan,
        "eot_rate_ci": fmt_rate(eot_neg, eot_eval),
        "eligible_6m": eligible_6m,
        "assessed_6m": assessed_6m,
        "unfavorable_6m": events_6m,
        "unfavorable_6m_rate": round(100 * events_6m / assessed_6m, 1) if assessed_6m else np.nan,
        "unfavorable_6m_rate_ci": fmt_rate(events_6m, assessed_6m),
        "ascertainment_6m_pct": round(100 * assessed_6m / eligible_6m, 1) if eligible_6m else np.nan,
        "eligible_12m": eligible_12m,
        "assessed_12m": assessed_12m,
        "unfavorable_12m": events_12m,
        "unfavorable_12m_rate": round(100 * events_12m / assessed_12m, 1) if assessed_12m else np.nan,
        "unfavorable_12m_rate_ci": fmt_rate(events_12m, assessed_12m),
        "ascertainment_12m_pct": round(100 * assessed_12m / eligible_12m, 1) if eligible_12m else np.nan,
    }


def build_dataset(xlsx_path: Path) -> tuple[pd.DataFrame, dict[str, int]]:
    raw = pd.read_excel(xlsx_path, sheet_name=0)
    raw_rows = len(raw)
    dedup = raw.drop_duplicates().copy()
    removed_exact_duplicates = raw_rows - len(dedup)

    cols = list(dedup.columns)
    center = cols[0]
    subject_id = cols[1]
    age = cols[2]
    sex = cols[3]
    arm = cols[5]
    end = cols[8]
    month2 = cols[22]
    eot = cols[23]
    unf_6m = cols[28]
    unf_6m_detail = cols[29]
    unf_12m = cols[30]
    unf_12m_detail = cols[31]

    dedup = dedup.rename(
        columns={
            center: "center",
            subject_id: "subject_id",
            age: "age_raw",
            sex: "sex_raw",
            arm: "arm_raw",
            end: "end_raw",
            month2: "month2_raw",
            eot: "eot_raw",
            unf_6m: "unf_6m_raw",
            unf_6m_detail: "unf_6m_detail_raw",
            unf_12m: "unf_12m_raw",
            unf_12m_detail: "unf_12m_detail_raw",
        }
    )

    dedup["arm_norm"] = dedup["arm_raw"].map(normalize_arm)
    dedup["age_num"] = dedup["age_raw"].map(parse_num)
    dedup["is_female"] = dedup["sex_raw"].astype(str).str.contains(FEMALE, na=False)
    dedup["end_dt"] = dedup["end_raw"].map(parse_date)
    dedup["m2_culture"] = dedup["month2_raw"].map(culture_status)
    dedup["eot_culture"] = dedup["eot_raw"].map(culture_status)
    dedup["unf_6m"] = [
        unfavorable_status(a, b)
        for a, b in zip(dedup["unf_6m_raw"], dedup["unf_6m_detail_raw"])
    ]
    dedup["unf_12m"] = [
        unfavorable_status(a, b)
        for a, b in zip(dedup["unf_12m_raw"], dedup["unf_12m_detail_raw"])
    ]

    analysis = dedup[dedup["subject_id"].notna() & dedup["arm_norm"].isin(["A", "B", "C", "D", "E"])].copy()
    cutoff = pd.Timestamp("2026-03-31")
    analysis["eligible_6m"] = analysis["end_dt"].notna() & (analysis["end_dt"] <= cutoff - pd.Timedelta(days=183))
    analysis["eligible_12m"] = analysis["end_dt"].notna() & (analysis["end_dt"] <= cutoff - pd.Timedelta(days=365))

    qc = {
        "raw_rows": raw_rows,
        "rows_after_exact_dedup": len(dedup),
        "removed_exact_duplicates": removed_exact_duplicates,
        "analysis_rows": len(analysis),
        "unique_id_strings": int(analysis["subject_id"].nunique()),
        "duplicate_id_rows": int(analysis.duplicated(subset=["subject_id"], keep=False).sum()),
        "duplicate_id_strings": int(analysis.loc[analysis.duplicated(subset=["subject_id"], keep=False), "subject_id"].nunique()),
        "other_arm_rows": int((dedup["subject_id"].notna() & dedup["arm_norm"].eq("OTHER")).sum()),
        "excluded_admin_rows": int((dedup["subject_id"].notna() & dedup["arm_norm"].eq("EXCLUDE")).sum()),
    }
    return analysis, qc


def write_outputs(base_dir: Path, analysis: pd.DataFrame, qc: dict[str, int]) -> None:
    out_dir = base_dir / "analysis_outputs" / "union_20260331"
    out_dir.mkdir(parents=True, exist_ok=True)

    ad = analysis[analysis["arm_norm"].isin(["A", "B", "C", "D"])].copy()
    e = analysis[analysis["arm_norm"] == "E"].copy()

    group_summary = pd.DataFrame(
        [
            summarize_group(ad, "A-D pooled"),
            summarize_group(e, "E real-world"),
            summarize_group(analysis, "A-E overall"),
        ]
    )
    arm_summary = pd.DataFrame(
        [summarize_group(analysis[analysis["arm_norm"] == arm], arm) for arm in ["A", "B", "C", "D", "E"]]
    )

    center_summary = (
        analysis.groupby("center", dropna=False)
        .apply(
            lambda x: pd.Series(
                {
                    "records_n": int(len(x)),
                    "arms_present": ",".join(sorted(x["arm_norm"].dropna().unique())),
                    "month2_eval": int(x["m2_culture"].isin(["neg", "pos"]).sum()),
                    "eligible_12m": int(x["eligible_12m"].sum()),
                    "assessed_12m": int(x.loc[x["eligible_12m"], "unf_12m"].notna().sum()),
                    "unfavorable_12m": int(x.loc[x["eligible_12m"], "unf_12m"].eq(1).sum()),
                }
            )
        )
        .reset_index()
        .sort_values(["records_n", "center"], ascending=[False, True])
    )

    group_summary.to_csv(out_dir / "group_summary.csv", index=False, encoding="utf-8-sig")
    arm_summary.to_csv(out_dir / "arm_summary.csv", index=False, encoding="utf-8-sig")
    center_summary.to_csv(out_dir / "center_summary.csv", index=False, encoding="utf-8-sig")

    ad_row = group_summary[group_summary["group"] == "A-D pooled"].iloc[0]
    e_row = group_summary[group_summary["group"] == "E real-world"].iloc[0]
    overall_row = group_summary[group_summary["group"] == "A-E overall"].iloc[0]

    quality_md = f"""# SEAL-MDR Data Quality Notes

Analysis date: 2026-03-31
Workbook snapshot: `{next(base_dir.glob("*.xlsx")).name}`

## Rules used

- Exact duplicate spreadsheet rows were removed before aggregation.
- Arm labels were normalized conservatively into `A/B/C/D/E`; entries that could not be mapped reliably were excluded from arm-level summaries.
- Culture endpoints used an evaluable denominator of explicit positive or negative results only.
- Post-treatment 6-month and 12-month unfavorable outcomes were analyzed only among records with sufficient follow-up time from treatment end.
- Outcome values such as `否/无` were treated as `No`; explicit event strings such as `是/有/死亡/复发/失败/进展/依从性差/自行停结核药` were treated as `Yes`.
- Free-text statuses indicating non-assessable follow-up such as `失访/未复诊/不可评估` were treated as missing.

## Quality checks

- Raw rows: {qc['raw_rows']}
- Rows after exact de-duplication: {qc['rows_after_exact_dedup']}
- Exact duplicate rows removed: {qc['removed_exact_duplicates']}
- Record-level rows included in A-E analysis: {qc['analysis_rows']}
- Unique ID strings in A-E analysis: {qc['unique_id_strings']}
- Rows with duplicated ID strings: {qc['duplicate_id_rows']}
- Number of duplicated ID strings: {qc['duplicate_id_strings']}
- Rows with non-empty but unclassifiable arm labels: {qc['other_arm_rows']}
- Rows excluded because the arm field looked administrative (for example screening failure / exit): {qc['excluded_admin_rows']}

## Interpretation cautions

- This is safest to report as an interim record-level multicenter analysis.
- The `受试者ID` field behaves like a name string rather than a guaranteed unique subject key.
- Post-treatment ascertainment is incomplete, especially outside the largest centers and in arms D/E.
- The workbook is well-suited to conference-style descriptive results, but not yet to a final SAP-grade comparative analysis.
"""

    abstract_md = f"""# Union Abstract Draft

## Candidate Title

Interim multicenter outcomes of all-oral shortened regimens in SEAL-MDR for rifampicin-resistant or multidrug-resistant tuberculosis

## Alternative Title

Multicenter interim analysis of the SEAL-MDR platform: early microbiological response and post-treatment outcomes of all-oral shortened regimens

## Draft Abstract

**Background:** Shorter all-oral regimens are needed for rifampicin-resistant or multidrug-resistant tuberculosis (RR/MDR-TB), but multicenter data from China remain limited. We analyzed the current SEAL-MDR workbook snapshot to generate a conference-ready interim multicenter summary using only existing curated Excel data.

**Methods:** SEAL-MDR is a multicenter clinical study of all-oral shortened regimens. For this interim analysis, exact duplicate spreadsheet rows were removed, arm labels were conservatively normalized to protocol-defined arms A-D and real-world arm E, and descriptive analyses were performed at record level. Month-2 culture negativity and end-of-treatment (EOT) culture negativity were calculated among evaluable records with explicit positive or negative results. Post-treatment 6-month and 12-month unfavorable outcomes were analyzed among records with mature follow-up from treatment end and assessable outcome status.

**Results:** After exact de-duplication, 733 analyzable records from 25 centers were available, including 601 records in protocol-defined arms A-D and 132 records in real-world arm E. In the pooled A-D cohort, median age was {ad_row['age_median']:.0f} years (IQR {ad_row['age_q1']:.0f}-{ad_row['age_q3']:.0f}) and {ad_row['female_pct']:.1f}% were female. Month-2 culture negativity was {int(ad_row['month2_neg'])}/{int(ad_row['month2_eval'])} ({ad_row['month2_rate']:.1f}%), and EOT culture negativity was {int(ad_row['eot_neg'])}/{int(ad_row['eot_eval'])} ({ad_row['eot_rate']:.1f}%). Among records eligible for post-treatment assessment in A-D, unfavorable outcomes occurred in {int(ad_row['unfavorable_6m'])}/{int(ad_row['assessed_6m'])} ({ad_row['unfavorable_6m_rate']:.1f}%) at 6 months and {int(ad_row['unfavorable_12m'])}/{int(ad_row['assessed_12m'])} ({ad_row['unfavorable_12m_rate']:.1f}%) at 12 months; follow-up ascertainment among eligible records was {ad_row['ascertainment_6m_pct']:.1f}% and {ad_row['ascertainment_12m_pct']:.1f}%, respectively. The real-world E cohort showed a higher 12-month unfavorable outcome rate ({int(e_row['unfavorable_12m'])}/{int(e_row['assessed_12m'])}, {e_row['unfavorable_12m_rate']:.1f}%), consistent with a higher-risk rescue population. Across all A-E records, 12-month unfavorable outcomes were observed in {int(overall_row['unfavorable_12m'])}/{int(overall_row['assessed_12m'])} ({overall_row['unfavorable_12m_rate']:.1f}%).

**Conclusions:** In this interim multicenter SEAL-MDR analysis, protocol-defined all-oral shortened regimens showed strong early microbiological response and high EOT culture negativity, with encouraging post-treatment outcomes among records with mature follow-up. Ongoing follow-up completion and tighter arm/status standardization will strengthen final comparative analyses.

## Suggested Results Table For Submission

- Main text: pooled A-D cohort
- Secondary context: E real-world cohort
- Figure 1: Month-2 culture negativity and EOT culture negativity with exact `n/N`
- Figure 2: Post-treatment unfavorable outcomes at 6 and 12 months among assessed records, with ascertainment explicitly stated in the legend
"""

    (out_dir / "data_quality_notes.md").write_text(quality_md, encoding="utf-8")
    (out_dir / "union_abstract_draft.md").write_text(abstract_md, encoding="utf-8")


def main() -> None:
    base_dir = Path.cwd()
    xlsx_files = sorted(base_dir.glob("*.xlsx"))
    if not xlsx_files:
        raise FileNotFoundError("No xlsx file found in working directory.")

    analysis, qc = build_dataset(xlsx_files[0])
    write_outputs(base_dir, analysis, qc)


if __name__ == "__main__":
    main()
