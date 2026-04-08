from __future__ import annotations

import argparse
import json
import shutil
from dataclasses import dataclass
from pathlib import Path

import pandas as pd


AS_OF_DATE_DEFAULT = "2026-03-31"


@dataclass(frozen=True)
class ColumnSet:
    subject_id: str = "受试者ID"
    age: str = "年龄（岁）"
    sex: str = "性别"
    arm: str = "研究组/方案（Arm）"
    start_date: str = "治疗开始日期"
    end_date: str = "治疗结束日期"
    duration_days: str = "疗程天数（为）"
    analysis_set: str = "分析集（ITT/mITT/PP）"
    resistance_type: str = "耐药类型"
    site: str = "结核部位"
    micro_confirmed: str = "微生物学确诊"
    baseline_culture: str = "基线培养"
    week8_culture: str = "Week 8培养"
    month2_culture: str = "治疗2个月培养"
    eot_culture: str = "停药时培养"
    followup_6m_culture: str = "停药后6月培养"
    followup_12m_culture: str = "停药后12月培养"
    unfavorable_eot: str = "停药时，是否不良结局（Unfavorable）"
    unfavorable_eot_component: str = "停药时，不良结局构成"
    unfavorable_6m: str = "停药后半年，是否不良结局（Unfavorable）"
    unfavorable_6m_component: str = "停药后半年，不良结局构成"
    unfavorable_12m: str = "停药后1年，是否不良结局（Unfavorable）"
    unfavorable_12m_component: str = "停药后1年，不良结局构成"
    final_outcome: str = "最终治疗结局"
    relapse_12m: str = "12月复发"
    grade3_ae: str = "≥3级不良事件（任意）"
    sae: str = "严重不良事件（SAE，任意）"
    ae_stop: str = "因AE停药/永久停药"
    death_any: str = "死亡（任意）"
    note: str = "备注.1"


COLS = ColumnSet()
PROTOCOL_ARMS = {"BDL", "BDLM"}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Create a protocol-aligned first-pass analysis from the BDLM Union workbook."
    )
    parser.add_argument("--input", type=Path, default=None, help="Path to the source workbook.")
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("reports") / "bdlm_union_first_pass",
        help="Directory for cleaned data and report outputs.",
    )
    parser.add_argument(
        "--as-of-date",
        default=AS_OF_DATE_DEFAULT,
        help="Analysis cutoff date in YYYY-MM-DD format.",
    )
    return parser.parse_args()


def find_default_workbook(root: Path) -> Path:
    candidates = sorted(root.rglob("*.xlsx"))
    ranked: list[Path] = []
    for path in candidates:
        name = path.name.lower()
        full = str(path).lower()
        if path.name.startswith("~$"):
            continue
        score = 0
        if "union" in name:
            score += 3
        if "bdlm" in name or "bdlm" in full:
            score += 2
        if "摘要" in path.name:
            score += 1
        if score > 0:
            ranked.append((score, path))
    if not ranked:
        raise FileNotFoundError("No candidate BDLM workbook was found under the working directory.")
    ranked.sort(key=lambda item: (-item[0], len(str(item[1])), str(item[1])))
    return ranked[0][1]


def normalize_text(value: object) -> str | None:
    if pd.isna(value):
        return None
    text = str(value).strip()
    return text or None


def normalize_yes_no(value: object) -> str | None:
    text = normalize_text(value)
    if text is None:
        return None
    mapping = {
        "是": "YES",
        "否": "NO",
        "阳": "YES",
        "阴": "NO",
        "阳性": "YES",
        "阴性": "NO",
    }
    return mapping.get(text, text)


def normalize_culture(value: object) -> str | None:
    text = normalize_text(value)
    if text is None:
        return None
    if "阳" in text:
        return "POS"
    if "阴" in text:
        return "NEG"
    if text in {"污染"}:
        return "CONTAM"
    if text in {"未查", "未做", "未测"}:
        return "ND"
    return text


def normalize_component(value: object) -> str | None:
    text = normalize_text(value)
    if text is None:
        return None
    if "失访" in text or "脱落" in text:
        return "失访/脱落"
    if "退出" in text:
        return "退出"
    if "死亡" in text:
        return "死亡"
    if "失败" in text:
        return "治疗失败"
    if "复发" in text or "复阳" in text:
        return "复发"
    if text in {"无", "否"}:
        return "无"
    if text.startswith("其他"):
        return "其他"
    return text


def classify_strict_outcome(
    *,
    unfavorable_std: object,
    component_std: object,
    culture_std: object,
    note: object,
    mature: bool,
) -> str:
    component = normalize_text(component_std)
    culture = normalize_text(culture_std)
    note_text = normalize_text(note) or ""
    unfavorable = normalize_text(unfavorable_std)

    if component in {"治疗失败", "复发", "死亡"}:
        return "TRUE_UNFAVORABLE"
    if culture == "POS":
        return "TRUE_UNFAVORABLE"
    if "复阳" in note_text or "复发" in note_text or "再感染" in note_text:
        return "TRUE_UNFAVORABLE"
    if component == "失访/脱落":
        return "LOST_TO_FOLLOWUP"
    if component == "退出":
        return "WITHDRAWN"
    if component == "其他":
        return "OTHER_UNFAVORABLE"
    if unfavorable == "NO":
        return "FAVORABLE"
    if mature:
        return "UNKNOWN_AFTER_WINDOW"
    return "PENDING"


def strict_outcome_summary(df: pd.DataFrame, arm_label: str | None, status_col: str) -> pd.DataFrame:
    subset = df if arm_label is None else df[df["arm_std"] == arm_label]
    counts = subset[status_col].value_counts(dropna=False)
    total = int(len(subset))
    rows: list[dict[str, object]] = []
    for status in [
        "TRUE_UNFAVORABLE",
        "LOST_TO_FOLLOWUP",
        "WITHDRAWN",
        "OTHER_UNFAVORABLE",
        "FAVORABLE",
        "PENDING",
        "UNKNOWN_AFTER_WINDOW",
    ]:
        count = int(counts.get(status, 0))
        rows.append(
            {
                "arm": arm_label or "Overall",
                "status_col": status_col,
                "status": status,
                "count": count,
                "total": total,
                "percent": (count / total * 100.0) if total else None,
            }
        )
    return pd.DataFrame(rows)


def culture_rate(df: pd.DataFrame, arm_label: str | None, source_col: str) -> dict[str, object]:
    subset = df if arm_label is None else df[df["arm_std"] == arm_label]
    evaluable = subset[source_col].isin(["NEG", "POS"])
    denom = int(evaluable.sum())
    numer = int((subset.loc[evaluable, source_col] == "NEG").sum())
    pct = (numer / denom * 100.0) if denom else None
    return {
        "arm": arm_label or "Overall",
        "metric": source_col,
        "numerator": numer,
        "denominator": denom,
        "percent": pct,
    }


def outcome_rate(df: pd.DataFrame, arm_label: str | None, source_col: str) -> dict[str, object]:
    subset = df if arm_label is None else df[df["arm_std"] == arm_label]
    evaluable = subset[source_col].isin(["YES", "NO"])
    denom = int(evaluable.sum())
    numer = int((subset.loc[evaluable, source_col] == "YES").sum())
    pct = (numer / denom * 100.0) if denom else None
    return {
        "arm": arm_label or "Overall",
        "metric": source_col,
        "numerator": numer,
        "denominator": denom,
        "percent": pct,
    }


def format_rate(row: dict[str, object]) -> str:
    pct = row["percent"]
    if pct is None:
        return f"{row['numerator']}/{row['denominator']}"
    return f"{row['numerator']}/{row['denominator']} ({pct:.1f}%)"


def format_count_percent(count: int, total: int) -> str:
    if total == 0:
        return "0/0"
    return f"{count}/{total} ({count / total * 100.0:.1f}%)"


def build_strict_summary_table(protocol_df: pd.DataFrame) -> pd.DataFrame:
    status_order = [
        "TRUE_UNFAVORABLE",
        "LOST_TO_FOLLOWUP",
        "WITHDRAWN",
        "OTHER_UNFAVORABLE",
        "FAVORABLE",
        "PENDING",
        "UNKNOWN_AFTER_WINDOW",
    ]
    status_labels = {
        "TRUE_UNFAVORABLE": "真正失败/复发",
        "LOST_TO_FOLLOWUP": "失访",
        "WITHDRAWN": "退出",
        "OTHER_UNFAVORABLE": "其他非方案性不良结局",
        "FAVORABLE": "随访完成且无事件",
        "PENDING": "尚未到对应随访时间窗",
        "UNKNOWN_AFTER_WINDOW": "已到时间窗但记录未明",
    }
    endpoint_map = {
        "strict_6m_status": "主要结局：停药后6个月",
        "strict_12m_status": "次要结局：停药后12个月",
    }
    rows: list[dict[str, object]] = []
    for endpoint_col, endpoint_label in endpoint_map.items():
        for arm in ["Overall", "BDL", "BDLM"]:
            subset = protocol_df if arm == "Overall" else protocol_df[protocol_df["arm_std"] == arm]
            total = int(len(subset))
            row: dict[str, object] = {"endpoint": endpoint_label, "arm": arm, "N": total}
            counts = subset[endpoint_col].value_counts(dropna=False)
            for status in status_order:
                count = int(counts.get(status, 0))
                row[status_labels[status]] = format_count_percent(count, total) if count else f"0/{total} (0.0%)"
            rows.append(row)
    return pd.DataFrame(rows)


def build_case_review_table(protocol_df: pd.DataFrame) -> pd.DataFrame:
    note_col = COLS.note

    def derive_basis(row: pd.Series, endpoint: str) -> str:
        if endpoint == "6M":
            component = row["unfavorable_6m_component_std"]
            culture = row["followup_6m_culture_std"]
        else:
            component = row["unfavorable_12m_component_std"]
            culture = row["followup_12m_culture_std"]
        bits: list[str] = []
        if pd.notna(component):
            bits.append(f"结局构成={component}")
        if pd.notna(culture):
            bits.append(f"培养={culture}")
        note = normalize_text(row.get(note_col))
        if note:
            bits.append(f"备注={note}")
        return "；".join(bits)

    rows: list[dict[str, object]] = []
    for endpoint, status_col in [("6M", "strict_6m_status"), ("12M", "strict_12m_status")]:
        sub = protocol_df[
            protocol_df[status_col].isin(
                ["TRUE_UNFAVORABLE", "LOST_TO_FOLLOWUP", "WITHDRAWN", "OTHER_UNFAVORABLE"]
            )
        ].copy()
        for _, row in sub.iterrows():
            rows.append(
                {
                    "endpoint": "主要结局：停药后6个月" if endpoint == "6M" else "次要结局：停药后12个月",
                    "strict_status": row[status_col],
                    "subject_id": row["subject_id_std"],
                    "arm": row["arm_std"],
                    "start_date": row["start_date_std"],
                    "end_date": row["end_date_std"],
                    "basis": derive_basis(row, endpoint),
                    "ongoing_contact_note": "仍在尝试联络" if endpoint == "12M" and row[status_col] == "LOST_TO_FOLLOWUP" else "",
                }
            )
    review_df = pd.DataFrame(rows)
    status_sort = {
        "TRUE_UNFAVORABLE": 1,
        "LOST_TO_FOLLOWUP": 2,
        "WITHDRAWN": 3,
        "OTHER_UNFAVORABLE": 4,
    }
    endpoint_sort = {"主要结局：停药后6个月": 1, "次要结局：停药后12个月": 2}
    if not review_df.empty:
        review_df["_endpoint_sort"] = review_df["endpoint"].map(endpoint_sort)
        review_df["_status_sort"] = review_df["strict_status"].map(status_sort)
        review_df = review_df.sort_values(
            by=["_endpoint_sort", "_status_sort", "arm", "start_date", "subject_id"],
            na_position="last",
        ).drop(columns=["_endpoint_sort", "_status_sort"])
    return review_df


def build_report(
    source_path: Path,
    output_dir: Path,
    as_of_date: pd.Timestamp,
    raw_df: pd.DataFrame,
    dedup_df: pd.DataFrame,
    protocol_df: pd.DataFrame,
    duplicate_rows: pd.DataFrame,
) -> str:
    arm_counts = protocol_df["arm_std"].value_counts().sort_index()
    mature_12m = protocol_df["followup_12m_mature"].sum()
    available_12m = protocol_df["unfavorable_12m_std"].isin(["YES", "NO"]).sum()
    age_out = protocol_df[~protocol_df["age_in_protocol_range"]]

    culture_rows: list[dict[str, object]] = []
    for arm in [None, "BDL", "BDLM"]:
        culture_rows.append(culture_rate(protocol_df, arm, "month2_culture_std"))
        culture_rows.append(culture_rate(protocol_df, arm, "eot_culture_std"))

    outcome_rows: list[dict[str, object]] = []
    for arm in [None, "BDL", "BDLM"]:
        outcome_rows.append(outcome_rate(protocol_df, arm, "unfavorable_eot_std"))
        outcome_rows.append(outcome_rate(protocol_df, arm, "unfavorable_6m_std"))
        outcome_rows.append(outcome_rate(protocol_df, arm, "unfavorable_12m_std"))

    strict_6m_rows = pd.concat(
        [strict_outcome_summary(protocol_df, arm, "strict_6m_status") for arm in [None, "BDL", "BDLM"]],
        ignore_index=True,
    )
    strict_12m_rows = pd.concat(
        [strict_outcome_summary(protocol_df, arm, "strict_12m_status") for arm in [None, "BDL", "BDLM"]],
        ignore_index=True,
    )

    comp_12m = (
        protocol_df.loc[protocol_df["unfavorable_12m_std"] == "YES", "unfavorable_12m_component_std"]
        .fillna("未分类")
        .value_counts()
    )

    lines: list[str] = []
    lines.append("# BDLM Union最小数据集首轮分析")
    lines.append("")
    lines.append(f"- 数据源：`{source_path}`")
    lines.append(f"- 分析日期截点：`{as_of_date.date().isoformat()}`")
    lines.append(f"- 输出目录：`{output_dir}`")
    lines.append("")
    lines.append("## 队列概览")
    lines.append(f"- 原始逐例数据：{len(raw_df)} 例")
    lines.append(f"- 去重后：{len(dedup_df)} 例")
    lines.append(
        f"- 严格保留方案内 `BDL/BDLM`：{len(protocol_df)} 例"
        f"（BDL {int(arm_counts.get('BDL', 0))}，BDLM {int(arm_counts.get('BDLM', 0))}）"
    )
    lines.append(f"- 非方案内其他方案：{int((dedup_df['protocol_arm'] == False).sum())} 例")
    lines.append(f"- 重复受试者ID：{duplicate_rows[COLS.subject_id].nunique()} 个")
    lines.append("")
    lines.append("## 主要数据质量提醒")
    lines.append(
        f"- 年龄超出方案 `14-75岁`：{len(age_out)} 例"
        + (f"（{', '.join(age_out[COLS.subject_id].astype(str).tolist())}）" if len(age_out) else "")
    )
    lines.append(
        f"- 12月结局随访成熟（按停药后满365天判断）：{int(mature_12m)} 例；"
        f"已录入标准 `是/否` 的12月不良结局：{int(available_12m)} 例"
    )
    lines.append(
        "- 这份工作簿是摘要最小数据集，`最终治疗结局`、`SAE`、`≥3级AE`、`因AE停药`、`死亡`"
        " 等字段目前基本为空，暂不适合做完整安全性或回归分析。"
    )
    lines.append(
        "- 培养结果原始录入混有 `阴/阴性`、`阳/阳性`、`未查`，本次已统一为标准类别后再统计。"
    )
    lines.append("")
    lines.append("## 首轮有效性结果")
    lines.append("### 培养阴转率")
    for row in culture_rows:
        label = "2个月培养阴转率" if row["metric"] == "month2_culture_std" else "停药时培养阴性率"
        lines.append(f"- {row['arm']} {label}：{format_rate(row)}")
    lines.append("")
    lines.append("### 不良结局率")
    for row in outcome_rows:
        metric_map = {
            "unfavorable_eot_std": "停药时不良结局率",
            "unfavorable_6m_std": "停药后6月不良结局率",
            "unfavorable_12m_std": "停药后1年不良结局率",
        }
        lines.append(f"- {row['arm']} {metric_map[row['metric']]}：{format_rate(row)}")
    lines.append("")
    lines.append("### 停药后1年不良结局构成")
    if len(comp_12m):
        for name, count in comp_12m.items():
            lines.append(f"- {name}：{int(count)} 例")
    else:
        lines.append("- 暂无可用记录")
    lines.append("")
    lines.append("## 严格终点重建")
    lines.append(
        "- 主要结局改为 `停药后6个月`，次要结局改为 `停药后12个月`。"
        " 失访和退出不再并入“真正失败/复发”终点，而是作为单独状态报告。"
    )
    lines.append(
        "- 严格终点规则：`治疗失败/复发` 组件、随访培养阳性、或备注中明确“复阳/复发/再感染”"
        " 记为 `真正失败/复发`；`失访/脱落`、`退出`、`其他` 单列；`NO` 记为 `随访完成且无事件`。"
    )
    lines.append("")
    lines.append("### 主要结局：停药后6个月")
    status_labels = {
        "TRUE_UNFAVORABLE": "真正失败/复发",
        "LOST_TO_FOLLOWUP": "失访",
        "WITHDRAWN": "退出",
        "OTHER_UNFAVORABLE": "其他非方案性不良结局",
        "FAVORABLE": "随访完成且无事件",
        "PENDING": "尚未到对应随访时间窗",
        "UNKNOWN_AFTER_WINDOW": "已到时间窗但记录未明",
    }
    for arm in ["Overall", "BDL", "BDLM"]:
        lines.append(f"- {arm}：")
        subset = strict_6m_rows[strict_6m_rows["arm"] == arm]
        parts = []
        for _, row in subset.iterrows():
            if row["count"]:
                parts.append(
                    f"{status_labels[row['status']]} {int(row['count'])}/{int(row['total'])}"
                    f" ({row['percent']:.1f}%)"
                )
        lines.append("  " + "；".join(parts))
    lines.append("")
    lines.append("### 次要结局：停药后12个月")
    for arm in ["Overall", "BDL", "BDLM"]:
        lines.append(f"- {arm}：")
        subset = strict_12m_rows[strict_12m_rows["arm"] == arm]
        parts = []
        for _, row in subset.iterrows():
            if row["count"]:
                parts.append(
                    f"{status_labels[row['status']]} {int(row['count'])}/{int(row['total'])}"
                    f" ({row['percent']:.1f}%)"
                )
        lines.append("  " + "；".join(parts))
    lines.append("")
    lines.append(
        "- 注：停药后12个月被标记为“失访”的患者，目前仍在继续尝试联络，因此这些病例应理解为"
        " 结局待进一步澄清，而不是已经证实发生真正失败/复发。"
    )
    lines.append("")
    lines.append("## 解释注意")
    lines.append(
        "- 这里的主要结局按工作簿字段 `停药后1年不良结局` 计算；它与研究方案正文里"
        " “治疗开始后12个月不良结局”并不完全同一口径，正式分析前需要统一定义。"
    )
    lines.append(
        "- `BDL` 和 `BDLM` 不是随机分配，而是按药敏策略决定，因此当前只能做描述性比较，"
        " 不建议直接解释为方案优劣。"
    )
    lines.append("")
    lines.append("## 输出文件")
    lines.append("- `cleaned_all_rows.csv`：标准化后的全部病例，含排除和质控标记")
    lines.append("- `cleaned_protocol_bdl_bdlm.csv`：去重后且仅保留 `BDL/BDLM` 的分析数据")
    lines.append("- `duplicate_resolution.csv`：重复ID及保留规则追踪")
    lines.append("- `arm_summary.csv`：主要指标按总体和分组汇总")
    lines.append("- `strict_outcome_summary.csv`：严格终点的互斥分类汇总")
    lines.append("- `strict_summary_table.csv` / `strict_summary_table.md`：可直接用于汇报的主次结局宽表")
    lines.append("- `strict_case_review.csv` / `strict_case_review.md`：严格终点逐例核查表")
    return "\n".join(lines) + "\n"


def dataframe_to_markdown(df: pd.DataFrame) -> str:
    if df.empty:
        return "_No rows_\n"
    cols = [str(c) for c in df.columns]
    header = "| " + " | ".join(cols) + " |"
    sep = "| " + " | ".join(["---"] * len(cols)) + " |"
    body = []
    for _, row in df.iterrows():
        vals = []
        for col in df.columns:
            value = row[col]
            text = "" if pd.isna(value) else str(value)
            text = text.replace("\n", "<br>")
            vals.append(text)
        body.append("| " + " | ".join(vals) + " |")
    return "\n".join([header, sep] + body) + "\n"


def main() -> None:
    args = parse_args()
    cwd = Path.cwd()
    source_path = args.input or find_default_workbook(cwd)
    output_dir = args.output_dir
    output_dir.mkdir(parents=True, exist_ok=True)

    as_of_date = pd.Timestamp(args.as_of_date)
    copied_path = output_dir / "source_workbook_copy.xlsx"
    shutil.copy2(source_path, copied_path)

    raw_df = pd.read_excel(copied_path, sheet_name=0).copy()
    raw_df["row_number_original"] = raw_df.index + 2
    raw_df["subject_id_std"] = raw_df[COLS.subject_id].map(normalize_text)
    raw_df["arm_std"] = raw_df[COLS.arm].map(normalize_text)
    raw_df["start_date_std"] = pd.to_datetime(raw_df[COLS.start_date], errors="coerce")
    raw_df["end_date_std"] = pd.to_datetime(raw_df[COLS.end_date], errors="coerce")
    raw_df["age_num"] = pd.to_numeric(raw_df[COLS.age], errors="coerce")

    culture_map = {
        COLS.baseline_culture: "baseline_culture_std",
        COLS.week8_culture: "week8_culture_std",
        COLS.month2_culture: "month2_culture_std",
        COLS.eot_culture: "eot_culture_std",
        COLS.followup_6m_culture: "followup_6m_culture_std",
        COLS.followup_12m_culture: "followup_12m_culture_std",
    }
    for source, target in culture_map.items():
        raw_df[target] = raw_df[source].map(normalize_culture)

    outcome_map = {
        COLS.unfavorable_eot: "unfavorable_eot_std",
        COLS.unfavorable_6m: "unfavorable_6m_std",
        COLS.unfavorable_12m: "unfavorable_12m_std",
        COLS.relapse_12m: "relapse_12m_std",
        COLS.grade3_ae: "grade3_ae_std",
        COLS.sae: "sae_std",
        COLS.ae_stop: "ae_stop_std",
        COLS.death_any: "death_any_std",
    }
    for source, target in outcome_map.items():
        raw_df[target] = raw_df[source].map(normalize_yes_no)

    component_map = {
        COLS.unfavorable_eot_component: "unfavorable_eot_component_std",
        COLS.unfavorable_6m_component: "unfavorable_6m_component_std",
        COLS.unfavorable_12m_component: "unfavorable_12m_component_std",
    }
    for source, target in component_map.items():
        raw_df[target] = raw_df[source].map(normalize_component)

    raw_df["protocol_arm"] = raw_df["arm_std"].isin(PROTOCOL_ARMS)
    raw_df["age_in_protocol_range"] = raw_df["age_num"].between(14, 75, inclusive="both")
    raw_df["followup_6m_mature"] = raw_df["end_date_std"].notna() & (
        raw_df["end_date_std"] <= as_of_date - pd.Timedelta(days=183)
    )
    raw_df["followup_12m_mature"] = raw_df["end_date_std"].notna() & (
        raw_df["end_date_std"] <= as_of_date - pd.Timedelta(days=365)
    )
    raw_df["available_12m_outcome"] = raw_df["unfavorable_12m_std"].isin(["YES", "NO"])

    completeness_fields = [
        "end_date_std",
        "month2_culture_std",
        "eot_culture_std",
        "followup_6m_culture_std",
        "followup_12m_culture_std",
        "unfavorable_eot_std",
        "unfavorable_6m_std",
        "unfavorable_12m_std",
        "unfavorable_12m_component_std",
    ]
    raw_df["completeness_score"] = raw_df[completeness_fields].notna().sum(axis=1)

    duplicate_mask = raw_df["subject_id_std"].duplicated(keep=False) & raw_df["subject_id_std"].notna()
    duplicate_rows = raw_df.loc[duplicate_mask].copy()

    dedup_df = (
        raw_df.sort_values(
            by=[
                "subject_id_std",
                "available_12m_outcome",
                "followup_12m_mature",
                "completeness_score",
                "end_date_std",
                "start_date_std",
                "row_number_original",
            ],
            ascending=[True, False, False, False, False, False, False],
        )
        .drop_duplicates(subset=["subject_id_std"], keep="first")
        .sort_values(by=["start_date_std", "subject_id_std"], na_position="last")
        .reset_index(drop=True)
    )

    protocol_df = dedup_df.loc[dedup_df["protocol_arm"]].copy().reset_index(drop=True)
    protocol_df["strict_6m_status"] = protocol_df.apply(
        lambda row: classify_strict_outcome(
            unfavorable_std=row["unfavorable_6m_std"],
            component_std=row["unfavorable_6m_component_std"],
            culture_std=row["followup_6m_culture_std"],
            note=row[COLS.note],
            mature=bool(row["followup_6m_mature"]),
        ),
        axis=1,
    )
    protocol_df["strict_12m_status"] = protocol_df.apply(
        lambda row: classify_strict_outcome(
            unfavorable_std=row["unfavorable_12m_std"],
            component_std=row["unfavorable_12m_component_std"],
            culture_std=row["followup_12m_culture_std"],
            note=row[COLS.note],
            mature=bool(row["followup_12m_mature"]),
        ),
        axis=1,
    )

    summary_rows: list[dict[str, object]] = []
    for arm in [None, "BDL", "BDLM"]:
        summary_rows.append(culture_rate(protocol_df, arm, "month2_culture_std"))
        summary_rows.append(culture_rate(protocol_df, arm, "eot_culture_std"))
        summary_rows.append(outcome_rate(protocol_df, arm, "unfavorable_eot_std"))
        summary_rows.append(outcome_rate(protocol_df, arm, "unfavorable_6m_std"))
        summary_rows.append(outcome_rate(protocol_df, arm, "unfavorable_12m_std"))

    summary_df = pd.DataFrame(summary_rows)
    summary_df["label"] = summary_df["metric"].map(
        {
            "month2_culture_std": "2个月培养阴转率",
            "eot_culture_std": "停药时培养阴性率",
            "unfavorable_eot_std": "停药时不良结局率",
            "unfavorable_6m_std": "停药后6月不良结局率",
            "unfavorable_12m_std": "停药后1年不良结局率",
        }
    )
    strict_summary_df = pd.concat(
        [
            strict_outcome_summary(protocol_df, arm, "strict_6m_status")
            for arm in [None, "BDL", "BDLM"]
        ]
        + [
            strict_outcome_summary(protocol_df, arm, "strict_12m_status")
            for arm in [None, "BDL", "BDLM"]
        ],
        ignore_index=True,
    )
    strict_summary_table_df = build_strict_summary_table(protocol_df)
    strict_case_review_df = build_case_review_table(protocol_df)

    duplicate_resolution = duplicate_rows[
        [
            "subject_id_std",
            "row_number_original",
            COLS.arm,
            COLS.start_date,
            COLS.end_date,
            COLS.unfavorable_12m,
            COLS.unfavorable_12m_component,
            "available_12m_outcome",
            "followup_12m_mature",
            "completeness_score",
            COLS.note,
        ]
    ].copy()
    if not duplicate_resolution.empty:
        kept_rows = dedup_df[dedup_df["subject_id_std"].isin(duplicate_resolution["subject_id_std"])]
        kept_rows = kept_rows[["subject_id_std", "row_number_original"]].rename(
            columns={"row_number_original": "kept_row_number"}
        )
        duplicate_resolution = duplicate_resolution.merge(kept_rows, on="subject_id_std", how="left")
        duplicate_resolution["kept_this_row"] = (
            duplicate_resolution["row_number_original"] == duplicate_resolution["kept_row_number"]
        )

    report_text = build_report(
        source_path=source_path,
        output_dir=output_dir,
        as_of_date=as_of_date,
        raw_df=raw_df,
        dedup_df=dedup_df,
        protocol_df=protocol_df,
        duplicate_rows=duplicate_rows,
    )

    cleaned_all_path = output_dir / "cleaned_all_rows.csv"
    cleaned_protocol_path = output_dir / "cleaned_protocol_bdl_bdlm.csv"
    duplicate_path = output_dir / "duplicate_resolution.csv"
    summary_path = output_dir / "arm_summary.csv"
    strict_summary_path = output_dir / "strict_outcome_summary.csv"
    strict_summary_table_path = output_dir / "strict_summary_table.csv"
    strict_summary_table_md_path = output_dir / "strict_summary_table.md"
    strict_case_review_path = output_dir / "strict_case_review.csv"
    strict_case_review_md_path = output_dir / "strict_case_review.md"
    report_path = output_dir / "analysis_report.md"
    metadata_path = output_dir / "analysis_metadata.json"

    raw_df.to_csv(cleaned_all_path, index=False, encoding="utf-8-sig")
    protocol_df.to_csv(cleaned_protocol_path, index=False, encoding="utf-8-sig")
    duplicate_resolution.to_csv(duplicate_path, index=False, encoding="utf-8-sig")
    summary_df.to_csv(summary_path, index=False, encoding="utf-8-sig")
    strict_summary_df.to_csv(strict_summary_path, index=False, encoding="utf-8-sig")
    strict_summary_table_df.to_csv(strict_summary_table_path, index=False, encoding="utf-8-sig")
    strict_case_review_df.to_csv(strict_case_review_path, index=False, encoding="utf-8-sig")
    report_path.write_text(report_text, encoding="utf-8")
    strict_summary_table_md_path.write_text(dataframe_to_markdown(strict_summary_table_df), encoding="utf-8")
    strict_case_review_md_path.write_text(dataframe_to_markdown(strict_case_review_df), encoding="utf-8")
    metadata_path.write_text(
        json.dumps(
            {
                "source_workbook": str(source_path),
                "copied_workbook": str(copied_path),
                "analysis_cutoff_date": as_of_date.date().isoformat(),
                "raw_rows": int(len(raw_df)),
                "deduplicated_rows": int(len(dedup_df)),
                "protocol_rows": int(len(protocol_df)),
                "strict_case_review_rows": int(len(strict_case_review_df)),
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )

    print(report_text)
    print(f"\nSaved outputs to: {output_dir}")


if __name__ == "__main__":
    main()
