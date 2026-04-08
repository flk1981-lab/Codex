from __future__ import annotations

from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd


INPUT_CSV = Path("reports") / "bdlm_union_first_pass" / "strict_outcome_summary.csv"
OUTPUT_DIR = Path("reports") / "bdlm_union_first_pass"


STATUS_ORDER = [
    "TRUE_UNFAVORABLE",
    "LOST_TO_FOLLOWUP",
    "WITHDRAWN",
    "OTHER_UNFAVORABLE",
    "FAVORABLE",
    "PENDING",
]

STATUS_LABELS = {
    "TRUE_UNFAVORABLE": "True failure/relapse",
    "LOST_TO_FOLLOWUP": "Lost to follow-up",
    "WITHDRAWN": "Withdrawn",
    "OTHER_UNFAVORABLE": "Other non-protocol outcome",
    "FAVORABLE": "Event-free",
    "PENDING": "Not yet due",
}

STATUS_COLORS = {
    "TRUE_UNFAVORABLE": "#B33A3A",
    "LOST_TO_FOLLOWUP": "#E59D2F",
    "WITHDRAWN": "#A8704E",
    "OTHER_UNFAVORABLE": "#7E7E7E",
    "FAVORABLE": "#2A9D8F",
    "PENDING": "#D9D9D9",
}

PANEL_META = {
    "strict_6m_status": {
        "title": "Primary endpoint: 6 months after treatment completion",
        "note": "Strict endpoints separate confirmed biological events from follow-up attrition.",
    },
    "strict_12m_status": {
        "title": "Secondary endpoint: 12 months after treatment completion",
        "note": "Participants classified as 12-month LTFU remain under active tracing.",
    },
}

ARM_ORDER = ["Overall", "BDL", "BDLM"]


def load_plot_data(path: Path) -> pd.DataFrame:
    df = pd.read_csv(path)
    df = df[df["status"].isin(STATUS_ORDER)].copy()
    df["status"] = pd.Categorical(df["status"], categories=STATUS_ORDER, ordered=True)
    df["arm"] = pd.Categorical(df["arm"], categories=ARM_ORDER, ordered=True)
    return df.sort_values(["status_col", "arm", "status"]).reset_index(drop=True)


def add_segment_labels(ax: plt.Axes, left: float, width: float, y: float, count: int, total: int) -> None:
    if count == 0:
        return
    pct = 100 * count / total if total else 0
    if width >= 8:
        ax.text(
            left + width / 2,
            y,
            f"{count}",
            ha="center",
            va="center",
            fontsize=9,
            color="white" if width >= 16 else "black",
            fontweight="bold",
        )
    elif pct >= 1.5:
        ax.text(
            left + width + 0.8,
            y,
            f"{count}",
            ha="left",
            va="center",
            fontsize=8,
            color="#333333",
        )


def plot_panel(ax: plt.Axes, panel_df: pd.DataFrame, panel_key: str) -> None:
    y_positions = list(range(len(ARM_ORDER)))
    ax.set_xlim(0, 100)
    ax.set_ylim(-0.6, len(ARM_ORDER) - 0.4)
    ax.set_yticks(y_positions)

    y_labels = []
    totals = {}
    for arm in ARM_ORDER:
        total = int(panel_df.loc[panel_df["arm"] == arm, "total"].iloc[0])
        totals[arm] = total
        y_labels.append(f"{arm} (n={total})")
    ax.set_yticklabels(y_labels, fontsize=10)

    for y, arm in zip(y_positions, ARM_ORDER):
        subset = panel_df[panel_df["arm"] == arm]
        left = 0.0
        total = totals[arm]
        for status in STATUS_ORDER:
            row = subset[subset["status"] == status].iloc[0]
            width = float(row["percent"])
            count = int(row["count"])
            if width > 0:
                ax.barh(
                    y,
                    width,
                    left=left,
                    color=STATUS_COLORS[status],
                    edgecolor="white",
                    linewidth=1.2,
                    height=0.62,
                )
                add_segment_labels(ax, left, width, y, count, total)
            left += width

    ax.invert_yaxis()
    ax.set_xticks([0, 20, 40, 60, 80, 100])
    ax.set_xticklabels([f"{x}%" for x in [0, 20, 40, 60, 80, 100]], fontsize=9)
    ax.grid(axis="x", color="#EAEAEA", linewidth=0.8)
    ax.set_axisbelow(True)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_visible(False)
    ax.spines["bottom"].set_color("#BEBEBE")
    ax.tick_params(axis="y", length=0)
    ax.set_title(
        PANEL_META[panel_key]["title"],
        loc="left",
        fontsize=12,
        fontweight="bold",
        pad=10,
    )
    ax.text(
        0,
        -0.95,
        PANEL_META[panel_key]["note"],
        fontsize=9,
        color="#555555",
        ha="left",
        va="bottom",
        transform=ax.transData,
    )


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    df = load_plot_data(INPUT_CSV)

    fig, axes = plt.subplots(2, 1, figsize=(11.2, 7.8), sharex=True)
    fig.patch.set_facecolor("white")

    for ax, panel_key in zip(axes, ["strict_6m_status", "strict_12m_status"]):
        plot_panel(ax, df[df["status_col"] == panel_key], panel_key)

    handles = [
        plt.Rectangle((0, 0), 1, 1, color=STATUS_COLORS[status])
        for status in STATUS_ORDER
    ]
    labels = [STATUS_LABELS[status] for status in STATUS_ORDER]
    fig.legend(
        handles,
        labels,
        ncol=3,
        frameon=False,
        loc="upper center",
        bbox_to_anchor=(0.5, 0.09),
        fontsize=9.5,
        columnspacing=1.6,
        handlelength=1.4,
    )

    fig.suptitle(
        "Strict post-treatment outcomes of a pretomanid-free BDL/M strategy for RR-TB in China",
        fontsize=15,
        fontweight="bold",
        y=0.98,
    )
    fig.text(
        0.01,
        0.935,
        "Bars show the full cohort composition, highlighting that apparent long-term unfavorable outcomes "
        "were driven mainly by follow-up attrition rather than confirmed failure/relapse.",
        fontsize=10,
        color="#444444",
    )
    fig.text(
        0.01,
        0.02,
        "Numbers inside or beside segments indicate case counts. Confirmed events were defined by failure/relapse components, "
        "positive follow-up culture, or explicit documentation of relapse/recurrence.",
        fontsize=8.8,
        color="#555555",
    )

    fig.tight_layout(rect=[0, 0.08, 1, 0.92])

    png_path = OUTPUT_DIR / "union_abstract_figure_strict_outcomes.png"
    pdf_path = OUTPUT_DIR / "union_abstract_figure_strict_outcomes.pdf"
    svg_path = OUTPUT_DIR / "union_abstract_figure_strict_outcomes.svg"
    fig.savefig(png_path, dpi=300, bbox_inches="tight", facecolor="white")
    fig.savefig(pdf_path, bbox_inches="tight", facecolor="white")
    fig.savefig(svg_path, bbox_inches="tight", facecolor="white")
    plt.close(fig)

    legend_text = (
        "Figure. Strict post-treatment outcomes after a pretomanid-free BDL/M strategy for rifampicin-resistant tuberculosis "
        "in China. Stacked bars show outcome composition overall and by strategy group at 6 and 12 months after treatment "
        "completion. Confirmed true failure/relapse was rare, whereas apparent unfavorable outcomes at 12 months were driven "
        "mainly by loss to follow-up and withdrawal. Participants classified as lost to follow-up at 12 months remain under "
        "active tracing."
    )
    (OUTPUT_DIR / "union_abstract_figure_legend.txt").write_text(legend_text, encoding="utf-8")

    print(f"Saved: {png_path}")
    print(f"Saved: {pdf_path}")
    print(f"Saved: {svg_path}")
    print(f"Saved: {OUTPUT_DIR / 'union_abstract_figure_legend.txt'}")


if __name__ == "__main__":
    main()
