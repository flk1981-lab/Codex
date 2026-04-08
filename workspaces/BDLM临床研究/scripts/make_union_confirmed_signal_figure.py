from __future__ import annotations

from pathlib import Path
import shutil

import matplotlib.pyplot as plt
import pandas as pd
from statsmodels.stats.proportion import proportion_confint


INPUT_CSV = Path("reports") / "bdlm_union_first_pass" / "cleaned_protocol_bdl_bdlm.csv"
OUTPUT_DIR = Path("reports") / "bdlm_union_first_pass"
DESKTOP_PNG = Path(r"C:\Users\Administrator\Desktop\union_abstract_figure_confirmed_signal(1).png")


ARM_ORDER = ["Overall", "BDL", "BDLM"]
ARM_LABELS = {"Overall": "Overall", "BDL": "BDL", "BDLM": "BDLM"}

MICRO_METRICS = [
    ("Month 2 culture conversion", "month2_culture_std", "NEG"),
    ("End-of-treatment culture negativity", "eot_culture_std", "NEG"),
]

FAIL_METRICS = [
    ("Confirmed failure/relapse by 6 months", "strict_6m_status"),
    ("Confirmed failure/relapse by 12 months", "strict_12m_status"),
]

COLORS = {
    "Month 2 culture conversion": "#2A9D8F",
    "End-of-treatment culture negativity": "#4FB3A5",
    "Confirmed failure/relapse by 6 months": "#B33A3A",
    "Confirmed failure/relapse by 12 months": "#D46A6A",
}


def compute_metrics(df: pd.DataFrame) -> pd.DataFrame:
    rows: list[dict[str, object]] = []
    for arm in ARM_ORDER:
        sub = df if arm == "Overall" else df[df["arm_std"] == arm]
        for label, col, success in MICRO_METRICS:
            evaluable = sub[col].isin(["NEG", "POS"])
            n = int(evaluable.sum())
            x = int((sub.loc[evaluable, col] == success).sum())
            lo, hi = proportion_confint(x, n, method="beta") if n else (None, None)
            rows.append(
                {
                    "panel": "micro",
                    "arm": arm,
                    "metric": label,
                    "x": x,
                    "n": n,
                    "pct": x / n * 100.0 if n else None,
                    "lo": lo * 100.0 if lo is not None else None,
                    "hi": hi * 100.0 if hi is not None else None,
                }
            )
        for label, col in FAIL_METRICS:
            n = int(len(sub))
            x = int((sub[col] == "TRUE_UNFAVORABLE").sum())
            lo, hi = proportion_confint(x, n, method="beta") if n else (None, None)
            rows.append(
                {
                    "panel": "fail",
                    "arm": arm,
                    "metric": label,
                    "x": x,
                    "n": n,
                    "pct": x / n * 100.0 if n else None,
                    "lo": lo * 100.0 if lo is not None else None,
                    "hi": hi * 100.0 if hi is not None else None,
                }
            )
    return pd.DataFrame(rows)


def plot_panel_micro(ax: plt.Axes, df: pd.DataFrame) -> None:
    y_map = {
        ("Overall", "Month 2 culture conversion"): 5.2,
        ("Overall", "End-of-treatment culture negativity"): 4.4,
        ("BDL", "Month 2 culture conversion"): 3.0,
        ("BDL", "End-of-treatment culture negativity"): 2.2,
        ("BDLM", "Month 2 culture conversion"): 0.8,
        ("BDLM", "End-of-treatment culture negativity"): 0.0,
    }
    for _, row in df.iterrows():
        y = y_map[(row["arm"], row["metric"])]
        ax.errorbar(
            row["pct"],
            y,
            xerr=[[row["pct"] - row["lo"]], [row["hi"] - row["pct"]]],
            fmt="o",
            color=COLORS[row["metric"]],
            ecolor=COLORS[row["metric"]],
            elinewidth=2,
            capsize=3,
            markersize=7,
        )
        ax.text(
            min(row["pct"] + 3.0, 101),
            y - 0.22,
            f"{row['x']}/{row['n']} ({row['pct']:.1f}%)",
            va="bottom",
            ha="left",
            fontsize=9,
            color="#333333",
        )

    ax.set_xlim(0, 105)
    ax.set_ylim(-0.6, 5.8)
    ax.set_yticks([4.8, 2.6, 0.4])
    ax.set_yticklabels(["Overall", "BDL", "BDLM"], fontsize=10)
    ax.set_xticks([0, 20, 40, 60, 80, 100])
    ax.set_xticklabels([f"{x}%" for x in [0, 20, 40, 60, 80, 100]], fontsize=9)
    ax.grid(axis="x", color="#EAEAEA", linewidth=0.8)
    ax.set_axisbelow(True)
    ax.set_title("Microbiological response", loc="left", fontsize=12.5, fontweight="bold", pad=8)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_visible(False)
    ax.tick_params(axis="y", length=0)
    ax.spines["bottom"].set_color("#BEBEBE")

    for y, label in [(5.2, "Month 2"), (4.4, "EOT"), (3.0, "Month 2"), (2.2, "EOT"), (0.8, "Month 2"), (0.0, "EOT")]:
        ax.text(-8, y, label, ha="right", va="center", fontsize=8.8, color="#666666", clip_on=False)


def plot_panel_fail(ax: plt.Axes, df: pd.DataFrame) -> None:
    y_map = {
        ("Overall", "Confirmed failure/relapse by 6 months"): 5.2,
        ("Overall", "Confirmed failure/relapse by 12 months"): 4.4,
        ("BDL", "Confirmed failure/relapse by 6 months"): 3.0,
        ("BDL", "Confirmed failure/relapse by 12 months"): 2.2,
        ("BDLM", "Confirmed failure/relapse by 6 months"): 0.8,
        ("BDLM", "Confirmed failure/relapse by 12 months"): 0.0,
    }
    for _, row in df.iterrows():
        y = y_map[(row["arm"], row["metric"])]
        ax.errorbar(
            row["pct"],
            y,
            xerr=[[row["pct"] - row["lo"]], [row["hi"] - row["pct"]]],
            fmt="o",
            color=COLORS[row["metric"]],
            ecolor=COLORS[row["metric"]],
            elinewidth=2,
            capsize=3,
            markersize=7,
        )
        ax.text(
            min(row["hi"] + 2.0, 27.5),
            y - 0.22,
            f"{row['x']}/{row['n']} ({row['pct']:.1f}%)",
            va="bottom",
            ha="left",
            fontsize=9,
            color="#333333",
        )

    ax.set_xlim(0, 28)
    ax.set_ylim(-0.6, 5.8)
    ax.set_yticks([4.8, 2.6, 0.4])
    ax.set_yticklabels(["Overall", "BDL", "BDLM"], fontsize=10)
    ax.set_xticks([0, 5, 10, 15, 20, 25])
    ax.set_xticklabels([f"{x}%" for x in [0, 5, 10, 15, 20, 25]], fontsize=9)
    ax.grid(axis="x", color="#EAEAEA", linewidth=0.8)
    ax.set_axisbelow(True)
    ax.set_title("Confirmed post-treatment failure/relapse", loc="left", fontsize=12.5, fontweight="bold", pad=8)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_visible(False)
    ax.tick_params(axis="y", length=0)
    ax.spines["bottom"].set_color("#BEBEBE")

    for y, label in [(5.2, "6 months"), (4.4, "12 months"), (3.0, "6 months"), (2.2, "12 months"), (0.8, "6 months"), (0.0, "12 months")]:
        ax.text(-2.4, y, label, ha="right", va="center", fontsize=8.8, color="#666666", clip_on=False)


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    df = pd.read_csv(INPUT_CSV)
    metrics = compute_metrics(df)

    fig, axes = plt.subplots(
        1,
        2,
        figsize=(12.8, 6.0),
        gridspec_kw={"width_ratios": [1.02, 0.88]},
    )
    fig.patch.set_facecolor("white")

    plot_panel_micro(axes[0], metrics[metrics["panel"] == "micro"])
    plot_panel_fail(axes[1], metrics[metrics["panel"] == "fail"])

    handles = [
        plt.Line2D([0], [0], marker="o", color=COLORS["Month 2 culture conversion"], linestyle="", markersize=7),
        plt.Line2D([0], [0], marker="o", color=COLORS["End-of-treatment culture negativity"], linestyle="", markersize=7),
        plt.Line2D([0], [0], marker="o", color=COLORS["Confirmed failure/relapse by 6 months"], linestyle="", markersize=7),
        plt.Line2D([0], [0], marker="o", color=COLORS["Confirmed failure/relapse by 12 months"], linestyle="", markersize=7),
    ]
    labels = [
        "Month 2 culture conversion",
        "End-of-treatment culture negativity",
        "Confirmed failure/relapse by 6 months",
        "Confirmed failure/relapse by 12 months",
    ]
    fig.legend(
        handles,
        labels,
        ncol=2,
        frameon=False,
        loc="lower center",
        bbox_to_anchor=(0.5, -0.01),
        fontsize=9.5,
        columnspacing=1.6,
        handletextpad=0.5,
    )

    fig.suptitle(
        "Key efficacy signals of a pretomanid-free BDL/M strategy for RR-TB in China",
        fontsize=15,
        fontweight="bold",
        y=0.97,
    )

    fig.subplots_adjust(left=0.08, right=0.98, top=0.86, bottom=0.16, wspace=0.42)

    png_path = OUTPUT_DIR / "union_abstract_figure_confirmed_signal.png"
    pdf_path = OUTPUT_DIR / "union_abstract_figure_confirmed_signal.pdf"
    svg_path = OUTPUT_DIR / "union_abstract_figure_confirmed_signal.svg"
    fig.savefig(png_path, dpi=300, bbox_inches="tight", facecolor="white")
    fig.savefig(pdf_path, bbox_inches="tight", facecolor="white")
    fig.savefig(svg_path, bbox_inches="tight", facecolor="white")
    if DESKTOP_PNG.parent.exists():
        shutil.copy2(png_path, DESKTOP_PNG)
    plt.close(fig)

    legend = (
        "Figure. Key efficacy signals of a pretomanid-free BDL/M strategy for rifampicin-resistant tuberculosis in China. "
        "Left panel shows month-2 culture conversion and end-of-treatment culture negativity; right panel shows confirmed "
        "failure/relapse by 6 and 12 months after treatment completion. Points represent proportions and whiskers exact 95% "
        "confidence intervals. Loss to follow-up and withdrawal are described in the text; participants classified as lost to "
        "follow-up at 12 months remain under active tracing."
    )
    (OUTPUT_DIR / "union_abstract_figure_confirmed_signal_legend.txt").write_text(legend, encoding="utf-8")

    print(f"Saved: {png_path}")
    print(f"Saved: {pdf_path}")
    print(f"Saved: {svg_path}")
    print(f"Saved: {DESKTOP_PNG}")
    print(f"Saved: {OUTPUT_DIR / 'union_abstract_figure_confirmed_signal_legend.txt'}")


if __name__ == "__main__":
    main()
