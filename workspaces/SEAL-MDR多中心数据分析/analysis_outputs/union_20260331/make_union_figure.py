from __future__ import annotations

import math
from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd


def wilson_ci(x: int, n: int, z: float = 1.959963984540054) -> tuple[float, float]:
    if n == 0:
        return (0.0, 0.0)
    p = x / n
    denom = 1 + z * z / n
    center = (p + z * z / (2 * n)) / denom
    half = z * math.sqrt((p * (1 - p) + z * z / (4 * n)) / n) / denom
    return center - half, center + half


def build_layout(groups: list[str], sublabels: list[str], gap: float = 0.9) -> tuple[list[float], list[str], dict[str, float]]:
    y_positions: list[float] = []
    y_labels: list[str] = []
    group_centers: dict[str, float] = {}

    current = 0.0
    for group in groups:
        group_rows = []
        for sub in sublabels:
            y_positions.append(current)
            y_labels.append(sub)
            group_rows.append(current)
            current += 1.0
        group_centers[group] = sum(group_rows) / len(group_rows)
        current += gap

    return y_positions, y_labels, group_centers


def add_rate_columns(df: pd.DataFrame, event_col: str, total_col: str, prefix: str) -> pd.DataFrame:
    rates = []
    low = []
    high = []
    labels = []
    for x, n in zip(df[event_col], df[total_col]):
        p = x / n if n else float("nan")
        lo, hi = wilson_ci(int(x), int(n)) if n else (float("nan"), float("nan"))
        rates.append(100 * p)
        low.append(100 * lo)
        high.append(100 * hi)
        labels.append(f"{int(x)}/{int(n)} ({100 * p:.1f}%)")
    df[f"{prefix}_rate"] = rates
    df[f"{prefix}_low"] = low
    df[f"{prefix}_high"] = high
    df[f"{prefix}_label"] = labels
    return df


def main() -> None:
    base = Path(__file__).resolve().parent
    arm_path = base / "arm_summary.csv"
    group_path = base / "group_summary.csv"
    out_png = base / "SEAL-MDR_union_figure_20260331.png"
    out_pdf = base / "SEAL-MDR_union_figure_20260331.pdf"

    arm_df = pd.read_csv(arm_path)
    group_df = pd.read_csv(group_path)

    arm_df = arm_df[arm_df["group"].isin(["A", "B", "C", "D"])].copy()
    arm_df["group"] = pd.Categorical(arm_df["group"], ["A", "B", "C", "D"], ordered=True)
    arm_df = arm_df.sort_values("group").reset_index(drop=True)

    overall = group_df[group_df["group"] == "A-D pooled"].copy()
    overall.loc[:, "group"] = "Overall"
    overall = overall[
        [
            "group",
            "month2_neg",
            "month2_eval",
            "eot_neg",
            "eot_eval",
            "unfavorable_6m",
            "assessed_6m",
            "unfavorable_12m",
            "assessed_12m",
        ]
    ]

    arm_keep = arm_df[
        [
            "group",
            "month2_neg",
            "month2_eval",
            "eot_neg",
            "eot_eval",
            "unfavorable_6m",
            "assessed_6m",
            "unfavorable_12m",
            "assessed_12m",
        ]
    ].copy()
    plot_df = pd.concat([overall, arm_keep], ignore_index=True)

    for event_col, total_col, prefix in [
        ("month2_neg", "month2_eval", "m2"),
        ("eot_neg", "eot_eval", "eot"),
        ("unfavorable_6m", "assessed_6m", "u6"),
        ("unfavorable_12m", "assessed_12m", "u12"),
    ]:
        add_rate_columns(plot_df, event_col, total_col, prefix)

    groups = ["Overall", "A", "B", "C", "D"]
    left_y, left_labels, left_centers = build_layout(groups, ["6 months", "12 months"], gap=0.95)
    right_y, right_labels, right_centers = build_layout(groups, ["Month 2", "EOT"], gap=0.95)

    plt.rcParams.update(
        {
            "font.family": "DejaVu Sans",
            "font.size": 10,
            "axes.titlesize": 12.5,
            "axes.labelsize": 10.5,
            "xtick.labelsize": 9.5,
            "ytick.labelsize": 9.5,
        }
    )

    fig, (ax_left, ax_right) = plt.subplots(
        1,
        2,
        figsize=(12.8, 6.2),
        gridspec_kw={"width_ratios": [1.0, 1.15]},
    )
    fig.subplots_adjust(left=0.16, right=0.98, top=0.82, bottom=0.14, wspace=0.30)

    colors = {
        "u6": "#991B1B",
        "u12": "#C62828",
        "m2": "#0F766E",
        "eot": "#0B8F83",
        "grid": "#D1D5DB",
        "text": "#111827",
        "muted": "#6B7280",
    }

    # Left panel: post-treatment unfavorable outcomes
    for idx, row in plot_df.iterrows():
        y6 = left_y[idx * 2]
        y12 = left_y[idx * 2 + 1]
        for prefix, yy, label in [("u6", y6, "6 months"), ("u12", y12, "12 months")]:
            rate = row[f"{prefix}_rate"]
            lo = row[f"{prefix}_low"]
            hi = row[f"{prefix}_high"]
            ax_left.errorbar(
                rate,
                yy,
                xerr=[[rate - lo], [hi - rate]],
                fmt="o",
                ms=7,
                color=colors[prefix],
                ecolor=colors[prefix],
                elinewidth=2,
                capsize=3.5,
                capthick=1.8,
                zorder=3,
            )
            ax_left.text(
                min(rate + 1.8, 97),
                yy - 0.34,
                row[f"{prefix}_label"],
                fontsize=8.9,
                ha="left",
                va="bottom",
                color=colors["text"],
            )

    ax_left.set_title("Post-treatment unfavorable outcome", fontweight="bold", pad=10)
    ax_left.set_xlim(0, 100)
    ax_left.set_xlabel("Rate (%)", fontweight="bold")
    ax_left.set_yticks(left_y)
    ax_left.set_yticklabels(left_labels)
    ax_left.invert_yaxis()
    ax_left.grid(axis="x", color=colors["grid"], linewidth=1)
    ax_left.set_axisbelow(True)
    ax_left.spines["top"].set_visible(False)
    ax_left.spines["right"].set_visible(False)
    for spine in ["left", "bottom"]:
        ax_left.spines[spine].set_linewidth(1)
        ax_left.spines[spine].set_color(colors["muted"])

    for group, yc in left_centers.items():
        ax_left.text(
            -0.02,
            yc,
            group,
            transform=ax_left.get_yaxis_transform(),
            ha="right",
            va="center",
            fontsize=10.7,
            fontweight="bold",
            color=colors["text"],
        )

    ax_left.scatter([], [], color=colors["u6"], s=80, label="Unfavorable outcome by 6 months")
    ax_left.scatter([], [], color=colors["u12"], s=80, label="Unfavorable outcome by 12 months")
    ax_left.legend(
        loc="lower left",
        bbox_to_anchor=(0.0, -0.24),
        frameon=False,
        fontsize=9,
        handletextpad=0.6,
    )

    # Right panel: microbiological response
    for idx, row in plot_df.iterrows():
        ym2 = right_y[idx * 2]
        yeot = right_y[idx * 2 + 1]
        for prefix, yy in [("m2", ym2), ("eot", yeot)]:
            rate = row[f"{prefix}_rate"]
            lo = row[f"{prefix}_low"]
            hi = row[f"{prefix}_high"]
            ax_right.errorbar(
                rate,
                yy,
                xerr=[[rate - lo], [hi - rate]],
                fmt="o",
                ms=7,
                color=colors[prefix],
                ecolor=colors[prefix],
                elinewidth=2,
                capsize=3.5,
                capthick=1.8,
                zorder=3,
            )
            ax_right.text(
                min(rate + 1.8, 98.5),
                yy - 0.34,
                row[f"{prefix}_label"],
                fontsize=8.9,
                ha="left",
                va="bottom",
                color=colors["text"],
            )

    ax_right.set_title("Microbiological response", fontweight="bold", pad=10)
    ax_right.set_xlim(0, 105)
    ax_right.set_xlabel("Rate (%)", fontweight="bold")
    ax_right.set_yticks(right_y)
    ax_right.set_yticklabels(right_labels)
    ax_right.invert_yaxis()
    ax_right.grid(axis="x", color=colors["grid"], linewidth=1)
    ax_right.set_axisbelow(True)
    ax_right.spines["top"].set_visible(False)
    ax_right.spines["right"].set_visible(False)
    for spine in ["left", "bottom"]:
        ax_right.spines[spine].set_linewidth(1)
        ax_right.spines[spine].set_color(colors["muted"])

    for group, yc in right_centers.items():
        ax_right.text(
            -0.02,
            yc,
            group,
            transform=ax_right.get_yaxis_transform(),
            ha="right",
            va="center",
            fontsize=10.7,
            fontweight="bold",
            color=colors["text"],
        )

    ax_right.scatter([], [], color=colors["m2"], s=80, label="Month 2 culture negative")
    ax_right.scatter([], [], color=colors["eot"], s=80, label="End-of-treatment culture negative")
    ax_right.legend(
        loc="lower left",
        bbox_to_anchor=(0.0, -0.24),
        frameon=False,
        fontsize=9,
        handletextpad=0.6,
    )

    fig.suptitle(
        "Key efficacy signals of SEAL-MDR across protocol-defined regimens in China",
        fontsize=18,
        fontweight="bold",
        y=0.98,
    )

    fig.savefig(out_png, dpi=320, bbox_inches="tight", facecolor="white")
    fig.savefig(out_pdf, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print(out_png)
    print(out_pdf)


if __name__ == "__main__":
    main()
