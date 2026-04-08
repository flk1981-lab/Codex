# SEAL-MDR Data Quality Notes

Analysis date: 2026-03-31
Workbook snapshot: `结核病临床试验《SEAL-MDR》_Union摘要_2023-03-16【深圳三院+其他分中心】.xlsx`

## Rules used

- Exact duplicate spreadsheet rows were removed before aggregation.
- Arm labels were normalized conservatively into `A/B/C/D/E`; entries that could not be mapped reliably were excluded from arm-level summaries.
- Culture endpoints used an evaluable denominator of explicit positive or negative results only.
- Post-treatment 6-month and 12-month unfavorable outcomes were analyzed only among records with sufficient follow-up time from treatment end.
- Outcome values such as `否/无` were treated as `No`; explicit event strings such as `是/有/死亡/复发/失败/进展/依从性差/自行停结核药` were treated as `Yes`.
- Free-text statuses indicating non-assessable follow-up such as `失访/未复诊/不可评估` were treated as missing.

## Quality checks

- Raw rows: 782
- Rows after exact de-duplication: 756
- Exact duplicate rows removed: 26
- Record-level rows included in A-E analysis: 733
- Unique ID strings in A-E analysis: 730
- Rows with duplicated ID strings: 6
- Number of duplicated ID strings: 3
- Rows with non-empty but unclassifiable arm labels: 20
- Rows excluded because the arm field looked administrative (for example screening failure / exit): 2

## Interpretation cautions

- This is safest to report as an interim record-level multicenter analysis.
- The `受试者ID` field behaves like a name string rather than a guaranteed unique subject key.
- Post-treatment ascertainment is incomplete, especially outside the largest centers and in arms D/E.
- The workbook is well-suited to conference-style descriptive results, but not yet to a final SAP-grade comparative analysis.
