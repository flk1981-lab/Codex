# TB Research Email Automation Prompts

These prompts are designed for Codex automations that both learn from the web and email a `.docx` report.

Recommended model settings:

- model: `gpt-5.4`
- reasoning_effort: `high`
- cwd: `D:/Codex`

## Daily Mentor Email

Suggested name: `TB Daily Mentor Email`

Prompt:

```text
Use [$clinical-research-watch](C:/Users/Administrator/.codex/skills/clinical-research-watch/SKILL.md) and [$tb-research-advisor](D:/Codex/.agents/skills/tb-research-advisor/SKILL.md) to scan the newest tuberculosis clinical research from top journals, PubMed, WHO, and official guideline issuers. Limit to the last 72 hours. Prioritize title-and-abstract triage, read in depth only up to 2 high-value items, and ignore low-signal or repetitive content.

Produce one bilingual mentor-style report in the same document.
Order is mandatory:
1. English version first
2. Chinese version second

Use the same section structure in both languages:
1. 3 to 5 most worthwhile items today / 今日最值得看的3到5条
2. Mentor-style deep dive / 重点论文导师拆解
3. New methods lessons from this run / 本次新增的方法学启发
4. What matters most for our study design thinking / 对我们研究设计最有价值的提醒
5. The one thing you should study next / 你今天最值得补的一件事

Clearly separate FACT, INFERENCE, and HYPOTHESIS in both language sections.

Save the markdown report to `D:/Codex/reports/tb-research/pending/tb-daily-<today>.md`.
Then run `powershell -File D:/Codex/scripts/Send-TBReportEmail.ps1 -MarkdownPath <that markdown path> -ReportTitle "TB Daily Mentor Report <today>" -Subject "TB Daily Mentor Report <today>"`.

Update local TB knowledge files only when there is a genuinely durable, decision-relevant finding.
```

Recommended schedule:

- Every day at `07:30`

Suggested RRULE:

- `FREQ=WEEKLY;BYDAY=MO,TU,WE,TH,FR,SA,SU;BYHOUR=7;BYMINUTE=30`

## Weekly Mentor Email

Suggested name: `TB Weekly Mentor Email`

Prompt:

```text
Use [$clinical-research-watch](C:/Users/Administrator/.codex/skills/clinical-research-watch/SKILL.md), [$systematic-review-top-journal](C:/Users/Administrator/.codex/skills/systematic-review-top-journal/SKILL.md), and [$tb-research-advisor](D:/Codex/.agents/skills/tb-research-advisor/SKILL.md) to review the last 14 days of tuberculosis clinical research from top journals, PubMed, WHO, and official guideline issuers. Read in depth only up to 4 high-value studies or updates.

Produce one bilingual mentor-style report in the same document.
Order is mandatory:
1. English version first
2. Chinese version second

Use the same section structure in both languages:
1. Evidence changes this week / 本周证据变化
2. The one paper most worth reading closely / 最值得精读的一篇论文
3. The most important methods lesson this week / 本周最重要的方法学收获
4. Key unresolved questions / 仍未解决的关键问题
5. The one research direction most worth advancing / 最值得推进的1个研究方向
6. The one thing you should study next week / 你下周最值得补的1件事

Clearly separate FACT, INFERENCE, and HYPOTHESIS in both language sections.

Save the markdown report to `D:/Codex/reports/tb-research/pending/tb-weekly-<today>.md`.
Then run `powershell -File D:/Codex/scripts/Send-TBReportEmail.ps1 -MarkdownPath <that markdown path> -ReportTitle "TB Weekly Mentor Review <today>" -Subject "TB Weekly Mentor Review <today>"`.

Update local TB knowledge files only when the addition is durable and decision-relevant.
```

Recommended schedule:

- Every Sunday at `20:00`

Suggested RRULE:

- `FREQ=WEEKLY;BYDAY=SU;BYHOUR=20;BYMINUTE=0`
