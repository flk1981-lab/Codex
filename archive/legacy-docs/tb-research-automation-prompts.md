# TB Research Automation Prompts

These prompts are written for Codex automations. They describe only the task, not the schedule.

Recommended model settings:

- model: `gpt-5.4`
- reasoning_effort: `high`
- cwd: `D:/Codex`

## Daily TB Radar

Suggested name: `TB Daily Radar`

Prompt:

```text
Use [$clinical-research-watch](C:/Users/Administrator/.codex/skills/clinical-research-watch/SKILL.md) and [$tb-research-advisor](D:/Codex/.agents/skills/tb-research-advisor/SKILL.md) to scan the newest tuberculosis clinical research that could matter for practice, study design, or research opportunities. Use primary sources on the web and verify unstable facts. Focus on eligible items from the last 3 days, plus any new official TB guideline or policy updates and any high-signal preprints clearly labeled as not peer reviewed.

For each eligible item, extract:
1. question and design
2. population and setting
3. intervention or exposure and comparator
4. primary endpoint and follow-up
5. main result or key methodological signal
6. major validity limitations
7. external validity for TB practice or TB study design

Then update the repo knowledge files in `knowledge/tb-research/` with dated bullets and source links:
- `guideline-watch.md`
- `landmark-trials.md`
- `methods-playbook.md`
- `research-gap-backlog.md`
- `paper-notes.md`

Finally produce a concise Chinese inbox report with these sections:
1. 今日最值得关注
2. 新增指南/政策
3. 新增方法学启发
4. 新增研究空白
5. 对我最有价值的下一步

If no eligible items are found, say so explicitly, keep the report concise, and still state whether any knowledge files were updated.

Do not fabricate study details, sample sizes, effect sizes, guideline claims, or dates.
```

Recommended schedule:

- Every day at `07:30`

Suggested RRULE:

- `FREQ=WEEKLY;BYDAY=MO,TU,WE,TH,FR,SA,SU;BYHOUR=7;BYMINUTE=30`

## Weekly TB Methods And Ideas

Suggested name: `TB Weekly Methods`

Prompt:

```text
Use [$clinical-research-watch](C:/Users/Administrator/.codex/skills/clinical-research-watch/SKILL.md), [$systematic-review-top-journal](C:/Users/Administrator/.codex/skills/systematic-review-top-journal/SKILL.md), and [$tb-research-advisor](D:/Codex/.agents/skills/tb-research-advisor/SKILL.md) to review the last 14 days of tuberculosis clinical research and convert it into one methods-first weekly synthesis plus a prioritized idea list.

Tasks:
1. Identify the most decision-relevant TB studies, guideline updates, and high-signal preprints from the last 14 days using primary sources.
2. Summarize what changed in the evidence base.
3. Extract reusable methodological lessons for TB trial design, cohort design, endpoint selection, causal inference, implementation studies, and evidence synthesis.
4. Decide which uncertainties remain truly unanswered.
5. Generate and rank 3 candidate TB research directions for me.

For each ranked research direction, include:
- one-sentence question
- why the question matters clinically
- why current evidence is still insufficient
- best design
- preferred primary endpoint
- main bias or confounding risk
- feasibility constraints
- minimum viable next step

Update these files with dated bullets and source links:
- `guideline-watch.md`
- `landmark-trials.md`
- `methods-playbook.md`
- `research-gap-backlog.md`
- `question-ideas.md`
- `paper-notes.md`

Produce a Chinese inbox report with these sections:
1. 本周证据变化
2. 本周最重要的方法学启发
3. 仍未解决的关键问题
4. 排名前3的研究方向
5. 下周最值得推进的一项

If no strong new items are found, say so clearly and focus on consolidating durable methodological lessons from the recent evidence.

Do not fabricate study details, effect sizes, or novelty claims.
```

Recommended schedule:

- Every Sunday at `20:00`

Suggested RRULE:

- `FREQ=WEEKLY;BYDAY=SU;BYHOUR=20;BYMINUTE=0`
