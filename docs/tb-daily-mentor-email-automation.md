# TB Daily Mentor Email Automation

This machine uses a Mac-native rebuild of the TB daily mentor email workflow.

Files:

- `scripts/set_tb_report_mail_config.py`: stores SMTP settings for the report email.
- `scripts/send_tb_report_email.py`: converts a markdown report to `.docx` and emails it.

Recommended setup for QQ Mail:

```bash
python3 /Users/flk1981/Codex/scripts/set_tb_report_mail_config.py \
  --sender-email "your_sender@qq.com" \
  --recipient-email "your_recipient@qq.com" \
  --smtp-password "your_app_password"
```

Expected config location:

- `/Users/flk1981/.codex/secrets/tb-report-mail.json`

The Codex automation should:

1. Read the newest TB clinical research from primary sources.
2. Save a bilingual daily report to `/Users/flk1981/Codex/reports/tb-research/pending/`.
3. Run `python3 /Users/flk1981/Codex/scripts/send_tb_report_email.py ...` to email the report as a `.docx` attachment.
