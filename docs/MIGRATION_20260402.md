# Migration Record: 2026-04-02

This workspace was refreshed to match the GitHub `origin/main` structure on April 2, 2026.

## Backups created before refresh

- Full workspace zip: `D:\Codex_backups\codex-backup-20260402-132718.zip`
- Git history bundle: `D:\Codex_backups\codex-history-20260402-132623.bundle`
- Local recovery branch: `backup/pre-github-refresh-20260402-132811`

## Active baseline after refresh

The active root now follows the GitHub layout:

- `docs/`
- `knowledge/`
- `reports/`
- `scripts/`
- `workspaces/`
- `archive/`

## Restored legacy content

Moved into `workspaces/`:

- `BDLM临床研究`
- `SEAL-MDR多中心数据分析`
- `TDSC_培训与执行版文档包_2026-03-25`
- `TDSC_研究操作手册`
- `TDSC项目 研究方案和SOPs`
- `四个队列中的NTM分析`
- `胸腰椎结核病系统综述_20260327`

Moved into `reports/`:

- prior `reports/` content from the local backup

Moved into `archive/`:

- `旧文档`
- `_installers`
- `_docx_edit_tmp`
- `bdlm_temp`
- legacy GitHub helper scripts from the old root
- older automation prompt docs
- superseded PowerShell mail/report scripts

Project-specific BDLM analysis scripts were restored under:

- `workspaces/BDLM临床研究/scripts/`

## Intentionally left out of the active root

The following were preserved only in backup artifacts and not restored into the live root:

- `.codex/`
- `.agents/`
- old root `.git/` contents
- duplicate or superseded top-level helper files
- older `knowledge/` snapshots that would overwrite the new GitHub baseline

## Notes

- `workspaces/` content is git-ignored by the root repository and is meant to be managed as local project material or nested repositories.
- `reports/` includes many generated files that are intentionally excluded from root Git tracking by `.gitignore`.
