# GitHub Upload Guide

本文件说明如果把 `~/Codex` 上传到 GitHub，默认应该提交什么、不应该提交什么。

## 建议提交

- `README.md`
- `docs/`
- `knowledge/`
- `scripts/`
- `archive/` 中明确需要留档的说明性材料

## 默认不要提交

- `.codex/`
- `workspaces/` 下的独立项目仓库
- `reports/` 中自动生成的 `.docx`、`.pdf`、`.pptx`
- `reports/**/pending/`
- 本地虚拟环境、缓存、临时文件和系统垃圾文件

## 原则

- `~/Codex` 顶层仓库负责“工作区说明、知识沉淀、可复用脚本”
- `workspaces/` 下每个项目自己负责自己的 Git 历史
- 如果某份报告需要长期共享，先确认内容适合公开，再考虑移动到一个可跟踪位置

## 上传前检查

```bash
git status
git add -A
git diff --cached
```

确认没有以下内容后，再提交：

- `.codex/`
- `.venv/`
- `.DS_Store`
- `reports/**/pending/`
- 生成型二进制报告
- 私有数据或密钥
