# Codex Structure

本文件说明 `~/Codex` 的目录组织规则，方便长期维护和跨电脑同步。

## 目标结构

```text
Codex/
├── README.md
├── docs/
├── workspaces/
├── knowledge/
├── reports/
├── scripts/
└── archive/
```

## 放置规则

- `workspaces/`
  用于独立项目。每个子目录最好都是一个可以独立协作的仓库。
- `knowledge/`
  用于长期主题知识库，例如方法学笔记、阅读队列、研究问题池、证据索引。
- `reports/`
  用于阶段性输出。建议按主题或年月组织。
- `scripts/`
  用于通用工具脚本，不要把只服务单个项目的脚本长期放这里。
- `docs/`
  用于工作区级别的说明、自动化文档和操作手册。
- `archive/`
  用于不再活跃但需要保留的内容，包括旧草稿、迁移中间态和历史说明。

## 当前结构说明

- `workspaces/clinical-research-toolkit/`
  当前主要项目仓库，已经单独初始化 Git。
- `knowledge/tb-research/`
  当前结核研究知识库。
- `reports/tb-research/`
  当前结核研究相关输出目录。
- `scripts/`
  当前存放邮件和报告相关辅助脚本。
- `.codex/`
  目前仍在 `~/Codex/` 下保留，但职责尚未完全明确；在确认依赖关系前，先不要移动或删除。

## 建议维护方式

- 打开 `~/Codex` 时，优先从 `workspaces/` 和 `knowledge/` 开始找内容
- 项目级规则写进项目自己的 `README.md`、`AGENTS.md`、`.codex/config.toml`
- 工作区级规则写进这里，不要和项目文档混在一起
- 如果以后把 `~/Codex` 上传到 GitHub，务必排除 `~/.codex/` 和各类本地运行状态
- 顶层 `~/Codex` 仓库建议不要直接纳入 `workspaces/` 下的独立项目内容；这些项目应各自单独推送
- `reports/` 下的生成型文档建议默认忽略，只有确认适合共享的材料再手动整理到可跟踪位置

## 迁移记录

- `clinical-research-toolkit/` 已移动到 `workspaces/clinical-research-toolkit/`
- `.codex.backup.20260402-002934` 已移动到 `~/Archive/codex/`
- `~/.codex/` 保持不动
