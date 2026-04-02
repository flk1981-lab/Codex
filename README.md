# Codex Workspace

这个目录是人工可维护的 Codex 工作总入口，不是 Codex 应用的内部状态目录。

## 目录职责

- `workspaces/`
  放独立项目或独立 Git 仓库。每个项目应能单独执行 `git pull`、`git commit`、`git push`。
- `knowledge/`
  放长期维护的主题知识库、方法笔记、阅读队列和研究备忘。
- `reports/`
  放阶段性输出、周期性报告和待处理结果。
- `scripts/`
  放跨项目复用的小脚本。
- `docs/`
  放这个工作区级别的说明文档。
- `archive/`
  放不再活跃但需要保留的历史材料。

## 与 `~/.codex` 的区别

- `~/.codex/` 是 Codex 应用内部状态目录，包含配置、缓存、会话、日志、数据库和密钥，不应按项目目录来整理。
- `~/Codex/` 是你自己维护的工作区，可以按主题、项目和输出结果进行结构化管理。

## 当前活跃入口

- 项目仓库：
  `workspaces/clinical-research-toolkit/`
- 主题知识库：
  `knowledge/tb-research/`
- 自动化与研究相关文档：
  `docs/`
- 跨项目脚本：
  `scripts/`

## 使用规则

- 新的长期项目优先放进 `workspaces/`
- 新的长期知识沉淀优先放进 `knowledge/`
- 导出结果和阶段性材料放进 `reports/`
- 过期但要保留的材料放进 `archive/`
- 不要把 `~/.codex/` 中的运行状态文件混进这里
- 如果把 `~/Codex` 本身上传到 GitHub，顶层仓库默认只跟踪文档、知识库和脚本；`workspaces/` 下的独立项目应各自单独管理 Git

详细结构规则见 `docs/STRUCTURE.md`。
