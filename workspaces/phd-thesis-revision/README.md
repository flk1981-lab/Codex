# 博士论文修改环境

这个工作区用于管理博士论文的修改、检查、拆章和输出，默认适配 `WPS + docx + Python + Pandoc`。

## 目录

```text
phd-thesis-revision/
├─ README.md
├─ run.cmd
├─ run.ps1
├─ config/
├─ input/
├─ output/
└─ tools/
```

## 推荐放法

- `input/original/`
  放原始论文文件，尽量只读保存。
- `input/working/`
  放当前修改中的论文文件。
- `output/backups/`
  放自动备份。
- `output/outline/`
  放大纲和结构检查结果。
- `output/chapters/`
  放按一级标题拆出的章节文件。
- `output/reports/`
  放术语一致性、结构检查等报告。
- `config/`
  放术语表、缩写表、检查规则。

## 常用命令

优先推荐在本目录打开 `CMD` 后使用：

```cmd
run.cmd check
run.cmd outline .\input\working\thesis.docx
run.cmd stats .\input\working\thesis.docx
run.cmd split .\input\working\thesis.docx
run.cmd terms .\input\working\thesis.docx
run.cmd backup .\input\working\thesis.docx
```

如果你习惯 PowerShell，也可以用：

```powershell
.\run.ps1 check
.\run.ps1 outline .\input\working\thesis.docx
.\run.ps1 stats .\input\working\thesis.docx
.\run.ps1 split .\input\working\thesis.docx
.\run.ps1 terms .\input\working\thesis.docx
.\run.ps1 backup .\input\working\thesis.docx
```

如果 PowerShell 提示脚本被系统策略阻止，请改用 `run.cmd`，或者使用：

```powershell
powershell -ExecutionPolicy Bypass -File .\run.ps1 check
```

## 术语表

默认术语表文件为 `config/terms.csv`，格式如下：

```csv
canonical,variants,notes
结核分枝杆菌,结核杆菌|MTB,首次定义后全文统一
耐多药结核病,MDR-TB|MDR TB,中英文写法统一
```

`variants` 用 `|` 分隔多个候选写法。

## 建议工作流

1. 把原始论文放到 `input/original/`
2. 复制一份到 `input/working/` 开始改
3. 每次大改前运行 `.\run.ps1 backup`
4. 修改一轮后运行 `outline`、`stats`、`terms`
5. 需要按章节处理时运行 `split`

## 下一步

你把论文主文件放进 `input/working/` 后，我就可以继续帮你做：

- 全文结构体检
- 标题层级与图表编号检查
- 术语和缩写一致性检查
- 按章节分拆润色
- 批量生成修改清单
