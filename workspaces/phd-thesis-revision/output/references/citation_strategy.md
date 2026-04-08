# 博士论文参考文献重整策略（第 1 轮）

## 1. 管理工具与工作流

- 文献管理主工具：Zotero（本机已运行，已完成批量导入）。
- 本地结构化库：
  - `D:\Codex\workspaces\phd-thesis-revision\output\references\reference_library.json`
  - `D:\Codex\workspaces\phd-thesis-revision\output\references\reference_library.bib`
  - `D:\Codex\workspaces\phd-thesis-revision\output\references\review_collection.bib`
  - `D:\Codex\workspaces\phd-thesis-revision\output\references\body_collection.bib`
- 当前候选库规模：182 篇。
  - 综述线：99 篇
  - 论文主体线：87 篇

## 2. 分流策略

### 论文主体

- 处理原则：优先补“研究背景、讨论、结论”中的论断性句子；每处只补 1-3 篇，不堆砌。
- 编号体系：沿用主体原有独立编号，新增文献续接到 `[102]` 之后。
- 当前已补主体新增文献：
  - `[102]` SEAL-MDR 研究方案论文
  - `[103]` WHO 2025 module 4 guideline
  - `[104]` WHO 2025 module 4 operational handbook

### 文献综述

- 处理原则：优先补“概念总述、机制判断、精准化路径、终点框架、总结性判断”中的关键论断。
- 编号体系：沿用综述原有独立编号，新增文献续接到 `[28]` 之后。
- 当前已补综述新增文献：
  - `[28]` post-TB lung impairment systematic review
  - `[29]` Advancing host-directed therapy for tuberculosis
  - `[30]` immunogenetics of TB susceptibility

## 3. 已采用的检索主题/关键词

### 综述线关键词

- `tuberculosis AND (multidrug-resistant OR MDR-TB OR RR-TB OR pre-XDR) AND review`
- `tuberculosis AND "host-directed therapy"`
- `(sulfasalazine OR SLC7A11 OR xCT) AND (macrophage OR oxidative stress OR inflammation OR tuberculosis)`
- `tuberculosis AND macrophage AND (oxidative stress OR glutathione OR ROS OR inflammation)`
- `tuberculosis AND (host genetics OR polymorphism OR SNP OR pharmacogenetic*)`

### 主体线关键词

- `(multidrug-resistant tuberculosis OR pre-XDR OR drug-resistant tuberculosis) AND (short regimen OR all-oral OR bedaquiline OR linezolid) AND (trial OR cohort OR outcome*)`
- `tuberculosis AND linezolid AND (neuropathy OR optic neuropathy OR toxicity OR adverse event*)`
- `tuberculosis AND bedaquiline AND (QT OR safety OR electrocardiogram)`
- `(sulfasalazine OR SASP) AND (tuberculosis OR macrophage OR host-directed)`
- `(SLC7A11 OR xCT) AND (polymorphism OR SNP OR genotype OR genetics)`

## 4. 第 1 轮已落地的正文补引位置

### 主体

- 第二章研究平台总述：将原有缺号 `[103-105]` 校正为 `[102-104]`
- 第三章安全性讨论：补入 `QTcF` 与 `Bdq` 心电安全性证据 `[7,21,84]`
- 第五章机制讨论：补入 `xCT/SLC7A11` 机制支持文献 `[30,56,59]`
- 第七章结论边界：补入核心临床来源文献 `[17,22,60]`

### 综述

- 综述引言总述：补入全球负担/HDT/PTLD 证据 `[1,28,29]`
- xCT 节点判断：补入功能性节点证据 `[7,8]`
- SASP 双通路整合：补入 `xCT`/炎症/HDT 框架证据 `[9,11,29]`
- rs13120371 机制解释：补入遗传学与功能解释证据 `[8,30]`
- 精准路径框架：补入监测与终点评估证据 `[25-27]`
- 终点框架讨论：补入病灶恢复/长期损伤相关证据 `[4,26,28]`
- 综述总结：补入临床转化与精准化综述证据 `[12,29,30]`

## 5. 下一轮优先补引区域

- 主体第四章讨论中关于 36 个月持续性 `Lzd` 神经/视觉毒性的两段解释
- 主体第五章讨论中“功能支持”与“机制终证”边界的补强
- 综述中关于动态标志物、药物暴露桥接、患者中心终点的 3-5 个关键段落
- 全文检查是否还存在“正文引用号 > 对应参考文献表最大编号”的情况

## 6. 第 2 轮新增处理

### 主体

- 新增 `[105]-[107]`，用于补强第四章关于 `Lzd` 长期毒性、严重不良事件负担和长期功能结局的论证。
- 在 `4.4.2 利奈唑胺长期毒性的临床启示` 下新增 2 段说明性讨论：
  - 一段强调神经/视觉毒性可能持续存在
  - 一段强调方案优化不能只看近期培养学结局，还要兼顾长期功能结局

### 综述

- 在精准路径部分进一步补入已有综述内编号文献：
  - 多维标志物/监测框架 `[25,26]`
  - 动态标志物服务治疗决策 `[25,26]`
  - 患者中心终点 `[28]`

## 7. 第 3 轮新增处理

### 主体讨论段加固

- 第五章讨论补强了 5 个关键位置的引文：
  - 实验结果总述
  - 与临床发现一致性
  - 机制终证边界
  - 局限性段
  - 方法学与转化价值段

### 第七章总讨论加固

- 在 `7.1.2`、`7.4.2`、`7.5.1` 后新增了 5 段口径更稳的总结性文字。
- 处理目标不是重复结果，而是把以下三点写得更清楚：
  - 第五章是“功能支持”，不是“机制终证”
  - 第七章结论必须与第五章证据边界一致
  - 长期毒性与长期功能结局应被纳入总讨论，而不是只当作附属安全性信息

## 8. 第 4 轮新增处理

### 第一章前言收口补引

- 对第一章中 5 个仍偏概括的关键论断段补入更贴切的 1-3 篇文献：
  - HDT 在耐药结核中的潜在价值
  - HDT 研究谱系与局限
  - SASP 的临床前与临床转化基础
  - 宿主遗传异质性与分层应用

### 第七章未来展望收口

- 在 `7.5.2 未来展望` 下新增 2 段可直接用于答辩口径的展望性总结。
- 重点统一三点：
  - 未来研究要做“证据链桥接”，不是重复现象学比较
  - 长期毒性与长期功能结局应作为核心终点
  - SASP 的价值取决于疗效、机制、分层和长期结局四条线是否同时成立

## 9. 参考文献体例清理

- 统一修正主体参考文献 `[99]-[101]` 中缺失的卷期/页码空格。
- 删除综述参考文献列表中新增文献前遗留的空白段，保证列表连续。

