<div align="center">

# 公文 Skills · gongwen-skills

**懂规矩的公文写作助手：懂条例国标，懂单位差异，懂岗位角色；不编造，守保密，帮你做减法。**

22 个 skill，覆盖起草、改稿、核稿、国标排版，另有乡镇街道社区、医院、高校、中小学、国企五个行业包。

简体中文 | [English](README_EN.md)

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
![Skills](https://img.shields.io/badge/skills-22-brightgreen)
![Version](https://img.shields.io/badge/version-1.0.0-orange)
[![GitHub stars](https://img.shields.io/github/stars/duyunzhoudyz/gongwen-skills?style=flat&logo=github)](https://github.com/duyunzhoudyz/gongwen-skills/stargazers)
![Platforms](https://img.shields.io/badge/platforms-Claude%20Code%20%7C%20Codex%20%7C%20Cursor%20%7C%20Gemini%20CLI%20%7C%20Trae%20%7C%20Kimi%20%7C%20Qwen-lightgrey)
![Prompts](https://img.shields.io/badge/prompts-DeepSeek%20%7C%20Doubao%20%7C%20Kimi-lightgrey)

[效果](#效果) · [30 秒上手](#30-秒上手) · [技能地图](#技能地图) · [直接这样说](#直接这样说) · [完整流程](#一份材料的完整流程) · [常见问题](#常见问题) · [关注与交流](#关注与交流)

</div>

> **保密提示**：涉密信息严禁输入任何联网 AI 工具。处理内部材料请使用单位内网部署的模型；本项目的提示词版和离线包可以直接在内网模型上使用。详见 [使用须知](使用须知.md)。

---

## 为什么做这个

用 AI 写材料的人越来越多，翻车的也越来越多：

- **编造事实**：让 AI 配几个本地案例，结果张冠李戴、查无此企业；
- **一股 AI 味**：通篇“赋能”“闭环”，“不是……而是……”，破折号满天飞；
- **态度含糊**：请示里没有倾向性意见，通知里全是“建议”“可以考虑”，末尾还附一句“以上内容仅供参考”；
- **格式不对**：粘贴进 Word 满屏 `**` 和 `#`，还得自己按国标重新排版；
- **不懂单位**：医院的国考指标、高校的两级发文、国企的党委前置研究，通用 AI 一概不知。

这套 Skills 把老笔杆子的经验和一线摸索出的写作技巧，固化成可复用、可检验的工作流：

| | 一般的 AI 写作 | gongwen-skills |
|---|---|---|
| 数据和案例 | 缺了就编 | 只用你给的材料，缺的标【待补】 |
| 篇幅 | 越写越长，素材不够就注水 | 你定了字数就按字数写，素材不够用文种常规的工作性表述补足，不新增事实、不注水 |
| 文风 | AI 腔、新八股 | 去掉 AI 腔，保留“一是二是”、排比小标题这些公文正当写法 |
| 态度 | 对冲、免责声明 | 措辞明确，需要核实的事项放到正文之外 |
| 格式 | 自己排 | 一键生成 GB/T 9704 国标 docx，机检标点、数字、发文字号 |
| 单位差异 | 一刀切 | 7 类单位画像、5 种角色模式、5 个行业包 |

## 效果

### AI 稿体检：一篇通知的前后对比

模拟 AI 起草的区教育局冬季校园安全检查通知，依次经过 `gongwen-direct` → `gongwen-humanize` → `gongwen-review`（完整过程见 [examples/B-AI稿体检](examples/B-AI稿体检/)）：

| 机检 | P0 必改 | P1 建议改 | 篇幅 |
|---|---:|---:|---:|
| 原稿 | 5 | 31 | 1092 字 |
| 体检后 | 0 | 0 | 807 字 |

**原稿（节选）**

> 随着冬季的到来，校园安全形势日趋复杂严峻，做好冬季校园安全检查工作具有十分重要的意义。……
>
> \*\*一、充分认识冬季校园安全检查的重要意义\*\*
>
> 毋庸置疑，校园安全是教育事业发展的底线和生命线。冬季是校园安全事故的高发期——取暖用电增多、道路结冰、流感高发，各类风险交织叠加。值得注意的是，部分学校仍存在重检查轻整改的问题。安全检查不是走过场，而是守护师生生命安全的关键防线。各学校要以此次检查为抓手，打造全方位、全链条的校园安全管理体系，充分运用信息化手段赋能校园安全管理……
>
> 以上内容仅供参考，请各学校结合实际情况调整。

**体检后**

> 一、总体要求
>
> 冬季取暖用电增多、道路结冰、流感高发，是校园安全事故的高发期。部分学校仍存在重检查轻整改的问题。各学校要以此次检查为抓手，建设校园安全管理体系，运用信息化手段提升管理能力……

其他改动：
- “建议各学校成立……专班”改成“各学校要成立……专班”；
- 发文字号“[2026]第12号”改成“〔2026〕12号”，日期“11月05日”改成“11月5日”，“3～5万元”改成“3万～5万元”；
- 对家长“一律移交公安机关予以拘留”的表述，核查后发现区教育局没有处罚权，予以删除；
- 所有数字前后一致，脚本逐项比对确认。

另有两个示例：[医院科室年终总结](examples/A-医院科室年终总结/)（要求 2000 字左右而素材不多；示例按 v1.0 早期规则生成，发布前将按现行规则更新）和 [申请经费该用函还是请示](examples/C-文种判断-函/)。三个示例的说明见 [examples/README.md](examples/README.md)。

### 与同类项目对比

我们把本项目和 GitHub 上 8 个同类公文 skill 放在同一套题上真跑：同一个模型（claude-sonnet-5-5）、同样的用户消息和材料，Claude 与 DeepSeek 两家大模型盲评。下表取自两套**留出题**（出题后没有据此改过 skill），共 54 道题、552 次运行，材料全部虚构。

| 维度 | 本项目 | 同类项目中最好的 |
|---|---|---|
| 去 AI 味后事实保持 | **682/682 项全部保留，编造 0 次** | 98.4%，两家评审合计判出编造 15 次 |
| 排版：GB/T 9704 版式审计 | **96.6 分** | 93.2 分 |
| 排版：缺字体时如实处理 | **32 次全部写标准字体名，不替代** | 另一项目 32 次全部写入替代字体，12 次未告知 |
| 写作：出现编造的运行（Claude 评审） | **11/40** | 16/40 |
| 带密级材料按要求停止处理 | 首页标密级：**唯一识别**；仅页眉标密级：各方均漏检，本项目已修复并复测通过 | 均未识别 |
| 写作整体质量（平均名次，1 最好） | 3.20，第 2（Claude 评审） | 3.15 |
| 排版整体质量 | 2.72，第 3 | 2.12 |
| 去 AI 味整体质量 | 1.97，第 2（DeepSeek 评审）；2.19，第 3（Claude 评审） | 1.69 / 1.78 |

**防编造、事实保持、保密、版式合规、字体诚实这些“不能出错”的地方，本项目领先所有参赛项目；文稿整体质量处于第一梯队，与最好的项目差距在 0.1～0.6 个名次之间，还在继续改进。**

参赛项目：[lieflat-gongwen](https://github.com/larashero3-dotcom/lieflat-gongwen)、[official-document-skill](https://github.com/Liuxiangjian-ai/official-document-skill)、[official-document-writing-skill](https://github.com/KaguraNanaga/official-document-writing-skill)、[super-official-writer](https://github.com/Amer-CN/super-official-writer)、[gongwen-gbt9704-skill](https://github.com/mizzlelover/gongwen-gbt9704-skill)、[document-format-skills](https://github.com/KaguraNanaga/document-format-skills)、[gbt-9704-2012-skills](https://github.com/cj0103/gbt-9704-2012-skills)、[lieflat-less-ai-tone](https://github.com/larashero3-dotcom/lieflat-less-ai-tone)，以及“不装 skill”对照。本项目借鉴过其中多数项目的思路，在此致谢；对比只针对评测时的版本。各项目的提交号、分赛道明细、迭代过程和局限（单一执行模型、每题 2 次、题目由本项目设计、评审员为大模型等）见 [竞品对标评测报告](evals/results/compete-report.md)，评测脚本和题目在 [evals/compete](evals/compete/)，可以自行复现。

### 评测结果

v0.1 发布前，我们用真实模型对首批 7 个 skill 做了对比评测：同一个模型（claude-opus-5-5），同样的用户消息和材料，一组装 skill，一组不装，评分员盲评。共 104 条用例、621 项检查，材料全部虚构。

| 口径 | 装 skill | 不装 skill |
|---|---:|---:|
| 全部检查项通过率（621 项） | **94.8%** | 59.7% |
| 通用检查项通过率（422 项） | **95.0%** | 74.6% |
| 34 条防编造用例中出现编造的运行 | **0** | 15 |
| 7 条涉密、内部资料用例中按要求停止处理的 | **7** | 2 |

- “通用”检查项指熟悉公文写作的人只看用户消息和材料也应做到的要求，不含【起草说明】栏目、素材台账列名这类本项目自己的约定。不装 skill 的一组天然做不到这些约定，所以两个口径都列出，看实际差距以“通用”为准。
- 不装 skill 时，编造以凭空写出的工作事例、成效评价和廉洁表态为主，很少编统计数字。
- 每条用例只跑一次，评分员与被评模型是同一个模型。方法、各 skill 分项和局限见 [v0.1 评测报告](evals/results/v0.1-benchmark.md)。

**v1.0 全量评测（22 个 skill）**：被测模型换成 claude-sonnet-5-5，共 372 条用例、2671 项检查，材料全部虚构，评分员盲评。先评了修复前的版本（第一轮），按暴露的问题逐个 skill 修复后，再复测装 skill 一组（第二轮，即发布版本）。

| 口径 | 装 skill（第二轮） | 不装 skill |
|---|---:|---:|
| 全部检查项通过率（2671 项） | **98.6%** | 54.1% |
| 通用检查项通过率（1726 项） | **99.1%** | 68.7% |
| 136 条防编造用例中出现编造的运行 | **0** | 66 |
| 21 条带密级、内部资料标识的用例中按要求停止处理的 | **21** | 3 |

- 修复前的第一轮（两组条件完全一致）：全部检查项 97.1% 对 54.0%，通用检查项 97.7% 对 68.5%。全部用例里，装 skill 一组有编造的运行：第一轮 19 次，第二轮评分员标出 2 次，复核后确认是引用工具内置资料，不算编造。
- 触发评测：440 条说法选对 439 条（99.8%）。
- 17 个核心 skill 的通过率均不低于 97.4%，防编造 0 次，触发达标，标 Stable；5 个行业包数字也达标，按约定统一标 Beta，等一线测试者反馈再升级。
- 局限：每条用例只跑一次，评分员与被评模型是同一个模型；第二轮装 skill 一组用更高的推理强度重跑，第二轮的差距不是等条件对比；复测用的是同一批用例；装 skill 一组仍有 37 项检查未通过（分布在 16 个 skill）；没有做跨模型评测和人工盲评，也未经一线测试者实测。方法、各 skill 分项和局限见 [v1.0 评测报告](evals/results/v1.0-benchmark.md)。

v1.0 修订了首轮交付、篇幅等规则，评测用例随之调整，作为新的基线，与 v0.1 的数字不直接比较。

## 30 秒上手

先看你在用什么工具：

| 你在用 | 安装方式 |
|---|---|
| DeepSeek、豆包、Kimi 等网页版，或单位内网部署的模型 | 不用安装，复制提示词即可，见方式四 |
| Claude Code | 插件市场两行命令，见方式二 |
| Codex、Cursor、Gemini CLI、GitHub Copilot、Qwen Code、Kimi Code CLI、Trae、CodeBuddy | 一键脚本，自动找到本机装了哪些工具，见方式一；也可以用方式三 |
| 已经在用 Agent，不想敲命令 | 让 Agent 自己装，见方式五 |

### 方式一：一键安装脚本（Windows、macOS、Linux）

1. 下载本仓库。会用 git 的运行 `git clone https://github.com/duyunzhoudyz/gongwen-skills.git`；不会的，点本页上方绿色的 **Code** 按钮，选 **Download ZIP**，下载后解压。
2. 在仓库文件夹里运行一行命令：

```powershell
# Windows：在文件夹地址栏输入 powershell 回车，在弹出的窗口里运行
powershell -ExecutionPolicy Bypass -File .\install.ps1
```

```bash
# macOS、Linux：在终端里进入仓库文件夹后运行
bash install.sh
```

脚本会检查 Python（3.9 及以上），问你要不要安装排版依赖，然后把 17 个核心 skill 装进本机找得到的全部 Agent 工具。想先看看会装到哪里、不写入任何文件：Windows 在命令末尾加 `-WhatIf`，macOS、Linux 加 `--dry-run`。

只装到某一个工具：

```bash
python tools/install.py --list-targets     # 查看支持哪些工具、各装在哪个目录
python tools/install.py --target qwen      # 只装到 Qwen Code
```

### 方式二：Claude Code 插件市场

```text
/plugin marketplace add duyunzhoudyz/gongwen-skills
/plugin install gongwen@gongwen-skills
```

### 方式三：npx skills（Codex、Cursor、Gemini CLI、GitHub Copilot 等）

```bash
npx skills add duyunzhoudyz/gongwen-skills
```

### 方式四：网页大模型或内网模型（不用安装）

打开 [`dist/prompts/`](dist/prompts/)，复制提示词文件的全文，粘贴到对话开头，再说出你的需求。第一次用，从 `gongwen.md`（总入口）开始。

- 每个 skill 对应一份提示词，每份约 2.6 万～2.9 万字。
- 需要换到下一个 skill 时，当前对话会给你一张交接卡。新开一个对话，先粘贴下一份提示词，再粘贴交接卡和材料，接着往下做。
- 在单位内网使用，可以下载 [Releases](https://github.com/duyunzhoudyz/gongwen-skills/releases) 页面的离线包，或自己运行 `python tools/build_prompts.py --offline` 生成。

### 方式五：让 Agent 自己装

把这段话发给你的 Agent：

```text
请把 https://github.com/duyunzhoudyz/gongwen-skills 中 skills/ 目录下的 skill 安装到你的 skills 目录。其中 gongwen-grassroots、gongwen-hospital、gongwen-university、gongwen-k12、gongwen-soe 是行业包，只装我需要的 gongwen-hospital。装好后告诉我有哪些可用。
```

### 选装行业包

行业包装在核心 skill 之上，按需选装。没装行业包时，核心 skill 按单位画像照样能写本行业的材料；装上以后，本行业的专用文书（如医院的等级评审自评报告、国企的“三重一大”议题材料）交给行业包处理。

| 行业 | 包名 |
|---|---|
| 乡镇、街道、村（社区） | `grassroots` |
| 公立医院 | `hospital` |
| 高校 | `university` |
| 中小学 | `k12` |
| 国有企业 | `soe` |

以医院为例：

```text
# Claude Code
/plugin install gongwen-hospital@gongwen-skills
```

```bash
# 一键脚本（Windows）
powershell -ExecutionPolicy Bypass -File .\install.ps1 --pack hospital
# 一键脚本（macOS、Linux）
bash install.sh --pack hospital
# 安装器
python tools/install.py --target auto --pack hospital
```

可以多选，如 `--pack hospital --pack soe`；`--pack all` 装全部五个。网页版直接用 `dist/prompts/` 里对应的提示词，如 `gongwen-hospital.md`。

> **依赖**：排版和 Excel 图表需要 Python 3.9+ 以及 `python-docx`、`openpyxl`（`pip install python-docx openpyxl`，一键脚本会问你要不要装）。没有 Python 也能用，排版改为给出 WPS/Word 手动设置清单。
>
> **卸载**：一键脚本和安装器的命令后加 `--uninstall`，看过要删除的清单再加 `--yes`。只删除本项目装上的 skill。

## 技能地图

17 个核心 skill 装在一个插件里，5 个行业包按需选装。点 skill 名可以看它的说明和更多用法。

成熟度：**Stable** 表示真实模型评测达标（通过率不低于 80%、防编造 0 次、能被正确触发）；**Beta** 表示已用样例测试，边界情况可能还有问题；5 个行业包也跑过 v1.0 真实模型评测，数字达标，按约定等一线反馈和人工盲评后再升级。

**入口与配置**

| Skill | 能力 | 成熟度 |
|---|---|---|
| [`gongwen`](skills/gongwen/) 总入口 | 接住说不清的需求，判断单位、角色、文书类型和行文方向，规划路线；3 个问题建简版单位画像；提示写作日历；目标 skill 没装时给交接卡和安装办法 | Stable |
| [`gongwen-profile`](skills/gongwen-profile/) 单位画像与文风学习 | 完整建档，可读取本单位公文处理办法；学本单位历史稿的文风；对比领导改过的稿子归纳偏好；管理禁用词；给画像做体检 | Stable |

**起草**

| Skill | 能力 | 成熟度 |
|---|---|---|
| [`gongwen-draft`](skills/gongwen-draft/) 法定公文 | 通知（会议、转发、印发、部署）、请示、就一件事的报告、函、通报、批复、意见、签报、决定、通告、公告；先定文种、查行文规则再成稿 | Stable |
| [`gongwen-report`](skills/gongwen-report/) 综合文稿 | 总结、计划、工作要点、汇报材料、调研报告、经验材料、述职述责述廉、自主制定的实施方案、大会工作报告；下级材料汇总，一份台账出多版材料 | Stable |
| [`gongwen-speech`](skills/gongwen-speech/) 讲话类 | 领导讲话、主持词、致辞、表态发言、交流发言、代表发言；按时长定字数（约 200 字/分钟），写成口语稿 | Stable |
| [`gongwen-party`](skills/gongwen-party/) 党建与组织生活 | 对照检查材料、党建工作报告、党课讲稿、主题党日方案、谈心谈话提纲、研讨发言、心得体会、思想汇报、“三会一课”记录整理；没有本人事例只出提纲 | Stable |
| [`gongwen-brief`](skills/gongwen-brief/) 信息简报与宣传 | 政务信息、简报、新闻稿、发布词、公众号推文；长材料压成信息；评估一篇信息能不能被采用 | Stable |
| [`gongwen-minutes`](skills/gongwen-minutes/) 会议纪要与议题 | 会议记录、会议纪要、上会议题材料、督办清单；只整理不添加，分清同意、原则同意与暂缓 | Stable |
| [`gongwen-implement`](skills/gongwen-implement/) 上级文件转化 | 提取上级文件要点、任务分解表、实施方案和贯彻落实方案、落实情况报告、传达要点；检测照抄上级原文，提示层层加码 | Stable |
| [`gongwen-rectify`](skills/gongwen-rectify/) 整改材料 | 整改方案、整改台账（可转 Excel）、整改进展和情况报告、按检查清单写自查报告；报告与台账逐条对账 | Stable |

**改稿质检**

| Skill | 能力 | 成熟度 |
|---|---|---|
| [`gongwen-title`](skills/gongwen-title/) 标题与提纲 | 公文标题、主副标题、三套小标题方案、提纲，检查标题与正文是否对得上 | Stable |
| [`gongwen-polish`](skills/gongwen-polish/) 改稿润色 | 压缩、只用已有事实扩写、上行下行语气互转、大白话转公文、按领导批注改、多稿合并 | Stable |
| [`gongwen-humanize`](skills/gongwen-humanize/) 公文去 AI 味 | 改掉 AI 腔，保留四字格、“一是二是”等公文正当写法，事实一字不动 | Stable |
| [`gongwen-direct`](skills/gongwen-direct/) 去防御性写作 | 删免责声明和对冲语气，把“建议”“可以考虑”改成要求，让请示有倾向性意见、批复明确表态 | Stable |
| [`gongwen-review`](skills/gongwen-review/) 核稿审校 | 按“八关”核稿，出分级问题清单；机检标点、数字、格式和保密隐私；按本单位禁用词表比对 | Stable |

**交付工具**

| Skill | 能力 | 成熟度 |
|---|---|---|
| [`gongwen-format`](skills/gongwen-format/) 国标排版 | 按 GB/T 9704-2012 生成红头、信函、纪要、白头、材料、简报六种版式的 Word（简报无国家标准格式，按机关常见做法）；修复已有 docx；核稿意见写成 Word 批注；简单表格转 Word 表格 | Stable |
| [`gongwen-data`](skills/gongwen-data/) 数据与图表 | 报表体检、多表合并、同比环比和百分点计算（附算式）、数据写成能进公文的句子、本机生成黑白灰 Excel 图表 | Stable |

**行业包**（单独安装）

| Skill | 能力 | 成熟度 |
|---|---|---|
| [`gongwen-grassroots`](skills/gongwen-grassroots/) 乡镇街道社区 | 报上级的政务信息、迎检汇报与自查、居民通知和告知书、防汛防火应急通知、工作台账、村（居）务公开；查面向群众的处罚和强制表述 | Beta |
| [`gongwen-hospital`](skills/gongwen-hospital/) 公立医院 | 等级评审自评、绩效考核分析、医疗质量分析、不良事件根因分析（RCA）和 PDCA 案例、科室主任述职（DRG、CMI）、党委会和院长办公会议题、院感与安全检查整改 | Beta |
| [`gongwen-university`](skills/gongwen-university/) 高校 | 审核评估自评报告、学科评估材料、申报书形式审查、学院年度总结、辅导员谈心谈话提纲、校内签报和“学校发还是部门发”的判断 | Beta |
| [`gongwen-k12`](skills/gongwen-k12/) 中小学 | 学期和班主任计划总结、督导评估自评、家长会讲话、学生评语、致家长的一封信、教研和安全教育记录、课后服务方案 | Beta |
| [`gongwen-soe`](skills/gongwen-soe/) 国有企业 | “三重一大”议题材料、董事会议案与决议、总经理办公会纪要、经营分析、改革任务台账、ESG 和社会责任报告、合规管理报告 | Beta |

## 直接这样说

不用记 skill 名，用平时说话的方式提需求就行：

| 你说 | 由谁处理 |
|---|---|
| “我刚到办公室，领导让我写个东西，我也说不清是啥。” | `gongwen` 总入口 |
| “领导让我给县自然资源局写个请示，景区停车场要临时用一块地。”（会告诉你两家不相隶属，应当用函） | `gongwen-draft` |
| “下周要交科室年终总结，这是全年的工作记录，帮我写。” | `gongwen-report` |
| “局长周三在全市冬季供暖保障会上讲 20 分钟，材料发你。” | `gongwen-speech` |
| “下个月开民主生活会，组织反馈的意见和我今年做的几件事都在这，帮我理个发言提纲。” | `gongwen-party` |
| “昨天院长办公会的录音转成文字了，出个纪要，定了的事列个督办清单。” | `gongwen-minutes` |
| “这篇经验材料 3000 多字，帮我压成 800 字的信息报市里。” | `gongwen-brief` |
| “县委巡察组反馈了 9 个问题，原文发你，帮我们弄个整改方案和台账。” | `gongwen-rectify` |
| “这篇是 DeepSeek 写的，AI 味太重，帮我改改，数字别动。” | `gongwen-humanize` |
| “这个通知满篇‘建议各单位’，像在跟下级商量，改成要求的口气。” | `gongwen-direct` |
| “明天送签，帮我核一下，没问题就排成红头 Word。” | `gongwen-review` → `gongwen-format` |
| “社区下周二停水检修，写个贴在单元门口的通知，再弄一条发业主群的。” | `gongwen-grassroots` 行业包 |
| “国考结果下来了，四级手术和 CMI 都掉了，院长让写个分析。” | `gongwen-hospital` 行业包 |
| “学院想改造实验室，要给分管副校长打个签报。” | `gongwen-university` 行业包 |
| “期末了，帮我给班上 42 个孩子写评语，每个孩子的情况我列在表里了。” | `gongwen-k12` 行业包 |
| “集团要给子公司增资，下周上党委会，帮我弄个三重一大的上会材料。” | `gongwen-soe` 行业包 |

标题提纲、压缩改稿、上级文件转化、数据图表、单位画像等更多说法，见各 skill 目录下的 README。

## 一份材料的完整流程

| 场景 | 路线 |
|---|---|
| 年终总结、述职、汇报材料 | `gongwen` 确认需求 → `report` 建素材台账、出提纲、成稿 → `direct`、`humanize` 清理防御腔和 AI 腔 → `review` 核稿（停下来，等你处理核稿意见）→ `format` 排成国标 Word。要多套小标题的，成稿后加 `title` |
| AI 稿体检（手上已有一篇 AI 写的稿子） | `direct` → `humanize` → `review`，要出 Word 的再加 `format`。上面“效果”一节就是这条路线 |
| 上级文件转化（上级发文，要出本单位方案） | `implement` 提取要点、拆任务分解表、写实施方案并检测照抄 → `direct` → `review`（停）→ `format`。到了报送节点，`implement` 再按任务表写落实情况报告 |

表中省略了 `gongwen-` 前缀。在 Claude Code、Codex 这类能调用多个 skill 的工具里，总入口按路线逐站交接；在网页版里，每一站结束时给你一张交接卡，粘贴到下一份提示词后面接着做。

## 六条写作红线

所有 skill 都遵守这六条，优先级高于任何具体流程（全文见 [shared/guardrails.md](shared/guardrails.md)）：

1. **保密先于生成**：发现密级标识、“内部资料”“工作秘密”字样或个人敏感信息，先停下来提醒。
2. **事实只来自用户**：数据、案例、文号、政策依据都只取自你的材料；必须由你补的标【待补】，存疑的列在正文后的起草说明里；“下周三”按当天日期推算并注明，“去年”这类年份原样保留，成文日期留到签发日；加引号的本单位提法原样保留。
3. **正文干净，提示外置**：正文里没有免责声明和 AI 自述，需要核实的事项放到正文后的“起草说明”。
4. **默认做减法**：通知、请示、函等开门见山，遵守减负规定的篇幅要求，不照抄上级文件；你指定了字数的按字数写，素材撑不起就用工作性表述补足并在起草说明中说明，不新增事实、不注水。
5. **规则可溯源、可机检**：每条格式、标点、数字规则都标注条款号，能用脚本检查的就不靠模型自觉。
6. **辅助不代笔**：对照检查、述职等需要本人经历的材料，没有素材就只给提纲。

## 不同单位，不同写法

内置 7 类单位画像：党政机关、基层（县乡、街道）、社区、公立医院、高校、中小学、国有企业。每类都整理了发文体系、常用文种、决策程序、术语、红线和年度写作节点。行业包里经过核实、通用文书也用得上的要点（如医院的指标口径、国企的前置研究程序），已经回灌到这 7 份画像里，没装行业包时核心 skill 也按本行业的规矩写。

内置 5 种角色模式：

| 模式 | 适合谁 | 有什么不同 |
|---|---|---|
| 教学模式 | 新入职公务员、选调生 | 首轮照常出稿，附写作要点，讲清为什么选这个文种、为什么这样写 |
| 高效模式 | 专职笔杆子 | 少解释，直接出多个版本；问题类型和依据条款仍然每条一行写明 |
| 转译模式 | 兼职写材料的业务骨干 | 把大白话和数据表转成公文语言 |
| 审稿模式 | 中层领导、主要领导 | 只出问题清单和一页纸摘要 |
| 严谨模式 | 党务、宣传岗 | 提法核验最严，引用必须注明出处 |

单位画像分两级：

- **简版**：第一次使用时，总入口用 3 个问题建档，保存在 `.gongwen/单位画像.md`。
- **完整版**：`gongwen-profile` 建立单位画像、领导偏好卡、禁用词、文风参数四个文件。它可以读本单位的公文处理办法，从历史稿里学文风，对比领导改过的稿子归纳偏好；`gongwen-review` 核稿时按禁用词表逐条比对。

之后所有 skill 都按你的单位和角色调整写法。画像可能含内部信息，`.gongwen/` 已写入 `.gitignore`，不要提交到公开仓库。

## 规则依据

《党政机关公文处理工作条例》、GB/T 9704-2012《党政机关公文格式》、GB/T 15834-2011《标点符号用法》、GB/T 15835-2011《出版物上数字用法》、《整治形式主义为基层减负若干规定》。同比、环比、百分点等数据写法，另参照国家统计局的公开解释。要点整理见 [shared/rules/](shared/rules/)，每条都标注了条款号。

机检引擎按这些规则做 12 类检查，覆盖 41 类文书。

## 常见问题

**1. 这个项目是什么？收费吗？**

gongwen-skills 是给机关、事业单位、国企文字工作者用的公文写作 Skills 集合，v1.0.0 有 22 个 skill。按 MIT 许可开源，免费使用。

**2. 没有 Claude，能用吗？**

能。Skill 版支持 Claude Code、Codex、Cursor、Gemini CLI、GitHub Copilot、Qwen Code、Kimi Code CLI 等工具，Trae、CodeBuddy 的安装目录来自社区资料，装的时候会提示。没有这些工具的，用提示词版：打开 `dist/prompts/`，复制文件内容粘贴到 DeepSeek、豆包、Kimi 等网页大模型，或者单位内网部署的模型。

**3. 单位的材料能直接贴进去吗？**

涉密材料和标注“内部资料”“工作秘密”的材料，不能输入任何联网 AI 工具。处理内部材料请用单位内网部署的模型，提示词版和离线包可以在内网使用。个人信息先脱敏再输入。skill 发现密级标识或敏感信息时会先停下来提醒，但这只是最后一道提醒，判断责任在使用者。

**4. 它会不会编数据？**

规则要求只用你给的材料，必须由你补的标【待补】，存疑的列在起草说明里。v0.1 评测中，34 条故意少给材料的用例，装了 skill 后出现编造 0 次，不装的有 15 次。模型输出有随机性，定稿前请逐条核对数字、人名、文号。

**5. 去 AI 味，能让稿子查不出是 AI 写的吗？**

不能，也不承诺这种效果。去 AI 味改的是文风质量，比如段首评论语、新八股词、破折号过多，目的是让稿子更好读。

**6. 我在医院（高校、中小学、国企、乡镇社区），有专门的吗？**

有 5 个行业包：gongwen-hospital、gongwen-university、gongwen-k12、gongwen-soe、gongwen-grassroots，各自带本行业的专用文书模板、红线和核稿清单。本行业的通用文书，比如总结、通知、讲话，仍由核心 skill 处理。没装行业包时，核心 skill 按单位画像处理。

**7. 排出来的 Word 字体不对，或者电脑上没有 Python，怎么办？**

项目不随附字体，也不会用别的字体顶替。电脑上缺仿宋_GB2312、楷体_GB2312、方正小标宋简体等标准字体时，排版 skill 会在回复开头指出缺哪几款，并告诉你到哪里获取、怎么安装（方正小标宋简体可到方正字库官网搜索，单位使用需商业授权；仿宋_GB2312、楷体_GB2312 先向单位信息中心或办公室索取），装好后重启 Word、WPS 再打开核对。缺字体的电脑上看到的不是正式效果，不能据此核对版式或打印。没有 Python 时，排版 skill 会给出 WPS、Word 的手动设置清单。不同电脑上显示可能有细微差别，正式印发前请打印核对。

**8. 提法会不会过时？**

项目不内置“最新提法库”。政治性提法、上级精神只对照你提供的原文。核稿时会把疑似过时的提法和机构旧称标出来，请你对照原文确认，不会自行替换。

## 路线图

v0.2 到 v0.4 原计划分三次发布，开发提前完成，与 v0.1 合并为 v1.0.0 一次发布，用成熟度标签区分。各版本的变化见 [CHANGELOG.md](CHANGELOG.md)。

- [x] **v0.1「年终季版」**：总入口、综合文稿、标题、去 AI 味、去防御性、核稿、排版
- [x] **v0.2「开年版」**：单位画像与文风学习、法定公文、讲话稿、党建材料、改稿润色
- [x] **v0.3「落实版」**：上级文件转化、整改材料、会议纪要、信息简报、数据与图表
- [x] **v0.4「行业包」**：基层、医院、高校、中小学、国企
- [x] **v1.0**（2026 年 10 月 31 日）：22 个 skill 一次发布；真实模型评测；一键安装脚本和行业包选装

**后续**

- [ ] 跨模型评测：提示词版在 DeepSeek、Qwen、Kimi 上的表现
- [ ] 收集测试者反馈、做人工盲评，据此修订规则，行业包达标后从 Beta 升级
- [ ] 在 Word、WPS 中实测排版效果；在 macOS、Linux 上实测字体检测
- [ ] 输出 Word 修订痕迹（目前只写批注）
- [ ] 各 skill 自带检查脚本的规则并入机检引擎

想帮忙测试的同事，请看 [测试者手册](docs/评测/测试者手册.md)：用虚构材料试用，不需要懂编程。

## 关注与交流

- 抖音：（待填写）
- 微信：（待填写）
- 抖音交流群：（二维码待添加）
- 微信交流群：（二维码待添加）

规则写错了、遇到新场景、想投稿行业素材、发现 bug，请到 [Issues](https://github.com/duyunzhoudyz/gongwen-skills/issues/new/choose) 选对应的表单；使用心得和写作经验欢迎到 [Discussions](https://github.com/duyunzhoudyz/gongwen-skills/discussions) 交流。参与开发见 [CONTRIBUTING.md](CONTRIBUTING.md)。

## 致谢

本项目借鉴了以下开源项目的思路或内容。许可证和借鉴方式见 [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)，每个模块的逐条明细见 [docs/借鉴记录/](docs/借鉴记录/)。

| 方向 | 项目 |
|---|---|
| Skill 集合的组织方式 | [Yuan1z0825/nature-skills](https://github.com/Yuan1z0825/nature-skills)、[JimLiu/baoyu-skills](https://github.com/JimLiu/baoyu-skills) |
| 中文去 AI 味 | [op7418/Humanizer-zh](https://github.com/op7418/Humanizer-zh)、[larashero3-dotcom/lieflat-less-ai-tone](https://github.com/larashero3-dotcom/lieflat-less-ai-tone)、[LifelongLazyLearner/qu-ai-wei](https://github.com/LifelongLazyLearner/qu-ai-wei)、[MrGeDiao/shuorenhua](https://github.com/MrGeDiao/shuorenhua)、[VincentOld/stop-slop-zh](https://github.com/VincentOld/stop-slop-zh)、[blader/humanizer](https://github.com/blader/humanizer)、[MohamedAbdallah-14/unslop](https://github.com/MohamedAbdallah-14/unslop) |
| 去防御性表达 | [hexiecs/talk-normal](https://github.com/hexiecs/talk-normal)、[molang163/chinese-natural-reply](https://github.com/molang163/chinese-natural-reply) |
| 公文文种与写作规则 | [zhaohui-yang/official-document-drafting](https://github.com/zhaohui-yang/official-document-drafting)、[gongyu0918-debug/chinese-official-writing-skill](https://github.com/gongyu0918-debug/chinese-official-writing-skill)、[Liuxiangjian-ai/official-document-skill](https://github.com/Liuxiangjian-ai/official-document-skill)、[KaguraNanaga/official-document-writing-skill](https://github.com/KaguraNanaga/official-document-writing-skill) |
| 公文排版与检查 | [mizzlelover/gongwen-gbt9704-skill](https://github.com/mizzlelover/gongwen-gbt9704-skill)、[KaguraNanaga/document-format-skills](https://github.com/KaguraNanaga/document-format-skills)、[imhuhe/gongwen-kit](https://github.com/imhuhe/gongwen-kit)、[linhut/gongwen-skill](https://github.com/linhut/gongwen-skill)、[huacnlee/autocorrect](https://github.com/huacnlee/autocorrect) |
| 数据、表格与重合检测 | [alchaincyf/huashu-excel](https://github.com/alchaincyf/huashu-excel)、[hustcc/mcp-echarts](https://github.com/hustcc/mcp-echarts)、[bowenliang123/markdown-exporter](https://github.com/bowenliang123/markdown-exporter)、[blingenf/copydetect](https://github.com/blingenf/copydetect) |
| 行业包与评测写法 | [PIGU-PPPgu/class-teacher-skills](https://github.com/PIGU-PPPgu/class-teacher-skills)、[anthropics/k12-teacher-skills](https://github.com/anthropics/k12-teacher-skills)、[huangwb8/ChineseResearchLaTeX](https://github.com/huangwb8/ChineseResearchLaTeX)、[a5c-ai/babysitter](https://github.com/a5c-ai/babysitter) |
| 只参考公开思路，未使用文字和代码 | [larashero3-dotcom/lieflat-gongwen](https://github.com/larashero3-dotcom/lieflat-gongwen)、[Rimagination/gongwen-draft](https://github.com/Rimagination/gongwen-draft)、[jzOcb/writing-style-skill](https://github.com/jzOcb/writing-style-skill)、[alvis961/sop-writing](https://github.com/alvis961/sop-writing)、[cuberfry68-coder/wenxi-skills](https://github.com/cuberfry68-coder/wenxi-skills) |

### 贡献者

欢迎提交规则纠错、行业素材和评测用例，合并后在此署名。署名前会征得本人同意，也可以匿名。

## 许可证

[MIT](LICENSE)。规则要点整理自公开的法规和国家标准，不替代原文。
