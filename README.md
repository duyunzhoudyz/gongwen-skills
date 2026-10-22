<!-- README_SYNC: source=working-tree; updated=2026-10-22 -->

<div align="center">

# 公文 Skills · gongwen-skills

**面向机关事业单位和国企文字工作者的公文写作 Skills：防编造、守保密、一键 GB/T 9704 排版。**

简体中文 · [English](README_EN.md)

如果这个项目帮了你，点一个 Star，方便回来，也让更多有同样需求的人更容易发现它。

[![GitHub stars](https://img.shields.io/github/stars/duyunzhoudyz/gongwen-skills?style=flat&logo=github&label=Stars)](https://github.com/duyunzhoudyz/gongwen-skills/stargazers)
[![License: PolyForm NC](https://img.shields.io/badge/License-PolyForm%20NC%201.0-blue.svg)](LICENSE)
![Skills](https://img.shields.io/badge/skills-22-brightgreen)
[![Version](https://img.shields.io/badge/version-1.0.0-orange)](CHANGELOG.md)
[![Platforms](https://img.shields.io/badge/platforms-7%20Agent%20tools%20%7C%20web%20LLM-lightgrey)](#快速开始)

[快速开始](#快速开始) · [技能地图](#技能地图) · [直接这样说](#直接这样说) · [效果](#效果) · [常见问题](#常见问题) · [关注与交流](#关注与交流)

</div>

> **保密提示**：涉密信息严禁输入任何联网 AI 工具。处理内部材料请使用单位内网部署的模型；本项目的提示词版和离线包可以直接在内网使用。详见 [使用须知](使用须知.md)。

<div align="center">

<img src="docs/assets/format-sample-notice.png" width="300" alt="gongwen-format 生成的红头通知样稿第 1 页" />

*`gongwen-format` 按 GB/T 9704-2012 生成的红头通知样稿第 1 页（测试样稿经 LibreOffice 渲染，内容虚构；实际字体以本机安装为准）*

</div>

---

## 和普通 AI 写材料有什么不同

| | 一般的 AI 写作 | gongwen-skills |
|---|---|---|
| 数据和案例 | 缺了就编 | 只用你给的材料，缺的标【待补】 |
| 篇幅 | 越写越长，素材不够就注水 | 你定了字数就按字数写，不新增事实、不注水 |
| 文风 | AI 腔、新八股 | 去掉 AI 腔，保留"一是二是"、排比小标题这些公文正当写法 |
| 态度 | 对冲、免责声明 | 措辞明确，需要核实的事项放到正文之外 |
| 格式 | 自己排 | 一键生成 GB/T 9704 国标 docx，机检标点、数字、发文字号 |
| 单位差异 | 一刀切 | 7 类单位画像、5 种角色模式、5 个行业包 |

---

**同台评测摘要**：和 GitHub 上 8 个同类公文 skill 在同一模型、同一题目下真跑，由 Claude、DeepSeek 两个 LLM 评审双盲评分（不是人工评审），数据取自两套留出题。

- **防编造、事实保持**：去 AI 味过程 682 项事实全部保留，两家评审判出编造 0 次；写作赛道出现编造的运行最少（11/40）
- **保密、版式、字体**：首页标密级的材料能停下来拒绝处理（涉密题只有 2 道，样本小）；GB/T 9704 版式审计 96.6 分；缺字体时如实写标准字体名并告知，不替代
- **整体写作质量**：第一梯队，尚未稳定第一。平均名次 Claude 评审第 2、DeepSeek 评审第 6，与最强对比项目差 0.1～0.6 个名次

各项数据、对比对象、局限和复现脚本见下方 [效果](#效果) 与 [竞品对标评测报告](evals/results/compete-report.md)。

---

## 快速开始

先看你在用什么工具：

| 你在用 | 安装方式 |
|---|---|
| DeepSeek、豆包、Kimi 等网页版，或单位内网模型 | 不用安装，复制提示词即可，见[方式四](#方式四网页大模型或内网模型不用安装) |
| Claude Code | 插件市场两行命令，见[方式二](#方式二claude-code-插件市场) |
| Codex、Cursor、Gemini CLI、GitHub Copilot、Qwen Code、Kimi Code CLI | 一键脚本，见[方式一](#方式一一键安装脚本windowsmacoslinux) |
| 已经在用 Agent，不想敲命令 | 见[方式五](#方式五让-agent-自己装) |

Trae、CodeBuddy 的 skills 目录路径尚未完全核实，一键脚本提供了对应目标（`--target trae`、`--target codebuddy`），可以试用，遇到问题请提 Issue。

### 方式一：一键安装脚本（Windows、macOS、Linux）

1. 下载本仓库。会用 git 的运行 `git clone https://github.com/duyunzhoudyz/gongwen-skills.git`；不会的，点本页上方绿色 **Code** 按钮，选 **Download ZIP**，下载后解压。
2. 在仓库文件夹里运行（Windows 在文件夹地址栏输入 `powershell` 回车；macOS、Linux 在终端里进入该文件夹）：

```powershell
powershell -ExecutionPolicy Bypass -File .\install.ps1
```

```bash
bash install.sh
```

脚本检查 Python（3.9+），询问是否安装排版依赖，然后把 17 个核心 skill 装进本机找到的全部 Agent 工具。只想看会装到哪里而不写入文件：Windows 末尾加 `-WhatIf`，macOS/Linux 加 `--dry-run`。

查看支持哪些工具，或只装到某一个工具（以 Qwen Code 为例）：

```bash
python tools/install.py --list-targets
python tools/install.py --target qwen
```

### 方式二：Claude Code 插件市场

```text
/plugin marketplace add duyunzhoudyz/gongwen-skills
/plugin install gongwen@gongwen-skills
```

### 方式三：npx skills（需要 Node.js）

```bash
npx skills add duyunzhoudyz/gongwen-skills
```

### 方式四：网页大模型或内网模型（不用安装）

打开 [`dist/prompts/`](dist/prompts/)，复制提示词文件全文，粘贴到对话开头，再说出你的需求。第一次用，从 `gongwen.md`（总入口）开始。

- 每个 skill 对应一份提示词，每份约 2.6 万～2.9 万字。
- 需要换到下一个 skill 时，当前对话会给你一张交接卡。新开对话，先粘贴下一份提示词，再粘贴交接卡和材料。
- 单位内网使用，运行 `python tools/build_prompts.py --offline` 生成离线包（输出到 `dist/offline/`）。

### 方式五：让 Agent 自己装

把这段话发给你的 Agent：

```text
请把 https://github.com/duyunzhoudyz/gongwen-skills 中 skills/ 目录下的 skill 安装到你的 skills 目录。
其中 gongwen-grassroots、gongwen-hospital、gongwen-university、gongwen-k12、gongwen-soe 是行业包，
只装我需要的 gongwen-hospital。装好后告诉我有哪些可用。
```

### 选装行业包

行业包装在核心 skill 之上，按需选装。以医院为例：

```text
/plugin install gongwen-hospital@gongwen-skills
```

```bash
bash install.sh --pack hospital
```

一键脚本支持多选：`--pack hospital --pack soe`；装全部：`--pack all`。

网页版直接用 `dist/prompts/` 里对应的提示词，如 `gongwen-hospital.md`。

> **依赖**：排版和 Excel 图表需要 Python 3.9+ 及 `python-docx`、`openpyxl`（`pip install python-docx openpyxl`，一键脚本会询问是否安装）。没有 Python 时，排版 skill 给出 WPS/Word 手动设置清单。
>
> **卸载**：安装命令末尾加 `--uninstall`，确认清单后再加 `--yes`。

---


## 技能地图

17 个核心 skill 装在一个插件里，5 个行业包按需选装。

成熟度：**Stable** = 真实模型评测通过率 ≥ 80%、防编造 0 次、触发正确；**Beta** = 样例测试通过，一线反馈后升级。

**入口与配置**

| Skill | 能力摘要 | 成熟度 |
|---|---|---|
| [`gongwen`](skills/gongwen/) 总入口 | 接住说不清的需求，判断文书类型和行文方向，3 个问题建简版单位画像，规划路线 | Stable |
| [`gongwen-profile`](skills/gongwen-profile/) 单位画像 | 完整建档；学本单位历史稿文风；归纳领导改稿偏好；管理禁用词 | Stable |

**起草**

| Skill | 能力摘要 | 成熟度 |
|---|---|---|
| [`gongwen-draft`](skills/gongwen-draft/) 法定公文 | 通知、请示、报告、函、通报、批复、意见、签报、决定、通告、公告；先定文种再成稿 | Stable |
| [`gongwen-report`](skills/gongwen-report/) 综合文稿 | 总结、计划、工作要点、汇报、调研报告、经验材料、述职述责述廉、实施方案、大会报告 | Stable |
| [`gongwen-speech`](skills/gongwen-speech/) 讲话类 | 领导讲话、主持词、致辞、表态/交流/代表发言；按时长定字数，写成口语稿 | Stable |
| [`gongwen-party`](skills/gongwen-party/) 党建与组织生活 | 对照检查、党建报告、党课、研讨发言、心得汇报、"三会一课"记录整理；无本人事例只给提纲 | Stable |
| [`gongwen-brief`](skills/gongwen-brief/) 信息简报 | 政务信息、简报、新闻稿、公众号推文；长材料压成信息；评估能否被采用 | Stable |
| [`gongwen-minutes`](skills/gongwen-minutes/) 会议纪要 | 会议记录、纪要、上会议题材料、督办清单；只整理不添加，分清同意/原则同意/暂缓 | Stable |
| [`gongwen-implement`](skills/gongwen-implement/) 上级文件转化 | 提取要点、任务分解表、实施方案、落实情况报告；检测照抄原文，提示层层加码 | Stable |
| [`gongwen-rectify`](skills/gongwen-rectify/) 整改材料 | 整改方案、台账（可转 Excel）、整改进展报告、自查报告；报告与台账逐条对账 | Stable |

**改稿质检**

| Skill | 能力摘要 | 成熟度 |
|---|---|---|
| [`gongwen-title`](skills/gongwen-title/) 标题与提纲 | 公文标题、主副标题、三套小标题方案、提纲，检查标题与正文是否对应 | Stable |
| [`gongwen-polish`](skills/gongwen-polish/) 改稿润色 | 压缩、扩写、语气互转、大白话转公文、按领导批注改、多稿合并 | Stable |
| [`gongwen-humanize`](skills/gongwen-humanize/) 去 AI 味 | 改掉 AI 腔，保留公文正当写法，事实一字不动 | Stable |
| [`gongwen-direct`](skills/gongwen-direct/) 去防御性 | 删免责声明，把"建议""可以考虑"改成要求，让请示有倾向性意见 | Stable |
| [`gongwen-review`](skills/gongwen-review/) 核稿审校 | 按"八关"核稿，出分级问题清单；机检标点、数字、格式、保密隐私、禁用词 | Stable |

**交付工具**

| Skill | 能力摘要 | 成熟度 |
|---|---|---|
| [`gongwen-format`](skills/gongwen-format/) 国标排版 | 按 GB/T 9704-2012 生成红头、信函、纪要、白头、材料、简报六种版式 Word；修复已有 docx | Stable |
| [`gongwen-data`](skills/gongwen-data/) 数据与图表 | 报表体检、同比环比计算（附算式）、数据写成公文句子、本机生成黑白灰 Excel 图表 | Stable |

**行业包**（单独安装）

| Skill | 能力摘要 | 成熟度 |
|---|---|---|
| [`gongwen-grassroots`](skills/gongwen-grassroots/) 乡镇街道社区 | 迎检汇报、居民通知、应急通知、工作台账、村居务公开 | Beta |
| [`gongwen-hospital`](skills/gongwen-hospital/) 公立医院 | 等级评审自评、绩效分析、RCA/PDCA、科室述职（DRG/CMI）、党委会/院长办公会议题 | Beta |
| [`gongwen-university`](skills/gongwen-university/) 高校 | 审核评估自评、学科评估、申报书审查、学院总结、签报和两级发文判断 | Beta |
| [`gongwen-k12`](skills/gongwen-k12/) 中小学 | 学期计划总结、督导评估自评、家长会讲话、学生评语、致家长一封信 | Beta |
| [`gongwen-soe`](skills/gongwen-soe/) 国有企业 | "三重一大"议题、董事会议案、总经理办公会纪要、ESG 报告、合规管理报告 | Beta |

---

## 直接这样说

不用记 skill 名，把你的需求直接发给它：

| 你发给它 | 由谁处理 |
|---|---|
| "帮我判断一下这个材料应该用什么文种，然后起草。（附背景材料）" | `gongwen` 总入口 |
| "这是今天上午安全生产专题会的录音转文字，帮我出纪要，定了的事列个督办清单。" | `gongwen-minutes` |
| "写一篇科室年度工作总结，2000 字左右，素材在附件里，不要编数据。" | `gongwen-report` |
| "这是 AI 写的初稿，把'赋能''闭环''全链条'这类 AI 腔去掉，改成正常公文的语气，数字和案例一个不动。（附稿件）" | `gongwen-humanize` |
| "这篇通知里的'建议各单位''可以参考'改成要求的口气，明确是必须执行。（附通知）" | `gongwen-direct` |
| "明天这份请示要送签，先帮我核一遍有没有问题，核完排成标准红头 Word。（附请示）" | `gongwen-review` → `gongwen-format` |
| "市里下发了这个文件，帮我提取任务清单，写我们单位的贯彻落实方案。（附文件）" | `gongwen-implement` |
| "巡察组反馈了 9 个问题，帮我出整改方案和整改台账，台账最后要能导出 Excel。（附反馈清单）" | `gongwen-rectify` |
| "下周等级评审，帮我写科室自评报告，重点体现 DRG 和 CMI 指标，材料在这里。（附数据）" | `gongwen-hospital` 行业包 |
| "学校审核评估，帮我写学院自评报告里'教学资源与利用'这一条，填写要求在附件。" | `gongwen-university` 行业包 |
| "子公司增资需要上三重一大，下周党委会，帮我准备上会材料，议题说明和决议草案都要。" | `gongwen-soe` 行业包 |
| "社区下周二停水检修，帮我写一份贴在单元门口的通知，另外写一条简短的发业主群消息。" | `gongwen-grassroots` 行业包 |

更多用法见各 skill 目录下的 README。

---

## 一份材料的完整流程

| 场景 | 路线 |
|---|---|
| 年终总结、述职、汇报材料 | `gongwen` 确认需求 → `report` 建素材台账出稿 → `direct`、`humanize` 清理腔调 → `review` 核稿 → `format` 排国标 Word |
| AI 稿体检（手上已有 AI 稿） | `direct` → `humanize` → `review`，要 Word 的再加 `format` |
| 上级文件转化 | `implement` 提取要点、拆任务表、写实施方案 → `direct` → `review` → `format` |

表中省略了 `gongwen-` 前缀。在能调用多个 skill 的工具里，总入口按路线逐站交接；在网页版里，每站结束给交接卡，粘贴到下一份提示词后继续。

---

## 效果

### AI 稿体检：一篇通知的前后对比

模拟 AI 起草的区教育局冬季校园安全检查通知，依次经过 `gongwen-direct` → `gongwen-humanize` → `gongwen-review`（完整过程见 [examples/B-AI稿体检](examples/B-AI稿体检/)）：

| 机检 | P0 必改 | P1 建议改 | 篇幅 |
|---|---:|---:|---:|
| 原稿 | 5 | 31 | 1092 字 |
| 体检后 | 0 | 0 | 814 字 |

**原稿（节选）**

> 随着冬季的到来，校园安全形势日趋复杂严峻……
>
> **一、充分认识冬季校园安全检查的重要意义**
>
> 毋庸置疑，校园安全是教育事业发展的底线和生命线。……各学校要以此次检查为抓手，打造全方位、全链条的校园安全管理体系，充分运用信息化手段赋能校园安全管理……
>
> 以上内容仅供参考，请各学校结合实际情况调整。

**体检后**

> 一、总体要求
>
> 冬季取暖用电增多、道路结冰、流感高发，是校园安全事故的高发期。部分学校仍存在重检查轻整改的问题。各学校要以此次检查为抓手，建设校园安全管理体系，运用信息化手段提升校园安全管理水平……

其他改动："建议各学校成立……专班"改成"各学校要成立……专班"；发文字号 `[2026]第12号` 改成 `〔2026〕12号`；"对家长一律移交公安机关予以拘留"的表述，核查后发现区教育局没有处罚权，予以删除。

另有两个示例：[医院科室年终总结](examples/A-医院科室年终总结/)（要求 2000 字而素材不多：正文约 1890 字，只有 2 处【待补】，没有新增事实）和[申请经费该用函还是请示](examples/C-文种判断-函/)。

### 与同类项目对比

本项目和 GitHub 上 8 个同类公文 skill 同台真跑：同一模型（claude-sonnet-5-5）、同样的用户消息和材料。结论取自两套**留出题**（c4、c5，共 54 道题、552 次运行，材料全部虚构）；c2、c3 在同一套题上反复调试过，不作为结论依据。

| 维度 | 本项目 | 对比项目（该项的最好成绩） |
|---|---|---|
| 去 AI 味：事实保持（682 项） | **100%，编造 0 次** | 98.4%，编造 15 次；另一项目也是 100%，但改动比例仅 0.6%，11 次未交付正文 |
| 排版：GB/T 9704 版式审计 | **96.6 分** | 93.2 分 |
| 排版：字体诚实 | **32 次全部写标准字体名** | 32 次全部写替代字体，12 次未告知 |
| 写作：出现编造的运行（Claude 评审） | **11/40 次**（最少） | 16/40、17/40 |
| 涉密材料识别 | c4 的 2 道题 2/2（唯一识别）；c5 密级只在页眉，全部参赛方漏检 | c4 为 0/2 |
| 写作整体质量（平均名次） | Claude 评审第 2（3.20）；DeepSeek 评审第 6（3.88） | Claude 评审第 1 名 3.15；DeepSeek 评审第 1 名是另一个项目 |
| 排版整体质量（平均名次） | 第 3（2.72） | 第 1 名 2.12 |
| 去 AI 味整体质量（平均名次） | Claude 评审第 3（2.19）；DeepSeek 评审第 2（1.97） | 第 1 名 1.78 / 1.69 |

防编造、事实保持、版式合规、字体诚实这些"不能出错"的维度，本项目在参赛方中成绩最好；整体质量是第一梯队，与最强对比项目差 0.1～0.6 个名次，尚未稳定第一。

几点限定：

- 评审员是 Claude 与 DeepSeek 两个 LLM，不是人工；其中 Claude 评审与被测模型同为 claude-sonnet-5-5。人工盲评在[路线图](#路线图)中。
- 涉密题只有 2 道，样本很小。c5 的页眉密级本项目当时也漏检，看到结果后才修复，修复后只重跑了本项目并复测通过。
- 评测只针对当时的版本，各项目此后可能已有改进。

完整方法、各 skill 分项、局限和可复现脚本见 [竞品对标评测报告](evals/results/compete-report.md)。

### v1.0 全量评测

被测模型 claude-sonnet-5-5，共 372 条用例、2671 项检查，材料全部虚构，评分员盲评。

| 口径 | 装 skill | 不装 skill |
|---|---:|---:|
| 全部检查项通过率（2671 项） | **98.6%** | 54.1% |
| 通用检查项通过率（1726 项） | **99.1%** | 68.7% |
| 136 条防编造用例中出现编造的运行 | **0** | 66 |
| 21 条带密级用例中按要求停止处理的 | **21** | 3 |

> 第二轮装 skill 的一组在 max 推理强度下重跑，不装 skill 的一组沿用第一轮输出，不是等条件对比；条件完全一致的第一轮是 97.1% 对 54.0%（通用检查项 97.7% 对 68.5%）。每条用例只跑一次，评分员与被评模型是同一个模型。

17 个核心 skill 通过率均不低于 97.4%，防编造 0 次，触发达标，标 Stable；5 个行业包数字也达标，按约定等一线反馈再升级。完整方法和局限见 [v1.0 评测报告](evals/results/v1.0-benchmark.md)。

---

## 六条写作红线

所有 skill 都遵守这六条，优先级高于任何具体流程（全文见 [shared/guardrails.md](shared/guardrails.md)）：

1. **保密先于生成**：发现密级标识、"内部资料""工作秘密"字样或个人敏感信息，先停下来提醒。
2. **事实只来自用户**：数据、案例、文号、政策依据都只取自你的材料；必须由你补的标【待补】，存疑的列在起草说明里。
3. **正文干净，提示外置**：正文里没有免责声明和 AI 自述，需要核实的放到"起草说明"。
4. **默认做减法**：开门见山，遵守减负规定的篇幅要求，不照抄上级文件；你指定字数就按字数写，不新增事实、不注水。
5. **规则可溯源、可机检**：每条格式、标点、数字规则都标注条款号，能用脚本检查的就不靠模型自觉。
6. **辅助不代笔**：对照检查、述职等需要本人经历的材料，没有素材就只给提纲。

规则依据：《党政机关公文处理工作条例》、GB/T 9704-2012、GB/T 15834-2011、GB/T 15835-2011、《整治形式主义为基层减负若干规定》。要点整理见 [shared/rules/](shared/rules/)，每条标注条款号。机检引擎按这些规则做 12 类检查，覆盖 41 类文书。

---

## 适用人群

不需要有多好的文笔，也不需要懂编程，只要你在用 AI 写公文：

- **机关事业单位文字工作者**：通知、请示、报告、函、纪要走量大、时间紧，格式和态度一个不能差
- **需要自己写材料的中层干部**：讲话稿、述职报告、工作要点，写好了提形象，写差了丢分
- **公立医院行政与科室人员**：等级评审自评报告、绩效分析、RCA/PDCA、科室主任述职——这些文书涉及 DRG、CMI、院感等指标口径，容易写错；医院行业包按公立医院的发文体系和指标规范处理
- **高校学院与职能部门**：审核评估自评、学科评估材料、申报书形式审查、校内签报和"学校发还是部门发"的判断——高校有一套两级发文和评估特有的文书逻辑，行业包按这套逻辑判断发文层级、检查漏项
- **国有企业综合与法务部门**："三重一大"议题材料、董事会议案与决议、总经理办公会纪要、ESG 报告、合规管理报告——国企的决策前置程序和合规边界是硬要求，行业包会核对前置研究，并区分决议和纪要
- **中小学教师与教务**：学期计划总结、督导评估自评、家长会讲话、学生评语、致家长一封信——面向家长的文风和督导迎检的口径都有讲究
- **乡镇街道社区工作人员**：迎检汇报、居民通知、应急通知、工作台账——面向群众的通知要通俗准确，迎检材料要对上检查清单
- **被 AI 生成稿折磨过的人**：编造、AI 腔、格式错误、态度含糊，每一条都要自己动手改

---

## 常见问题

**1. 这个项目是什么，收费吗？**

gongwen-skills 是给机关事业单位、国企文字工作者用的公文写作 Skills 集合，v1.0.0 有 22 个 skill。按 PolyForm Noncommercial 1.0.0 许可证源码公开（source-available，不是 OSI 意义上的开源），个人及非商业用途免费；商业用途请联系作者。

**2. 没有 Claude，能用吗？**

能。Skill 版支持 Claude Code、Codex、Cursor、Gemini CLI、GitHub Copilot、Qwen Code、Kimi Code CLI 等工具。没有这些工具的，用提示词版：打开 `dist/prompts/`，复制文件内容粘贴到 DeepSeek、豆包、Kimi 等网页大模型，或单位内网部署的模型。

**3. 单位的材料能直接贴进去吗？**

涉密材料和标注"内部资料""工作秘密"的材料，不能输入任何联网 AI 工具。处理内部材料请用单位内网部署的模型；提示词版和离线包可以在内网使用。skill 发现密级标识时会先停下来提醒，但判断责任在使用者。

**4. 它会不会编数据？**

规则要求只用你给的材料，必须由你补的标【待补】，存疑的列在起草说明。v1.0 评测中 136 条故意少给材料的用例，装 skill 出现编造 0 次。模型输出有随机性，定稿前请逐条核对数字、人名、文号。

**5. 去 AI 味能让稿子查不出是 AI 写的吗？**

不能，也不承诺这种效果。去 AI 味改的是文风质量，目的是让稿子更好读。

**6. 排出来的 Word 字体不对，或者没有 Python，怎么办？**

项目不随附字体，也不会用别的字体顶替。缺字体时，排版 skill 会在回复开头指出缺哪几款，并告诉你怎么安装；装好后重启 Word/WPS 再核对。没有 Python 时，排版 skill 给出 WPS/Word 手动设置清单。

更多问题见 [使用须知](使用须知.md)。

---

## 路线图

- [x] **v0.1「年终季版」**：总入口、综合文稿、标题、去 AI 味、去防御性、核稿、排版
- [x] **v0.2「开年版」**：单位画像、法定公文、讲话稿、党建材料、改稿润色
- [x] **v0.3「落实版」**：上级文件转化、整改材料、会议纪要、信息简报、数据与图表
- [x] **v0.4「行业包」**：基层、医院、高校、中小学、国企
- [x] **v1.0.0**（当前版本）：22 个 skill 一次发布；真实模型评测；一键安装脚本和行业包选装

**后续计划**

- [ ] 跨模型评测：提示词版在 DeepSeek、Qwen、Kimi 上的表现
- [ ] 收集测试者反馈，做人工盲评，据此修订规则，行业包达标后从 Beta 升级
- [ ] 在 Word、WPS 中实测排版效果；在 macOS、Linux 上实测字体检测
- [ ] 输出 Word 修订痕迹（目前只写批注）

想帮忙测试的同事，请看 [测试者手册](docs/评测/测试者手册.md)：用虚构材料试用，不需要懂编程。

---

## 关注与交流

<div align="center">

<img src="docs/assets/wecom.jpg" width="180" alt="企业微信" />

**扫码加我企业微信，进公文 skill 交流群**

商业授权 / 合作咨询也可以通过这里联系。

</div>

规则写错了、遇到新场景、发现 bug，请到 [Issues](https://github.com/duyunzhoudyz/gongwen-skills/issues/new/choose) 选对应表单；使用心得也可以提 Issue，或扫上面的二维码进群交流。参与开发见 [CONTRIBUTING.md](CONTRIBUTING.md)。

如果这个项目在你最需要的时候帮了你，可以点一个 Star，或者请作者喝杯茶。

<div align="center">

<img src="docs/assets/wechat-pay.jpg" width="180" alt="微信收款码" />

</div>

---

## 致谢

本项目借鉴了多个开源项目的思路或内容，许可证和借鉴方式见 [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)，逐条明细见 [docs/借鉴记录/](docs/借鉴记录/)。

---

## 许可证

源码公开。个人、学习、教育、研究及其他非商业用途可按许可证免费使用；商业使用需事先取得商业授权，包括商业集成、付费服务、SaaS、企业内部业务流程或商业再分发。

[PolyForm Noncommercial 1.0.0](LICENSE) — 规则要点整理自公开法规和国家标准，不替代原文。
