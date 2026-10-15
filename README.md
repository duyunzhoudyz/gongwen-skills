<div align="center">

# 公文 Skills · gongwen-skills

**懂规矩的公文写作助手：懂条例国标，懂单位差异，懂岗位角色；不编造，守保密，帮你做减法。**

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
![Skills](https://img.shields.io/badge/skills-7-brightgreen)
![Version](https://img.shields.io/badge/version-0.1.0-orange)
![Platforms](https://img.shields.io/badge/platforms-Claude%20Code%20%7C%20Codex%20%7C%20Cursor%20%7C%20Trae%20%7C%20Kimi%20%7C%20Qwen-lightgrey)
![Prompts](https://img.shields.io/badge/prompts-DeepSeek%20%7C%20Doubao%20%7C%20Kimi-lightgrey)

</div>

> **保密提示**：涉密信息严禁输入任何联网 AI 工具。处理内部材料请使用单位内网部署的模型；本项目的提示词版和离线包可以直接在内网模型上使用。详见 [使用须知](使用须知.md)。

---

## 为什么做这个

用 AI 写材料的人越来越多，翻车的也越来越多：

- **编造事实**：让 AI 配几个本地案例，结果张冠李戴、查无此企业；
- **一股 AI 味**：通篇"赋能""闭环"，"不是……而是……"，破折号满天飞；
- **态度含糊**：请示里没有倾向性意见，通知里全是"建议""可以考虑"，末尾还附一句"以上内容仅供参考"；
- **格式不对**：粘贴进 Word 满屏 `**` 和 `#`，还得自己按国标重新排版；
- **不懂单位**：医院的国考指标、高校的两级发文、国企的党委前置研究，通用 AI 一概不知。

这套 Skills 把老笔杆子的经验和一线摸索出的写作技巧，固化成可复用、可检验的工作流：

| | 一般的 AI 写作 | gongwen-skills |
|---|---|---|
| 数据和案例 | 缺了就编 | 只用你给的材料，缺的标【待补】 |
| 篇幅 | 越写越长 | 默认做减法，遵守减负规定的篇幅要求 |
| 文风 | AI 腔、新八股 | 去掉 AI 腔，保留"一是二是"、排比小标题这些公文正当写法 |
| 态度 | 对冲、免责声明 | 措辞明确，需要核实的事项放到正文之外 |
| 格式 | 自己排 | 一键生成 GB/T 9704 国标 docx，机检标点、数字、发文字号 |
| 单位差异 | 一刀切 | 7 类单位画像、5 种角色模式 |

## 效果：一篇 AI 起草的通知，体检前后

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
- "建议各学校成立……专班"改成"各学校要成立……专班"；
- 发文字号"[2026]第12号"改成"〔2026〕12号"，日期"11月05日"改成"11月5日"，"3～5万元"改成"3万～5万元"；
- 对家长"一律移交公安机关予以拘留"的表述，核查后发现区教育局没有处罚权，予以删除；
- 所有数字前后一致，脚本逐项比对确认。

另有两个示例：[医院科室年终总结](examples/A-医院科室年终总结/)（缺的 9 个数据全部标【待补】，没有编造）和 [申请经费该用函还是请示](examples/C-文种判断-函/)。

## 30 秒上手

**方式一：Claude Code 插件市场**

```bash
/plugin marketplace add OWNER/gongwen-skills
/plugin install gongwen@gongwen-skills
```

**方式二：通用安装（Codex、Cursor、Gemini CLI、GitHub Copilot 等）**

```bash
npx skills add OWNER/gongwen-skills
```

**方式三：国产 Agent 工具（Qwen Code、Kimi CLI、Trae、CodeBuddy 等）**

```bash
git clone https://github.com/OWNER/gongwen-skills.git
cd gongwen-skills
python tools/install.py --list-targets
python tools/install.py --target qwen
```

**方式四：网页大模型（DeepSeek、豆包、Kimi）或内网模型**

打开 [`dist/prompts/`](dist/prompts/)，复制对应的提示词文件，粘贴到对话开头，再说出你的需求。第一次用，建议从 `gongwen.md`（总入口）开始。

**方式五：让 Agent 自己装**

把这句话发给你的 Agent：

```text
请把 https://github.com/OWNER/gongwen-skills 中 skills/ 目录下的所有 skill 安装到你的 skills 目录，装好后告诉我有哪些可用。
```

> 排版功能需要 Python 3.9+ 和 `python-docx`：`pip install python-docx openpyxl`。没有 Python 也能用，排版会改为输出 WPS/Word 手动设置清单。

## 直接这样说

| 你说 | 由谁处理 |
|---|---|
| "我刚到办公室，领导让我写个东西报给区财政局要经费，该用什么文种？" | `gongwen` 总入口 |
| "下周要交科室年终总结，这是全年的工作记录，帮我写" | `gongwen-report` |
| "用这份年终总结的素材，再给我出一份个人述职" | `gongwen-report` |
| "帮我给这篇总结起三套小标题" | `gongwen-title` |
| "这篇是 DeepSeek 写的，AI 味太重，帮我改改" | `gongwen-humanize` |
| "这个通知里全是'建议''可以考虑'，帮我改得硬气一点" | `gongwen-direct` |
| "帮我看看这篇稿子有没有问题" | `gongwen-review` |
| "排成 Word，红头文件格式" | `gongwen-format` |

## 技能地图

| Skill | 能力 | 成熟度 |
|---|---|---|
| [`gongwen`](skills/gongwen/) | 总入口：识别任务、规划路线、建立单位画像、提示写作日历；v0.1 中还提供法定公文、讲话稿、纪要、信息的基础起草 | Beta |
| [`gongwen-report`](skills/gongwen-report/) | 综合文稿：总结、计划、工作要点、汇报材料、调研报告、经验材料、述职报告、实施方案 | Beta |
| [`gongwen-title`](skills/gongwen-title/) | 标题与提纲：公文标题、主副标题、三套小标题方案、提纲、标题与正文一致性检查 | Beta |
| [`gongwen-humanize`](skills/gongwen-humanize/) | 公文去 AI 味：带公文白名单，事实一字不改 | Beta |
| [`gongwen-direct`](skills/gongwen-direct/) | 去防御性写作：删免责声明和对冲语气，让请示、批复、通知态度明确 | Beta |
| [`gongwen-review`](skills/gongwen-review/) | 核稿审校：按"八关"核稿，分级问题清单，机检标点、数字、格式、隐私 | Beta |
| [`gongwen-format`](skills/gongwen-format/) | 国标排版：生成 GB/T 9704 docx，修复已有 docx，把核稿意见写成 Word 批注 | Beta |

**一份年终总结的完整流程**：总入口确认需求 → report 建素材台账、出提纲 → title 提炼小标题 → report 成稿 → direct 和 humanize 清理 AI 腔 → review 核稿 → format 排成国标 Word。

## 六条写作红线

所有 skill 都遵守这六条，优先级高于任何具体流程（全文见 [shared/guardrails.md](shared/guardrails.md)）：

1. **保密先于生成**：发现密级标识、内部资料字样或个人敏感信息，先停下来提醒。
2. **事实只来自用户**：数据、案例、文号、政策依据都只取自你的材料，缺的标【待补】，存疑的标【待核实】。
3. **正文干净，提示外置**：正文里没有免责声明和 AI 自述，需要核实的事项放到正文后的"起草说明"。
4. **默认做减法**：开门见山，遵守减负规定的篇幅要求，不照抄上级文件。
5. **规则可溯源、可机检**：每条格式、标点、数字规则都标注条款号，能用脚本检查的就不靠模型自觉。
6. **辅助不代笔**：对照检查、述职等需要本人经历的材料，没有素材就只给提纲。

## 不同单位，不同写法

内置 7 类单位画像：党政机关、基层（县乡、街道）、社区、公立医院、高校、中小学、国有企业。每类都整理了发文体系、常用文种、决策程序、术语、红线和年度写作节点。

内置 5 种角色模式：

| 模式 | 适合谁 | 有什么不同 |
|---|---|---|
| 教学模式 | 新入职公务员、选调生 | 解释为什么选这个文种、为什么这样改 |
| 高效模式 | 专职笔杆子 | 少解释，直接出多个版本 |
| 转译模式 | 兼职写材料的业务骨干 | 把大白话和数据表转成公文语言 |
| 审稿模式 | 中层领导、主要领导 | 只出问题清单和一页纸摘要 |
| 严谨模式 | 党务、宣传岗 | 提法核验最严，引用必须注明出处 |

第一次使用时，总入口会用 3 个问题为你建立单位画像（保存在 `.gongwen/单位画像.md`）。之后所有 skill 都会按你的单位和角色调整写法。

## 规则依据

《党政机关公文处理工作条例》、GB/T 9704-2012《党政机关公文格式》、GB/T 15834-2011《标点符号用法》、GB/T 15835-2011《出版物上数字用法》、《整治形式主义为基层减负若干规定》。要点整理见 [shared/rules/](shared/rules/)，每条都标注了条款号。

## 路线图

- [x] **v0.1「年终季版」**（2026 年 10 月）：总入口、综合文稿、标题、去 AI 味、去防御性、核稿、排版
- [ ] **v0.2「开年版」**（2026 年 11 月下旬）：单位画像与文风学习、法定公文、讲话稿、党建材料、改稿润色
- [ ] **v0.3「落实版」**（2027 年 1–2 月）：上级文件转化、整改材料、会议纪要、信息简报、数据与图表
- [ ] **v0.4「行业包」**（2027 年 3–5 月）：基层、医院、高校、中小学、国企
- [ ] **v1.0**：公开评测基准、稳定版

## 关注与交流

- 抖音：（待填写）
- 微信：（待填写）
- 抖音交流群、微信交流群：（二维码待添加）

欢迎提 issue 反馈规则错误，也欢迎在 Discussions 分享你所在行业的写作经验。贡献方式见 [CONTRIBUTING.md](CONTRIBUTING.md)。

## 致谢

本项目借鉴了以下开源项目的思路或内容，详见 [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)：

- [Yuan1z0825/nature-skills](https://github.com/Yuan1z0825/nature-skills)、[JimLiu/baoyu-skills](https://github.com/JimLiu/baoyu-skills)：Skill 集合的组织方式
- [op7418/Humanizer-zh](https://github.com/op7418/Humanizer-zh)、[larashero3-dotcom/lieflat-less-ai-tone](https://github.com/larashero3-dotcom/lieflat-less-ai-tone)、[LifelongLazyLearner/qu-ai-wei](https://github.com/LifelongLazyLearner/qu-ai-wei)、[MrGeDiao/shuorenhua](https://github.com/MrGeDiao/shuorenhua)：中文去 AI 味
- [hexiecs/talk-normal](https://github.com/hexiecs/talk-normal)、[molang163/chinese-natural-reply](https://github.com/molang163/chinese-natural-reply)：去防御性表达
- [mizzlelover/gongwen-gbt9704-skill](https://github.com/mizzlelover/gongwen-gbt9704-skill)、[KaguraNanaga/document-format-skills](https://github.com/KaguraNanaga/document-format-skills)、[imhuhe/gongwen-kit](https://github.com/imhuhe/gongwen-kit)、[linhut/gongwen-skill](https://github.com/linhut/gongwen-skill)：公文排版与检查

## 许可证

[MIT](LICENSE)。规则要点整理自公开的法规和国家标准，不替代原文。
