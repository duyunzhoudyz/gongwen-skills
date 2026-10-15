# 第三方声明

本项目在开发中参考了以下开源项目、法规和标准。所有外部仓库均为在线阅读，没有克隆或复制到本仓库。

**本项目的代码、规则文字和例句均为自写**，没有整段照录任何来源。借鉴方式分为四种：

- **思路**：分类方法、判定原则、工作流程；
- **改写**：参考原规则，用公文语境的自拟表述重写；
- **数据**：引用原项目公开的统计结论，用来确定级别和阈值；
- **参数**：参考公开的版式参数取值，与国家标准交叉核对。

每个模块逐条的借鉴明细（读取版本、借鉴内容、落到哪些文件）见 [docs/借鉴记录/](docs/借鉴记录/)。

## 一、MIT 许可项目

| 项目 | 版权声明 | 借鉴方式 | 用于 |
|---|---|---|---|
| [mizzlelover/gongwen-gbt9704-skill](https://github.com/mizzlelover/gongwen-gbt9704-skill) | Copyright (c) 2026 mizzlelover | 参数、思路 | gongwen-format 版式参数、字体属性写法、页码与版记实现 |
| [KaguraNanaga/document-format-skills](https://github.com/KaguraNanaga/document-format-skills) | Copyright (c) 2025 Gloria LaRoche | 思路 | gongwen-format 已有 docx 的段落角色识别与修复 |
| [imhuhe/gongwen-kit](https://github.com/imhuhe/gongwen-kit) | Copyright (c) 2026 HUANGHE | 思路、改写 | 数字自洽检查；"必错标本、必过标本"测试法 |
| [linhut/gongwen-skill](https://github.com/linhut/gongwen-skill) | Copyright (c) 2026 Jose AI | 思路 | P0/P1/P2 分级口径；Word 批注定位 |
| [KaguraNanaga/official-document-writing-skill](https://github.com/KaguraNanaga/official-document-writing-skill) | 见原仓库 LICENSE | 思路、改写 | 标题拟定检查角度；总结类检查项 |
| [Liuxiangjian-ai/official-document-skill](https://github.com/Liuxiangjian-ai/official-document-skill) | 见原仓库 LICENSE | 思路、改写 | 文种判断顺序；方案与计划的区分 |
| [gongyu0918-debug/chinese-official-writing-skill](https://github.com/gongyu0918-debug/chinese-official-writing-skill) | 见原仓库 LICENSE | 改写 | 总结、述职、要点、调研报告的写作要求 |
| [op7418/Humanizer-zh](https://github.com/op7418/Humanizer-zh) | 见原仓库 LICENSE | 思路、改写 | 中文去 AI 味规则骨架；改写工作流 |
| [larashero3-dotcom/lieflat-less-ai-tone](https://github.com/larashero3-dotcom/lieflat-less-ai-tone) | 见原仓库 LICENSE | 数据、思路、改写 | 11 项有量化证据的中文 AI 特征及判定原则 |
| [LifelongLazyLearner/qu-ai-wei](https://github.com/LifelongLazyLearner/qu-ai-wei) | 见原仓库 LICENSE | 思路 | "何时保留"的组织方式；公文平行结构保留 |
| [MrGeDiao/shuorenhua](https://github.com/MrGeDiao/shuorenhua) | 见原仓库 LICENSE | 思路 | 改写时保持事实不变的约束 |
| [VincentOld/stop-slop-zh](https://github.com/VincentOld/stop-slop-zh) | 见原仓库 LICENSE | 思路、改写 | 词表分组方式 |
| [blader/humanizer](https://github.com/blader/humanizer) | 见原仓库 LICENSE | 思路 | 限定词堆叠、聊天残留、知识边界免责的保留边界 |
| [hexiecs/talk-normal](https://github.com/hexiecs/talk-normal) | 见原仓库 LICENSE | 思路 | 否定对举、总结标签、平衡论述 |
| [molang163/chinese-natural-reply](https://github.com/molang163/chinese-natural-reply) | 见原仓库 LICENSE | 思路 | 中文防御性回复的分类与判定 |
| [MohamedAbdallah-14/unslop](https://github.com/MohamedAbdallah-14/unslop) | 见原仓库 LICENSE | 思路 | 严重度分级；保留真实不确定性 |
| [huacnlee/autocorrect](https://github.com/huacnlee/autocorrect) | 见原仓库 LICENSE | 思路 | 半角标点紧邻中文才判错的判定思路 |

## 二、Apache-2.0 许可项目

| 项目 | 借鉴方式 | 用于 |
|---|---|---|
| [Yuan1z0825/nature-skills](https://github.com/Yuan1z0825/nature-skills) | 思路 | Skill 集合组织方式、路由协议、成熟度标签 |

只借鉴了设计思路，没有使用其代码或文字。

## 三、只学思路、未使用任何文字和代码的项目

以下项目的许可证不允许或未声明允许再分发。本项目只参考了公开的思路或统计结论，没有复制其中任何文字、表格或代码。

| 项目 | 许可证 | 参考内容 |
|---|---|---|
| [larashero3-dotcom/lieflat-gongwen](https://github.com/larashero3-dotcom/lieflat-gongwen) | PolyForm Noncommercial 1.0.0 | 公开的公文语料统计结论（用于校准阈值）；【待补】占位机制的思路 |
| [Rimagination/gongwen-draft](https://github.com/Rimagination/gongwen-draft) | 未声明 | "先查、先核、再写"的思路 |
| [alvis961/sop-writing](https://github.com/alvis961/sop-writing) | 未声明 | 按材料类型区分标题风格的思路 |
| [cuberfry68-coder/wenxi-skills](https://github.com/cuberfry68-coder/wenxi-skills) | 未声明 | 标题逐级展开同一核心意思的思路 |

## 四、明确未参考的内容

- [anthropics/skills](https://github.com/anthropics/skills) 中的 docx、xlsx、pdf、pptx 四个 skill 为专有许可（source-available），本项目开发中未阅读、未参考。

## 五、法规与标准

以下文件的要点整理见 [shared/rules/](shared/rules/)，每条标注条款号，不替代原文：

- 《党政机关公文处理工作条例》（中办发〔2012〕14号）
- GB/T 9704-2012《党政机关公文格式》
- GB/T 15834-2011《标点符号用法》
- GB/T 15835-2011《出版物上数字用法》
- 《整治形式主义为基层减负若干规定》（2024年）

国家标准原文受著作权保护，本项目只收录用自己的话整理的要点，不收录原文。

## 六、运行依赖

以下 Python 库在运行时由用户自行安装，本项目不随附：

| 库 | 许可证 | 用途 |
|---|---|---|
| [python-docx](https://github.com/python-openxml/python-docx) | MIT | 生成、读取、修复 docx，写入批注 |
| [openpyxl](https://foss.heptapod.net/openpyxl/openpyxl) | MIT | 读取 Excel（预留） |

本项目不随附任何字体文件。仿宋_GB2312、方正小标宋简体等字体请使用者自行合法获取。
