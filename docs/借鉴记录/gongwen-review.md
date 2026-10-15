# 借鉴记录：核稿审校（gongwen-review）与机检引擎

- 模块范围：`skills/gongwen-review/`（SKILL.md、references、evals）、`shared/scripts/gongwen_lint.py`、`shared/scripts/lint_rules.json`、`tests/lint/`
- 阅读方式：2026-10-16 通过 GitHub API 和 raw.githubusercontent.com 在线读取默认分支，未克隆、未下载到本仓库。当日 GitHub API 对提交记录接口限流，未取得提交号，以读取日期为准。
- 借鉴方式：**思路**（分级方法、判定原则、测试方法，不涉及文字和代码）；**改写**（参考原规则的检查项，用本项目的公文语境和正则重新实现）。本模块代码、规则、例句均为自写，未复制任何来源的代码或文字。

## 一、MIT 许可来源

| 来源 | 许可证 | 借鉴内容 | 方式 | 落到的文件 |
|---|---|---|---|---|
| [linhut/gongwen-skill](https://github.com/linhut/gongwen-skill) | MIT（LICENSE 文件核对） | P0 必须修复、P1 建议、P2 可选的三级分级；核稿意见逐轮“采纳／不采纳（写明理由）”闭环；修订加批注两种落地方式（由 gongwen-format 承接） | 思路 | SKILL.md 第五步定级标准；`references/opinion-template.md`；核稿意见.json 交 gongwen-format 写批注的分工 |
| [imhuhe/gongwen-kit](https://github.com/imhuhe/gongwen-kit) | MIT（LICENSE 文件核对） | `checks/check-num.js` 的数字自洽检查思路：“以下四个方面”与其后下一级标题或“一是……”条数比对；宁可漏报不误报、提示类不阻断；`tests/negative.js` 的“必错标本、必过标本”负对照测试法（证明每条规则都能报错，防止检查器空跑）；`check-style.js` 中工具故障与稿件问题用不同退出码区分 | 思路、改写 | `gongwen_lint.py` 的 STR-002（数字概括与条数，P2）；`lint_rules.json` 每条规则的 `hit_examples`、`miss_examples`；`tests/lint/test_rule_examples.py`；退出码 2 表示输入错误 |
| [huacnlee/autocorrect](https://github.com/huacnlee/autocorrect) | MIT（GitHub API 核对） | `rule/fullwidth.rs` 的判定思路：半角标点只在紧邻中文时才判为错误，小数点、网址、英文句子因不与中文相邻而自然排除；逗号、问号看单侧，句号、冒号看两侧或行尾 | 思路 | `lint_rules.json` PUN-012（半角标点）的正则设计、`skip_line` 排除网址 |

## 二、权威公开资料（规则依据，非开源项目）

| 来源 | 用途 | 落到的文件 |
|---|---|---|
| 共产党员网 2021-04-17 转载《秘书工作》文章《公文校核中常见的文字错谬有哪些？案例来啦~》 | “老少边穷”已停用及规范提法；“以社区党组织为核心”“总任务”两个校核案例 | `references/tifa-risk.md` 第三部分（一）（三） |
| 《国务院机构改革方案》（2018-03-17 十三届全国人大一次会议批准） | 国土资源部、环境保护部、农业部、国家卫生和计划生育委员会不再保留 | `references/tifa-risk.md` 第三部分（二） |
| 《国务院机构改革方案》（2023 年 3 月十四届全国人大一次会议批准，新华社 2023-03-10 播发） | 不再保留中国银行保险监督管理委员会，组建国家金融监督管理总局 | 同上 |

GB/T 9704-2012、GB/T 15834-2011、GB/T 15835-2011、《党政机关公文处理工作条例》、《整治形式主义为基层减负若干规定》的条款均引自本仓库 `shared/rules/` 的要点整理，未另行照录标准原文。

## 三、有意没有采用的内容

1. linhut/gongwen-skill 的“六维质量评分、低于 80 分须重写”：分数缺少可溯源的依据，与本项目“规则可溯源”和“默认只提意见”冲突，不采用。
2. gongwen-kit 对表格合计、跨表数值的硬性比对：v0.1 机检只处理正文文本，表格合计交模型在文字关人工核算；docx 表格单元格只做隐私扫描。
3. gongwen-kit 把引号内文字整体排除在文风扫描之外：本项目文风词表由 style 模块维护，引擎按词表约定的语义实现，不额外加引号豁免，避免与其他 skill 的调用结果不一致。
4. 新华社《新闻信息报道中的禁用词和慎用词》：只找到转载版本，未能核对官方原文，tifa-risk.md 只作为核对渠道提及，不列具体条目。
5. 未能找到可靠出处的“最新提法”一律不收（宁缺毋滥）。
