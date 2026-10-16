# 第三方声明

本项目在开发中参考了以下开源项目、法规和标准。所有外部仓库均为在线阅读，没有克隆或复制到本仓库。各仓库的许可证于 2026-10-16 通过 GitHub API 的 license 接口逐个核对，版权声明取自各仓库的 LICENSE 文件。

**本项目的代码、规则文字和例句均为自写**，没有整段照录任何来源。借鉴方式分为五种：

- **思路**：分类方法、判定原则、工作流程；
- **改写**：参考原规则，用公文语境的自拟表述重写；
- **数据**：引用原项目公开的统计结论，用来确定级别和阈值；
- **参数**：参考公开的版式参数取值，与国家标准交叉核对；
- **反例**：把来源中的写法或做法当作要避免的情形，例句自拟。

每个模块逐条的借鉴明细（读取版本、借鉴内容、落到哪些文件）见 [docs/借鉴记录/](docs/借鉴记录/)。

## 一、MIT 许可项目

### （一）公文写作与排版

| 项目 | 版权声明 | 借鉴方式 | 用于 |
|---|---|---|---|
| [mizzlelover/gongwen-gbt9704-skill](https://github.com/mizzlelover/gongwen-gbt9704-skill) | Copyright (c) 2026 mizzlelover | 参数、思路 | gongwen-format 版式参数、字体属性写法、页码与版记实现 |
| [KaguraNanaga/document-format-skills](https://github.com/KaguraNanaga/document-format-skills) | Copyright (c) 2025 Gloria LaRoche | 思路 | gongwen-format 已有 docx 的段落角色识别与修复 |
| [imhuhe/gongwen-kit](https://github.com/imhuhe/gongwen-kit) | Copyright (c) 2026 HUANGHE | 思路、改写、参数 | gongwen-review 数字自洽检查；各 skill 脚本测试沿用的“必错标本、必过标本”方法；gongwen-format 文档网格和页码位置取值 |
| [linhut/gongwen-skill](https://github.com/linhut/gongwen-skill) | Copyright (c) 2026 Jose AI | 思路 | P0/P1/P2 分级口径；核稿意见逐轮闭环；Word 批注定位 |
| [KaguraNanaga/official-document-writing-skill](https://github.com/KaguraNanaga/official-document-writing-skill) | Copyright (c) 2026 Official Document Writing Skill Contributors | 思路、改写、反例 | gongwen-title 标题拟定检查角度；gongwen-report 总结类检查项；gongwen-draft 会议通知、请示、复函的要素；gongwen-speech 开场套话反例 |
| [Liuxiangjian-ai/official-document-skill](https://github.com/Liuxiangjian-ai/official-document-skill) | Copyright (c) 2026 Xiangjian Liu | 思路、改写 | 文种判断顺序与易混文种辨析；方案与计划的区分；表态发言结构；简报、宣传稿的写作要求 |
| [gongyu0918-debug/chinese-official-writing-skill](https://github.com/gongyu0918-debug/chinese-official-writing-skill) | Copyright 2026 chinese-official-writing contributors | 思路、改写 | 综合文稿、标题、法定公文、讲话、党建、纪要、信息、改稿各类文书的写作要求；基层、医院、高校、国企行业包中倡议书、公开信、整改材料、申报书、签报、决议的写法 |
| [zhaohui-yang/official-document-drafting](https://github.com/zhaohui-yang/official-document-drafting) | Copyright (c) 2026 official-document-drafting contributors | 改写 | gongwen-draft 请示、批复、函、通报、通告、决定、通知的写作规格；gongwen-minutes 按会议类型区分纪要结构 |

zhaohui-yang/official-document-drafting 的 LICENSE 写明 MIT 只适用于仓库中的原创代码和文档，字体等第三方资产另行授权（GitHub API 因此显示为 NOASSERTION）。本项目只参考其文书规格，没有使用其字体。

### （二）去 AI 味与去防御性写作

| 项目 | 版权声明 | 借鉴方式 | 用于 |
|---|---|---|---|
| [op7418/Humanizer-zh](https://github.com/op7418/Humanizer-zh) | Copyright (c) 2026 歸藏 | 思路、改写 | 中文去 AI 味规则骨架；改写工作流 |
| [larashero3-dotcom/lieflat-less-ai-tone](https://github.com/larashero3-dotcom/lieflat-less-ai-tone) | Copyright (c) 2026 shiujan | 数据、思路、改写 | 11 项有量化证据的中文 AI 特征及判定原则 |
| [LifelongLazyLearner/qu-ai-wei](https://github.com/LifelongLazyLearner/qu-ai-wei) | Copyright (c) 2026 @LifelongLazyLearner | 思路 | “何时保留”的组织方式；公文平行结构保留；改写前冻结事实 |
| [MrGeDiao/shuorenhua](https://github.com/MrGeDiao/shuorenhua) | Copyright (c) 2026 MrGeDiao | 思路、改写 | 改写时保持事实不变的约束；只标注不改的模式；评测用例的设计 |
| [VincentOld/stop-slop-zh](https://github.com/VincentOld/stop-slop-zh) | Copyright (c) 2026 文森 (Wen Sen) | 思路、改写 | 词表分组方式；候选词逐词筛选 |
| [blader/humanizer](https://github.com/blader/humanizer) | Copyright (c) 2025 Siqi Chen | 思路、改写 | 限定词堆叠、聊天残留、知识边界免责的保留边界 |
| [hexiecs/talk-normal](https://github.com/hexiecs/talk-normal) | Copyright (c) 2026 hexie | 思路、改写 | 否定对举、总结标签、平衡论述、结尾征询 |
| [molang163/chinese-natural-reply](https://github.com/molang163/chinese-natural-reply) | Copyright (c) 2026 molang163 | 思路 | 中文防御性回复的分类与判定 |
| [MohamedAbdallah-14/unslop](https://github.com/MohamedAbdallah-14/unslop) | Copyright (c) 2026 Mohamed Abdallah | 思路 | 严重度分级；保留真实不确定性 |
| [huacnlee/autocorrect](https://github.com/huacnlee/autocorrect) | Copyright (c) 2020 Jason Lee | 思路 | 半角标点紧邻中文才判错的判定思路 |

### （三）Skill 组织、数据工具与行业包

| 项目 | 版权声明 | 借鉴方式 | 用于 |
|---|---|---|---|
| [JimLiu/baoyu-skills](https://github.com/JimLiu/baoyu-skills) | Copyright (c) 2026 Jim Liu | 思路 | gongwen-profile 配置文件的读取顺序（项目级优先于用户级）；共享资料在构建时复制进各 skill 的“自包含”做法 |
| [alchaincyf/huashu-excel](https://github.com/alchaincyf/huashu-excel) | Copyright (c) 2026 huashu | 思路、改写 | gongwen-data 报表体检：合计行对账、空值词表、会计负数、编码识别 |
| [hustcc/mcp-echarts](https://github.com/hustcc/mcp-echarts) | Copyright (c) 2025 hustcc | 思路 | gongwen-data 图表在本机生成、不依赖远程服务 |
| [antvis/mcp-server-chart](https://github.com/antvis/mcp-server-chart) | Copyright (c) 2025 AntV Visualization Team | 反例 | 其默认配置把图表数据发往远程渲染接口，gongwen-data 因此不使用这类插件 |
| [blingenf/copydetect](https://github.com/blingenf/copydetect) | Copyright (c) 2020 Bryson Lingenfelter | 思路 | `shared/scripts/overlap_check.py` 的重合率定义和最小匹配长度 |
| [huangwb8/ChineseResearchLaTeX](https://github.com/huangwb8/ChineseResearchLaTeX) | Copyright © 2024 Weibin Huang | 思路 | gongwen-university 申报书形式检查：只读不改、分级意见、栏目篇幅分布 |
| [a5c-ai/babysitter](https://github.com/a5c-ai/babysitter) | Copyright (c) 2025 A5C AI | 思路 | gongwen-hospital 不良事件根因分析（RCA）报告的工作顺序 |
| [PIGU-PPPgu/class-teacher-skills](https://github.com/PIGU-PPPgu/class-teacher-skills) | Copyright (c) 2026（LICENSE 未写权利人） | 改写 | gongwen-k12 学生评语、家长会讲话的写作要求 |

## 二、Apache-2.0 许可项目

| 项目 | 借鉴方式 | 用于 |
|---|---|---|
| [Yuan1z0825/nature-skills](https://github.com/Yuan1z0825/nature-skills) | 思路 | Skill 集合组织方式、路由协议、成熟度标签、相邻 skill 的转交边界 |
| [anthropics/k12-teacher-skills](https://github.com/anthropics/k12-teacher-skills) | 思路 | gongwen-k12 评测期望的写法：每条只判一件事，提问行为单列，部分条目按条件适用 |
| [bowenliang123/markdown-exporter](https://github.com/bowenliang123/markdown-exporter) | 思路 | `shared/scripts/md2xlsx.py` 每张表占一个工作表、编号类单元格存为文本 |

只借鉴了设计思路，没有使用其代码或文字，因此无需随附其 NOTICE 文件。

## 三、只学思路、未使用任何文字和代码的项目

以下项目和资料的许可证不允许或未声明允许再分发。本项目只参考了公开的思路或统计结论，没有复制其中任何文字、表格或代码。

| 项目 | 许可证 | 参考内容 |
|---|---|---|
| [larashero3-dotcom/lieflat-gongwen](https://github.com/larashero3-dotcom/lieflat-gongwen) | PolyForm Noncommercial 1.0.0（GitHub API 显示为 NOASSERTION，以仓库 LICENSE 文件为准） | 公开的公文语料统计结论（用于校准阈值）；【待补】占位机制；文风参数化的方法（gongwen-profile 的参数只从用户自己的历史稿计算，不内置其统计值）；领导讲话中力度词的用法 |
| [Rimagination/gongwen-draft](https://github.com/Rimagination/gongwen-draft) | 未声明 | “先查、先核、再写”；先按行文目的匹配文种、再核发文权限的顺序 |
| [alvis961/sop-writing](https://github.com/alvis961/sop-writing) | 未声明 | 按材料类型区分标题风格 |
| [cuberfry68-coder/wenxi-skills](https://github.com/cuberfry68-coder/wenxi-skills) | 未声明 | 标题逐级展开同一核心意思 |
| [jzOcb/writing-style-skill](https://github.com/jzOcb/writing-style-skill) | 未声明 | 比对初稿与定稿提取改稿规则；规则按出现次数分级；更新前备份 |
| [anthropics/skills](https://github.com/anthropics/skills) 的 `.claude-plugin/marketplace.json` | 仓库根目录未声明许可证，各 skill 分别授权 | 插件市场清单的字段结构（`plugins[].source`、`plugins[].skills`）；本项目的清单自行编写 |
| RCA²: Improving Root Cause Analyses and Actions to Prevent Harm（美国国家患者安全基金会，2015 年） | 受版权保护的指南，未取得全文 | gongwen-hospital 改进措施按强、中、弱分级的思路；分级描述和例子自拟，标为工具约定 |

## 四、明确未参考的内容

- [anthropics/skills](https://github.com/anthropics/skills) 中的 docx、xlsx、pdf、pptx 四个 skill 为专有许可（source-available），本项目开发中未阅读、未参考。
- larashero3-dotcom/lieflat-gongwen 的 `示例/` 目录中的领导讲话范文：未阅读、未使用。
- 调研中见到的 AGPL-3.0 许可的 Word 修订解析实现（scitex-msword）：未参考其代码。gongwen-profile 的修订还原按 ECMA-376 标准自行实现。
- 网上流传的对照检查材料、个人发言提纲等“范文”：来源和许可不明，未参考其文字。gongwen-party 的示范稿只给结构，事例位置一律占位。

## 五、法规与标准

### （一）公文处理规范与国家标准

以下文件的要点整理见 [shared/rules/](shared/rules/)，每条标注条款号，不替代原文：

- 《党政机关公文处理工作条例》（中办发〔2012〕14号）
- GB/T 9704-2012《党政机关公文格式》
- GB/T 15834-2011《标点符号用法》
- GB/T 15835-2011《出版物上数字用法》
- 《整治形式主义为基层减负若干规定》（2024年）
- 《国务院办公厅关于实施〈国家行政机关公文处理办法〉涉及的几个具体问题的处理意见》（国办函〔2001〕1号），只用于意见文种的行文方向

国家标准原文受著作权保护，本项目只收录用自己的话整理的要点，不收录原文。

### （二）各 skill 和行业包引用的法律法规与规范性文件

党内法规、法律、行政法规和规范性文件，依《中华人民共和国著作权法》第五条不适用该法。本项目只按条款号概括要点或短句引用，不整段照录。以下只列名称和发文字号（没有文号的注明年份或发布机关）；用到的条款、核实日期见 [docs/借鉴记录/](docs/借鉴记录/) 和各 skill 的 references，如 gongwen-k12 的 `references/sources.md`。

**党内法规和中央文件**

- 《中国共产党纪律处分条例》（2023年修订）
- 《中国共产党巡视工作条例》（2024年修订）
- 《县以上党和国家机关党员领导干部民主生活会若干规定》（2016年）
- 《中国共产党支部工作条例（试行）》（2018年）
- 《中国共产党党员教育管理工作条例》（2019年）
- 《党委（党组）书记抓基层党建工作述职评议考核办法（试行）》（2019年）
- 《中国共产党农村基层组织工作条例》（2019年）
- 《中国共产党国有企业基层组织工作条例（试行）》（2019年）
- 《中国共产党普通高等学校基层组织工作条例》（2021年修订）
- 《关于加强公立医院党的建设工作的意见》（中办发〔2018〕35号）
- 《关于建立中小学校党组织领导的校长负责制的意见（试行）》（2022年）
- 《关于进一步推进国有企业贯彻落实“三重一大”决策制度的意见》（中办发〔2010〕17号）
- 《关于在深化国有企业改革中坚持党的领导加强党的建设的若干意见》（2015年）
- 《关于中央企业在完善公司治理中加强党的领导的意见》（2021年）
- 《中共中央 国务院关于加强基层治理体系和治理能力现代化建设的意见》（2021年）
- 《关于规范村级组织工作事务、机制牌子和证明事项的意见》（2022年）
- 《关于全面建立乡镇（街道）履行职责事项清单的意见》（2024年）
- 《关于进一步减轻义务教育阶段学生作业负担和校外培训负担的意见》（2021年）
- 《关于减轻中小学教师负担进一步营造教育教学良好环境的若干意见》（2019年）

**法律**

- 《中华人民共和国民法典》
- 《中华人民共和国行政处罚法》（2021年修订）
- 《中华人民共和国行政强制法》
- 《中华人民共和国个人信息保护法》
- 《中华人民共和国广告法》
- 《中华人民共和国审计法》（2021年修正）
- 《中华人民共和国公司法》（2023年修订）
- 《中华人民共和国城市居民委员会组织法》（2025年修订）
- 《中华人民共和国村民委员会组织法》（2025年修改）
- 《中华人民共和国突发事件应对法》（2024年修订）
- 《中华人民共和国未成年人保护法》（2020年修订）
- 《中华人民共和国家庭教育促进法》

**行政法规和国务院文件**

- 《重大行政决策程序暂行条例》（国务院令第713号）
- 《保障农民工工资支付条例》（国务院令第724号）
- 《森林草原防灭火条例》（2025年）
- 《国务院办公厅关于加强行政规范性文件制定和监督管理工作的通知》（国办发〔2018〕37号）
- 《国务院办公厅关于推进政务新媒体健康有序发展的意见》（国办发〔2018〕123号）
- 《国务院办公厅关于加强三级公立医院绩效考核工作的意见》（国办发〔2019〕4号）
- 《国务院办公厅关于进一步完善国有企业法人治理结构的指导意见》（国办发〔2017〕36号）
- 《国务院办公厅转发司法部关于审理政府信息公开行政复议案件若干问题指导意见的通知》（国办函〔2021〕132号），只作转发标题写法的实例
- 《安全生产治本攻坚三年行动方案（2024—2026年）》（国务院安委会，2024年），只用于核对评测材料中的年份表述
- 《基层报表数据“只报一次”工作指引》（国务院办公厅）

**部门规章和规范性文件：民政、应急、气象**

- 《气象灾害预警信号发布与传播办法》（中国气象局令第16号）
- 《最低生活保障审核确认办法》（民发〔2021〕57号）
- 《关于改进和规范基层群众性自治组织出具证明工作的指导意见》（民政部等六部门，2020年）

**部门规章和规范性文件：卫生健康、医疗保障**

- 《医疗质量管理办法》（国家卫生计生委令第10号）
- 《三级医院评审标准（2025年版）》（国卫医政发〔2025〕4号）
- 《三级中医医院评审标准（2026年版）》（国中医药医政函〔2026〕39号）
- 《三级综合医院评审标准实施细则（2011年版）》（原卫生部），只用于说明“星号”一词的来历
- 国家三级、二级公立医院绩效考核操作手册（2024版）（国卫办医政函〔2024〕87号、86号）
- 《关于启动2025年度二级和三级公立医院绩效监测有关工作的通知》（国家卫生健康委办公厅，2026年）
- 《患者安全专项行动方案（2023—2025年）》（国卫办医政发〔2023〕13号）
- 《2026年国家医疗质量安全改进目标》（国家卫生健康委）
- 《医疗质量安全核心制度要点》（国家卫生健康委，2018年）
- 《医疗机构手术分级管理办法》（国家卫生健康委，2022年）
- 《医疗机构工作人员廉洁从业九项准则》（国卫医发〔2021〕37号）
- 《2026年纠正医药购销领域和医疗服务中不正之风工作要点》（国卫医急函〔2026〕123号）
- 按病组和病种分值付费2.0版分组方案（医保办发〔2024〕9号）
- DRG/DIP 支付方式改革三年行动计划（国家医保局，2021年）

**部门规章和规范性文件：教育、科研**

- 《普通高等学校本科教育教学审核评估实施方案（2021—2025年）》（教督〔2021〕1号）
- 《普通高等学校本科教育教学审核评估指标体系（试行）》（教育部）
- 《第五轮学科评估指标体系框架》（教育部学位与研究生教育发展中心，2020年）
- 《普通高等学校辅导员队伍建设规定》（教育部令第43号）
- 《学校教职工代表大会规定》（教育部令第32号）
- 《2026年国家社会科学基金年度项目申报公告》（全国哲学社会科学工作办公室）
- 《关于2026年度国家自然科学基金项目申请与结题等有关事项的通告》（国科金发计〔2026〕2号）
- 《义务教育学校管理标准》（教基〔2017〕9号）
- 《关于加强义务教育学校考试管理的通知》（教基厅函〔2021〕34号）
- 《关于进一步加强中小学日常考试管理的通知》（教基厅〔2025〕3号）
- 《中小学教辅材料管理办法》（新广出发〔2015〕45号）
- 《关于加强和改进普通高中学生综合素质评价的意见》（教基二〔2014〕11号）
- 《义务教育质量评价指南》（教基〔2021〕3号）
- 《县域义务教育优质均衡发展督导评估办法》（教督〔2017〕6号）
- 《中小学校责任督学挂牌督导办法》（国教督办〔2013〕2号）
- 《关于进一步减轻中小学教师非教育教学负担若干措施的通知》（教师厅〔2025〕2号）
- 《中小学班主任工作规定》（教基一〔2009〕12号）
- 《关于健全学校家庭社会协同育人机制的意见》（教基〔2022〕7号）
- 《关于加强和改进新时代基础教育教研工作的意见》（教基〔2019〕14号）
- 《中小学幼儿园安全管理办法》（教育部等十部门令第23号）
- 《严禁中小学校和在职中小学教师有偿补课的规定》（教师〔2015〕5号）
- 教育部历年招生工作通知中严禁宣传炒作“状元”和升学率的要求（未取得条款号，正文写“以当年文件为准”）

**部门规章和规范性文件：国资、证券、财政、工会**

- 《中央企业合规管理办法》（国资委令第42号）
- 《中央企业商业秘密保护暂行规定》（国资发〔2010〕41号）
- 《关于新时代中央企业高标准履行社会责任的指导意见》（国资委，2024年）
- 《关于转发〈央企控股上市公司ESG专项报告编制研究〉的通知》（国资委办公厅，2023年）
- 《上市公司信息披露管理办法》（证监会令第226号）
- 《企业可持续披露准则——基本准则（试行）》（财会〔2024〕17号）
- 《企业可持续披露准则第1号——气候（试行）》（财政部，2025年）
- 沪深北证券交易所《上市公司可持续发展报告指引》（2024年）
- 《企业民主管理规定》（总工发〔2012〕12号）

### （三）其他国家标准和行业标准

- GB/T 36001-2015《社会责任报告编写指南》，只引用编号和名称，未读取正文。
- WS/T 598.7—2018《卫生统计指标 第7部分：医疗服务》，只引用平均住院日的定义。

### （四）技术规范

- Agent Skills 开放规范：frontmatter 字段约束，用于 `tools/validate.py`。
- ECMA-376 Office Open XML：修订标记和批注的公开定义，用于 gongwen-profile 的修订还原和批注读取。
- GitHub Flavored Markdown 规范（CC BY-SA 4.0）：表格语法，按规范自行实现解析，未复制规范文字，用于 `shared/scripts/md2xlsx.py`。
- GitHub issue forms 语法文档、PEP 668：Issue 模板和安装脚本的提示按其约定自行编写。

### （五）其他公开资料

核实制度性事实时，还在线查阅了政府网站和官方媒体的公开报道，以及国家统计局“统计知识”栏目等公开资料，只引用事实、定义和公式，表述自拟。部分机关、高校、医院和企业在官网公开的本单位制度，只用于了解通行做法，没有收录其文字，这里不逐一列出。本声明不列任何单位的内部文件。

## 六、运行依赖

以下 Python 库在运行时由用户自行安装，本项目不随附：

| 库 | 许可证 | 用途 | 是否必需 |
|---|---|---|---|
| [python-docx](https://github.com/python-openxml/python-docx) | MIT | 生成、读取、修复 docx，写入批注 | 排版功能需要 |
| [openpyxl](https://foss.heptapod.net/openpyxl/openpyxl) | MIT | 读取 Excel；生成整改台账表格和 Excel 原生图表 | 数据、台账功能需要 |
| [matplotlib](https://github.com/matplotlib/matplotlib) | Matplotlib License（PSF 风格，与 BSD 兼容） | gongwen-data 另出 PNG 图表 | 可选 |
| [PyMuPDF](https://github.com/pymupdf/PyMuPDF) | AGPL-3.0（另有商业许可） | gongwen-format 预览时把 PDF 转成逐页 PNG | 可选 |

可选依赖只在用户本机已安装时调用，缺少时自动跳过。PyMuPDF 为 AGPL-3.0 许可，如需在对外提供的服务中使用，请自行评估其许可义务。gongwen-format 的预览功能还会调用本机已安装的 LibreOffice（MPL-2.0）把 docx 转成 PDF，未安装时跳过。

本项目不随附任何字体文件。仿宋_GB2312、方正小标宋简体等字体请使用者自行合法获取。
