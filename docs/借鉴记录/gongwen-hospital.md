# 借鉴记录：公立医院行业包（gongwen-hospital）

- 模块范围：`skills/gongwen-hospital/` 下的 SKILL.md、README.md、references/（不含 `_shared/`）、scripts/hospital_check.py、scripts/hospital_rules.json、evals/；`tests/hospital/`
- 阅读方式：2026-10-16 通过 GitHub API 和 raw.githubusercontent.com 在线读取默认分支，未克隆、未下载到本仓库。“读取版本”为当日默认分支的提交号前 7 位。政策文件通过政府网站、新华网、行业媒体在线核实；《三级医院评审标准（2025年版）》附件为扫描版 PDF，由网页工具取回后在会话临时目录中渲染成图片逐页阅读，未存入本仓库。
- 借鉴方式：**思路**（流程、判定原则，不涉及文字和代码）；**改写**（参考原文的检查项，用本项目的医院语境和例句重写）。本模块的规则、代码、例句、评测材料均为自写，例句和材料中的单位、人员、数据全部虚构。

## 一、MIT 许可来源

| 来源 | 许可证 | 读取版本 | 借鉴内容 | 方式 | 落到的文件 |
|---|---|---|---|---|---|
| [gongyu0918-debug/chinese-official-writing-skill](https://github.com/gongyu0918-debug/chinese-official-writing-skill) | MIT | 23e124c | `references/genre-playbook-remediation-plan.md`：整改材料“状态先行”；把形成方案当作完成整改、把建立台账开会发文培训当作整改完成，属于常见毛病；原因分析不越过材料依据去断定个人动机和责任 | 改写 | `references/templates/zhenggai.md` 第五部分第2、4、5条；`references/templates/rca-pdca.md` 第三部分第1、6条 |
| [a5c-ai/babysitter](https://github.com/a5c-ai/babysitter) | MIT | feb68ab | `library/specializations/domains/social-sciences-humanities/healthcare/skills/patient-safety-event-analysis/SKILL.md`：RCA 的工作顺序（上报、组建团队、收集信息、重建时间线、识别促成因素、确定根本原因、制定措施、实施监测） | 思路 | `references/templates/rca-pdca.md` 第二部分的报告结构 |

## 二、只学思路的来源

| 来源 | 性质 | 借鉴内容 | 处理方式 | 影响的设计 |
|---|---|---|---|---|
| RCA²: Improving Root Cause Analyses and Actions to Prevent Harm（美国国家患者安全基金会，2015年；现由 IHI 维护） | 受版权保护的指南，未取得原文全文 | 改进措施按“强、中、弱”分级：改环境设备、强制功能、简化和统一流程属强；增加人手、调整排班、系统提醒、核查表属中；双人核对、警示、新制度、培训属弱 | 只学思路，未复制文字；分级描述和例子为自拟，并标为“工具约定”（2026-10-16 按评审意见统一：增加人手一律算中） | `references/templates/rca-pdca.md` 第三部分第7条；`references/quick-cards.md` 第四部分；评测第2条 |

## 三、权威公开资料（制度依据，非开源项目）

| 资料 | 核实出处 | 用到的内容 | 落到的文件 |
|---|---|---|---|
| 《三级医院评审标准（2025年版）》（国卫医政发〔2025〕4号） | 国家卫生健康委网站；福建省卫生健康委转发页及附件 PDF；深圳市卫生健康委解读 | 两个部分 97 节 232 条；前置要求 4 节 29 条及各节名称；监测指标 5 章名称；评审周期 4 年；甲乙丙符合比例；数据核查规则；二级医院可参照；2022年版废止 | `references/profile-detail.md`；`references/templates/dengji-ziping.md`；HOS-016 |
| 《三级中医医院评审标准（2026年版）》（国中医药医政函〔2026〕39号） | 中国政府网政策库通知全文（2026-10-16 复核） | 2026年3月13日印发；2017年版三级中医医院、中西医结合医院、民族医医院评审标准及实施细则同时废止；中医医院用中医系列标准 | `references/profile-detail.md` 第一、四部分 |
| 《国务院办公厅关于加强三级公立医院绩效考核工作的意见》（国办发〔2019〕4号） | 中国政府网 | 4 个一级指标、55 项三级指标名称、26 项国家监测指标、数据来源 | `references/glossary.md` 第一部分；`references/templates/jixiao-fenxi.md`；HOS-008～013 |
| 国家三级、二级公立医院绩效考核操作手册（2024版）（国卫办医政函〔2024〕87号、86号） | 健康报 2024-03-22 报道；二级手册 86号见法规转载站全文；三级手册 87号见行业文章转述（未取得官方原文，与 86号同批印发） | 三级 4/14/55（定量 50、定性 5）；二级 4/10/28，国家监测指标 21 | `references/profile-detail.md` 第一部分 |
| 关于启动2025年度二级和三级公立医院绩效监测有关工作的通知（2026-07-03） | 健康界转载 | “绩效监测”名称；不含中医医院；病案首页每月15日前上传 | `references/profile-detail.md`；`references/calendar.md`；HOS-015 |
| 中共中央办公厅《关于加强公立医院党的建设工作的意见》（2018年） | 中国政府网（2018-06-25） | 院长办公会议与党委会议的关系原文；不得以党政联席会议代替党委会议；医德“一票否决”制 | `references/templates/yiti.md`；`references/redlines.md` 第五条；HOS-004、005 |
| 《医疗质量管理办法》（国家卫生计生委令第10号） | 中国政府网 | 第十一条、第二十六条、第三十四条、第四十七条 | `references/templates/zhiliang-fenxi.md`；`references/glossary.md` 第三、五部分 |
| 《患者安全专项行动方案（2023—2025年）》（国卫办医政发〔2023〕13号） | 中国政府网；附件2 核对福建省卫生健康委转发件；附件1 核对长治市转发件 PDF（2026-10-16 渲染成图片阅读） | 非惩罚性报告机制（第十条、第十二条）；每百出院人次主动报告不良事件年均大于2.5例次（行动目标原句）；附件1 监测指标名“每百出院人次报告不良事件例数”、计量单位“例次”；不良事件分级分类 | `references/templates/rca-pdca.md`；`references/glossary.md` 第一、四部分；`references/quick-cards.md` 第十部分；HOS-018～020 |
| 《2026年国家医疗质量安全改进目标》 | 国家卫生健康委网站附件；新华社报道 | 10 项目标名称 | `references/profile-detail.md`；`references/calendar.md` |
| 《医疗机构工作人员廉洁从业九项准则》（国卫医发〔2021〕37号） | 新华网（2021-11-18）；地方卫生健康部门转发件 | 九项标题；“九不准”同时废止 | `references/templates/yifeng.md`；`references/quick-cards.md`；HOS-017 |
| 《2026年纠正医药购销领域和医疗服务中不正之风工作要点》（国卫医急函〔2026〕123号） | 新华网、中国新闻网（2026-06-08） | 文号、成文和公布时间 | `references/profile-detail.md`；`references/calendar.md` |
| 按病组和病种分值付费2.0版分组方案（医保办发〔2024〕9号） | 人民日报（2024-07-24）及地方转发件 | 2025年起统一使用 2.0 版 | `references/glossary.md` 第二部分；`references/templates/zhuren-shuzhi.md` |
| 《中华人民共和国广告法》第九条第（三）项、第十六条，《中华人民共和国个人信息保护法》第二十八条 | 中国政府网及法律文本网站 | 最高级用语（第九条，2026-10-16 按评审意见补引）；医疗广告禁止内容（第十六条）；医疗健康信息属敏感个人信息 | `references/redlines.md` 第一、六条；HOS-001～003 |

核实不了、未写入本包的内容见 `references/profile-detail.md` 第六部分（各省细则、当年手册逐项公式、各院议事规则、医德考评办法、院感质控指标最新定义）。《医师法》相关条款本次未核实到原文，未引用。

## 四、自研内容

1. 指标台账的七列格式，尤其是“统计周期”“口径与来源系统”两列，以及卫生健康部门绩效评价口径与医保结算口径分开登记、分开写的规则（`references/data-rules.md`）。
2. 8 类专用文书模板的结构、写法要点、正误对照和“只出框架的情形”（`references/templates/`）。
3. 行业机检 `scripts/hospital_check.py` 与 `scripts/hospital_rules.json`：疗效与绝对化表述（对外材料升级）、党政联席会议决策、议题“三重一大”程序（“尚未提交党委会议”一类否定说法不算已履行）、指数误用百分点（含“CMI为0.94，较上年提高0.05个百分点”这类跨逗号写法）、CMI 带百分号、不规范指标名称、已废止依据、RCA 归咎个人和写入处分（含扣绩效、通报批评），共 20 条规则；单元测试采用必错标本和必过标本（`tests/hospital/`）。这一测试方法沿用本仓库 gongwen-review 的做法。
4. 行业核稿附加清单和没有 Python 时的人工搜索词（`references/review-addon.md`）。
5. 医院类型差异表（三级、二级、中医、专科、高校附属、国企举办、基层医疗卫生机构、卫生健康行政部门）（`references/profile-detail.md` 第四部分）。
6. 整改状态口径不另立一套，沿用本仓库 gongwen-rectify 的 `references/status.md`（含临时状态“未启动”、立行立改的时间界限、没有佐证不算完成），在 `references/templates/zhenggai.md` 第三部分和 `references/quick-cards.md` 第九部分摘要转写。医院检查中常见的“已当场整改、佐证尚未附上”，同样记持续推进，提示佐证补齐后改记立行立改（2026-10-16 按评审意见删去原先“按动作认定为立行立改”的例外，与 rectify_check 一致）。
7. 改进措施“只来自材料”：措施内容、牵头部门、时限只用用户材料，起草者想到的思路放在起草说明“附：可供参考的改进思路”；RCA 的强、中、弱分级用于标注材料中的措施和这份思路（2026-10-16 按评审意见同步到 RCA、质量分析、绩效分析、整改模板和速查卡）。
8. 科室主任述职的现场口头稿由本包一并给出（计划 Q4），时长换算（默认约 200 字/分钟，以目标±10%为准）和口语化要点沿用本仓库 gongwen-speech 的讲稿规则，写入 `references/templates/zhuren-shuzhi.md` 第四部分；述职场合数字用原数、不取整，是本包按考核场合作的调整。

## 五、有意没有采用的内容

1. 2025年版评审标准的逐条条文和各省细则的分值：条文篇幅大，且省细则各不相同，本包只写结构和使用原则，要求用户提供本省细则。
2. 绩效考核 55 项指标的逐项公式和导向：每年手册会调整，本包只写名称和出处，导向按当年手册。
3. “星号指标”作为正式术语：一线说法所指不一，本包只作为不规范说法处理（HOS-014）。
4. 带完整数据的范文：不提供可直接套用的全文，避免示例数据被误用。
5. 自动识别患者真实姓名的规则：在“患者”后直接匹配姓名误报太多（如“患者张口困难”），改由 gongwen_lint 的隐私类检查和人工核对承担。
