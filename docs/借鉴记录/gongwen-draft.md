# 借鉴记录：法定公文（gongwen-draft）

- 模块范围：`skills/gongwen-draft/` 下的 SKILL.md、README.md、references/（不含 `_shared/`）、scripts/draft_check.py、evals/，以及 `tests/draft/`
- 阅读方式：2026-10-16 通过 GitHub API 和 raw.githubusercontent.com 在线读取，未克隆、未下载到本仓库。"读取版本"为当日默认分支的提交号前 7 位。
- 借鉴方式分三种：**思路**（分类方法、判定原则，不涉及文字）；**改写**（参考原文的规则，用自拟的公文语境表述和例句重写）；**反例**（把来源中的写法当作本 skill 要纠正的毛病，例句自拟）。本模块所有例句、示例材料、评测材料均为自拟，未整段照录任何来源。
- 承接：本模块的文种决策、行文规则检查表和分文种骨架，承接并深化了 v0.1 总入口 `skills/gongwen/references/interim-drafting.md` 第二部分（本仓库自研内容，未修改该文件）。逐节对照结果见第六部分。

## 一、MIT 许可来源

| 来源 | 许可证 | 读取版本 | 借鉴内容 | 方式 | 落到的文件 |
|---|---|---|---|---|---|
| [zhaohui-yang/official-document-drafting](https://github.com/zhaohui-yang/official-document-drafting) | MIT（字体除外，本模块未使用其字体） | 6b4166a | `prompts/doc-types/request-请示/spec.md`：请示中本单位打算用"拟"、不写"决定"；请求事项不能埋在背景里；先办后请是常见错误。`approval-批复/spec.md`：表态前置、逐项回应、一请示一批复、不同意的写明理由。`letter-函/spec.md`：函分商洽、询问、请批、答复四类，答复函先给结论。`circular-通报/spec.md`：未查清的责任和结论不写成定论。`public-notice-通告/spec.md`：通告一般不设主送，时间和空间范围不可遗漏，处罚条款注明依据。`decision-决定/spec.md`：决定分部署、奖惩、变更撤销三类，与决议区分。`notice-通知/spec.md`：印发、批转、转发三个动词不混用，附件全文不抄进正文 | 改写 | `references/doc-types/qingshi.md` 第三部分第4条、第六部分；`references/doc-types/pifu.md` 第三部分；`references/doc-types/han.md` 第一部分、第三部分第4条；`references/doc-types/baogao.md` 第三部分第3条；`references/doc-types/tonggao.md` 第三部分；`references/doc-types/jueding.md` 第二部分；`references/doc-types/tongzhi.md` 第一部分；`scripts/draft_check.py` DRF-011、DRF-041、DRF-043、DRF-090 |
| [gongyu0918-debug/chinese-official-writing-skill](https://github.com/gongyu0918-debug/chinese-official-writing-skill) | MIT | 23e124c | `genre-playbook-request.md`：同一事项的配套内容不另算一事；不把用户的申请改写成另一种性质的请求（如把"拨付"改成"追加"）。`genre-playbook-bulletin.md`：不因文种提高评价或处理强度；只有问题、没有处理决定时以问题和当前状态收束；名单与正文一一对应 | 改写 | `references/doc-types/qingshi.md` 第三部分第4条、第四部分第1条；`references/doc-types/tongbao.md` 第三部分第2、3、5条；`scripts/draft_check.py` DRF-010 的判断口径 |
| [Liuxiangjian-ai/official-document-skill](https://github.com/Liuxiangjian-ai/official-document-skill) | MIT | cbe5f8c | 易混文种辨析（请示与报告、请示与函、批复与函、通知通告公告）；平行文语气平等协商、不用命令口吻；下行文写明要求、责任和时限 | 思路 | `references/decision.md` 第三部分；`references/doc-types/han.md` 第四部分 |
| [KaguraNanaga/official-document-writing-skill](https://github.com/KaguraNanaga/official-document-writing-skill) | MIT | e76a6f5 | `references/document-templates.md` 会议通知模板的三段组织（开会原因、会议安排、会议要求，含报名截止和联系人）；请示模板"紧急事项注明批复时限"；复函先引用来函标题和文号 | 改写 | `references/doc-types/tongzhi.md` 第三部分要素表；`references/doc-types/qingshi.md` 第三部分第6条；`references/doc-types/han.md` 第三部分第5条 |

## 二、只学思路的来源

| 来源 | 许可证 | 读取版本 | 借鉴内容 | 处理方式 | 影响的设计 |
|---|---|---|---|---|---|
| [Rimagination/gongwen-draft](https://github.com/Rimagination/gongwen-draft) | 无许可证 | 559d28a | `references/document-types.md` 的组织方法：先按行文目的匹配文种，再核对行文关系与发文权限，必要时改用其他文种 | 只学思路，未复制任何文字、表格或代码 | SKILL.md 第一步"四问定文种"的顺序（名义、对象、目的、可用文种）；P1 文种"先核发文权限" |

## 三、公开规则与公文实例

以下为规则依据和写法实例，均在线核验（2026-10-16），只引条款号和短句，未整段照录：

| 来源 | 用途 | 落到的文件 |
|---|---|---|
| 《中华人民共和国行政处罚法》（2021年修订）第十六条 | 其他规范性文件不得设定行政处罚 | `references/doc-types/tonggao.md` 第四部分；`references/quick-ref.md` 第九部分 |
| 《中华人民共和国行政强制法》第十条第四款、第十三条、第四十三条第二款（2026-10-16 评审修复轮在维基文库全文核对） | 其他规范性文件不得设定行政强制措施；行政强制执行由法律设定；不得以停止供水、供电、供热、供燃气迫使当事人履行 | 同上；`references/checklist.md` 第十三部分；SKILL.md 第三步 |
| 《国务院办公厅关于加强行政规范性文件制定和监督管理工作的通知》（国办发〔2018〕37号）一、（一） | 行政规范性文件不得违法减损公民、法人和其他组织的合法权益或者增加其义务 | 同上 |
| 《国务院办公厅转发司法部关于审理政府信息公开行政复议案件若干问题指导意见的通知》（国办函〔2021〕132号，中国政府网2022-01-04） | 转发标题去掉本文"关于"、事由与原文种之间不加"的"的写法实例 | `references/doc-types/tongzhi.md` 第四部分；`scripts/draft_check.py` 标题折叠及单元测试 |
| 深圳市政府信息公开目录收录的《广东省人民政府办公厅转发国务院办公厅关于加强政府网站域名管理的通知》（2018-10-19，被转发文件为国办函〔2018〕55号） | 被转发的是通知时两个"的通知"合为一个的写法实例 | 同上 |
| 中共中央办公厅《关于加强公立医院党的建设工作的意见》（中办发〔2018〕35号，中国政府网2018-06-25；第二轮评审修复时核对） | 公立医院不以党政联席会议代替党委会议 | `references/units.md` 第五部分（四） |
| 《党政机关公文处理工作条例》、GB/T 9704-2012、《整治形式主义为基层减负若干规定》、交通运输部和湖南大学签报制度 | 已由 `shared/rules/` 核验整理，本模块直接引用其条款号 | 全模块 |

## 四、自研内容

以下设计在调研的项目中没有找到对应实现，为本模块自研：

1. 文种决策表增加"本单位领导（内设部门报送）"和"社会公众"两行，以及口语说法速判表（`references/decision.md` 第二部分）。
2. 行文规则检查表的"发现问题后的四种处理"（直接改、拆分、标待补、提示确认），以及与 `shared/rules/tiaoli.md` X 编号一一对应（`references/rules-check.md`）。
3. 三类判断由用户定的规则：批复同意与否、请示的本单位倾向、通报的处理决定，没给的标【待补】而不替用户选择（SKILL.md 第三步通用规则第3条）。
4. 单位发文体系差异：党政两套代字按事项归属定名义、校级与部门两级发文、各类单位签报流转（`references/units.md`）。
5. 转发标题折叠算法：追溯转发链、省略中间层次，按计划 D19 给出“本机关转发……”和“关于转发……的通知”两种写法（不传 `--issuer` 时只给后一种）；批转和政府、办公厅（室）转发部门文件的层次保留，法规、规章名称（含"管理办法""实施细则"）补书名号（`scripts/draft_check.py --fold-title`）。
6. 会议通知五项必备要素与只认"与会议相连的日期"的机检口径（`scripts/draft_check.py` DRF-050～DRF-055）。
7. 面向群众文书的法律表述核查表（`references/doc-types/tonggao.md` 第四部分）。

## 五、有意没有采用的内容

1. 各来源模板中的"指导思想""重要意义"等铺垫段落：与《整治形式主义为基层减负若干规定》第2条冲突，本模块默认不写。
2. zhaohui-yang 通告规格中"罚则写明依法给予处罚"的默认条款：本模块要求罚则只写有权机关和用户提供的依据，并对社区、学校、医院改用通知或温馨提示。
3. 各来源中带具体数据的完整范文：本模块只给骨架和短例，避免用户照抄、示例数据被误用。
4. zhaohui-yang 对通报篇幅的参考值（表彰类约1500～2000字）：本模块按工具约定给区间，并以用户或画像要求为准。

## 六、v0.1 过渡起草的迁移对照（W1-DR-3，2026-10-16）

按计划 D1，逐节对照 `skills/gongwen/references/interim-drafting.md` 第一、二、六、七部分与本 skill 的 SKILL.md 和 references。结论：本 skill 已覆盖全部条目，“缺失”0 项，没有需要迁入的内容；第一部分第5条“先提纲还是直接成稿”已被计划 D11 取代，不迁。

| interim-drafting.md | 条目 | 本 skill 中的位置 | 结果 |
|---|---|---|---|
| 一、1 | 保密；纪要、通报先查人名 | SKILL.md 开工检查第1条；`references/doc-types/tongbao.md` 第四部分 | 已有 |
| 一、2 | 只用用户事实，列素材清单；转译模式展示清单 | SKILL.md 第三步通用规则第2条；角色适配·转译模式（素材台账） | 已有 |
| 一、3 | 定篇幅：4000字、5000字上限，医院、学校、国企只作提示，用户指定字数±10% | SKILL.md 第三步通用规则第6条（本轮按 D12 补篇幅不足的处理）；`references/checklist.md` 第十部分 | 已有 |
| 一、4 | 开门见山及正误例 | SKILL.md 第三步通用规则第4条；`references/doc-types/tongzhi.md` 第二部分 | 已有 |
| 一、5 | 教学、审稿模式先提纲后成稿 | SKILL.md 开工检查第7条、角色适配（按 D11 首轮成稿，四种情形才停在提纲） | 不迁（被 D11 取代） |
| 一、6 | 正文后接【起草说明】 | SKILL.md 第五步 | 已有 |
| 二（一） | 四问定文种 | SKILL.md 第一步；`references/decision.md` 第一部分 | 已有 |
| 二（一） | 易混文种七组及条款 | SKILL.md 第一步；`references/decision.md` 第三部分 | 已有 |
| 二（一） | 正误对照（区教育局向区财政局请批用函；院感报告夹带购置请示） | `references/doc-types/han.md` 第六部分、`references/doc-types/baogao.md` 第四部分，改用自拟场景 | 已有（原例与本 skill 评测1、2同场景，有意不照搬） |
| 二（一） | 用户指定文种不当的三步处理 | SKILL.md 第一步；`references/decision.md` 第四部分 | 已有 |
| 二（二） | 行文规则11项，含请示有本单位意见、转报下级请示有倾向性意见 | SKILL.md 第二步；`references/rules-check.md` 第一、二部分；`references/doc-types/qingshi.md` 第三部分第5条、第五部分 | 已有 |
| 二（三） | 请示骨架、结语、附注、正误 | `references/doc-types/qingshi.md`；`references/quick-ref.md` 第二部分 | 已有 |
| 二（三） | 报告分界：一件事的报告归本 skill，全面工作报告、自查报告转出 | `references/doc-types/baogao.md` 第一部分（情况、答复、交办事项、报送四种归本 skill；全面工作报告、落实情况报告、整改和自查报告转出） | 已有 |
| 二（三） | 函、通知（会议要素、印发、转发）、批复、通报、意见、签报 | `references/doc-types/` 对应文件；`references/quick-ref.md` | 已有 |
| 二（三） | 决定、通告、公告可起草；决议、命令、公报、议案只给结构；国企董事会议案按章程 | SKILL.md 适用范围；`references/doc-types/jueding.md`、`tonggao.md`、`gonggao.md`、`p2.md` 第三部分 | 已有 |
| 二（四） | 版头不进正文；成文日期、依据文件缺的标【待补】；不自造简称 | SKILL.md 第三步通用规则第1、2条；`references/units.md` 第六部分；`references/doc-types/tongzhi.md` 第二部分第1条 | 已有 |
| 六、1 | 起草说明的顺序 | SKILL.md 第五步第3条（本轮按 D17 补重新编号） | 已有 |
| 六、2 | 交付文件位置 | SKILL.md 第五步第1条 | 已有 |
| 六、3 | 机检；不能运行按人工清单 | SKILL.md 第四步；`references/checklist.md` | 已有 |
| 六、4 | 成稿后的下一步路线；网页环境交接卡 | SKILL.md 第五步第2、4条；适用范围“转交方式”（本轮按 D14 补未安装兜底） | 已有 |
| 七 | 法定公文自检5项 | SKILL.md 自检清单 | 已有 |
| 三、四、五 | 讲话稿、纪要、信息 | 不迁入本 skill。按关键词核对：gongwen-speech 有200字/分钟换算、1小时上限、主持词、表态发言、不编金句、不设督查考核事项；gongwen-minutes 有记录与纪要之分、未形成结论的事项、出席名单、两委会议记录、前置研究；gongwen-brief 有开头即结论、800字、“全国领先”“首创”、疗效保证、授权同意、重复报送 | 由对应 skill 覆盖 |
