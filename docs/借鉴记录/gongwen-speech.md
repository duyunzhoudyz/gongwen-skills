# 借鉴记录：讲话类（gongwen-speech）

- 模块范围：`skills/gongwen-speech/` 下的 SKILL.md、README.md、references/（不含 `_shared/`）、scripts/speech_check.py、evals/，以及 `tests/speech/`
- 阅读方式：2026-10-16 通过 GitHub API 和 raw.githubusercontent.com 在线读取，未克隆、未下载到本仓库。"读取版本"为当日默认分支的提交号前 7 位。许可证于同日通过 GitHub API 的 license 接口核对。
- 借鉴方式分三种：**思路**（分类方法、判定原则，不涉及文字）；**改写**（参考原文的规则，用自拟的公文语境表述和例句重写）；**反例**（把来源中的写法当作本 skill 要纠正的毛病，例句自拟）。本模块所有例句、示例材料、评测材料均为自拟，未整段照录任何来源。

## 一、MIT 许可来源

| 来源 | 许可证 | 读取版本 | 借鉴内容 | 方式 | 落到的文件 |
|---|---|---|---|---|---|
| [gongyu0918-debug/chinese-official-writing-skill](https://github.com/gongyu0918-debug/chinese-official-writing-skill) | MIT | 23e124c | `references/genre-playbook-meeting-host.md`：主持词服务会议流程；会议名称、议程顺序、发言人、表决方式和结果都沿用材料；人数核验、审议、表决只在会议需要且材料给出时才写；现场未形成的结果不写成既成事实；用户只要一段时只交一段；常见易错点（编人数和表决结果、添加不需要的程序、调换议程顺序） | 改写 | `references/types/host.md` 第二、五部分；SKILL.md"主持词、表态发言、交流发言"第 1～3 条 |
| 同上 | MIT | 23e124c | `references/speech-person-order.md`：开场称谓中，具体人物排在泛称之前；同组按已知职务和正副关系排列；受邀主宾可排在本地主官之前，属礼遇安排；无法确定次序的跨单位人物保持材料原序，不凭头衔推断级别；不补充材料中没有的人物；称谓排序不改变主持介绍和发言顺序 | 改写 | `references/types/ceremony.md` 第五部分 |
| 同上 | MIT | 23e124c | `references/genre-playbook-speech-address.md`：讲话内容服从场合、身份、听众和主题；号召和判断要有材料事实支撑；不用排比、口号代替实质内容；不为凑篇幅补写成效、受众、职责；用户只要开场时只交开场 | 改写 | SKILL.md 第五步第 2 条、"素材不足时"；`references/types/leader-speech.md` 第八部分第 3 条；`references/oral-style.md` 第六部分第 5 条 |
| [Liuxiangjian-ai/official-document-skill](https://github.com/Liuxiangjian-ai/official-document-skill) | MIT | cbe5f8c | `SKILL.md` 中表态发言"认识、态度、措施、承诺"的结构，以及"不能只有态度没有措施，每项承诺对应具体动作"；讲话稿"结尾可以升华，但升华要有事实支撑" | 改写 | `references/types/statement.md` 第一部分骨架和第（一）节；`references/types/leader-speech.md` 第一部分 |
| [KaguraNanaga/official-document-writing-skill](https://github.com/KaguraNanaga/official-document-writing-skill) | MIT | e76a6f5 | `references/document-templates.md` 未收讲话类模板；其总结、报告范文以"在……的正确领导下、在……的大力支持下"三段式开头，本模块将这种写法列为讲话开场的常见毛病 | 反例 | `references/types/leader-speech.md` 第一部分第（二）节误例；SKILL.md 第四步第 2 条 |

## 二、只学思路的来源

| 来源 | 许可证 | 读取版本 | 借鉴内容 | 处理方式 | 影响的设计 |
|---|---|---|---|---|---|
| [larashero3-dotcom/lieflat-gongwen](https://github.com/larashero3-dotcom/lieflat-gongwen) | PolyForm Noncommercial 1.0.0 | e0a5aba | `SKILL.md` 中对领导讲话的风格参数思路：力度词集中在"要求"一类，强制词少用；破折号基本不用；引用密度由论证需要决定，涉及重要论述拿不准时标待核对 | 只学思路，未复制任何文字、表格或代码；其 `示例/` 目录下的三篇领导讲话范文未阅读、未使用 | 部署内容用"要"、"必须"少用（沿用 `shared/style/defensive.md`）；引语缺原文标【待核实】（`references/quotes-and-facts.md`） |

## 三、自研内容

以下设计在调研的项目中没有找到对应实现，为本模块自研：

1. "讲的稿、念的稿"的判断表、两种稿子的写法标准，以及讲话版、印发版两版并出和事实一致性校验（`references/oral-style.md` 第一至四部分）。
2. 时长换算表和计数口径（从称谓行算起，不含标题署名和标记文字），超时删减顺序，素材撑不起时按共享规则在回复第一段说明差距并附补料清单（`references/oral-style.md` 第五、六部分）。
3. `scripts/speech_check.py`：按语速估算时长、各部分字数，检查长句、数字过密、段落过长、缺称谓、称谓穿插不足，列出引语清单和来源不明的引文；配单元测试 `tests/speech/`。
4. 议程避重表：工作报告、表彰决定、分管领导部署、交流发言、前一位领导讲话分别怎么避开（`references/agenda-and-preference.md` 第二部分）。
5. 领导偏好卡逐字段落到稿子上的对照表，以及偏好与红线冲突时的处理（`references/agenda-and-preference.md` 第四部分）。
6. 引语、提法、名言古诗文、事例、数据、讲话人本人的经历和见闻六类内容的来源要求和缺失写法，可选概括句放在起草说明、由用户挑选的机制，引语核对清单（`references/quotes-and-facts.md`）。
7. 表态发言的权限边界：不替上级和其他部门承诺，不承诺资金、编制、政策，结果目标只用上级下达或用户确认的（`references/types/statement.md` 第一部分第（二）节）。
8. 讲话人身份差别表和 7 类单位的调整项（`references/types/leader-speech.md` 第七部分、`references/units.md`）。
9. 代表发言的规则：以发言人本人经历为主体，没有本人素材时按 `shared/guardrails.md` 第六条第 4 款只出提纲和填写指引（`references/types/statement.md` 第三部分）。
10. 督查考核事项、超过 1 小时、上市公司未披露信息三种"用户坚持"情形的确定处理（SKILL.md 开工检查第 1 条、第四步第 6 条、第五步第 3 条）。
11. `speech_check.py` 第二版修正：识别"所指出的：""强调的那样，""重要论述："等引法；长礼宾称谓行不按整句检查；"1—9月"一类日期范围不计入数字；文字摘要区分"篇幅不足"和"篇幅超出"。
12. 第二轮评审后的修正：部署任务的责任主体、完成时限、标准和主持词小结的贯彻要求只用要点清单，缺的标【待补】；按领导意思展开时不新增任务、时限和标准；带具体数字、时限、责任人的正例注明"由材料提供"；执法监管挂牌督办与变相考核分开处理；`speech_check.py` 识别"强调，要“……”"式的部分引用。
13. 第四阶段（W2）修正：按计划 D11 删去“教学模式、8 分钟以上先交提纲”的闸门，任何模式首轮成稿，只有四种情形停在提纲；按 D12 统一篇幅规则；按计划第七节补齐转交（述职含现场口头稿交 gongwen-report，医院科室主任述职交 gongwen-hospital，发布词交 gongwen-brief）；fact_diff 已单列序数和概数，删去原先的人工忽略说明。

说明：独立评审后，SKILL.md、references/ 和单元测试中的正误例句全部改用与评测材料无关的场景（冬季供暖、森林防火、校园食堂等），避免评测时提前看到答案；第二轮评审后，`references/types/host.md` 中与评测 4 相近的承接语例句也已替换。

## 四、有意没有采用的内容

1. gongyu0918 讲话页在材料不足时"按原级别讲清该事项的目的、意义"的做法：与《整治形式主义为基层减负若干规定》第2条开门见山的要求冲突，本模块素材不足时交提纲或框架稿，不写意义铺垫。
2. Liuxiangjian-ai 讲话稿结构中"先阐述对形势的认识"：本模块把认识部分限制在全文五分之一以内，不写宏大背景。
3. Liuxiangjian-ai 的内部打分机制（90 分以上直接输出等）：分数无法复核，本模块改用机检脚本加人工清单。
4. 各来源中带具体数据的完整讲话范文：本模块不提供可直接套用的范文全文，只给骨架和短例，避免照抄和示例数据被误用。
