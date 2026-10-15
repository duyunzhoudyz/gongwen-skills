# 借鉴记录：公文去 AI 味（gongwen-humanize）

- 模块范围：`skills/gongwen-humanize/` 全部自有文件，`shared/scripts/fact_diff.py`，`tests/fact_diff/`。
- 阅读方式：2026-10-16 通过 WebFetch、GitHub API 和 raw.githubusercontent.com 在线读取，未克隆、未下载到本仓库。下表"读取版本"为当日默认分支的提交号。lieflat-less-ai-tone 的 GitHub API license 接口当日返回 403，许可证以仓库中的 LICENSE 文件为准。
- 借鉴方式：**思路**（工作流组织、判定原则，不涉及文字）；**改写**（参考原规则，用公文语境和自拟例句重写）。本模块所有规则文字和例句均为自拟，没有整段照录任何来源。
- 词表层面（AI-001～051、白名单 WL-001～017）的借鉴见 `docs/借鉴记录/style.md`，本文件只记录 gongwen-humanize 自身的借鉴。

## 一、MIT 许可来源

| 来源 | 许可证 | 读取版本 | 借鉴内容 | 方式 | 落到的文件 |
|---|---|---|---|---|---|
| [op7418/Humanizer-zh](https://github.com/op7418/Humanizer-zh) | MIT | f4518a8 | 四步工作流（通读记下要保留的主张和数据、只改真实存在的问题、对照原文查增删和断言强度、重读语气）；"交付前核对"的检查角度（新增细节能否在原文找到出处、"可能"有没有变成"确定"、连接词关系是否保持）；不用自评分、字数下降代替核对 | 思路 | SKILL.md 工作流程第一步、第四步；rewrite-guide.md 第一部分第 4 条、第五部分 |
| [larashero3-dotcom/lieflat-less-ai-tone](https://github.com/larashero3-dotcom/lieflat-less-ai-tone) | MIT | 27d2923 | "没命中规则的句子逐字保留、拿不准保持原文"；"每处改动都能对应到一条规则，对应不上就撤销"；不能作为改写理由的特征（被动句、名词化、长句、"首先……其次"、句内排比）；话题壳、复述句的改法 | 思路、改写 | rewrite-guide.md 第一部分第 2、3、7、8 条，第三部分（五） |
| [LifelongLazyLearner/qu-ai-wei](https://github.com/LifelongLazyLearner/qu-ai-wei) | MIT | 1d32e80 | 改写前先"冻结事实"、建信息账本，改后对照账本复扫；复扫后没有可验证的改善就恢复原文；条目按"触发、保护、动作、复扫"组织；无依据的拔高先写事实再另起评价 | 思路 | SKILL.md 第一步"冻结事实"；rewrite-guide.md 每类"识别、改法、保留、对照"四项结构，第一部分第 7 条 |
| [MrGeDiao/shuorenhua](https://github.com/MrGeDiao/shuorenhua) | MIT | e7c2b86 | 事实保持清单（谁做什么、数字、条件、否定、情态、比较关系、责任归属、判断强度、来源归属不变；原文没写执行者的不替它补）；只标问题、不改原文的标注模式；按改动范围分档；不判断作者身份、不承诺规避检测的表态；应改与不应改两类评测用例的设计 | 思路、改写 | rewrite-guide.md 第一部分"七个不变"；SKILL.md 三种模式、"降 AI 率"应对；evals.json 中保留类用例（2、3、7、14） |

## 二、按公文场景补充的内容（自研）

1. 义务强度（必须、应当、要、可以、原则上、力争）、责任分工、时限列入"不变"清单，这是公文特有的事实。
2. 三种模式与公文岗位对应：只标注不改对应审稿模式，轻度改写按机检 P1、P2 分级。
3. 按文书类型和单位类型的放宽、收紧规则（doc-type-notes.md、unit-notes.md），如医疗"闭环管理"、国企专项行动名称不按新八股处理。
4. `fact_diff.py`：按开发规范 8.1 节自行实现，未参考任何外部代码。

## 三、有意没有采用的内容

1. lieflat-less-ai-tone 第 6 条"连续编号当小标题就删编号"：会破坏公文层次序号。
2. lieflat-less-ai-tone 第 5 条把"如下："当作空提示删除：会误删"现将有关事项通知如下："。
3. lieflat-less-ai-tone 第 2 条"三项以上并列改成不同句法"：公文任务分工本来就是并列。
4. lieflat-less-ai-tone 10.4 建议把句首连接词换成"不过""其实"：口语化，不合公文语体。
5. qu-ai-wei 的"句长节奏过齐"和"任意分层"：会冲击"一是……二是……"和排比小标题。
6. Humanizer-zh 第 31 条"套话收尾"：可能误删"特此通知""妥否，请批示"。
7. 各来源中"强制句长参差""改用第一人称"之类的做法，以及基于博客语料得出的频率阈值。
8. "风格样本优先于规则"：改为"本单位范文只影响语气，不能突破只改表达、不动事实的约束"（unit-notes.md 通用做法第 1 条）。
