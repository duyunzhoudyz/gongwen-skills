# 借鉴记录：改稿润色（gongwen-polish）

- 模块范围：`skills/gongwen-polish/`（SKILL.md、references、scripts/polish_stat.py、evals）、`tests/polish/`
- 阅读方式：2026-10-16 通过 GitHub API 和 raw.githubusercontent.com 在线读取默认分支，未克隆、未下载到本仓库。
- 借鉴方式：**思路**（判定原则、流程顺序，不涉及文字和代码）。本模块的规则、对照表、例句和脚本均为自写，未复制任何来源的代码或文字。

## 一、MIT 许可来源

| 来源 | 许可证 | 借鉴内容 | 方式 | 落到的文件 |
|---|---|---|---|---|
| [gongyu0918-debug/chinese-official-writing-skill](https://github.com/gongyu0918-debug/chinese-official-writing-skill) `references/compression-details.md` | MIT（根目录 LICENSE 核对，“Copyright 2026 chinese-official-writing contributors”） | 超限时先合并重复、再压铺垫和修饰；主体、对象、数字、期限、责任等要素压缩时不得丢失；低于下限时以事项为单位整体补足，不零散加词 | 思路 | `references/length.md` 第三至六部分（四级删减、必保内容、扩写手法） |
| 同上 `references/writing-rules.md` | 同上 | 上限低于必保内容时，询问放宽或允许删减哪些要素，不为达标删必保事实，也不以估计宣称达标；不能由结果反推调查、检查等过程；把表述改为“拟”“力争”并不免除事实核对；局部修改时保留修改范围外的原文 | 思路 | SKILL.md 第一步第 4、5 条、第二步“压缩”第 5 条和“扩写”第 2 条；`references/length.md` 第五部分、第六部分（三） |
| 同上 `references/formal-addressing.md` | 同上 | 称谓随行文关系变化；同一篇内自称保持一致；上行文语气克制、不替上级作决定；结尾语与文种功能对应 | 思路 | `references/tone.md` 第一、二部分（措辞对照表和互转做法；惯用语本身取自本仓库 `shared/rules/doc-types.md`） |
| 同上 `scripts/edit_scope.py` | 同上 | 用脚本核对改稿是否只改了计划内的地方；输出文件用独占创建、不覆盖已有文件 | 思路 | `scripts/polish_stat.py` 的 compare 子命令（改为逐段配对，列出修改、移动、拆分、删除、新增和每段多出的文字，供核对改动范围）；body --out 用独占创建写文件，同名时加 -v2，不覆盖 |

## 二、本仓库内的依据

| 资料 | 用途 | 落到的文件 |
|---|---|---|
| `shared/rules/jianfu.md`（《整治形式主义为基层减负若干规定》要点） | 压缩目标的篇幅上限（第2条）、讲话时长（第6条）、调语气时不加码（JF-05、JF-07、JF-12）、部门不向下级党委政府发指令性要求（JF-16） | SKILL.md 第二步；`references/length.md` 第一部分；`references/tone.md` 第二部分 |
| `shared/rules/doc-types.md`（各文种惯用语、文种决策表、易混文种辨析） | 行文方向措辞对照表中的开头语、结语；调语气时确定新文种 | `references/tone.md` 第一、二部分 |
| `shared/style/defensive.md` 第六部分（确定性措辞层级） | 下行文“要、应当、必须、不得”的分层；批注“语气太硬、太软”的改法 | `references/tone.md`、`references/annotations.md` |
| `shared/profiles/profile-template.md` 第二部分（领导偏好卡“历次改稿记录”） | 按批注改稿后追加一行改稿记录的格式和规则 | `references/annotations.md` 第四部分 |
| `skills/gongwen-profile/references/maintenance.md` 第一部分（三）（四）、`skills/gongwen-profile/SKILL.md` 写入门槛 | 改偏好卡前备份到 `.gongwen/历史/`、文件末尾“维护记录”加一行；一个文件放多张卡时按职务取卡；只出现在一份稿件中的偏好标“观察中” | SKILL.md 第二步“按领导批注改稿”第 7 条；`references/annotations.md` 第四部分 |
| `shared/rules/data-expression.md` 第三部分（百分比与百分点） | 转译“涨了3个点”时区分百分点和百分比 | `references/tone.md` 第三部分 |
| `shared/guardrails.md` 第一条第2款 | 遇密级、内部资料标识停止后，三种可以继续的情形 | SKILL.md 开工检查第 1 条 |
| `shared/scripts/fact_diff.py`（含 `--ledger` 模式） | 改稿前后事实比对；合并、补充材料时以拼接的参照文件做集合比对 | SKILL.md 第三步；`references/merge.md` 第五部分 |

## 三、官方依据（2026-10-16 在线核对）

| 依据 | 核对内容 | 落到的文件 |
|---|---|---|
| 《党政机关公文处理工作条例》第八条（[中国政府网](https://www.gov.cn/zwgk/2013-02/22/content_2337704.htm)） | 第（十一）项请示“向上级机关请求指示、批准”；第（十四）项函“不相隶属机关之间商洽工作、询问和答复问题，请求批准和答复审批事项” | `references/tone.md` 第二部分“请示与函” |
| 《党政机关公文处理工作条例》第十六条第（二）项（同上，修复轮次再次核对） | 只有党委、政府的办公厅（室）经授权可以向下级党委、政府行文，其他部门不得向下级党委、政府发布指令性公文或提出指令性要求 | `SKILL.md` 调语气第 5、6 条，`references/tone.md` 第二部分，评测第 18 条 |
| 国务院安委会《安全生产治本攻坚三年行动方案（2024—2026年）》（2024年4月印发，据地方政府网站转发和解读核对） | 行动期限为 2024—2026 年，2026 年为收官之年；“十四五”止于 2025 年，2026 年为“十五五”开局之年 | 评测材料 `evals/files/某县应急管理局2026年安全生产工作汇报.txt`、`某市某局2027年工作要点（送审稿）.txt` |
| 《国务院办公厅关于加强三级公立医院绩效考核工作的意见》（国办发〔2019〕4号，[中国政府网](https://www.gov.cn/zhengce/content/2019-01/30/content_5362266.htm)）附件 | 指标 6 的正式名称为“出院患者四级手术比例”，本模块不再把“四级手术占比”写成官方指标名 | `references/tone.md` 第三部分、`references/units.md` 第五部分 |

## 四、有意没有采用的内容

1. chinese-official-writing-skill 的精确替换 JSON 计划（`edit_scope.py` 按计划逐字生成改稿）：网页提示词版无法执行，模型手写 JSON 计划又容易出错；本模块改为改后再逐段对照，网页版用人工清单核对。
2. 同一来源中“试写稿可补拟具体值”的做法：与本项目红线二“事实只来自用户”冲突，不采用。扩写和按批注补事实时，缺的一律标【待补】。
3. 按“相似度分数”判定改稿质量的做法：分数缺少可溯源的依据。脚本只报告每段文字的来源和改动类型，是否合适由模型按任务和批注判断。
4. 讲话每分钟约 200 字为本项目的经验约定（见 `docs/下一轮工作计划.md` 5.3、`shared/profiles/units/dangzheng.md` 第九部分），不是规定原文，文中均注明“约”，画像另有语速的按画像。
5. Word 修订痕迹不逐处判断接受或拒绝：脚本按“接受全部修订”读取（含 w:ins、不含 w:del），要拒绝某处修订的由用户先在 Word 中处理。docx 读取改为标准库解析 XML，不再依赖 python-docx 的 `Paragraph.text`（它不含修订插入的文字）。
