# 借鉴记录：去防御性写作（gongwen-direct）

- 模块范围：`skills/gongwen-direct/` 下的 SKILL.md、README.md、`references/rewrite-guide.md`、`references/certainty-by-doctype.md`、`references/unit-adjust.md`、`evals/`。
- 共享规则 `shared/style/defensive.md` 和 `style_patterns.json` 中 DEF-001～028 的来源，已记在 `docs/借鉴记录/style.md`，本文件不重复，只记本 skill 自有文件的借鉴。
- 阅读方式：2026-10-16 通过 GitHub API 和 raw.githubusercontent.com 在线读取，未克隆、未下载到本仓库。"读取版本"为当日默认分支的根树 SHA 前 7 位。
- 借鉴方式：**思路**（分类方法、判定原则、流程，不涉及文字）；**改写**（参考原规则，用公文语境的自拟例句重写）。本模块所有例句、评测材料均为自拟，未照录任何来源的文字。

## 一、MIT 许可来源

| 来源 | 许可证 | 读取版本 | 读取的文件 | 借鉴内容 | 方式 | 落到的文件 |
|---|---|---|---|---|---|---|
| [hexiecs/talk-normal](https://github.com/hexiecs/talk-normal) | MIT | d89cf32 | `prompt.md` | 第 7 条"比较类问题直接给推荐和简短理由，不写平衡论述"；第 12 条"结尾不附条件式提议和'告诉我 X 我再……'" | 思路、改写 | rewrite-guide.md 第四部分（平衡腔按文种改，请示写倾向）、第八部分（删结尾征询）；certainty-by-doctype.md 第一、五部分（条件式请求改直接请求） |
| [molang163/chinese-natural-reply](https://github.com/molang163/chinese-natural-reply) | MIT | 8f4ea91 | `README.md`、`skills/chinese-natural-reply/SKILL.md` | 防御性回复分类中的"划边界""提前纠正对方没有的误解""汇报没动什么"；"删掉后对方得到的信息不减少就删"的判定；"不确定且影响对方下一步的，说一次；不影响的，不说"；"只改说话方式，不减少该做的核查" | 思路 | rewrite-guide.md 第二部分判断句、第三部分按顺序判断、第八部分第 3 条；SKILL.md 第四步总规则第 4 条和第五步（去对冲不跳过事实核对） |
| [blader/humanizer](https://github.com/blader/humanizer) | MIT | 默认分支 HEAD（当日 API 限流未取得树 SHA；style.md 同日记录为 225a6f3） | `SKILL.md` | §9 限定词叠加（多为修补前文的过度断言，只在有依据且意思需要时保留一个限定）；§22 聊天残留（去掉包装、保留内容）；§23 知识边界免责与猜测填充（写明来源未记载，或删去该句）；"改写后核对没有增删事实"的终检步骤 | 思路、改写 | rewrite-guide.md 第二部分改法第 2、3 条，第三部分第 5 条，第八部分第 1 条；SKILL.md 第五步 |
| [MohamedAbdallah-14/unslop](https://github.com/MohamedAbdallah-14/unslop) | MIT | 29d4360 | `.cursor/skills/unslop/SKILL.md`、`anti-aiisms.md` | 反映真实不确定性的限定词保留；数字、人名、引文、安全警示原样不动；改后逐项回核事实，核不实的打待核标记，不编细节；对冲词"组合出现才报"的分级 | 思路 | rewrite-guide.md 第一部分第 1、4 条，第九部分；SKILL.md 第五步人工核对清单；速查表的级别设定 |

## 二、只学思路的来源

无。本 skill 未参考非商业许可、专有许可或无许可证项目的内容。

## 三、自研内容

以下内容为本项目根据《党政机关公文处理工作条例》、共享规则和单位画像自行编写，没有外部来源：

1. 各文种确定性要求（certainty-by-doctype.md）：请示的倾向性意见与结语、批复表态、部署性文件措辞分层、报告剥离请示事项、讲话稿承诺只改到"力争"、文种不明时的推断顺序。
2. "建议"改写前的三项判断：文种、对象和权限（含《条例》第十六条第（二）项部门对下级政府不提指令性要求）、必选还是可选。
3. 改过头的六种情况、保留清单的六类例子。
4. 单位差异（unit-adjust.md）：依据 `shared/profiles/units/` 七份画像整理。
5. 人工诊断速查表中标"人工"的变体说法，来自对评测材料的模拟机检，用于弥补现有规则的漏检。

## 四、有意没有采用的内容

1. unslop 的 anti-detector 模式（以降低 AI 检测分数为目标的改写步骤）：与红线第六条"不承诺、不宣传规避 AI 检测"冲突，不采用。
2. unslop 提供的英文概率阶梯（unlikely、plausible、probable 等）：公文不用概率措辞表达判断，本 skill 改为"有事实直接陈述、无事实标【待核实】、预期写成'力争'"。
3. talk-normal 第 2 条"禁用一切否定式对举"：属于 AI 句式，归 gongwen-humanize；公文中"严禁""不得"等否定式要求是正当写法。
4. chinese-natural-reply 中"官腔"一类：在聊天场景下是毛病，在公文中多为正当写法，不作为本 skill 的改写对象。
5. chinese-natural-reply 的"安慰套话""随口一问回一大篇""干活时的进度消息"：属于对话场景，与公文改稿无关。
