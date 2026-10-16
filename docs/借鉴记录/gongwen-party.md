# 借鉴记录：党建与组织生活材料（gongwen-party）

- 模块范围：`skills/gongwen-party/`（SKILL.md、README.md、manifest.json、references、scripts/party_check.py、evals）、`tests/party/`
- 阅读方式：2026-10-16 通过 GitHub API 在线读取外部仓库默认分支；党内法规原文通过中国政府网、共产党员网、人民网的公开网页在线阅读。未克隆、未下载任何仓库或文件到本仓库。
- 借鉴方式：**思路**（判定原则、材料区分方法，不涉及文字和代码）。本模块的规则、例句、脚本均为自写，未复制任何来源的代码或文字。党内法规条款只摘要点并标注条款号，不整段照录。

## 一、MIT 许可来源

| 来源 | 许可证 | 借鉴内容 | 方式 | 落到的文件 |
|---|---|---|---|---|
| [gongyu0918-debug/chinese-official-writing-skill](https://github.com/gongyu0918-debug/chinese-official-writing-skill) 的 `chinese-official-writing/references/party-organizational-life.md` | MIT（GitHub API 核对） | 书记代表支委会报告工作与书记个人抓党建述职分开写；党员个人自评不用支部成绩代替；民主测评结果不直接写成组织评定结论；会议召开、形成整改清单不等于整改完成；个人、班子、支部的问题和措施分别对应；年度主题、对照项目以本次通知为准 | 思路 | `references/frameworks.md` 第三部分；`references/types/baogao.md` 第一部分；`references/types/duizhao.md` 第四部分；`references/writing-rules.md` 第三部分第 5 款 |

## 二、本仓库内部复用（同为 MIT）

| 来源 | 借鉴内容 | 落到的文件 |
|---|---|---|
| `skills/gongwen-review/references/tifa-risk.md` | 提法“只标记、不替换、不凭记忆判断”的处理原则；不内置最新提法库 | `references/writing-rules.md` 第八部分；`references/basis.md` 第七部分 |
| `skills/gongwen-report/references/frameworks.md` | 缺本人事例时交付“提纲＋填写指引（需要的事实、句式）”的模板形式 | `references/frameworks.md` |
| `skills/gongwen-report/SKILL.md` | 素材台账的五列格式和“台账没有的不写”的对账做法 | SKILL.md 第二步、第六步 |

## 三、权威公开资料（规则依据，非开源项目）

| 来源 | 用途 | 落到的文件 |
|---|---|---|
| 《县以上党和国家机关党员领导干部民主生活会若干规定》，中共中央印发，2016-12-23 施行；新华社 2017-01-12 播发全文，人民网转载：https://cpc.people.com.cn/n1/2017/0112/c64094-29019459.html | 第二、五、六、七、八、九、十、十一、十三、十五、十九条：适用对象、主题、对照检查六项基本内容、频次、会前准备（个人发言提纲应当自己动手撰写、说明个人有关事项）、会议程序、批评要求、整改时限、会后报告、参照执行 | `references/basis.md` 第一部分；SKILL.md 开工检查、第四步；`references/types/duizhao.md`、`references/types/baogao.md` |
| 《中国共产党支部工作条例（试行）》，中共中央印发，2018 年 10 月；中国政府网 2018-11-25：https://www.gov.cn/zhengce/2018-11/25/content_5343236.htm | 第十一至十四、十六至十九、二十七条：党员大会和支委会的频次与表决条件、党小组会、党课和主题党日、组织生活会、民主评议党员、谈心谈话、支部书记述职 | `references/basis.md` 第二部分；`references/types/jilu.md` 第四部分；`references/types/zhibu.md`；`references/types/zhutidangri.md` |
| 《中国共产党党员教育管理工作条例》，中共中央印发，2019 年 5 月；中国政府网 2019-05-21：https://www.gov.cn/zhengce/2019-05/21/content_5393554.htm | 第十六条“三会一课”、双重组织生活；第十九条谈心谈话 | `references/basis.md` 第三部分 |
| 《党委（党组）书记抓基层党建工作述职评议考核办法（试行）》，中共中央组织部印发，2019-12-30 施行；共产党员网 2020-01-07：https://www.12371.cn/2020/01/07/ARTI1578382177881722.shtml | 第三、四、五、七条：覆盖范围、述职内容、现场述职与“把自己摆进去”、整改清单 | `references/basis.md` 第四部分；`references/types/baogao.md` 第四部分；`references/frameworks.md` 第三部分 |
| 中办《关于加强公立医院党的建设工作的意见》（2018 年 6 月）；中国政府网 2018-06-25：https://www.gov.cn/zhengce/2018-06/25/content_5301208.htm | 第（八）项内设机构党支部执行“三会一课”等制度；第（十）项党支部书记一般由内设机构负责人中的党员担任 | `references/basis.md` 第五部分；`references/units.md` 第四部分 |
| 共产党员网“开好民主生活会”专题页：https://www.12371.cn/special/khmzshh/ | 核对《若干规定》会前准备四项要求的表述 | `references/basis.md` 第一部分 |
| 共产党员网转载《若干规定》全文：https://news.12371.cn/2017/01/12/ARTI1484218248891973.shtml（2026-10-16 按评审意见复核） | 第二条第二款（经济组织等的党组（党委）成员执行本规定）；第六条末段（受到诫勉谈话的，应当说明整改情况）；第七条第二款（到会人数须达应到会人数的三分之二以上）；第九条（会议方案提前10日报上级党组织审核；第（四）项个人发言提纲应当自己动手撰写，并按规定说明个人有关事项）；第十五条报告内容 | `references/basis.md` 第一部分；SKILL.md 开工检查、结构速查、第七步；`references/types/duizhao.md`、`references/types/jilu.md`、`references/types/baogao.md`、`references/units.md` |

高校、国有企业、中小学的党组织领导体制，引用本仓库 `shared/profiles/units/` 已核实的条目，未另行引用。

## 四、有意没有采用的内容

1. 新闻报道中当年民主生活会的发言重点（如某年度“几个带头”）：属于年度提法，每年变化，不收入任何文件，一律以用户提供的通知原文为准。
2. “党支部书记每年至少讲1次党课”“党课每季度至少1次”：在上述党内法规全文中没有查到对应条文（支部工作条例只规定党委（党组）书记每年至少讲1次党课），可能出自地方或单位实施办法，不作为规则写入，`references/types/zhibu.md` 中保留【待核实】。
3. 高校、医院党支部书记“双带头人”：只查到高校转载件和地方文件，未能核对中央文件原文，不写入。
4. 网上流传的对照检查“范文”和个人发言提纲范本：多为个人作品，来源和许可不明，不参考其文字；本 skill 的示范稿只给结构，事例位置一律占位。
5. 新华社报道中的中央政治局民主生活会做法：属于新闻报道，不作为写作规则依据。
