# 借鉴记录：基层行业包（gongwen-grassroots）

- 模块范围：`skills/gongwen-grassroots/` 下的 SKILL.md、README.md、references/（不含 `_shared/`）、scripts/notice_check.py、evals/，以及 `tests/grassroots/`。
- 阅读方式：2026-10-16 通过 GitHub API、raw.githubusercontent.com 和政府、媒体网站在线读取，未克隆、未下载到本仓库。"读取版本"为当日默认分支的提交号前 7 位。
- 借鉴方式：**思路**（分类方法、判定原则）；**改写**（参考原文规则，用自拟的基层语境表述和例句重写）；**反例**（把来源中的写法当作要纠正的毛病，例句自拟）。本模块所有例句、示例材料、评测材料均为自拟和虚构，未整段照录任何来源。

## 一、MIT 许可来源

| 来源 | 许可证 | 读取版本 | 借鉴内容 | 方式 | 落到的文件 |
|---|---|---|---|---|---|
| [gongyu0918-debug/chinese-official-writing-skill](https://github.com/gongyu0918-debug/chinese-official-writing-skill) | MIT | 23e124c | `genre-playbook-initiative.md`：倡议邀请自愿参与，处罚、检查、考核、追责不能从倡议中推导，需要执行性管理时改用通知；不擅自扩大参与对象 | 改写 | `references/templates/jumin-wenshu.md` 第五部分；`scripts/notice_check.py` NTC-009 |
| 同上 | MIT | 23e124c | `genre-playbook-open-letter.md`：分清发信方、执行方和受众，"我们"所指要一致；不把内部工作部署照搬给公众；已确定和待定的安排不混写 | 改写 | `references/templates/jumin-wenshu.md` 第五部分第 4 条、第十部分第 7、8 条 |
| 同上 | MIT | 23e124c | `government-work-information.md`：动态、经验、问题建议三类信息的各自写法；数据保留统计期间、比较基准等限定条件 | 改写 | `references/templates/zhengwu-xinxi.md` 第四、六部分 |
| 同上 | MIT | 23e124c | `material-submission.md`：零值、无此事项、待核实分别表述 | 改写 | `references/templates/taizhang.md` 第四部分第 3 条 |

## 二、只学思路的来源

无。v0.1 借鉴记录（`docs/借鉴记录/gongwen-report.md`）中已登记的"先查先核再写""【待补】占位"思路，本模块沿用项目既有做法，未另行读取。

## 三、政策与制度来源

党内法规、法律、行政法规和规范性文件不适用著作权法（著作权法第五条），本模块只做条款化概括，未整段照录。

| 来源 | 发布机关、时间 | 核实内容 | 落到的文件 |
|---|---|---|---|
| 《中华人民共和国城市居民委员会组织法》（修订），https://www.news.cn/politics/20251029/e37962ac66a440fda6b5e7d8f95a6349/c.html | 全国人大常委会，2025-10-28 修订，2026-01-01 施行 | 第十条职责；第二十八条居民代表会议每半年至少一次；第三十条居民会议职权、重要事项先经社区党组织研究讨论；第三十四条居务公开（公开栏、公布期限不少于十五日）；第四十三条协助工作经费、委托事项经费由委托部门承担 | glossary.md、profile-detail.md、templates/wugongkai.md、templates/minsheng-shishi.md、templates/yibiaotong.md、notice_check.py |
| 《关于修改〈中华人民共和国村民委员会组织法〉的决定》，https://www.news.cn/politics/20251029/4b8965e39bab4029917a5ad05c107df6/c.html ；修改后全文转载，https://www.6laws.net/6law/lawcn/中华人民共和国村民委员会组织法.htm | 全国人大常委会，2025-10-28 修改，2026-01-01 施行 | 第二十八条村民会议讨论决定事项、第三款先经村党组织研究讨论；第三十五条村务公开（一般事项至少每季度、财务往来较多的每月、重大事项随时，公布期限不少于十五日）；第四十三条经费 | 同上 |
| 《中国共产党农村基层组织工作条例》，https://www.gov.cn/zhengce/2019-01/10/content_5356764.htm | 中共中央，2019-01 公布，2018-12-28 施行 | 第十九条第三款：村级重大事项决策实行"四议两公开"，条例原文列明四个步骤（村党组织提议、村"两委"会议商议、党员大会审议、村民会议或者村民代表会议决议）和决议公开、实施结果公开 | glossary.md、profile-detail.md、templates/wugongkai.md、templates/minsheng-shishi.md |
| 《关于规范村级组织工作事务、机制牌子和证明事项的意见》，https://www.gov.cn/gongbao/content/2022/content_5708934.htm | 中办、国办，2022-08 | 建立健全准入制和动态调整制度；压减证明事项；不得要求对口挂牌 | glossary.md、templates/yibiaotong.md |
| 民政部等六部门《关于改进和规范基层群众性自治组织出具证明工作的指导意见》及第一批清单（媒体报道），https://www.thepaper.cn/newsDetail_forward_7160362 | 民政部等六部门，2020 | 第一批《不应由基层群众性自治组织出具证明事项清单》20 项 | glossary.md、templates/yibiaotong.md |
| 《关于全面建立乡镇（街道）履行职责事项清单的意见》（新闻报道），https://www.12371.cn/2024/08/30/ARTI1725018545250289.shtml ；编制完成情况，https://www.news.cn/politics/20250730/3d69add6a27f43a7a72ccdc95c96c649/c.html | 中共中央，2024-08；新华社，2025-07 | 三类清单；2025 年 6 月底全部编制完成 | glossary.md、templates/yibiaotong.md |
| 《整治形式主义为基层减负若干规定》，https://www.gov.cn/zhengce/202508/content_7035439.htm | 中办、国办，2024-08 | 第7、8、12、14、16、17条（与 `shared/rules/jianfu.md` 一致）；第15条（不得将村（社区）作为行政执法、城市管理、招商引资等事务的责任主体）为本包另行核实，`shared/rules/jianfu.md` 暂缺；第16条两句分开写：未经省级党委和政府统一部署不得擅自下放清单外职责事项，不得随意以属地管理、签订责任状等方式转嫁任务（后一句没有省级部署的前提） | redlines.md、templates/yingjian.md、templates/yingji-tongzhi.md、templates/yibiaotong.md |
| "一表通"和《基层报表数据"只报一次"工作指引》，https://www.news.cn/politics/20250616/4b7cf5b064e74e3ea9cc2ea5d9624c19/c.html ；https://paper.people.com.cn/rmrb/html/2024-10/29/nw.D110000renmrb_20241029_2-10.htm | 新华社 2025-06；人民日报 2024-10 | 各地依据国务院办公厅指引推进"一表通"；报道提到重庆（"六步工作法"）、浙江、贵州、呼和浩特赛罕区等地做法；地方做法、名称各地不同 | glossary.md、templates/yibiaotong.md、templates/taizhang.md |
| 《中共中央 国务院关于加强基层治理体系和治理能力现代化建设的意见》，https://www.12371.cn/2021/07/12/ARTI1626046486313923.shtml | 中共中央、国务院，2021-04 | 网格化管理；依托村（社区）统一划分综合网格；专职网格员纳入社区工作者管理 | glossary.md |
| 《中华人民共和国突发事件应对法》（修订），https://paper.people.com.cn/rmrbwap/html/2024-07/02/nw.D110000renmrb_20240702_1-14.htm | 全国人大常委会，2024-06 修订，2024-11-01 施行 | 第七条不得编造、故意传播虚假信息；第二十二条、第六十条、第六十四条警报和预警期；第七十七条、第七十九条自救互救和服从指挥安排 | glossary.md、templates/yingji-tongzhi.md、redlines.md |
| 《气象灾害预警信号发布与传播办法》，https://www.gov.cn/gongbao/content/2008/content_912553.htm | 中国气象局令第16号，2007 | 第三条四级颜色；第七条统一发布；第十条不得更改删减、不得传播虚假过时预警 | 同上、notice_check.py |
| 《森林草原防灭火条例》，https://www.news.cn/politics/20251121/3cb4e11ef6254b14a9bbcef6c9bca4c5/c.html | 国务院，2025-11 公布，2026-01-01 施行 | 第十六条居（村）委会协助；第二十五条防火期；第二十七条第一款防火期内禁止在防火区内野外用火（限于防火期内）；第三十条禁火令；第五十四条防火期内擅自在防火区野外用火的处罚主体；第六十三条《森林防火条例》《草原防火条例》同时废止 | glossary.md、templates/yingji-tongzhi.md、notice_check.py NTC-016 |
| 《中华人民共和国行政处罚法》（修订），https://www.12371.cn/2021/01/23/ARTI1611368129158773.shtml | 全国人大常委会，2021-01 修订 | 第二十四条省级可将县级部门部分处罚权交由乡镇、街道行使 | profile-detail.md、templates/jumin-wenshu.md |
| 《最低生活保障审核确认办法》，https://www.gov.cn/zhengce/zhengceku/2021-07/29/content_5628296.htm | 民政部，民发〔2021〕57号，2021-06 | 第二条县级民政部门审核确认，有条件的地方可下放至乡镇（街道）；第十九条初审公示 7 天；第二十四条公布内容和隐私保护 | templates/wugongkai.md、templates/jumin-wenshu.md、notice_check.py |
| 《中华人民共和国民法典》第二百零一条，https://www.spp.gov.cn/spp/fl/202006/t20200602_463888.shtml | 全国人大，2020-05 | 按日计算期间的，开始的当日不计入，自下一日开始计算；用于提示公示期限的起算差异 | templates/wugongkai.md、notice_check.py NTC-019 |
| 《中华人民共和国个人信息保护法》，https://paper.people.com.cn/rmrbwap/html/2021-08/23/nw.D110000renmrb_20210823_1-14.htm | 全国人大常委会，2021-08 | 第六条最小范围、不得过度收集；第二十八条敏感个人信息 | templates/taizhang.md、templates/wangge-zoufang.md |
| 某县政府办公室 2024 年三季度政务信息采用情况通报（政务快报第 47 期），https://yanjin.gov.cn/uploads/ueditor/20241108/97043d93e1bec13e261f25093e2eafdf.pdf | 县级政府办公室，2024-11-01 | 按采用刊物层级计分、领导批示加分、数据失实和抄袭扣分、多级采用只计最高分、乡镇按月报送问题类专报课题 | templates/zhengwu-xinxi.md 第一部分（只写规则类型，不写分值） |
| 民生实事项目人大代表票决制相关报道，https://www.rmzxw.com.cn/c/2018-10-18/2194020.shtml | 人民政协网，2018-10 | 2017 年 3 月起浙江省人大常委会在全省推行；地方做法 | glossary.md、templates/minsheng-shishi.md |

## 四、核实中发现的问题与处理

1. 两部组织法已修改、修订，2026 年 1 月 1 日起施行，条号变化（如村务公开由原第三十条改为第三十五条），并新增"公布期限不少于十五日"。本模块按新文本写；`shared/profiles/units/shequ.md` 只写法律名称、未写条号，内容不冲突，建议主 Agent 在下一轮更新时补充"2025 年修订"和公布期限。
2. 《森林防火条例》《草原防火条例》已于 2026 年 1 月 1 日废止，由《森林草原防灭火条例》替代。仓库共享资料中未发现对旧条例的引用；本模块在 `notice_check.py` 中加了 NTC-016 提示旧条例引用。
3. "一表通"是地方做法，不是全国统一制度名称；《基层报表数据"只报一次"工作指引》的印发时间未从官方文本核实，本模块只写发布机关、不写日期。
4. 《关于全面建立乡镇（街道）履行职责事项清单的意见》未查到全文，三类清单的表述依据新华社、人民日报等报道。
5. 政务信息"积分导向"与减负规定第8条（不得将刊物刊发文章信息作为督查检查考核内容）存在张力。本模块只借采用导向提高信息质量，不写分值，不为积分编数据、拆稿。
6. 顶层设计和下一轮工作计划中的"准入清单"在中央文件中的表述为"准入制"（2022 年中办、国办意见），各地另有社区工作事项准入清单；本模块要求写本地清单全称。
7. "四议两公开"：条例第十九条第三款原文已列明四个步骤和两公开，本模块按条例原文写，只把各步骤的具体操作注明"以当地规定为准"。初稿曾误写为"条例只写实行四议两公开、步骤为通行表述"，并误引另一份文件的链接，评审后已改正。
8. 初稿写"票决制 2008 年起源于浙江宁海乡镇试点""一表通（如重庆、海南）"，所引报道中没有起源地、年份和海南的内容，评审后改为有出处的表述（浙江 2017 年起全省推行；重庆、浙江）。
9. 独立评审指出：初稿把《森林草原防灭火条例》第二十七条写成"防火区内禁止野外用火"，漏了"防火期内"的限定，等于把限时禁令写成常年禁令；减负规定第16条把"未经省级党委和政府统一部署"的前提错套到"不得以签订责任状等方式转嫁"一句。两处已对照原文改正（glossary.md、templates/yingji-tongzhi.md、quick-cards.md 卡四、templates/yibiaotong.md）。

## 五、自研内容

1. `scripts/notice_check.py`：面向群众文书体检（要素、日期与星期、通俗度、倡议书语气、群消息篇幅、村（居）务公开期限与低保初审公示期分开判断、公示中的个人信息、预警来源、已废止法规、村（社区）常见的取消资格和责任书等强制表述 NTC-018、期限恰好卡在临界值的提示 NTC-019），及 `tests/grassroots/` 中的必过、必错标本。
2. 9 类专用文书模板与速查卡；发文主体判断不了时按文书取默认值（居民文书等默认社区）；张贴版与群消息版成对交付。
3. 行业红线、核稿附加清单、行业画像详版、写作日历。

## 六、有意没有采用的内容

1. 各地政务信息计分的具体分值：各地不同且随年度调整，写进模板会误导用户。
2. 自建房、燃气等专项整治的全国统一检查标准：各地方案不同，迎检材料以用户提供的检查清单为准。
3. 带具体降雨量、风力的应急通知范文：预警数据只能来自发布单位原文，模板不提供可套用的数字。

## 七、需主 Agent 处理

1. 解冻后在 `shared/rules/jianfu.md` 第二部分补第15条（明晰基层权责：村（居）委会依法履行职责，不得将村（社区）作为行政执法、城市管理、招商引资等事务的责任主体），并把第五部分第 4 条"与写作无关的条款"中的第15条去掉，然后同步。本包 redlines.md 第六条、templates/yibiaotong.md 已按核实结果引用第15条。
2. 第四阶段边界统一：gongwen-draft 的 description 补"居民通知、告知书、倡议书和乡镇发村（社区）的防汛防火应急通知，装有 gongwen-grassroots 时交它"；`shared/profiles/units/shequ.md` 第十部分 gongwen-draft 一行改为指向本包；gongwen-brief 的转交范围改为"乡镇（街道）、村（社区）报上级的政务信息"，与本包一致。
3. `shared/profiles/units/shequ.md` 补充两部组织法"2025 年修订、修改，2026 年 1 月 1 日起施行"和公布期限不少于十五日（见第四部分第 1 条）。
