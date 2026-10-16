# 借鉴记录：数据整理与图表（gongwen-data）

- 模块范围：`skills/gongwen-data/` 下的 SKILL.md、README.md、references/（不含 `_shared/`）、scripts/data_tool.py、scripts/chart_tool.py、evals/；`shared/rules/data-expression.md`；`tests/data/`。
- 阅读方式：2026-10-16 通过 GitHub API、raw.githubusercontent.com 和国家统计局网站在线读取，未克隆、未下载到本仓库。"读取版本"为当日默认分支的提交号前 7 位。
- 借鉴方式：**思路**（方法、判定原则，不涉及文字和代码）；**改写**（参考来源的规则或启发式，用自己的代码和公文语境的表述重写）；**反例**（把来源的做法当作本模块要避开的做法）。本模块所有代码、例句、示例和评测材料均为自写或自拟，未复制任何来源的代码和成段文字。

## 一、MIT 许可来源

| 来源 | 许可证 | 读取版本 | 借鉴内容 | 方式 | 落到的文件 |
|---|---|---|---|---|---|
| [alchaincyf/huashu-excel](https://github.com/alchaincyf/huashu-excel) | MIT | 5b88312 | SKILL.md 和 README：把表中自带的"合计、小计"行当作免费的校验和，与清洗后的明细对账，不通过就不交付；先读原始单元格再计算（合并单元格只有左上角有值、文本型数字求和被跳过）；口径要声明（同比的分母是哪一期）；小基数的百分比增长、0% 和 100% 先查分母；不画双坐标轴。`scripts/profile_table.py`：表头前说明行、多级表头、"注："等说明行前缀、伪装成数值的空值词表（"—""/""暂无"等）、会计负数"(123)"、千分位、单位后缀、CSV 编码按 utf-8-sig、gb18030 依次尝试、按文件头字节识别旧版 .xls、隐藏行列提示 | 改写、思路 | `scripts/data_tool.py` 的 `parse_cell`、`analyze`、`check_totals`、`inspect_table`；`references/inspect-checklist.md` 第二部分；`references/chart-guide.md` 第一部分；`shared/rules/data-expression.md` 第八部分第 4 条 |
| [hustcc/mcp-echarts](https://github.com/hustcc/mcp-echarts) | MIT | 3ed2496 | README：图表完全在本地生成、不依赖远程服务的设计取向 | 思路 | `scripts/chart_tool.py` 只用 openpyxl 生成 Excel 原生图表，matplotlib 可选；SKILL.md 开工检查第 2 条 |
| [antvis/mcp-server-chart](https://github.com/antvis/mcp-server-chart) | MIT | e2a8fb7 | `src/utils/env.ts`：未设置 `VIS_REQUEST_SERVER` 时默认把图表数据发到服务方的远程渲染接口 | 反例 | SKILL.md 开工检查第 2 条“不用……远程渲染的图表插件”；`references/chart-guide.md` 第三部分第 4 条；README 注意事项第 2 条 |

## 二、公开资料（事实性依据，非开源项目）

以下为国家统计局"统计知识"栏目和法律条文，只引用定义、公式和条款要点，表述为本项目自拟，例句为自拟（第一产业比重例出自原文，标注了出处）。

| 来源 | 引用内容 | 落到的文件 |
|---|---|---|
| 国家统计局《"百分数"与"百分点"》（stats.gov.cn，统计知识） | 百分点是以百分数表示的相对指标的变动幅度；"增长20%"与"降低到20%"的区别；第一产业比重由21.8%降至18.2%应说下降3.6个百分点 | `shared/rules/data-expression.md` 第二、三部分 |
| 国家统计局《"番"与"倍"》 | 一倍与一番都是增加100%；翻两番是增加300%；番按几何级数、倍按算术级数；翻番数公式 | `shared/rules/data-expression.md` 第四部分 |
| 国家统计局《同比与环比有何不同》 | 同比、环比的定义和公式；二者基期不同，不宜直接比较 | `shared/rules/data-expression.md` 第二部分 |
| 国家统计局《什么是经济增长贡献率和拉动点》 | 贡献率＝某项增量÷总量增量；拉动＝贡献率×总量增速，单位为百分点；按不变价计算 | `shared/rules/data-expression.md` 第五部分；`scripts/data_tool.py` 的 pull、contrib |
| 国家统计局《平均发展速度和平均增长速度》 | 平均增长速度由平均发展速度减1得到，一般用水平法（几何平均），不是各期增速的简单平均 | `shared/rules/data-expression.md` 第七部分；`scripts/data_tool.py` 的 cagr |
| 《中华人民共和国民法典》第一千二百五十九条 | "以上""以下""以内""届满"包括本数，"不满""超过""以外"不包括本数 | `shared/rules/data-expression.md` 第八部分 |
| GB/T 15835-2011 | 千分位分节（5.1.1）、万亿与阿拉伯数字连用（5.3）等，已整理在 `shared/rules/gbt15835.md` | `shared/rules/data-expression.md` 第十部分 |
| 《三级公立医院绩效考核指标》（国家卫生健康委，地方转发件，2026-10-16 在线核验） | 出院患者四级手术比例＝出院患者四级手术台次数÷同期出院患者手术台次数×100%，数据来源为病案首页；口径以当年国家操作手册为准 | `references/units.md` 第三部分 |
| WS/T 598.7—2018《卫生统计指标 第7部分：医疗服务》（nhc.gov.cn，2026-10-16 在线核验） | 平均住院日（IHS03.02.002）＝出院者占用总床日数÷同期出院人数 | `references/units.md` 第三部分 |

## 三、自研内容

以下设计在调研的项目中没有找到对应实现，为本模块自研：

1. 数据表述规范的公文写法：百分点与百分数的适用场合、增速变化用"加快、回落"、亏损的"扭亏为盈、减亏、增亏"写法、完成率与时间进度的差距用百分点、"同比"与"比上年同期"不叠用、"环比"不跨年、数据段落写法和正误对照表、10 条可机检规则（`shared/rules/data-expression.md`）。
2. 下级报表合并的口径规则：同量纲换算并记录、不同量纲不相加、时段与多数表不同的表不计入合计、比率和平均数不相加、应报未报和重复报送检查，以及记录式和指标式两种表格布局的自动判断（`scripts/data_tool.py merge`）。
3. 计算结果直接生成公文句子，第一句为推荐写法，基数小时改为先写增减量；指标名以单位结尾时去重（"出院人次6171人次"写成"出院6171人次"）（`scripts/data_tool.py phrase`）。
4. 计算底稿格式（序号、指标、算式、结果、文字表述、来源），与 `fact_diff.py --ledger` 配合做正文对账（`references/calc-sheet.md` 第四部分）。
5. 公文图表规范：黑白灰、数据标签、纵轴从0开始、单位和数据来源必备、合计行不入图、空值不画成0、条形图按表中顺序自上而下，以及网页环境下的图表规格卡（`scripts/chart_tool.py`、`references/chart-guide.md`）。
6. 输出脱敏：体检、合并的输出中姓名改为"张某"、手机号改为"138****0000"一类写法；姓名、手机号列不作图表分类。
7. 密级和内部标注拦截（DT-SEC-01）：读表时查文件名、各工作表、表头前文字、单元格和表注，命中即停止，inspect 只报标注位置；用户说明后加 `--marking-ok` 才继续。
8. 合计不符时的表格模式计算：各类型改用明细相加数，句子和底稿行带【待核实：表中合计为X，明细相加为Y】；修约容差上限为末位2个单位，超出的报 P1。

## 四、有意没有采用的内容

1. huashu-excel 依赖 pandas 做清洗和分析：本模块只用标准库和 openpyxl，便于在内网办公电脑上运行。
2. huashu-excel 的离群值统计（四分位距）和相关性分析：公文数据工作以对账、计算和表述为主，统计推断不在本模块范围内，避免模型据此推测原因。
3. huashu-excel 用另一个 Agent 独立重算的质控环节：本模块改为"算式写进计算底稿＋反算验证＋fact_diff 对账"，网页环境也能执行。
4. mcp-echarts、mcp-server-chart 依赖 Node.js 和浏览器渲染生成彩色交互图：公文需要能黑白打印、能在 Excel 和 WPS 里修改的图，本模块改用 Excel 原生图表。
