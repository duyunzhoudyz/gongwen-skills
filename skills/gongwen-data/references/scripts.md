# 脚本用法

两个脚本都在本机运行，不联网，不修改输入文件，输出文件已存在时加 -v2、-v3。读表时先查密级和内部标注（文件名、各工作表名、标题、表头前文字、单元格、表注）：命中的，inspect 只报 DT-SEC-01 和标注位置，不输出标题、表注、统计数和其他问题；merge、count、calc 和 phrase 的表格模式、chart_tool 直接拦截，退出码 2。用户说明使用内网部署模型、已删去相关内容、经本单位保密管理部门确认标识系误加或材料已解密、已公开后，才可加 `--marking-ok` 重新运行，此后 DT-SEC-01 降为 P2。需要 Python 3.9 及以上和 openpyxl（`pip install openpyxl`）；读 CSV 只用标准库；出 PNG 另需 matplotlib（可选）。不支持旧版 .xls，请先另存为 .xlsx。

## 目录

一、data_tool.py inspect：报表体检
二、data_tool.py merge：多表合并
三、data_tool.py calc 和 phrase：计算和转文字
四、data_tool.py count：分类计数
五、chart_tool.py：公文图表
六、退出码和常见报错

## 一、data_tool.py inspect：报表体检

```text
python scripts/data_tool.py inspect <表格> [--sheet 名称 | --all-sheets] [--header-row N] [--json] [--strict] [--marking-ok]
```

- 自动识别标题行、表头（含两级表头、年份表头）、单位行、合计、小计、“其中”、平均行和说明行；识别不准时用 `--header-row` 指定表头行号。
- JSON 主要字段：`stopped`（带密级或内部标注、已停止，此时只有 `marks` 和 DT-SEC-01）、`tables[].title`、`pre_texts`（标题与表头之间的文字，如填报单位、口径说明）、`notes`（表注全文）、`periods`（识别出的时段）、`unit_note`、`source_note`（数据来源）、`columns`（列名、单位、类型：key 关键列、number 数量、rate 比率、avg 平均数、id 编号、period 时段）、`checks`（合计校验：ok 相符、mismatch 不符、uncertain 超出修约容差、rounding 修约误差、skipped 未校验及原因）、`stats`（各数值列的个数、空值数、合计、最小、最大）、`issues`。`pre_texts`、`notes` 已脱敏；表注说明了空值含义的，DT-NUL 的建议中直接引用表注原文。
- 合计校验的修约容差为末位2个单位：在容差内报 P2（DT-SUM-03）；超过容差、但不超过（行数＋1）×0.5个末位单位的报 P1（DT-SUM-06）；再大的报 P0（DT-SUM-01、02）。文本输出另列“未校验：列名（原因）”。
- 问题编号含义见 `references/inspect-checklist.md` 第二部分。个人信息在输出中已脱敏。

## 二、data_tool.py merge：多表合并

```text
python scripts/data_tool.py merge <表格> [<表格>...] [--all-sheets] [--key 列名] [--layout auto|record|indicator]
                                  [--expect 名单] [--out 汇总.xlsx] [--json] [--strict]
```

- record：每行一个单位或事项（如30个科室各一行），按列汇总出一行合计。indicator：每行一个指标（如各子公司的“营业收入、利润总额”快报），按指标汇总。默认自动判断。
- 一个工作簿里每个工作表是一个下级单位的，加 `--all-sheets`。
- `--expect` 给应报单位名单（文件每行一个，或用逗号分隔），输出应报未报和名单外单位。
- 处理规则：同量纲单位按多数表换算并记录（`conversions`）；不同量纲不汇总（DT-MG-05）；时段与多数表不同的表不计入汇总（DT-MG-03）；比率、平均数不相加（DT-MG-08）；有空值或缺列的合计标为不完整（DT-MG-07）。
- 重复报送（DT-MG-09）：同一单位出现在两张以上的表中时，`value` 不含这些行，另列 `value_with_dups`（重复的都计入）和 `value_dedup`（各表报的数相同时，重复单位只计一次）。确认是重复报送的采用 `value_dedup`；各表报的数不同时不给 `value_dedup`，说明中列出各表的数，请填报单位确认。
- `--out` 写出 .xlsx（工作表：汇总、明细、问题清单、单位换算）或 .csv。
- 问题编号：DT-MG-01 某表缺列（P1）；03 时段不同、未计入（P0）；04 未写时段、按多数表处理（P1）；05 量纲不同、未汇总（P0）；06 已换算单位（P1）；07 汇总不完整（P1）；08 比率、平均数未汇总（P2）；09 重复报送（P0，另列含重复的合计和只计一次的合计）；10 关键列名不同（P2）；11 应报未报（P1）；12 名单外单位（P2）。人工合并规则见 `references/inspect-checklist.md` 第三部分。

## 三、data_tool.py calc 和 phrase：计算和转文字

单个计算：

```text
python scripts/data_tool.py phrase --type growth --name 新登记经营主体 --unit 户 --current 4356 --base 4021 --basis 同比 --period 2026年1—6月
```

| --type | 需要的参数 |
|---|---|
| growth、change | --current --base；--basis 同比/环比/比上年；--base-label 上月、2021年等 |
| pp | --current --base（两个比率，可带%） |
| share | --part --total；--part-name --total-name |
| pull | --part-current --part-base --total-base |
| contrib | --part-current --part-base --total-current --total-base |
| completion | --actual --target；可加 --elapsed 6 --span 12 算时间进度 |
| multiple | --current --base |
| cagr | --current --base --years |
| mean、wmean | --values（逗号分隔）；wmean 另加 --weights |
| ratio | --numerator --denominator；--scale 100 表示百分比 |

表格批量计算（每行一个结果，合计行也算）：

```text
python scripts/data_tool.py phrase --type growth --table 表.csv --current-col 新登记户数 --base-col 上年同期户数 --basis 同比 --unit 户 --md
```

- 表格模式的列参数：`--name-col`、`--current-col`、`--base-col`、`--part-col`（share，分母取合计行，合计不符时取明细相加）、`--actual-col`、`--target-col`、`--numerator-col`、`--denominator-col`。
- `calc` 输出结果、算式和提示；`phrase` 另输出 `sentences`（第一句为推荐写法）。`--md` 另输出计算底稿表格行，“来源”列为【待补：表名和行号】，交付前补写。
- 合计、小计行与明细不符（DT-SUM-01、02、06）时，各类型一律改用明细相加数计算：`inputs`、结果、句子和 `--md` 行都用明细相加数，句子末尾加【待核实：××表中合计为X，明细相加为Y】（该标记进正文前删去，写进起草说明“待核实”一行），表中原数只写进 `warnings`。占比、拉动、贡献率的总量同样取明细相加数。
- 未给 `--unit` 时，growth、change、multiple、share、pull、contrib、completion 取所用列的单位（如“本年累计（万元）”取万元）。
- 默认百分比、百分点保留1位小数，`--decimals` 可改。提示（`warnings`）要逐条看：基期为0或负数、基数小、未写基期、合计不符。

## 四、data_tool.py count：分类计数

```text
python scripts/data_tool.py count <表格> --by 事件类型[,事件分级] [--unit 件] [--json]
```

- 按一列或两列统计行数和占比，按件数从多到少排列，每项附算式和可直接引用的句子（“食品安全7件，占35.0%”）。
- 分组列为空的行不计入并提示；各项占比修约后相加不等于100%时提示写表注。
- 不按姓名、证件号等个人信息列分组，遇到时报错退出。

## 五、chart_tool.py：公文图表

```text
python scripts/chart_tool.py <表格> --type bar|barh|line|pie|stacked [--x 分类列] [--y 数值列,数值列]
        [--title 图题] [--unit 单位] [--source 数据来源] [--note 注] [--sort desc] [--top 10]
        [--highlight 本单位] [--out 公文/<目录>/图表-新登记户数.xlsx] [--png] [--json]
```

- 生成 Excel 原生图表：黑白灰配色、数据标签、纵轴从0开始、纵轴标题“单位：……”，表格下方写数据来源。在 Excel、WPS 中可以直接改字、改色，复制进 Word。
- 不同单位的指标、饼图里的负数或比率会报错退出（退出码 2），不会硬画。
- 姓名、手机号等个人信息列不作分类（`--x` 或默认关键列），报错退出（退出码 2）；需要时先用 count 按类别汇总再出图。
- `--png` 在装有 matplotlib 时另出 PNG；没有中文字体时汉字可能显示为方框，提示中会写明。

## 六、退出码和常见报错

| 退出码 | 含义 |
|---|---|
| 0 | 正常 |
| 1 | `--strict` 时存在 P0 问题 |
| 2 | 输入错误：文件不存在、格式不支持、列名不存在、参数缺失、图表类型不适用、缺少 openpyxl；材料带密级或内部标注而未加 `--marking-ok`（DT-SEC-01） |

- “是旧版 .xls 格式”：在 Excel 或 WPS 中另存为 .xlsx。
- “无法识别编码”：用 Excel 另存为“CSV UTF-8”。
- “表中没有名为……的列”：报错信息会列出现有列名，按原样复制。
- Windows 下路径过长（超过260个字符）会读不到文件，把表格放到短一些的目录再运行。
