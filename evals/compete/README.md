# 竞品对标评测

用同一个执行模型、同一批真实场景任务，实际运行本项目和 GitHub 上同类公文 skill，比较交付结果。与 `evals/bench`（装 skill 与不装 skill 对比，用例按本项目约定设计）不同，这里的任务、检查点和评分口径**不依赖本项目的任何约定**，只用熟悉公文写作的人都会认同的通用标准。

## 一、参赛方

见 `contenders.json`。每个参赛方记录仓库、评测时的提交号、许可证、参加的赛道。竞品仓库只在评测目录中临时克隆运行，不复制进本仓库。

| 赛道 | 本项目 | 竞品 | 对照 |
|---|---|---|---|
| W 写作 | 全套 22 个 skill（与用户实际安装一致，由执行模型按 description 自选） | lieflat-gongwen、Liuxiangjian-ai/official-document-skill、KaguraNanaga/official-document-writing-skill、Amer-CN/super-official-writer | 不装 skill |
| F 排版 | 全套 22 个 skill | mizzlelover/gongwen-gbt9704-skill、KaguraNanaga/document-format-skills、cj0103/gbt-9704-2012-skills | 不装 skill |
| H 去 AI 味 | 全套 22 个 skill | larashero3-dotcom/lieflat-less-ai-tone、Liuxiangjian-ai/official-document-skill（其 description 含降 AI 味） | 不装 skill |

## 二、执行

- 执行模型统一为 claude-sonnet-5-5，推理强度 max，Claude Code 子智能体。
- 每个参赛方 × 每道题 × 2 次重复，每次一个独立的匿名工作目录（目录名为哈希）。
- 一个子智能体处理同一参赛方、同一赛道、同一重复轮次的 3～4 道题。
- 任务说明除“已安装的 skill”一段外逐字相同（`templates/`）。只允许读取本参赛方的 skill 目录和本题工作目录。
- 当前日期统一设为 2026 年 10 月 21 日。单轮对话：认为应先提问的，以提问结束。
- 每次运行写 `_回复.md`、`_过程.md`、`_正文/`（文书全文），排版赛道另有交付的 .docx。

## 三、评分

### W 写作

1. **检查点**（每题 5～8 条，通用口径，pass/fail）：文种、要素、与材料一致、缺失信息的处理、计算、不编造等。
2. **五维评分**（各 0～10）：文种与要素、事实忠实、结构逻辑、语言文风、可用性。
3. **排序**：同一题同一轮的全部参赛方输出匿名、随机排序后由评审员整体排名。
4. **编造核查**：`fact_diff` 列出输出中有、输入中没有的数字，评审员逐个判定“编造 / 派生 / 非事实”，另列非数字编造。

### F 排版

1. **中立版式审计**：`docx_audit.py` 按 GB/T 9704-2012 的可测量条款逐项检查（页面、版心、字体、字号、行距、标题层级、页码、版头、版记等），独立编写，不调用本项目的 `check_docx.py`。
2. **内容保真**：交付 docx 的文字与输入文字逐字比对。
3. **字体诚实**：本机缺少标准字体时，是否在文件中静默改写为替代字体、是否告知用户。
4. **视觉评审**：LibreOffice 渲染为页面图片，评审员看图评分和排序（渲染机缺少的字体由 LibreOffice 代替显示，各参赛方条件相同，只用于判断版面布局）。

### H 去 AI 味

1. **事实保持**：输入中的数字、人名、单位名、日期在输出中是否保留；输出新增的数字。
2. **改动幅度**：字符级编辑比例（只报告，不计分）。
3. **AI 痕迹**：同时报告 lieflat-less-ai-tone 自带脚本和本项目 `gongwen_lint` 的计数（两者分别是一个参赛方的“主场工具”，并列报告、不单独计分）。
4. **评审**：AI 味、公文语体保持、事实保持三维评分（0～10）和整体排序。

### 评审员

- Claude 评审：claude-sonnet-5-5 子智能体，max。
- DeepSeek 评审：通过 API 调用，同一份匿名材料按正序、倒序各评一次，用于检查位置偏差。Key 只从仓库根目录 `deepseek.key.local` 读取（已加入 .gitignore），不写入任何文件和日志。
- 两家评审结果分别报告，并给出一致性（Kendall τ）。评审员看不到参赛方名称；输出中可能残留能识别来源的写法（如特定标注格式），这是局限之一。

## 四、局限（预先登记）

1. 执行模型只有一个（Sonnet 5.5）。各 skill 在其他模型上的表现可能不同。
2. 本项目的全套 skill 与单个竞品 skill 比较，比较的是“用户实际安装后的产品”，不是单个文件的质量。
3. 题目由本项目设计，虽然只用通用口径，仍可能偏向本项目擅长的场景。题目覆盖的文种中，部分竞品声明不覆盖（如 lieflat-gongwen 只覆盖 6 类文体），分文种结果单独列出。
4. Claude 评审可能偏好 Claude 风格的文字；DeepSeek 评审用于对照。
5. 每题每方 2 次重复，样本仍然有限。

## 五、文件

| 文件 | 作用 |
|---|---|
| `contenders.json` | 参赛方清单：仓库、提交号、许可证、参加的赛道、`setup` 环境准备命令、`smoke` 烟雾测试、`env_path`/`env_commands`（需要 PATH 替身的命令）、`exclude`（复制时排除的目录）、`identify_terms`（脱敏时替换为“[工具]”的词） |
| `tasks/<题号>/task.json` 与材料 | 题目、用户消息、输入材料（`files/`）、检查点、事实单（`facts`）、可选的 `layout`（F 题版式）；`reference/` 放评审与机检用的参照物，不复制进工作目录 |
| `templates/run-common.md`、`run-ours.md`、`run-single.md`、`run-none.md` | 执行任务说明模板：通用要求逐字相同，三种模板只有“已安装的 skill”一段不同 |
| `templates/judge-W.md`、`judge-F.md`、`judge-H.md` | Claude 评审任务说明模板；其中 `<!--RULES-->` 块同时是 DeepSeek 评审提示词的评审规则 |
| `prep.py` | 复制各参赛方 skill、执行 setup 与烟雾测试、建匿名工作目录、分批生成执行任务说明 |
| `machine.py` | 机检：fact_diff、事实保持、编辑比例、AI 痕迹计数、docx 审计、LibreOffice 渲染、生成脱敏副本 |
| `docx_audit.py` | 中立版式审计（GB/T 9704-2012），独立编写，不读不调用本项目 skill 目录下的任何文件 |
| `judge_prep.py` | 把同一题同一轮的全部输出随机分配字母，生成 Claude 评审任务说明与校验脚本 |
| `deepseek_judge.py` | DeepSeek 评审（W、H 赛道，正序倒序各一次） |
| `aggregate.py` | 解盲与统计，输出原始结果 JSON 和统计表 |
| `common.py` | 以上脚本共用的函数（路径、题目与评测目录读取、评分维度口径、评审结果校验） |
| `shims/` | `zip`、`unzip`、`python3` 三个命令的 Go 替身源码（Windows 默认没有 `zip`，`python3` 指向应用商店占位程序），由 `setup` 命令构建到参赛方副本的 `_shim/` 目录 |

## 六、运行步骤

前置：Python 3.13 加 `python-docx`、`openpyxl`；LibreOffice（`C:\Program Files\LibreOffice\program\soffice.exe`）；`uv`（渲染图片用 `uv run --with pymupdf`）；Node.js 与 npm（Amer-CN、mizzlelover 两个竞品）；Go（构建命令替身）；git。评测目录必须放在仓库之外，建议短路径（如 `D:\bench-compete`），竞品的 `node_modules` 层级较深。

```bash
# 0. 提交本项目的改动（ours 取仓库 HEAD 的提交号；有未提交改动时 prep 会警告）

# 1. 准备：复制 skill、跑 setup 与烟雾测试、建匿名工作目录、生成执行任务说明
python evals/compete/prep.py --out <评测目录> --reps 2
#    竞品克隆目录默认 ~/.gongwen-bench/competitors（不存在的会按 contenders.json 克隆并检出指定提交）；
#    也可以 --competitors-dir 指向已有克隆；--tasks W01,F01 只评部分题；--tracks W,F,H 选赛道；
#    --tasks-dir 换题目目录；--ours-src 换本项目 skill 的来源仓库。

# 2. 运行：对 <评测目录>/run_batches.json 中每个批次各起一个子智能体（Claude Code 的 Agent，
#    model: sonnet，effort: max，并发不超过 10 个），提示词为
#    "请完整阅读任务说明文件 <instruction_path>，然后严格按其中的要求逐个处理请求。"

# 3. 机检：fact_diff、事实保持、编辑比例、AI 痕迹、docx 审计、渲染、脱敏副本
python evals/compete/machine.py --bench <评测目录>

# 4. Claude 评审：随机分配字母，生成评审任务说明；对 judge/judge_batches.json 每项各起一个子智能体
#    （model: sonnet，effort: max），提示词为"请完整阅读任务说明文件 <instruction_path>，然后严格按其中的要求评审。"
python evals/compete/judge_prep.py --bench <评测目录>

# 5. DeepSeek 评审：Key 放在仓库根目录 deepseek.key.local（已加入 .gitignore）；先 --dry-run 看提示词长度
python evals/compete/deepseek_judge.py --bench <评测目录> --dry-run
python evals/compete/deepseek_judge.py --bench <评测目录>

# 6. 汇总：解盲，输出原始结果 JSON 和统计表
python evals/compete/aggregate.py --bench <评测目录> --out <结果.json> --tables <表.md>
```

单个 docx 也可以直接审计：`python evals/compete/docx_audit.py <docx> --layout 红头下行 --ref <原文.txt>`。

注意：

- 第 1 步会重建 `runs/`。评测目录里已有运行产出时 `prep.py` 拒绝重建，确需重建加 `--force`。
- 机检摘要里 H 赛道的两组 AI 痕迹计数写作“统计脚本一”（lieflat-less-ai-tone 自带的 `check-translationese.py`、`check-structure.py`）和“统计脚本二”（本仓库 `gongwen_lint`），摘要中不出现参赛方名称，避免脱敏后反而暴露来源。
- 评审员只许读评审任务说明里列出的文件（`sanitized/`、`render/`、`inputs/`）；`index.json`、`judge/blind.json`、`skills/` 不许读。
- 模型输出有随机性：重跑第 2 步会得到不同样本，重跑第 3～6 步只会重新计算。


## 七、docx_audit 检查项

`docx_audit.py` 只依据 GB/T 9704-2012 的条款，用 lxml 直接解析 docx（样式链、主题字体、自动编号、文本框、页眉页脚都会展开），共 29 项；得分是加权通过率（必需 3、重要 2、次要 1，不适用的项不计）。输出 JSON 里另有 `fonts_used`（按生效的东亚字体统计字符数）、`fonts_nonstandard`、`fidelity`（给了 `--ref` 时，去空白后与原文逐字比对的相似度和前 20 处差异）。

| 级别 | 检查项 |
|---|---|
| 必需 | 01 纸张 A4；02 页边距（上 37、左 28、右 26、下 35 mm）；03a 正文字体仿宋_GB2312；03b 正文三号；06 标题（二号方正小标宋简体居中）；12a 发文机关标志/纪要标志（红色居中小标宋）；12b 发文字号格式；15 字体合规总表 |
| 重要 | 04 行距（每页 22 行）；05 首行缩进两字；07a/07b/07c 一级黑体、二级楷体_GB2312、三四级仿宋_GB2312；08 主送机关顶格；09a～09d 页码（PAGE 域、四号宋体、一字线、单右双左）；09e 信函首页无页码；10 成文日期；11a～11c 附件说明、附件另页、附件标识；12c 发文字号位置；12d 签发人字体；12e 红色分隔线 |
| 次要 | 13a 版记字体字号；13b 版记分隔线；14 正文颜色 |

判定口径：字体名必须与标准中文名完全一致（“仿宋”“FangSong”“方正仿宋_GBK”等不合规），仅 `FangSong_GB2312`、`KaiTi_GB2312`、`SimHei`、`SimSun`、`FZXiaoBiaoSong-B05S` 五个同一字体的系统别名默认视为等同（`--strict-font-names` 关闭）；版式（红头下行、上行、信函、纪要、普通材料、讲话）不同，适用的项不同；信函版式不能自动识别，须用 `--layout` 指定。F 题的版式优先取 `task.json` 的 `layout` 字段，缺失时按题号映射，F06、F07 按 `reference/_原文.txt` 的内容判断。

审计量的是文件结构，遇到少见的画线、定位写法可能漏判（版记分隔线、红色分隔线最常见），所以评审员同时看渲染图；渲染机没有安装标准公文字体，LibreOffice 会用别的字体代替显示，字体是否规范只看审计结果里的字体名。
