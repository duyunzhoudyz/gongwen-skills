# 借鉴记录：gongwen-format（国标排版）

- 模块：`skills/gongwen-format/`（scripts/、assets/gbt9704_params.json、references/）及 `tests/format/`
- 核验日期：2026-10-16
- 阅读方式：通过 GitHub 网页和 GitHub API 在线阅读，未克隆、未下载到本仓库。
- 复制情况：未复制外部项目的代码或文字。下表借鉴的是参数取值、实现思路和测试方法，代码全部用 Python（python-docx）重写。
- 专有许可：anthropics/skills 中的 docx skill 为专有许可，本模块开发中未阅读、未参考。

## 一、借鉴来源

| 来源 | 许可证 | 读过的文件 | 借鉴内容 | 落到文件 |
|---|---|---|---|---|
| [mizzlelover/gongwen-gbt9704-skill](https://github.com/mizzlelover/gongwen-gbt9704-skill)（main） | MIT，Copyright (c) 2026 mizzlelover | README.md、SKILL.md、LICENSE、references/gbt9704-2012-summary.md、references/gbt9704-audit-matrix.md、references/gbt9704-visual-audit.md、scripts/generate_gongwen_docx.mjs、scripts/verify_gongwen_docx.mjs | 页边距 twips 取值作交叉核对；rFonts 四个属性加 hint=eastAsia、lang eastAsia=zh-CN 的写法；页码单页右缩进、双页左缩进各 280 twips 并开 evenAndOddHeaders；版记用 tblpYSpec=bottom 浮动表格贴版心底边，线宽 0.35/0.25 mm；版头红线约 0.5 mm；页脚距边界 20 mm；缺小标宋时告警 | scripts/gw_docx.py、scripts/gw_fonts.py、assets/gbt9704_params.json |
| [KaguraNanaga/document-format-skills](https://github.com/KaguraNanaga/document-format-skills)（main） | MIT，Copyright (c) 2025 Gloria LaRoche | SKILL.md、README_CN.md、LICENSE、scripts/formatter.py、scripts/process.py、scripts/analyzer.py、scripts/punctuation.py | 用正则识别各级标题；主送机关按“冒号结尾、长度上限”识别；标题只在开头几段中找；落款在末尾若干段中找；附件说明用悬挂缩进；结构性空行保留、多余空段压到 1 磅而不删除；同时写 firstLineChars 和 twips | scripts/gw_layout.py、scripts/gw_format.py、scripts/fix_docx.py |
| [imhuhe/gongwen-kit](https://github.com/imhuhe/gongwen-kit)（main） | MIT，Copyright (c) 2026 HUANGHE | README.md、package.json、LICENSE、tests/run.sh、tests/negative.js、tests/fix.js、tests/messy-docx.js、checks/check-doc.js、lib/gongwen-core.js、lib/fix/comments.js、lib/fix/revisions.js、docs/字体.md、docs/版式.md | 标本测试法：每条规则造一份最小“必错标本”，断言 --strict 退出码和规则号；合规样稿为“必过标本”；合规样稿再修一遍零改动（幂等）；乱稿标本。文档网格 linesAndChars 579/-842；页码位置参考其实测的页脚距边界 1392 twips | tests/format/test_check.py、tests/format/test_fix_annotate.py、tests/format/helpers.py、assets/gbt9704_params.json |
| [linhut/gongwen-skill](https://github.com/linhut/gongwen-skill)（master） | MIT，Copyright (c) 2026 Jose AI | SKILL.md、LICENSE、engine/core/document/tracked_common.py、engine/core/document/tracked_changes.py、engine/core/document/tracked_annotator.py、engine/core/document/annotator.py、engine/core/document/reviewer_comments.py、engine/core/rules/checker.py、rules/official/_common.yaml | P0/P1/P2 的分级口径（纸张、边距、字体字号为 P0；行距、缩进、页码为 P1；空行等为 P2）；批注锚点须为段落直接子元素；批注作者按工具命名；对定位失败的意见单独处理 | scripts/check_docx.py、scripts/annotate_docx.py、references/check-items.md |

## 二、没有采用的做法

| 来源 | 做法 | 原因 |
|---|---|---|
| mizzlelover/gongwen-gbt9704-skill | 发文机关标志用 framePr 定位，红线用 VML 矩形 | 改用段前距和段落边框，WPS、Word、LibreOffice 都能稳定显示，check_docx 也容易回读 |
| mizzlelover/gongwen-gbt9704-skill | 版记分隔线用红色 | GB/T 9704-2012 只规定版头分隔线为红色，版记线按黑色处理 |
| mizzlelover/gongwen-gbt9704-skill | 不设 charSpace | 改为设置字符网格 -842，保证每行 28 字 |
| KaguraNanaga/document-format-skills | 用 rPrChange、pPrChange 把格式修改记为修订；强制 pStyle=Normal | 本模块输出新文件并附修复报告；保留原样式，只加直接格式和大纲级别 |
| KaguraNanaga/document-format-skills | 三级标题默认加粗 | 国标未要求，默认不加粗，单位要求时用“三级标题加粗”开关 |
| linhut/gongwen-skill | 页边距 2.8/2.8/2.7/2.7 cm、行距 33 磅 | 与 GB/T 9704-2012 不符 |
| linhut/gongwen-skill | 定位不到时挂到前后段落；批注锚定整段 | 本模块按 excerpt 精确拆分 run 锚定；定位不到的只列入报告，不写入文档 |
| gongwen-kit、linhut/gongwen-skill | 写修订（tracked changes） | v0.1 只写批注，修订留待后续版本 |

## 三、实践值来源

| 参数 | 取值 | 来源与处理 |
|---|---|---|
| 固定行距 | 28.9 磅 | `shared/rules/gbt9704.md` 第九部分（实践值），用版心高度推算验证 |
| 文档网格 | linePitch 579、charSpace -842 | 按版心尺寸推算，与 gongwen-kit 取值一致 |
| 页脚距边界、页码段后 | 20 mm、13 磅 | 页脚距边界取信函下双线位置（10.1）；段后距使页码下沿距纸张下边约 24.6 mm，与 gongwen-kit 实测值一致；LibreOffice 导出 PDF 测得一字线距版心下边缘约 8 mm |
| 版头红线宽 | 1.5 磅（约 0.5 mm） | 参照 mizzlelover/gongwen-gbt9704-skill |
| 发文机关标志字号 | 最大小初 36 磅，按版心宽 90% 自动缩小 | 实践值，标准只要求醒目、美观、庄重 |
| 正文表格 | 仿宋四号（可选宋体、小四），表头加粗，全框线 0.5 磅，与版心同宽，单元格行距为字号的 1.5 倍、不对齐文档网格 | 实践值，GB/T 9704-2012 未规定表格字号；按计划第三节 Q9 的公文表格样式 |
| 表题 | 黑体三号，居中，与表格同页 | 实践值 |
| 纪要出席名单前空行 | 一行（28.9 磅） | GB/T 9704-2012 10.3 |

## 四、v1.0 新增内容的来源（第四阶段）

| 内容 | 来源 | 落到文件 |
|---|---|---|
| Markdown 管道表格转 Word 表格（计划 Q9） | 自行实现。识别口径（表头行、分隔行、未转义竖线拆格、`<br>` 换行、表题判断）与本仓库 `shared/scripts/md2xlsx.py` 一致；未阅读、未复制外部项目代码 | scripts/gw_layout.py、scripts/gw_docx.py、scripts/build_docx.py、scripts/check_docx.py、scripts/fix_docx.py、tests/format/test_table.py |
| `--allow` 按本单位制度保留 | 与本仓库 `shared/scripts/gongwen_lint.py` 的 `--allow` 口径一致（降为 P2、注明“按本单位制度保留”、不计入 `--strict`） | scripts/check_docx.py、tests/format/test_eval_fixes.py |
| 文号序号占位、纪要出席空行、层次标题与下段同页 | v0.1 评测待修问题 FM-1、FM-4、FM-5（`evals/results/v0.1-待修问题.md`） | scripts/gw_layout.py、scripts/build_docx.py、scripts/fix_docx.py、scripts/gw_format.py、scripts/check_docx.py（序号占位时仍查括号和“第”字）、tests/format/test_eval_fixes.py |
