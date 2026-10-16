# 借鉴记录：整改材料（gongwen-rectify）

- 模块范围：`skills/gongwen-rectify/` 下的 SKILL.md、README.md、`references/`（status、ledger、plan、writing、sources、units、checklist）、`scripts/rectify_check.py`、`evals/`；共享脚本 `shared/scripts/md2xlsx.py`；单元测试 `tests/rectify/`。
- 共享规则（guardrails、output-format、ai-tone、defensive、jianfu、gbt15835、roles、units）的来源已记在 `docs/借鉴记录/` 下对应文件，本文件不重复。
- 阅读方式：2026-10-16 通过网页在线读取，未克隆、未下载到本仓库。
- 借鉴方式：**依据**（法规、党内法规条文，只作要点整理并注明条款）；**思路**（做法和设计取舍，不涉及文字和代码）。本模块所有例句、评测材料均为自拟的虚构内容。

## 一、法规和官方文件（依据）

| 来源 | 核对页面 | 借鉴内容 | 落到的文件 |
|---|---|---|---|
| 《中国共产党巡视工作条例》（2024年2月修订印发） | 中国政府网 2024 年 2 月发布稿（www.gov.cn/zhengce/202402/content_6932874.htm） | 第二十五条立行立改；第三十一条整改情况公开；第三十三条整改主体责任和第一责任人；第三十四条集中整改6个月、问题清单任务清单责任清单、集中整改进展情况报告的报送对象、集中整改后的常态化整改；第四十六条第（四）项虚假整改问责；第五十条巡察参照巡视 | `references/sources.md` 第一部分；`references/plan.md` 第五部分；`references/writing.md` 第二、六部分；SKILL.md 硬规则、任务卡示例 |
| 《中华人民共和国审计法》（2021年10月23日修正） | 全国人大常委会关于修改审计法的决定（共产党员网 2021-10-23 转载稿） | 第五十二条：按规定时间整改、整改情况报审计机关并报本级政府或主管部门、按规定向社会公布、整改弄虚作假依法追责 | `references/sources.md` 第二部分；`references/plan.md` 第五部分 |
| 审计署向全国人大常委会所作审计整改报告（2024年12月） | 中国新闻网 2024-12-22 报道 | 按立行立改、分阶段整改、持续整改三类提出整改要求 | `references/status.md` 第六部分（区分"整改要求类型"和"整改状态"）；`references/sources.md` 第二部分 |

未能取得原文、因而没有收录具体要求的：中办、国办关于审计查出问题整改长效机制的文件；各地审计整改销号办法；各省巡察工作规定。已在 `references/sources.md` 第五部分写明。

## 二、开源项目（思路）

| 来源 | 许可证 | 读取的内容 | 借鉴内容 | 方式 | 落到的文件 |
|---|---|---|---|---|---|
| [bowenliang123/markdown-exporter](https://github.com/bowenliang123/markdown-exporter) | Apache-2.0 | `SKILL.md`（经 jsDelivr 在线读取，main 分支） | `md_to_xlsx` 每张表各占一个工作表；`--force-text` 默认把单元格存为文本，避免 Excel 改写 | 思路 | `shared/scripts/md2xlsx.py`：`--all-tables` 每表一个工作表；数值转换只做"写法不变"的安全转换，编号类列、前导 0、末尾 0、长数字串一律存文本 |
| GitHub Flavored Markdown 规范（表格扩展） | CC BY-SA 4.0 | 表格一节 | 表头行加分隔行、分隔行列数与表头一致、`\|` 转义、对齐冒号的识别规则 | 思路（按规范自行实现，未复制规范文字） | `shared/scripts/md2xlsx.py` 的 `parse_tables`、`parse_delimiter`；`rectify_check.py` 的台账解析 |

本模块未复制任何外部项目的代码或文字。未参考非商业许可、专有许可或无许可证项目的内容。

## 三、自研内容

1. 状态口径（`references/status.md`）：四种状态加临时"未启动"的认定标准、判定顺序、七种不能升级的情形、完成率算法和统计句写法。"立行立改"的时间界限是工具约定。
2. 台账规范（`references/ledger.md`）：列和填法、子项编号、问题照录的三条规则、三清单视图、更新台账的另存规则。
3. 对账脚本 `scripts/rectify_check.py`：编号识别（含子项汇总、带字母编号）、按段落切分报告、完成说法与计划说法的区分（"将""确保""……前"，"已于……前"）、否定和下属单位部分完成的排除、照录问题表述的剔除、统计句和完成率核对、问题表述弱化检测、与反馈原文的逐字核对、超期判断。
4. `shared/scripts/md2xlsx.py` 的公文表格样式（宋体、表题跨列居中、表头浅灰底、全边框、按内容估算行高、冻结表头、每页重复表头、多列横向打印）和打码文字保护（`138****0000` 不被当作加粗标记）。
5. 七类单位的整改材料调整（`references/units.md`），依据 `shared/profiles/units/` 七份画像整理。

## 四、有意没有采用的内容

1. 把单元格一律存为文本（markdown-exporter 的默认做法）：整改台账、督办清单中有金额和数量，存为数值便于求和；改为只对写法不变的纯数字做转换。
2. 用"基本完成"作为过渡状态：多地整改材料中常见，但容易把未完成写成完成，本模块不收为状态口径，脚本对其报 P1。
3. 内置"最新提法库"和各地整改工作规定：提法和地方规定变化快，按 guardrails 第二条只用用户提供的原文。
