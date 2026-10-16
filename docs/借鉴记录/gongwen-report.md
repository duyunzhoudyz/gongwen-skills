# 借鉴记录：综合文稿（gongwen-report）

- 模块范围：`skills/gongwen-report/` 下的 SKILL.md、README.md、references/（不含 `_shared/`）、evals/
- 阅读方式：2026-10-16 通过 GitHub API 和 raw.githubusercontent.com 在线读取，未克隆、未下载到本仓库。"读取版本"为当日默认分支的提交号前 7 位。
- 借鉴方式分三种：**思路**（分类方法、判定原则，不涉及文字）；**改写**（参考原文的规则，用自拟的公文语境表述和例句重写）；**反例**（把来源中的写法当作本 skill 要纠正的毛病，例句自拟）。本模块所有例句、示例材料、评测材料均为自拟，未整段照录任何来源。

## 一、MIT 许可来源

| 来源 | 许可证 | 读取版本 | 借鉴内容 | 方式 | 落到的文件 |
|---|---|---|---|---|---|
| [gongyu0918-debug/chinese-official-writing-skill](https://github.com/gongyu0918-debug/chinese-official-writing-skill) | MIT | 23e124c | `genre-playbook-work-summary.md`：系统部署、流程衔接完成只能证明事项已完成，有运行数据或反馈才写成效；`genre-playbook-duty-report.md`：述职区分个人作用、班子集体贡献和单位整体工作；`genre-playbook-work-priorities.md`：要点中保留"既定、拟议、待评估"各自状态；`genre-playbook-research.md`：数据、直接观察和分析判断分开写，不把有限样本写成普遍结论；`genre-playbook-research-outline.md`：调研提纲不预设受访者存在某种困难或成效 | 改写 | SKILL.md 第四步第 3、7 条；`references/types/shuzhi.md` 第三部分第 3 条；`references/types/yaodian.md` 第三部分第 5 条；`references/types/diaoyan.md` 第三部分第 3、4 条；`references/frameworks.md` 第二部分 |
| [KaguraNanaga/official-document-writing-skill](https://github.com/KaguraNanaga/official-document-writing-skill) | MIT | e76a6f5 | `references/document-templates.md` 总结模板中"经验提炼四步法"（罗列分类、归纳亮点、划分层次、严谨行文）的工作顺序；`checklists/quality-checklist.md` 总结类检查项（数据准确、问题准确、打算可行）。其总结模板以"在……的正确领导下……取得了显著成绩"开头、问题部分用"与上级要求和先进单位相比还有一定差距"引出，本模块将这两种写法列为常见毛病 | 思路、反例 | 素材台账"先拆条、再归类、再搭提纲"的流程；`references/types/zongjie.md` 第五部分第 3、7 条；SKILL.md 第四步第 1 条 |
| [Liuxiangjian-ai/official-document-skill](https://github.com/Liuxiangjian-ai/official-document-skill) | MIT | cbe5f8c | 方案与计划的区分（方案重在如何组织实施专项任务，计划重在一段时间内做什么）；汇报材料突出可供决策的信息；调研报告每条建议回应一个问题 | 改写 | `references/types/jihua.md` 第一部分；`references/types/fangan.md` 第一部分；`references/types/huibao.md` 第三部分；`references/types/diaoyan.md` 第三部分第 5 条 |

## 二、只学思路的来源

| 来源 | 许可证 | 读取版本 | 借鉴内容 | 处理方式 | 影响的设计 |
|---|---|---|---|---|---|
| [Rimagination/gongwen-draft](https://github.com/Rimagination/gongwen-draft) | 无许可证 | 559d28a | `references/material-workflow.md` 的"先查先核再写"：起草前把材料整理成可审计的底账，区分已确认事实和待核实事项；不把计划写成已完成、不把建议写成决定；来源冲突时不悄悄合并 | 只学思路，未复制任何文字、表格或代码 | 素材台账的"已确认、待核实"两种状态；SKILL.md 第四步第 7 条"状态不升级"；汇总模式"比不齐的不加总"和口径冲突表 |
| [larashero3-dotcom/lieflat-gongwen](https://github.com/larashero3-dotcom/lieflat-gongwen) | PolyForm Noncommercial 1.0.0 | e0a5aba | 缺失的数字、文号、人名、日期用【待补：xx】占位，补齐后方可定稿的机制 | 只学思路，未复制任何文字、表格或代码；本项目的标记格式以 `shared/style/output-format.md` 为准 | 正文【待补】【待核实】标记与起草说明逐条汇总 |

## 三、自研内容

以下设计在调研的项目中没有找到对应实现，为本模块自研：

1. 素材台账的固定五列（序号、类别、内容、来源、状态）与五个类别，以及"正文与台账对账"的逐项流程（`references/ledger.md`）。
2. 下级材料汇总的口径比对（统计口径、时间范围、计量单位、数值矛盾、归属重复）和口径冲突表（`references/merge-and-versions.md` 第一部分）。
3. 一份台账多版输出的视角转换规则和"本人角色"标记（`references/merge-and-versions.md` 第二部分）。
4. 述职缺事例、调研缺记录、素材为零三种情形的框架交付模板（`references/frameworks.md`）。
5. 8 类材料按单位类型的调整项（`references/units.md`），以及装有行业包时专用文书的转交口径。
6. 大会工作报告（职代会、教代会、党代会）的写法要点（`references/types/zongjie.md` 第七部分，v1.0 第四阶段新增）。

## 四、有意没有采用的内容

1. KaguraNanaga 和 Liuxiangjian-ai 的总结、方案模板中包含"形势背景""指导思想"等铺垫段落：与《整治形式主义为基层减负若干规定》第2条的开门见山要求冲突，本模块默认不写。
2. 各来源中带具体数据的完整范文（如年度工作总结示例）：本模块不提供可直接套用的范文全文，只给结构骨架和短例，避免用户照抄、避免示例数据被误用。
3. gongyu0918 各文体页中的试写样稿约定（未给篇幅时按约 400 字出样稿）：本模块按材料类型给默认篇幅，不出试写样稿。
