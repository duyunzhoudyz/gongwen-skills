# 借鉴记录：中小学行业包（gongwen-k12）

- 模块范围：`skills/gongwen-k12/` 下的 SKILL.md、README.md、references/（不含 `_shared/`）、scripts/k12_check.py、evals/，以及 `tests/k12/`。
- 阅读方式：2026-10-16 通过 GitHub API 和 raw.githubusercontent.com 在线读取，未克隆、未下载到本仓库。"读取版本"为当日默认分支的提交号前 7 位。
- 借鉴方式分三种：**思路**（分类方法、判定原则，不涉及文字）；**改写**（参考原文的规则，用自拟的中小学语境表述和例句重写）；**反例**（把来源中的写法当作本模块要纠正的毛病，例句自拟）。本模块所有例句、示例材料、评测材料均为自拟，未整段照录任何来源。

## 一、MIT、Apache-2.0 许可来源

| 来源 | 许可证 | 读取版本 | 借鉴内容 | 方式 | 落到的文件 |
|---|---|---|---|---|---|
| [PIGU-PPPgu/class-teacher-skills](https://github.com/PIGU-PPPgu/class-teacher-skills) | MIT | fd8487f | `scenarios/03-student-comments/SCENARIO.md`：评语每句对应真实表现，没有进步点时不虚构进步、只保留优点和建议；批量写作可统一风格，但仍基于每名学生的情况。`references/parent_meeting_guide.md`：家长会不给学生排名、不点名批评、不在学生之间比较、不指责家长、不作做不到的承诺，并给家长留出提问时间 | 改写 | `references/templates/xuesheng-pingyu.md` 第三部分第 1、8 条；`references/templates/jiazhanghui-jianghua.md` 第三部分末段、第四部分第 3、6、7 条 |
| [anthropics/k12-teacher-skills](https://github.com/anthropics/k12-teacher-skills) | Apache-2.0 | 31fd6db | `evals/README.md` 的评测细则写法：每条判定标准独立打分，对话行为（是否追问缺失信息）与产出文件分开判定，部分标准只在条件成立时适用 | 思路 | `evals/evals.json` 的 expectations 写法：每条只判一件事，提问行为单列一条，"如计算……""如附回执……"之类的条件式判定 |

## 二、只学思路的来源

本模块没有借鉴无许可证或非商业许可项目的内容。

## 三、政策与制度来源

行业制度、指标、术语来自政府网站、官方媒体和地方教育部门转发的公开文件，2026-10-16 在线核实，逐条列在 `skills/gongwen-k12/references/sources.md`（S01～S23；S23 为评审后增补的 2025 年中小学日常考试管理通知，S12 的督导项数、进校园次数按新京报摘引的原文改为"不超过"），写明文号、年份、用到的要点和核实到哪一级。只能从媒体转述或地方转引核到的，标注"条文以原文为准"；县域优质均衡的指标标准值没有拿到原文附件，本模块不内置标准值。

## 四、自研内容

以下设计在调研的项目中没有找到对应实现，为本模块自研：

1. 学生名单本地脱敏和还原（`scripts/k12_check.py mask`、`unmask`）：模型只读脱敏稿，对照表留在本机，屏幕输出不含姓名。
2. 学生评语批量检查（`comments`）：贬损和标签用语、考试排名（区分比赛名次）、分数、家庭健康处分信息、雷同度两两比较、共用套句、人称、套话占比，以及与事实表对账（缺评语、无事实依据、数字和作品名找不到）。
3. 面向家长文书检查（`parents`）：按双减意见第4、5、9、10条，考试管理通知，教辅和保险自愿购买等规定写成规则，带否定语境排除和学段降级。
4. 十条行业红线（`references/redlines.md`）和核稿附加清单（`references/review-addon.md`），以及"记录只整理不补造""写实记录不代写"的处理办法。
5. 九类专用文书模板，以及教育局作为写作方时的单位差异处理。

## 五、有意没有采用的内容

1. class-teacher-skills 中自动排座位、积分奖励、成绩分析等班级管理功能：不属于公文写作范围，成绩分析交 gongwen-data。
2. 各来源中可以直接套用的评语、家长会发言范文全文：本模块只给结构骨架和短例，避免千人一面和示例被误用。
3. k12-teacher-skills 的教案设计、分层教学：属于教学设计，不属于学校公文和事务性文书。
