# 借鉴记录：单位画像与文风学习（gongwen-profile）

- 模块范围：`skills/gongwen-profile/` 下的 SKILL.md、README.md、manifest.json、references/（不含 `_shared/`）、scripts/style_profile.py、scripts/draft_diff.py、scripts/secret_marks.py、evals/；`tests/profile/`
- 阅读方式：2026-10-16 通过 GitHub API 和 raw.githubusercontent.com 在线读取，未克隆、未下载到本仓库。"读取版本"为当日默认分支的提交号前 7 位。
- 借鉴方式分三种：**思路**（方法和判定原则，不涉及文字和代码）；**改写**（参考原文规则，用自拟表述重写）；**反例**（把来源中的做法当作本模块要避免的情形）。本模块所有文字、例句、示例材料、评测材料、代码均为自拟，未复制任何来源的文字或代码。

## 一、MIT 许可来源

| 来源 | 许可证 | 读取版本 | 借鉴内容 | 方式 | 落到的文件 |
|---|---|---|---|---|---|
| [JimLiu/baoyu-skills](https://github.com/JimLiu/baoyu-skills) | MIT | 1567581 | `docs/creating-skills.md` 中 EXTEND.md 的机制：项目级配置优先于用户级配置；skill 开工时先查找并读取配置文件，读到后向用户简要报告 | 思路 | SKILL.md"四个文件"读写约定第 1 条（工作目录 `.gongwen/` 优先于 `~/.gongwen/`）、开工检查第 2 条（读到什么先用一行报告）；`references/maintenance.md` 第一部分（一） |

## 二、只学思路的来源

| 来源 | 许可证 | 读取版本 | 借鉴内容 | 处理方式 | 影响的设计 |
|---|---|---|---|---|---|
| [larashero3-dotcom/lieflat-gongwen](https://github.com/larashero3-dotcom/lieflat-gongwen) | PolyForm Noncommercial 1.0.0（仓库 LICENSE 文件；GitHub 显示为 Other） | e0a5aba | `参数卡.md` 的"风格参数化"方法：把文风写成可测量的参数（篇幅、句长、各级标题数、"一是二是"、顿号密度、强制词和建议词、开头和结尾方式）；参考范围用中位数和四分位范围而不用平均数；参数只用来发现明显冲突，不要求把稿子凑成平均值 | 只学思路，未复制任何文字、参数取值、表格或代码；本模块的参数从用户自己的历史稿计算，不内置任何语料统计值 | `scripts/style_profile.py` 的统计项和"中位数（P25～P75）"汇总；`references/style-params.md` 第一部分模板和第五部分末段"参数是参考范围" |
| [jzOcb/writing-style-skill](https://github.com/jzOcb/writing-style-skill) | 无许可证 | d94826e | 用"初稿与定稿"两端比对提取改稿规则；规则按出现次数和人工确认分级，只出现一次的先观察；更新前备份、可回滚 | 只学思路，未复制任何文字或代码 | 花脸稿学习的写入门槛（2 份以上稿件写入，1 份只进改稿记录"观察中"，用户确认可提前写入）；读写约定第 4 条"改前备份到 `.gongwen/历史/`" |

## 三、规范与标准

| 来源 | 用途 | 落到的文件 |
|---|---|---|
| ECMA-376 Office Open XML（WordprocessingML）中修订标记的公开定义：`w:ins`、`w:del`、`w:delText`、`w:moveFrom`、`w:moveTo` 及其 `w:author`、`w:date` 属性，批注 `word/comments.xml`、`w:commentRangeStart`、`w:commentRangeEnd` | 按标准自行实现修订还原和批注读取；嵌套修订（`w:ins` 内含 `w:del`）按"祖先已隐藏则子节点隐藏"处理；页眉页脚（`word/header*.xml`、`word/footer*.xml`）只用于查标注。调研时看到的现成 Python 实现为 AGPL-3.0（scitex-msword）或商业服务，均未参考其代码 | `scripts/draft_diff.py` 的 `DocxRevisions`；`scripts/style_profile.py` 的 docx 读取；`scripts/secret_marks.py` 的 `docx_mark_lines` |

## 四、自研内容

以下设计在调研的项目中没有找到对应实现，为本模块自研：

1. 四个画像文件的读写约定：先读后写、变更单、改前备份、维护记录、一个文件放多张领导偏好卡（SKILL.md"四个文件"，`references/maintenance.md` 第一部分）。
2. 从本单位公文处理办法提取画像字段的"字段与条文"对照表和提取规则，含国标差异括注、旧版办法标【待核实】（`references/intake.md` 第二部分）。
3. 分轮问卷：每轮不超过 3 问、带单位类型选项、可跳过（`references/intake.md` 第一、三部分）。
4. 文风学习的"不采纳的习惯"：历史稿中与本项目文风规则冲突的写法单列，不作为模仿对象；政治性提法不进常用表述（`references/style-params.md` 第四部分）。
5. 花脸稿改动分类表，以及事实更正、错别字、一次性取舍、对人评价四类不归纳为偏好（`references/edit-patterns.md` 第二部分）。
6. `style_profile.py`：公文层次标题识别、标题句式分类、开头和结尾方式分类、四字格左右邻字多样性过滤、隐私命中只计数并在统计前屏蔽、密级标注整篇跳过。
7. `draft_diff.py`：句级锚点加动态规划对齐（一对一、一拆二、二合一），字级改动合并，启发式标签（评价词删减、换词、数据前移、语气改为要求、改为百分点，以及区分"数字改动"和"改为具体数字"等），比较时剥离层次序号，修订人姓名脱敏，`--author` 按修订人筛选修订和批注；含密级或内部标注的不输出对比清单。
8. 画像体检的 13 个检查项和单位类型常见异常（`references/maintenance.md` 第二部分）。
9. `secret_marks.py`：两个脚本共用的密级和内部标注识别。优先读本项目共享机检 `lint_rules.json` 的 PRI-005、PRI-006、PRI-007，读不到时用内置副本，另加"严禁外传""不得翻印""内部参阅"等补充写法；引号内的举例不算标注；docx 同时查页眉、页脚和修订中删除的文字。
10. 口述偏好录入、"观察中"记录合并计份后升级为偏好条目、从已发文件版式提取画像字段（SKILL.md 功能一、功能三；`references/edit-patterns.md` 第三部分；`references/intake.md` 第二部分（三））。

## 五、有意没有采用的内容

1. lieflat-gongwen 内置的语料统计参考值：本模块只从用户本单位的历史稿计算参数，不提供全国通用的"标准文风"，避免把某一批语料的习惯当成规范。
2. writing-style-skill 调用外部大模型命令行工具自动提取规则、自动改写 SKILL.md 的做法：本模块的脚本不联网、不调用任何模型，规则由模型在对话中归纳、经用户同意后写入。
3. 自动把偏好应用到稿件上的改稿功能：改稿由 gongwen-polish 负责，本模块只记录偏好。
