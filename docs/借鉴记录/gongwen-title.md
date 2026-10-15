# 借鉴记录：标题与提纲（gongwen-title）

- 模块范围：`skills/gongwen-title/` 下的 SKILL.md、README.md、references/（不含 `_shared/`）、evals/。
- 阅读方式：2026-10-16 通过 GitHub API 和 raw.githubusercontent.com 在线读取，未克隆、未下载到本仓库。下表“读取版本”为当日默认分支的提交号。
- 借鉴方式：**思路**（分类方法、判定原则，不涉及文字）；**改写**（参考原文的规则，用本项目自拟的公文例句重写）。本模块所有示例、例句均为自拟，未整段照录任何来源。

## 一、MIT 许可来源

| 来源 | 许可证 | 读取版本 | 借鉴内容 | 方式 | 落到的文件 |
|---|---|---|---|---|---|
| [KaguraNanaga/official-document-writing-skill](https://github.com/KaguraNanaga/official-document-writing-skill) `references/writing-techniques.md`“标题拟定技巧” | MIT | e76a6f5 | 拟标题的几项检查角度：题文相符、宽窄适度；标题明确集中；小标题与总标题是领属关系、同级是并列关系；同级句式字数匀称；“帽子大、内容小”与“帽子小、内容大” | 改写 | `references/heading-rules.md` 第二部分第 1～3 条、第三部分 3.2；SKILL.md 任务五判断表中的“大而无当”“标题过窄” |
| [gongyu0918-debug/chinese-official-writing-skill](https://github.com/gongyu0918-debug/chinese-official-writing-skill) `references/structure-editing.md` | MIT | 23e124c | 改动结构时同步核对小标题的数量、名称和顺序 | 思路 | SKILL.md 任务三第 6 条“只替换标题行，正文一字不动”；`references/heading-rules.md` 4.1 列标题树的做法 |

## 二、只学思路的来源（无许可证）

| 来源 | 许可证 | 读取版本 | 借鉴内容 | 处理方式 |
|---|---|---|---|---|
| [alvis961/sop-writing](https://github.com/alvis961/sop-writing) `SKILL.md` Step 4 | 无许可证 | abc8db0 | 不同文体的标题风格不同（总结、述职、汇报、经验、调研各有侧重） | 只学“按材料类型区分标题风格”的思路，未复制文字；`references/title-patterns.md` 2.3 的分类和示例由本项目自拟 |
| [cuberfry68-coder/wenxi-skills](https://github.com/cuberfry68-coder/wenxi-skills) `公文写作算法SKILL.md` | 无许可证 | fd8752b | 全文标题、章节标题、段落主题句逐级展开同一核心意思 | 只学思路，未复制文字；落为 `references/heading-rules.md` 第二部分第 6 条和一致性检查中的“关键词回查” |

## 三、本模块自研内容

1. 转发、批转、印发类标题的处理规则和层层转发示例，以 `shared/rules/doc-types.md` 2.8 为基线展开。
2. 小标题三套方案（对仗式、递进式、数字概括式）的构造方法、适用场合和推荐依据；“条数核对”行的输出约定。
3. 标题与正文一致性检查的判断表和 P0、P1、P2 分级。
4. 常见材料提纲骨架与篇幅比例建议；述职、调研报告提纲的“填写指引”和“调研前提纲”处理。
5. 七类单位的标题与提纲调整，依据 `shared/profiles/units/` 整理。

## 四、有意没有采用的内容

1. KaguraNanaga 项目“标题艺术性”部分建议多用形象化语句和“原话”式标题：本模块只在经验材料、讲话稿中允许，并要求不生造词、不拔高。
2. alvis961/sop-writing 一次生成 10 个标题供挑选的做法：与“默认做减法”冲突，本模块固定为 3 套方案加推荐。
3. KaguraNanaga 项目中“以高质量党建引领高质量发展”一类示例标题：属于政治性提法，本模块不自拟此类标题，严谨模式下只用用户提供的原文。
