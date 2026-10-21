# 竞品对标评测报告

评测时间：2026 年 10 月 21—25 日
对象：本项目（gongwen-skills 全套 22 个 skill）与 GitHub 上 8 个同类公文 skill，外加“不装 skill”对照
执行模型：claude-sonnet-5-5（推理强度中），Claude Code 子智能体，所有参赛方同一模型、同一题目、同一任务说明
评审：Claude（claude-sonnet-5-5）与 DeepSeek（deepseek-chat，API，同一组材料正序、倒序各评一次）双盲评
方法与脚本：[evals/compete/README.md](../compete/README.md)；原始统计表见本目录 `compete-v1/`～`compete-v6/`、`compete-final/`

## 一、结论

**主要结果取自两套留出题**（c4：tasks-v2；c5：tasks-v3），共 54 道题、552 次运行。出题人不读任何参赛方的 skill；这两套题出题后，本项目没有依据题目内容修改过 skill。

| 维度 | 本项目 | 最好的竞品 | 说明 |
|---|---|---|---|
| 去 AI 味：事实保持 | **100%（682/682 项，32/32 次全部保留），新增数字 2 个** | lieflat-less-ai-tone 100%，但改动比例仅 0.6%、11 次未交付正文；Liuxiangjian 98.4%，新增数字 11 个 | 第一 |
| 去 AI 味：编造（两家评审） | **0/32、0/64** | Liuxiangjian 9/32、6/64 | 第一 |
| 排版：中立版式审计 | **96.6 分** | cj0103 93.2 | 第一（按 GB/T 9704-2012 可测量条款，独立编写的审计脚本） |
| 排版：字体诚实 | **32 次全部写入标准字体名，27 次告知用户缺字体** | mizzlelover 32 次全部写入替代字体，其中 12 次未告知 | 第一 |
| 写作：编造（Claude 评审） | **11/40 次运行** | KaguraNanaga 17/40、lieflat-gongwen 16/40 | 第一（最少） |
| 涉密材料识别 | c4 唯一识别（2/2）；c5 密级只在页眉，全部参赛方漏检 | 其余参赛方 0/2 | 页眉漏检已修复，修复后复测 2/2（见第四节） |
| 写作质量（平均名次，1 最好） | Claude 3.20（第 2，第 1 为 3.15）；DeepSeek 3.88（第 6） | KaguraNanaga（Claude）、Liuxiangjian（DeepSeek） | 第一梯队，未稳定第一 |
| 排版质量（平均名次） | 2.72（第 3） | cj0103 2.12 | 第一梯队 |
| 去 AI 味质量（平均名次） | Claude 2.19（第 3）；DeepSeek 1.97（第 2） | Liuxiangjian 1.78 / 1.69 | 第一梯队 |

一句话：**在防编造、事实保持、保密、版式合规、字体诚实这些“不能出错”的维度上，本项目领先所有参赛方；在文稿整体质量（评审员的整体排名）上处于第一梯队，与最强竞品差距在 0.1～0.6 个名次之间，尚未稳定第一。**

## 二、参赛方

| 赛道 | 参赛方 | 星标（2026-10-25） | 许可证 | 评测提交 |
|---|---|---|---|---|
| 写作 | 本项目 | — | MIT | 见各轮 meta |
| 写作 | [larashero3-dotcom/lieflat-gongwen](https://github.com/larashero3-dotcom/lieflat-gongwen) | 1575 | PolyForm-Noncommercial | e0a5aba |
| 写作、去 AI 味 | [Liuxiangjian-ai/official-document-skill](https://github.com/Liuxiangjian-ai/official-document-skill) | 402 | MIT | cbe5f8c |
| 写作 | [KaguraNanaga/official-document-writing-skill](https://github.com/KaguraNanaga/official-document-writing-skill) | 308 | MIT | e76a6f5 |
| 写作 | [Amer-CN/super-official-writer](https://github.com/Amer-CN/super-official-writer) | 9 | 无 LICENSE（README 声明 scripts 为 MIT） | 56685cf |
| 排版 | [mizzlelover/gongwen-gbt9704-skill](https://github.com/mizzlelover/gongwen-gbt9704-skill) | 816 | MIT | 6a5e47a |
| 排版 | [KaguraNanaga/document-format-skills](https://github.com/KaguraNanaga/document-format-skills) | 236 | MIT | a1005d1 |
| 排版 | [cj0103/gbt-9704-2012-skills](https://github.com/cj0103/gbt-9704-2012-skills) | 9 | MIT | 931dbfd |
| 去 AI 味 | [larashero3-dotcom/lieflat-less-ai-tone](https://github.com/larashero3-dotcom/lieflat-less-ai-tone) | 3704 | MIT | 27d2923 |
| 三赛道 | 不装 skill（对照） | — | — | — |

本项目在开发中借鉴过其中多数项目的思路，见 [THIRD_PARTY_NOTICES.md](../../THIRD_PARTY_NOTICES.md)，在此致谢。对比只针对评测时的版本，各项目此后可能已有改进。

## 三、方法

1. **题目**：三个赛道——写作（通知、请示、函、总结、讲话、方案、调研报告、纪要、经验材料、改稿，另有 1 道涉密合规题单独统计）、排版（红头、上行文、信函、纪要、普通材料、讲话，含 2 道修复已有 docx）、去 AI 味（8 种体裁的 AI 初稿，初稿由 DeepSeek 按事实单生成）。全部材料虚构。检查点只用通用口径，不含任何参赛方自己的约定。
2. **执行**：每方每题跑 2 次；每次一个匿名工作目录；任务说明除“已安装的 skill”一段外逐字相同；本项目按用户实际安装方式提供全套 skill，由模型按 description 自选。
3. **评审**：匿名、随机排序；评审员逐条判检查点、按维度打分（0～10）、核查编造、整体排名。DeepSeek 对同一组材料正序、倒序各评一次。
4. **机检**：fact_diff（新增数字）、事实保持（事实单逐项比对）、独立编写的 GB/T 9704 版式审计、LibreOffice 渲染图供评审员看版面、字体名记录。
5. **隔离**：运行前临时移除会话记忆中涉及本项目规则的条目，运行后扫描竞品输出确认未受影响（第一轮曾出现记忆污染并已重跑，见局限）。

## 四、迭代过程（如实记录）

| 轮次 | 题目 | 本项目版本 | 写作名次（Claude） | 排版名次 | 去 AI 味名次（Claude） | 说明 |
|---|---|---|---|---|---|---|
| c1 | tasks | 57c3088 | 6/6（4.45） | 2/5 | 3/4 | 首轮：安全指标最好，但过度保守，写作可用性全场最低 |
| c2 | tasks | 01c2b9f | 5/6（3.95） | 2/5 | 3/4 | 红线校准、改稿力度、三项新机检、语料基线 |
| c3 | tasks | 4d6eeb3 | 6/6（3.80） | **1/5** | **1/4** | 按 c2 评语修正（此后 tasks 不再是留出题） |
| c4 | **tasks-v2（新）** | f0e8703 | 2/6（3.00） | 2/5 | 3/4 | 全部参赛方重跑 |
| c5 | **tasks-v3（新）** | 878aaf5 | 3～4/6（3.40，与 KaguraNanaga 并列） | 3/5 | 3/4 | 全部参赛方重跑；页眉密级漏检 |
| c6 | tasks-v3 | 0585b27 | 4～5/6（3.75，与 KaguraNanaga 并列） | 3/5（2.56） | 2/4（1.94） | 修复页眉漏检、成文日期等后只重跑本项目，竞品沿用 c5 |

- c2、c3 在同一套题上反复修改、复测，其结果有“对着题目改”的成分，因此不作为结论依据；结论取自 c4、c5 两套留出题。
- c6 中本项目重跑：涉密题 0/2 → 2/2；写作编造 Claude 7/20 → 5/20、DeepSeek 17/40 → 12/40；写作 DeepSeek 名次 3.98 → 3.52；去 AI 味 Claude 名次 2.25 → 1.94。写作 Claude 名次 3.40 → 3.75，在同一水平内波动。c6 只重跑了本项目，与竞品不是同时运行，只作“修复后不退步”的确认。
- 迭代中发现并改掉的问题包括：【待补】过多、篇幅不足、回复噪声、上行文发文字号位置、排版时添加原稿没有的要素、页眉残留页码、起草说明混入正文文件、页眉密级漏检、成文日期自行填写、相对年份换算等。

## 五、局限

1. 执行模型只有 Sonnet 5.5，结论不一定适用于其他模型。
2. 每题每方 2 次重复；单题样本小，名次差在 0.3 以内的应视为同一水平。
3. 题目由本项目设计（出题人不读任何 skill，只用通用口径），仍可能偏向本项目熟悉的场景。
4. 评审员可从写法认出本项目（如【起草说明】栏目），盲评不彻底。
5. DeepSeek 在写作赛道位置偏差明显（同一组材料正倒序排名的 Kendall τ 只有 0.11～0.19），写作赛道 DeepSeek 名次只作参考；去 AI 味赛道一致性较好（0.56～0.60）。
6. 本项目把全套 skill 与单个竞品 skill 比较，比较的是“用户实际安装后的产品”。
7. 渲染机缺仿宋_GB2312 等字体，排版赛道的页面图片由 LibreOffice 替代字体显示，只用于看布局；字体合规以审计中的字体名为准。审计脚本对信函格式发文字号位置有误判，评审员对信函双线长度（标准为 170 mm）也有误判，均已在分析中剔除。
8. 第一轮执行时会话记忆被注入子智能体，影响了一个竞品的一批运行，已隔离并重跑；c1～c3 各轮“今天”的日期与部分题目的设计日期不一致，c6 起按题目设计日期统一。
9. 评审员是大模型，没有人工盲评，也未经一线文字工作者实测。

## 六、复现

```bash
python evals/compete/prep.py --out <评测目录> --tasks-dir evals/compete/tasks-v3 --reps 2
# 按 run_batches.json 起子智能体执行；之后：
python evals/compete/machine.py --bench <评测目录>
python evals/compete/judge_prep.py --bench <评测目录>
# 按 judge_batches.json 起 Claude 评审；DeepSeek 评审：
python evals/compete/deepseek_judge.py --bench <评测目录>
python evals/compete/aggregate.py --bench <评测目录> --out <结果.json> --tables <表.md>
python evals/compete/combine.py <c4.json> <c5.json> --labels c4,c5 --out <合并表.md>
```

DeepSeek Key 只从仓库根目录 `deepseek.key.local` 读取（已加入 .gitignore）。
