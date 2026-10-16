# 借鉴记录：总入口 gongwen

- 模块：`skills/gongwen/`（SKILL.md、references/routes.md、skill-map.md、intake.md、adaptation.md、quick-ref.md、evals/）
- 记录日期：2026-10-16
- 更新：v1.0 第四阶段按计划 D1 停用过渡起草，总入口只做路由。原 `references/interim-drafting.md` 已删除，其中法定公文部分由 gongwen-draft 承接，讲话稿、纪要、信息部分分别由 gongwen-speech、gongwen-minutes、gongwen-brief 覆盖（对照结果见 `docs/借鉴记录/gongwen-draft.md` 第六部分）。

## 一、外部开源项目

以下仓库均通过 GitHub 网页、GitHub API 或 raw.githubusercontent.com 在线阅读，未克隆、未下载到本仓库。本模块没有复制外部仓库的文字或代码，只吸收设计思路并用中文重新写成。

| 来源 | 许可证 | 阅读的内容 | 借鉴的思路 | 落到哪些文件 |
|---|---|---|---|---|
| Yuan1z0825/nature-skills，`skills/nature-writing/SKILL.md` 的 Routing protocol | Apache-2.0 | 先识别请求在几个维度上的取值，用一行写出识别结果，然后继续执行；只在选择有歧义且影响结果时才提问 | 第三步“判断完用一行回显，然后接着做，不等确认”；提问只针对会影响路线或文种的维度 | `SKILL.md` 第三步，`references/intake.md` 第二部分 |
| Yuan1z0825/nature-skills，README 的技能索引和成熟度标签 | Apache-2.0 | 技能索引按“技能、状态、用途、触发词”列出；Draft、Beta、Stable 三档成熟度及其含义 | 能力地图按“skill、能做什么、示例说法”列出，标明 Beta 的含义；核心 skill 与行业包分开列 | `references/skill-map.md` 第一、二部分 |
| Yuan1z0825/nature-skills，`skills/nature-writing/SKILL.md` 的 Submission boundary | Apache-2.0 | 在 skill 内写明哪些请求转给哪个相邻 skill | “适用范围”中的转交清单；“已转交的步骤不顺手代做” | `SKILL.md` 适用范围 |
| Liuxiangjian-ai/official-document-skill，`SKILL.md` | MIT | 按发文主体、受文对象和行文关系选文种；只在缺失事实会让文稿无法使用时才提问，其余用占位符并在文末列出缺项；用户指定文种与目的冲突时指出 | 文种判断“先定关系、再定目的” | v0.1 落在过渡起草的文种判断中；v1.0 总入口只保留规划路线时的一句文种初判（`SKILL.md` 第三步第 4 条、`references/quick-ref.md` 第四部分），完整的文种判断由 gongwen-draft 承接 |

Apache-2.0 项目只借鉴了流程设计思路，没有复制其文字或代码，因此无需随附其 NOTICE；汇总进 THIRD_PARTY_NOTICES.md 时按“思路参考”登记。

## 二、仓库内资料

本模块的规则性内容均来自仓库内的共享资料，按 manifest.json 同步，未另行改写规则本身：

| 内容 | 来源 |
|---|---|
| 文种初判用的决策表、易混辨析 | `shared/rules/doc-types.md` |
| 文种初判引用的条款 | `shared/rules/tiaoli.md` |
| 5 种角色模式、识别信号、无画像降级、首轮交付 | `shared/profiles/roles.md` |
| 交接卡、未安装兜底、安装命令 | `shared/style/output-format.md` 第四部分 |
| 简版画像模板、画像修改规则 | `shared/profiles/profile-template.md` |
| 写作日历 | `shared/profiles/calendar.md` |
| 单位差异 | `shared/profiles/units/*.md` |
| 22 个 skill 的分工、行业包与核心 skill 的分工 | `docs/下一轮工作计划.md` 第七节、D14、D21、Q1 |
| 路线设计 | `docs/顶层设计.md` 第 4.0、4.1、4.19 节 |

## 三、示例与评测材料

`evals/files/` 下的会议记录、数据、送审稿、画像均为虚构，单位一律写作“某市某局”“某医院”“某集团有限公司”，人名、病案号、身份证号为虚构并已部分遮挡，用于测试脱敏提醒。
