# 贡献指南

欢迎一线文字工作者和开发者一起完善这个项目。最需要的贡献有四类：

1. **规则纠错**：发现某条规则与《条例》、国标原文或本行业惯例不符，请提 issue，并附上依据（文件名称和条款号）。
2. **行业包**：医院、高校、中小学、国企、基层 5 个行业包已随 v1.0.0 发布（Beta）。欢迎在 Discussions 中补充本行业的模板、术语和红线，或指出与实际做法不符之处。
3. **评测用例**：你在工作中遇到的真实写作场景，脱敏、虚构化之后可以变成评测用例，帮助我们检验效果。
4. **试用反馈**：不写代码也能帮上忙。用虚构的材料试一试，把效果好坏告诉我们，做法见 [docs/评测/测试者手册.md](docs/评测/测试者手册.md)。

## 提交入口

- **Issue 表单。**到 [Issues](https://github.com/duyunzhoudyz/gongwen-skills/issues/new/choose) 选对应的表单，表单文件在 `.github/ISSUE_TEMPLATE/`。
  - 规则纠错，用于规则与文件原文或行业惯例不符。
  - 新场景需求，用于现有 skill 写不好的材料。
  - 行业包投稿，用于行业模板、术语、红线、评测用例。
  - Bug 报告，用于安装失败、脚本报错、skill 没有触发。
  - 测试反馈，用于试用中发现的问题，也可以记录好用的地方。
- **Pull Request。**提交时按 [PR 模板](.github/PULL_REQUEST_TEMPLATE.md) 填写，里面的“提交前检查”要全部过一遍。
- **Discussions。**使用心得、写作经验、行业素材的讨论，到 [Discussions](https://github.com/duyunzhoudyz/gongwen-skills/discussions)。

## 提交前必读

- [docs/开发规范.md](docs/开发规范.md)，讲目录结构、接口约定、写作要求。
- [使用须知.md](使用须知.md)，讲保密与合规底线。

## 红线

- **不提交任何真实的内部文件**，哪怕已经公开。示例和评测材料一律虚构，单位写作“某市某局”“某医院”。
- **不提交字体文件。**
- 宣传和文档中不使用“规避 AI 检测”“代写”之类的说法。
- 借鉴外部项目时，先确认许可证：
  - MIT、Apache-2.0 等宽松许可：可以借鉴，在 `docs/借鉴记录/` 里记录来源；
  - 非商业许可、专有许可、无许可证：只能学思路，不能复制。

## 新增 skill 或行业包

新增前先在 Issue 里说明场景，确认与现有 22 个 skill 的分工不重叠（分工以 `docs/下一轮工作计划.md` 第七节“边界总表”为准）。动手后按下面的清单逐项完成，细节见开发规范对应章节。

1. **目录结构**（开发规范第 1 至 3 节）：`skills/<名称>/` 下有 `SKILL.md`、`README.md`、`manifest.json`、`references/`、`evals/`，需要脚本时加 `scripts/`。目录名与 `name` 一致，统一 `gongwen-` 前缀。`SKILL.md` 的 `description` 写在一行内，正文不超过 300 行，细节放进 `references/`。`references/_shared/` 由脚本生成，不要手改。
2. **`manifest.json`**（开发规范第 4 节）。写 `name`、`version`、`status`、`shared`（共享资料清单）和 `prompt_bundle`（提示词版打包清单）。新 skill 的 `status` 一律写 `beta`。改完运行 `python tools/sync_shared.py` 把共享资料同步进来。
3. **插件清单。**在 `.claude-plugin/marketplace.json` 加条目。核心 skill 加进 `gongwen` 插件的 `skills` 数组；行业包单独加一个 plugin 条目，描述里写明“需先安装 gongwen”。漏登记时 `validate.py` 会报错。
4. **评测用例**（开发规范第 9 节）。`evals/evals.json` 不少于 12 条，格式和写法见下一节。
5. **触发评测。**`evals/trigger.json` 20 条，10 条应触发、10 条不应触发，格式是 `{"query": "……", "should_trigger": true}`。不应触发的优先选与相邻 skill 容易混淆的说法。
6. **测试。**带检查脚本的 skill，在 `tests/<短名>/` 下补测试和标本，标本要有必错和必过两类，参照 `tests/soe/`。改了 `tools/` 的，同时更新 `tests/tools/`。
7. **借鉴记录**（开发规范第 10 节）：借鉴了外部项目的，在 `docs/借鉴记录/<模块>.md` 记录来源仓库、许可证、借鉴内容和落到哪些文件，并汇总进 `THIRD_PARTY_NOTICES.md`；没有借鉴的，在 PR 里注明“无借鉴”。
8. **文档与构建。**更新 `README.md` 和 `README_EN.md` 的技能地图、`CHANGELOG.md`，运行 `python tools/build_prompts.py` 重新生成 `dist/prompts/`。

## 评测用例怎么写

评测用例的格式和硬性要求见开发规范第 9 节，这里补充写法。

1. **格式。**每条用例有 `id`、`prompt`（用户原话，口语化）、`expected_output`（期望结果简述）、`files`（输入材料，放在 `evals/files/`）、`expectations`（逐条可判定的期望）、`unit_type`、`role_mode`、`tags`。
2. **覆盖面。**每个 skill 至少覆盖 4 种单位类型、3 种角色模式，至少 2 条“防编造”用例（故意少给数据）、1 条“保密”用例（材料含“内部资料”之类的标识）、1 条交接卡用例。
3. **虚构化。**单位写“某市某局”“某医院”，人名写“张某”“李某”，数字、日期、文号都自己编。真实场景只借用它的结构和难点，不搬材料。提交前先读测试者手册里的“怎样脱敏”一节。
4. **不得复用示例数字。**`SKILL.md` 和 `references/` 里的示例，不能用评测材料里的场景、数字和日期，否则评测就成了背答案。提交前运行 `python tools/validate.py --eval-overlap`，有警告的先改掉。法规日期这类合法重合，在 PR 里说明即可。
5. **期望分通用与约定。**
   - 通用期望是熟悉公文写作的人，只看用户消息和材料也应做到的要求，如不编造数据、文种选对、正文没有 Markdown 符号、篇幅符合要求、见到密级标识就停下。
   - 约定期望要知道本项目的做法才能满足，如【起草说明】栏目、素材台账的列名、交接卡格式、问题分级 P0/P1/P2。
   - 评测报告把两类分开统计，因为不装 skill 的对照组天然做不到约定。一条用例里通用期望要占多数；一条期望里只要夹了约定，整条算约定。
6. **期望要能逐条判定。**写“素材中没有的数字一律标【待补】，未编造任何数据”，不写“内容质量较高”。首轮交付、篇幅目标和 `role_mode` 标签的写法见开发规范第 9 节“期望写法”。

## 成熟度从 Beta 升到 Stable

新 skill 一律从 Beta 起步。同时满足下面三条，才能升到 Stable。

1. 真实模型评测的通过率不低于 80%；
2. 防编造用例中出现编造的次数为 0；
3. 触发评测中能被正确触发。

升级时，同一个 PR 里改 `manifest.json` 的 `status`、skill 自己 README 的“成熟度”小节、根目录 `README.md` 技能地图里的标签，并在 `CHANGELOG.md` 记一笔，附上评测报告的链接。数字以 `evals/results/` 里的报告为准。

行业包要等一线测试者反馈和人工盲评之后再评定，不能只凭评测数字升级。

## 开发流程

1. 共享资料只改 `shared/` 下的文件。改完运行：
   ```bash
   python tools/sync_shared.py
   ```
2. 提交前，按 PR 模板里的“提交前检查”逐项运行：
   ```bash
   python tools/validate.py
   python tools/validate.py --eval-overlap
   python -m unittest discover -s tests -t .
   python tools/sync_shared.py --check
   python tools/build_prompts.py
   python tools/build_prompts.py --check
   ```
   改了 skill 内容的，先运行不带 `--check` 的 `build_prompts.py` 重新生成 `dist/prompts/`，再运行带 `--check` 的确认一致。
3. PR 说明中写清楚：改了什么、依据是什么、影响哪些 skill。借鉴了外部项目的，写明来源；没有借鉴的，注明“无借鉴”。

## 行为准则

就事论事，尊重不同单位、不同岗位的写作习惯。规则有争议时，以文件原文为准；原文没有规定的，以多数一线从业者的做法为准。
