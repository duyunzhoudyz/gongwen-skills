# 对比评测脚本

用真实模型跑各 skill 的 `evals/evals.json`，比较"装 skill（A 组）"和"不装 skill（B 组）"的表现。结果和方法说明见 `evals/results/v0.1-benchmark.md`。

## 流程

```bash
# 1. 准备：每条用例建 A、B 两个匿名工作目录，复制输入材料，生成子智能体任务说明
python evals/bench/prep.py --out <评测目录>

# 2. 运行：对 <评测目录>/tasks/ 下每个任务说明，各起一个子智能体（Claude Code 的 Agent），提示词为
#    "请完整阅读任务说明文件 <路径>，然后严格按其中的要求逐个处理请求。"
#    A、B 两组任务说明除"已安装的 skill"一段外逐字相同，模板见 templates/run-A.md、run-B.md、common.md

# 3. 机检：gongwen_lint、fact_diff --ledger、check_docx、输入文件完整性
python evals/bench/machine.py --bench <评测目录>

# 4. 期望分类（只需一次）：按 templates/classify.md 起一个子智能体，产出 <评测目录>/exp_class.json

# 5. 盲评：A、B 随机标为"甲""乙"，生成评分说明；对 grade_tasks/ 下每个说明各起一个子智能体
python evals/bench/grade_prep.py --bench <评测目录>

# 6. 汇总：解盲，输出原始结果 JSON 和统计表
python evals/bench/aggregate.py --bench <评测目录> --out evals/results/v0.1-raw.json --tables <统计表.md>

# 7. 缺陷归因：列出 A 组未通过的期望，按 defects/<skill>.md 各起一个子智能体
python evals/bench/defect_prep.py --bench <评测目录>

# 8. 归档：把各次运行的回复和交付文件复制到仓库
python evals/bench/archive.py --bench <评测目录> --out evals/results/v0.1-runs
```

## skill 范围与版本

- `prep.py --skills`：默认评测插件市场清单（`.claude-plugin/marketplace.json`）中的全部 skill；写 `--skills v0.1` 复现 v0.1 的 7 个 skill，也可以列出具体名称。
- `prep.py --tag`：评测版本标签（默认 `v1.0`），用于生成工作目录编号，并写入 `meta.json`；`aggregate.py` 据此填写原始结果中的版本。
- `prep.py --groups A`：修改 skill 后复测，只重建 A 组（装 skill）的工作目录和任务说明，B 组的目录、输出和输入哈希保持不动；之后照常运行机检、评分准备、盲评和汇总。新增的用例、输入材料改过的用例，B 组也一并重建，B 组任务说明里只列这些用例，保证两组输入一致。本次要运行的任务说明清单写在 `tasks_written.json`。评测目录中必须已有上一轮完整的 `hashes.json`。复测前建议把整个评测目录另存一份，保留上一轮的评分。
- 用例的 `file_dest` 字段指定材料放在工作目录中的位置，例如 `{"evals/files/已有单位画像.md": ".gongwen/单位画像.md"}`。
- 输入材料为 docx 时按"接受全部修订"读取（与 fact_diff 一致）；为 xlsx 时逐表逐行转为文本，供 fact_diff 比对。

## 跨 skill 触发评测

模拟 Agent 工具按 description 选 skill：给模型看全部 skill 的描述和各 skill `evals/trigger.json` 中的说法，为每条说法选一个 skill 或"无"。

```bash
python evals/bench/trigger_prep.py --out <评测目录>          # 生成 <评测目录>/trigger/tasks/batch-*.md
# 对每个任务说明各起一个子智能体，答案写到 trigger/answers/batch-*.json
python evals/bench/trigger_score.py --bench <评测目录> --out evals/results/<版本>-trigger.json --tables <表.md>
```

判定：正例选中来源 skill 为对；反例没有选中来源 skill 为对。

## 注意

- 评测目录放在仓库之外（`prep.py` 会拒绝仓库内路径），避免 B 组读到 skill 文件。
- 模型输出有随机性，重跑第 2、5 步会得到不同样本。报告中的统计数字可以不经评测目录、直接从仓库内的原始结果复算：`python evals/bench/aggregate.py --from-raw evals/results/v0.1-raw.json`。
- `archive.py` 会把本机绝对路径替换为 `<评测目录>`、`<仓库>`、`<用户目录>`，归档前不必手工脱敏。
- 并发子智能体不超过 10 个。
