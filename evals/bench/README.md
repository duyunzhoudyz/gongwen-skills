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

## 注意

- 评测目录放在仓库之外（`prep.py` 会拒绝仓库内路径），避免 B 组读到 skill 文件。
- 模型输出有随机性，重跑第 2、5 步会得到不同样本。报告中的统计数字可以不经评测目录、直接从仓库内的原始结果复算：`python evals/bench/aggregate.py --from-raw evals/results/v0.1-raw.json`。
- `archive.py` 会把本机绝对路径替换为 `<评测目录>`、`<仓库>`、`<用户目录>`，归档前不必手工脱敏。
- 并发子智能体不超过 10 个。
