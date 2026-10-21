export const meta = {
  name: "compete-v6-exec",
  description: "竞品对标 v6 执行阶段：72 个子智能体并发跑各批次（Sonnet 5.5，medium，并发≤10）",
  phases: ["exec"]
};

const BATCHES = "D:/bench-compete-v6/run_batches.json";
const MAX_CONCURRENCY = 10;

export default async function (args) {
  const fs = await import("node:fs");
  const batches = JSON.parse(fs.readFileSync(BATCHES, "utf-8"));

  phase("exec");
  const results = await parallel(
    batches.map((b) => ({
      id: b.batch_id,
      run: () => agent({
        model: "sonnet",
        effort: "medium",
        prompt: `请完整阅读任务说明文件 ${b.instruction_path}，然后严格按其中的要求逐个处理请求。`
      })
    })),
    MAX_CONCURRENCY
  );

  return { total: results.length, done: results.filter(r => r).length };
}
