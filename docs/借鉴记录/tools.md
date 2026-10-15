# 借鉴记录：工具链（tools/）

本模块未复制任何外部仓库的代码或文字，以下只记录参照的公开约定和设计思路。

| 来源 | 许可证 | 借鉴内容 | 落到哪些文件 |
|---|---|---|---|
| Agent Skills 开放规范（frontmatter 字段约束） | 规范文本，未复制 | name 只用小写字母、数字、连字符，≤64 字符；description ≤1024 字符；name 不含保留词 | `tools/validate.py` |
| anthropics/skills 的 `.claude-plugin/marketplace.json` 写法（顶层设计 3.4） | 只参照字段结构，未复制内容 | `plugins[].source`、`plugins[].skills` 的路径解析方式 | `tools/validate.py`、`tools/build_prompts.py`（读取 metadata.version） |
| baoyu-skills 的"自包含"思路（见顶层设计 3.3） | 只学思路 | 共享资料以 shared/ 为唯一源，构建时复制进每个 skill，CI 校验一致性 | `tools/sync_shared.py` |

`tools/install.py` 的安装目录映射来自项目任务说明中的路径表，其中 codebuddy（user 级）和 trae 为社区资料，安装时会打印提醒。

## v1.0 安装体验与协作模板（追加）

以下同样未复制外部代码或文字。

| 来源 | 许可证 | 借鉴内容 | 落到哪些文件 |
|---|---|---|---|
| anthropics/skills 的 `.claude-plugin/marketplace.json` 写法（同上） | 只参照字段结构，未复制内容 | 按 `plugins[].skills` 读取核心插件与行业包清单，供 `--pack`、`--uninstall` 使用 | `tools/install.py` |
| GitHub 官方文档中的 issue forms 语法与 `config.yml` 约定 | 规范文本，未复制 | YAML 表单的字段类型、必填校验、联系链接写法 | `.github/ISSUE_TEMPLATE/*.yml` |
| PEP 668（externally-managed-environment）的公开说明 | 规范文本，未复制 | pip 安装失败时给出 apt 包或 `--user --break-system-packages` 的处理提示 | `install.sh` |

`tools/install.py --detect` 的探测依据（各工具默认的用户配置目录名和命令名）来自各工具的公开安装说明，只判断目录或命令是否存在，不运行这些工具。
