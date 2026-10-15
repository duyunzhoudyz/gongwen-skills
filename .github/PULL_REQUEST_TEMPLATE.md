## 改了什么

<!-- 一两句话说明改动内容。 -->

## 依据

<!-- 规则类改动写明文件名称和条款号；其他改动说明原因。 -->

## 影响范围

<!-- 影响哪些 skill、共享资料或工具脚本。 -->

## 提交前检查

- [ ] `python tools/validate.py` 通过
- [ ] `python -m unittest discover -s tests -t .` 全部通过（改了 `tools/` 的，已补充或更新 `tests/tools/` 中的测试）
- [ ] `python tools/sync_shared.py --check` 通过（改了 `shared/` 的，已先运行 `python tools/sync_shared.py`）
- [ ] `python tools/build_prompts.py --check` 通过（改了 skill 内容的，已重新生成 `dist/prompts/`）
- [ ] 不含任何真实的内部文件（哪怕已经公开）；示例和评测材料均已虚构化，不含涉密信息和个人信息
- [ ] 不含字体文件（`.ttf`、`.otf`、`.ttc`、`.woff` 等）
- [ ] 借鉴了外部项目的，已确认许可证，并在 `docs/借鉴记录/` 中记录来源；没有借鉴的，请在此注明"无借鉴"
