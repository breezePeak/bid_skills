# 执行与续跑命令

本文件只说明如何调用。合格标准在 `SKILL.md` 第 6 节，失败回路在第 7 节；命令运行成功不代替这些验收。

## 1. 初次执行

在技能目录运行，未指定模板时使用内置默认：

```bash
python scripts/review_pipeline.py input.docx --work-dir work
```

上传模板追加 `--template template.docx`；已确认的模板约定追加 `--template-style-json template.style.json`；正文覆盖追加 `--text-rules text-rules.json`。实际文件名和路径从本次任务读取，不凭示例猜测。

程序要求 `requires_image_review` 时，按 `references/visual-inspection.md` 实际调用宿主视觉接口，生成绑定原图及完整/局部/边缘视图的初检结果，不手填 PASS。配置好 `--vision-worker-config` 时入口会自动执行初检。有缺题注、豁免或编号语义需要定位时，按真实对象生成 `object-plan.json`/`numbering-plan.json`。随后仍以原始输入执行，带上 `--initial-visual-review` 及存在的相应计划；不要传不存在的占位文件，也不要让用户代填内部记录。

默认 `--renderer auto` 与 `--field-engine auto`，分别按 **Word → WPS → LibreOffice** 顺序选择可调用引擎；渲染失败会按顺序重试。明确指定某引擎时只执行该引擎，不静默降级；只指定 `--field-engine word` 并不等于限定渲染器，要求两者都是 Word 时同时传 `--renderer word --field-engine word`。环境可用和初检完成均不是验收通过。依赖、记录及失败处理见 `rendering-engines.md`。

## 2. 局部修复后实际更新域

先按失败类型完成真实修复，保存为新文件，保留原文与本轮修复前副本；完成对应修改范围和内容复核。然后更新域：

```bash
python scripts/field_refresh.py repaired.docx --out refreshed.docx --json-out refreshed-fields.json
```

刷新脚本使用 `--engine word`、`--engine wps` 或 `--engine libreoffice` 指定已确定引擎；流水线对应参数是 `--field-engine`，两者不能写混。需要指定 UNO Python 时沿用 `--uno-python`。不能把刷新报告对应到另一个后来修改过的文件。

涉及编号结构修复或引擎恢复域指令时，用输出再实际刷新一次，复核再次计算的序号和域结构：

```bash
python scripts/field_refresh.py refreshed.docx --out refreshed-again.docx --json-out refreshed-again-fields.json
```

这一步由 Agent 按任务需要调用，不声称当前流水线会自动追加第二次刷新。引擎要求保持一致；再次刷新失败按第 7 节处理。后续验收必须改用 `refreshed-again.docx` 和它的报告，不退回第一次的结果。

## 3. 只验收续跑

下面示例使用默认模板，且前一步只需一次刷新：

```bash
python scripts/review_pipeline.py refreshed.docx --work-dir work --audit-only --source input.docx --field-update-report refreshed-fields.json --initial-visual-review work/reports/initial-visual-review.json --content-plan work/reports/content-plan.json
```

有第二次刷新时，将输入和 `--field-update-report` 同时换成第二次输出。存在明确的 `--renderer` / `--field-engine` 要求时，续跑必须保留；默认模式仍使用 `auto`。上传模板任务必须补上**原来的** `--template`、`--template-style-json`、`--text-rules`；没有用户覆盖时也沿用本次生效规则，不临时改脚本默认值。

有对象豁免等语义决定时，追加与当前对象哈希绑定的有效 `--object-plan`。沿用原文绑定的内容保护计划，不重新以候选生成基准；脚注搬移登记不得遗漏。对象真的发生变化时，先按实际源与候选核对对应关系，再用支持的工具重新定位，不能只改哈希躲过校验。

`--audit-only` **不会自动修好问题**，它负责验证局部修复后的文件。不要为省事重跑整套样式/布局修复覆盖已经修好的图与分页。

## 4. 按实际结果继续

| 结果或问题 | Agent 下一步 |
|---|---|
| `requires_image_review` | 实际看原图、填完整初检及必要语义计划，然后按原始输入继续。不能填未核验的 PASS。 |
| `requires_field_update` | 读取具体失败项，解决引擎/域结构问题；实际更新当前候选后进入只验收续跑。 |
| `failed`，有 `failed_gates` 或具体 `issues` | 读取对应 `gate-*.json` 或 `pipeline-failure.json`，定位具体对象，按主技能第 7 节修复再验收。渲染成功也不能覆盖审计失败。 |
| `awaiting_visual_review` | 仅表示程序审计及渲染已就绪；先完成主技能第 6 节的逐页、逐图和表格语义验收，尚不能交付。 |
| 存在表格 `review` | 需要修复的先改 Word 并重新审计；无需修改的在最新记录中填写现有 `keep_separate`/`acceptable` 决定和具体理由，并核对最新页面。不能新增一个假“已合并”状态替代真实修改。 |
| 清单、规则、原图、候选或页图不匹配 | 找出真正变化的对象/文件，重新执行受影响的检查并生成当前完整验收记录；不手改状态、版本号或哈希。 |
| 最终入口返回 `failed` | 仍回到第 7 节；保留已有正确成果，不能把当前 candidate 改名交付。 |

## 5. 最后验收和交付

程序复核完成后，依据最新 `work/reports/final-visual-review.template.json` 定位各图当前所在页，通过 `figure_review.py inspect-final` 执行两次从零检查并写入真实调用证据，随后补齐逐页和表格实际观察。具体命令见 `references/visual-inspection.md`。页面、图片和表格 review 都必须核对当前内容；不能复制旧 PASS。保留清单中原文绑定的 `image_discovery_ledger`，后续新发现也必须实际修复并复查关闭。若发现任何问题，先回到对应修复，不执行发布。

所有明确标准都满足后：

```bash
python scripts/finalize_review.py work/review-manifest.json final-visual-review.json --out final.docx --json-out release.json
```

只交付入口成功输出的 Word 及实际修复报告。回报原问题、位置、实际修改、复查结果和真实验证范围；报告同时记录实际渲染引擎、切换原因及实际域更新引擎；不能把 WPS/LibreOffice 的结果写成 Windows Word F9 实测。源码级回归与正常文档验收是不同任务，不把测试通过等同于某份 Word 已通过视觉验收。
