# 逐图视觉检查：执行、复查与失败处理

## 固定行为

每张实际图片单独检查。程序生成完整图、最长 720 像素的重叠局部图、左/右/上/下四周条带；小图也保留四周条带。必须逐个文字块对照所属框的四边检查，侧栏、竖框、外围说明不能漏掉。检查不得以整页缩略图、箭头数据正确或媒体外框未越界代替。

初检一次真实视觉调用；终检针对**当前最终图及当前所在页**发起两次独立调用。两次终检均不输入初检结果、修复结论或上一轮答案；它们都重新找全部问题，不只是核对已有缺陷。可使用同一视觉模型，但必须是新请求/新子代理，不复用带结论的会话。“独立”指执行和输入隔离，不代表错误概率相互独立。

任一调用发现缺陷即 FAIL；不确定即待检查；两次结果冲突不按多数票放行。文字压线/出框、重叠、裁切、不可读都必须定位到实际视图和区域。图内问题不能靠整体缩小图片、删除文字或填写通过消除。修复仍遵守原有模板、内容保护和图片语义/风格要求。

## 宿主视觉接口

`visual_evidence.py` 向命令适配器的标准输入发送 JSON，**内含真实 PNG 图片字节的 Base64**，不是仅发送路径、图名、OCR 文本或一个图结构。每次调用均新启动进程；标准输出必须是规定的 JSON 视觉结论。宿主可以用现有视觉插件/新子代理实现适配器，不绑定特定模型、供应商或版本。

适配器必须实际把全部图片交给具备视觉能力的模型；命令可以是本机桥接程序，命令参数用数组，不通过 shell 拼接。用 `--vision-worker-config path/to/worker.json` 或环境变量 `DLR_VISION_WORKER_CONFIG` 传递配置。自有适配器可写为：

```json
{"command": ["python", "/absolute/path/to/host_vision_bridge.py"], "timeout_seconds": 180}
```

请求和输出格式由 `visual_evidence.py` 中的 `make_request`、`PROMPT` 和 `evaluate` 定义。不要让 Agent 手造“工具调用成功”回执。适配器失败、超时、缺图片、缺观察、缺第二次检查，均保留原因并阻断。

包内 `vision_http_worker.py` 是**可选**的 Chat-Completions 兼容接口适配器。使用 `examples/visual-worker/http.json` 时，需要宿主从用户已授权的服务配置中提供：

- `DLR_VISION_API_URL`：完整请求地址，通常以 `/chat/completions` 结尾；没有默认外部地址。
- `DLR_VISION_MODEL`：实际可用的视觉模型名称，不在技能里硬编码。
- `DLR_VISION_API_KEY`：服务凭证，只通过环境传递，不写入报告或仓库。

不得自动把投标材料发给新供应商。未配置已授权的视觉接口时，普通的手动看图仍可帮助定位和修复，但**不能产生本版自动放行凭据**。该环境应停在待视觉执行状态；不是继续手填 JSON 当作完成。宿主内置“能看图”与“Python 可调用该视觉能力”是两件事，需要宿主桥接。

## 初检

流水线仍使用原入口。配置好宿主适配器后，初检由程序调用：

```bash
python scripts/review_pipeline.py input.docx --work-dir work --vision-worker-config worker.json
```

遇到 `requires_image_review` 先读取实际结果；缺题注/豁免等语义计划仍由 Agent 处理。已有准备清单也可单独调用：

```bash
python scripts/figure_review.py inspect-initial input.docx work/reports/initial-visual-review.json --worker-config worker.json --out-dir work/initial-inspections --json-out work/reports/initial-visual-review.checked.json
```

之后将 `.checked.json` 通过原来的 `--initial-visual-review` 参数传回流水线。初检发现 FAIL 可以进入对应修复；调用缺失或 uncertain 不能当作完成。旧版手写通过清单没有真实调用证据，需要重新初检，不能只补版本号。

## 终检

Agent 先核对最新页图，把每张图所在页填入现有最终检查模板的 `figures.objects[].pages`，页名和哈希从渲染清单读取。不要让用户逐张填内部 JSON。然后执行：

```bash
python scripts/figure_review.py inspect-final input.docx work/candidate.docx work/reports/initial-visual-review.checked.json final-visual-review.json work/review-manifest.json --worker-config worker.json --out-dir work/final-inspections --json-out final-visual-review.checked.json
```

程序为每张图实际执行两次检查，汇总后写入 `inspection`；原图、最终完整图、全部局部、四周条带和所在当前页必须齐全。最终文件中保留实际 `repair_note`，不编造修复动作。

有缺陷时返回失败并将调用凭据追加到 `reports/image-discoveries.json`，同步更新清单中的台账引用；**初检曾经通过也照样登记**。按缺陷修好图后，仍先更新域、用同一规则 `--audit-only` 全量审计和重渲染，再重新终检。台账与原始 Word 绑定并随续跑保留，不能重置初检或删除台账来消除问题。

图片终检全部通过时，整份文档的 `overall_status` 仍保持 `pending`；原有逐页检查、表格语义检查及 14 项审计全部完成后，才填写整体验收结论并调用 `finalize_review.py`。只改 `overall_status`、图片 `checks` 或哈希不能覆盖保留的失败调用。图内缺陷同时对比归一化像素；只重新编码 PNG 或改变文件哈希、像素未变，仍不算修复。

## 原生 Shape / 不可直接解码的媒体

没有媒体文件不是豁免。先用原有渲染入口渲染**实际待检文件**，再由 Agent 从最新页面定位对象像素范围，编写定位表（键使用实际对象 ID）：

```json
{"实际对象ID": [{"name": "实际页图文件名.png", "bbox": [100, 120, 900, 700]}]}
```

绑定与裁切哈希由程序生成，不手写：

```bash
python scripts/figure_review.py bind-rendered input.docx initial-review.json actual-render-report.json object-page-locations.json --json-out initial-review.bound.json
```

再将 `.bound.json` 送入 `inspect-initial`。终检时 `bind-rendered` 第一个参数换成当前候选，检查表换成最终表、渲染报告换成当前报告。裁片必须含完整图和必要边界余量；同时提供所在完整页，视觉审查器核对对象匹配，不能选一块正常区域代替整图。截图被裁切、图本身不可读时仍不能通过。

## 验证边界

程序复核图片来源和像素、视图覆盖、实际请求与响应、文件哈希和新旧缺陷闭环。它不是像素语义分类器，仍会受接入模型的误判影响；任何不支持视觉的假适配器都不能用于正式检查。这些本地回执也不是宿主权限隔离或防恶意篡改的安全凭证，宿主强制执行应进一步限制修改检查器和回执的权限。

`tests/test_visual_evidence.py` 使用明确标注的模拟视觉输出，验证执行和阻断逻辑；正式验收会拒绝带模拟标记的输出。旧 `test_figure_gate.py` 仅在测试其原有清单/闭环规则时显式隔离外部视觉依赖，另有不隔离的“手填 PASS 被拒绝”用例。不能用它的通过来声称模型能识别图片。真实坏图测试另外运行：

```bash
python tests/run_visual_regression.py --worker-config worker.json --out-dir work/visual-regression
```

该测试使用用户提供的右侧竖栏出框截图。预期位置仅供测试器比对，不发送给模型；要求模型主动定位两行文字的越界区域。一张负例通过也不代表所有图片都不会漏检。

接口参考：OpenAI 官方 images/vision 指南 `https://developers.openai.com/api/docs/guides/images-vision`；Pillow 官方图像操作文档 `https://pillow.readthedocs.io/en/stable/reference/ImageOps.html`。接口兼容性和实际视觉能力以所用服务实测为准。
