# 执行与续跑

执行顺序和验收标准以 `SKILL.md` 为准。正文、表格、图片按完整内容块推进；命令由 Agent 执行，不让用户逐块确认或填写内部 JSON。

## 开始与全局准备

```bash
python scripts/review_pipeline.py input.docx --work-dir work
```

有用户模板时追加 `--template template.docx`；明确模板约定和文字覆盖分别使用 `--template-style-json profile.json`、`--text-rules text-rules.json`。模板和规则在本会话锁定，续跑继承，不反复解析或重新选择。渲染和更新域默认 Word → WPS → LibreOffice，明确指定不静默降级。

入口扫描对象并返回 `global_setup`、第一块和工作文档。只准备共享样式、编号定义、页面和页眉页脚；不运行原来的全文九阶段修复。Agent 使用模板修正共同设置后提交：

```bash
python scripts/review_pipeline.py --work-dir work --action global --proposal global-ready.docx --note "实际修改的共同设置"
```

共同设置已正确时，`--proposal` 使用当前工作文档；不能为了结束准备虚构修改。提交全局设置会先做一次代表性 Office 保存往返检查，成功且条件未变时复用。真实格式改变或无法确认时保存诊断样例，不重复跑全文；以原模板为标准。全局设置完成后不再重复执行。

## 当前块

`next` 返回稳定块 ID、当前文档路径、部件和节点范围、完整文本/XML 文件及上下文。长表使用 `--view-offset`、`--view-size` 阅读后续窗口；这些参数不拆表、不改变修复范围，也不表示真实页数。

```bash
python scripts/review_pipeline.py --work-dir work
python scripts/review_pipeline.py --work-dir work --view-offset 14 --view-size 14
```

只在返回的完整块中检查和修改。正文与其标题、图表及其题注、表内图片保持关系。需要在页眉、页脚、脚注、尾注中修复的内容也有独立检查项，不遗漏正文外部件。

旧的全文修复脚本不作为每块的命令。Agent 可调用现有函数处理选中对象，或对工作文档的选中节点作最小修改；不能每处理一块都先跑一次全文批量修复，再尝试通过范围保护。

处理到含图块时，再执行该块图片初检：

```bash
python scripts/review_pipeline.py --work-dir work --action inspect-initial --vision-worker-config worker.json
```

没有自动 worker 时省略配置，转宿主看图任务；首次明确传入的已授权配置会保存并由后续步骤继承。需要修图内内容时，初检后先执行 `--action image-plan`，修改后执行 `--action inspect-repaired --proposal block-fixed.docx`。

实际修复并复核当前块后提交拟稿；块外误改、块内业务文字误改、实际字体字号不合规或域边界损坏时，程序拒绝拟稿，保留当前正确工作文档。

```bash
python scripts/review_pipeline.py --work-dir work --action checkpoint --proposal block-fixed.docx --note "本块实际问题、修改和复核结果"
```

无需修改的块同样需要检查，但直接提交当前文档，不复制一份新的完整 DOCX。程序返回下一块，Agent 连续推进；块完成不是最终分页已经通过。已登记的图内缺陷若图片仍未改变，不能推进成已修复。

## 题注和脚注计划

缺题名、对象豁免由 Agent 根据原文对象决定。`--object-plan` 必须绑定最初原文，提交部分对象时按 ID 累积保存；后续块不会丢失前面补题注的许可。ID 和哈希取自程序清单，不能手猜。

脚注内容迁移使用现有 `content_integrity.register_footnote_changes` / `validate_movements` 生成和校验准确的原文定位，再通过 `--content-plan` 登记。续跑合并已有题注及脚注保护项，不以候选重新建立原文。最终按当前对象核验豁免；旧的段落位置不能盲目重放。

## 最终验收

所有块完成后再调用：

```bash
python scripts/review_pipeline.py --work-dir work --action final
```

这一步统一实际更新域，执行原有全部必需程序审计并渲染。返回本轮 `manifest`、`candidate` 和 `review` 的真实路径。Agent 按一两页窗口检查全部最终页面，定位全部图片和全部表格；表格不只检查脚本产生的 `review` 候选。

填写实际页面/表格观察后，对需要的图片完成最终核验：

```bash
python scripts/review_pipeline.py --work-dir work --action inspect-final
python scripts/finalize_review.py <返回的manifest路径> <返回的review路径> --out final.docx --json-out release.json
```

图片协议不固定双次终检。原图内容未变且初检真正通过时，可复用图内证据并结合当前页面观察；修改过、原生 Shape、存在未闭合缺陷或证据不匹配的图仍须实际检查。页面/表格未通过时不会自动填写通过。交付仍要求完整程序审计、内容保护、图片和全部页面有效结果。

## 失败与恢复

终检发现问题，只返回具体块，保留其他块的修改。集中处理同一轮的实际失败项，再统一验收。

```bash
python scripts/review_pipeline.py --work-dir work --action reopen --block B0023 --note "真实问题与影响范围"
```

如果确需修正共享编号定义、页面或模板样式，使用 `--action shared --proposal shared-fixed.docx --note "缺陷和影响范围"`。该操作保护正文，不清空块进度，但使最终页面验收失效。

未改文档再次请求 `final` 时返回现有任务，不反复更新域、渲染。引擎或审计步骤失败且执行条件已修复时加 `--retry-final`；已成功并绑定当前输入的步骤可复用，失败步骤不缓存为通过。相同语义输入和执行条件连续两次出现同因错误后停止第三次全量重试；`--retry-final` 不能绕过这个限制。文件改名或重打包不重置计数。命令配置需要调整时使用 `--action configure`，只允许执行引擎、UNO Python 和已授权视觉适配器配置变化；模板/文字规则不是静默切换项。

断点恢复直接对同一 `--work-dir` 继续。系统锁随进程退出释放；旧分块会话补齐缺失的附属对象清单和保存检查能力，不清空已完成正文；以前的 checked 不会被伪装成本次新增局部审计的结果。工作文档被外部改动、原文被替换或规则变化时先定位原因，不能改哈希绕过。

最终返修后，变化页面及相邻边界必须重看。仅在当前图像/页面及规则完全匹配时复用有效观察；不可复用旧的整份交付结论。
