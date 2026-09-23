# 图表补注与真实图片验收

## 1. 先看真实对象

运行 `inspect_document.py` 得到图表对象清单、提取的原图和待填写的初检模板。图表 ID、位置和哈希由程序计算，Agent 不手编。不能只搜有向流程图 JSON，漏掉图片中的架构图、技术路线图和原生 Shape。

先看完整图，再按需放大局部，检查尤其容易漏掉的侧边窄框文字。照片等不适用的检查项须说明原因；架构图没有箭头时核对分层/分组关系，不编造流程连线。原生 Shape 没有 `word/media` 文件时，先渲染原 Word 查看。

原图与整页截图要区分：Word 独立图题出现在页面下方是正常的；只有真正位于图像像素中的外部图号/题注才需要从图像中清理。去除烧录题注须保留业务内容和原风格。缺清晰来源、无法分辨业务内容与题注时先确认，不随便裁图。

## 2. 补充缺注的语义计划

示例中的哈希、ID都从真实清单复制；题名由 Agent 根据实际内容填写。用户不需要填写这个 JSON。

```json
{
  "version": 1,
  "source_sha256": "sha256:当前输入文件哈希",
  "objects": [
    {"id": "F0001", "object_sha256": "sha256:图哈希", "title": "林草分库总体架构图"},
    {"id": "T0001", "object_sha256": "sha256:表哈希", "title": "设施设备配置明细表"}
  ]
}
```

题名用实际业务内容，简洁、中性；不要把所有缺题名都统一叫“示意图”。确定是封面标志/装饰图/布局表时，用 `exempt: true`、`reason: "具体用途"`，而不是给它造图号。旧独立题注未使用题注样式时，可加 `caption_paragraph` 和完整 `caption_text`，绑定原文件中的明确段落，不进行模糊搜索。

计划过期、图表变化、定位不唯一或缺题名时程序阻断，Agent 重新核对，不能直接将缺注记为通过。程序负责图下/表上位置、原生 SEQ、章内计数；不负责凭空理解一张栅格图的题名。

## 3. 图面检查记录

`figure_review.py prepare` 自动给出待检查的原图项目。Agent 查看后填写 `image_type`、`observation`、`checks`；有缺陷填写 fail，不能为了让流程继续把它填成 pass。初检可以有 fail，目的是进入修复。

终检记录必须绑定原文、当前候选、原图、最终图和最新页面哈希。六项检查是：文字在框内、无重叠、可读、不裁切、连接/分组关系正确、图像内没有外部题注。终检还核对语义及原有风格。

初检存在文字出框、重叠、连线错误或烧录题注，而最终图哈希未变时，交付入口拒绝“已修复”。原图只是被压得太小时，独占页/等比放大可以保持原媒体不变；不能为了排版把正常图换成新的美化风格。

`flowchart_audit.py` 检查语义节点与连线，不测量图片像素中的字框。`figure_review.py` 校验实际对象覆盖、哈希新鲜度和缺陷闭环，也不伪装成 OCR 或全自动视觉判断。视觉结论必须来自 Agent 真实看图，宿主权限/工具日志才可能进一步限制伪造检查记录。

## 4. 执行与继续

正常执行：

```bash
python scripts/inspect_document.py input.docx --work-dir work --template template.docx --template-style-json template-style.json
# Agent 看原图、填写初检记录；有缺注时填写 object-plan。
python scripts/review_pipeline.py input.docx --work-dir work --template template.docx --template-style-json template-style.json --initial-visual-review initial-review.json --object-plan object-plan.json
```

流水线生成候选，不直接交付。图像本体缺陷由 Agent 按图修复流程处理，不能只修改报告。修改后的候选先实际更新域，再以只读继续模式审计和渲染，不重跑批量修复来覆盖已经修好的图：

```bash
python scripts/field_refresh.py repaired-candidate.docx --out refreshed.docx --engine word --json-out field-update.json
python scripts/review_pipeline.py refreshed.docx --audit-only --source input.docx --work-dir final-review --template template.docx --template-style-json template-style.json --initial-visual-review initial-review.json --field-update-report field-update.json
```

最后按流水线生成的 `final-visual-review.template.json` 填写真实的逐页、逐图观察和表格 review 决定，再调用：

```bash
python scripts/finalize_review.py final-review/review-manifest.json final-visual-review.json --out final.docx --json-out release.json
```

有豁免图/布局表时，继续模式仍须传对应的最新 `--object-plan`；结构发生变化要重新确认对象对应关系。任一步失败，候选都不是交付成果。最后改动了任何内容、域、样式或图片，都必须重新更新域、检查和渲染。
