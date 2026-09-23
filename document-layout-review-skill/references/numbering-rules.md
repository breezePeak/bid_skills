# 标题、题注、脚注自动编号

## 检查和修复

先由 Agent 识别真实标题、题注、脚注及其引用关系。程序检查的是实际结构，不以“看起来有数字”作为通过条件。

- 标题：有效的多级编号定义及段落/样式引用，级别与标题层级一致；手工前缀转换后不得重复显示。
- 题注：真实 SEQ 域，图和表分别计数；分章题注的章节号使用 STYLEREF，序号按对应标题级别重启。普通数字、写成文字的域代码及逐条固定重启不合格。
- 脚注：正文中的 `footnoteReference`、脚注内容中的 `footnoteRef` 及对应内容部件、关系完整。脚注 ID 是关联键，不是可见序号；不能用修改 ID 或普通上标冒充自动编号。

分章题注使用输出编号的 `STYLEREF` 开关，不得引用成标题文字。新建分章序列在章节标题中设置隐藏的 `SEQ` 重启域，题注继续自动递增，避免不支持章节重启开关的渲染器连续跨章计数；不在每条题注上写死起始值。隐藏域不显示额外数字，也不得改变标题文字或版式。

已有正常自动编号保持不变。默认保留原有编号形式，不靠每个对象单独设置起始值掩盖断号。封面标题、目录项、正文列表、参考文献、公式指数不作为章节/脚注批量转换。

```bash
python scripts/numbering_audit.py input.docx --json-out numbering-before.json
python scripts/numbering_repair.py input.docx --out numbered.docx --json-out numbering-changes.json
python scripts/numbering_audit.py numbered.docx --json-out numbering-after.json
```

初检 `inspect_document.py` 和最终 `review_pipeline.py` 已接入编号审计；确定性修复先做编号转换，再做原有文字/样式等处理，避免文字清理使定位失效。模板应用时保留标题原有有效编号，防止样式替换把自动编号清掉。

编号转换使用独立 `automatic-numbering` 修改范围；修复脚本保存后检查表格结构/几何、图片、媒体与分节未被顺带修改。不要把这一步当成只改字符的 `text-content`，也不能因旧文字校验不允许域或脚注部件变化而退回手工编号。此范围检查不代替编号审计和页面复查。

## 需要语义确认的位置

脚本可直接处理已识别标题的常见数字/中文序号、连续及一级分章题注、能唯一配对的“脚注文本样式＋普通上标”，以及原生脚注内部缺失的编号标记。孤立引用、重复映射、混合序列、断号、复杂域、复杂脚注内容等返回待核对，不猜测，不输出声称修好的文件。

未使用标题/题注样式的对象、普通正文中的手工脚注，由 Agent 结合内容确认后，通过可选 `numbering-plan.json` 提供精确位置。能从文件确认的由 Agent 处理；只有无法确认含义时才询问用户。

```json
{
  "source_sha256": "sha256:当前输入文件的SHA256",
  "headings": [
    {"paragraph": 1, "text": "1 项目概述", "level": 1}
  ],
  "captions": [
    {"paragraph": 3, "text": "图1 系统架构"}
  ],
  "footnotes": [
    {
      "reference_paragraph": 4,
      "reference_text": "引用[1]。",
      "start": 2,
      "end": 5,
      "note_paragraph": 9,
      "note_text": "[1] 补充说明"
    }
  ]
}
```

段落编号是当前 `word/document.xml` 中按文档顺序的 `w:p`，从 1 开始，包含表格中的段落，不包含删除/移出修订；文字是 `w:t`，制表和换行各算一个字符。`start/end` 为从 0 开始的左闭右开范围。字段不认识、指纹或文字不匹配时直接拒绝。映射只用于当前输入，不可用于后续已改变的文件；没有特殊位置时无需生成此文件。

```bash
python scripts/numbering_repair.py input.docx --out numbered.docx --numbering-plan numbering-plan.json
python scripts/review_pipeline.py input.docx --work-dir work --text-rules text-rules.json --numbering-plan numbering-plan.json
```

复杂多段脚注、页脚中的模拟脚注、带图片/分节/书签的脚注，以及脚本未识别的编号形式，Agent 仍须按真实关系转换、复查；不能因为脚本未识别就宣称整份文档没有问题。

## 最终验收

修复后重新运行编号审计，再更新域并查看最新渲染页面，检查编号顺序、章节重启、题注引用、脚注顺序和分页。程序设置了打开时更新域的标记，但这不等于已经刷新并视觉验收。存在错误或未处理的待核对项，不得交付或标为通过；修复报告简要说明三类编号实际改动。

结构参考：Microsoft Learn 的 NumberingProperties、FootnoteReference，以及 Microsoft Support 的 SEQ field 文档。
