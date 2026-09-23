# 文字规则接口

## 使用方式

Agent 根据本次用户要求生成 `text-rules.json`。用户不需要手写 JSON；没有额外要求时写 `{}`，沿用默认值。不要为了单次任务改默认文件或脚本。

以下入口均支持 `--text-rules <文件>`：

- `inspect_document.py`：初检；
- `review_pipeline.py`：现有修复流程，会把同一份规则传给模板应用、文字修复和最终审计；
- `template_style_enforce.py`、`template_usage_repair.py`、`template_usage_audit.py`：单独运行时也必须传同一份规则。

```bash
python scripts/inspect_document.py input.docx --work-dir work --text-rules text-rules.json
# Agent 完成全面检查、问题清单和语义判断后，再执行修复。
python scripts/review_pipeline.py input.docx --work-dir work --text-rules text-rules.json
```

原有 `--template`、`--template-style-json` 等参数仍可使用。不传 `--text-rules` 时兼容旧调用，读取 `assets/default-text-rules.json`；用户提出了额外要求时不得漏传。

## 字段

当前接口仅处理普通正文字体、字号、加粗和斜体。表格、标题、题注、图片、空格和标点继续按原规则执行；不能把正文规则扩展到这些对象。

| `body` 字段 | 含义 | 默认值 |
|---|---|---|
| `preserve_fonts` | 允许保留的已有字体；`[]` 清空额外保留项，不关闭字体检查 | `["黑体", "SimHei"]` |
| `forbidden_fonts` | 明确禁止的字体；禁止项会取消默认保留项 | `[]` |
| `fonts` | 指定替代/统一字体；可分别设置 `eastAsia`（东亚文字）、`ascii`、`hAnsi`、`cs` | `{}`，沿用模板 |
| `size_pt` | 正文字号，单位磅，按 0.5 磅递增 | `null`，沿用模板 |
| `bold` | 加粗策略 | `"template"` |
| `italic` | 斜体策略 | `"template"` |

加粗/斜体策略：`template` 沿用原检查逻辑，清除与正文基准冲突的直接格式；`preserve` 保留原状态；`forbid` 禁止（包括样式继承）；`require` 统一启用。保留“黑体字体”不等于保留“加粗”。

只覆盖用户明确提及的字段，未提及的继续沿用默认值。`fonts` 只覆盖指定的字体槽，不替用户决定其他文字的字体；该槽的指定字体优先于保留列表。要求“只能使用某字体”时，须覆盖对应字体槽，而非仅修改保留列表。

`黑体/SimHei`、`宋体/SimSun` 按同名字体处理，也识别大小写和 `@` 前缀；其他字体名称按规范化后的名称比较，不把所有带“黑”字的字体视为黑体。

## 示例

用户要求正文不允许黑体，模板正文为宋体：

```json
{
  "body": {
    "forbidden_fonts": ["黑体", "SimHei"]
  }
}
```

用户要求正文中文统一宋体、12 磅，保留原有加粗：

```json
{
  "body": {
    "fonts": {"eastAsia": "宋体"},
    "size_pt": 12,
    "bold": "preserve"
  }
}
```

如果模板正文也使用了被禁止的字体，且用户没有确定替代字体，停止并询问；确认后才填入 `fonts`。不能在禁止黑体的同时仍按黑体基准修复。

## 执行和复查

规则在修改文档前校验。未知字段、错误类型、同一字体同时明确保留和禁止、指定字体落入禁止列表等情况返回失败，不静默忽略。

初检和修复流程会在 `reports/text-rules.effective.json` 保存合并后的规则；相关审计/修复报告记录 `text_rules_sha256`。流程检查该标识，防止子脚本漏用规则。用户改变要求后，更新任务规则并重新检查、修复和验收，不复用旧结论。

字体检查会读取段落样式、字符样式和直接格式的继承结果。无法解析正在使用的主题字体、又无法确定安全替代值时，说明未能验证，确认明确字体后再继续，不声称通过。

本次覆盖值只写入正文，不修改原模板文件及共享样式定义；模板结构一致性检查仍使用原模板。规则检查通过也不能替代流程要求的最终渲染和页面验收。

## 回归测试

在技能目录运行：

```bash
python -m unittest discover -s tests -p test_text_rules.py -v
```

测试覆盖规则合并、默认行为、用户覆盖、样式继承、字符属性保护、命令行传参和流程规则一致性。测试中的表格/图片等旁路以及渲染使用模拟结果，不能作为真实 Word 页面的验收证据。
