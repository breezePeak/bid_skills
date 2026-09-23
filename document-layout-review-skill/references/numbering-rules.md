# 标题、题注、脚注自动编号

## 模板是标题编号的来源

检查实际模板的 `styles.xml`、`numbering.xml`、style → numId → abstractNum/lvl、层级覆盖及有效示例。采用模板的数字类型、`lvlText`、`suff`、层级及重启约定。编号已经自动化不代表正确：不符合模板仍须修复。

默认模板的实际定义是 decimal、`%1`、`%1.%2`、`%1.%2.%3`，编号后空格。不得把坏文档中的“第一章/第1章”保留下来；也不得把默认模板数字规则强加给使用另一套明确约定的模板。模板有多套冲突定义或无法确认时先解决冲突，不依据源文档猜一个。

恢复真实标题样式，给段落和标题样式绑定同一原生多级列表。复制编号定义时分配不冲突的 ID，不覆盖正文列表的编号实例。标题中的实际业务数字、年份、正文列表、目录和封面标题不是待删序号。目录由 Office 更新，不改成手写数字。

## 本项目图表规则

每个一级章节中，图和表分别从1递增：`图1 题名`、`图2 题名`；`表1 题名`、`表2 题名`。不显示章节前缀，不使用 STYLEREF 拼接图号。每条可见题注分别使用原生 `SEQ 图 \* ARABIC \s 1` 或 `SEQ 表 \* ARABIC \s 1`；在章节边界附加不显示结果的序列重置，不能在每条题注中写死 `\r 1`。

独立图题放图下，表题放表上；题注从图像内容分离。删除失效题注域时同时处理真实域代码和缓存，不只改“错误！未定义样式。”的显示文字。不锁域，不把序号全部扁平化为普通文字。

缺题注时，Agent 看实际图表、上下文和表头，给出简洁、不编造事实的题名；通过 `object-plan.json` 交给程序插入和计数。不要把整个题注画入图中。装饰图、封面标志、布局表经确认用途后可以豁免，理由必须记录。

正文旧图表引用先唯一定位，再按真实先后位置改为“见下图/见上图/见下表/见上表”。引用跨越其他同类对象或有多个候选时，由 Agent 调整对象位置/语句并复核含义，不把“见相关图表”当解决方案；文件名、代码、路径中的图号不是正文引用。

## 接口与流程

```bash
python scripts/numbering_policy.py inventory input.docx --json-out objects.json
python scripts/numbering_audit.py input.docx --template template.docx --template-style-json template-style.json --json-out before.json
python scripts/numbering_repair.py input.docx --out numbered.docx --template template.docx --template-style-json template-style.json --object-plan object-plan.json --json-out changes.json
python scripts/numbering_audit.py numbered.docx --template template.docx --template-style-json template-style.json --json-out after.json
```

不传模板时沿用本项目已确认的默认模板；有用户模板必须一致传入，不得初检用默认、修复用上传、终检又用默认。无缺题注/豁免等语义决定时不必生成 `object-plan`。流水线内部在最后样式步骤后再次按同一规则恢复，避免样式应用把正确编号冲掉。

## 原生脚注保持原有能力

脚注检查正文 `footnoteReference`、脚注内容 `footnoteRef`、部件及关系。ID 是关联键，不是可见序号。普通上标、公式指数、参考文献不能冒充脚注。原生脚注转换仍调用原有 `numbering_core.py`，标题/题注模板规则由新策略层接管；没有重写脚注算法。

对未标注语义的标题/题注及手工脚注仍可使用已有 `--numbering-plan`：

```json
{
  "source_sha256": "sha256:当前输入文件的SHA256",
  "headings": [{"paragraph": 1, "text": "1 项目概述", "level": 1}],
  "captions": [{"paragraph": 3, "text": "图1 系统架构"}],
  "footnotes": [{
    "reference_paragraph": 4,
    "reference_text": "引用[1]。",
    "start": 2,
    "end": 5,
    "note_paragraph": 9,
    "note_text": "[1] 补充说明"
  }]
}
```

段落为当前 `word/document.xml` 中文档顺序的 `w:p`，从1开始，包含表格段落，排除删除/移出修订；文字由 `w:t`、制表和换行组成。范围 start/end 为0起点左闭右开。计划先对原文件验证，再处理脚注，之后重新绑定图表；不能用旧索引错改新位置。复杂或歧义脚注仍需 Agent 确认，不猜配对。

## 真正更新域后再验收

```bash
python scripts/field_refresh.py numbered.docx --out refreshed.docx --engine word --json-out field-update.json
```

Windows Word 引擎更新各 story 的域、目录、图表目录，保存到新文件，禁止覆盖输入；中文/英文域错误直接失败。`updateFields=true` 或手写缓存不算执行 F9。

可以显式选择 `--engine libreoffice` 做替代引擎检查，但不能将其等同于 Word。引擎更新后还要重新跑本项目编号/样式审计；发现其丢失 `\s`、隐藏重置开关或改变模板格式时必须拒绝交付，不能为了通过而降低检查标准。没有适用引擎时状态是待验证，不是假通过。

新增/删除**带有自动题注**的图表后，序列可在更新域时重算。F9 本身不会为一张裸图片自动命名或生成题注；新裸图/表再次执行技能补注，或在 Word 中使用同一“图/表”标签插入自动题注。上/下图表是位置性正文，不是编号域；移动对象跨过引用段落后也须重新核对方向。

`tests/word_f9_integration.py` 用真实 Word 测试更新、增删首图题注及更名标题；增删后不调用修复器，防止用程序重写缓存冒充 F9。非 Windows 平台明确记录 skipped。
