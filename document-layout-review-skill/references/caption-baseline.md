# 题注与自动编号基准

**图下表上；图表自动编号须能持续刷新。模板不能覆盖这两条。**

## 位置与样式

图题是图片下方独立段落，表题是表格上方独立段落。模板只决定题注字体、字号、对齐、行距及段间距。模板示例放反或没有示例时，仍执行图下表上，不再让用户选择位置。旧 `caption_placement` 字段不再具有覆盖权。

## 原生序号

图和表使用独立 SEQ 序列，各自在一级章节从 1 递增，显示“图1 题名”“表1 题名”，不带章节前缀。可见域使用 `SEQ 图 \* ARABIC \s 1` / `SEQ 表 \* ARABIC \s 1`。保留章级隐藏重置 `SEQ 图/表 \r 0 \h`，不把重置绑在第一张图表上。

标题文字、样式显示名及标题编号显示格式不是题注域的参数。题注不得拼 STYLEREF、REF/PAGEREF 章节前缀，不得带失效书签参数、手写数字、逐项固定重启、锁域或伪造域结果。重命名标题、改变标题编号字形或正常增删首张图表后，编号仍须能够重新计算。

修复仍由原生编号核心执行。包装层不再反向移动题注，也不再过滤核心的 `caption-position` 错误；新增 `caption_field_guard.py` 验证简单域和复杂域的完整性及指令，不替代实际 Office 更新。

## 刷新流程

更新前核对原生指令和域边界，此时允许待更新的旧数字缓存；更新后检查实际章内序号和错误结果。`field_refresh.py`、编号审计及最终交付均接入检查。

若 Office 导出时改写了 SEQ 开关，刷新组件只按当前源文件已验证的原生指令恢复开关，保留引擎实际计算的结果。匹配依据包括段落非域文字、域次序及序列类型；数量、对应关系不一致或引擎算出的数字不正确时直接失败。记录引擎原始输出哈希、指令恢复清单和最终哈希，不能把这个过程说成未修改过的 Office 原始输出，更不能自行填写假缓存冒充 F9。

最终文件有其他修改时，仍需重新更新、检查与渲染，不得复用旧验收。

## 回归

```bash
python -m unittest discover -s tests -p 'test_caption_baseline_regression.py' -v
```

默认执行结构与接入测试。真实引擎测试须显式启用；下例在 Windows PowerShell 中使用 Word，不会回退到 LibreOffice：

```powershell
$env:DLR_RUN_OFFICE_TESTS = "1"
$env:DLR_FIELD_ENGINE = "word"
python -m unittest discover -s tests -p 'test_caption_baseline_regression.py' -v
```

真实引擎案例包括标题改名、标题样式改名、标题编号字形变化、删除/插入首张图和表、复杂域改名；每例先刷新，再修改，再连续刷新两次。Word 结果只能由实际 Word 运行证明，LibreOffice 不等同于 Windows F9 验证。
