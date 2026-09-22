# Mutation Scope Contract

本文件定义 `document-layout-review` 的对象级修改白名单。它的目的不是告诉 Agent “最好不要改坏其他对象”，而是把“非目标对象不得变化”变成可验证的交付条件。

## 1. 基本原则

每轮修复只能声明一个 `repair_scope`。

修复前：

```bash
python3 scripts/layout_invariant_guard.py snapshot input.docx --out guard.json
```

修复后：

```bash
python3 scripts/layout_invariant_guard.py compare guard.json candidate.docx --scope <scope>
```

比较失败时，candidate 不得交付，也不得继续作为下一轮输入。

## 2. Scope 白名单

| scope | 本轮允许变化 | 必须保持不变 |
|---|---|---|
| `text-style` | 普通正文/标题 run、paragraph 样式 | 可见文字内容、表格结构与几何、图片/Shape、媒体、关系、页面/分节 |
| `text-content` | 明确命中的字符内容 | run/p 样式、表格结构与几何、图片/Shape、媒体、关系、页面/分节 |
| `table-text-style` | 表格单元格内 run/p 样式 | 可见文字内容、表格结构与几何、表外文字样式、图片/Shape、媒体、关系、页面/分节 |
| `table-layout` | 表格布局/几何属性 | 可见文字内容、文字样式、表格语义结构、图片/Shape、媒体、关系、页面/分节 |
| `image-layout` | 被命中的图片/Shape 布局及其媒体/关系 | 可见文字内容、文字样式、表格结构与几何、页面/分节 |
| `flowchart-redraw` | 被命中的流程图 drawing/media/relationship | 可见正文内容、文字样式、表格结构与几何、页面/分节 |
| `page-layout` | section/page 属性 | 可见文字内容、文字样式、表格结构与几何、图片/Shape、媒体、关系 |

禁止 `scope=all`。

## 3. 表格语义结构始终受保护

任何 scope 都不得无授权改变：

- 表格行数；
- 每行单元格数量；
- `gridSpan`；
- `vMerge`；
- 单元格顺序。

`table-layout` 只允许解决几何/分页，不允许改变表格业务结构。

## 4. 文字修复时的冻结项

`text-style` 和 `text-content` 期间，至少冻结：

- `w:tblPr`；
- `w:tblGrid`；
- `w:trPr`；
- `w:tcPr`；
- `w:drawing`；
- `w:pict`；
- `word/media/*`；
- 相关 relationship；
- `w:sectPr`。

因此，修字体/空格后如果表格列宽、图片宽高、图片位置、页面方向等发生任何变化，必须判定当前候选 FAIL。

## 5. 表格文字与表格布局必须拆开

禁止在同一轮同时“统一表格文字样式 + 自动重排列宽”。

先执行 `table-text-style`：

- 修字体/字号/段落；
- 首行缩进 = 0；
- hanging = 0；
- 表格几何完全冻结。

通过 guard 后，若表格本身仍有布局问题，再单独执行 `table-layout`。

## 6. 图片默认冻结

图片不是普通正文装饰，不得在文字修复中被自动归一化。

只有问题清单明确命中的图片，才允许进入 `image-layout`。

默认禁止：

- “所有图片统一宽度”；
- “按版心批量缩小全部图片”；
- 因字体/表格修改顺带重算图片 extents；
- 重写 DrawingML/VML 造成 anchor/inline 变化。

## 7. 最小修改

优先顺序：

1. 修改单个属性；
2. 修改单个 run/paragraph/table cell/drawing；
3. 修改单个表格或图片；
4. 只有确实必要时才重建一个对象；
5. 禁止为了局部问题重建整个 document.xml。

## 8. 失败处理

Guard FAIL 时：

- 不得说“虽然有变化但视觉没问题”；
- 不得把变化视为“自动优化”；
- 不得继续下一轮；
- 必须回滚或从上一份 PASS 候选重新修。

唯一例外是用户明确要求同时改变该非目标对象，此时应结束当前轮并重新声明新的 scope。
