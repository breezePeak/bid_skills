# Image flowchart redraw contract

## Purpose

Most real flowcharts arrive as raster images or screenshots, not editable graph data. The repair workflow therefore reconstructs the complete semantic graph from the **whole original image** and redraws a new diagram from that graph.

Do not treat the original image as an editable canvas. Do not patch individual arrows, move a few boxes, or shrink the image and call it repaired.

## Whole-image-first rule

Before extracting any individual node, inspect the complete flowchart image at a scale that shows the whole diagram.

Record the global picture first:

- title and purpose, if visible;
- main reading direction: top-to-bottom or left-to-right;
- start and end regions;
- major phases, groups, swimlanes, or repeated subflows;
- primary path;
- decision branches;
- merge points;
- loops / retries / rollback paths;
- legends or annotations that affect meaning.

Only after that global pass may the model enumerate individual nodes and arrows.

If the image is too large to read text clearly, keep the whole-image observation as context, then inspect crops for text detail. A crop may refine a node or label but must never replace the whole-image understanding.

## Semantic reconstruction output

Reconstruct the image into this JSON shape before redrawing:

```json
{
  "type": "flowchart",
  "schema_version": 1,
  "id": "FLOW-001",
  "title": "流程图标题",
  "purpose": "这张图表达什么",
  "direction": "TB",
  "groups": [
    { "id": "G1", "label": "阶段一" }
  ],
  "nodes": [
    {
      "id": "N1",
      "type": "start",
      "text": "开始",
      "group": "G1",
      "source_text_confidence": "high"
    }
  ],
  "edges": [
    {
      "from": "N1",
      "to": "N2",
      "label": "是",
      "source_confidence": "high"
    }
  ],
  "uncertainties": []
}
```

The redraw script only needs `id`, `title`, `direction`, `nodes`, and `edges`. The other fields preserve image-understanding evidence for review.

Supported node types:

- `start`
- `end`
- `process`
- `decision`
- `document`
- `subprocess`

## No-invention rule

The reconstruction must preserve the source image's meaning.

- Do not add a missing business step because it seems logical.
- Do not reverse arrows to make the graph cleaner unless the source arrow is clearly read in that direction.
- Do not silently rename node text.
- Do not collapse two visually separate steps into one without evidence.
- Do not create Yes/No labels that are not visible or otherwise unambiguous in the source.

When an arrow, label, or node cannot be read confidently, record it in `uncertainties` and keep the repair blocked until the ambiguity is resolved by a clearer image or user confirmation.

## Completeness check before redraw

Before drawing, compare the semantic spec against the source image and confirm:

1. every visible process/decision/start/end box is represented;
2. every visible directed connector is represented;
3. every visible connector label is represented;
4. every merge, branch, and loop has the same topology;
5. every group/swimlane that changes interpretation is represented;
6. no node or edge exists in the spec without source evidence.

Do not redraw until this check passes.

## Redraw acceptance

A redrawn flowchart is acceptable only when both conditions hold:

### Semantic equivalence

- same process steps;
- same directed edges;
- same decision outcomes;
- same merge and loop structure;
- same meaningful labels;
- same phase/group relationships when present.

### Visual quality

- text fits inside nodes;
- arrows attach at node borders;
- arrowheads clearly point toward targets;
- connectors do not pass through unrelated nodes;
- branches and loops are visually distinguishable;
- labels do not overlap boxes or lines;
- the diagram fits the target page without making text unreadably small.

If semantics pass but layout fails, re-layout and redraw. Do not change semantics merely to obtain a prettier picture.

## 标点语境要求

语义重建完成后、进入布局前，标题、节点和边标签必须执行上下文标点规范化：

- 中文语境：全角中文标点；
- 英文、代码、版本、路径、URL、标识符和公式：半角标点；
- 混合文本：按局部语境处理。

例如：

- `满足条件?` → `满足条件？`
- `否,整改` → `否，整改`
- `API（v1.2）` → `API(v1.2)`
- `f（x）＝x＋1` → `f(x)=x+1`

## 几何验收要求

语义等价并不代表绘图合格。重绘结果还必须通过几何检查：

- 连线不得穿过无关节点；
- 普通业务流程连接线必须使用正交路由；
- 不得存在无法判定归属的连线交叉；
- 回退、整改、重试、复测等循环边必须走外围回廊；
- 箭头必须落在正确目标节点的边界；
- 判断分支标签必须贴近自己的出边。

复杂流程没有 ELK／Graphviz 等可靠布局引擎时应停止重绘，不允许退化为模型手工猜坐标。
