# Flowchart inspection and redraw rules

## 本项目图面与题注约束

本规则也覆盖架构图、技术路线图和分层结构图，但不得把无箭头的架构图强行转成有向流程并补造节点/箭头。先查看真正的图像或原生 Shape，不以节点 JSON 的审计结果代替图面检查。

图内文字出框、重叠、裁切必须修；原配色和图形风格正常的部分保留。图片输出只含业务图形，禁止写入“图2.1 …”等外部编号和题注。题注由 Word 独立段落与 SEQ 域承担。检测到已烧录题注时，只清除已确定的题注区域或从可靠源重出图；不能裁掉业务内容，也不能把 Word 页上的正常图题误认为图像内容。

初检、修复、终检按 `caption-and-image-review.md` 逐对象闭环；缺检查或初检硬缺陷未修复，不得通过。模型必须真实查看图片；脚本只验证证据覆盖、哈希与闭环，不会从栅格图片自动推断所有文字边界。


## Default assumption: the source is an image

Assume a flowchart is a raster image, screenshot, scanned figure, or embedded picture unless an editable graph specification is explicitly available.

Being an image is not a reason to redraw. First distinguish defects inside the
figure from insufficient space on the document page.

For a sound source figure:

```text
whole source image
  -> keep original media and visual design
  -> proportionate placement/size adjustment if needed
  -> give figure + caption a dedicated page if still crowded
  -> render and verify readability and surrounding text
```

Do not shrink text below readable size to keep the current page count. A dedicated
page keeps the original paper orientation/margins and keeps the caption with its
figure. Use pagination flags rather than inserting piles of blank paragraphs.

Only an intrinsic defect that placement/pagination cannot fix, or an explicit
user redesign request, permits semantic reconstruction and redraw. A clear,
intentional source diagonal connector is not by itself a defect. Do not guess
unreadable source text; obtain a clearer source or confirmation.

Read `flowchart-redraw-contract.md` and `flowchart-drawing-standard.md` before reconstructing an image flowchart.

## Whole-image understanding comes first

Never start by reading one box at a time and independently guessing its outgoing arrow.

First inspect the complete diagram and establish:

- overall purpose;
- main reading direction;
- start/end;
- phases or groups;
- main path;
- decision branches;
- merge points;
- loops and return paths.

If text is too small, use crops only after the global pass. Keep the global topology in context while reading each crop.

## Separation of responsibilities

The reconstructed flowchart source owns semantics:

- node identity;
- node type;
- node text;
- directed edges (`from -> to`);
- branch labels;
- group/swimlane membership;
- intended top-to-bottom or left-to-right orientation.

The layout engine owns:

- node coordinates;
- ports/attachment points;
- edge bends;
- crossing reduction;
- spacing;
- overall canvas size.

The vision model owns source-image understanding and final semantic/visual comparison. It must not use visual judgment as a substitute for a structurally valid graph.

## Structural checks

Before rendering:

- all edge endpoints exist;
- duplicate node IDs and duplicate edges are rejected;
- start nodes have no unexplained inbound edges;
- end nodes have no unexplained outbound edges;
- required nodes are reachable from a start;
- required nodes can reach an end;
- decision nodes expose the expected branch count;
- two-way or cyclic links are deliberate rather than accidental;
- branch labels are present when the source image contains them;
- the semantic spec accounts for every visible source-image node and connector.

Use:

```bash
python3 scripts/flowchart_audit.py <flowchart.json> --json-out <flowchart-audit.json>
```

## Redraw only after the intrinsic-defect gate

The rules below apply only after redraw has been justified. They do not prohibit
proportionate scaling or dedicated-page placement of an otherwise sound image.
Record the actual defect and why placement/pagination cannot solve it. Rebuild
from the verified graph without changing the source's good visual design.

Do not:

- erase and repaint one arrow on the original bitmap;
- move a few boxes while retaining broken connector geometry;
- globally shrink the original image to hide overflow;
- preserve bad coordinates just because they came from the source picture.

Do:

1. reconstruct semantics;
2. calculate new node dimensions from text;
3. calculate a fresh layout;
4. route connectors from source boundary to target boundary;
5. redraw all nodes, labels, arrows, and group boundaries consistently.

Preserve the source palette, fills, border style, square/rounded corners, font
style, grouping and reading direction unless explicitly authorized to change.
A layout engine's defaults are not a source style. The bundled renderer can
produce a candidate, but its default palette must not overwrite an existing
figure. When necessary, use a source-style-capable drawing implementation.

The bundled renderer command is:

```bash
python3 scripts/flowchart_redraw.py <flowchart.json> --out-dir <redraw-dir> --engine auto
```

`auto` prefers Graphviz `dot` when available and runs a geometric validation pass over the routed result. A host-provided ELK/elkjs layout is also valid. The dependency-free fallback is restricted to simple acyclic graphs. If the graph contains multiple decisions, loopbacks, dense branching, or other complex routing and no real layout engine exists, the redraw must be blocked rather than emitted with guessed connectors.

## Layout engine

For non-trivial directed business processes, use a layered graph-layout algorithm. ELK/elkjs and Graphviz `dot` are suitable references because they support directed layout and routed edges. Dagre is a simpler alternative when advanced edge routing is unnecessary.

Recommended behavior:

- TB orientation => main flow top to bottom;
- LR orientation => main flow left to right;
- orthogonal routed connectors for ordinary business-process diagrams;
- edges attach at node boundaries/ports;
- reserve spacing for edge labels;
- measure/wrap text before final node size is sent to the layout engine.

## Text sizing

Do not use a fixed node height when text can grow arbitrarily.

Before layout:

1. normalize node text;
2. wrap to a deliberate line length or measured width;
3. compute node width/height from the wrapped lines and font metrics;
4. add internal padding;
5. pass the real size to the layout engine.

## Source-versus-redraw verification

After redrawing, inspect the **source image and the complete redrawn image together**.

Check semantic equivalence before visual polish:

- all source nodes exist;
- all source arrows exist;
- arrow direction matches;
- branch labels match;
- loop and merge topology matches;
- no invented steps or edges appear.

Then compare the original and new complete figures side by side at comparable display size. Confirm the source style is retained, not merely that the new graph is valid. Check visual quality:

- arrowhead points toward the intended target;
- connector terminates at the correct node boundary;
- connector does not pass through an unrelated node;
- unexplained diagonal connectors are absent;
- connector crossings are absent unless explicitly justified and visually disambiguated;
- loopback/retry/整改 edges use outer corridors instead of cutting through the main flow;
- edge labels do not overlap nodes, other labels, or arrowheads;
- node text stays inside the border with padding;
- nodes do not overlap;
- branches are visually distinguishable;
- the diagram is not clipped by the target page;
- global shrinking does not make text unreadable when a dedicated page or necessary re-layout could solve the problem.

If the source is ambiguous, do not guess. Preserve the uncertainty and request a clearer source or confirmation.
