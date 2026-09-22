# Source projects and adopted ideas

This skill is an original implementation. It borrows **workflow ideas**, not source code, from the following projects reviewed during design.

## Anthropic `skills` / DOCX skill

Repository: https://github.com/anthropics/skills

Ideas adopted:

- a DOCX is not complete merely because it was generated successfully;
- render the final DOCX and visually inspect it;
- layout-sensitive tables need explicit, consistent width handling;
- document generation and document verification are separate steps.

The reviewed skill is not treated as copyable source. No implementation text or scripts from it are vendored here.

## `Yuanfang-Wyy/standard-word-doc-skill`

Repository: https://github.com/Yuanfang-Wyy/standard-word-doc-skill

Ideas adopted:

- separate document audit from document repair;
- treat a provided template as the visual authority;
- normalize repeated empty paragraphs and other deterministic formatting defects before visual review;
- use explicit table-format rules rather than relying on arbitrary model styling.

## `grapeot/docx-skill`

Repository: https://github.com/grapeot/docx-skill

Ideas adopted:

- inspect document structure before editing;
- preserve visible layout rather than assuming equal character count means equal visual width;
- render after layout-sensitive edits instead of trusting raw text/XML alone.

## `kieler/elkjs`

Repository: https://github.com/kieler/elkjs

Ideas adopted:

- use a dedicated graph-layout engine to compute node positions and routed edges;
- layered layout is a strong fit for directed business processes;
- rendering/styling should remain separate from graph layout.

elkjs is not bundled by this skill. If a project chooses to depend on it, follow its current license and distribution requirements.

## `dagrejs/dagre`

Repository: https://github.com/dagrejs/dagre

Ideas adopted:

- a lightweight directed-graph layout engine can be used when the project does not need ELK's richer routing features.

Dagre is not bundled by this skill.

## Graphviz `dot`

Repository: https://github.com/rossbar/graphviz

Ideas adopted:

- use a dedicated layered graph layout for directed processes;
- use orthogonal routed connectors for business process diagrams;
- separate semantic graph data from computed geometry;
- validate produced geometry instead of assuming a renderer output is readable.

The skill invokes an installed `dot` executable when available; Graphviz is not vendored in the skill bundle.
