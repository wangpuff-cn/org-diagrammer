---
name: org-diagrammer
description: >
  Generate interactive equity structure charts and organizational charts as
  standalone HTML files (or in-conversation Widgets). Use this skill whenever
  the user mentions equity/ownership/shareholding/capital structure diagrams,
  VIE structures, shareholder relationship charts, 股权架构图, 股权结构图,
  股权穿透图, 持股关系图, 画股权, 公司架构图, VIE 结构图, 出资结构, 股东关系图,
  or org/department/reporting-line charts, 组织架构图, 部门架构图, 组织结构图,
  汇报关系图.
  Input: a standard Excel template (equity or org; Chinese and English headers
  both supported — see assets/). Output: an interactive HTML with branch
  collapse/expand, mono⇄color toggle, zoom (incl. fit-to-window), search &
  locate with chain highlight, SVG/PNG/PDF export, and hover details.
  Unlike equity-chart (static matplotlib PNG), this skill produces an
  interactive web version.
---

# org-diagrammer · Interactive Equity / Org Charts

## What it does

- **Dual modes**: equity structure (equity) + organizational chart (org) — one Excel in, one interactive diagram out.
- **Interactions** (built into the front-end template, zero extra work):
  - Badge under each node **collapses/expands** its branch; siblings re-layout to normal spacing; the toggled node stays in view
  - Toolbar: mono⇄color, level-collapse menu (expand all / collapse all / collapse to level N), zoom dropdown (fit-to-window + 50%~1000%), ＋/− step buttons (gesture zoom disabled)
  - Search box with live dropdown (top 10): picking a result locks onto that exact node and highlights the chain **up to the root** with thickened links; auto-zooms (≤500%); typing previews substring matches
  - Click a node: bidirectional highlight (ancestors + descendants), thickened links
  - Export menu: SVG / PNG / PDF
  - Node hover: equity nodes show business & function + shareholders; org nodes show function / leader / headcount / parent
- **Visual language** (aligned with equity-chart): rounded rectangles, orthogonal links (bus centered between layers), solid = holding/direct management, dashed = contractual control (VIE)/dotted-line reporting, black-white default (color when fills exist), listed entities get bold borders.

## Repository layout

```
org-diagrammer/
├── SKILL.md                           ← this file (agent guide, English)
├── README.md / README_EN.md           ← user documentation
├── scripts/
│   └── build_equity_json.py           ← core: Excel → layout JSON / standalone HTML (needs openpyxl only)
└── assets/
    ├── 股权架构图填写模版.xlsx          ← equity template (中文, with example & instructions)
    ├── 组织架构图填写模版.xlsx          ← org template (中文, with example)
    ├── Equity_Structure_Template_EN.xlsx
    ├── Org_Structure_Template_EN.xlsx
    └── widget-template/
        └── index.html                 ← front-end template (placeholders: /*__EQUITY_DATA__*/, /*__CHART_TITLE__*/)
```

## Standard usage (one command)

```bash
python3 scripts/build_equity_json.py input.xlsx \
    --title "Example Corp Equity Chart" \
    --html out/MMDD_example_equity_chart.html
```

- `--mode auto` (default) detects equity/org from sheet names & headers; force with `--mode equity|org`
- `--json out.json`: layout JSON only (to feed a custom front end)
- Without `--json/--html`: JSON to stdout
- Template defaults to `<script_dir>/../assets/widget-template/index.html`; override with `--template`

The output HTML is **fully self-contained and shareable** (data inlined, no external deps); the browser tab title = `--title`.

## Excel parsing rules

Headers are matched **by name, bilingually** (column order is flexible). Both Chinese and English templates work.

### Equity template

- Headers: 节点颜色·Node Color / 层级编号·Level / 公司名称·Company Name / 国家·Country / 省·Province / 市·City / 区·District / 股权关系·Relationship / 境内·境外·Domestic·Overseas / 股东·母公司·Shareholder·Parent / 持股比例·Ownership % / 主体业务及职能·Business & Function / 特殊要求·Notes
- Country/province/city/district join into the node's second region line (`-`/`——` skipped), 12pt
- **Business & Function**: shown on hover; containing 上市/Listed → listed entity (bold border)
- Relationship: 协议控制 / Contractual Control / VIE → dashed; others → solid
- Shareholders: `/`-separated multi-shareholders with matching ratios (fan-in bus)
- Node color column = **cell fill** (theme color + tint auto-resolved); filled → colored node & color mode default; unfilled → white/black
- Notes: free text; supports 境内主体低于境外主体 / "domestic below offshore" (domestic nodes sink)
- >10 siblings under one parent on the same layer → wrap to rows (nodes with children sit on lower rows)

### Org template

- Headers: 节点颜色·Node Color / 层级编号·Level / 部门·组织名称·Department / 负责人·Leader / 组织职能·Function / 人数·Headcount / 管理关系·Management / 上级部门·Parent / 特殊要求·Notes
- Parent column supports formulas (e.g. `=C$2`; cached value read, formula fallback)
- Management: containing 虚线/兼管/协议 or dashed/virtual/VIE → dashed; otherwise solid
- Headcount: formulas OK; or text like "正编 156 外包 69" / "Regular 156 Outsourced 69"
- Function/leader/headcount/parent are **hover tooltip** content, not drawn on the node
- Same-layer equal width (longest name wins, tight fit); >15 siblings → wrap

## Layout & visual parameters (tunable in scripts)

`H_UNIT=42` (node height), name 14pt (auto-shrink to 6pt floor, never wraps), region line 12pt,
width: **same-layer equal width** — the layer's longest-name node sets the width, padding ≈5px (tight, no quantization),
layer gap: 1×height for one-to-one, 2×height for fan-in/out, node gap `0.6×H_UNIT`.
Mechanisms (tidy-tree with subtree footprint = max(own width, children span), sibling wrap, centered buses, orthogonal links) must not change; values may be tuned per chart size.

## Showing as a Kimi Work Widget (optional)

Standalone HTML is the default deliverable. To project into a conversation/dashboard:
1. Create a Widget with the `Widget` tool, write `assets/widget-template/index.html` into its `workspace/index.html`
2. Replace `/*__EQUITY_DATA__*/ null` with the `--json` payload, and `/*__CHART_TITLE__*/架构图` with the title
3. `Widget.show`. **Never** paste the whole HTML into a reply.

## Notes & boundaries

- Only dependency: `openpyxl` (bundled in Kimi Work's managed Python). Reading colors requires loading the workbook twice (styles + cached values)
- Edges referencing missing nodes are dropped
- viewBox is computed from actual node bounds — no cropping
- If you edit `index.html`, keep both placeholder comments or injection fails
- Link highlight attribution: shared segments own multiple endpoint ids joined with `\x01` — do not split on spaces (names contain spaces)
