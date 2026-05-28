---
name: routegraph
description: Generate editable PowerPoint technical route diagrams from paper figures, screenshots, sketches, or text descriptions using PowerShell and PowerPoint COM. Use when Codex needs to reproduce or draft research route maps, multi-panel technical workflows, AI/ML method diagrams, counterfactual analysis frameworks, or revise an existing generated diagram from a screenshot with special attention to arrow topology, routing, labels, and editable PPT output.
---

# RouteGraph

## Purpose

Create editable PowerPoint technical route diagrams, not flattened screenshots. Prefer PowerShell + PowerPoint COM scripts that draw native PPT shapes, text boxes, connectors, and grouped visual structure so the user can refine the result in PowerPoint.

RouteGraph optimizes for correct scientific and workflow logic first. Visual polish is secondary to arrow topology, readable layout, stable text, and editability.

## Workflow

1. Understand the input.
   - For an uploaded figure or sketch, identify panels, reading order, text blocks, boxes, model blocks, result blocks, background regions, and arrows.
   - For a text-only request, infer the same structure before writing code.
   - If the diagram is complex, read `references/prompts/image_to_structure.md`.

2. Output a structure plan before code.
   - Always include an arrow topology table before generating a full script.
   - Use this format:

```text
Arrow ID | From | To | Style | Meaning | Priority
```

3. Generate a PowerShell COM script.
   - Use native editable PPT elements.
   - Use absolute coordinates.
   - Reuse helper patterns from `assets/templates/*.ps1`.
   - Keep PowerShell special symbols stable, especially Greek letters and subscripts.
   - Do not paste the entire source figure as a background image unless the user explicitly asks.

4. Provide run instructions.

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\routegraph_generate.ps1
```

5. Iterate from screenshots.
   - Ask the user to return a screenshot of the generated PPT.
   - On revision, inspect arrows first, then layout, then text and styling.
   - If revising from a screenshot, read `references/prompts/screenshot_revision.md`.

## Critical Rules

- Arrow logic has highest priority. Wrong source, target, direction, missing critical arrows, or mismatched arrow style are high-severity failures.
- Prefer segmented routes for complex arrows:

```powershell
Add-LineSegment
Add-LineSegment
Add-LineArrow
```

- Prefer editable plain text for chemical/subscript-heavy labels:
  - `PM2.5` instead of `PM2.5` with rich text subscripts unless explicitly requested.
  - `NO2`, `SO2`, and `O3` instead of Unicode subscripts.
- Generate standalone special symbols through PowerShell code:

```powershell
$Delta = [char]0x0394
Add-TextBox $slide "$Delta PM2.5 air quality" 340 213 190 22 13 $true $black
```

- Use helper functions consistently: `RGB`, `Add-TextBox`, `Add-Box`, `Add-LineSegment`, and `Add-LineArrow`.
- Keep generated scripts easy to modify. Prefer semantic coordinate variables for complex diagrams.

## Resource Guide

Read only the resource needed for the task:

- `references/powershell_com_conventions.md`: PowerPoint COM helper conventions, layering, colors, fonts, and coordinates.
- `references/text_normalization_rules.md`: text, symbols, chemical labels, and encoding-safe notation.
- `references/arrow_topology_review.md`: arrow review checklist and topology failure modes.
- `references/second_pass_review.md`: screenshot-based revision process.
- `references/prompts/structure_to_powershell.md`: guidance for converting a structure plan into PowerShell.
- `assets/templates/four_panel_route_template.ps1`: starter for multi-panel technical route diagrams.
- `assets/templates/horizontal_pipeline_template.ps1`: starter for simple horizontal pipelines.
- `assets/templates/pm25_counterfactual_framework.ps1`: detailed PM2.5 counterfactual framework example.
- `assets/style_presets/manuscript_clean.json`: restrained manuscript-style colors and typography.
- `assets/examples/`: lightweight examples for expected use cases.

## Output Contract

When generating a diagram script, include:

1. A brief interpretation of the intended diagram.
2. The arrow topology table.
3. The PowerShell script or a path to the generated `.ps1` file.
4. The run command.
5. A request for the generated PPT screenshot so RouteGraph can perform the second-pass review.
