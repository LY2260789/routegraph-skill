---
name: routegraph
description: Generate editable PowerPoint technical route diagrams and reconstruct scientific instrument figures using native shapes through PowerShell and PowerPoint COM. Use for research workflows, paper schematics, exploded views, and reference-based instrument illustrations where topology, component structure and editability matter. Instrument reconstruction supports semantic scenes, cylinders, hollow cylinders, cuboids and detailed panels; photo interpretation remains an Agent review task.
---

# RouteGraph

## Purpose

Create editable PowerPoint technical route diagrams, not flattened screenshots. Prefer PowerShell + PowerPoint COM scripts that draw native PPT shapes, text boxes, connectors, and grouped visual structure so the user can refine the result in PowerPoint.

RouteGraph optimizes for correct scientific and workflow logic first. Visual polish is secondary to arrow topology, readable layout, stable text, and editability.

## Instrument reconstruction mode

For instrument diagrams, exploded views, or instrument photographs, read [references/instrument_reconstruction.md](references/instrument_reconstruction.md). Use the scene compiler and existing COM renderer described there. Keep the technical-route workflow below for ordinary diagrams.

- Inspect accessible references before creating a scene. Separate visible structure, justified simplification and unverified inference. Do not render an unknown interior as established fact.
- Use semantic components before geometry. Cylinder, HollowCylinder, Cuboid and Panel are implemented. NativeReference preserves selected editable components from an existing PPTX for incremental migration; it is not an image or SVG fallback.
- For an actual instrument photograph, read [references/photo_exterior.md](references/photo_exterior.md). Keep visible exterior features and record uncertainty; do not merge different panel configurations from multiple photographs.
- Scientific Metallic is implemented. Scientific Flat and Anime Tech are reserved interfaces, not completed presets.
- Render previews locally, inspect them and run structural checks. Apply explicit local scene patches when needed. Ask for a screenshot only if local rendering is unavailable.
- Routegraph owns scientific structure, parametric geometry and native COM rendering. General presentation tools may validate the resulting file; they do not need to redraw it with a second engine. See [the audited call chain](references/reconstruction_audit.md).
- Do not claim automated OCR, photo segmentation, internal reconstruction, physical CAD accuracy, or reference fidelity beyond the available evidence.

## Workflow

1. Understand the input.
   - For reference images, read `references/reference_reconstruction.md` for transparency, photographic insets and source-parameter handling.
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
   - Export a PNG with PowerPoint COM and inspect it when local rendering is available. Otherwise ask the user to return a screenshot.
   - On revision, inspect arrows first, then layout, then text and styling.
   - If revising from a screenshot, read `references/prompts/screenshot_revision.md`.

## Critical Rules

- Arrow logic has highest priority. Wrong source, target, direction, missing critical arrows, or mismatched arrow style are high-severity failures.
- Draw each logical arrow as one continuous PowerPoint shape whenever possible.
  - Use `Add-LineArrow` for direct straight arrows.
  - Use `Add-ElbowArrow` for simple L-shaped or Z-shaped routes.
  - Use `Add-PolylineArrow` for controlled multi-bend routes.
  - Do not build one logical arrow from several independent `Add-LineSegment` calls plus a final arrowhead unless there is a true branching junction.
  - For dashed arrows, this is mandatory: a dashed route must normally be one shape so the dash pattern is continuous across bends.
- Avoid this old segmented pattern for a single logical arrow:

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

- Use helper functions consistently: `RGB`, `Add-TextBox`, `Add-Box`, `Add-LineArrow`, `Add-ElbowArrow`, and `Add-PolylineArrow`. Use `Add-LineSegment` only for non-arrow helper lines or true branching trunks.
- Cast PowerPoint COM line weights and freeform coordinates to `[single]` so generated scripts remain compatible with PowerShell 7.
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
- `assets/templates/continuous_arrow_regression.ps1`: focused regression deck for straight, elbow, multi-bend dashed, and true-branch routes.
- `assets/style_presets/manuscript_clean.json`: restrained manuscript-style colors and typography.
- `assets/examples/`: lightweight examples for expected use cases.

## Output Contract

When generating a diagram script, include:

1. A brief interpretation of the intended diagram.
2. The arrow topology table.
3. The PowerShell script or a path to the generated `.ps1` file.
4. The run command.
5. The rendered preview and review outcome, or a request for a screenshot when local rendering is unavailable.
