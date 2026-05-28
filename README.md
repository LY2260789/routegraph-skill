# RouteGraph Skill

RouteGraph is a Codex skill for turning technical route diagrams, paper figures, screenshots, sketches, or structured descriptions into editable PowerPoint diagrams.

It is designed for scientific and engineering workflows where the final output should be native PowerPoint shapes, text boxes, and connectors rather than a flattened screenshot.

## What It Does

- Extracts panel structure, reading order, text blocks, model blocks, result blocks, background regions, and arrows from an input figure or description.
- Requires an arrow topology table before generating code.
- Generates PowerShell scripts that use PowerPoint COM to draw editable `.pptx` files.
- Prioritizes correct arrow source, target, direction, and route over decorative polish.
- Supports iterative refinement from screenshots of the generated PowerPoint.

## Requirements

- Windows
- Microsoft PowerPoint installed
- PowerShell
- Codex with local skill support

PowerPoint COM automation is Windows-specific. The generated diagrams are intended to be opened and edited in PowerPoint.

## Installation

Copy this folder into your Codex skills directory:

```powershell
$target = "$env:USERPROFILE\.codex\skills\routegraph"
New-Item -ItemType Directory -Force -Path $target
Copy-Item -Recurse -Force .\* $target
```

Restart Codex or reload skills if your environment requires it.

## Typical Use

Ask Codex to use RouteGraph on an image or route description:

```text
Use routegraph to convert this technical route diagram into an editable PowerPoint.
```

For a local file:

```text
Read image: C:\path\to\figure.png and output a PPT to the Desktop.
```

RouteGraph should inspect the image, produce an arrow topology table, generate a PowerShell COM script, run it when appropriate, and report the output `.pptx` path.

## Workflow

1. Understand the input diagram.
2. Extract structure and arrow topology.
3. Generate a PowerShell COM script with editable native PPT elements.
4. Save the `.pptx`.
5. Review a screenshot for second-pass fixes.

The required arrow topology format is:

```text
Arrow ID | From | To | Style | Meaning | Priority
```

## Repository Layout

- `SKILL.md`: skill entry point and operating rules.
- `assets/templates/`: reusable PowerShell COM drawing templates.
- `assets/style_presets/`: visual style presets.
- `assets/examples/`: lightweight example notes.
- `references/`: focused guidance for arrows, text normalization, PowerPoint COM, and second-pass review.
- `references/prompts/`: reusable prompts for structure extraction and script generation.

## Notes

RouteGraph intentionally avoids pasting the source image as the final slide background unless explicitly requested. The goal is an editable PowerPoint artifact that users can continue refining.

## License

MIT
