# RouteGraph Skill

RouteGraph is a Codex skill for turning technical route diagrams, paper figures, screenshots, sketches, or structured descriptions into editable PowerPoint diagrams.

It is designed for scientific and engineering workflows where the final output should be native PowerPoint shapes, text boxes, and connectors rather than a flattened screenshot.

## What It Does

- Guides the Agent to identify panel structure, text, components and arrows from an input figure or description; this is not an automatic OCR or segmentation engine.
- Requires an arrow topology table before generating code.
- Generates PowerShell scripts that use PowerPoint COM to draw editable `.pptx` files.
- Prioritizes correct arrow source, target, direction, and route over decorative polish.
- Keeps each logical arrow in one continuous PowerPoint shape whenever the route allows it.
- Uses native elbow connectors or editable freeform polylines for bent and multi-bend arrows.
- Supports iterative refinement from screenshots of the generated PowerPoint.

## Requirements

- Windows
- Microsoft PowerPoint installed
- PowerShell
- Codex with local skill support
- Python 3.10+ (standard library only, for instrument reconstruction and tests)

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

## Continuous Arrow Routing

RouteGraph treats an arrow as a logical object, not a collection of decorative line fragments.

- Straight routes use `Add-LineArrow`.
- Simple L-shaped or Z-shaped routes use `Add-ElbowArrow`.
- Controlled multi-bend routes use `Add-PolylineArrow`.
- Dashed routes should remain one shape so the dash pattern does not restart at every bend.
- Multiple shapes are reserved for true branching junctions or intentional occlusion breaks.

The templates cast PowerPoint COM line weights and freeform coordinates to `[single]` for compatibility with PowerShell 7.

## Repository Layout

- `SKILL.md`: skill entry point and operating rules.
- `assets/templates/`: reusable PowerShell COM drawing templates.
- `assets/templates/continuous_arrow_regression.ps1`: regression deck covering straight, elbow, polyline, and true-branch routes.
- `assets/style_presets/`: visual style presets.
- `assets/examples/`: lightweight example notes.
- `references/`: focused guidance for arrows, text normalization, PowerPoint COM, and second-pass review.
- `references/prompts/`: reusable prompts for structure extraction and script generation.
- `scripts/reconstruction/`: COM-independent scene validation, geometry compilation, local repair and PPTX inspection.
- `scripts/render_instrument.ps1`: native COM renderer for compiled instrument scenes.
- `tests/`: unit tests and Windows/PowerPoint integration regressions.
- `assets/examples/instrument_reconstruction/`: source-labelled migration of the previous editable instrument output, plus a review patch.

## Scientific Figure Reconstruction

The existing route templates are unchanged. Instrument reconstruction is an incremental path:

```text
Agent reference inspection → Instrument Scene JSON → scene.py
→ render plan JSON → render_instrument.ps1 → native PPTX + PNG + render receipt
→ verify.py + visual inspection → explicit local scene patch → new output
```

Implemented geometry: Cylinder, HollowCylinder, Cuboid and Panel. Parameters control dimensions, origin, rotation and material. Named native surfaces form a component group. Scientific Metallic separates steel and polymer materials from geometry. The compatible NativeReference adapter copies selected native PPT shapes for detailed parts that have not been parameterized.

The bundled exploded instrument case is derived from a previous generated PPT, **not from the original paper image**. It cannot establish scientific accuracy against the paper. Photonion photographs were unavailable during Phase 1; a subsequently supplied photograph has now been used for an independent exterior reconstruction. Photo segmentation, automatic OCR, Slotted Box and multi-image integration are not implemented.

Run from this skill folder, choosing unused output filenames:

```powershell
python scripts/reconstruction/scene.py assets/examples/instrument_reconstruction/charging_collection.reviewed.scene.json --out build/instrument.plan.json
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/render_instrument.ps1 -PlanPath build/instrument.plan.json -OutputPath build/instrument.pptx
python scripts/reconstruction/verify.py build/instrument.pptx --scene assets/examples/instrument_reconstruction/charging_collection.reviewed.scene.json --plan build/instrument.plan.json --receipt build/instrument.render.json --out build/instrument.validation.json
```

PowerPoint COM needs a desktop login session. Use `pwsh` if preferred. New commands refuse to overwrite inputs or outputs and close only the presentations they opened.

Tests:

```powershell
python -m unittest discover -s tests -v
powershell -NoProfile -ExecutionPolicy Bypass -File tests/run_regressions.ps1 -OutputDirectory build/regression-001 -PythonExecutable python
```

The integration runner uses a new directory, runs the instrument scene, an independent primitive gallery, the old continuous-arrow template, and a native ungroup/edit/regroup/save/reopen smoke test. It preserves the original legacy template and removes its application-wide Quit call only in the temporary test copy.

See [audit and architecture](references/reconstruction_audit.md), [scene contract and limits](references/instrument_reconstruction.md), and [Phase 1 acceptance report](references/phase1_acceptance.md).

## Photo exterior extension

The shared Cuboid compiler supports either visible side for an above-object oblique view. Panel adds named rectangle/ellipse features for visible controls, openings and grilles; features must remain inside the panel and use material tone keys. This extends the same renderer and leaves legacy scene output unchanged.

See [photo exterior workflow](references/photo_exterior.md). A photo-specific palette is supplied with `scene.py SCENE --style PALETTE --out PLAN`; it remains separate from the component geometry. The library does not infer concealed interiors or instrument readings.

## Update scope and validation

This update is intended for v0.3.0. The core instrument scene compiler, four geometry types, metallic palette, local patches, native renderer and verifier are implemented. Diagram Graph, automatic reference recognition, Scientific Flat and Anime Tech remain future work. Complex optical/flow diagrams currently use Agent-authored native COM scripts through the existing workflow; they do not yet use a general diagram compiler.

There are 26 standard-library unit tests. GitHub Actions runs these on Python 3.10 and 3.12; it does not validate PowerPoint rendering. The separate Windows desktop regression command above exercises COM rendering, old arrow geometry and native editing. The Phase 1 acceptance report is a historical snapshot; later local photo tests and an optical diagram reconstruction do not replace the original paper-reference verification that remains outstanding.

User photographs and case-specific output decks are not bundled. The small bundled native PPTX is a generated migration fixture, not the original paper figure. See [reference reconstruction guidance](references/reference_reconstruction.md) for photographic insets, transparency and editable-object claims.

## Notes

RouteGraph intentionally avoids pasting the source image as the final slide background unless explicitly requested. The goal is an editable PowerPoint artifact that users can continue refining.

## License

MIT
