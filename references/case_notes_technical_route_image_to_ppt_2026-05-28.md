# RouteGraph Case Notes: Technical Route Image to Editable PPT

Date: 2026-05-28

## Case Summary

User requested converting a local image into an editable PowerPoint route diagram.

- Input image: `<user-home>\Downloads\technical_route_graph.png`
- Output PPT: `<user-home>\Desktop\technical_route_graph.pptx`
- Generated script: `<workspace>\routegraph_generate_technical_route.ps1`
- User feedback: "效果非常不错，还原度接近95%"

The source image was a four-panel horizontal technical route figure labeled A-D. The final PPT was generated with native PowerPoint COM shapes, text boxes, and connector lines rather than using the original image as a flattened background.

## What Worked Well

The routegraph workflow was effective for this case because it prioritized structure before drawing:

- The image was inspected first to identify panels, titles, model blocks, data boxes, background regions, and arrows.
- Arrow topology was explicitly identified before script generation.
- Native editable PowerPoint elements were used for all major objects.
- Blue solid model-loop arrows and purple dashed cross-panel arrows were reproduced with separate connector segments.
- Absolute coordinates were sufficient for a close visual match because the source figure was a clean horizontal scientific workflow.

The generated output reached approximately 95% perceived fidelity according to user feedback.

## Diagram Structure Captured

Panel A:

- Title: training a model with inputs from non-intervention period.
- Left input stack: emission proxies, meteorological variables, chemical indicators, PM2.5.
- XGBoost model block to the right.
- Blue solid loop arrows connecting the model with the input/output stack.

Panel B:

- Title: predicting PM2.5 for intervention period using the trained model.
- BAU simulations label and downward arrow.
- Counterfactual concentration box.
- Observed concentration box.
- Bottom label: Delta PM2.5 air quality.

Panel C:

- Title: re-training PM2.5 model for all measurements.
- Input stack similar to Panel A, with counterfactual PM2.5 as the bottom input.
- XGBoost model block to the right.
- Blue solid model-loop arrows.

Panel D:

- Title: predicting PM2.5 under fixed meteorological states.
- Counterfactual emission strength box.
- Emission strength box.
- Bottom label: Delta Emission strength.

## Arrow Topology

```text
Arrow ID | From | To | Style | Meaning | Priority
A1 | A XGBoost | A Emission proxies / PM2.5 | Blue solid loop | Model input-output relation | High
B1 | A XGBoost | B Counterfactual conc. | Purple dashed | Trained model predicts counterfactual concentration | High
C1 | B Counterfactual conc. | C Counterfactual PM2.5 | Purple dashed routed | Air-quality delta feeds retraining input | High
D1 | C XGBoost | D Counterfactual emission strength | Purple dashed branch | Fixed-meteorology prediction | High
D2 | C XGBoost | D Emission strength | Purple dashed branch | Fixed-meteorology prediction | High
```

Important implementation detail:

- For complex arrows, manually segmented connectors gave better control than a single connector.
- The B-to-C purple dashed route was handled as horizontal, vertical, then horizontal arrow segments.
- The C-to-D output split was handled as a dashed trunk with two dashed branches.

## Implementation Choices

PowerPoint dimensions:

- Single-slide 16:9-style canvas, set to `1028 x 251` points to match the wide source image proportions.

Editable objects:

- Text boxes for panel labels, panel titles, BAU label, and bottom delta labels.
- Rectangles for input/output/model blocks.
- Light background rectangles for grouped regions.
- PowerPoint connector lines for arrows.

Text normalization:

- Used plain `PM2.5` instead of rich-text subscripts for stability.
- Used PowerShell `[char]0x0394` for `Delta`.
- Avoided relying on non-ASCII-heavy script text except where needed by user file paths.

Layering:

- Background blocks were sent behind other objects.
- Model blocks were drawn after internal arrows where overlap control was useful.

## High-Value Patterns To Promote Into RouteGraph

1. Keep the "image inspection -> topology table -> script generation" sequence mandatory.
2. For clean scientific figures, use a canvas ratio close to the image aspect ratio before drawing.
3. Reproduce panel groups first, then text hierarchy, then arrows.
4. Draw cross-panel arrows after all panel blocks so routes can be visually aligned to final object positions.
5. Use semantic coordinate clusters for each panel, even if coordinates are absolute.
6. Prefer simple editable text for chemical labels unless the user explicitly requests typographic subscripts.
7. Save the generated script next to the working directory so the user can iterate without rebuilding from scratch.

## Gaps And Future Improvements

The current output was strong but not a perfect clone. The remaining 5% likely comes from:

- Fine spacing differences between panels.
- Minor differences in box dimensions and background region sizes.
- Arrow turn points, especially the long B-to-C dashed connector.
- Font metrics and exact line wrapping compared with the source image.
- Subscript rendering for `PM2.5`, if publication-quality typography is required.

Recommended RouteGraph upgrades:

- Add a reusable four-panel horizontal scientific workflow template.
- Add helper functions for "input stack + model + blue loop arrows".
- Add helper functions for dashed cross-panel routed connectors with branch outputs.
- Add a post-generation checklist focused on:
  - arrow source and target correctness;
  - route turn-point alignment;
  - panel title wrapping;
  - bottom label placement;
  - background block sizing.
- Consider an optional image-as-transparent-reference mode during script development, removed before final save.

## Suggested Acceptance Criteria For Similar Tasks

For image-to-PPT route diagrams, consider the output acceptable when:

- All major panels and their reading order are preserved.
- All critical arrows have the correct source, target, direction, color, and dash style.
- Text is editable and readable.
- Boxes and background regions are editable PowerPoint shapes.
- The original raster image is not used as the final diagram background.
- The generated PPT opens successfully and is non-empty.
- User-perceived fidelity is at least 90% for first pass.

## Reusable Prompt Notes

When handling similar user requests, RouteGraph should explicitly say:

```text
I will create editable native PowerPoint shapes rather than placing the source image as a flat background.
I will inspect the image, extract panel structure and arrow topology, then generate a PowerShell COM script.
```

Before scripting, include an arrow topology table in this form:

```text
Arrow ID | From | To | Style | Meaning | Priority
```

After generation, report:

- output PPT path;
- script path;
- whether the file was verified on disk;
- a request for a screenshot if the user wants second-pass refinement.
