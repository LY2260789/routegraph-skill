# Instrument Scene v1

Read this for instrument reconstruction, not for ordinary technical-route diagrams. Python scripts use only the standard library. All dimensions are visual points (1 pt = 1/72 inch), never measured physical dimensions unless separately established by the source.

## Minimal workflow and responsibilities

1. Agent inspects each accessible image and records components, labels, source regions, assembly and occlusion. The compiler does not recognize an image. If a reference is missing, report it and use an explicitly labelled synthetic or legacy regression only.
2. Author an Instrument Scene JSON. Derive each component from evidence before choosing its primitive. Keep source files unchanged. Use Cylinder/HollowCylinder/Cuboid where appropriate, and native references only for existing editable detail.
3. Run `scripts/reconstruction/scene.py SCENE --out NEW_PLAN.json`, then `scripts/render_instrument.ps1 -PlanPath NEW_PLAN.json -OutputPath NEW.pptx`.
4. Run `scripts/reconstruction/verify.py NEW.pptx --scene SCENE --plan NEW_PLAN.json --receipt NEW.render.json --out NEW.validation.json`. Inspect the PNG at readable size against the reference. Automated checks cover structure and native object properties, not scientific interpretation.
5. Write a reasoned local patch and create a new scene using `scene.py SCENE --patch PATCH --out REPAIRED_SCENE.json`. Recompile, render, inspect and validate. Repair only demonstrated errors; retain unresolved ambiguity in the report.

No presentation-library authoring dependency is required. The COM backend owns native shapes and previews. An installed presentation finalizer can additionally validate the PPTX package and import compatibility, without replacing Routegraph's renderer.

## Scene contract

The authoritative executable validator is `scene.py:validate`. Use the bundled `charging_collection.reviewed.scene.json` as a runnable example.

| Field | Meaning |
|---|---|
| schema_version | `routegraph.instrument.v1` |
| canvas | Width and height in points |
| references | reference_id, source type, path/hash when available; distinguish photo, paper image, prior output, synthetic test |
| assemblies | Named semantic frames. In v1 these are labels for grouping related parts; origins below are already slide coordinates, not a parent transform hierarchy |
| projection | Oblique x_basis, depth_basis, vertical_basis, and ellipse_ratio shared by primitive compilation |
| components | Semantic components listed below |
| annotations | Panel names and section captions as editable text |

| Component field | Meaning |
|---|---|
| component_id / component_type | Stable identifier and scientific role; do not encode a material as its scientific identity |
| geometry_type | Cylinder, HollowCylinder, Cuboid, Panel, or NativeReference |
| geometry_parameters | Dimensions or exact source PPTX shape names |
| bounding_box | Declared axis-aligned visual bounds [x,y,w,h]; checked against compiled geometry |
| relative_position | frame and origin [x,y] in slide coordinates. No implicit transform inheritance |
| rotation | Screen-plane degrees about the local origin; actual group positioning preserves this origin |
| z_order | Draw order. Larger values are in front. Explicit occlusion relations must agree |
| material | Material key in the selected style; native references preserve their encoded source fills |
| label / label_anchor | Plain text and box, arrow start/target, font_size; additional_targets preserves a shared label with multiple leaders |
| relationships | connected_to, contains, adjacent_to, occludes, assembled_with; endpoints checked, occlusion order checked |
| confidence | [0,1] evidence assessment supplied by Agent, not a probability calculated by this library |
| source_reference | reference_id, evidence, note and optional independent reference bounding_box/native shape names |

`visible`, `simplified`, `legacy_output`, `unverified` distinguish evidence. Simplification and unverified inference require an explanation. Unverified geometry is rejected at render compilation: leave it in analysis or an explicit text uncertainty note rather than silently drawing it as established structure. No unknown scientific values, connections or interiors may be supplied by the compiler.

## Geometry and projection boundaries

- Cylinder: length, radius. Named rear cap, outer wall, front face. The front is at local (0,0) and the axis runs along positive x.
- HollowCylinder: length, radius, inner_radius, bore_depth. Adds inner wall and bore shadow. A layered elliptical opening expresses depth; its dark area is a shading convention, not an assertion that the tube has a closed back. It is not a boolean solid model.
- Cuboid: width, height, depth. Front, side and top polygons derive from shared basis vectors. The photo extension selects the visible side from the depth basis x sign. It accepts x toward right/up, depth toward either left/up or right/up and vertical [0,1]. Cameras below the object remain unsupported.
- Panel: width, height and optional features. Each feature has a unique semantic name, kind (rectangle/ellipse), local box [x,y,w,h], and a material tone key. All features remain inside the panel. They express visible surface details, not hidden electronics. Surface and detail shapes form a named editable group.
- NativeReference: path, sha256, slide, shape_names, source_bounds. Selects disjoint top-level native shapes in a PPTX, preserves detail, and permits translation. Rotation is already encoded and must be zero at this adapter level. Reject images, SVG/blip fills, OLE and graphic frames. Source checksum is rechecked immediately before COM opens it read-only.

The inherited example has two independently composed panels, not one measured 3D camera. Charging cylinders share a -11.5 degree local orientation. Collection uses a consistent oblique box basis, with local tube orientation inherited from the prior output. Native contours retain their old projection. Do not claim a globally solved camera or full perspective consistency for retained legacy parts.

Ring, Rod, Tube, Plate and Connector can later become validated semantic wrappers on these primitives. Slotted Box needs parameterized openings, not a declaration that arbitrary rectangular cutouts have already been reconstructed. Retain native details until that work is verified. Do not silently downgrade unsupported geometry to a box.

## Style separation

`assets/style_presets/scientific_metallic.json` supplies steel, gray polymer and blue polymer tones. Geometry does not contain those colors. Changing material does not change surfaces, sizes or labels. Original 3D-printed polymer components stay polymer, not falsely metallic steel.

Scientific Flat will use the same geometry with solid fills and outlines; Anime Tech may add discrete shading. Both are reserved, and attempts to render them fail explicitly today. Imported native reference components are not automatically restyled. Their eventual conversion to parametric geometry is a prerequisite for complete style switching.

## Repair and validation limits

Patches contain the original scene SHA256, component_id, reason and a `set` mapping. They can change geometry, position, rotation, bounds, label anchor or draw order. They cannot change scientific labels, source evidence or relationship semantics. Patches do not mutate the input object or overwrite files; the CLI re-resolves native paths when output is in another folder.

The verifier checks component presence, compiled bounds, canvas containment, labels, leader arrow presence, occlusion order, COM-measured text fit, native shape counts and absence of raster/SVG/OLE/media. It compares component centres and proportions with independent reference regions when supplied. A warning for drift or a leader away from its component requires review, not an invented automatic correction.

`contains`/`connected_to`/`assembled_with` are currently semantic evidence, not geometric collision/contact constraints. There is no automatic segmentation, OCR, image similarity scoring, 3D inference, general layout solver or automated scientific judgement. Local repair is Agent-directed and rule-checked. Rendering plus visual inspection is mandatory for delivered artifacts.

## Diagram Graph and multi-image design (deferred)

Use `routegraph.diagram.v1` for nodes/ports/directed edges, routes, panels and source text. Share only IDs, source evidence, confidence, labels, units and style references with Instrument Scene. A diagram edge expresses flow; an instrument relation expresses assembly or space. Preserve the current continuous-arrow helpers as its future backend.

For multi-image integration, introduce a composition manifest referencing separate diagram/instrument scenes, each source image's role, per-panel transform, and explicit component identity matches. Keep conflicts and uncertain matches visible. Do not merge unrelated components because their shapes look similar. This design adds no new recognition or rendering code in Phase 1.

## Migration topology

These are annotation leaders, not airflow or electrical process connections:

| Arrow ID | From | To | Style | Meaning | Priority |
|---|---|---|---|---|---|
| leader/tungsten_wire | Tungsten wire label | Visible wire | Solid | Component label | High |
| leader/spacer | 3D-printed spacer label | Spacer body | Solid | Component label | High |
| leader/fitting_pipe | Fitting pipe label | Fitting pipe | Solid | Component label | High |
| leader/ground_collar | Stainless steel ground label | Ground collar | Solid | Component label | High |
| leader/guide_plate | Guide plate label | Guide plate | Solid | Component label | High |
| leader/cassette | Cassette label | Cassette | Solid | Component label | High |
| leader/collection_electrodes | Collection electrodes label | Upper electrode | Solid | Component label | High |
| leader/collection_electrodes/1 | Same label | Lower electrode | Solid | Component label | High |
| leader/collection_plate | Collection plate label | Removable plate | Solid | Component label | High |

The analytic primitive gallery has no arrows. The unchanged legacy regression contains A1 straight, A2 elbow dashed, A3 multi-bend dashed and A4 true branch; its existing source/target geometry remains the regression contract.
