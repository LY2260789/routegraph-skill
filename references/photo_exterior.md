# Photo to editable instrument exterior

Use an accessible photo as primary appearance evidence. If several figures show different configurations, choose the user's specified photograph for geometry. Other diagrams can explain context but do not authorize mixing panel layouts, adding hidden parts or inferring circuit/flow connections.

1. Identify the target instrument and excluded surrounding objects. Preserve input files unchanged with hashes. Record visible components and source pixel regions before drawing.
2. Establish one simple projection and a drawing scale. Screen points are not physical measurements. Treat cabinet depth and concealed surfaces as simplified, not measured.
3. Map housings/rails to Cuboid; visible controls, grille patterns and faceplates to Panel; suitable cylindrical parts to Cylinder/HollowCylinder. Panel features have semantic names and style tone references. An uncertain grille pattern may be simplified, but its hole count must not be presented as an instrument specification.
4. Keep visible labels editable. Do not fill unreadable screens with invented numeric readings. Do not draw hidden optical paths, ion sources or electronics just because a related instrument diagram contains them.
5. Export white-background native PPTX and PNG, inspect silhouette/proportions/occlusion/contact points, and apply local patches to demonstrated errors. Validate native object counts and editability. Check copied prototype data against the actual primary photo.

## Recent Photonion case

The user selected the third laboratory photograph, not the other product photos embedded in diagrams. The exterior includes a high tower, a lower cabinet, monitor, red upper faceplate, gray door, dark lower equipment panels, aluminum frame and two visible casters. Lab background, yellow barrier, chair, cylinders, separate analyzer and external hoses were excluded.

The scene uses only fresh parametric shapes, with no NativeReference or source image embedded. The photo-specific material file uses the existing `scientific_metallic` preset contract with additional coating/rack/screen colors. Exact dark-panel details are simplified; no instrument values or unknown internals are asserted.

The shared renderer and compiler own geometry. Photo-specific scene coordinates and palettes belong to the user's output directory, not the reusable implementation. The checked example is under the originating workspace's `output/phase2/` and is not a required installed runtime asset.

## Commands

```powershell
python scripts/reconstruction/scene.py PATH_TO_SCENE --style PATH_TO_PALETTE --out NEW_PLAN.json
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/render_instrument.ps1 -PlanPath NEW_PLAN.json -OutputPath NEW.pptx
python scripts/reconstruction/verify.py NEW.pptx --scene PATH_TO_SCENE --plan NEW_PLAN.json --receipt NEW.render.json --out NEW.validation.json
```

The optional native edit smoke test accepts `-ComponentId`, `-SurfaceName`, `-LabelName` and `-ExpectedPartCount`, so a new photo case can be tested without hardcoding its component identity in test code. Always write the edit test to a separate file.
