"""COM-independent instrument scene compiler. Units are PowerPoint points, not mm."""
from __future__ import annotations
import argparse
import copy
import hashlib
import json
import math
from pathlib import Path
import xml.etree.ElementTree as ET
import zipfile

SUPPORTED = {"Cylinder", "HollowCylinder", "Cuboid", "Panel", "NativeReference"}
RELATIONS = {"connected_to", "contains", "adjacent_to", "occludes", "assembled_with"}
EVIDENCE = {"visible", "simplified", "unverified", "legacy_output"}


def read_json(path):
    return json.loads(Path(path).read_text(encoding="utf-8-sig"))


def write_new(path, value):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("x", encoding="utf-8") as stream:
        json.dump(value, stream, ensure_ascii=False, indent=2, allow_nan=False)


def number(value, name, positive=False):
    if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value):
        raise ValueError(f"{name}: expected finite number")
    if positive and value <= 0:
        raise ValueError(f"{name}: must be positive")


def vector(value, n, name):
    if not isinstance(value, list) or len(value) != n:
        raise ValueError(f"{name}: expected {n} values")
    for item in value:
        number(item, name)


def bbox(value, name):
    vector(value, 4, name)
    if min(value[2:]) <= 0:
        raise ValueError(f"{name}: nonpositive extent")


def validate(scene):
    if scene.get("schema_version") != "routegraph.instrument.v1":
        raise ValueError("Unsupported scene schema (Diagram Graph uses a different schema)")
    vector(scene["canvas"], 2, "canvas")
    if min(scene["canvas"]) <= 0:
        raise ValueError("canvas must be positive")
    projection = scene["projection"]
    if projection["type"] != "oblique":
        raise ValueError("Only oblique projection is supported in Phase 1")
    if not 0 < projection["ellipse_ratio"] <= 1:
        raise ValueError("ellipse_ratio must be in (0,1]")
    for key in ("x_basis", "depth_basis", "vertical_basis"):
        vector(projection[key], 2, key)
        if math.hypot(*projection[key]) == 0:
            raise ValueError(f"{key}: degenerate projection")
    if not (projection['x_basis'][0]>0 and projection['x_basis'][1]<=0 and projection['depth_basis'][0]!=0 and projection['depth_basis'][1]<0 and projection['vertical_basis']==[0,1]):
        raise ValueError('Oblique faces require right/up x, upward nonzero depth and downward vertical')
    reference_ids={r['reference_id'] for r in scene['references']}
    frames={a['id'] for a in scene['assemblies']}
    ids = [c["component_id"] for c in scene["components"]]
    if len(ids) != len(set(ids)) or not ids:
        raise ValueError("Components must have unique IDs and cannot be empty")
    order = {}
    warnings = []
    for c in scene["components"]:
        cid = c["component_id"]
        if not cid or any(x in cid for x in "/\\"):
            raise ValueError("component_id must be nonempty and have no slash")
        for key in ("component_type", "geometry_parameters", "relative_position", "rotation", "z_order", "material", "label", "label_anchor", "relationships", "confidence", "source_reference"):
            if key not in c:
                raise ValueError(f"{cid}: missing {key}")
        if c["geometry_type"] not in SUPPORTED:
            raise ValueError(f"{cid}: unsupported geometry {c['geometry_type']}")
        bbox(c["bounding_box"], cid)
        vector(c["relative_position"]["origin"], 2, cid)
        number(c["rotation"], "rotation")
        number(c["z_order"], "z_order")
        number(c["confidence"], "confidence")
        if not 0 <= c["confidence"] <= 1:
            raise ValueError(f"{cid}: confidence out of range")
        src = c["source_reference"]
        if src.get("reference_id") not in reference_ids or src["evidence"] not in EVIDENCE:
            raise ValueError(f"{cid}: missing source evidence")
        if c['relative_position']['frame'] not in frames:
            raise ValueError(f'{cid}: unknown assembly frame')
        if src["evidence"] in {"unverified", "simplified"} and not src.get("note"):
            raise ValueError(f"{cid}: inference needs a note")
        if src["evidence"] == "unverified":
            raise ValueError(f"{cid}: unverified geometry cannot be rendered as established structure")
        if src["evidence"] == "legacy_output":
            warnings.append(f"{cid}: derived from prior output, not verified against original scientific figure")
        if c["label"]:
            if not c["label_anchor"]:
                raise ValueError(f"{cid}: label needs anchor")
            bbox(c["label_anchor"]["box"], "label box")
            vector(c["label_anchor"]["target"], 2, "label target")
            vector(c["label_anchor"]["start"], 2, "leader start")
            number(c['label_anchor']['font_size'], 'font_size', positive=True)
            for target in c['label_anchor'].get('additional_targets',[]):
                vector(target,2,'additional leader target')
        g = c["geometry_parameters"]
        if c["geometry_type"] in {"Cylinder", "HollowCylinder"}:
            for key in ("length", "radius"):
                number(g[key], key, positive=True)
            if c["geometry_type"] == "HollowCylinder":
                number(g["inner_radius"], "inner_radius", positive=True)
                number(g["bore_depth"], "bore_depth", positive=True)
                if g["inner_radius"] >= g["radius"] or g["bore_depth"] > g["length"]:
                    raise ValueError(f"{cid}: invalid bore dimensions")
        elif c["geometry_type"] == "Cuboid":
            for key in ("width", "height", "depth"):
                number(g[key], key, positive=True)
        elif c['geometry_type']=='Panel':
            for key in ('width','height'):
                number(g[key],key,positive=True)
            names=set()
            for f in g.get('features',[]):
                if f['kind'] not in {'rectangle','ellipse'} or not f.get('name') or f['name'] in names:
                    raise ValueError(f'{cid}: panel features need unique names and supported kinds')
                names.add(f['name'])
                bbox(f['box'],'panel feature')
                x,y,w,h=f['box']
                if x<0 or y<0 or x+w>g['width'] or y+h>g['height']:
                    raise ValueError(f'{cid}: feature outside panel')
        else:
            if not g.get("shape_names") or len(g["shape_names"]) != len(set(g["shape_names"])):
                raise ValueError(f"{cid}: native reference needs unique source shape names")
            if c["rotation"] != 0:
                raise ValueError("NativeReference rotation is baked in; transform support is deferred")
            bbox(g["source_bounds"], "native source bounds")
        order[cid] = c["z_order"]
    for c in scene["components"]:
        for relation in c["relationships"]:
            if relation["type"] not in RELATIONS or relation["target"] not in order:
                raise ValueError(f"{c['component_id']}: invalid relationship")
            if relation["target"] == c["component_id"]:
                raise ValueError("Self relationship is not allowed")
            if relation["type"] == "occludes" and c["z_order"] <= order[relation["target"]]:
                raise ValueError(f"{c['component_id']}: occlusion contradicts z_order")
    return warnings


def rotated_bounds(bounds, angle, origin):
    x, y, w, h = bounds
    a = math.radians(angle)
    points = [(origin[0] + px * math.cos(a) - py * math.sin(a), origin[1] + px * math.sin(a) + py * math.cos(a)) for px, py in [(x,y),(x+w,y),(x+w,y+h),(x,y+h)]]
    xs, ys = zip(*points)
    return [min(xs), min(ys), max(xs)-min(xs), max(ys)-min(ys)]


def primitives(component, projection, material):
    """Local named surfaces, grouped and rotated by backend; no PPT constants here."""
    g = component["geometry_parameters"]
    shapes = []
    def ellipse(name, x, y, w, h, tone, gradient=None):
        shapes.append(dict(kind="ellipse", name=name, box=[x,y,w,h], fill=material[tone], gradient=gradient))
    def poly(name, pts, tone, gradient=None):
        shapes.append(dict(kind="polygon", name=name, points=pts, fill=material[tone], gradient=gradient))
    if component["geometry_type"] in {"Cylinder", "HollowCylinder"}:
        length, radius = g["length"], g["radius"]
        rx = radius * projection["ellipse_ratio"]
        ellipse("rear_cap", length-rx, -radius, 2*rx, 2*radius, "dark")
        poly("outer_wall", [[0,-radius],[length,-radius],[length,radius],[0,radius]], "base", [material["light"], material["dark"]])
        ellipse("front_face", -rx, -radius, 2*rx, 2*radius, "rim", [material["light"], material["rim"]])
        if component["geometry_type"] == "HollowCylinder":
            ri = g["inner_radius"]
            irx = ri * projection["ellipse_ratio"]
            # Concentric opening plus inset dark ellipse show visible inner-wall depth.
            # This is a stylized opening, not a claim of a closed internal end cap.
            ellipse("inner_wall", -irx, -ri, 2*irx, 2*ri, "inside", [material["inside"], material["dark"]])
            depth_fraction = min(g["bore_depth"] / (2*radius), .65)
            inset = .12 + .24 * depth_fraction
            ellipse("bore_shadow", -irx*(1-inset), -ri*(1-inset), 2*irx*(1-inset), 2*ri*(1-inset), "bore")
        bounds = [-rx, -radius, length+2*rx, 2*radius]
    elif component['geometry_type']=='Panel':
        w,h=g['width'],g['height']
        shapes.append(dict(kind='rectangle',name='surface',box=[0,0,w,h],fill=material['base'],gradient=[material['light'],material['base']]))
        for f in g.get('features',[]):
            if f['tone'] not in material:
                raise ValueError(f"Unknown panel tone: {f['tone']}")
            shapes.append(dict(kind=f['kind'],name=f['name'],box=f['box'],fill=material[f['tone']],gradient=None))
        bounds=[0,0,w,h]
    else:
        w, h, d = g["width"], g["height"], g["depth"]
        u, v, z = projection["x_basis"], projection["depth_basis"], projection["vertical_basis"]
        def point(a,b,c):
            return [a*u[0]+b*v[0]+c*z[0], a*u[1]+b*v[1]+c*z[1]]
        side_x = 0 if v[0] < 0 else w
        poly("side_wall", [point(side_x,d,0),point(side_x,0,0),point(side_x,0,h),point(side_x,d,h)], "dark")
        poly("front_wall", [point(0,0,0),point(w,0,0),point(w,0,h),point(0,0,h)], "base", [material["base"],material["dark"]])
        poly("top_face", [point(0,0,0),point(0,d,0),point(w,d,0),point(w,0,0)], "light")
        points = [p for s in shapes for p in s["points"]]
        xs, ys = zip(*points)
        bounds = [min(xs), min(ys), max(xs)-min(xs), max(ys)-min(ys)]
    return shapes, bounds


def compile_scene(scene, scene_path, style):
    warnings = validate(scene)
    if style.get("status") != "implemented" or style["id"] != "scientific_metallic":
        raise ValueError("Only Scientific Metallic is implemented in Phase 1")
    groups = []
    selected_native=set()
    for c in sorted(scene["components"], key=lambda c: (c["z_order"], c["component_id"])):
        group = dict(id=c["component_id"], geometry_type=c["geometry_type"], origin=c["relative_position"]["origin"], rotation=c["rotation"], z_order=c["z_order"], source=c["source_reference"])
        if c["geometry_type"] == "NativeReference":
            group["native"] = copy.deepcopy(c["geometry_parameters"])
            source_path = (Path(scene_path).parent / group["native"]["path"]).resolve()
            if not source_path.is_file():
                raise ValueError(f"Missing native source: {source_path}")
            digest = hashlib.sha256(source_path.read_bytes()).hexdigest()
            if digest != group["native"]["sha256"]:
                raise ValueError("Native source checksum mismatch")
            if source_path.suffix.lower()!='.pptx':
                raise ValueError('Native references must be PPTX')
            ns={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
            with zipfile.ZipFile(source_path) as package:
                root=ET.fromstring(package.read(f"ppt/slides/slide{group['native']['slide']}.xml"))
                available={}
                for shape in root.find('p:cSld/p:spTree',ns):
                    if shape.tag.rsplit('}',1)[-1] not in {'sp','grpSp','cxnSp','pic','graphicFrame'}:
                        continue
                    props=shape.find('.//p:cNvPr',ns)
                    if props is not None: available[props.get('name')]=shape
                for name in group['native']['shape_names']:
                    key=(str(source_path),group['native']['slide'],name)
                    if key in selected_native or name not in available:
                        raise ValueError(f'Duplicate or missing native shape: {name}')
                    selected_native.add(key)
                    shape=available[name]
                    if any(e.tag.rsplit('}',1)[-1] in {'pic','graphicFrame','blip','oleObj'} for e in shape.iter()):
                        raise ValueError(f'Noneditable visual in native reference: {name}')
            group["native"]["path"] = str(source_path)
            group["expected_bounds"] = list(group["origin"])+group["native"]["source_bounds"][2:]
        else:
            if c["material"] not in style["materials"]:
                raise ValueError(f"Unknown material {c['material']}")
            shapes, bounds = primitives(c, scene["projection"], style["materials"][c["material"]])
            group.update(shapes=shapes, local_bounds=bounds, expected_bounds=rotated_bounds(bounds,c["rotation"],group["origin"]))
        group["declared_bounds"] = c["bounding_box"]
        if max(abs(a-b) for a,b in zip(group['declared_bounds'],group['expected_bounds']))>1.5:
            raise ValueError(f"{group['id']}: declared bounding_box disagrees with compiled geometry; update both in a local repair")
        groups.append(group)
    labels = []
    for c in scene["components"]:
        if c["label"]:
            labels.append(dict(id=c["component_id"], text=c["label"], **c["label_anchor"]))
    return dict(schema_version="routegraph.render-plan.v1", canvas=scene["canvas"], style_id=style["id"], font=scene.get("font","Arial"), groups=groups, labels=labels, annotations=scene.get("annotations",[]), warnings=warnings, source_scene_sha256=hashlib.sha256(json.dumps(scene,sort_keys=True).encode()).hexdigest())


def apply_patch(scene, patch):
    """Explicit local repairs only; never infer scientific connections automatically."""
    if patch.get("source_scene_sha256") != hashlib.sha256(json.dumps(scene, sort_keys=True).encode()).hexdigest():
        raise ValueError("Patch is stale or belongs to another scene")
    result = copy.deepcopy(scene)
    by_id = {c["component_id"]: c for c in result["components"]}
    allowed = {"relative_position", "rotation", "geometry_parameters", "bounding_box", "label_anchor", "z_order"}
    for change in patch["changes"]:
        if not change.get("reason"):
            raise ValueError("Repair needs an evidence-based reason")
        cid = change["component_id"]
        if cid not in by_id or not set(change["set"]).issubset(allowed):
            raise ValueError("Unknown component or forbidden patch field")
        by_id[cid].update(copy.deepcopy(change["set"]))
    validate(result)
    result.setdefault("repair_history", []).append(copy.deepcopy(patch))
    return result


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("scene", type=Path)
    parser.add_argument("--style", type=Path, default=Path(__file__).resolve().parents[2]/"assets/style_presets/scientific_metallic.json")
    parser.add_argument("--out", type=Path, required=True)
    parser.add_argument("--patch", type=Path)
    args = parser.parse_args()
    scene = read_json(args.scene)
    if args.patch:
        # Keep relative native source resolution valid when writing to another directory.
        repaired = apply_patch(scene, read_json(args.patch))
        for c in repaired["components"]:
            if c["geometry_type"] == "NativeReference":
                c["geometry_parameters"]["path"] = str((args.scene.parent/c["geometry_parameters"]["path"]).resolve())
        write_new(args.out, repaired)
    else:
        write_new(args.out, compile_scene(scene,args.scene,read_json(args.style)))


if __name__ == "__main__":
    main()
