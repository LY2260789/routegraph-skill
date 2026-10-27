"""Explainable PPTX and render-receipt checks; no image similarity claims."""
import argparse
import hashlib
import json
import math
from pathlib import Path
import xml.etree.ElementTree as ET
import zipfile

NS = {"p":"http://schemas.openxmlformats.org/presentationml/2006/main", "a":"http://schemas.openxmlformats.org/drawingml/2006/main"}
EMU = 12700


def native_bounds(shape):
    transform = shape.find("p:grpSpPr/a:xfrm",NS) if shape.tag.endswith("}grpSp") else shape.find("p:spPr/a:xfrm",NS)
    if transform is None:
        raise ValueError("Missing native geometry transform")
    off, ext = transform.find("a:off",NS), transform.find("a:ext",NS)
    x,y,w,h = [int(v)/EMU for v in (off.get("x"),off.get("y"),ext.get("cx"),ext.get("cy"))]
    angle = math.radians(int(transform.get("rot","0"))/60000)
    bw,bh = abs(w*math.cos(angle))+abs(h*math.sin(angle)),abs(w*math.sin(angle))+abs(h*math.cos(angle))
    geometry=shape.find("p:spPr/a:prstGeom",NS)
    if geometry is not None and geometry.get("prst")=="ellipse":
        bw,bh=math.hypot(w*math.cos(angle),h*math.sin(angle)),math.hypot(w*math.sin(angle),h*math.cos(angle))
    return [x+w/2-bw/2,y+h/2-bh/2,bw,bh]


def union(boxes):
    left=min(b[0] for b in boxes); top=min(b[1] for b in boxes)
    return [left,top,max(b[0]+b[2] for b in boxes)-left,max(b[1]+b[3] for b in boxes)-top]


def inspect_pptx(path):
    counts=dict(slides=0,groups=0,native_shapes=0,connectors=0,text_shapes=0,pictures=0,graphic_frames=0,media_files=0)
    texts=[]; names=[]; top={}; arrows={}
    with zipfile.ZipFile(path) as package:
        counts["media_files"]=sum(n.startswith("ppt/media/") for n in package.namelist())
        for name in package.namelist():
            if not (name.startswith("ppt/slides/slide") and name.endswith(".xml")):
                continue
            counts["slides"]+=1
            root=ET.fromstring(package.read(name))
            tree=root.find("p:cSld/p:spTree",NS)
            for shape in tree:
                if shape.tag.rsplit("}",1)[-1] not in {"sp","grpSp","cxnSp","pic"}:
                    continue
                props=shape.find(".//p:cNvPr",NS)
                if props is not None:
                    top[props.get("name")]=native_bounds(shape)
            for tag,key in [("grpSp","groups"),("sp","native_shapes"),("cxnSp","connectors"),("pic","pictures"),("graphicFrame","graphic_frames")]:
                counts[key]+=len(root.findall(f".//p:{tag}",NS))
            for shape in root.findall(".//p:sp",NS):
                body=shape.find("p:txBody",NS)
                if body is not None and body.findall(".//a:t",NS):
                    counts["text_shapes"]+=1
                    texts.append("".join(t.text or "" for t in body.findall(".//a:t",NS)))
            names.extend(p.get("name") for p in root.findall(".//p:cNvPr",NS))
            for shape in list(root.findall(".//p:sp",NS))+list(root.findall(".//p:cxnSp",NS)):
                props=shape.find(".//p:cNvPr",NS)
                if props is None:
                    continue
                head=shape.find("p:spPr/a:ln/a:tailEnd",NS)
                if head is not None and head.get("type") not in {None,"none"}:
                    arrows[props.get("name")]=head.get("type")
    return dict(counts=counts,texts=texts,names=names,top_level_bounds=top,arrow_shapes=arrows)


def verify(scene,plan,receipt,pptx):
    native=inspect_pptx(pptx)
    errors=[]; warnings=[]; comparisons=[]
    counts=native["counts"]
    if plan.get('source_scene_sha256')!=hashlib.sha256(json.dumps(scene,sort_keys=True).encode()).hexdigest():
        errors.append('Render plan belongs to a different scene revision')
    if counts["slides"]!=1 or counts["pictures"] or counts["graphic_frames"] or counts["media_files"]:
        errors.append("Expected one slide containing only native shapes, groups, connectors and text")
    if receipt["sha256"]!=hashlib.sha256(Path(pptx).read_bytes()).hexdigest():
        errors.append("Render receipt does not match PPTX")
    actual={g["id"]:g for g in receipt["groups"]}
    component_by_id={c["component_id"]:c for c in scene["components"]}
    if set(actual)!=set(component_by_id):
        errors.append("Receipt component set differs from scene")
    for group in plan["groups"]:
        cid=group["id"]
        if cid not in actual or 'component/'+cid not in native["top_level_bounds"]:
            errors.append(f"Missing editable component {cid}")
            continue
        measured=native["top_level_bounds"]['component/'+cid]
        error=max(abs(a-b) for a,b in zip(measured,group["expected_bounds"]))
        if error>1.5:
            errors.append(f"{cid}: geometry differs from plan by {error:.2f} pt")
        x,y,w,h=measured
        if x<-.5 or y<-.5 or x+w>scene["canvas"][0]+.5 or y+h>scene["canvas"][1]+.5:
            errors.append(f"{cid}: component extends outside canvas")
        c=component_by_id[cid]
        ref=c["source_reference"].get("bounding_box")
        if ref:
            # Independent baseline from old native output, not a pixel similarity score.
            position=math.hypot((x+w/2)-(ref[0]+ref[2]/2),(y+h/2)-(ref[1]+ref[3]/2))
            ratio=max(abs(w/ref[2]-1),abs(h/ref[3]-1))
            comparisons.append(dict(id=cid,center_shift_pt=round(position,3),relative_size_change=round(ratio,4)))
            if position>8 or ratio>.18:
                warnings.append(f"{cid}: baseline position/size drift needs visual review")
        if c["label"]:
            if c["label"] not in native["texts"]:
                errors.append(f"{cid}: label text missing or changed")
            if 'leader/'+cid not in native["arrow_shapes"]:
                errors.append(f"{cid}: label arrow missing")
            for i,_ in enumerate(c['label_anchor'].get('additional_targets',[]),1):
                if f'leader/{cid}/{i}' not in native['arrow_shapes']:
                    errors.append(f'{cid}: additional label arrow missing')
            # A reference anchor may touch a component boundary, allow 12 pt tolerance.
            tx,ty=c["label_anchor"]["target"]
            if not x-12<=tx<=x+w+12 or not y-12<=ty<=y+h+12:
                warnings.append(f"{cid}: label target is away from component bounds")
        for rel in c["relationships"]:
            if rel["type"]=="occludes" and actual[cid]["z_order"]<=actual[rel["target"]]["z_order"]:
                errors.append(f"{cid}: wrong occlusion order")
    for label in receipt["text_checks"]:
        if label["overflow"]:
            errors.append(f"{label['id']}: PowerPoint measured text overflow")
    for a in scene.get("annotations",[]):
        if a["text"] not in native["texts"]:
            errors.append(f"Missing annotation {a['id']}")
    reference_kinds={r.get('kind') for r in scene.get('references',[])}
    has_photo='user_instrument_photo' in reference_kinds
    has_native=any(c['geometry_type']=='NativeReference' for c in scene['components'])
    return dict(status="fail" if errors else "pass",counts=counts,errors=errors,warnings=warnings,reference_comparison=comparisons,
        dimensions={"scientific_structure":("photo exterior interpreted by Agent; no internal structure or physical dimensions verified" if has_photo else "original scientific reference unavailable or synthetic; scene relationships and content checked only"),
        "editability":"native OOXML inspected; no raster/SVG/OLE allowed",
        "visual_style":("requires rendered preview review; native references retain their source style" if has_native else "geometry uses a shared style preset; requires rendered preview review"),
        "reference_similarity":("component bounds against manually identified photo regions; no pixel score" if has_photo else "component bounds against prior output when provided; no pixel score")})


def main():
    p=argparse.ArgumentParser()
    p.add_argument("pptx",type=Path); p.add_argument("--scene",type=Path); p.add_argument("--plan",type=Path); p.add_argument("--receipt",type=Path); p.add_argument("--out",type=Path,required=True)
    args=p.parse_args()
    if args.out.exists():
        raise FileExistsError(args.out)
    def read(path): return json.loads(path.read_text(encoding="utf-8-sig"))
    if args.scene:
        result=verify(read(args.scene),read(args.plan),read(args.receipt),args.pptx)
        if read(args.receipt)['plan_sha256']!=hashlib.sha256(args.plan.read_bytes()).hexdigest():
            result['errors'].append('Receipt does not match render plan')
            result['status']='fail'
    else:
        result=inspect_pptx(args.pptx)
    args.out.parent.mkdir(parents=True,exist_ok=True)
    args.out.write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding="utf-8")
    print(json.dumps(result.get("counts",{})))
    if result.get("status")=="fail":
        raise SystemExit(1)


if __name__=="__main__":
    main()
