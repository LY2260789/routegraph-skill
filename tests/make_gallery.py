"""Generate independent analytic primitives, not a claimed reconstruction of an instrument."""
from pathlib import Path
import sys
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'scripts/reconstruction'))
from scene import primitives,rotated_bounds,compile_scene,read_json,write_new

def make_scene():
    style=read_json(ROOT/'assets/style_presets/scientific_metallic.json')
    projection=dict(type='oblique',ellipse_ratio=.38,x_basis=[1,-.25],depth_basis=[-.8,-.5],vertical_basis=[0,1])
    scene=dict(schema_version='routegraph.instrument.v1',canvas=[840,360],font='Arial',projection=projection,
      references=[dict(reference_id='analytic',kind='synthetic_geometry_test')],assemblies=[dict(id='test')],components=[],annotations=[])
    specs=[('solid','Cylinder',dict(length=135,radius=35),[90,190],-12),('hollow','HollowCylinder',dict(length=135,radius=35,inner_radius=26,bore_depth=24),[355,190],-12),('box','Cuboid',dict(width=120,height=65,depth=65),[635,170],0)]
    for index,(cid,kind,g,origin,rotation) in enumerate(specs):
        c=dict(component_id=cid,component_type='analytic_test',geometry_type=kind,geometry_parameters=g,relative_position=dict(frame='test',origin=origin),rotation=rotation,z_order=index,material='steel',label='',label_anchor=None,relationships=[],confidence=1,
          source_reference=dict(reference_id='analytic',evidence='simplified',note='Analytic primitive demonstration only; no physical instrument or measured dimensions represented'))
        _,b=primitives(c,projection,style['materials']['steel']);c['bounding_box']=rotated_bounds(b,rotation,origin)
        scene['components'].append(c)
        scene['annotations'].append(dict(id=cid,text=kind,box=[60+index*270,270,250,35],font_size=22,bold=False))
    scene['annotations'].insert(0,dict(id='title',text='Scientific Metallic: editable primitives',box=[45,30,750,50],font_size=28,bold=True))
    return scene,style

if __name__=='__main__':
    out=Path(sys.argv[1]);scene,style=make_scene()
    write_new(out/'primitives.scene.json',scene)
    write_new(out/'primitives.plan.json',compile_scene(scene,out/'primitives.scene.json',style))
