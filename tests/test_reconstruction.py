import copy
import hashlib
import json
from pathlib import Path
import sys
import tempfile
import unittest

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'scripts/reconstruction'))
from scene import validate,compile_scene,primitives,rotated_bounds,apply_patch,read_json,write_new
from verify import inspect_pptx

FIXTURE=ROOT/'assets/examples/instrument_reconstruction/charging_collection.scene.json'
STYLE=ROOT/'assets/style_presets/scientific_metallic.json'

class SceneTests(unittest.TestCase):
    def setUp(self):
        self.scene=read_json(FIXTURE); self.style=read_json(STYLE)
        self.c=next(c for c in self.scene['components'] if c['component_id']=='spacer')

    def test_fixture_compiles_with_native_detail(self):
        p=compile_scene(self.scene,FIXTURE,self.style)
        self.assertEqual(19,len(p['groups']))
        self.assertEqual(10,sum('shapes' in g for g in p['groups']))

    def test_duplicate_ids_rejected(self):
        self.scene['components'].append(copy.deepcopy(self.c))
        with self.assertRaises(ValueError):validate(self.scene)

    def test_nan_rejected(self):
        self.c['geometry_parameters']['radius']=float('nan')
        with self.assertRaises(ValueError):validate(self.scene)

    def test_missing_label_anchor_rejected(self):
        self.c['label_anchor']=None
        with self.assertRaises(ValueError):validate(self.scene)

    def test_unknown_geometry_rejected(self):
        self.c['geometry_type']='SlottedBox'
        with self.assertRaises(ValueError):validate(self.scene)

    def test_unverified_internal_structure_rejected(self):
        self.c['source_reference'].update(evidence='unverified',note='Not visible in photo')
        with self.assertRaises(ValueError):validate(self.scene)

    def test_hole_cannot_exceed_outer_radius(self):
        self.c['geometry_parameters']['inner_radius']=self.c['geometry_parameters']['radius']
        with self.assertRaises(ValueError):validate(self.scene)

    def test_invalid_occlusion_rejected(self):
        self.c['relationships']=[dict(type='occludes',target='collection_plate')]
        with self.assertRaises(ValueError):validate(self.scene)

    def test_missing_relationship_target_rejected(self):
        self.c['relationships']=[dict(type='connected_to',target='missing')]
        with self.assertRaises(ValueError):validate(self.scene)

    def test_unknown_reference_rejected(self):
        self.c['source_reference']['reference_id']='invented'
        with self.assertRaises(ValueError):validate(self.scene)

    def test_solid_cylinder_analytic_extents(self):
        c=copy.deepcopy(self.c); c['geometry_type']='Cylinder'; c['geometry_parameters']=dict(length=100,radius=20)
        _,b=primitives(c,self.scene['projection'],self.style['materials']['steel'])
        self.assertEqual([-7.6,-20,115.2,40],b)
        out=rotated_bounds(b,90,[200,100])
        for a,e in zip(out,[180,92.4,40,115.2]):self.assertAlmostEqual(a,e)

    def test_material_change_keeps_geometry(self):
        a,b=primitives(self.c,self.scene['projection'],self.style['materials']['steel'])
        x,y=primitives(self.c,self.scene['projection'],self.style['materials']['blue_polymer'])
        self.assertEqual(b,y)
        for s,t in zip(a,x):
            self.assertEqual(s.get('box'),t.get('box')); self.assertEqual(s.get('points'),t.get('points'))

    def test_local_patch_preserves_unrelated_components(self):
        before=copy.deepcopy(self.scene)
        patch=read_json(FIXTURE.parent/'label_fit.patch.json')
        after=apply_patch(self.scene,patch)
        self.assertEqual(before,self.scene)
        self.assertEqual(before['components'][0],after['components'][0])
        for a,b in zip(before['components'],after['components']):
            self.assertEqual(a['geometry_parameters'],b['geometry_parameters'])

    def test_stale_patch_rejected(self):
        patch=read_json(FIXTURE.parent/'label_fit.patch.json'); patch['source_scene_sha256']='old'
        with self.assertRaises(ValueError):apply_patch(self.scene,patch)

    def test_scientific_relationship_patch_rejected(self):
        patch=read_json(FIXTURE.parent/'label_fit.patch.json'); patch['changes'][0]['set']={'relationships':[]}
        with self.assertRaises(ValueError):apply_patch(self.scene,patch)

    def test_missing_native_file_fails_before_render(self):
        self.scene['components'][0]['geometry_parameters']['path']='missing.pptx'
        with self.assertRaises(ValueError):compile_scene(self.scene,FIXTURE,self.style)

    def test_native_checksum_mismatch_rejected(self):
        self.scene['components'][0]['geometry_parameters']['sha256']='bad'
        with self.assertRaises(ValueError):compile_scene(self.scene,FIXTURE,self.style)

    def test_missing_native_shape_rejected(self):
        self.scene['components'][0]['geometry_parameters']['shape_names']=['missing']
        with self.assertRaises(ValueError):compile_scene(self.scene,FIXTURE,self.style)

    def test_no_overwrite(self):
        with tempfile.TemporaryDirectory() as tmp:
            p=Path(tmp)/'input.json';write_new(p,{'original':True})
            with self.assertRaises(FileExistsError):write_new(p,{})
            self.assertTrue(read_json(p)['original'])

    def test_native_reference_editability(self):
        data=inspect_pptx(FIXTURE.parent/'legacy_reference.pptx')
        self.assertEqual(0,data['counts']['pictures'])
        self.assertEqual(0,data['counts']['media_files'])
        self.assertEqual(12,data['counts']['text_shapes'])

    def test_reserved_style_is_not_silently_implemented(self):
        self.style['id']='anime_tech'
        with self.assertRaises(ValueError):compile_scene(self.scene,FIXTURE,self.style)

if __name__=='__main__':unittest.main()
