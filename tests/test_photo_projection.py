import copy
from pathlib import Path
import sys
import unittest
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'scripts/reconstruction'))
from scene import primitives,read_json,validate

class PhotoProjectionTests(unittest.TestCase):
    def setUp(self):
        self.scene=read_json(ROOT/'assets/examples/instrument_reconstruction/charging_collection.scene.json')
        self.material=read_json(ROOT/'assets/style_presets/scientific_metallic.json')['materials']['steel']
        self.component={'geometry_type':'Cuboid','geometry_parameters':{'width':100,'height':200,'depth':20}}

    def test_right_visible_side_starts_at_front_right_edge(self):
        p=copy.deepcopy(self.scene['projection']);p.update(x_basis=[1,0],depth_basis=[1,-.6])
        parts,bounds=primitives(self.component,p,self.material)
        self.assertEqual(parts[0]['points'],[[120,-12],[100,0],[100,200],[120,188]])
        self.assertEqual(bounds,[0,-12,120,212])

    def test_left_visible_side_stays_at_front_left_edge(self):
        p=copy.deepcopy(self.scene['projection']);p.update(x_basis=[1,0],depth_basis=[-1,-.6])
        parts,bounds=primitives(self.component,p,self.material)
        self.assertEqual(parts[0]['points'],[[-20,-12],[0,0],[0,200],[-20,188]])
        self.assertEqual(bounds,[-20,-12,120,212])

    def test_camera_below_object_is_rejected(self):
        self.scene['projection']['depth_basis']=[1,.6]
        with self.assertRaises(ValueError):validate(self.scene)

    def test_panel_features_keep_exact_relative_positions(self):
        c={'geometry_type':'Panel','geometry_parameters':{'width':100,'height':40,'features':[{'name':'display','kind':'rectangle','box':[10,5,40,20],'tone':'inside'}]}}
        parts,bounds=primitives(c,self.scene['projection'],self.material)
        self.assertEqual(bounds,[0,0,100,40])
        self.assertEqual(parts[1]['box'],[10,5,40,20])
        self.assertEqual(parts[1]['name'],'display')

    def test_panel_feature_outside_surface_is_rejected(self):
        c=self.scene['components'][0]
        c['geometry_type']='Panel';c['geometry_parameters']={'width':100,'height':40,'features':[{'name':'overflow','kind':'ellipse','box':[95,5,10,10],'tone':'inside'}]}
        with self.assertRaises(ValueError):validate(self.scene)
