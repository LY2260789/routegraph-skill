from pathlib import Path
import json,sys
import xml.etree.ElementTree as ET
import zipfile
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/reconstruction'))
from verify import inspect_pptx,NS
path=Path(sys.argv[1]);out=Path(sys.argv[2]);data=inspect_pptx(path)
expected=['A1_Straight_Solid','A2_Elbow_Dashed','A3_Polyline_Dashed','A4_Branch_Target_1','A4_Branch_Target_2']
assert set(expected)==set(data['arrow_shapes']),data['arrow_shapes']
assert all(data['names'].count(n)==1 for n in expected)
assert data['counts']['pictures']==0
with zipfile.ZipFile(path) as z:
    root=ET.fromstring(z.read('ppt/slides/slide1.xml'))
    shapes={s.find('.//p:cNvPr',NS).get('name'):s for s in root.find('p:cSld/p:spTree',NS) if s.tag.rsplit('}',1)[-1] in {'sp','cxnSp'}}
    for name in ['A2_Elbow_Dashed','A3_Polyline_Dashed']:
        assert shapes[name].find('p:spPr/a:ln/a:prstDash',NS).get('val')!='solid'
    custom=shapes['A3_Polyline_Dashed'].find('p:spPr/a:custGeom/a:pathLst/a:path',NS)
    endpoints=[]
    for segment in custom:
        kind=segment.tag.rsplit('}',1)[-1]
        points=[(int(p.get('x')),int(p.get('y'))) for p in segment.findall('a:pt',NS)]
        if kind=='cubicBezTo':
            # The existing COM helper serializes a straight segment as a degenerate
            # cubic with all three control/end points identical. Preserve this behavior.
            assert len(set(points))==1
        if kind in {'lnTo','cubicBezTo'}: endpoints.append(points[-1])
    assert endpoints==[(1270000,0),(1270000,1143000),(2413000,1143000)],endpoints
if out.exists():raise FileExistsError(out)
out.write_text(json.dumps(dict(status='pass',counts=data['counts'],single_shape_arrows=expected,assertions=['single shape for nonbranching arrows','continuous dashed polyline has three segments','true branch retains trunk and bus']),indent=2),encoding='utf-8')
print('Legacy arrow topology: PASS')
