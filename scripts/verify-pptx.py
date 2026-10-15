"""Independently inspect the generated PPTX using only Python's standard library."""
import io
import json
import re
import struct
import sys
import xml.etree.ElementTree as ET
from zipfile import ZipFile

ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main', 'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'}
with ZipFile(sys.argv[1]) as z:
    assert z.testzip() is None, 'PPTX archive CRC failure'
    names = z.namelist()
    slides = [n for n in names if re.fullmatch(r'ppt/slides/slide\d+\.xml', n)]
    assert len(slides) == 1
    root = ET.fromstring(z.read(slides[0]))
    assert 'Editable PPT environment test' in ''.join(root.itertext())
    assert len(root.findall('.//p:sp', ns)) >= 3, 'Missing independently editable shapes'
    connectors = root.findall('.//p:cxnSp', ns)
    assert len(connectors) == 1
    assert connectors[0].find('.//a:tailEnd', ns).get('type') == 'triangle'
    tables = root.findall('.//a:tbl', ns)
    assert len(tables) == 1 and len(tables[0].findall('a:tr', ns)) == 3
    assert all(len(row.findall('a:tc', ns)) == 2 for row in tables[0].findall('a:tr', ns))
    assert not root.findall('.//a:normAutofit', ns) and not root.findall('.//a:spAutoFit', ns)
    assert all(el.get('noGrp') != '1' for el in root.iter())
    for tag in ['defRPr', 'endParaRPr']:
        for el in root.findall('.//a:' + tag, ns):
            assert el.get('b') is None and el.get('sz') is None
            assert all(el.find('a:' + font, ns) is None for font in ['latin', 'ea', 'cs'])
    assert any(el.get('sz') for el in root.findall('.//a:rPr', ns)), 'Missing editable run font sizes'
    assert any(el.find('a:latin', ns) is not None for el in root.findall('.//a:rPr', ns)), 'Missing editable run font faces'
    for name in names:
        if name.endswith('.rels'):
            for rel in ET.fromstring(z.read(name)):
                assert not (rel.get('Type', '').endswith('/image') and rel.get('TargetMode') == 'External'), 'External image relationship'
    svgs = [n for n in names if n.startswith('ppt/media/') and n.endswith('.svg')]
    assert svgs, 'SVG embedding was not exercised'
    for svg in svgs:
        png = z.read(svg[:-4] + '.png')
        assert len(png) > 1000 and png[:8] == b'\x89PNG\r\n\x1a\n'
        assert struct.unpack('>II', png[16:24]) == (1024, 1024)
        assert png[25] in (4, 6), 'Fallback must have transparency'
    assert len(root.findall('.//p:pic', ns)) == 1, 'Only the small test icon may be an image'
    print(json.dumps({'slides': 1, 'editable_text': True, 'native_connectors': 1, 'native_tables': 1, 'svg_png_fallbacks': len(svgs), 'no_external_images': True, 'typography_and_grouping': 'passed', 'zip_integrity': 'passed'}, indent=2))
