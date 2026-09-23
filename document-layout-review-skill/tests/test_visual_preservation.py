"""Targeted regressions for preserving source appearance during layout repair.

Run: python -m unittest discover -s tests -p test_visual_preservation.py -v
Fixtures use small OOXML packages; no user documents or external downloads.
"""
from __future__ import annotations

import base64
from copy import deepcopy
import importlib.util
from pathlib import Path
import sys
import tempfile
import unittest
import zipfile
from xml.etree import ElementTree as ET

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))
from layout_invariant_guard import snapshot_docx, changed_invariants, W, A, R, q
from table_layout_repair import repair_document_xml, write_docx

WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
PNG = base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9ZQmcAAAAASUVORK5CYII=")


def fixture() -> dict[str, bytes]:
    document = f'''<w:document xmlns:w="{W}" xmlns:a="{A}" xmlns:r="{R}" xmlns:wp="{WP}" xmlns:pic="{PIC}"><w:body>
<w:p><w:r><w:t>原说明</w:t></w:r></w:p>
<w:p><w:r><w:drawing><wp:inline><wp:extent cx="2000" cy="1000"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rId1"/><a:srcRect/></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2000" cy="1000"/></a:xfrm><a:solidFill><a:srgbClr val="70AD47"/></a:solidFill></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
<w:p><w:r><w:t>图1 技术路线</w:t></w:r></w:p>
<w:tbl><w:tblPr><w:tblStyle w:val="Plain"/><w:tblLook w:firstRow="1"/><w:tblW w:w="8000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="7000"/></w:tblGrid>
<w:tr><w:tc><w:tcPr><w:tcW w:w="1000" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:r><w:t>类别</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="7000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>配置要求</w:t></w:r></w:p></w:tc></w:tr>
<w:tr><w:tc><w:tcPr><w:tcW w:w="1000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>设备</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="7000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>满足项目需求的设备配置</w:t></w:r></w:p></w:tc></w:tr>
</w:tbl><w:p><w:r><w:t>后续正文</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:left="1800" w:right="1800"/></w:sectPr></w:body></w:document>'''
    styles = f'''<w:styles xmlns:w="{W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="table" w:styleId="Base"><w:name w:val="Base"/></w:style><w:style w:type="table" w:styleId="Plain"><w:basedOn w:val="Base"/></w:style><w:style w:type="table" w:styleId="Fancy"><w:tblStylePr w:type="firstRow"><w:tcPr><w:shd w:fill="D9E1F2"/></w:tcPr></w:tblStylePr></w:style></w:styles>'''
    return {
        "word/document.xml": document.encode(), "word/styles.xml": styles.encode(),
        "word/_rels/document.xml.rels": f'<Relationships xmlns="{REL}"><Relationship Id="rId1" Type="{R}/image" Target="media/image1.png"/><Relationship Id="rId2" Type="{R}/image" Target="media/image2.png"/></Relationships>'.encode(),
        "word/media/image1.png": PNG, "word/media/image2.png": PNG + b"second",
    }


def mutate_xml(files: dict, fn, part: str = "word/document.xml") -> dict:
    result = dict(files)
    root = ET.fromstring(result[part])
    fn(root)
    result[part] = ET.tostring(root, encoding="utf-8")
    return result


class AppearanceTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.files = fixture()
        self.source = self.save("source.docx", self.files)
        self.baseline = snapshot_docx(self.source)

    def tearDown(self):
        self.tmp.cleanup()

    def save(self, name, files):
        path = self.root / name
        with zipfile.ZipFile(path, "w") as z:
            for key, data in files.items():
                z.writestr(key, data)
        return path

    def diff(self, files, scope="table-layout"):
        return changed_invariants(self.baseline, snapshot_docx(self.save("candidate.docx", files)), scope)

    def test_geometry_repair_preserves_original_style(self):
        repaired, _ = repair_document_xml(self.files["word/document.xml"])
        output = self.root / "repaired.docx"
        write_docx(self.source, output, repaired)
        candidate = snapshot_docx(output)
        self.assertEqual(changed_invariants(self.baseline, candidate, "table-layout"), [])
        self.assertNotEqual(self.baseline["table_geometry"], candidate["table_geometry"])

    def test_frequent_template_style_is_not_copied(self):
        template_files = mutate_xml(self.files, lambda root: root.find('.//' + q(W, 'tblStyle')).set(q(W, 'val'), 'Fancy'))
        template = self.save("template.docx", template_files)
        repaired, _ = repair_document_xml(self.files["word/document.xml"], template)
        root = ET.fromstring(repaired)
        self.assertEqual(root.find('.//' + q(W, 'tblStyle')).get(q(W, 'val')), "Plain")
        output = self.root / "repaired.docx"
        write_docx(self.source, output, repaired)
        self.assertEqual(changed_invariants(self.baseline, snapshot_docx(output), "table-layout"), [])

    def test_direct_header_shading_is_blocked(self):
        files = mutate_xml(self.files, lambda root: ET.SubElement(root.find('.//' + q(W, 'tcPr')), q(W, 'shd'), {q(W, 'fill'): 'D9E1F2'}))
        self.assertIn("table_appearance", self.diff(files))

    def test_changed_table_style_is_blocked(self):
        files = mutate_xml(self.files, lambda root: root.find('.//' + q(W, 'tblStyle')).set(q(W, 'val'), 'Fancy'))
        self.assertIn("table_appearance", self.diff(files))

    def test_conditional_style_change_under_same_id_is_blocked(self):
        def change(root):
            style = root.find(f"{q(W, 'style')}[@{q(W, 'styleId')}='Plain']")
            cond = ET.SubElement(style, q(W, 'tblStylePr'), {q(W, 'type'): 'firstRow'})
            ET.SubElement(ET.SubElement(cond, q(W, 'tcPr')), q(W, 'shd'), {q(W, 'fill'): 'AAAAAA'})
        files = mutate_xml(self.files, change, "word/styles.xml")
        self.assertIn("table_appearance", self.diff(files))

    def test_inherited_style_shading_is_blocked(self):
        def change(root):
            style = root.find(f"{q(W, 'style')}[@{q(W, 'styleId')}='Base']")
            ET.SubElement(ET.SubElement(style, q(W, 'tblPr')), q(W, 'shd'), {q(W, 'fill'): 'AAAAAA'})
        self.assertIn("table_appearance", self.diff(mutate_xml(self.files, change, "word/styles.xml")))

    def test_existing_colored_header_is_preserved(self):
        files = mutate_xml(self.files, lambda root: ET.SubElement(root.find('.//' + q(W, 'tcPr')), q(W, 'shd'), {q(W, 'fill'): '70AD47'}))
        source = self.save("colored.docx", files)
        repaired, _ = repair_document_xml(files['word/document.xml'])
        output = self.root / "colored-repaired.docx"
        write_docx(source, output, repaired)
        self.assertEqual(snapshot_docx(source)['table_appearance'], snapshot_docx(output)['table_appearance'])

    def test_table_look_change_is_blocked(self):
        files = mutate_xml(self.files, lambda root: root.find('.//' + q(W, 'tblLook')).set(q(W, 'firstRow'), '0'))
        self.assertIn("table_appearance", self.diff(files))

    def test_table_border_color_change_is_blocked(self):
        def change(root):
            borders = ET.SubElement(root.find('.//' + q(W, 'tblPr')), q(W, 'tblBorders'))
            ET.SubElement(borders, q(W, 'top'), {q(W, 'val'): 'single', q(W, 'color'): '00FF00'})
        self.assertIn("table_appearance", self.diff(mutate_xml(self.files, change)))

    def test_resizing_keeps_media_and_passes_image_layout(self):
        def change(root):
            root.find('.//' + q(WP, 'extent')).set('cx', '4000')
            root.find('.//' + q(WP, 'extent')).set('cy', '2000')
            root.find('.//' + q(A, 'ext')).set('cx', '4000')
            root.find('.//' + q(A, 'ext')).set('cy', '2000')
        self.assertEqual(self.diff(mutate_xml(self.files, change), 'image-layout'), [])

    def test_image_byte_replacement_is_blocked(self):
        files = dict(self.files, **{'word/media/image1.png': PNG + b'redrawn'})
        self.assertIn('media', self.diff(files, 'image-layout'))

    def test_switching_to_another_embedded_image_is_blocked(self):
        files = mutate_xml(self.files, lambda root: root.find('.//' + q(A, 'blip')).set(q(R, 'embed'), 'rId2'))
        self.assertIn('drawing_content', self.diff(files, 'image-layout'))

    def test_shape_color_change_is_blocked(self):
        files = mutate_xml(self.files, lambda root: root.find('.//' + q(A, 'srgbClr')).set('val', 'D9E1F2'))
        self.assertIn('drawing_content', self.diff(files, 'image-layout'))

    def test_crop_is_not_a_layout_fix(self):
        files = mutate_xml(self.files, lambda root: root.find('.//' + q(A, 'srcRect')).set('l', '12000'))
        self.assertIn('drawing_content', self.diff(files, 'image-layout'))

    def test_dedicated_page_flags_preserve_media(self):
        def change(root):
            ps = root.findall(f"{q(W, 'body')}/{q(W, 'p')}")
            for paragraph, flags in ((ps[1], ['pageBreakBefore', 'keepNext', 'keepLines']), (ps[2], ['keepLines']), (ps[3], ['pageBreakBefore'])):
                ppr = ET.Element(q(W, 'pPr')); paragraph.insert(0, ppr)
                for flag in flags:
                    ET.SubElement(ppr, q(W, flag), {q(W, 'val'): '1'})
        self.assertEqual(self.diff(mutate_xml(self.files, change), 'figure-pagination'), [])

    def test_dedicated_page_does_not_allow_font_change(self):
        def change(root):
            run = root.find('.//' + q(W, 'r'))
            rpr = ET.Element(q(W, 'rPr')); run.insert(0, rpr)
            ET.SubElement(rpr, q(W, 'sz'), {q(W, 'val'): '40'})
        self.assertIn('text_style_without_pagination', self.diff(mutate_xml(self.files, change), 'figure-pagination'))

    def test_dedicated_page_does_not_allow_landscape(self):
        files = mutate_xml(self.files, lambda root: root.find('.//' + q(W, 'pgSz')).set(q(W, 'orient'), 'landscape'))
        self.assertIn('sections', self.diff(files, 'figure-pagination'))

    def test_non_color_font_fixes_still_work(self):
        def change(root):
            run = root.find('.//' + q(W, 'r'))
            rpr = ET.Element(q(W, 'rPr')); run.insert(0, rpr)
            ET.SubElement(rpr, q(W, 'rFonts'), {q(W, 'eastAsia'): '宋体'})
        self.assertEqual(self.diff(mutate_xml(self.files, change), 'text-style'), [])

    def test_old_snapshots_fail_closed(self):
        old = dict(self.baseline); old.pop('table_appearance')
        self.assertIn('table_appearance', changed_invariants(old, self.baseline, 'table-layout'))

    def test_writer_does_not_publish_unauthorized_shading(self):
        files = mutate_xml(self.files, lambda root: ET.SubElement(root.find('.//' + q(W, 'tcPr')), q(W, 'shd'), {q(W, 'fill'): 'DDDDDD'}))
        output = self.root / 'blocked.docx'
        with self.assertRaisesRegex(ValueError, 'table_appearance'):
            write_docx(self.source, output, files['word/document.xml'])
        self.assertFalse(output.exists())
        self.assertEqual(list(self.root.glob('.table-layout-*')), [])

    def test_writer_preserves_existing_output_on_failure(self):
        files = mutate_xml(self.files, lambda root: ET.SubElement(root.find('.//' + q(W, 'tcPr')), q(W, 'shd'), {q(W, 'fill'): 'DDDDDD'}))
        output = self.save('previous-pass.docx', self.files)
        before = output.read_bytes()
        with self.assertRaises(ValueError):
            write_docx(self.source, output, files['word/document.xml'])
        self.assertEqual(output.read_bytes(), before)

    def test_cell_text_color_change_is_blocked(self):
        def change(root):
            run = root.find('.//' + q(W, 'tbl')).find('.//' + q(W, 'r'))
            rpr = ET.Element(q(W, 'rPr')); run.insert(0, rpr)
            ET.SubElement(rpr, q(W, 'color'), {q(W, 'val'): 'CC0000'})
        self.assertIn('table_appearance', self.diff(mutate_xml(self.files, change)))

    def test_theme_color_change_is_blocked_but_font_theme_is_not(self):
        themed = mutate_xml(self.files, lambda root: ET.SubElement(root.find('.//' + q(W, 'tcPr')), q(W, 'shd'), {q(W, 'themeFill'): 'accent1'}))
        themed['word/theme/theme1.xml'] = f'<a:theme xmlns:a="{A}"><a:themeElements><a:clrScheme name="colors"><a:accent1><a:srgbClr val="70AD47"/></a:accent1></a:clrScheme><a:fontScheme name="fonts"><a:majorFont><a:latin typeface="Arial"/></a:majorFont></a:fontScheme></a:themeElements></a:theme>'.encode()
        base = snapshot_docx(self.save('theme-source.docx', themed))
        changed_color = mutate_xml(themed, lambda root: root.find('.//' + q(A, 'srgbClr')).set('val', '111111'), 'word/theme/theme1.xml')
        self.assertIn('table_appearance', changed_invariants(base, snapshot_docx(self.save('theme-color.docx', changed_color)), 'table-layout'))
        changed_font = mutate_xml(themed, lambda root: root.find('.//' + q(A, 'latin')).set('typeface', 'Calibri'), 'word/theme/theme1.xml')
        self.assertNotIn('table_appearance', changed_invariants(base, snapshot_docx(self.save('theme-font.docx', changed_font)), 'table-layout'))

    def test_media_relationship_retargeting_is_blocked(self):
        files = mutate_xml(self.files, lambda root: root.find(q(REL, 'Relationship')).set('Target', 'media/image2.png'), 'word/_rels/document.xml.rels')
        self.assertIn('relationships', self.diff(files, 'image-layout'))

    def test_vml_resizing_allowed_but_recoloring_blocked(self):
        def insert(root):
            run = root.find('.//' + q(W, 'r'))
            pict = ET.SubElement(run, q(W, 'pict'))
            ET.SubElement(pict, '{urn:schemas-microsoft-com:vml}shape', {'style': 'width:100pt;height:50pt', 'fillcolor': '#70AD47'})
        files = mutate_xml(self.files, insert)
        base = snapshot_docx(self.save('vml-source.docx', files))
        resized = mutate_xml(files, lambda root: root.find('.//{urn:schemas-microsoft-com:vml}shape').set('style', 'width:200pt;height:100pt'))
        self.assertEqual(changed_invariants(base, snapshot_docx(self.save('vml-sized.docx', resized)), 'image-layout'), [])
        colored = mutate_xml(files, lambda root: root.find('.//{urn:schemas-microsoft-com:vml}shape').set('fillcolor', '#D9E1F2'))
        self.assertIn('drawing_content', changed_invariants(base, snapshot_docx(self.save('vml-color.docx', colored)), 'image-layout'))

    def test_text_content_fix_scope_remains_usable(self):
        files = mutate_xml(self.files, lambda root: setattr(root.find('.//' + q(W, 't')), 'text', '原说明。'))
        self.assertEqual(self.diff(files, 'text-content'), [])

    def test_width_repair_does_not_rewrite_cell_text(self):
        files = mutate_xml(self.files, lambda root: setattr(root.find('.//' + q(W, 'tbl')).find('.//' + q(W, 't')), 'text', '类\u200b别'))
        source = self.save('zero-width.docx', files)
        repaired, _ = repair_document_xml(files['word/document.xml'])
        output = self.root / 'zero-width-repaired.docx'
        write_docx(source, output, repaired)
        self.assertEqual(snapshot_docx(source)['text_content'], snapshot_docx(output)['text_content'])

    def test_cli_reports_guard_only_after_success(self):
        import json
        import subprocess
        output, report = self.root / 'cli.docx', self.root / 'cli.json'
        cp = subprocess.run([sys.executable, str(SCRIPTS / 'table_layout_repair.py'), str(self.source), '--output', str(output), '--json-out', str(report)], capture_output=True, text=True)
        self.assertEqual(cp.returncode, 0, cp.stderr)
        self.assertEqual(json.loads(report.read_text())['scope_guard'], 'passed')
        self.assertTrue(output.is_file())

    def test_skill_preserves_existing_checks_and_removes_blanket_redraw(self):
        skill = (SCRIPTS.parent / 'SKILL.md').read_text(encoding='utf-8')
        for required in ('numbering_audit.py', '--text-rules', 'hard_text_audit.py', 'table_layout_audit.py', 'figure-pagination', 'layout_invariant_guard.py'):
            self.assertIn(required, skill)
        self.assertNotIn('不能只缩放、擦线或局部挪框。', skill)
        self.assertIn('独占一页', skill)
        self.assertIn('原来没有底色', skill)

    def test_writer_rejects_overwriting_input(self):
        with self.assertRaises(ValueError):
            write_docx(self.source, self.source, self.files['word/document.xml'])


if __name__ == '__main__':
    unittest.main()
