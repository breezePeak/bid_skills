"""Regression tests for the rule interface; no Word/LibreOffice/network required."""
from __future__ import annotations
import contextlib
import io
import json
import shutil
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch
from lxml import etree as ET

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))
from text_rules import (TextRulesError, load_rules, load_profile, expected_body_props,
                        effective_run_props, rules_digest, font_key)
from template_usage_audit import audit
from template_usage_repair import repair
from template_style_profile import extract_profile
import inspect_document
import review_pipeline

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}
Q = lambda n: f"{{{W}}}{n}"
STYLES = f'''<w:styles xmlns:w="{W}">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体" w:cs="Times New Roman"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:styleId="Body" w:default="1"><w:name w:val="正文"/><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体" w:cs="Times New Roman"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Body"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Caption"><w:name w:val="caption"/><w:basedOn w:val="Body"/></w:style>
<w:style w:type="paragraph" w:styleId="TableText"><w:name w:val="表格文字"/><w:basedOn w:val="Body"/></w:style>
<w:style w:type="character" w:styleId="Emph"><w:name w:val="Emph"/><w:rPr><w:b/><w:bCs/><w:i/><w:iCs/></w:rPr></w:style>
<w:style w:type="character" w:styleId="Hei"><w:name w:val="Hei"/><w:rPr><w:rFonts w:eastAsia="黑体"/><w:sz w:val="30"/><w:b/><w:bCs/></w:rPr></w:style>
</w:styles>'''
DOCUMENT = f'''<w:document xmlns:w="{W}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14"><w:body>
<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>第一章</w:t></w:r></w:p>
<w:p w14:paraId="00112233"><w:pPr><w:pStyle w:val="Body"/><w:ind w:firstLineChars="200"/></w:pPr><w:bookmarkStart w:id="1" w:name="body"/><w:r><w:t>正文内容 ABC，2026年。</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>
<w:p><w:pPr><w:pStyle w:val="Caption"/></w:pPr><w:r><w:t>图1 题注</w:t></w:r></w:p>
<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:pStyle w:val="TableText"/></w:pPr><w:r><w:rPr><w:rFonts w:eastAsia="黑体"/></w:rPr><w:t>表格文字</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
</w:body></w:document>'''


def make_doc(path):
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/document.xml", DOCUMENT)
        z.writestr("word/media/sentinel.bin", b"image bytes must not change")
        z.writestr("word/numbering.xml", f'<w:numbering xmlns:w="{W}"/>')
        z.writestr("word/header1.xml", f'<w:hdr xmlns:w="{W}"><w:p><w:r><w:t>原页眉</w:t></w:r></w:p></w:hdr>')
        z.writestr("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>')


def mutate(path, fn, part="word/document.xml"):
    with zipfile.ZipFile(path) as z:
        files = {n: z.read(n) for n in z.namelist()}
    root = ET.fromstring(files[part]); fn(root)
    files[part] = ET.tostring(root, encoding="utf-8")
    with zipfile.ZipFile(path, "w") as z:
        for name, value in files.items():
            z.writestr(name, value)


def body_run(root):
    return root.find("w:body/w:p[2]/w:r", NS)


def put_run(path, inner):
    def update(root):
        r = body_run(root)
        old = r.find("w:rPr", NS)
        if old is not None:
            r.remove(old)
        r.insert(0, ET.fromstring(f'<w:rPr xmlns:w="{W}">{inner}</w:rPr>'))
    mutate(path, update)


def props(path):
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    return effective_run_props(load_profile(path), "Body", body_run(root).find("w:rPr", NS))


class RulesTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.dir = Path(self.temp.name)
        self.template = self.dir / "template.docx"
        self.src = self.dir / "source.docx"
        self.out = self.dir / "fixed.docx"
        make_doc(self.template); shutil.copy2(self.template, self.src)

    def rule_file(self, body):
        path = self.dir / "rules.json"
        path.write_text(json.dumps({"body": body}, ensure_ascii=False), encoding="utf-8")
        return path

    def fix_and_check(self, body=None):
        rules = {"body": body} if body is not None else None
        result = repair(self.template, self.src, self.out, rules)
        checked = audit(self.template, self.out, rules)
        self.assertEqual(checked["status"], "passed", checked)
        self.assertEqual(result["text_rules_sha256"], checked["text_rules_sha256"])
        return props(self.out)

    def test_01_default_hei_font_only_is_preserved(self):
        put_run(self.src, '<w:rFonts w:eastAsia="黑体"/>')
        self.assertEqual(audit(self.template, self.src)["status"], "passed")
        self.assertEqual(self.fix_and_check()["fonts"]["eastAsia"], "黑体")

    def test_02_default_size_bold_italic_still_repaired(self):
        put_run(self.src, '<w:rFonts w:eastAsia="SimHei"/><w:sz w:val="32"/><w:b/><w:i/>')
        issues = audit(self.template, self.src)["issues"]
        self.assertTrue(any("size" in i["code"] for i in issues))
        self.assertTrue(any("bold" in i["code"] for i in issues))
        result = self.fix_and_check()
        self.assertEqual(result["fonts"]["eastAsia"], "SimHei")
        self.assertEqual(result["size_half_points"], "24")
        self.assertFalse(result.get("bold", False)); self.assertFalse(result.get("italic", False))

    def test_03_forbidden_cancels_default_without_extra_fields(self):
        put_run(self.src, '<w:rFonts w:eastAsia="黑体"/>')
        rule = {"forbidden_fonts": ["SimHei"]}
        self.assertEqual(audit(self.template, self.src, {"body": rule})["status"], "failed")
        self.assertEqual(self.fix_and_check(rule)["fonts"]["eastAsia"], "宋体")

    def test_04_empty_preserve_list(self):
        put_run(self.src, '<w:rFonts w:eastAsia="SimHei"/>')
        self.assertEqual(self.fix_and_check({"preserve_fonts": []})["fonts"]["eastAsia"], "宋体")

    def test_05_alias_case_and_vertical_font(self):
        for name in ("黑体", "SimHei", "simhei", "@SIMHEI"):
            with self.subTest(name=name):
                put_run(self.src, f'<w:rFonts w:eastAsia="{name}"/>')
                self.assertEqual(self.fix_and_check({"forbidden_fonts": ["黑体"]})["fonts"]["eastAsia"], "宋体")

    def test_06_other_font_can_be_preserved_by_config(self):
        put_run(self.src, '<w:rFonts w:eastAsia="仿宋"/>')
        self.assertEqual(audit(self.template, self.src)["status"], "failed")
        self.assertEqual(self.fix_and_check({"preserve_fonts": ["仿宋"]})["fonts"]["eastAsia"], "仿宋")

    def test_07_explicit_font_overrides_default_preservation(self):
        put_run(self.src, '<w:rFonts w:eastAsia="黑体"/>')
        result = self.fix_and_check({"fonts": {"eastAsia": "仿宋"}})
        self.assertEqual(result["fonts"]["eastAsia"], "仿宋")
        self.assertEqual(result["fonts"]["ascii"], "Times New Roman")

    def test_08_inherited_character_font_ban(self):
        put_run(self.src, '<w:rStyle w:val="Hei"/>')
        self.assertEqual(self.fix_and_check({"forbidden_fonts": ["黑体"]})["fonts"]["eastAsia"], "宋体")

    def test_09_inherited_paragraph_font_ban_no_rpr(self):
        mutate(self.src, lambda root: root.find("w:style[@w:styleId='Body']/w:rPr/w:rFonts", NS).set(Q("eastAsia"), "黑体"), "word/styles.xml")
        self.assertEqual(audit(self.template, self.src, {"body": {"forbidden_fonts": ["黑体"]}})["status"], "failed")
        self.assertEqual(self.fix_and_check({"forbidden_fonts": ["黑体"]})["fonts"]["eastAsia"], "宋体")

    def test_10_size_override_half_point(self):
        result = self.fix_and_check({"size_pt": 10.5})
        self.assertEqual(result["size_half_points"], "21")
        self.assertEqual(result["size_cs_half_points"], "21")

    def test_11_direct_emphasis_preserve(self):
        put_run(self.src, '<w:b/><w:bCs/><w:i/><w:iCs/><w:sz w:val="32"/>')
        result = self.fix_and_check({"bold": "preserve", "italic": "preserve"})
        self.assertTrue(all(result[k] for k in ("bold", "bold_cs", "italic", "italic_cs")))

    def test_12_inherited_emphasis_forbid(self):
        put_run(self.src, '<w:rStyle w:val="Emph"/>')
        rule = {"bold": "forbid", "italic": "forbid"}
        self.assertEqual(audit(self.template, self.src, {"body": rule})["status"], "failed")
        result = self.fix_and_check(rule)
        self.assertFalse(any(result.get(k, False) for k in ("bold", "bold_cs", "italic", "italic_cs")))

    def test_13_emphasis_require(self):
        result = self.fix_and_check({"bold": "require", "italic": "require"})
        self.assertTrue(all(result[k] for k in ("bold", "bold_cs", "italic", "italic_cs")))

    def test_14_template_banned_conflict_no_output(self):
        mutate(self.template, lambda root: root.find("w:style[@w:styleId='Body']/w:rPr/w:rFonts", NS).set(Q("eastAsia"), "黑体"), "word/styles.xml")
        with self.assertRaises(TextRulesError):
            repair(self.template, self.src, self.out, {"body": {"forbidden_fonts": ["黑体"]}})
        self.assertFalse(self.out.exists())
        self.fix_and_check({"forbidden_fonts": ["黑体"], "fonts": {"eastAsia": "宋体"}})

    def test_15_unknown_fields_are_not_ignored(self):
        for rule in ({"boddy": {}}, {"body": {"fontz": {}}}, {"body": {"fonts": {"eastasia": "宋体"}}}):
            with self.subTest(rule=rule), self.assertRaises(TextRulesError):
                load_rules(rule)

    def test_16_invalid_values(self):
        for body in ({"size_pt": 12.2}, {"size_pt": True}, {"size_pt": float("nan")},
                     {"size_pt": 0}, {"preserve_fonts": "黑体"}, {"fonts": {"eastAsia": ""}},
                     {"bold": "ignore"}, {"italic": None}):
            with self.subTest(body=body), self.assertRaises(TextRulesError):
                load_rules({"body": body})
        for root in ([], None, False):
            if root is None:
                continue
            with self.subTest(root=root), self.assertRaises(TextRulesError):
                load_rules(root)

    def test_17_explicit_conflicting_rules(self):
        for body in ({"preserve_fonts": ["黑体"], "forbidden_fonts": ["SimHei"]},
                     {"fonts": {"eastAsia": "黑体"}, "forbidden_fonts": ["SimHei"]}):
            with self.assertRaises(TextRulesError):
                load_rules({"body": body})

    def test_18_bom_and_duplicate_keys(self):
        file = self.dir / "bom.json"
        file.write_text('{"body":{"forbidden_fonts":["黑体"]}}', encoding="utf-8-sig")
        self.assertEqual(load_rules(file)["body"]["preserve_fonts"], [])
        file.write_text('{"body":{"bold":"preserve","bold":"forbid"}}', encoding="utf-8")
        with self.assertRaises(TextRulesError): load_rules(file)

    def test_19_no_cross_task_state(self):
        one = load_rules({"body": {"forbidden_fonts": ["黑体"]}})
        default = load_rules()
        self.assertEqual(one["body"]["preserve_fonts"], [])
        self.assertIn("黑体", default["body"]["preserve_fonts"])
        one["body"]["bold"] = "forbid"
        self.assertEqual(load_rules()["body"]["bold"], "template")

    def test_20_idempotent_and_other_parts_unchanged(self):
        put_run(self.src, '<w:rFonts w:eastAsia="黑体"/><w:b/><w:sz w:val="30"/>')
        rules = {"body": {"forbidden_fonts": ["黑体"], "bold": "preserve", "size_pt": 11}}
        repair(self.template, self.src, self.out, rules)
        out2 = self.dir / "fixed-again.docx"
        result = repair(self.template, self.out, out2, rules)
        self.assertEqual(result["change_count"], 0)
        with zipfile.ZipFile(self.src) as src, zipfile.ZipFile(self.out) as out, zipfile.ZipFile(out2) as again:
            for name in src.namelist():
                if name != "word/document.xml": self.assertEqual(src.read(name), out.read(name), name)
            self.assertEqual(out.read("word/document.xml"), again.read("word/document.xml"))
            a=ET.fromstring(src.read("word/document.xml")); b=ET.fromstring(out.read("word/document.xml"))
            self.assertEqual(a.xpath("//w:t/text()", namespaces=NS), b.xpath("//w:t/text()", namespaces=NS))
            self.assertEqual(ET.tostring(a.find("w:body/w:tbl", NS)), ET.tostring(b.find("w:body/w:tbl", NS)))
            self.assertIn(b'mc:Ignorable="w14"', out.read("word/document.xml"))

    def test_21_cli_invalid_rules_no_output(self):
        file = self.dir / "bad.json"; file.write_text('{"body":{"wrong":true}}')
        result = subprocess.run([sys.executable, str(SCRIPTS/"template_usage_repair.py"),
            str(self.template), str(self.src), "--out", str(self.out), "--text-rules", str(file)], capture_output=True, text=True)
        self.assertEqual(result.returncode, 2); self.assertFalse(self.out.exists())
        self.assertEqual(json.loads(result.stdout)["code"], "text-rules-invalid")

    def test_22_enforce_repair_audit_cli_rule_chain(self):
        put_run(self.src, '<w:rStyle w:val="Hei"/>')
        file = self.rule_file({"forbidden_fonts": ["黑体"], "fonts": {"eastAsia": "仿宋"},
                               "size_pt": 11, "bold": "preserve"})
        style_json = self.dir / "profile.json"; style_json.write_text(json.dumps(extract_profile(self.template)))
        applied = self.dir / "applied.docx"
        commands = [
            ["template_style_enforce.py", self.template, style_json, self.src, "--out", applied],
            ["template_usage_repair.py", self.template, applied, "--out", self.out],
            ["template_usage_audit.py", self.template, self.out],
        ]
        hashes=[]
        for cmd in commands:
            p = subprocess.run([sys.executable, str(SCRIPTS/cmd[0]), *map(str, cmd[1:]), "--text-rules", str(file)], capture_output=True, text=True)
            self.assertEqual(p.returncode, 0, p.stdout+p.stderr)
            hashes.append(json.loads(p.stdout)["text_rules_sha256"])
        self.assertEqual(len(set(hashes)), 1)
        result=props(self.out)
        self.assertEqual(result["fonts"]["eastAsia"], "仿宋")
        self.assertTrue(result["bold"]); self.assertEqual(result["size_half_points"], "22")

    def test_23_enforce_default_preserves_inherited_hei(self):
        mutate(self.src, lambda root: root.find("w:style[@w:styleId='Body']/w:rPr/w:rFonts", NS).set(Q("eastAsia"), "黑体"), "word/styles.xml")
        style_json = self.dir / "profile.json"; style_json.write_text(json.dumps(extract_profile(self.template)))
        cp = subprocess.run([sys.executable, str(SCRIPTS/"template_style_enforce.py"), str(self.template), str(style_json), str(self.src), "--out", str(self.out)], capture_output=True, text=True)
        self.assertEqual(cp.returncode, 0, cp.stdout+cp.stderr)
        self.assertEqual(props(self.out)["fonts"]["eastAsia"], "黑体")

    def test_24_theme_font_cannot_hide_forbidden_name(self):
        theme = '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme><a:minorFont><a:ea typeface="SimHei"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>'
        with zipfile.ZipFile(self.src, "a") as z: z.writestr("word/theme/theme1.xml", theme)
        put_run(self.src, '<w:rFonts w:eastAsia="宋体" w:eastAsiaTheme="minorEastAsia"/>')
        rule={"forbidden_fonts": ["黑体"]}
        self.assertEqual(audit(self.template,self.src,{"body":rule})["status"],"failed")
        self.assertEqual(self.fix_and_check(rule)["fonts"]["eastAsia"],"宋体")

    def test_25_forbidden_baseline_then_override_is_safe(self):
        mutate(self.template, lambda root: root.find("w:style[@w:styleId='Body']/w:rPr/w:rFonts", NS).set(Q("eastAsia"), "SimHei"), "word/styles.xml")
        expected=expected_body_props(load_profile(self.template), load_rules({"body":{"forbidden_fonts":["黑体"], "fonts":{"eastAsia":"宋体"}}}))
        self.assertEqual(expected["fonts"]["eastAsia"], "宋体")

    def test_26_source_and_template_not_overwritten(self):
        with self.assertRaises(TextRulesError): repair(self.template,self.src,self.src)
        with self.assertRaises(TextRulesError): repair(self.template,self.src,self.template)

    def test_27_size_only_leaves_defaults_intact(self):
        rule=load_rules({"body":{"size_pt":14}})
        self.assertIn("黑体",rule["body"]["preserve_fonts"])
        self.assertEqual(rule["body"]["bold"],"template")
        self.assertEqual(rule["body"]["fonts"],{})

    def test_28_roundtrip_effective_snapshot(self):
        rules=load_rules({"body":{"forbidden_fonts":["SimHei"],"size_pt":14}})
        file=self.dir/'effective.json';file.write_text(json.dumps(rules))
        self.assertEqual(rules_digest(rules), rules_digest(load_rules(file)))

    def _fake_runner(self, calls, mismatch=False, default_allow=(0,)):
        def run(name,*args,allow=None):
            allow = default_allow if allow is None else allow
            args=list(map(str,args));calls.append((name,args))
            if name in {"template_usage_audit.py","template_style_enforce.py","template_usage_repair.py"}:
                self.assertIn("--text-rules",args)
                cp=subprocess.run([sys.executable,str(SCRIPTS/name),*args],capture_output=True,text=True)
                if mismatch and name=="template_style_enforce.py":
                    report=Path(args[args.index('--json-out')+1]);data=json.loads(report.read_text());data['text_rules_sha256']='wrong';report.write_text(json.dumps(data))
                return {'ok':cp.returncode in allow,'returncode':cp.returncode,'stdout':cp.stdout,'stderr':cp.stderr}
            if name=='render_docx.py':
                page=Path(args[args.index('--out-dir')+1])/'page-1.png';page.write_bytes(b'test render placeholder')
                return {'ok':True,'returncode':0,'stdout':json.dumps({'pages':[str(page)]}),'stderr':''}
            if '--out' in args or '--output' in args:
                index=args.index('--out') if '--out' in args else args.index('--output')
                shutil.copy2(args[0],args[index+1])
            if '--json-out' in args:
                path=Path(args[args.index('--json-out')+1]);path.write_text(json.dumps({'status':'passed','issues':[],'issue_count':0}))
            return {'ok':True,'returncode':0,'stdout':'{}','stderr':''}
        return run

    def test_29_pipeline_and_inspection_forward_rules(self):
        file=self.rule_file({'forbidden_fonts':['黑体'], 'size_pt':13})
        put_run(self.src,'<w:rFonts w:eastAsia="黑体"/>')
        style_json=self.dir/'profile.json';style_json.write_text(json.dumps(extract_profile(self.template)))
        calls=[];work=self.dir/'work'
        arguments=[str(self.src),'--template',str(self.template),'--template-style-json',str(style_json),
                   '--work-dir',str(work),'--text-rules',str(file)]
        for module in (inspect_document,review_pipeline):
            with patch.object(sys,'argv',[module.__name__,*arguments]),patch.object(module,'run_script',self._fake_runner(calls, default_allow=(0,2,3,4) if module is inspect_document else (0,))),contextlib.redirect_stdout(io.StringIO()):
                self.assertEqual(module.main(),0)
        first=json.loads((work/'inspection.json').read_text())
        final=json.loads((work/'review-manifest.json').read_text())
        self.assertEqual(first['text_rules_sha256'],final['text_rules_sha256'])
        self.assertEqual(final['status'],'awaiting_visual_review')
        for name,args in calls:
            if name in {'template_usage_audit.py','template_usage_repair.py','template_style_enforce.py'}:
                self.assertEqual(args[args.index('--text-rules')+1],str(work/'reports/text-rules.effective.json'))
        self.assertEqual(props(work/'candidate.docx')['fonts']['eastAsia'],'宋体')

    def test_30_rule_mismatch_stops_pipeline(self):
        file=self.rule_file({'forbidden_fonts':['黑体']})
        style_json=self.dir/'profile.json';style_json.write_text(json.dumps(extract_profile(self.template)))
        calls=[];work=self.dir/'work'
        args=[str(self.src),'--template',str(self.template),'--template-style-json',str(style_json),'--work-dir',str(work),'--text-rules',str(file)]
        with patch.object(sys,'argv',['review_pipeline',*args]),patch.object(review_pipeline,'run_script',self._fake_runner(calls,True)),contextlib.redirect_stdout(io.StringIO()):
            self.assertEqual(review_pipeline.main(),2)
        self.assertFalse((work/'candidate.docx').exists())
        self.assertFalse(any(n=='template_usage_repair.py' for n,a in calls))

    def test_31_implicit_body_without_heading_is_checked(self):
        def change(root):
            body=root.find("w:body",NS)
            body.remove(body.find("w:p",NS))
            paragraph=body.find("w:p",NS)
            ppr=paragraph.find("w:pPr",NS)
            ppr.remove(ppr.find("w:pStyle",NS))
            run=paragraph.find("w:r",NS)
            run.insert(0,ET.fromstring(f'<w:rPr xmlns:w="{W}"><w:rFonts w:eastAsia="黑体"/></w:rPr>'))
        mutate(self.src,change)
        rules={"body":{"forbidden_fonts":["黑体"]}}
        self.assertEqual(audit(self.template,self.src,rules)["status"],"failed")
        repair(self.template,self.src,self.out,rules)
        self.assertEqual(audit(self.template,self.out,rules)["status"],"passed")

    def test_32_named_title_is_not_a_body_rule_target(self):
        def add_style(root):
            root.append(ET.fromstring(f'<w:style xmlns:w="{W}" w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Body"/></w:style>'))
        mutate(self.src,add_style,"word/styles.xml")
        title=f'<w:p xmlns:w="{W}"><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:rPr><w:rFonts w:eastAsia="黑体"/><w:b/></w:rPr><w:t>封面标题</w:t></w:r></w:p>'
        mutate(self.src,lambda root:root.find("w:body",NS).insert(0,ET.fromstring(title)))
        repair(self.template,self.src,self.out,{"body":{"forbidden_fonts":["黑体"],"bold":"forbid","size_pt":11}})
        with zipfile.ZipFile(self.src) as before, zipfile.ZipFile(self.out) as after:
            a=ET.fromstring(before.read("word/document.xml")).find("w:body/w:p",NS)
            b=ET.fromstring(after.read("word/document.xml")).find("w:body/w:p",NS)
            self.assertEqual(ET.tostring(a),ET.tostring(b))


if __name__ == '__main__':
    unittest.main(verbosity=2)
