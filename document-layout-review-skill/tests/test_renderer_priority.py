"""Focused renderer regression tests; native Office calls are explicitly mocked.

Run: python -m unittest discover -s tests -p test_renderer_priority.py -v
The optional smoke script tests an installed real renderer separately.
"""
from __future__ import annotations
import ast
import copy
import importlib.util
import json
import re
import struct
import subprocess
import sys
import tempfile
import types
import unittest
import zipfile
import zlib
from pathlib import Path
from unittest.mock import patch

SCRIPTS = Path(__file__).resolve().parents[1] / 'scripts'
sys.path.insert(0, str(SCRIPTS))
import office_backends as backend
import render_docx as renderer
import runtime_preflight as preflight


def png():
    def chunk(kind, data):
        return struct.pack('>I', len(data)) + kind + data + struct.pack('>I', zlib.crc32(kind+data) & 0xffffffff)
    return b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', 1, 1, 8, 2, 0, 0, 0)) + chunk(b'IDAT', zlib.compress(b'\0\xff\xff\xff')) + chunk(b'IEND', b'')


def docx(path, extra=''):
    with zipfile.ZipFile(path, 'w') as z:
        z.writestr('[Content_Types].xml', '<Types/>')
        z.writestr('word/styles.xml', '<styles/>')
        z.writestr('word/document.xml', '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>test</w:t></w:r>'+extra+'</w:p></w:body></w:document>')


class RendererTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.source = self.root / '中文 input.docx'
        docx(self.source)
        self.probed = []; self.exported = []
        self.availability = {e: True for e in backend.PRIORITY}
        self.failures = set()
        self.patches = [
            patch.object(renderer, 'executable', return_value='/test/pdftoppm'),
            patch.object(renderer, 'probe_renderer', side_effect=self.probe),
            patch.object(renderer, 'export_pdf', side_effect=self.export),
            patch.object(renderer, 'rasterize', side_effect=self.raster),
        ]
        for p in self.patches: p.start()
        self.addCleanup(self.tmp.cleanup)
        for p in self.patches: self.addCleanup(p.stop)

    def probe(self, engine, soffice=None):
        self.probed.append(engine)
        return {'available': self.availability[engine], 'detail': 'not installed', 'soffice': soffice}

    def export(self, engine, source, pdf, work, probe, timeout):
        self.exported.append(engine)
        if engine in self.failures:
            raise backend.OfficeError('office-command-failed', 'test engine failed')
        pdf.write_bytes(b'%PDF-1.7\nunit fixture\n%%EOF')
        return {'engine': backend.ENGINE_NAMES[engine]}

    def raster(self, pdf, folder, ppm, dpi, timeout):
        folder.mkdir()
        page = folder/'page-1.png'; page.write_bytes(png())
        return [page]

    def execute(self, **kwargs):
        return renderer.render(self.source, self.root/'out', **kwargs)

    def valid(self, result, requested='auto'):
        return renderer.validate_render(self.source, result, requested, result['page_records'])

    def test_word_first_and_stop(self):
        result=self.execute(); self.assertEqual(result['renderer'],'word')
        self.assertEqual(self.probed,['word']); self.valid(result)

    def test_word_unavailable_then_wps(self):
        self.availability['word']=False
        result=self.execute(); self.assertEqual(result['renderer'],'wps')
        self.assertEqual(self.probed,['word','wps']); self.valid(result)

    def test_word_runtime_failure_then_wps(self):
        self.failures={'word'}
        result=self.execute(); self.assertEqual(self.exported,['word','wps'])
        self.assertEqual(result['attempts'][0]['status'],'failed'); self.valid(result)

    def test_wps_runtime_failure_then_lo(self):
        self.failures={'word','wps'}
        result=self.execute(); self.assertEqual(result['renderer'],'libreoffice')
        self.assertEqual(self.exported,list(backend.PRIORITY)); self.valid(result)

    def test_only_lo_installed(self):
        self.availability.update(word=False,wps=False)
        result=self.execute(); self.assertEqual(result['renderer'],'libreoffice'); self.valid(result)

    def test_all_unavailable_fail_closed(self):
        self.availability={e:False for e in backend.PRIORITY}
        with self.assertRaises(renderer.RenderError): self.execute()
        result=json.loads((self.root/'out/render-report.json').read_text())
        self.assertEqual(result['status'],'failed'); self.assertEqual(result['pages'],[])
        self.assertEqual([a['engine'] for a in result['attempts']],list(backend.PRIORITY))

    def test_all_runtime_failures(self):
        self.failures=set(backend.PRIORITY)
        with self.assertRaises(renderer.RenderError): self.execute()
        self.assertEqual(self.exported,list(backend.PRIORITY))

    def test_timeout_falls_back(self):
        real=self.export
        def action(engine,*args):
            if engine=='word': raise backend.OfficeError('office-timeout','timeout')
            return real(engine,*args)
        with patch.object(renderer,'export_pdf',side_effect=action): result=self.execute()
        self.assertEqual(result['renderer'],'wps'); self.valid(result)

    def test_explicit_word_does_not_fall_back(self):
        self.failures={'word'}
        with self.assertRaises(renderer.RenderError): self.execute(renderer='word')
        self.assertEqual(self.probed,['word'])

    def test_explicit_wps_works(self):
        result=self.execute(renderer='wps')
        self.assertEqual(self.probed,['wps']); self.valid(result,'wps')

    def test_soffice_path_does_not_override_priority(self):
        result=self.execute(soffice='/custom/soffice')
        self.assertEqual(result['renderer'],'word')

    def test_missing_png_program_blocks_before_export(self):
        with patch.object(renderer,'executable',return_value=None):
            with self.assertRaises(renderer.RenderError): self.execute()
        self.assertEqual(self.exported,[])

    def test_source_never_changed(self):
        before=backend.digest(self.source); self.execute()
        self.assertEqual(before,backend.digest(self.source))

    def test_source_mutation_does_not_fall_back(self):
        real=self.export
        def action(*args):
            result=real(*args); self.source.write_bytes(b'changed'); return result
        with patch.object(renderer,'export_pdf',side_effect=action):
            with self.assertRaises(renderer.RenderError) as ctx: self.execute()
        self.assertEqual(ctx.exception.code,'render-source-changed'); self.assertEqual(self.probed,['word'])

    def test_no_new_pdf_does_not_reuse_old_pdf(self):
        out=self.root/'out'; out.mkdir(); (out/'document.pdf').write_bytes(b'old')
        self.availability.update(wps=False,libreoffice=False)
        with patch.object(renderer,'export_pdf',return_value={}):
            with self.assertRaises(renderer.RenderError): self.execute()

    def test_separate_runs_do_not_mix_old_pages(self):
        a=self.execute(); b=self.execute()
        self.assertNotEqual(a['pages'],b['pages']); self.valid(a); self.valid(b)

    def test_render_success_is_not_visual_pass(self):
        self.assertEqual(self.execute()['visual_review_status'],'pending')

    def test_stale_source_rejected(self):
        result=self.execute(); self.source.write_bytes(b'changed')
        with self.assertRaises(renderer.RenderError): self.valid(result)

    def test_tampered_pdf_rejected(self):
        result=self.execute(); Path(result['pdf']).write_bytes(b'changed')
        with self.assertRaises(renderer.RenderError): self.valid(result)

    def test_tampered_page_rejected(self):
        result=self.execute(); Path(result['pages'][0]).write_bytes(b'changed')
        with self.assertRaises(renderer.RenderError): self.valid(result)

    def test_tampered_report_rejected(self):
        result=self.execute(); Path(result['report_path']).write_text('{}')
        with self.assertRaises(renderer.RenderError): self.valid(result)

    def test_engine_label_mismatch_rejected(self):
        result=self.execute(); result['engine']='WPS Writer'
        with self.assertRaises(renderer.RenderError): self.valid(result)

    def test_skipped_wps_rejected(self):
        self.availability.update(word=False,wps=False)
        result=self.execute(); result['attempts'].pop(1)
        with self.assertRaises(renderer.RenderError): self.valid(result)

    def test_manifest_page_omission_rejected(self):
        result=self.execute()
        with self.assertRaises(renderer.RenderError): renderer.validate_render(self.source,result,'auto',[])

    def test_invalid_engine(self):
        with self.assertRaises(backend.OfficeError): self.execute(renderer='unknown')


class RasterTests(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory(); self.root=Path(self.tmp.name)
        self.addCleanup(self.tmp.cleanup)
        self.pdf=self.root/'a.pdf'; self.pdf.write_bytes(b'%PDF-1.7\n%%EOF')

    def fake_run(self,numbers,corrupt=False):
        def action(cmd,*args,**kwargs):
            folder=Path(cmd[-1]).parent
            for n in numbers: (folder/f'page-{n}.png').write_bytes(b'bad' if corrupt else png())
        return action

    def test_numeric_page_sort(self):
        with patch.object(renderer,'pdf_page_count',return_value=12), patch.object(renderer,'run',side_effect=self.fake_run(range(1,13))):
            result=renderer.rasterize(self.pdf,self.root/'pages','ppm',120,30)
        self.assertEqual([p.name for p in result],[f'page-{i}.png' for i in range(1,13)])

    def test_missing_page_rejected(self):
        with patch.object(renderer,'pdf_page_count',return_value=3), patch.object(renderer,'run',side_effect=self.fake_run([1,3])):
            with self.assertRaises(renderer.RenderError): renderer.rasterize(self.pdf,self.root/'pages','ppm',120,30)

    def test_page_count_mismatch_rejected(self):
        with patch.object(renderer,'pdf_page_count',return_value=3), patch.object(renderer,'run',side_effect=self.fake_run([1,2])):
            with self.assertRaises(renderer.RenderError): renderer.rasterize(self.pdf,self.root/'pages','ppm',120,30)

    def test_corrupt_png_rejected(self):
        with patch.object(renderer,'pdf_page_count',return_value=1), patch.object(renderer,'run',side_effect=self.fake_run([1],True)):
            with self.assertRaises(renderer.RenderError): renderer.rasterize(self.pdf,self.root/'pages','ppm',120,30)

    def test_invalid_pdf_header(self):
        self.pdf.write_bytes(b'not a pdf')
        with self.assertRaises(renderer.RenderError): renderer.pdf_page_count(self.pdf,30)


class PreflightTests(unittest.TestCase):
    def test_auto_field_prefers_word(self):
        with patch.object(preflight,'probe_word',return_value={'available':True}), patch.object(preflight,'probe_wps') as wps:
            self.assertEqual(preflight.select_engine()['engine'],'word'); wps.assert_not_called()

    def test_auto_field_wps_before_lo(self):
        with patch.object(preflight,'probe_word',return_value={'available':False}), patch.object(preflight,'probe_wps',return_value={'available':True}), patch.object(preflight,'probe_libreoffice') as lo:
            self.assertEqual(preflight.select_engine()['engine'],'wps'); lo.assert_not_called()

    def test_explicit_field_unavailable_no_fallback(self):
        with patch.object(preflight,'probe_word',return_value={'available':False}), patch.object(preflight,'probe_wps') as wps:
            with self.assertRaises(preflight.PreflightError): preflight.select_engine('word')
            wps.assert_not_called()

    def test_excluded_failed_word_is_not_retried(self):
        with patch.object(preflight,'probe_word') as word, patch.object(preflight,'probe_wps',return_value={'available':True}):
            self.assertEqual(preflight.select_engine(exclude={'word'})['engine'],'wps'); word.assert_not_called()

    def test_word_only_machine_does_not_require_libreoffice(self):
        with tempfile.TemporaryDirectory() as td:
            source=Path(td)/'input.docx'; docx(source)
            with patch.object(preflight.shutil,'which',side_effect=lambda n:'/bin/pdftoppm' if n.startswith('pdftoppm') else None), patch.object(preflight,'run_probe',return_value={'available':True}), patch.object(preflight,'select_renderer',return_value={'engine':'word'}), patch.object(preflight,'select_engine',return_value={'engine':'word'}):
                result=preflight.check(source,source,Path(td)/'work')
            self.assertEqual(result['status'],'passed'); self.assertEqual(result['renderer']['engine'],'word')

    def test_audit_only_does_not_require_uno(self):
        with tempfile.TemporaryDirectory() as td:
            source=Path(td)/'input.docx'; docx(source)
            with patch.object(preflight.shutil,'which',return_value='/bin/pdftoppm'), patch.object(preflight,'run_probe',return_value={'available':True}), patch.object(preflight,'select_renderer',return_value={'engine':'wps'}), patch.object(preflight,'select_engine') as fields:
                result=preflight.check(source,source,Path(td)/'work',audit_only=True)
            fields.assert_not_called(); self.assertEqual(result['status'],'passed')


class SafetyTests(unittest.TestCase):
    def test_external_split_field_blocked(self):
        with tempfile.TemporaryDirectory() as td:
            source=Path(td)/'a.docx'
            docx(source,'<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>INCLUDE</w:instrText></w:r><w:r><w:instrText>TEXT file</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>')
            with self.assertRaises(renderer.RenderError): renderer.guard_source(source)

    def test_external_hyperlink_allowed(self):
        with tempfile.TemporaryDirectory() as td:
            source=Path(td)/'a.docx'; docx(source)
            with zipfile.ZipFile(source,'a') as z:
                z.writestr('word/_rels/document.xml.rels','<Relationships><Relationship TargetMode="External" Type="test/hyperlink" Target="https://example.invalid"/></Relationships>')
            renderer.guard_source(source)

    def test_external_template_blocked(self):
        with tempfile.TemporaryDirectory() as td:
            source=Path(td)/'a.docx'; docx(source)
            with zipfile.ZipFile(source,'a') as z:
                z.writestr('word/_rels/document.xml.rels','<Relationships><Relationship TargetMode="External" Type="test/attachedTemplate" Target="https://example.invalid"/></Relationships>')
            with self.assertRaises(renderer.RenderError): renderer.guard_source(source)

    def test_finalizer_and_pipeline_call_render_gate(self):
        for name in ('review_pipeline.py','finalize_review.py'):
            tree=ast.parse((SCRIPTS/name).read_text())
            calls=[n.func.id for n in ast.walk(tree) if isinstance(n,ast.Call) and isinstance(n.func,ast.Name)]
            self.assertIn('validate_render',calls)

    def test_required_fourteen_gates_preserved(self):
        tree=ast.parse((SCRIPTS/'review_pipeline.py').read_text())
        value=next(n.value for n in tree.body if isinstance(n,ast.Assign) and any(isinstance(t,ast.Name) and t.id=='REQUIRED_GATES' for t in n.targets))
        self.assertEqual(len(ast.literal_eval(value)),14)

    def test_com_worker_does_not_kill_apps_by_name(self):
        text=(SCRIPTS/'office_com.ps1').read_text(encoding='utf-8-sig')
        self.assertNotRegex(text,r'(?i)\btaskkill\s+/|Stop-Process\s+-Name')
        self.assertIn('$app.Documents.Count -ne 0',text)


class FieldFallbackTests(unittest.TestCase):
    """Isolate existing policy modules; test dispatch, not their internal semantics."""
    def setUp(self):
        class PolicyError(ValueError):
            def __init__(self,code,message,**details):
                super().__init__(message); self.code=code; self.details=details
            def as_issue(self): return {'code':self.code,'message':str(self),**self.details}
        policy=types.ModuleType('numbering_policy')
        for name in ('parse','fields','command','Q'): setattr(policy,name,lambda *a:None)
        policy.file_digest=backend.digest; policy.ERROR_TEXT=re.compile('Error!'); policy.PolicyError=PolicyError
        caption=types.ModuleType('caption_field_guard')
        caption.CaptionFieldError=PolicyError
        caption.require_valid=lambda *a,**k:None
        caption.preserve_native_instructions=lambda *a,**k:None
        with patch.dict(sys.modules,{'numbering_policy':policy,'caption_field_guard':caption}):
            spec=importlib.util.spec_from_file_location('_field_dispatch_test',SCRIPTS/'field_refresh.py')
            self.module=importlib.util.module_from_spec(spec); spec.loader.exec_module(self.module)
        self.error=PolicyError

    def selected(self,requested='auto',uno_python=None,*,exclude=()):
        return {'engine':next(e for e in backend.PRIORITY if e not in exclude),'capabilities':{}}

    def test_actual_field_runtime_fallback(self):
        def action(source,out,engine,*a,**k):
            if engine=='word': raise self.error('office-timeout','test timeout')
            return {'status':'passed','engine':'WPS Writer'}
        with patch.object(preflight,'select_engine',side_effect=self.selected), patch.object(self.module,'_refresh_once',side_effect=action):
            result=self.module.refresh('source','out')
        self.assertEqual(result['engine'],'WPS Writer')
        self.assertEqual([a['engine'] for a in result['engine_attempts']],['word','wps'])

    def test_field_result_error_is_not_hidden_by_fallback(self):
        with patch.object(preflight,'select_engine',side_effect=self.selected) as select, patch.object(self.module,'_refresh_once',side_effect=self.error('field-result-error','invalid caption')):
            with self.assertRaises(self.error): self.module.refresh('source','out')
        self.assertEqual(select.call_count,1)

    def test_explicit_field_engine_does_not_fall_back(self):
        with patch.object(preflight,'select_engine',return_value={'engine':'wps'}), patch.object(self.module,'_refresh_once',side_effect=self.error('office-timeout','timeout')) as once:
            with self.assertRaises(self.error): self.module.refresh('source','out','wps')
        self.assertEqual(once.call_count,1)

    def test_wps_report_is_accepted_when_hash_bound(self):
        with tempfile.TemporaryDirectory() as td:
            source=Path(td)/'x.docx'; source.write_bytes(b'test')
            report={'status':'passed','engine':'WPS Writer','fields_updated':True,'indexes_updated':True,'errors':[],'output_sha256':backend.digest(source)}
            with patch.object(self.module,'result_errors',return_value=[]), patch.object(self.module,'require_caption_fields'):
                self.assertEqual(self.module.validate_report(source,report)['engine'],'WPS Writer')

    def test_incomplete_wps_report_rejected(self):
        with self.assertRaises(self.error): self.module.validate_report('unused',{'status':'passed','engine':'WPS Writer'})


if __name__=='__main__': unittest.main()
