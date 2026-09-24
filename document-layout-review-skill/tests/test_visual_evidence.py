"""Deterministic guard/worker-transport tests. Semantic outputs are test doubles."""
import copy
import io
import json
import sys
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from PIL import Image

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent / 'scripts'))
import visual_evidence as ve
import figure_inspection as fi


class VisualEvidenceTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.image = self.root / 'image.png'
        # Small geometric fixture; this suite does not pretend to run a vision model.
        Image.new('RGB', (120, 90), (240, 245, 249)).save(self.image)
        self.page = self.root / 'page.png'
        Image.new('RGB', (220, 300), 'white').save(self.page)
        self.pages = [{'name': self.page.name, **ve.reference(self.page)}]
        self.source = self.root / 'source.docx'
        self.source.write_bytes(b'source binding fixture')
        self.inspection_patch = patch.object(fi, 'validate_inspection', lambda *a, **kw: ve.validate_inspection(*a, **kw, allow_test_double=True))
        self.inspection_patch.start(); self.addCleanup(self.inspection_patch.stop)
        self.binding = {'source_sha256': ve.file_sha(self.source), 'candidate_sha256': ve.file_sha(self.source),
                        'object_id': 'fixture-1', 'object_sha256': 'original-image', 'media_sha256': [ve.file_sha(self.image)], 'pages': []}

    def prepare(self, final=False, binding=None):
        return ve.prepare([self.image], self.root, binding or self.binding, phase='final' if final else 'initial',
            pages=self.pages if final else (), originals=[self.image] if final else ())

    def worker(self, mode='pass', timeout=10):
        return {'command': [sys.executable, str(HERE / 'fake_vision_worker.py'), mode], 'timeout_seconds': timeout, 'allow_test_double': True}

    def run_review(self, mode='pass', final=False, binding=None):
        return ve.run(self.prepare(final, binding), self.worker(mode), self.root)

    def validate(self, *a, **kw):
        return ve.validate_inspection(*a, **kw, allow_test_double=True)

    def test_marked_mock_cannot_be_used_for_production_acceptance(self):
        with self.assertRaisesRegex(ve.VisualError, '模拟视觉输出'):
            ve.validate_inspection(self.run_review())

    def revise(self, ref, callback):
        path = ve.checked_file(ref)
        value = ve.read(path); callback(value); ve.write(path, value)
        return ve.reference(path)

    def test_small_image_has_full_and_all_edges(self):
        b = ve.validate_bundle(self.prepare())
        self.assertEqual({v['label'] for v in b['views']}, {'full', 'edge-left', 'edge-right', 'edge-top', 'edge-bottom'})

    def test_real_uploaded_negative_gets_overlapping_tiles_and_right_edge(self):
        path = HERE / 'fixtures/visual/right-column-overflow.png'
        ref = ve.prepare([path], self.root, self.binding, phase='initial')
        b = ve.validate_bundle(ref)
        self.assertEqual(len(b['views']), 9)
        right = next(v for v in b['views'] if v['label'] == 'edge-right')
        self.assertLessEqual(right['box'][0], 881)
        self.assertGreaterEqual(right['box'][2], 977)
        self.assertLessEqual(right['box'][1], 464)
        self.assertGreaterEqual(right['box'][3], 516)

    def test_tile_coverage_including_last_pixel(self):
        w, h = 1555, 1031
        tiles = [box for name, box in ve.view_boxes(w, h) if name.startswith('tile')]
        for y in range(h):
            intervals = sorted((a, c) for a, b, c, d in tiles if b <= y < d)
            edge = 0
            for start, end in intervals:
                self.assertLessEqual(start, edge); edge = max(edge, end)
            self.assertEqual(edge, w)

    def test_missing_right_edge_is_blocked_even_with_new_hash(self):
        ref = self.revise(self.prepare(), lambda d: d['views'].pop())
        with self.assertRaisesRegex(ve.VisualError, '漏项'):
            ve.validate_bundle(ref)

    def test_crop_pixels_cannot_be_replaced_by_clean_picture(self):
        ref = self.prepare(); b = ve.read(ve.checked_file(ref)); row = b['views'][1]
        Image.new('RGB', tuple(row['size']), 'black').save(row['path'])
        row['sha256'] = ve.file_sha(row['path']); ve.write(ref['path'], b)
        with self.assertRaisesRegex(ve.VisualError, '裁片不是真实'):
            ve.validate_bundle(ve.reference(ref['path']))

    def test_corrupt_original_asset_rejected(self):
        ref = self.prepare(); b = ve.read(ve.checked_file(ref))
        Path(b['assets'][0]['path']).write_bytes(b'corrupt')
        with self.assertRaises(ve.VisualError):
            ve.validate_bundle(ref)

    def test_no_worker_no_pass(self):
        with self.assertRaisesRegex(ve.VisualError, '没有可调用'):
            ve.run(self.prepare(), None, self.root)

    def test_invalid_command_type_rejected(self):
        with self.assertRaises(ve.VisualError):
            ve.load_worker({'command': 'echo pass'})

    def test_initial_real_subprocess_receipt(self):
        result = self.validate(self.run_review())
        self.assertEqual(result['verdict'], 'pass')
        self.assertEqual(result['phase'], 'initial')

    def test_final_two_fresh_blind_requests(self):
        ref = self.run_review(final=True); data = ve.read(ve.checked_file(ref))
        self.assertEqual([c['role'] for c in data['calls']], ['final-primary', 'final-independent'])
        self.assertNotEqual(data['calls'][0]['request_id'], data['calls'][1]['request_id'])
        request = ve.read(ve.checked_file(data['calls'][1]['request']))
        self.assertNotIn('observation', request)
        self.assertNotIn('previous_result', request)
        self.assertGreater(len(request['images']), len(request['required_view_ids']))
        self.assertEqual(self.validate(ref)['verdict'], 'pass')

    def test_second_review_detecting_defect_blocks(self):
        result = self.validate(self.run_review('second-fail', final=True))
        self.assertEqual(result['verdict'], 'fail')
        self.assertEqual(result['checks']['text_inside_bounds'], 'fail')

    def test_contradictory_pass_and_finding_fails(self):
        result = self.validate(self.run_review('contradictory'))
        self.assertEqual(result['verdict'], 'fail')

    def test_uncertain_not_pass(self):
        self.assertEqual(self.validate(self.run_review('uncertain'))['verdict'], 'uncertain')

    def test_each_invalid_worker_response_blocks(self):
        for mode in ('missing', 'na', 'no-observation', 'wrong-id', 'unlocated-fail', 'invalid-json', 'exit'):
            with self.subTest(mode=mode), self.assertRaises(ve.VisualError):
                self.run_review(mode)

    def test_timeout_records_blocked_and_does_not_reuse_output(self):
        with self.assertRaises(ve.VisualError):
            ve.run(self.prepare(), self.worker('timeout', 1), self.root)
        reports = list(self.root.glob('inspection-*/inspection.json'))
        self.assertEqual(len(reports), 1)
        self.assertEqual(ve.read(reports[0])['status'], 'blocked')

    def test_missing_current_page_is_blocked(self):
        with self.assertRaises(ve.VisualError):
            ve.prepare([self.image], self.root, self.binding, phase='final', originals=[self.image])

    def test_missing_original_is_blocked(self):
        with self.assertRaises(ve.VisualError):
            ve.prepare([self.image], self.root, self.binding, phase='final', pages=self.pages)

    def test_current_page_mismatch_fails(self):
        self.assertEqual(self.validate(self.run_review('page-mismatch', final=True))['verdict'], 'fail')

    def test_semantics_mismatch_fails(self):
        self.assertEqual(self.validate(self.run_review('semantic-mismatch', final=True))['verdict'], 'fail')

    def test_classification_disagreement_fails(self):
        self.assertEqual(self.validate(self.run_review('type-disagreement', final=True))['verdict'], 'fail')

    def test_stale_candidate_binding_rejected(self):
        ref = self.run_review()
        changed = {**self.binding, 'candidate_sha256': 'new-file'}
        with self.assertRaises(ve.VisualError):
            self.validate(ref, changed)

    def test_cannot_reuse_initial_as_final(self):
        with self.assertRaises(ve.VisualError):
            self.validate(self.run_review(), phase='final')

    def test_reused_request_id_rejected(self):
        ref = self.run_review(final=True)
        ref = self.revise(ref, lambda d: d['calls'][1].update(request_id=d['calls'][0]['request_id']))
        with self.assertRaisesRegex(ve.VisualError, '复用'):
            self.validate(ref)

    def test_missing_second_pass_rejected(self):
        ref = self.revise(self.run_review(final=True), lambda d: d['calls'].pop())
        with self.assertRaises(ve.VisualError):
            self.validate(ref)

    def test_summary_only_pass_cannot_override_real_fail(self):
        ref = self.run_review('fail')
        ref = self.revise(ref, lambda d: d['calls'][0]['result'].update(verdict='pass'))
        with self.assertRaisesRegex(ve.VisualError, '真实调用结果不一致'):
            self.validate(ref)

    def test_text_only_request_with_refreshed_hash_rejected(self):
        ref = self.run_review(); data = ve.read(ve.checked_file(ref)); call = data['calls'][0]
        request_path = ve.checked_file(call['request']); request = ve.read(request_path)
        request['images'] = []; ve.write(request_path, request)
        call['request'] = ve.reference(request_path); ve.write(ref['path'], data)
        with self.assertRaisesRegex(ve.VisualError, '实际请求没有包含'):
            self.validate(ve.reference(ref['path']))

    def test_actual_docx_media_pixels_are_compared_not_just_id(self):
        from docx import Document
        import zipfile
        docx = self.root / 'real.docx'; document = Document(); document.add_picture(str(self.image)); document.save(docx)
        with zipfile.ZipFile(docx) as z:
            media = [n for n in z.namelist() if n.startswith('word/media/')]
            doc = SimpleNamespace(files={n: z.read(n) for n in media})
        obj = {'id': 'fixture-1', 'paths': media}
        b = ve.validate_bundle(self.prepare())
        fi._assert_pixels(doc, obj, b, [], docx, {})
        replacement = io.BytesIO(); Image.new('RGB', (120, 90), 'black').save(replacement, format='PNG')
        doc.files[media[0]] = replacement.getvalue()
        with self.assertRaisesRegex(ve.VisualError, '不是当前 DOCX'):
            fi._assert_pixels(doc, obj, b, [], docx, {})

    def make_discovery(self):
        proof = self.run_review('fail', final=True)
        ledger_path = self.root / 'discoveries.json'
        ve.write(ledger_path, {'version': 1, 'source_sha256': ve.file_sha(self.source), 'inspections': [proof]})
        result = self.validate(proof)
        row = {'id': 'fixture-1', 'discovered_inspections': [proof],
               'resolved_discovered_defects': [f['id'] for f in result['findings']]}
        return proof, ve.reference(ledger_path), row

    def test_late_defect_after_initial_false_pass_can_authorize_scoped_repair(self):
        _, ledger, row = self.make_discovery()
        result = fi.late_findings(self.source, row, 'original-image', 'actually-repaired-image', ledger)
        self.assertTrue(result['authorizes_change'])

    def test_same_bad_image_cannot_be_cleared_by_later_pass(self):
        _, ledger, row = self.make_discovery()
        with self.assertRaisesRegex(ve.VisualError, '真实图片未改变'):
            fi.late_findings(self.source, row, 'original-image', 'original-image', ledger)

    def test_reencoding_same_pixels_does_not_close_intrinsic_defect(self):
        _, ledger, row = self.make_discovery()
        changed_binding = {**self.binding, 'object_sha256': 'new-file-hash-same-pixels'}
        row['inspection'] = self.run_review(final=True, binding=changed_binding)
        with self.assertRaisesRegex(ve.VisualError, '真实图片未改变'):
            fi.late_findings(self.source, row, 'original-image', 'new-file-hash-same-pixels', ledger)

    def test_omitted_late_failure_cannot_be_ignored(self):
        _, ledger, row = self.make_discovery(); row['discovered_inspections'] = []
        with self.assertRaises(ve.VisualError):
            fi.late_findings(self.source, row, 'original-image', 'repaired', ledger)

    def test_unclosed_late_issue_blocks(self):
        _, ledger, row = self.make_discovery(); row['resolved_discovered_defects'] = []
        with self.assertRaises(ve.VisualError):
            fi.late_findings(self.source, row, 'original-image', 'repaired', ledger)

    def test_different_document_cannot_authorize_redesign(self):
        _, ledger, row = self.make_discovery(); self.source.write_bytes(b'different document')
        with self.assertRaises(ve.VisualError):
            fi.late_findings(self.source, row, 'original-image', 'repaired', ledger)

    def test_all_pass_receipt_is_not_discovered_defect(self):
        proof = self.run_review(final=True)
        path = self.root / 'discoveries.json'; ve.write(path, {'version': 1, 'source_sha256': ve.file_sha(self.source), 'inspections': [proof]})
        with self.assertRaises(ve.VisualError):
            fi.ledger(self.source, path)

    def test_empty_picture_is_not_a_native_shape_bypass(self):
        with self.assertRaisesRegex(ve.VisualError, '原生'):
            fi._extract(SimpleNamespace(files={}), {'id': 'fixture-1', 'paths': []}, self.root)

    def test_pipeline_still_has_all_14_gates_and_default_renderer_priority(self):
        import ast
        text = (HERE.parent / 'scripts/review_pipeline.py').read_text()
        tree = ast.parse(text)
        gates = next(ast.literal_eval(n.value) for n in tree.body if isinstance(n, ast.Assign) and any(isinstance(t, ast.Name) and t.id == 'REQUIRED_GATES' for t in n.targets))
        self.assertEqual(len(gates), 14)
        self.assertIn('image_discovery_ledger', text)
        import office_backends
        self.assertEqual(office_backends.engine_order(), ('word', 'wps', 'libreoffice'))


if __name__ == '__main__':
    unittest.main()
