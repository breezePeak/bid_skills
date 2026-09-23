"""Orchestrator boundary tests; these do not substitute for a full Word run."""
import contextlib
import io
import json
import sys
import unittest
from pathlib import Path
from unittest.mock import patch
from test_numbering_policy import Fixtures
import review_pipeline as pipeline

class PipelineBoundary(Fixtures):
    def invoke(self,*args):
        stream=io.StringIO()
        with patch.object(sys,'argv',['review_pipeline.py',*map(str,args)]),contextlib.redirect_stdout(stream):
            rc=pipeline.main()
        return rc,json.loads(stream.getvalue())
    def test_actual_image_review_required_before_legacy_mutation(self):
        self.sample();before=self.source.read_bytes();work=self.root/'work'
        code,result=self.invoke(self.source,'--work-dir',work)
        self.assertEqual(code,4);self.assertEqual(result['status'],'requires_image_review')
        self.assertFalse((work/'candidate.docx').exists());self.assertEqual(before,self.source.read_bytes())
        review=json.loads(Path(result['review']).read_text())
        self.assertEqual(len(review['objects']),2)
        self.assertTrue(all('pending' in x['checks'].values() for x in review['objects']))
    def test_input_in_reserved_stage_path_is_not_overwritten(self):
        self.sample();work=self.root/'work';work.mkdir();source=work/'candidate.docx';source.write_bytes(self.source.read_bytes());before=source.read_bytes()
        code,result=self.invoke(source,'--work-dir',work)
        self.assertEqual(code,2);self.assertEqual(result['issues'][0]['code'],'input-output-collision')
        self.assertEqual(before,source.read_bytes())
    def test_audit_only_requires_real_original_file(self):
        self.sample();code,result=self.invoke(self.source,'--work-dir',self.root/'work','--audit-only')
        self.assertEqual(code,2);self.assertEqual(result['issues'][0]['code'],'input-missing')


def load_tests(loader, tests, pattern):
    return unittest.TestSuite(PipelineBoundary(name) for name in PipelineBoundary.__dict__ if name.startswith('test_'))
