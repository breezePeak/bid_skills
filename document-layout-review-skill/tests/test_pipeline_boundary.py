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
        with patch('runtime_preflight.check', return_value={'status':'passed'}),patch.object(sys,'argv',['review_pipeline.py',*map(str,args)]),contextlib.redirect_stdout(stream):
            rc=pipeline.main()
        return rc,json.loads(stream.getvalue())
    def test_startup_prepares_blocks_without_global_image_calls(self):
        self.sample();before=self.source.read_bytes();work=self.root/'work'
        code,result=self.invoke(self.source,'--work-dir',work)
        self.assertEqual(code,0);self.assertEqual(result['status'],'global_setup')
        self.assertFalse((work/'candidate.docx').exists());self.assertEqual(before,self.source.read_bytes())
        state=json.loads((work/'block-state.json').read_text());review=json.loads(Path(state['initial_review']).read_text())
        self.assertEqual(len(review['objects']),2)
        self.assertTrue(all('pending' in x['checks'].values() for x in review['objects']))
    def test_legacy_candidate_filename_is_preserved_as_input(self):
        self.sample();work=self.root/'work';work.mkdir();source=work/'candidate.docx';source.write_bytes(self.source.read_bytes());before=source.read_bytes()
        code,result=self.invoke(source,'--work-dir',work)
        self.assertEqual(code,0);self.assertEqual(result['status'],'global_setup')
        self.assertEqual(before,source.read_bytes())
    def test_audit_only_cannot_skip_unfinished_blocks(self):
        self.sample();code,result=self.invoke(self.source,'--work-dir',self.root/'work','--audit-only')
        self.assertEqual(code,2);self.assertEqual(result['status'],'blocked')


def load_tests(loader, tests, pattern):
    return unittest.TestSuite(PipelineBoundary(name) for name in PipelineBoundary.__dict__ if name.startswith('test_'))

# DLR_BLOCK_BOUNDARY_TESTS
