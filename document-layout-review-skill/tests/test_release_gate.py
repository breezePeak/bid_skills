import copy
import json
import unittest
from pathlib import Path
from unittest.mock import patch
from test_figure_gate import FigureGate
import numbering_policy as n
import figure_review as f
import finalize_review as g
import field_refresh as r
from review_pipeline import REQUIRED_GATES, write_json


class ReleaseGate(FigureGate):
    def bundle(self):
        a,b,pages=self.data();initial=self.root/'initial.json';write_json(initial,a)
        style=self.root/'template-style.json';write_json(style,{})
        gates=[]
        for ident in sorted(REQUIRED_GATES):
            report=self.root/(ident+'.json');result={'status':'passed','issues':[]}
            # Synthetic gate metadata fixtures; no claim that Word executed here.
            if ident=='field-update':result={'status':'passed','engine':'Microsoft Word','fields_updated':True,'indexes_updated':True,'errors':[],'output_sha256':n.file_digest(self.out)}
            write_json(report,result);gates.append({'id':ident,'passed':True,'report':str(report),'report_sha256':n.file_digest(report)})
        m={'version':4,'status':'awaiting_visual_review','candidate':str(self.out),'candidate_sha256':n.file_digest(self.out),
           'source':str(self.source),'source_sha256':n.file_digest(self.source),'template':str(self.template),'template_sha256':n.file_digest(self.template),
           'template_style_json':str(style),'template_style_sha256':n.file_digest(style),'initial_visual_review':str(initial),'initial_visual_review_sha256':n.file_digest(initial),
           'gates':gates,'render':{'passed':True,'pages':pages}}
        v={'overall_status':'pass','candidate_sha256':n.file_digest(self.out),'pages':[{'name':p['name'],'sha256':p['sha256'],'status':'pass','observations':'Synthetic test page review.'} for p in pages],'figures':b,'table_reviews':[]}
        return m,v
    def test_complete_contract_passes(self):
        m,v=self.bundle();self.assertEqual(g.validate_release(m,v)['status'],'passed')
    def test_missing_gate_cannot_vacuously_pass(self):
        m,v=self.bundle();m['gates']=[]
        with self.assertRaises(n.PolicyError) as c:g.validate_release(m,v)
        self.assertEqual(c.exception.code,'required-gate-missing')
    def test_missing_image_gate_rejected(self):
        m,v=self.bundle();m['gates']=[x for x in m['gates'] if x['id']!='image-inventory']
        with self.assertRaises(n.PolicyError):g.validate_release(m,v)
    def test_duplicate_gate_rejected(self):
        m,v=self.bundle();m['gates'][-1]=m['gates'][0]
        with self.assertRaises(n.PolicyError):g.validate_release(m,v)
    def test_empty_pages_rejected(self):
        m,v=self.bundle();m['render']['pages']=[];v['pages']=[]
        with self.assertRaises(n.PolicyError):g.validate_release(m,v)
    def test_duplicate_page_review_rejected(self):
        m,v=self.bundle();v['pages']*=2
        with self.assertRaises(n.PolicyError):g.validate_release(m,v)
    def test_missing_figure_entries_rejected(self):
        m,v=self.bundle();v['figures']['objects']=[]
        with self.assertRaises(n.PolicyError):g.validate_release(m,v)
    def test_mutated_audit_json_rejected(self):
        m,v=self.bundle();Path(m['gates'][0]['report']).write_text('{}')
        with self.assertRaises(n.PolicyError):g.validate_release(m,v)
    def test_changed_template_rejected(self):
        m,v=self.bundle();m['template_sha256']='old'
        with self.assertRaises(n.PolicyError):g.validate_release(m,v)
    def test_no_observation_rejected(self):
        m,v=self.bundle();v['pages'][0]['observations']=''
        with self.assertRaises(n.PolicyError):g.validate_release(m,v)
    def test_failed_release_preserves_existing_deliverable(self):
        m,v=self.bundle();v['overall_status']='fail';final=self.root/'final.docx';final.write_bytes(b'old')
        with self.assertRaises(n.PolicyError):g.finalize(m,v,final)
        self.assertEqual(final.read_bytes(),b'old')
    def test_release_copies_exact_reviewed_bytes(self):
        m,v=self.bundle();final=self.root/'final.docx';result=g.finalize(m,v,final)
        self.assertEqual(n.file_digest(final),m['candidate_sha256'])
    def test_updatefields_flag_is_not_engine_evidence(self):
        m,v=self.bundle()
        with self.assertRaises(n.PolicyError):r.validate_report(self.out,{'status':'passed','updateFields':True,'output_sha256':n.file_digest(self.out)})
    def test_field_refresh_after_modification_is_stale(self):
        m,v=self.bundle()
        with self.assertRaises(n.PolicyError):r.validate_report(self.out,{'status':'passed','engine':'Microsoft Word','fields_updated':True,'indexes_updated':True,'errors':[],'output_sha256':'old'})
    def test_table_review_unresolved_is_blocked(self):
        m,v=self.bundle();gate=next(x for x in m['gates'] if x['id']=='table-layout')
        write_json(gate['report'],{'issues':[{'severity':'review','code':'merge-candidate','table':1}]});gate['report_sha256']=n.file_digest(Path(gate['report']))
        with self.assertRaises(n.PolicyError) as c:g.validate_release(m,v)
        self.assertEqual(c.exception.code,'table-review-missing')

    def test_release_cannot_overwrite_original_source(self):
        m,v=self.bundle();before=self.source.read_bytes()
        with self.assertRaises(n.PolicyError):g.finalize(m,v,self.source)
        self.assertEqual(self.source.read_bytes(),before)


def load_tests(loader, tests, pattern):
    return unittest.TestSuite(ReleaseGate(name) for name in ReleaseGate.__dict__ if name.startswith('test_'))
