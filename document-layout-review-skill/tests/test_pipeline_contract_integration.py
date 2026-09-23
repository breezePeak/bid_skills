"""Integration of NEW gates using real DOCX data.

Unchanged legacy checks, image observation and Office execution are test doubles;
these tests are not a claim of full-pipeline or Windows Word compatibility.
"""
from __future__ import annotations
import hashlib
import importlib.util
import io
import json
import sys
import types
import unittest
from contextlib import redirect_stdout
from pathlib import Path
from unittest.mock import patch

from test_template_contract_regression import TemporaryTest,SCRIPTS,content,fmt,env,Document,property_node
import template_caption_policy as captions


def file_hash(path):return 'sha256:'+hashlib.sha256(Path(path).read_bytes()).hexdigest()
def data_hash(data):return 'sha256:'+hashlib.sha256(data).hexdigest()


class TestError(ValueError):
    def __init__(self,code,message,**details):super().__init__(message);self.code=code;self.details=details
    def as_issue(self):return {'severity':'error','code':self.code,'message':str(self),**self.details}


def module(name,**members):
    m=types.ModuleType(name)
    m.__dict__.update(members)
    return m


def load(name):
    spec=importlib.util.spec_from_file_location(name,SCRIPTS/(name+'.py'))
    m=importlib.util.module_from_spec(spec);sys.modules[name]=m;spec.loader.exec_module(m);return m


class GateIntegrationTests(TemporaryTest):
    def setUp(self):
        super().setUp()
        self.doubles={
            'numbering_policy':module('numbering_policy',PolicyError=TestError,file_digest=file_hash,digest=data_hash,
                 Doc=lambda p:p,inventory=lambda d:[],public_inventory=lambda d:[],
                 load_object_plan=lambda *a,**k:{},audit_document=lambda *a,**k:{'status':'passed','issues':[]}),
            'figure_review':module('figure_review',make_review_template=lambda *a:{'objects':[]},validate_initial=lambda *a:{},validate_final=lambda *a:{'status':'passed'}),
            'field_refresh':module('field_refresh',refresh=lambda *a,**k:{'status':'passed'},validate_report=lambda *a:{'status':'passed'},safe_fields=lambda *a:None),
            'template_table_style':module('template_table_style',audit=lambda *a:{'status':'passed','issues':[]})}
        ctx=patch.dict(sys.modules,self.doubles);ctx.start();self.addCleanup(ctx.stop)
        self.pipeline=load('review_pipeline');self.release=load('finalize_review')

    def save_json(self,path,data):path.write_text(json.dumps(data,ensure_ascii=False),encoding='utf-8');return path

    def manifest(self,changed=False):
        source=self.document('original.docx','工期30天')
        candidate=self.document('candidate.docx','工期3天' if changed else '工期30天')
        template=self.document('template.docx',None)
        p=self.root/'style.json';self.save_json(p,{})
        initial=self.save_json(self.root/'initial.json',{})
        plan=self.save_json(self.root/'content-plan.json',content.make_plan(source))
        rules=self.save_json(self.root/'rules.json',{})
        pre=self.save_json(self.root/'preflight.json',{'status':'passed'})
        gates=[]
        for name in sorted(self.pipeline.REQUIRED_GATES):
            path=self.save_json(self.root/(name+'.json'),{'status':'passed','issues':[]})
            gates.append({'id':name,'passed':True,'report':str(path),'report_sha256':file_hash(path)})
        # This is deliberately only a page-record fixture, NOT visual verification.
        page=self.root/'page-1.png';page.write_bytes(b'page-record-fixture')
        m={'version':5,'status':'awaiting_visual_review','source':str(source),'source_sha256':file_hash(source),
           'candidate':str(candidate),'candidate_sha256':file_hash(candidate),'template':str(template),'template_sha256':file_hash(template),
           'template_style_json':str(p),'template_style_sha256':file_hash(p),'initial_visual_review':str(initial),'initial_visual_review_sha256':file_hash(initial),
           'content_plan':str(plan),'content_plan_sha256':file_hash(plan),'text_rules':str(rules),'text_rules_file_sha256':file_hash(rules),
           'runtime_preflight':str(pre),'runtime_preflight_sha256':file_hash(pre),'gates':gates,
           'render':{'passed':True,'pages':[{'name':page.name,'path':str(page),'sha256':file_hash(page)}]}}
        v={'candidate_sha256':file_hash(candidate),'overall_status':'pass','pages':[{'name':page.name,'sha256':file_hash(page),'status':'pass','observations':'test record only'}],
           'figures':{},'table_reviews':[]}
        return m,v

    def test_all_three_new_gates_are_mandatory(self):
        self.assertTrue({'content-integrity','table-layout-template','template-effective-format'}<=self.pipeline.REQUIRED_GATES)

    def test_old_manifest_cannot_bypass_new_checks(self):
        m,v=self.manifest();m['version']=4
        with self.assertRaises(TestError) as error:self.release.validate_release(m,v)
        self.assertEqual(error.exception.code,'manifest-not-ready')

    def test_missing_content_gate_prevents_release(self):
        m,v=self.manifest();m['gates']=[g for g in m['gates'] if g['id']!='content-integrity']
        with self.assertRaises(TestError) as error:self.release.validate_release(m,v)
        self.assertEqual(error.exception.code,'required-gate-missing')

    def test_fabricated_pass_reports_do_not_hide_changed_business_number(self):
        m,v=self.manifest(changed=True);out=self.root/'final.docx'
        with self.assertRaises(TestError) as error:self.release.finalize(m,v,out)
        self.assertEqual(error.exception.code,'final-content-integrity-failed')
        self.assertFalse(out.exists())

    def test_release_rechecks_real_current_content(self):
        m,v=self.manifest();result=self.release.validate_release(m,v)
        self.assertEqual(result['final_rechecks']['content-integrity']['status'],'passed')
        self.assertEqual(result['final_rechecks']['template-effective-format']['status'],'passed')

    def test_changed_content_plan_invalidates_manifest(self):
        m,v=self.manifest();Path(m['content_plan']).write_text('{}')
        with self.assertRaises(TestError) as error:self.release.validate_release(m,v)
        self.assertEqual(error.exception.code,'baseline-stale')

    def test_failed_preflight_occurs_before_image_preparation_and_repairs(self):
        source=self.document();template=self.document('template.docx',None);profile=self.save_json(self.root/'style.json',{})
        argv=['review_pipeline.py',str(source),'--template',str(template),'--template-style-json',str(profile),'--work-dir',str(self.root/'work')]
        original=source.read_bytes()
        with patch.object(sys,'argv',argv),patch.object(self.pipeline,'runtime_check',side_effect=env.PreflightError('word-engine-unavailable','test')),patch.object(self.pipeline,'make_review_template') as images,patch.object(self.pipeline,'run_script') as repairs,redirect_stdout(io.StringIO()):
            self.assertEqual(self.pipeline.main(),2)
            images.assert_not_called();repairs.assert_not_called()
        self.assertEqual(source.read_bytes(),original)
        self.assertFalse((self.root/'work'/'candidate.docx').exists())

    def test_audit_only_gate_compares_original_not_candidate_to_itself(self):
        m,v=self.manifest(changed=True);reports=self.root/'gates';reports.mkdir()
        original_runner=self.pipeline.run_script
        real={'content_integrity.py','template_layout_contract.py','template_format_contract.py'}
        def runner(name,*args,**kw):
            if name in real:return original_runner(name,*args,**kw)
            target=Path(args[args.index('--json-out')+1])
            self.save_json(target,{'status':'passed','issues':[],'text_rules_sha256':'test','issue_count':0})
            return {'ok':True,'returncode':0,'stdout':'','stderr':''}
        with patch.object(self.pipeline,'run_script',runner):
            gates=self.pipeline.check_gates(Path(m['candidate']),Path(m['template']),Path(m['template_style_json']),Path(m['text_rules']),'test',reports,None,Path(m['initial_visual_review']),Path(m['source']),{},Path(m['content_plan']))
        checks={g['id']:g for g in gates}
        self.assertFalse(checks['content-integrity']['passed'])
        self.assertTrue(checks['template-effective-format']['passed'])
        self.assertTrue(checks['table-layout-template']['passed'])
        self.assertEqual(set(checks),self.pipeline.REQUIRED_GATES)


class CaptionContractTests(TemporaryTest):
    def template(self,position):
        d=Document();p=d.add_paragraph('表1 清单',style='Caption');t=d.add_table(rows=1,cols=1);t.cell(0,0).text='内容'
        if position=='after':t._tbl.addnext(p._p)
        path=self.root/(position+'.docx');d.save(path);return path

    def test_table_caption_below_template_is_still_forced_above(self):
        self.assertEqual(captions.placement_rules(self.template('after'),required={'table'}),{'table':'before'})

    def test_template_table_caption_above_is_preserved(self):
        self.assertEqual(captions.placement_rules(self.template('before'),required={'table'}),{'table':'before'})

    def test_missing_user_template_placement_is_not_inferred_from_broken_input(self):
        source=self.document('empty-template.docx',None)
        self.assertEqual(captions.placement_rules(source,required={'table'}),{'table':'before'})

    def test_deprecated_template_placement_cannot_reverse_baseline(self):
        source=self.document('empty-template.docx',None);profile=self.root/'profile.json'
        profile.write_text(json.dumps({'caption_placement':{'table':'after'}}))
        self.assertEqual(captions.placement_rules(source,profile,{'table'}),{'table':'before'})

    def test_deprecated_position_value_has_no_authority(self):
        source=self.document('empty-template.docx',None);profile=self.root/'profile.json'
        profile.write_text(json.dumps({'caption_placement':{'table':'guess'}}))
        self.assertEqual(captions.placement_rules(source,profile,{'table'}),{'table':'before'})

    def test_all_core_numbering_errors_including_position_remain_blocking(self):
        mockcore=module('numbering_policy',Doc=lambda x:x,audit_document=lambda *a:{'status':'failed','issues':[
            {'code':'caption-position','severity':'error'},{'code':'heading-not-automatic','severity':'error'}]})
        with patch.dict(sys.modules,{'numbering_policy':mockcore}),patch.object(captions,'placement_rules',return_value={'figure':'after','table':'before'}),patch.object(captions,'audit_caption_fields',return_value={'status':'passed','issues':[]}):
            result=captions.audit_document('candidate','template')
        self.assertEqual(result['status'],'failed')
        self.assertEqual([i['code'] for i in result['issues']],['caption-position','heading-not-automatic'])

    def test_generated_heading_alias_maps_only_to_template_role(self):
        d=Document();p=self.root/'styles.docx';d.save(p)
        from test_template_contract_regression import bytes_part
        resolver=fmt.StyleResolver(bytes_part(p,'word/styles.xml'))
        self.assertEqual(fmt.expected_style_id('DLR_heading1',resolver,{'heading1':'Heading1'}),'Heading1')
        self.assertIsNone(fmt.expected_style_id('SourceRandomStyle',resolver))


if __name__=='__main__':unittest.main()
