import copy
import json
import unittest
from pathlib import Path
from PIL import Image
from test_numbering_policy import Fixtures, rewrite
import numbering_policy as n
import figure_review as f


class FigureGate(Fixtures):
    def data(self):
        self.sample();self.normalize()
        initial=f.make_review_template(self.source,self.root/'source-images')
        for row in initial['objects']:
            row.update(image_type='diagram',observation='Synthetic test record, not a real visual review.',checks={k:'pass' for k in f.CHECKS})
        page=self.root/'page-1.png';Image.new('RGB',(200,300),'white').save(page)
        pages=[{'path':str(page),'name':page.name,'sha256':n.file_digest(page)}]
        final={'source_sha256':n.file_digest(self.source),'candidate_sha256':n.file_digest(self.out),'objects':[]}
        for obj in f.figures(self.out):
            old=next(r for r in initial['objects'] if r['id']==obj['id'])
            final['objects'].append({'id':obj['id'],'object_sha256':obj['hash'],'source_object_sha256':old['object_sha256'],
                                     'image_type':'diagram','observation':'Synthetic final test record.',
                                     'checks':{k:'pass' for k in f.CHECKS},'style_and_semantics_preserved':True,
                                     'resolved_initial_defects':[],'pages':[{'name':page.name,'sha256':pages[0]['sha256']}]})
        return initial,final,pages

    def verify(self,initial,final,pages):return f.validate_final(self.source,self.out,initial,final,pages)

    def test_all_actual_images_have_review_entries(self):
        initial,final,pages=self.data();self.assertEqual(self.verify(initial,final,pages)['figure_count'],2)
    def test_missing_initial_image_rejected(self):
        a,b,p=self.data();a['objects'].pop()
        with self.assertRaises(n.PolicyError) as c:self.verify(a,b,p)
        self.assertEqual(c.exception.code,'image-review-missing')
    def test_missing_final_image_rejected(self):
        a,b,p=self.data();b['objects'].pop()
        with self.assertRaises(n.PolicyError):self.verify(a,b,p)
    def test_same_bad_bitmap_cannot_be_declared_fixed(self):
        a,b,p=self.data();a['objects'][0]['checks']['text_inside_bounds']='fail';b['objects'][0]['resolved_initial_defects']=['text_inside_bounds'];b['objects'][0]['repair_note']='test claims it was fixed'
        with self.assertRaises(n.PolicyError) as c:self.verify(a,b,p)
        self.assertEqual(c.exception.code,'image-defect-unchanged')
    def test_same_burned_caption_bitmap_rejected(self):
        a,b,p=self.data();a['objects'][0]['checks']['no_embedded_caption']='fail';b['objects'][0]['resolved_initial_defects']=['no_embedded_caption'];b['objects'][0]['repair_note']='only changed report'
        with self.assertRaises(n.PolicyError) as c:self.verify(a,b,p)
        self.assertEqual(c.exception.code,'image-defect-unchanged')
    def test_unresolved_initial_defect_rejected(self):
        a,b,p=self.data();a['objects'][0]['checks']['readable']='fail'
        with self.assertRaises(n.PolicyError) as c:self.verify(a,b,p)
        self.assertEqual(c.exception.code,'image-issue-not-closed')
    def test_layout_readability_can_improve_without_repainting_bitmap(self):
        a,b,p=self.data();a['objects'][0]['checks']['readable']='fail';b['objects'][0]['resolved_initial_defects']=['readable'];b['objects'][0]['repair_note']='Figure placed on its own page and enlarged.'
        self.assertEqual(self.verify(a,b,p)['status'],'passed')
    def test_final_text_overflow_is_error(self):
        a,b,p=self.data();b['objects'][0]['checks']['text_inside_bounds']='fail'
        with self.assertRaises(n.PolicyError) as c:self.verify(a,b,p)
        self.assertEqual(c.exception.code,'image-visual-defect')
    def test_architecture_cannot_escape_as_photo(self):
        a,b,p=self.data();b['objects'][0]['image_type']='photo'
        with self.assertRaises(n.PolicyError) as c:self.verify(a,b,p)
        self.assertEqual(c.exception.code,'image-type-changed')
    def test_no_arrow_architecture_can_use_na_with_reason(self):
        a,b,p=self.data()
        for row in a['objects']+b['objects']:
            row['checks']['connections_correct']='not_applicable';row['not_applicable_reasons']={'connections_correct':'This architecture uses layers, no arrows; grouping checked.'}
        self.assertEqual(self.verify(a,b,p)['status'],'passed')
    def test_na_without_reason_rejected(self):
        a,b,p=self.data();a['objects'][0]['checks']['connections_correct']='not_applicable'
        with self.assertRaises(n.PolicyError):self.verify(a,b,p)
    def test_pending_not_pass(self):
        a,b,p=self.data();b['objects'][0]['checks']['readable']='pending'
        with self.assertRaises(n.PolicyError):self.verify(a,b,p)
    def test_old_candidate_hash_rejected(self):
        a,b,p=self.data();b['candidate_sha256']='old'
        with self.assertRaises(n.PolicyError):self.verify(a,b,p)
    def test_old_source_object_hash_rejected(self):
        a,b,p=self.data();b['objects'][0]['source_object_sha256']='old'
        with self.assertRaises(n.PolicyError):self.verify(a,b,p)
    def test_old_page_hash_rejected(self):
        a,b,p=self.data();b['objects'][0]['pages'][0]['sha256']='old'
        with self.assertRaises(n.PolicyError):self.verify(a,b,p)
    def test_no_final_page_is_not_enough_to_review_standalone_png(self):
        a,b,p=self.data();b['objects'][0]['pages']=[]
        with self.assertRaises(n.PolicyError):self.verify(a,b,p)
    def test_source_media_export_excludes_external_word_caption(self):
        a,b,p=self.data()
        for row in a['objects']:
            self.assertEqual(len(row['source_files']),1)
            self.assertEqual(Path(row['source_files'][0]['path']).read_bytes(),self.image.read_bytes())
    def test_unrequested_repaint_rejected(self):
        a,b,p=self.data()
        new=self.root/'new.png';Image.new('RGB',(100,60),'black').save(new)
        paths=f.figures(self.out)[0]['paths']
        def patch(data):
            for path in paths:data[path]=new.read_bytes()
        rewrite(self.out,patch);b['candidate_sha256']=n.file_digest(self.out)
        for row,obj in zip(b['objects'],f.figures(self.out)):row['object_sha256']=obj['hash']
        with self.assertRaises(n.PolicyError) as c:self.verify(a,b,p)
        self.assertEqual(c.exception.code,'image-unrequested-redesign')


def load_tests(loader, tests, pattern):
    return unittest.TestSuite(FigureGate(name) for name in FigureGate.__dict__ if name.startswith('test_'))
