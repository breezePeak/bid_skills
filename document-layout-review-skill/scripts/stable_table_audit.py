#!/usr/bin/env python3
"""Table appearance audit retaining real property differences, ignoring style UI metadata."""
from __future__ import annotations
import argparse,json,copy
from pathlib import Path
from format_equivalence import canonical


def audit(source,template,style_json=None):
    import template_table_style as tables
    from numbering_policy import Doc,Q,PolicyError,file_digest
    doc=Doc(Path(source));standard=tables.TemplateTables(template,style_json)
    def project(d):
        rows=tables.project(d);colors=tables.theme_colors(d)
        for row,table in zip(rows,d.document.iter(Q('tbl'))):
            sid=tables.table_sid(d,table)
            flat=tables.style_flat(d,sid,colors)
            # Parent inheritance is resolved by the original table resolver.
            # Only style metadata is removed; all property/conditional regions remain.
            for child in list(flat):
                if child.tag not in {Q('pPr'),Q('rPr'),Q('tblPr'),Q('trPr'),Q('tcPr'),Q('tblStylePr')}:flat.remove(child)
            flat.attrib.clear()
            row['style']=canonical(flat)
        return rows
    before=project(doc);records=tables.apply(doc,standard);after=project(doc)
    issues=[{'severity':'error','code':'table-template-appearance-mismatch','table':a['table'],
             'message':'表格真实外观与模板不一致；等效样式元数据已排除。'} for a,b in zip(before,after) if a!=b]
    return {'status':'failed' if issues else 'passed','input':str(source),'input_sha256':file_digest(source),
        'template':str(template),'template_sha256':file_digest(template),'template_style_sha256':standard.profile_hash,
        'tables':records,'issues':issues,'visual_review_required':True}


def main():
    p=argparse.ArgumentParser();p.add_argument('action',choices=['audit']);p.add_argument('input',type=Path)
    p.add_argument('--template',type=Path,required=True);p.add_argument('--template-style-json',type=Path);p.add_argument('--json-out',type=Path)
    a=p.parse_args()
    try:result=audit(a.input,a.template,a.template_style_json)
    except Exception as exc:result={'status':'failed','issues':[{'code':'table-equivalence-error','message':str(exc)}]}
    text=json.dumps(result,ensure_ascii=False,indent=2)
    if a.json_out:
        if a.json_out.resolve() in {a.input.resolve(),a.template.resolve()}:raise SystemExit('报告不能覆盖文档')
        a.json_out.parent.mkdir(parents=True,exist_ok=True);a.json_out.write_text(text+'\n',encoding='utf-8')
    print(text);return 0 if result['status']=='passed' else 2


if __name__=='__main__':raise SystemExit(main())
