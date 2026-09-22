#!/usr/bin/env python3
"""检测用户上传模板中需要用户明确选择的样式冲突。

输出若为 requires_user_choice，调用方必须弹框；禁止模型自行采用 recommended。
"""
from __future__ import annotations
import argparse, json, subprocess, sys, tempfile, zipfile, re
from pathlib import Path
from xml.etree import ElementTree as ET

HERE = Path(__file__).resolve().parent
W="http://schemas.openxmlformats.org/wordprocessingml/2006/main"; NS={"w":W}; Q=lambda n:f"{{{W}}}{n}"


def load_profile(template: Path) -> dict:
    tmp = Path(tempfile.mkstemp(suffix='.json')[1])
    try:
        cp=subprocess.run([sys.executable,str(HERE/'template_style_profile.py'),str(template),'--out',str(tmp)],capture_output=True,text=True)
        if cp.returncode != 0: raise RuntimeError(cp.stderr or cp.stdout)
        return json.loads(tmp.read_text(encoding='utf-8'))
    finally:
        tmp.unlink(missing_ok=True)


def pstyle(profile, sid): return profile.get('styles',{}).get(str(sid),{})

def sig(item): return item.get('signature')

def main():
    ap=argparse.ArgumentParser(description='检测模板样式冲突并生成弹框数据')
    ap.add_argument('template',type=Path); ap.add_argument('--json-out',type=Path)
    a=ap.parse_args(); profile=load_profile(a.template); conflicts=[]
    body_id=str(profile.get('semantic_roles',{}).get('body',{}).get('style_id') or '')
    body=pstyle(profile,body_id); ind=body.get('paragraph',{}).get('indent',{})
    if ind.get('firstLine') and ind.get('firstLineChars'):
        conflicts.append({
          'conflict_id':'body-first-line-indent-dual-definition','scope':'body',
          'description':'正文样式同时定义物理首行缩进和字符首行缩进，两套基准并存。',
          'current_values':{'firstLine':ind.get('firstLine'),'firstLineChars':ind.get('firstLineChars')},
          'options':[{'id':'use_chars','label':'以字符缩进为准'},{'id':'use_physical','label':'以物理长度为准'},{'id':'preserve_both','label':'保持模板原状'}],
          'recommended':'use_chars'})
    # 题注候选冲突。
    cap=profile.get('semantic_roles',{}).get('caption',{})
    cids=[str(x) for x in cap.get('style_ids',[]) if x]
    if len(cids)>1 and len({sig(pstyle(profile,x)) for x in cids})>1:
        conflicts.append({'conflict_id':'caption-style-duplicate','scope':'caption','description':'模板存在多套定义不同的题注样式。',
          'current_values':[{'style_id':x,'name':pstyle(profile,x).get('name')} for x in cids],
          'options':[{'id':x,'label':f"使用 {pstyle(profile,x).get('name') or x}（style {x}）"} for x in cids],
          'recommended':cids[0]})
    # 多套表格文字样式。
    candidates=[]
    for sid,item in profile.get('styles',{}).items():
        name=str(item.get('name') or '')
        if '表格文字' in name and item.get('type')=='paragraph': candidates.append(str(sid))
    if len(candidates)>1 and len({sig(pstyle(profile,x)) for x in candidates})>1:
        actual=str(profile.get('semantic_roles',{}).get('table_text',{}).get('style_id') or '')
        conflicts.append({'conflict_id':'table-text-style-duplicate','scope':'table_text','description':'模板存在多套定义不同的表格文字样式。',
          'current_values':[{'style_id':x,'name':pstyle(profile,x).get('name'),'actual_usage':x==actual} for x in candidates],
          'options':[{'id':x,'label':f"使用 {pstyle(profile,x).get('name') or x}（style {x}）"} for x in candidates],
          'recommended':actual or candidates[0]})
    # 标题样式明显缺项/继承风险。
    h5=profile.get('semantic_roles',{}).get('heading5',{}).get('style_id'); h5s=pstyle(profile,h5)
    if h5s and h5s.get('based_on')==body_id:
        hi=h5s.get('paragraph',{}).get('indent',{}); bi=body.get('paragraph',{}).get('indent',{})
        if hi.get('firstLineChars') in {'0',0} and bi.get('firstLine') not in {None,'0',0} and hi.get('firstLine') is None:
            conflicts.append({'conflict_id':'heading5-inherits-normal-indent','scope':'heading5','description':'五级标题可能继续继承正文的物理首行缩进。',
             'options':[{'id':'force_zero_indent','label':'明确把两种首行缩进都设为 0'},{'id':'preserve','label':'保持模板原状'}], 'recommended':'force_zero_indent'})
    h7=profile.get('semantic_roles',{}).get('heading7',{}).get('style_id'); h7s=pstyle(profile,h7)
    if h7s and not h7s.get('paragraph',{}).get('spacing',{}).get('line'):
        conflicts.append({'conflict_id':'heading7-line-spacing-missing','scope':'heading7','description':'七级标题没有显式行距，与其他标题层级不一致。',
          'options':[{'id':'set_1_5','label':'设置为 1.5 倍行距'},{'id':'preserve','label':'保持模板原状'}], 'recommended':'set_1_5'})
    # 技术偏离表物理列与表头分组不一致时视为语义歧义，要求用户确认其用途。
    with zipfile.ZipFile(a.template) as z:
        root=ET.fromstring(z.read('word/document.xml'))
    for tbl in root.findall('.//w:tbl',NS):
        grid=len(tbl.findall('w:tblGrid/w:gridCol',NS)); rows=tbl.findall('w:tr',NS)
        if not rows: continue
        texts=[''.join(t.text or '' for t in tc.findall('.//w:t',NS)).replace('\u200b','').strip() for tc in rows[0].findall('w:tc',NS)]
        if all(x in ''.join(texts) for x in ['序号','标的名称','招标技术要求','投标响应内容','偏离程度','备注']) and grid!=len(texts):
            conflicts.append({'conflict_id':'technical-deviation-table-header-vs-physical-columns','scope':'technical_deviation_table',
              'description':f'技术偏离表表头显示 {len(texts)} 个分组，但底层有 {grid} 个物理列，需要确认额外列的业务含义。',
              'current_values':{'header_groups':texts,'physical_grid_columns':grid},
              'options':[{'id':'preserve_physical_columns','label':'保留物理列，并由用户说明额外列用途'},{'id':'normalize_to_header_groups','label':'合并为表头分组数量'}],
              'recommended':'preserve_physical_columns'})
            break
    status='requires_user_choice' if conflicts else 'no_conflicts'
    result={'status':status,'template':str(a.template),'template_profile':profile,'conflicts':conflicts,
      'ui':{'type':'modal','title':'模板样式存在冲突，请确认','blocking':True,'allow_continue_without_selection':False}}
    text=json.dumps(result,ensure_ascii=False,indent=2)
    if a.json_out:
        a.json_out.parent.mkdir(parents=True,exist_ok=True); a.json_out.write_text(text+'\n',encoding='utf-8')
    print(text); return 4 if conflicts else 0
if __name__=='__main__': raise SystemExit(main())
