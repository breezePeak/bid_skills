"""Bound cross-command retries by semantic input, not candidate filenames/revisions."""
from __future__ import annotations
import copy
import hashlib
import json
from pathlib import Path
from block_progress import package, parse, signature, Q, BlockError


def key_for(state, scripts):
    from office_stability import snapshot
    parts,_,_=package(state['current'])
    value=[snapshot(state['current'])]
    # Preserve true table/shape geometry and field instructions as retry inputs;
    # omit raw style serialization, save IDs and display-only field caches.
    for name in sorted(parts):
        if name.startswith('word/media/'):
            value.append(hashlib.sha256(parts[name]).hexdigest())
        elif name.startswith('word/') and name.endswith('.xml') and Path(name).stem.startswith(('document','header','footer','footnotes','endnotes')):
            root=parse(parts[name])
            selected=[]
            for n in root.iter():
                if n.tag==Q('fldSimple'):
                    selected.append((n.tag,sorted(n.attrib.items())))
                elif n.tag in {Q('tblGrid'),Q('tcPr'),Q('sectPr'),Q('instrText'),Q('drawing'),Q('pict')}:
                    selected.append(signature(n))
            value.append((name,selected))
    settings={k:state['settings'].get(k) for k in ('renderer','field_engine','uno_python','baselines')}
    settings['runtime']=(state.get('runtime_preflight') or {}).get('field_engine')
    digest=hashlib.sha256(json.dumps([value,settings],ensure_ascii=False,sort_keys=True).encode())
    for p in sorted(Path(scripts).iterdir()):
        if p.is_file() and p.suffix in {'.py','.ps1'}:
            digest.update(p.name.encode());digest.update(p.read_bytes())
    return digest.hexdigest()


def failure_key(issues):
    def stable(value):
        if isinstance(value,dict):
            return {k:stable(v) for k,v in value.items() if k not in {'path','report','detail','stdout','stderr','candidate_sha256','source_sha256','timestamp','revision'}}
        if isinstance(value,list):return [stable(v) for v in value]
        if isinstance(value,str):
            import re
            return re.sub(r'(?:[A-Za-z]:[\\/]|/)[^\s]+','<path>',value)
        return value
    rows=[stable(i) for i in issues]
    return hashlib.sha256(json.dumps(rows,sort_keys=True,ensure_ascii=False).encode()).hexdigest()


def before_retry(state, scripts):
    key=key_for(state,scripts)
    history=state.get('failure_history',[])
    same=[r for r in history if r['input_key']==key]
    if same and same[-1].get('repeated',0)>=2:
        raise BlockError('相同输入和执行条件已连续两次得到相同错误；禁止再跑全文终检。读取 failure_history 定位原因，修正输入、共享规则或真实执行条件后继续。')
    return key


def record_failure(state, key, issues):
    history=state.setdefault('failure_history',[])
    problem=failure_key(issues)
    prior=next((r for r in reversed(history) if r['input_key']==key),None)
    repeated=prior.get('repeated',0)+1 if prior and prior['failure_key']==problem else 1
    history.append({'input_key':key,'failure_key':problem,'repeated':repeated,
                    'issues':copy.deepcopy(issues),'status':'requires_cause_fix' if repeated>=2 else 'failed'})
    # Preserve distinct diagnostic failures; only cap duplicate history volume.
    if len(history)>40:del history[:-40]
    return history[-1]


def record_success(state):
    state['failure_history']=[]
