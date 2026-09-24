"""Office round-trip diagnostics and cleanup: preserve outputs, never a false PASS.

Serialization differences are compared through the shared effective-style resolver.
True kerning/widow-control changes are NOT ignored. No renderer is switched to hide
format failures. A representative probe is advisory coverage, not full acceptance.
"""
from __future__ import annotations
import copy
import hashlib
import json
import shutil
import tempfile
import zipfile
from contextlib import contextmanager
from pathlib import Path
from lxml import etree as E

from block_progress import package, package_session, parse, Q, DOC, NS, sha, write_json
from template_format_contract import (StyleResolver, paragraph_style, own_runs, value_key,
                                      EXTRA_RUN, EXTRA_PARAGRAPH)
from block_checks import resolved_fonts
import content_integrity as content

RUN = EXTRA_RUN | {'sz','szCs','b','bCs','i','iCs','vertAlign','vanish'}
PARAGRAPH = EXTRA_PARAGRAPH | {'spacing','ind','jc','keepNext','keepLines','pageBreakBefore','outlineLvl'}


@contextmanager
def office_workspace(parent, warnings):
    """Only the operation's private directory is eligible for cleanup."""
    work = Path(tempfile.mkdtemp(prefix='dlr-office-',dir=parent))
    try:
        yield str(work)
    finally:
        try:
            shutil.rmtree(work)
        except OSError as exc:
            warnings.append({'code':'office-cleanup-incomplete','path':str(work),
                             'message':str(exc),'action':'cleanup-only; do not repeat document processing'})


def complete_output(path):
    """Require a stable, readable ZIP before accepting a cleanup-only warning."""
    path = Path(path)
    first = path.stat()
    with zipfile.ZipFile(path) as archive:
        if archive.testzip() or len(archive.namelist()) != len(set(archive.namelist())):
            raise ValueError('Office 输出包损坏或包含重复部件。')
        for name in ('[Content_Types].xml','word/document.xml','word/styles.xml'):
            parse(archive.read(name))
    second = path.stat()
    if (first.st_size,first.st_mtime_ns)!=(second.st_size,second.st_mtime_ns):
        raise ValueError('Office 输出仍在变化，不得作为成功结果。')
    return {'sha256':sha(path),'size':second.st_size}


def prop_values(props, names):
    return {n:value_key(props,n) for n in sorted(names)}


def snapshot(path):
    parts, _, _ = package(path)
    resolver = StyleResolver(parts['word/styles.xml'])
    result = {}
    for name in sorted(parts):
        if not name.startswith('word/') or not name.endswith('.xml') or not Path(name).stem.startswith(('document','header','footer','footnotes','endnotes')):
            continue
        root = parse(parts[name]); rows=[]
        for rec in content.visible_records(root):
            p=rec['node']; text=''.join(rec['text'])
            if not text.strip():
                continue  # field-result-only paragraphs are audited by field/TOC checks
            sid=paragraph_style(p,resolver)
            pr=resolver.properties('pPr',sid,p.find(Q('pPr')))
            spans=[]
            for r in own_runs(p):
                if r.xpath('ancestor::w:fldSimple',namespaces=NS):
                    continue
                rt=''.join(r.xpath('./w:t/text()',namespaces=NS))
                if not rt:continue
                rp=r.find(Q('rPr')); c=rp.find(Q('rStyle')) if rp is not None else None
                cs=c.get(Q('val')) if c is not None else None
                props=resolver.properties('rPr',sid,rp,character=cs)
                v=prop_values(props,RUN);v['fonts']=resolved_fonts(props,parts)
                if spans and spans[-1][1]==v:spans[-1][0]+=rt
                else:spans.append([rt,v])
            # Field-result text may legitimately change. For paragraphs containing
            # fields compare stable paragraph properties; field checks own results.
            if rec['fields']:spans=[]
            rows.append({'text':text,'paragraph':prop_values(pr,PARAGRAPH),'runs':spans})
        result[name]=rows
    return result


def compare(before, after):
    """Return located effective-property differences, not raw styles.xml diffs."""
    with package_session():
        a,b=snapshot(before),snapshot(after)
    issues=[]
    for part in sorted(set(a)|set(b)):
        left,right=a.get(part,[]),b.get(part,[])
        if len(left)!=len(right):
            issues.append({'code':'office-content-structure-drift','part':part,'before_count':len(left),'after_count':len(right)})
            continue
        for n,(old,new) in enumerate(zip(left,right),1):
            if old['text']!=new['text']:
                issues.append({'code':'office-content-drift','part':part,'paragraph':n,'before':old['text'][:120],'after':new['text'][:120]})
            props={k:{'before':old['paragraph'][k],'after':new['paragraph'][k]}
                   for k in old['paragraph'] if old['paragraph'][k]!=new['paragraph'][k]}
            if props:
                issues.append({'code':'office-format-drift','part':part,'paragraph':n,'properties':props})
            if old['runs']!=new['runs']:
                issues.append({'code':'office-run-format-drift','part':part,'paragraph':n,
                               'before':old['runs'][:3],'after':new['runs'][:3]})
    return {'status':'failed' if issues else 'passed','issues':issues,
            'source_sha256':sha(before),'output_sha256':sha(after),
            'scope':'effective text/paragraph formats; full template/table/page checks remain required'}


def make_probe(source, out):
    """Small representative sample using current shared styles/settings, never source text."""
    parts,root,body=package(source);root=copy.deepcopy(root);body=root.find(Q('body'))
    samples=[];seen=set()
    for p in body.iter(Q('p')):
        if p.xpath('ancestor::w:drawing|ancestor::w:pict|ancestor::w:txbxContent',namespaces=NS):continue
        sid=p.find('w:pPr/w:pStyle',NS)
        key=sid.get(Q('val')) if sid is not None else None
        if key in seen:continue
        seen.add(key)
        n=E.Element(Q('p'));pp=p.find(Q('pPr'))
        if pp is not None:
            pp=copy.deepcopy(pp)
            for tag in ('sectPr','pageBreakBefore','numPr'):
                for x in list(pp.findall(Q(tag))):pp.remove(x)
            n.append(pp)
        r=E.SubElement(n,Q('r'));rp=next(iter(p.findall('.//'+Q('rPr'))),None)
        if rp is not None:r.append(copy.deepcopy(rp))
        E.SubElement(r,Q('t')).text='排版稳定性样例 Abc 123'
        samples.append(n)
        if len(samples)==24:break
    if not samples:
        n=E.Element(Q('p'));E.SubElement(E.SubElement(n,Q('r')),Q('t')).text='排版样例';samples.append(n)
    section=body.find(Q('sectPr'));section=copy.deepcopy(section) if section is not None else None
    for child in list(body):body.remove(child)
    body.extend(samples)
    if section is not None:
        # Probe layout does not require the source document's header/footer contents.
        for tag in ('headerReference','footerReference'):
            for n in list(section.findall(Q(tag))):section.remove(n)
        body.append(section)
    copied=dict(parts);copied[DOC]=E.tostring(root,encoding='UTF-8',xml_declaration=True,standalone=True)
    for name in list(copied):
        if name.startswith('word/') and Path(name).stem.startswith(('header','footer','footnotes','endnotes')) and name.endswith('.xml'):
            n=parse(copied[name])
            for c in list(n):n.remove(c)
            copied[name]=E.tostring(n)
    out=Path(out);out.parent.mkdir(parents=True,exist_ok=True)
    with zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED) as archive:
        for name,data in copied.items():archive.writestr(name,data)
    return {'sampled_style_count':len(samples),'coverage':'representative used text styles (up to 24), not complete document verification'}


def probe_once(work, state, proposal, refresh):
    """Run before block mutations; rerun only if shared settings/engine/code change."""
    from block_cache import code_digest
    settings=state['settings'];parts,_,_=package(proposal)
    fingerprint=hashlib.sha256()
    for name in ('word/styles.xml','word/numbering.xml','word/settings.xml','word/theme/theme1.xml'):
        fingerprint.update(name.encode());fingerprint.update(parts.get(name,b''))
    fingerprint.update(json.dumps({k:settings.get(k) for k in ('field_engine','uno_python')},sort_keys=True).encode())
    # Actual probe capability/version changes can release a prior environmental block.
    runtime=(state.get('runtime_preflight') or {}).get('field_engine') or {}
    fingerprint.update(json.dumps(runtime,sort_keys=True,default=str).encode())
    fingerprint.update(code_digest(Path(__file__).parent).encode())
    key=fingerprint.hexdigest();folder=Path(work)/'office-probe'/key;report=folder/'result.json'
    if report.is_file():
        prior=json.loads(report.read_text(encoding='utf-8'))
        if prior.get('status')=='passed' and all(Path(prior[k]['path']).is_file() and sha(prior[k]['path'])==prior[k]['sha256'] for k in ('sample','saved')):
            return prior
        if prior.get('status')=='failed':
            raise ValueError('同一保存条件的样例已失败；先修正原因，不能重复运行：'+str(report))
    folder.mkdir(parents=True,exist_ok=True)
    sample=folder/'sample.docx';saved=folder/'saved.docx';coverage=make_probe(proposal,sample)
    result={'key':key,'coverage':coverage,'status':'failed'}
    try:
        engine=refresh(sample,saved,settings.get('field_engine','auto'),settings.get('uno_python'),
                       preflight=(state.get('runtime_preflight') or {}).get('field_engine'))
        result.update(compare(sample,saved));result['engine']=engine
        result.update(sample={'path':str(sample),'sha256':sha(sample)},saved={'path':str(saved),'sha256':sha(saved)})
    except Exception as exc:
        result['issues']=[exc.as_issue() if hasattr(exc,'as_issue') else {'code':'office-probe-failed','message':str(exc)}]
    write_json(report,result)
    if result['status']!='passed':
        raise ValueError('Office 保存稳定性检查未通过，已保留前后样例；停止重复保存：'+str(report))
    return result
