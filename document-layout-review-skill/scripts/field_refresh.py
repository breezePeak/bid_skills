#!/usr/bin/env python3
"""Refresh with a real office engine; setting updateFields alone is not proof."""
from __future__ import annotations
import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from numbering_policy import file_digest, parse, fields, command, Q, ERROR_TEXT, PolicyError

HERE = Path(__file__).resolve().parent


def result_errors(path):
    errors=[]
    with zipfile.ZipFile(path) as z:
        for name in z.namelist():
            if not name.startswith('word/') or not name.endswith('.xml'):continue
            if Path(name).name.startswith(('document','header','footer','footnotes','endnotes')):
                root=parse(z.read(name))
                for p in root.iter(Q('p')):
                    for f in fields(p):
                        if ERROR_TEXT.search(f.get('cached','')):errors.append({'part':name,'result':f['cached']})
    return errors


def safe_fields(source):
    with zipfile.ZipFile(source) as z:
        for name in z.namelist():
            if name.startswith('word/') and name.endswith('.xml'):
                root=parse(z.read(name))
                for p in root.iter(Q('p')):
                    for f in fields(p):
                        if command(f) in {'DDE','DDEAUTO','INCLUDETEXT','INCLUDEPICTURE','DATABASE','LINK','RD'}:
                            raise PolicyError('external-field-update-blocked','更新外部内容域需要单独确认，不能自动联网刷新。',part=name)


def refresh(source, out, engine='word', uno_python=None):
    source,out=Path(source),Path(out)
    if source.resolve()==out.resolve():raise PolicyError('in-place-write','刷新域输出不能覆盖输入。')
    safe_fields(source);source_hash=file_digest(source);out.parent.mkdir(parents=True,exist_ok=True)
    with tempfile.TemporaryDirectory(dir=out.parent) as work:
        candidate=Path(work)/'refreshed.docx';report=Path(work)/'engine.json'
        if engine=='word':
            shell=shutil.which('powershell') or shutil.which('pwsh')
            if os.name!='nt' or not shell:raise PolicyError('word-engine-unavailable','当前环境没有 Windows Word COM；不能把写入缓存或 updateFields 当作 F9 已验证。')
            cmd=[shell,'-NoProfile','-ExecutionPolicy','Bypass','-File',str(HERE/'refresh_fields_word.ps1'),'-InputPath',str(source.resolve()),'-OutputPath',str(candidate.resolve()),'-ReportPath',str(report.resolve())]
        elif engine=='libreoffice':
            office=shutil.which('libreoffice') or shutil.which('soffice')
            if not office:raise PolicyError('office-engine-unavailable','未安装 LibreOffice。')
            python=uno_python or ('/usr/bin/python3' if Path('/usr/bin/python3').exists() else sys.executable)
            cmd=[python,str(HERE/'refresh_fields_uno.py'),str(source.resolve()),str(candidate.resolve()),str(report.resolve()),'--soffice',office]
        else:raise PolicyError('office-engine-invalid','不支持的域刷新引擎。')
        cp=subprocess.run(cmd,capture_output=True,text=True,timeout=180,check=False)
        data=json.loads(report.read_text(encoding='utf-8-sig')) if report.is_file() else {}
        if cp.returncode or data.get('status')!='passed' or not candidate.is_file():
            raise PolicyError('field-refresh-failed','实际更新域失败。',engine=engine,detail=data,stderr=cp.stderr[-4000:])
        errors=result_errors(candidate)
        if errors:raise PolicyError('field-result-error','实际更新域后仍有错误，未写入输出。',errors=errors)
        if file_digest(source)!=source_hash:raise PolicyError('source-changed-by-office','Office 更新域时改动了输入文件，停止交付。')
        os.replace(candidate,out)
    data.update({'source_sha256':file_digest(source),'output_sha256':file_digest(out),'output':str(out),
                 'errors':[], 'note':'已由所列引擎更新并保存；LibreOffice 结果不等同于 Windows Word F9 的兼容性保证。'})
    return data


def validate_report(candidate, report):
    data=report if isinstance(report,dict) else json.loads(Path(report).read_text(encoding='utf-8-sig'))
    if data.get('status')!='passed' or data.get('engine') not in {'Microsoft Word','LibreOffice UNO'} or data.get('fields_updated') is not True or data.get('indexes_updated') is not True or data.get('errors'):
        raise PolicyError('field-refresh-unverified','缺少真实 Office 更新域和目录的成功记录。')
    if data.get('output_sha256')!=file_digest(candidate):raise PolicyError('field-refresh-stale','域刷新后文档又变了，必须重新更新、检查和渲染。')
    errors=result_errors(candidate)
    if errors:raise PolicyError('field-result-error','最新文件的域结果仍有错误。',errors=errors)
    return data


def main():
    ap=argparse.ArgumentParser();ap.add_argument('input',type=Path);ap.add_argument('--out',type=Path,required=True);ap.add_argument('--engine',choices=['word','libreoffice'],default='word');ap.add_argument('--uno-python');ap.add_argument('--json-out',type=Path,required=True);a=ap.parse_args()
    try:result=refresh(a.input,a.out,a.engine,a.uno_python)
    except Exception as exc:result={'status':'failed','issues':[exc.as_issue() if isinstance(exc,PolicyError) else {'severity':'error','code':'field-refresh-error','message':str(exc)}]}
    a.json_out.parent.mkdir(parents=True,exist_ok=True);a.json_out.write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8');print(json.dumps(result,ensure_ascii=False,indent=2))
    return 0 if result['status']=='passed' else 2


if __name__=='__main__':raise SystemExit(main())
