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
from office_backends import PRIORITY, OfficeError, native_operation, find_office
from numbering_policy import file_digest, parse, fields, command, Q, ERROR_TEXT, PolicyError

from caption_field_guard import (
    CaptionFieldError,
    require_valid as _require_caption_fields,
    preserve_native_instructions as _preserve_native_instructions,
)

HERE = Path(__file__).resolve().parent


def _caption_call(action, *args, **kwargs):
    try:
        return action(*args, **kwargs)
    except CaptionFieldError as exc:
        # Keep the pipeline's existing structured failure/retry protocol.
        raise PolicyError(exc.code, str(exc), **exc.details) from exc


def require_caption_fields(*args, **kwargs):
    return _caption_call(_require_caption_fields, *args, **kwargs)


def preserve_native_instructions(*args, **kwargs):
    return _caption_call(_preserve_native_instructions, *args, **kwargs)


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


def _refresh_once(source, out, engine='auto', uno_python=None, *, preflight=None):
    source,out=Path(source),Path(out)
    if source.resolve()==out.resolve():raise PolicyError('in-place-write','刷新域输出不能覆盖输入。')
    from runtime_preflight import select_engine, PreflightError
    try:
        selected = preflight or select_engine(engine, uno_python)
    except PreflightError as exc:
        raise PolicyError(exc.code, str(exc), **exc.details) from exc
    if selected.get('engine') not in {'word', 'wps', 'libreoffice'} or (engine != 'auto' and selected['engine'] != engine):
        raise PolicyError('office-engine-mismatch', '预检查引擎与本次明确要求不一致。')
    engine = selected['engine']; uno_python = selected.get('uno_python') or uno_python
    safe_fields(source)
    # Check native instructions before asking Office to update them; stale numeric
    # caches after an edit are allowed here, malformed references are not.
    require_caption_fields(source, check_cache=False)
    source_hash=file_digest(source);out.parent.mkdir(parents=True,exist_ok=True)
    with tempfile.TemporaryDirectory(dir=out.parent) as work:
        candidate=Path(work)/'refreshed.docx';report=Path(work)/'engine.json'
        if engine in {'word', 'wps'}:
            snapshot = Path(work) / 'input.docx'
            shutil.copyfile(source, snapshot)
            try:
                data = native_operation(engine, 'refresh', source=snapshot, output=candidate,
                                        work_dir=work, timeout=180)
            except OfficeError as exc:
                raise PolicyError(exc.code, str(exc), **exc.details) from exc
        elif engine == 'libreoffice':
            office = find_office()
            if not office:
                raise PolicyError('office-engine-unavailable', '未安装 LibreOffice。')
            python = uno_python or ('/usr/bin/python3' if Path('/usr/bin/python3').exists() else sys.executable)
            cmd = [python,str(HERE/'refresh_fields_uno.py'),str(source.resolve()),str(candidate.resolve()),str(report.resolve()),'--soffice',office]
            try:
                cp = subprocess.run(cmd,capture_output=True,text=True,timeout=180,check=False)
            except (OSError, subprocess.TimeoutExpired) as exc:
                raise PolicyError('field-refresh-failed', 'Office 更新域启动失败或超时。', detail=str(exc)) from exc
            try:
                data = json.loads(report.read_text(encoding='utf-8-sig')) if report.is_file() else {}
            except (OSError, ValueError) as exc:
                raise PolicyError('field-refresh-failed', '更新域引擎未返回有效报告。') from exc
            if cp.returncode:
                raise PolicyError('field-refresh-failed','实际更新域失败。',engine=engine,detail=data,stderr=cp.stderr[-4000:])
        else:
            raise PolicyError('office-engine-invalid','不支持的域刷新引擎。')
        if data.get('status') != 'passed' or not candidate.is_file():
            raise PolicyError('field-refresh-failed', '没有实际更新域的有效输出。', engine=engine, detail=data)
        errors=result_errors(candidate)
        if errors:raise PolicyError('field-result-error','实际更新域后仍有错误，未写入输出。',errors=errors)
        # Keep real engine-calculated results, but do not let an exporter discard
        # native Word SEQ switches needed for the next F9. This operation is fully
        # recorded, source-bound and is part of refresh serialization.
        preserved = Path(work) / 'native-preserved.docx'
        preservation = preserve_native_instructions(source, candidate, preserved)
        os.replace(preserved, candidate)
        caption_check = require_caption_fields(candidate)
        if file_digest(source)!=source_hash:raise PolicyError('source-changed-by-office','Office 更新域时改动了输入文件，停止交付。')
        os.replace(candidate,out)
    data.update({'source_sha256':file_digest(source),'output_sha256':file_digest(out),'output':str(out),
                 'errors':[], 'native_caption_preservation':preservation, 'caption_fields':caption_check, 'note':'已由所列引擎更新并保存；WPS/LibreOffice 结果不等同于 Windows Word F9 的兼容性保证。'})
    return data


# Only engine/runtime failures allow fallback. A malformed field, invalid cached
# result, unsafe external field or content/preservation error remains a hard FAIL.
RETRYABLE_ENGINE_ERRORS = {
    'word-engine-unavailable', 'wps-engine-unavailable', 'office-engine-unavailable',
    'office-com-unavailable', 'office-launch-failed', 'office-command-failed',
    'office-timeout', 'office-worker-failed', 'field-refresh-failed',
}


def refresh(source, out, engine='auto', uno_python=None, *, preflight=None):
    from runtime_preflight import select_engine, PreflightError
    attempts = []
    excluded = set()
    selected = preflight
    while True:
        try:
            selected = selected or select_engine(engine, uno_python, exclude=excluded)
        except PreflightError as exc:
            raise PolicyError(exc.code, str(exc), attempts=attempts, **exc.details) from exc
        actual = selected.get('engine')
        if actual not in PRIORITY or (engine != 'auto' and actual != engine) or actual in excluded:
            raise PolicyError('office-engine-mismatch', '域更新预检与本次要求不一致。')
        for name, probe in selected.get('capabilities', {}).items():
            if name != actual and not probe.get('available'):
                attempts.append({'engine': name, 'status': 'unavailable', 'reason': probe.get('detail', '不可用')})
        try:
            data = _refresh_once(source, out, actual, uno_python, preflight=selected)
            attempts.append({'engine': actual, 'status': 'passed'})
            data.update(engine_requested=engine, engine_priority=list(PRIORITY), engine_attempts=attempts)
            return data
        except PolicyError as exc:
            attempts.append({'engine': actual, 'status': 'failed', 'issue': exc.as_issue()})
            if engine != 'auto' or exc.code not in RETRYABLE_ENGINE_ERRORS:
                raise
            excluded.update(PRIORITY[:PRIORITY.index(actual)+1])
            if len(excluded) == len(PRIORITY):
                raise PolicyError('field-refresh-failed', '所有域更新引擎失败；不得跳过更新和验收。', attempts=attempts) from exc
            selected = None


def validate_report(candidate, report):
    data=report if isinstance(report,dict) else json.loads(Path(report).read_text(encoding='utf-8-sig'))
    if data.get('status')!='passed' or data.get('engine') not in {'Microsoft Word','WPS Writer','LibreOffice UNO'} or data.get('fields_updated') is not True or data.get('indexes_updated') is not True or data.get('errors'):
        raise PolicyError('field-refresh-unverified','缺少真实 Office 更新域和目录的成功记录。')
    if data.get('output_sha256')!=file_digest(candidate):raise PolicyError('field-refresh-stale','域刷新后文档又变了，必须重新更新、检查和渲染。')
    errors=result_errors(candidate)
    if errors:raise PolicyError('field-result-error','最新文件的域结果仍有错误。',errors=errors)
    require_caption_fields(candidate)
    return data


def main():
    ap=argparse.ArgumentParser();ap.add_argument('input',type=Path);ap.add_argument('--out',type=Path,required=True);ap.add_argument('--engine',choices=['auto','word','wps','libreoffice'],default='auto');ap.add_argument('--uno-python');ap.add_argument('--json-out',type=Path,required=True);a=ap.parse_args()
    try:result=refresh(a.input,a.out,a.engine,a.uno_python)
    except Exception as exc:result={'status':'failed','issues':[exc.as_issue() if isinstance(exc,PolicyError) else {'severity':'error','code':'field-refresh-error','message':str(exc)}]}
    a.json_out.parent.mkdir(parents=True,exist_ok=True);a.json_out.write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8');print(json.dumps(result,ensure_ascii=False,indent=2))
    return 0 if result['status']=='passed' else 2


if __name__=='__main__':raise SystemExit(main())
