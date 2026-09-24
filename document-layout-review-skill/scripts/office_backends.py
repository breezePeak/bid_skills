#!/usr/bin/env python3
"""Office capability/worker helpers. Default priority is never installation order."""
from __future__ import annotations
import hashlib
import json
import os
import shutil
import subprocess
import tempfile
from pathlib import Path

PRIORITY = ('word', 'wps', 'libreoffice')
ENGINE_NAMES = {'word': 'Microsoft Word', 'wps': 'WPS Writer', 'libreoffice': 'LibreOffice'}
HERE = Path(__file__).resolve().parent


class OfficeError(ValueError):
    def __init__(self, code, message, **details):
        super().__init__(message)
        self.code, self.details = code, details

    def as_issue(self):
        return {'severity': 'error', 'code': self.code, 'message': str(self), **self.details}


def digest(path):
    h = hashlib.sha256()
    with Path(path).open('rb') as f:
        for block in iter(lambda: f.read(1024 * 1024), b''):
            h.update(block)
    return 'sha256:' + h.hexdigest()


def engine_order(requested='auto'):
    if requested == 'auto':
        return PRIORITY
    if requested not in PRIORITY:
        raise OfficeError('office-engine-invalid', '不支持的 Office 引擎。', requested=requested)
    return (requested,)


def executable(explicit, names):
    if explicit:
        p = Path(explicit)
        return str(p.resolve()) if p.is_file() else shutil.which(str(explicit))
    return next((p for name in names if (p := shutil.which(name))), None)


def find_office(explicit=None):
    found = executable(explicit, ('soffice.com', 'libreoffice', 'soffice', 'soffice.exe'))
    if found or explicit:
        return found
    # Desktop installers do not necessarily add their program directory to PATH.
    candidates = [Path('/Applications/LibreOffice.app/Contents/MacOS/soffice')]
    for key in ('ProgramFiles', 'ProgramFiles(x86)'):
        if os.environ.get(key):
            candidates.append(Path(os.environ[key]) / 'LibreOffice/program/soffice.com')
            candidates.append(Path(os.environ[key]) / 'LibreOffice/program/soffice.exe')
    return next((str(p) for p in candidates if p.is_file()), None)


def run(command, timeout=120, *, env=None):
    try:
        cp = subprocess.run(command, capture_output=True, text=True, encoding='utf-8',
                            errors='replace', check=False, timeout=timeout, env=env)
    except subprocess.TimeoutExpired as exc:
        raise OfficeError('office-timeout', 'Office/页面转换超时。', timeout=timeout) from exc
    except OSError as exc:
        raise OfficeError('office-launch-failed', '无法启动转换程序。', detail=str(exc)) from exc
    if cp.returncode:
        raise OfficeError('office-command-failed', '转换程序执行失败。', returncode=cp.returncode,
                          stdout=cp.stdout[-3000:], stderr=cp.stderr[-3000:])
    return cp


def native_operation(engine, mode, *, source=None, output=None, work_dir=None, timeout=120):
    if engine not in ('word', 'wps'):
        raise OfficeError('office-engine-invalid', 'COM 入口仅支持 Word/WPS。')
    shell = executable(None, ('powershell.exe', 'powershell', 'pwsh.exe', 'pwsh'))
    if os.name != 'nt' or not shell:
        raise OfficeError('office-com-unavailable', '当前执行环境没有可调用的 Windows Office COM。', engine=engine)
    with tempfile.TemporaryDirectory(prefix='office-com-', dir=work_dir) as tmp:
        report = Path(tmp) / 'worker.json'
        command = [shell, '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
                   '-File', str(HERE / 'office_com.ps1'), '-Engine', engine, '-Mode', mode,
                   '-ReportPath', str(report)]
        if source is not None:
            command += ['-InputPath', str(Path(source).resolve())]
        if output is not None:
            command += ['-OutputPath', str(Path(output).resolve())]
        failure = None
        try:
            run(command, timeout)
        except OfficeError as exc:
            failure = exc
        try:
            data = json.loads(report.read_text(encoding='utf-8-sig')) if report.is_file() else {}
        except (OSError, ValueError):
            data = {}
        if failure or data.get('status') != 'passed' or data.get('engine') != ENGINE_NAMES[engine]:
            if data.get('error_code') in ('external-field-update-blocked', 'field-result-error'):
                raise OfficeError(data['error_code'], 'Office 文档检查失败，不能用降级掩盖。', detail=data)
            raise OfficeError(failure.code if failure else 'office-worker-failed',
                              '实际 Office 调用未成功。', engine=engine, report=data,
                              cause=failure.as_issue() if failure else None)
        return data


def probe_native(engine):
    try:
        result = native_operation(engine, 'probe', timeout=20)
        return {'available': True, 'version': result.get('version'), 'application': result.get('application')}
    except OfficeError as exc:
        return {'available': False, 'detail': str(exc), 'issue': exc.as_issue()}


def probe_renderer(engine, soffice=None):
    if engine in ('word', 'wps'):
        return probe_native(engine)
    office = find_office(soffice)
    if not office:
        return {'available': False, 'detail': '未找到 LibreOffice/soffice。'}
    try:
        cp = run([office, '--headless', '--version'], timeout=15)
        return {'available': True, 'soffice': office, 'version': cp.stdout.strip()}
    except OfficeError as exc:
        return {'available': False, 'soffice': office, 'detail': str(exc), 'issue': exc.as_issue()}


def select_renderer(requested='auto', soffice=None):
    attempts = {}
    for engine in engine_order(requested):
        attempts[engine] = probe_renderer(engine, soffice)
        if attempts[engine]['available']:
            return {'engine': engine, 'requested': requested, 'priority': list(PRIORITY),
                    'capabilities': attempts, 'soffice': attempts[engine].get('soffice')}
    raise OfficeError('renderer-unavailable', '没有可用的渲染引擎；不能跳过视觉验收。',
                      requested=requested, probes=attempts)
