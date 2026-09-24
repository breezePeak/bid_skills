#!/usr/bin/env python3
"""Resolve a usable Office path before modifying a document.

Explicit engine requests never fall back to another engine. No capability result is a
field-update or visual-acceptance result; those existing gates still run later.
"""
from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import os
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
from office_backends import PRIORITY, OfficeError, probe_native, select_renderer, find_office as _find_office


class PreflightError(ValueError):
    def __init__(self, code: str, message: str, **details):
        super().__init__(message)
        self.code, self.details = code, details

    def as_issue(self):
        return {'severity': 'error', 'code': self.code, 'message': str(self), **self.details}


def digest(path):
    return 'sha256:' + hashlib.sha256(Path(path).read_bytes()).hexdigest()


def run_probe(command, timeout=15):
    try:
        cp = subprocess.run(command, capture_output=True, text=True, check=False, timeout=timeout)
        return {'available': cp.returncode == 0, 'command': command, 'returncode': cp.returncode,
                'detail': (cp.stdout + '\n' + cp.stderr).strip()[-1500:]}
    except (OSError, subprocess.TimeoutExpired) as exc:
        return {'available': False, 'command': command, 'detail': str(exc)}


def probe_word():
    return probe_native('word')


def probe_wps():
    return probe_native('wps')


def find_office():
    return _find_office()


def probe_libreoffice(uno_python=None):
    office = find_office()
    if not office:
        return {'available': False, 'detail': '未找到 LibreOffice/soffice。'}
    version = run_probe([office, '--headless', '--version'])
    if not version['available']:
        return {**version, 'soffice': office}
    if uno_python:
        candidates = [str(uno_python)]
    else:
        candidates = list(dict.fromkeys([sys.executable, '/usr/bin/python3']))
    attempts = []
    for python in candidates:
        found = python if Path(python).is_file() else shutil.which(python)
        if not found:
            continue
        probe = run_probe([found, '-c', 'import uno; import unohelper; print("UNO ready")'])
        attempts.append(probe)
        if probe['available']:
            return {'available': True, 'soffice': office, 'uno_python': found,
                    'version': version['detail'], 'uno_probe': probe}
    return {'available': False, 'soffice': office, 'detail': 'LibreOffice 存在，但没有可导入 UNO 的 Python。',
            'attempts': attempts}


def select_engine(requested='auto', uno_python=None, *, exclude=()):
    if requested not in {'auto', *PRIORITY}:
        raise PreflightError('office-engine-invalid', '不支持的域更新引擎。', requested=requested)
    attempts = {}
    for engine in (PRIORITY if requested == 'auto' else (requested,)):
        if engine in exclude:
            continue
        probe = probe_word() if engine == 'word' else probe_wps() if engine == 'wps' else probe_libreoffice(uno_python)
        attempts[engine] = probe
        if probe.get('available'):
            return {'engine': engine, 'requested': requested, 'priority': list(PRIORITY),
                    'uno_python': probe.get('uno_python'), 'capabilities': attempts}
    code = 'office-engine-unavailable' if requested == 'auto' else requested + '-engine-unavailable'
    raise PreflightError(code, '没有可用的真实域更新引擎；明确指定引擎时不自动降级。',
                         requested=requested, probes=attempts)


def validate_docx(path):
    path = Path(path)
    if not path.is_file():
        raise PreflightError('document-missing', '文档或模板不存在。', path=str(path))
    try:
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
            if len(names) != len(set(names)):
                raise ValueError('ZIP 含重复成员')
            for required in ('[Content_Types].xml', 'word/document.xml', 'word/styles.xml'):
                ET.fromstring(z.read(required))
            bad = z.testzip()
            if bad:
                raise ValueError('ZIP CRC 失败: ' + bad)
    except (OSError, ValueError, KeyError, zipfile.BadZipFile, ET.ParseError) as exc:
        raise PreflightError('document-unreadable', '文档或模板无法完整解析。', path=str(path), detail=str(exc)) from exc
    return {'path': str(path.resolve()), 'sha256': digest(path)}


def validate_profile(template, profile):
    if profile is None:
        return None
    try:
        data = json.loads(Path(profile).read_text(encoding='utf-8-sig'))
        if not isinstance(data, dict):
            raise ValueError('JSON 根节点必须为对象')
    except (OSError, ValueError) as exc:
        raise PreflightError('template-profile-invalid', '模板规则 JSON 不可读取。', detail=str(exc)) from exc
    expected = data.get('template', {}).get('canonical_template_sha256')
    if expected and expected != digest(template):
        raise PreflightError('template-profile-stale', '模板规则不属于当前模板，必须重新解析。')
    return {'path': str(Path(profile).resolve()), 'sha256': digest(profile)}


def check(source, template, work_dir, *, requested='auto', profile=None,
          audit_only=False, uno_python=None, renderer='auto'):
    documents = {'source': validate_docx(source), 'template': validate_docx(template)}
    rule_profile = validate_profile(template, profile)
    work_dir = Path(work_dir)
    work_dir.mkdir(parents=True, exist_ok=True)
    try:
        with tempfile.TemporaryFile(dir=work_dir) as probe:
            probe.write(b'preflight')
            probe.flush()
    except OSError as exc:
        raise PreflightError('work-directory-unwritable', '工作目录不可写；未开始文档修改。') from exc
    ppm = shutil.which('pdftoppm') or shutil.which('pdftoppm.exe')
    if not ppm:
        raise PreflightError('renderer-unavailable', '缺少 pdftoppm；不能跳过页面图片验收。')
    probe = run_probe([ppm, '-v'])
    if not probe['available']:
        raise PreflightError('renderer-unusable', '页面转换程序无法执行。', probe=probe)
    try:
        rendering = select_renderer(renderer)
    except OfficeError as exc:
        raise PreflightError(exc.code, str(exc), **exc.details) from exc
    selected = {'engine': None, 'requested': requested, 'reason': 'audit-only 使用已核验的更新域报告'} if audit_only else select_engine(requested, uno_python)
    return {'status': 'passed', 'documents': documents, 'profile': rule_profile,
            'field_engine': selected, 'renderer': {**rendering, 'pdftoppm': ppm},
            'note': '此结果只证明执行条件可用，不代表文档修复、域更新或视觉验收已通过。'}


def main():
    ap = argparse.ArgumentParser(description='在修改前检查模板、更新域引擎和渲染条件')
    ap.add_argument('input', type=Path)
    ap.add_argument('--template', type=Path, required=True)
    ap.add_argument('--template-style-json', type=Path)
    ap.add_argument('--work-dir', type=Path, required=True)
    ap.add_argument('--field-engine', choices=['auto', 'word', 'wps', 'libreoffice'], default='auto')
    ap.add_argument('--renderer', choices=['auto', 'word', 'wps', 'libreoffice'], default='auto')
    ap.add_argument('--uno-python')
    ap.add_argument('--audit-only', action='store_true')
    ap.add_argument('--json-out', type=Path)
    args = ap.parse_args()
    try:
        result = check(args.input, args.template, args.work_dir, requested=args.field_engine,
                       profile=args.template_style_json, audit_only=args.audit_only, uno_python=args.uno_python, renderer=args.renderer)
    except PreflightError as exc:
        result = {'status': 'blocked', 'issues': [exc.as_issue()]}
    payload = json.dumps(result, ensure_ascii=False, indent=2)
    if args.json_out:
        args.json_out.parent.mkdir(parents=True, exist_ok=True)
        args.json_out.write_text(payload + '\n', encoding='utf-8')
    print(payload)
    return 0 if result['status'] == 'passed' else 2


if __name__ == '__main__':
    raise SystemExit(main())
