#!/usr/bin/env python3
"""Render using Word -> WPS -> LibreOffice, then verify fresh complete PNG output.

The input is never saved by Office. A render is not a visual acceptance result.
"""
from __future__ import annotations
import argparse
import json
import os
import re
import shutil
import struct
import tempfile
import uuid
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
from office_backends import (PRIORITY, ENGINE_NAMES, OfficeError, digest, engine_order,
                             executable, find_office, native_operation, probe_renderer, run)


class RenderError(OfficeError):
    pass


def write_json(path, value):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(path.name + '.' + uuid.uuid4().hex + '.tmp')
    try:
        temporary.write_text(json.dumps(value, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)


def guard_source(source):
    """Block automatic external-content loading; hyperlinks remain ordinary links."""
    W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    forbidden = re.compile(r'^\s*(DDEAUTO|DDE|INCLUDETEXT|INCLUDEPICTURE|DATABASE|LINK|RD)\b', re.I)
    try:
        with zipfile.ZipFile(source) as z:
            if z.testzip() or len(z.namelist()) != len(set(z.namelist())):
                raise ValueError('DOCX ZIP invalid')
            ET.fromstring(z.read('word/document.xml'))
            for name in z.namelist():
                if name.lower().endswith('vbaproject.bin'):
                    raise RenderError('document-macros-blocked', '不能自动执行含宏文档。')
                if name.endswith('.rels'):
                    for rel in ET.fromstring(z.read(name)):
                        if rel.get('TargetMode') == 'External' and not rel.get('Type', '').endswith('/hyperlink'):
                            raise RenderError('external-content-blocked', '文档含外部内容关系，须先核验。', part=name)
                if name.startswith('word/') and name.endswith('.xml'):
                    root = ET.fromstring(z.read(name))
                    for field in root.iter(W + 'fldSimple'):
                        if forbidden.match(field.get(W + 'instr', '')):
                            raise RenderError('external-content-blocked', '文档含外部内容域。', part=name)
                    stack = []
                    for el in root.iter():
                        if el.tag == W + 'fldChar':
                            kind = el.get(W + 'fldCharType')
                            if kind == 'begin':
                                stack.append([])
                            elif kind in ('separate', 'end') and stack:
                                if forbidden.match(''.join(stack[-1])):
                                    raise RenderError('external-content-blocked', '文档含外部内容域。', part=name)
                                if kind == 'end':
                                    stack.pop()
                        elif el.tag == W + 'instrText' and stack:
                            stack[-1].append(el.text or '')
    except (OSError, ValueError, KeyError, zipfile.BadZipFile, ET.ParseError) as exc:
        if isinstance(exc, RenderError):
            raise
        raise RenderError('document-unreadable', '无法读取完整 DOCX。', detail=str(exc)) from exc


def export_pdf(engine, source, pdf, work, probe, timeout):
    if engine in ('word', 'wps'):
        return native_operation(engine, 'render', source=source, output=pdf,
                                work_dir=work, timeout=timeout)
    office = probe.get('soffice') or find_office()
    if not office:
        raise RenderError('renderer-unavailable', '未找到 LibreOffice。')
    profile = Path(work) / 'lo-profile'
    profile.mkdir()
    # Use an isolated profile, not the user's current interactive LibreOffice.
    cp = run([office, '-env:UserInstallation=' + profile.resolve().as_uri(),
              '--headless', '--nologo', '--nodefault', '--norestore',
              '--convert-to', 'pdf:writer_pdf_Export', '--outdir', str(pdf.parent), str(source)], timeout)
    actual = pdf.parent / (source.stem + '.pdf')
    if actual != pdf and actual.is_file():
        os.replace(actual, pdf)
    return {'engine': ENGINE_NAMES[engine], 'version': probe.get('version'),
            'stdout': cp.stdout[-2000:], 'stderr': cp.stderr[-2000:]}


def pdf_page_count(pdf, timeout):
    with Path(pdf).open('rb') as f:
        if f.read(5) != b'%PDF-':
            raise RenderError('invalid-pdf', '渲染器没有生成有效 PDF 文件头。')
    info = executable(None, ('pdfinfo', 'pdfinfo.exe'))
    if info:
        env = os.environ.copy(); env['LC_ALL'] = 'C'; env['LANG'] = 'C'
        cp = run([info, str(pdf)], timeout, env=env)
        match = re.search(r'^Pages:\s*(\d+)\s*$', cp.stdout, re.M)
        if not match or int(match.group(1)) < 1:
            raise RenderError('invalid-pdf', 'PDF 没有可用页面。')
        return int(match.group(1))
    return None


def rasterize(pdf, folder, ppm, dpi, timeout):
    folder.mkdir()
    expected = pdf_page_count(pdf, timeout)
    run([ppm, '-png', '-r', str(dpi), str(pdf), str(folder / 'page')], timeout)
    numbered = []
    for page in folder.iterdir():
        match = re.fullmatch(r'page-(\d+)\.png', page.name)
        if match:
            numbered.append((int(match.group(1)), page))
    numbered.sort(key=lambda item: item[0])
    if not numbered or [n for n, _ in numbered] != list(range(1, len(numbered) + 1)):
        raise RenderError('render-pages-incomplete', '页面图片为空、缺页或页码重复。')
    if expected is not None and len(numbered) != expected:
        raise RenderError('render-pages-incomplete', 'PNG 数量与 PDF 页数不一致。', expected=expected, actual=len(numbered))
    for _, page in numbered:
        data = page.read_bytes()
        if len(data) < 45 or data[:8] != b'\x89PNG\r\n\x1a\n' or data[12:16] != b'IHDR' or data[-8:-4] != b'IEND':
            raise RenderError('invalid-page-image', 'PNG 图片损坏或不完整。', page=page.name)
        width, height = struct.unpack('>II', data[16:24])
        if width < 1 or height < 1:
            raise RenderError('invalid-page-image', 'PNG 尺寸无效。', page=page.name)
    return [p for _, p in numbered]


def render(docx, out_dir, *, renderer='auto', soffice=None, pdftoppm=None, dpi=120, timeout=120):
    source, out_dir = Path(docx).resolve(), Path(out_dir).resolve()
    order = engine_order(renderer)
    if not source.is_file() or not 36 <= dpi <= 600 or timeout <= 0:
        raise RenderError('render-input-invalid', '输入文件、DPI 或超时设置无效。')
    guard_source(source)
    source_hash = digest(source)
    ppm = executable(pdftoppm, ('pdftoppm', 'pdftoppm.exe'))
    if not ppm:
        raise RenderError('page-renderer-unavailable', '缺少 pdftoppm，不能仅生成 PDF 后跳过逐页验收。')
    out_dir.mkdir(parents=True, exist_ok=True)
    attempts = []
    for engine in order:
        probe = probe_renderer(engine, soffice)
        if not probe.get('available'):
            attempts.append({'engine': engine, 'status': 'unavailable', 'reason': probe.get('detail', '不可用'), 'probe': probe})
            continue
        try:
            with tempfile.TemporaryDirectory(prefix='.render-work-', dir=out_dir) as tmp:
                work = Path(tmp)
                snapshot = work / 'input.docx'
                shutil.copyfile(source, snapshot)
                if digest(snapshot) != source_hash:
                    raise RenderError('render-source-changed', '开始渲染前源文档已改变。')
                artifacts = work / 'artifacts'; artifacts.mkdir()
                pdf = artifacts / 'document.pdf'
                detail = export_pdf(engine, snapshot, pdf, work, probe, timeout)
                if not pdf.is_file() or pdf.stat().st_size == 0:
                    raise RenderError('render-output-missing', '转换程序未生成新的非空 PDF。')
                pages = rasterize(pdf, artifacts / 'pages', ppm, dpi, timeout)
                if digest(source) != source_hash or digest(snapshot) != source_hash:
                    raise RenderError('render-source-changed', '渲染期间 DOCX 被改动，停止验收。')
                attempts.append({'engine': engine, 'status': 'passed', 'reason': '', 'details': detail})
                # Publish a new run directory only after all pages passed structural checks.
                published = out_dir / ('render-' + uuid.uuid4().hex)
                os.replace(artifacts, published)
                pdf = published / 'document.pdf'
                pages = [published / 'pages' / p.name for p in pages]
                report_path = published / 'render-report.json'
                result = {'status': 'passed', 'renderer': engine, 'engine': ENGINE_NAMES[engine],
                          'renderer_requested': renderer, 'priority': list(PRIORITY),
                          'attempts': attempts, 'docx': str(source), 'source_sha256': source_hash,
                          'pdf': str(pdf), 'pdf_sha256': digest(pdf), 'pages': list(map(str, pages)),
                          'page_records': [{'name': p.name, 'path': str(p), 'sha256': digest(p)} for p in pages],
                          'page_count': len(pages), 'dpi': dpi, 'png_renderer_available': True,
                          'report_path': str(report_path), 'visual_review_status': 'pending',
                          'note': '实际渲染引擎如实记录；渲染成功不等于逐页视觉验收通过。'}
                write_json(report_path, result)
                write_json(out_dir / 'render-report.json', result)
                return result
        except OfficeError as exc:
            if exc.code in ('render-source-changed', 'external-field-update-blocked', 'field-result-error'):
                raise
            if digest(source) != source_hash:
                raise RenderError('render-source-changed', '失败的渲染调用期间源文档被改变。') from exc
            attempts.append({'engine': engine, 'status': 'failed', 'reason': str(exc), 'issue': exc.as_issue()})
        except (OSError, ValueError) as exc:
            if digest(source) != source_hash:
                raise RenderError('render-source-changed', '渲染调用期间源文档被改变。') from exc
            attempts.append({'engine': engine, 'status': 'failed', 'reason': str(exc)})
    failure = {'status': 'failed', 'renderer_requested': renderer, 'priority': list(PRIORITY),
               'source_sha256': source_hash, 'attempts': attempts, 'pages': [], 'visual_review_status': 'blocked'}
    write_json(out_dir / 'render-report.json', failure)
    raise RenderError('all-renderers-failed', '所有允许的渲染器均失败，停止交付；不得跳过视觉验收。', **failure)


def validate_render(candidate, data, requested='auto', pages=None):
    """Fail closed on stale output, missing provenance, wrong engine or broken order."""
    if not isinstance(data, dict) or data.get('status') != 'passed' or data.get('renderer') not in PRIORITY:
        raise RenderError('render-unverified', '缺少真实渲染成功记录。')
    engine = data['renderer']
    if data.get('engine') != ENGINE_NAMES[engine] or data.get('renderer_requested') != requested:
        raise RenderError('render-engine-mismatch', '渲染引擎记录与明确要求不一致。')
    order = engine_order(requested)
    expected = list(order[:order.index(engine)+1]) if engine in order else []
    attempts = data.get('attempts', [])
    if not expected or [a.get('engine') for a in attempts] != expected:
        raise RenderError('render-order-invalid', '没有按 Word → WPS → LibreOffice 顺序执行。')
    if attempts[-1].get('status') != 'passed' or any(a.get('status') not in ('failed', 'unavailable') or not a.get('reason') for a in attempts[:-1]):
        raise RenderError('render-attempt-invalid', '缺少真实失败原因或成功记录。')
    if digest(candidate) != data.get('source_sha256'):
        raise RenderError('render-stale', '渲染不属于最后保存的 DOCX。')
    report_path = Path(data.get('report_path', ''))
    if not report_path.is_file() or json.loads(report_path.read_text(encoding='utf-8-sig')) != data:
        raise RenderError('render-report-stale', '原始渲染报告缺失或发生变化。')
    pdf = Path(data.get('pdf', ''))
    if not pdf.is_file() or digest(pdf) != data.get('pdf_sha256'):
        raise RenderError('render-pdf-stale', 'PDF 缺失或被替换。')
    records = data.get('page_records', [])
    if not records or len(records) != data.get('page_count') or data.get('pages') != [p.get('path') for p in records]:
        raise RenderError('render-pages-incomplete', '渲染页面证据不完整。')
    numbers = [int(m.group(1)) if (m := re.fullmatch(r'page-(\d+)\.png', p.get('name', ''))) else -1 for p in records]
    if numbers != list(range(1, len(records)+1)):
        raise RenderError('render-pages-incomplete', '页码缺失、重复或乱序。')
    for p in records:
        path = Path(p['path'])
        if path.name != p['name'] or not path.is_file() or digest(path) != p.get('sha256'):
            raise RenderError('render-pages-stale', 'PNG 页图缺失或被替换。')
    if pages is not None and pages != records:
        raise RenderError('render-pages-mismatch', '验收清单不是渲染器本轮完整页面清单。')
    return {'renderer': engine, 'engine': data['engine'], 'attempts': attempts, 'page_count': len(records)}


def main():
    parser = argparse.ArgumentParser(description='Word → WPS → LibreOffice 渲染 DOCX，并生成最新完整页图')
    parser.add_argument('docx', type=Path)
    parser.add_argument('--out-dir', type=Path, required=True)
    parser.add_argument('--renderer', choices=('auto', *PRIORITY), default='auto')
    parser.add_argument('--soffice'); parser.add_argument('--pdftoppm')
    parser.add_argument('--dpi', type=int, default=120)
    parser.add_argument('--timeout', type=int, default=120, help='每次引擎/页面转换调用的超时秒数')
    args = parser.parse_args()
    try:
        result = render(args.docx, args.out_dir, renderer=args.renderer, soffice=args.soffice,
                        pdftoppm=args.pdftoppm, dpi=args.dpi, timeout=args.timeout)
    except Exception as exc:
        result = {'status': 'failed', 'pages': [], 'issues': [exc.as_issue() if hasattr(exc, 'as_issue') else {'code': 'render-error', 'message': str(exc)}]}
        # Always write a failed current-run report, not a stale prior PASS.
        if args.out_dir.resolve() != args.docx.resolve().parent or args.docx.name != 'render-report.json':
            try:
                write_json(args.out_dir / 'render-report.json', result)
            except OSError:
                pass
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result['status'] == 'passed' else 2


if __name__ == '__main__':
    raise SystemExit(main())
