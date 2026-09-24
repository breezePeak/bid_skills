#!/usr/bin/env python3
"""Agent-facing block workflow; existing auditors run once at finalization.

No language-model scheduling is hidden here: the host inspects/repairs the current
block, checkpoints it, and continues without asking the user per block.
"""
from __future__ import annotations
import argparse
import copy
import json
import os
import shutil
import sys
import io
import hashlib
from contextlib import redirect_stdout
from pathlib import Path

from block_progress import (BlockError, STATE, Q, initialize, load, next_block,
    block_by_id, accept_global, accept_shared_repair, checkpoint, reopen, sha, read_json, write_json, session_lock)

HERE = Path(__file__).resolve().parent


def settings_from_args(args, reports):
    from review_pipeline import resolve_template
    from text_rules import load_rules, write_rules
    template, profile = resolve_template(args, reports)
    baseline = reports.parent / 'baseline'
    baseline.mkdir(exist_ok=True)
    paths = {}
    for name, src in [('template', template), ('template_style_json', profile)]:
        dst = baseline / ('template.docx' if name == 'template' else 'template-style.json')
        if dst.resolve() != Path(src).resolve():
            shutil.copyfile(src, dst)
        paths[name] = str(dst.resolve())
    rules = baseline / 'text-rules.effective.json'
    write_rules(rules, load_rules(args.text_rules))
    paths['text_rules'] = str(rules.resolve())
    settings = {**paths, 'baselines': {k: {'path': p, 'sha256': sha(p)} for k, p in paths.items()},
                'renderer': args.renderer or 'auto', 'field_engine': args.field_engine or 'auto',
                'uno_python': args.uno_python,
                'vision_worker_config': str(args.vision_worker_config) if args.vision_worker_config else None}
    return settings


def _initialize_objects_unlocked(work, state):
    from numbering_policy import Doc, inventory, public_inventory
    from figure_review import make_review_template
    from content_integrity import make_plan
    reports = work / 'reports'
    doc = Doc(Path(state['source']))
    body = doc.document.find(Q('body'))
    objects = inventory(doc)
    for block in state['blocks']:
        block['figures'] = []
        block['objects'] = []
        for obj in objects:
            if block.get('part', 'word/document.xml') != 'word/document.xml':
                continue
            owner = obj['element']
            while owner.getparent() is not None and owner.getparent() is not body:
                owner = owner.getparent()
            if owner.getparent() is body and block['source_start'] <= body.index(owner) < block['source_end']:
                block['objects'].append(obj['id'])
                if obj['kind'] == 'figure':
                    block['figures'].append(obj['id'])
    initial = reports / 'initial-visual-review.json'
    # Inventory only. No model call and no full-document image review at startup.
    write_json(initial, make_review_template(Path(state['source']), work / 'source-images'))
    content = reports / 'content-plan.json'
    write_json(content, make_plan(Path(state['source']), None))
    write_json(reports / 'object-inventory.json', {'objects': public_inventory(doc)})
    state['initial_review'] = str(initial.resolve())
    state['content_plan'] = str(content.resolve())
    from figure_inspection import ledger
    state['discovery_ledger'] = str((reports / 'image-discoveries.json').resolve())
    ledger(Path(state['source']), state['discovery_ledger'])
    state['objects_initialized'] = True
    write_json(work / STATE, state)


def initialize_objects(work,state):
    with session_lock(work):
        current=load(work)
        if not current.get('objects_initialized'):
            _initialize_objects_unlocked(work,current)


def require_initial_for_block(state, block):
    if not block.get('figures'):
        return
    from numbering_policy import Doc, inventory
    from figure_inspection import verify_row
    rows = {r['id']: r for r in read_json(state['initial_review'])['objects']}
    doc = Doc(Path(state['source']))
    objects = {o['id']: o for o in inventory(doc)}
    for ident in block.get('figures', []):
        # A FAIL is a known defect to fix; uncertain/missing execution cannot pass.
        verify_row(Path(state['source']), Path(state['source']), objects[ident], rows[ident], doc, phase='initial')


def require_local_image_change(state, block, proposal):
    """Do not advance a block with a registered, still-identical intrinsic defect."""
    if not block.get('figures'):
        return
    from numbering_policy import Doc,inventory
    from visual_evidence import INTRINSIC,validate_inspection,validate_bundle,checked_file,normalized_image,pixel_sha
    rows={r['id']:r for r in read_json(state['initial_review'])['objects']}
    doc=Doc(Path(proposal));objects={o['id']:o for o in inventory(doc)}
    for ident in block['figures']:
        row=rows[ident];obj=objects.get(ident)
        if obj is None:
            raise BlockError('修复删除了原图片对象，不能推进。')
        failed={k for k,v in row.get('checks',{}).items() if v=='fail'} & INTRINSIC
        if not failed:
            continue
        if obj['hash']==row['object_sha256']:
            raise BlockError('当前块图内缺陷尚未实际修复：'+ident)
        if obj.get('paths') and not row.get('rendered_view'):
            proof=validate_inspection(row['inspection'],phase='initial')
            original=validate_bundle(proof['bundle'])['assets']
            before=[pixel_sha(normalized_image(checked_file(a))) for a in original]
            after=[pixel_sha(normalized_image(io.BytesIO(doc.files[p]))) for p in obj['paths']]
            if before==after:
                raise BlockError('图片仅重新编码，图内缺陷并未改变：'+ident)


def inspect_current_images(work, state, block, config):
    ids = block.get('figures', [])
    if not ids:
        return {'status': 'not_applicable', 'block': block['id'], 'figure_count': 0}
    if config is None:
        raise BlockError('当前块含图片；请提供本次已授权的视觉接口配置，不能手填 PASS。')
    from figure_inspection import inspect_initial
    result = inspect_initial(Path(state['source']), state['initial_review'], config,
                             work / 'initial-inspections', object_ids=ids)
    write_json(state['initial_review'], result)
    require_initial_for_block(state, block)
    return {'status': 'inspected', 'block': block['id'], 'figures': ids, 'review': state['initial_review']}


def reuse_unchanged_reviews(previous, template, manifest):
    """Reuse only exact current page evidence; changed pages + neighbours stay pending."""
    if not previous or not Path(previous.get('review', '')).is_file():
        return template
    old = read_json(previous['review'])
    if not Path(previous.get('manifest','')).is_file():
        return template
    old_m = read_json(previous['manifest'])
    for key in ('source_sha256','template_sha256','template_style_sha256','text_rules_sha256','review_policy_version'):
        if old_m.get(key) != manifest.get(key):
            return template
    if old.get('candidate_sha256') != previous.get('candidate_sha256') or old_m.get('candidate_sha256') != previous.get('candidate_sha256'):
        return template
    # Compare actual image files, not hand-maintained page labels.
    previous_pages = {p['name']: p for p in old_m['render']['pages']}
    current_pages = {p['name']: p for p in manifest['render']['pages']}
    prior_rows = {p['name']: p for p in old.get('pages', [])}
    names = [p['name'] for p in manifest['render']['pages']]
    unchanged = set()
    for name, p in current_pages.items():
        prior = previous_pages.get(name)
        if prior and prior['sha256'] == p['sha256'] and Path(prior['path']).is_file() and sha(prior['path']) == prior['sha256'] and sha(p['path']) == p['sha256']:
            unchanged.add(name)
    changed = {i for i, name in enumerate(names) if name not in unchanged}
    boundary = changed | {j for i in changed for j in (i-1, i+1) if 0 <= j < len(names)}
    if len(previous_pages) != len(current_pages) and names:
        boundary.add(len(names) - 1)
    for i, row in enumerate(template['pages']):
        old_row = prior_rows.get(row['name'], {})
        if i not in boundary and row['name'] in unchanged and old_row.get('status') == 'pass' and old_row.get('sha256') == row['sha256'] and str(old_row.get('observations', '')).strip():
            row.update(status='pass', observations=old_row['observations'],
                       reused_from={'path': previous['review'], 'sha256': sha(previous['review'])})
    previous_figures = {r['id']: r for r in old.get('figures', {}).get('objects', [])}
    for i, row in enumerate(template['figures']['objects']):
        prior = previous_figures.get(row['id'], {})
        if (prior.get('inspection') and not prior.get('rendered_view') and prior.get('object_sha256') == row['object_sha256'] and
                prior.get('source_object_sha256') == row['source_object_sha256'] and prior.get('pages') and
                all(p['name'] in unchanged and p['sha256'] == current_pages[p['name']]['sha256'] for p in prior['pages'])):
            reuse = copy.deepcopy(prior)
            reuse['reuse'] = {'previous_candidate_sha256': (prior.get('reuse') or {}).get('previous_candidate_sha256', previous['candidate_sha256']), 'reason': 'unchanged-image-and-page-hashes'}
            template['figures']['objects'][i] = reuse
    previous_tables = {r['id']:r for r in old.get('table_objects',[])}
    for i,row in enumerate(template.get('table_objects',[])):
        prior = previous_tables.get(row['id'],{})
        if prior.get('status') == 'pass' and prior.get('object_sha256') == row['object_sha256'] and prior.get('pages') and all(r['name'] in unchanged and r['sha256'] == current_pages[r['name']]['sha256'] for r in prior['pages']):
            template['table_objects'][i] = copy.deepcopy(prior)
    return template


def persist_plans(work, state, args):
    """Keep plans across resume, always bound to the original, never the candidate."""
    from numbering_policy import Doc, load_object_plan
    from content_integrity import make_plan, validate_movements
    changed = False
    if args.object_plan:
        plan = read_json(args.object_plan)
        load_object_plan(plan, Doc(Path(state['source'])))
        if state.get('object_plan'):
            previous = read_json(state['object_plan'])
            load_object_plan(previous,Doc(Path(state['source'])))
            merged = {r['id']:r for r in previous.get('objects',[])}
            merged.update({r['id']:r for r in plan.get('objects',[])})
            plan['objects'] = list(merged.values())
        plan_hash = hashlib.sha256(json.dumps(plan,sort_keys=True,ensure_ascii=False).encode()).hexdigest()
        dest = work / 'plans' / ('objects-' + plan_hash[:16] + '.json')
        write_json(dest,plan)
        if str(dest) != state.get('object_plan'):
            state['object_plan'] = str(dest)
            content = read_json(state['content_plan'])
            content['allowed_caption_insertions'] = make_plan(Path(state['source']),plan)['allowed_caption_insertions']
            write_json(state['content_plan'],content)
            changed = True
    if args.content_plan:
        plan = read_json(args.content_plan)
        if plan.get('version') != 1 or plan.get('source_sha256') != state['source_sha256']:
            raise BlockError('内容计划必须绑定最初原文，不得以候选替换基准。')
        previous = read_json(state['content_plan'])
        captions = {r['id']:r for r in previous.get('allowed_caption_insertions',[])}
        captions.update({r['id']:r for r in plan.get('allowed_caption_insertions',[])})
        plan['allowed_caption_insertions'] = list(captions.values())
        movement_key = lambda r:(r['reference_paragraph'],r['start'],r['end'])
        movements = {movement_key(r):r for r in previous.get('allowed_footnote_moves',[])}
        movements.update({movement_key(r):r for r in plan.get('allowed_footnote_moves',[])})
        if movements:
            plan['allowed_footnote_moves'] = list(movements.values())
        validate_movements(Path(state['source']), plan.get('allowed_footnote_moves',[]))
        if plan != previous:
            write_json(state['content_plan'],plan); changed = True
    if changed:
        if state.get('final'):
            state['previous_final'] = state['final']
        state['final'] = None
        state.pop('final_attempt',None)
        write_json(work / STATE,state)
    return state


def audit_fingerprint(state):
    from block_cache import code_digest
    refs = {key:sha(state[key]) for key in ('initial_review','content_plan','object_plan') if state.get(key)}
    data = {'revision':state['revision'],'current_sha256':state['current_sha256'],
            'source_sha256':state['source_sha256'],'settings':state['settings'],
            'plans':refs,'code':code_digest(HERE)}
    return hashlib.sha256(json.dumps(data,sort_keys=True,ensure_ascii=False).encode()).hexdigest()


def final_audit(work, state, args):
    if not state['global_ready'] or any(b['status'] != 'checked' for b in state['blocks']):
        raise BlockError('还有未完成块；不能提前执行全文终检。')
    fingerprint = audit_fingerprint(state)
    saved = state.get('final')
    if saved and saved.get('fingerprint') == fingerprint and not args.retry_final:
        try:
            prior = read_json(saved['manifest'])
            if (sha(saved['candidate']) == saved['candidate_sha256'] and
                    prior.get('candidate_sha256') == saved['candidate_sha256']):
                # Validate stored report and render bindings, without rerunning models/auditors.
                from block_cache import prepared_valid
                if prepared_valid(prior):
                    return {'status':prior['status'],**saved,'reused_prepared_audit':True}, 0 if prior['status']=='awaiting_visual_review' else 2
        except (OSError, ValueError, KeyError, TypeError):
            pass
    from figure_review import validate_initial
    from field_refresh import refresh, validate_report
    from review_pipeline import _audit_only_main
    from figure_inspection import ledger
    validate_initial(Path(state['source']),state['initial_review'])
    previous = saved or state.get('previous_final')
    attempt = state.get('final_attempt')
    if not attempt or attempt.get('fingerprint') != fingerprint:
        pass_root = work / 'final'; pass_root.mkdir(exist_ok=True)
        number = 1
        while (pass_root / f'pass-{number:03d}').exists():
            number += 1
        round_dir = pass_root / f'pass-{number:03d}'; round_dir.mkdir()
        attempt = {'fingerprint':fingerprint,'path':str(round_dir),'status':'running'}
        state['final_attempt'] = attempt; write_json(work / STATE,state)
    round_dir = Path(attempt['path'])
    settings = state['settings']
    refreshed = round_dir / 'field-updated.docx'; field_report = round_dir / 'field-update.json'
    valid_field = False
    if refreshed.is_file() and field_report.is_file():
        try:
            field = validate_report(refreshed,field_report)
            valid_field = field.get('source_sha256') == state['current_sha256']
        except Exception:
            pass
    if not valid_field:
        field = refresh(Path(state['current']),refreshed,settings['field_engine'],settings.get('uno_python'),
                        preflight=(state.get('runtime_preflight') or {}).get('field_engine'))
        write_json(field_report,field)
    object_plan = None
    if state.get('object_plan'):
        from numbering_policy import Doc, inventory
        plan = read_json(state['object_plan'])
        old = {o['id']:o for o in inventory(Doc(Path(state['source'])))}
        new = {o['id']:o for o in inventory(Doc(refreshed))}
        if set(old) != set(new):
            raise BlockError('图表对象对应关系发生变化，不能盲目复用题注/豁免计划。')
        effective = copy.deepcopy(plan); effective['source_sha256'] = sha(refreshed)
        for row in effective.get('objects',[]):
            if row['id'] not in new:
                raise BlockError('语义计划对象已失效。')
            row['object_sha256'] = new[row['id']]['hash']
            row.pop('caption_paragraph',None); row.pop('caption_text',None)
        object_plan = round_dir / 'object-plan.effective.json'; write_json(object_plan,effective)
    argv = ['review_pipeline.py',str(refreshed),'--work-dir',str(round_dir),'--audit-only',
            '--source',state['source'],'--template',settings['template'],
            '--template-style-json',settings['template_style_json'],'--text-rules',settings['text_rules'],
            '--initial-visual-review',state['initial_review'],'--content-plan',state['content_plan'],
            '--field-update-report',str(field_report),'--renderer',settings['renderer'],'--field-engine',settings['field_engine']]
    if object_plan:
        argv += ['--object-plan',str(object_plan)]
    if settings.get('uno_python'):
        argv += ['--uno-python',settings['uno_python']]
    old_argv = sys.argv; captured = io.StringIO()
    try:
        sys.argv = argv
        with redirect_stdout(captured):
            code = _audit_only_main()
    finally:
        sys.argv = old_argv
        (round_dir/'audit-stdout.txt').write_text(captured.getvalue(),encoding='utf-8')
    manifest_path = round_dir/'review-manifest.json'; manifest = read_json(manifest_path)
    if not manifest.get('candidate_sha256'):
        return {'status':'failed','manifest':str(manifest_path),'details':manifest},code or 2
    # A single source-bound ledger survives all final rounds and local reopens.
    discovery_path = state.get('discovery_ledger') or str(work/'reports'/'image-discoveries.json')
    manifest['image_discovery_ledger'] = ledger(Path(state['source']),discovery_path)
    manifest['review_policy_version'] = 'block-v2'
    review_path = round_dir/'final-visual-review.json'
    template_path = round_dir/'reports'/'final-visual-review.template.json'
    if template_path.is_file():
        template = read_json(template_path)
        from numbering_policy import Doc,inventory
        template['table_objects'] = [{'id':o['id'],'object_sha256':o['hash'],'status':'pending','observations':'','pages':[]}
                                     for o in inventory(Doc(Path(manifest['candidate']))) if o['kind']=='table']
        # Preserve manual page/table observations when retrying this same prepared candidate.
        reusable = previous
        if review_path.is_file() and read_json(review_path).get('candidate_sha256')==manifest['candidate_sha256']:
            reusable = {'candidate_sha256':manifest['candidate_sha256'],'review':str(review_path),'manifest':str(manifest_path)}
        template = reuse_unchanged_reviews(reusable,template,manifest)
        write_json(review_path,template)
    state['previous_final'] = previous
    state['final'] = {'revision':state['revision'],'fingerprint':fingerprint,
                      'candidate_sha256':manifest['candidate_sha256'],'candidate':manifest['candidate'],
                      'manifest':str(manifest_path),'review':str(review_path)}
    state['final_attempt']['status'] = manifest['status']
    write_json(work/STATE,state)
    manifest['workflow'] = 'block-v1'
    manifest['block_session'] = {'path':str((work/STATE).resolve()),'sha256':sha(work/STATE)}
    manifest['next'] = '处理本轮失败项；页面按一两页窗口查看。正常图复用图内检查并核对当前页面；返修只 reopen 指定块。'
    write_json(manifest_path,manifest)
    return {'status':manifest['status'],**state['final'],
            'failed_gates':[g['id'] for g in manifest.get('gates',[]) if not g['passed']]},code


def main(argv=None):
    ap = argparse.ArgumentParser(description='按完整内容块修复；最终统一更新域、审计和渲染。')
    ap.add_argument('input', type=Path, nargs='?')
    ap.add_argument('--work-dir', type=Path, required=True)
    ap.add_argument('--action', choices=['next', 'global', 'checkpoint', 'reopen', 'inspect-initial', 'final', 'inspect-final', 'shared', 'configure'], default='next')
    ap.add_argument('--block'); ap.add_argument('--proposal', type=Path); ap.add_argument('--note', default='')
    ap.add_argument('--template', type=Path); ap.add_argument('--template-style-json', type=Path)
    ap.add_argument('--text-rules', type=Path); ap.add_argument('--object-plan', type=Path); ap.add_argument('--content-plan', type=Path)
    ap.add_argument('--renderer', choices=['auto', 'word', 'wps', 'libreoffice'])
    ap.add_argument('--field-engine', choices=['auto', 'word', 'wps', 'libreoffice'])
    ap.add_argument('--uno-python')
    ap.add_argument('--vision-worker-config', type=Path, default=os.environ.get('DLR_VISION_WORKER_CONFIG'))
    ap.add_argument('--view-offset', type=int, default=0, help='长表/大块的查看窗口起点，不改变修复块')
    ap.add_argument('--view-size', type=int, default=14)
    ap.add_argument('--target-chars', type=int, default=1800)
    ap.add_argument('--max-paragraphs', type=int, default=14)
    ap.add_argument('--retry-final', action='store_true', help='执行条件已修复时显式重试终检；不重做块')
    ap.add_argument('--audit-only', action='store_true', help='兼容参数：转最终审计，不重新修复全文')
    args = ap.parse_args(argv)
    work = args.work_dir.resolve()
    work.mkdir(parents=True, exist_ok=True)
    try:
        if not (work / STATE).exists():
            if args.input is None or args.action != 'next':
                raise BlockError('首次运行请提供 input.docx，先生成分块进度。')
            reports = work / 'reports'; reports.mkdir(exist_ok=True)
            settings = settings_from_args(args, reports)
            from runtime_preflight import check as runtime_check
            preflight = runtime_check(args.input, Path(settings['template']), work, requested=settings['field_engine'],
                                      profile=Path(settings['template_style_json']), audit_only=False,
                                      uno_python=settings.get('uno_python'), renderer=settings['renderer'])
            write_json(reports / 'runtime-preflight.json', preflight)
            state = initialize(args.input, work, settings, target_chars=args.target_chars, max_paragraphs=args.max_paragraphs)
            state['runtime_preflight'] = preflight
            write_json(work / STATE, state)
        state = load(work)
        if args.input is not None and sha(args.input) != state['input_sha256']:
            raise BlockError('此工作目录属于另一份原文；不能静默更换输入。')
        if not state.get('objects_initialized'):
            initialize_objects(work, state)
            state = load(work)
        if args.object_plan or args.content_plan:
            with session_lock(work):
                state = persist_plans(work,load(work),args)
        # Resumes inherit pinned rules. An explicit different request is not silently ignored.
        for key in ('renderer', 'field_engine', 'uno_python'):
            value = getattr(args, key)
            if value is not None and value != state['settings'].get(key) and args.action != 'configure':
                raise BlockError(f'续跑 {key} 与本轮设置不同；不能静默更换执行条件。')
        for key in ('template', 'template_style_json', 'text_rules'):
            value = getattr(args, key)
            if value is not None:
                if key == 'text_rules':
                    from text_rules import load_rules
                    equal = load_rules(value) == load_rules(state['settings'][key])
                else:
                    equal = sha(value) == state['settings']['baselines'][key]['sha256']
                if not equal:
                    raise BlockError(f'续跑 {key} 与锁定基准不同。')
        config = args.vision_worker_config or state['settings'].get('vision_worker_config')
        code = 0
        action = 'final' if args.audit_only else args.action
        if action == 'next':
            result = next_block(work, args.view_offset, args.view_size)
        elif action == 'configure':
            with session_lock(work):
                state = load(work)
                for key in ('renderer','field_engine','uno_python'):
                    value = getattr(args,key)
                    if value is not None:
                        state['settings'][key] = value
                if args.vision_worker_config:
                    state['settings']['vision_worker_config'] = str(args.vision_worker_config.resolve())
                state['runtime_preflight'] = None
                if state.get('final'):
                    state['previous_final'] = state['final']
                state['final'] = None; state.pop('final_attempt',None)
                write_json(work / STATE,state)
            result = next_block(work)
        elif action == 'shared':
            accept_shared_repair(work,args.proposal or state['current'],args.note)
            result = next_block(work)
        elif action == 'global':
            accept_global(work, args.proposal or state['current'], args.note)
            result = next_block(work)
        elif action in {'checkpoint', 'reopen', 'inspect-initial'}:
            ident = args.block or (next_block(work).get('block') or {}).get('id')
            block = block_by_id(state, ident)
            if action == 'checkpoint':
                require_initial_for_block(state, block)
                require_local_image_change(state,block,args.proposal or state['current'])
                checkpoint(work, ident, args.proposal or state['current'], args.note)
                result = next_block(work)
            elif action == 'reopen':
                reopen(work, ident, args.note)
                result = next_block(work)
            else:
                with session_lock(work):
                    result = inspect_current_images(work, load(work), block, config)
        elif action == 'final':
            with session_lock(work):
                result, code = final_audit(work, load(work), args)
        else:
            from figure_inspection import inspect_final
            final = state.get('final')
            if not final:
                raise BlockError('先运行最终审计。')
            visual = read_json(final['review'])
            with session_lock(work):
                result = inspect_final(Path(state['source']), Path(final['candidate']), state['initial_review'],
                                       final['review'], final['manifest'], config, work / 'final-inspections')
            manifest = read_json(final['manifest'])
            expected_pages = {p['name']: p['sha256'] for p in manifest['render']['pages']}
            pages = result.get('pages', [])
            pages_ready = (len(pages) == len(expected_pages) and {p.get('name') for p in pages} == set(expected_pages) and
                           all(p.get('sha256') == expected_pages.get(p.get('name')) and p.get('status') == 'pass' and str(p.get('observations', '')).strip() for p in pages))
            tables_ready = all(r.get('decision') in {'keep_separate', 'acceptable'} and str(r.get('reason', '')).strip() for r in result.get('table_reviews', []))
            from block_evidence import validate_table_coverage
            try:
                validate_table_coverage(Path(final['candidate']), result, manifest['render']['pages'])
            except Exception:
                tables_ready = False
            if result.get('figure_inspection_status') == 'pass' and manifest.get('status') == 'awaiting_visual_review' and pages_ready and tables_ready:
                result['overall_status'] = 'pass'  # Summarize existing observations, never manufacture them.
            write_json(final['review'], result)
            code = 2 if result.get('figure_inspection_status') == 'fail' else 0
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return code or 0
    except Exception as exc:
        print(json.dumps({'status': 'blocked', 'message': str(exc)}, ensure_ascii=False, indent=2))
        return 2


if __name__ == '__main__':
    raise SystemExit(main())
