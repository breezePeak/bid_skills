"""Reuse successful deterministic steps only with identical inputs and outputs.

No failed step, changed report, stale render, or changed checker is cached as PASS.
The cache saves subprocess work, not host visual decisions.
"""
from __future__ import annotations
import hashlib
import json
from pathlib import Path
from block_progress import sha, read_json, write_json


def code_digest(scripts):
    digest = hashlib.sha256()
    for p in sorted(p for p in Path(scripts).iterdir() if p.is_file() and p.suffix in {'.py','.ps1'}):
        digest.update(p.name.encode()); digest.update(p.read_bytes())
    return digest.hexdigest()


def invocation_key(name, args, scripts):
    values = []
    for arg in args:
        text = str(arg)
        p = Path(text)
        # Only existing input paths; output JSON/dir is not an input.
        values.append(text)
    output_flags = {'--json-out', '--out-dir', '--output', '--out'}
    inputs = []
    for i, arg in enumerate(args):
        if i and str(args[i-1]) in output_flags:
            continue
        try:
            p = Path(str(arg))
            if p.is_file():
                inputs.append([str(p.resolve()), sha(p)])
        except (OSError, ValueError):
            pass
    return {'script':str(name), 'args':values, 'inputs':inputs, 'code_sha256':code_digest(scripts)}


def cached_run(name, args, scripts, execute):
    """execute(name, *args) returns the existing run_script result."""
    strings = list(map(str,args))
    if '--json-out' in strings:
        output = Path(strings[strings.index('--json-out') + 1])
        cache = output.with_suffix(output.suffix + '.execution-cache')
    elif str(name) == 'render_docx.py' and '--out-dir' in strings:
        cache = Path(strings[strings.index('--out-dir')+1]) / 'render.execution-cache'
        output = None
    else:
        return execute(name,*args)
    key = invocation_key(name,args,scripts)
    try:
        saved = read_json(cache)
        valid = saved.get('key') == key and saved.get('run',{}).get('returncode') == 0
        if output is not None:
            valid = valid and output.is_file() and sha(output) == saved.get('output_sha256')
        elif valid:
            from render_docx import validate_render
            validate_render(Path(strings[0]), json.loads(saved['run']['stdout']))
        if valid:
            return dict(saved['run'], reused_step=True)
    except Exception:
        pass
    result = execute(name,*args)
    if result.get('returncode') == 0 and (output is None or output.is_file()):
        write_json(cache,{'key':key,'run':result,'output_sha256':sha(output) if output else None})
    return result


def prepared_valid(manifest):
    """A cached preparation still has to match every current report and page file."""
    try:
        if manifest.get('status') not in {'awaiting_visual_review','failed'}:
            return False
        candidate = Path(manifest['candidate'])
        if sha(candidate)!=manifest['candidate_sha256']:
            return False
        gates = manifest.get('gates')
        if not isinstance(gates,list) or not gates or len({g.get('id') for g in gates}) != len(gates):
            return False
        if manifest['status']=='awaiting_visual_review' and (not all(g.get('passed') is True for g in gates) or manifest.get('render',{}).get('passed') is not True):
            return False
        for gate in gates:
            if sha(gate['report'])!=gate['report_sha256']:
                return False
        render=manifest.get('render',{})
        if render.get('passed'):
            from render_docx import validate_render
            validate_render(candidate,render['data'],manifest.get('renderer_requested','auto'),render['pages'])
        return True
    except Exception:
        return False
