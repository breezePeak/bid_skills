#!/usr/bin/env python3
"""对已完成第一轮检查的 DOCX 执行确定性修复并产出待视觉验收 candidate.docx。"""
from __future__ import annotations
import argparse, hashlib, json, shutil, subprocess, sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
BUNDLE = HERE.parent
DEFAULT_TEMPLATE = BUNDLE / "assets" / "default-template.docx"
DEFAULT_STYLE_JSON = BUNDLE / "assets" / "default-template-style.json"

def run(cmd, allow=(0,)):
    cp = subprocess.run(cmd, capture_output=True, text=True, check=False)
    return {"ok": cp.returncode in allow, "returncode": cp.returncode, "stdout": cp.stdout, "stderr": cp.stderr}

def run_script(name, *args, allow=(0,)):
    return run([sys.executable, str(HERE / name), *map(str, args)], allow=allow)

def load_json(path: Path):
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None

def sha256(path: Path):
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return "sha256:" + h.hexdigest()

def resolve_template(args, reports: Path):
    if args.template is None:
        return DEFAULT_TEMPLATE, DEFAULT_STYLE_JSON
    if not args.template.is_file():
        raise FileNotFoundError("template not found")
    if args.template_style_json is not None:
        if not args.template_style_json.is_file():
            raise FileNotFoundError("template style json not found")
        return args.template, args.template_style_json

    conflict = reports / "template-choice-request.json"
    r = run_script("template_conflict_detector.py", args.template, "--json-out", conflict, allow=(0, 4))
    data = load_json(conflict)
    if data and data.get("status") == "requires_user_choice":
        print(json.dumps({
            "status": "requires_user_choice",
            "choice_request": str(conflict),
            "conflicts": data.get("conflicts", []),
        }, ensure_ascii=False, indent=2))
        raise SystemExit(4)

    profile = reports / "uploaded-template.style-profile.json"
    r = run_script("template_style_profile.py", args.template, "--out", profile)
    if not r["ok"] or not profile.is_file():
        raise RuntimeError("无法解析用户模板样式")
    return args.template, profile

def main():
    ap = argparse.ArgumentParser(description="Word 排版确定性修复流程")
    ap.add_argument("input", type=Path)
    ap.add_argument("--work-dir", type=Path, required=True)
    ap.add_argument("--template", type=Path)
    ap.add_argument("--template-style-json", type=Path)
    args = ap.parse_args()

    if not args.input.is_file():
        ap.error("input not found")

    args.work_dir.mkdir(parents=True, exist_ok=True)
    reports = args.work_dir / "reports"
    render_dir = args.work_dir / "render"
    reports.mkdir(exist_ok=True)
    render_dir.mkdir(exist_ok=True)

    try:
        template, style_json = resolve_template(args, reports)
    except SystemExit:
        raise
    except Exception as e:
        print(json.dumps({"status": "failed", "error": str(e)}, ensure_ascii=False, indent=2))
        return 2

    current = args.input
    stages = []

    def run_stage(name: str, script: str, cmd_args, out: Path, report: Path):
        nonlocal_current = None  # only for readability in tracebacks
        r = run_script(script, *cmd_args)
        stages.append({"name": name, "run": r, "report": str(report)})
        if not r["ok"] or not out.is_file():
            raise RuntimeError(f"{name}失败")
        return out

    try:
        out = args.work_dir / "01-punctuation.docx"
        report = reports / "01-punctuation.json"
        current = run_stage(
            "标点修复",
            "contextual_punctuation.py",
            [current, "--out", out, "--json-out", report],
            out, report,
        )

        out = args.work_dir / "02-whitespace.docx"
        report = reports / "02-whitespace.json"
        current = run_stage(
            "空格修复",
            "whitespace_repair.py",
            [current, "--out", out, "--json-out", report],
            out, report,
        )

        out = args.work_dir / "03-template-style.docx"
        report = reports / "03-template-style.json"
        current = run_stage(
            "模板样式应用",
            "template_style_enforce.py",
            [template, style_json, current, "--out", out, "--json-out", report],
            out, report,
        )

        out = args.work_dir / "04-template-usage.docx"
        report = reports / "04-template-usage.json"
        current = run_stage(
            "直接格式污染修复",
            "template_usage_repair.py",
            [template, current, "--out", out, "--json-out", report],
            out, report,
        )

        out = args.work_dir / "05-tables.docx"
        report = reports / "05-tables.json"
        current = run_stage(
            "表格排版修复",
            "table_layout_repair.py",
            [current, "--output", out, "--json-out", report, "--template", template],
            out, report,
        )

        out = args.work_dir / "06-layout.docx"
        report = reports / "06-layout.json"
        current = run_stage(
            "通用布局策略修复",
            "docx_layout_policy.py",
            [current, "--output", out, "--json-out", report],
            out, report,
        )
    except Exception as e:
        print(json.dumps({"status": "failed", "error": str(e), "stages": stages}, ensure_ascii=False, indent=2))
        return 2

    candidate = args.work_dir / "candidate.docx"
    shutil.copy2(current, candidate)

    gates = []
    def gate(name, script, report_name, script_args, passed):
        report = reports / report_name
        r = run_script(script, *script_args(report), allow=(0, 2, 3, 4))
        data = load_json(report)
        ok = bool(passed(data))
        gates.append({"name": name, "passed": ok, "report": str(report), "result": data, "run": r})

    gate("标点", "contextual_punctuation.py", "gate-punctuation.json",
         lambda report: [candidate, "--json-out", report],
         lambda d: d is not None and d.get("issue_count", 0) == 0)
    gate("字符级文本", "hard_text_audit.py", "gate-hard-text.json",
         lambda report: [candidate, "--json-out", report],
         lambda d: d is not None and d.get("status") == "passed")
    gate("结构与版式", "docx_audit.py", "gate-docx-audit.json",
         lambda report: [candidate, "--json-out", report],
         lambda d: d is not None and not any(i.get("severity") == "error" for i in d.get("issues", [])))
    gate("模板实际使用", "template_usage_audit.py", "gate-template-usage.json",
         lambda report: [template, candidate, "--json-out", report],
         lambda d: d is not None and d.get("status") == "passed")
    gate("模板结构一致性", "template_conformance.py", "gate-template-conformance.json",
         lambda report: [template, candidate, "--json-out", report],
         lambda d: d is not None and (d.get("status") == "passed" or not d.get("errors")))

    render_run = run_script("render_docx.py", candidate, "--out-dir", render_dir, allow=(0,))
    render_data = None
    if render_run["ok"]:
        try:
            render_data = json.loads(render_run["stdout"])
        except Exception:
            render_data = None

    pages = []
    for pstr in (render_data or {}).get("pages", []):
        p = Path(pstr)
        if p.is_file():
            pages.append({"path": str(p), "name": p.name, "sha256": sha256(p)})

    structural_pass = all(g["passed"] for g in gates)
    render_pass = bool(render_run["ok"] and pages)
    status = "awaiting_visual_review" if structural_pass and render_pass else "failed"

    manifest = {
        "version": 3,
        "input": str(args.input),
        "template": str(template),
        "template_style_json": str(style_json),
        "candidate": str(candidate),
        "candidate_sha256": sha256(candidate),
        "status": status,
        "stages": stages,
        "gates": gates,
        "render": {"passed": render_pass, "run": render_run, "data": render_data, "pages": pages},
        "next": "Agent 必须对最新渲染页面执行视觉验收；发现问题则回到对应修复步骤，修改后重新运行本流程。",
    }
    manifest_path = args.work_dir / "review-manifest.json"
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({
        "status": status,
        "candidate": str(candidate),
        "manifest": str(manifest_path),
        "page_count": len(pages),
        "failed_gates": [g["name"] for g in gates if not g["passed"]],
    }, ensure_ascii=False, indent=2))
    return 0 if status == "awaiting_visual_review" else 2

if __name__ == "__main__":
    raise SystemExit(main())
