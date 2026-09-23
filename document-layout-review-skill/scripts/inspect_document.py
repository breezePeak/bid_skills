#!/usr/bin/env python3
"""对 DOCX 做第一轮只读检查，生成结构化问题清单，不修改原文件。"""
from __future__ import annotations
import argparse, json, subprocess, sys
from pathlib import Path
from text_rules import load_rules, load_profile, expected_body_props, write_rules, rules_digest

HERE = Path(__file__).resolve().parent
BUNDLE = HERE.parent
DEFAULT_TEMPLATE = BUNDLE / "assets" / "default-template.docx"
DEFAULT_STYLE = BUNDLE / "assets" / "default-template-style.json"

def run_script(name: str, *args: object, allow=(0, 2, 3, 4)):
    cp = subprocess.run(
        [sys.executable, str(HERE / name), *map(str, args)],
        capture_output=True, text=True, check=False
    )
    return {
        "ok": cp.returncode in allow,
        "returncode": cp.returncode,
        "stdout": cp.stdout,
        "stderr": cp.stderr,
    }

def load_json(path: Path):
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None

def main():
    ap = argparse.ArgumentParser(description="Word 排版第一轮只读检查")
    ap.add_argument("input", type=Path)
    ap.add_argument("--work-dir", type=Path, required=True)
    ap.add_argument("--template", type=Path)
    ap.add_argument("--template-style-json", type=Path)
    ap.add_argument("--text-rules", type=Path, help="本次文字规则 JSON")
    args = ap.parse_args()

    if not args.input.is_file():
        ap.error("input not found")

    args.work_dir.mkdir(parents=True, exist_ok=True)
    reports = args.work_dir / "reports"
    images = args.work_dir / "images"
    reports.mkdir(exist_ok=True)
    images.mkdir(exist_ok=True)

    template = args.template or DEFAULT_TEMPLATE
    style_json = args.template_style_json or (DEFAULT_STYLE if args.template is None else None)

    if args.template is not None and style_json is None:
        conflict_path = reports / "template-conflicts.json"
        r = run_script("template_conflict_detector.py", template, "--json-out", conflict_path)
        data = load_json(conflict_path)
        if data and data.get("status") == "requires_user_choice":
            result = {
                "status": "requires_user_choice",
                "template": str(template),
                "conflicts": data.get("conflicts", []),
                "choice_request": str(conflict_path),
            }
            out = args.work_dir / "inspection.json"
            out.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
            print(json.dumps(result, ensure_ascii=False, indent=2))
            return 4
        profile = reports / "template-style-profile.json"
        pr = run_script("template_style_profile.py", template, "--out", profile, allow=(0,))
        if not pr["ok"] or not profile.is_file():
            print(json.dumps({"status": "failed", "error": "无法解析模板样式"}, ensure_ascii=False))
            return 2
        style_json = profile

    try:
        rules = load_rules(args.text_rules)
        expected_body_props(load_profile(template), rules)
        rules_file = reports / "text-rules.effective.json"
        write_rules(rules_file, rules)
        rules_hash = rules_digest(rules)
    except (ValueError, OSError) as exc:
        print(json.dumps({"status": "failed", "error": str(exc)}, ensure_ascii=False))
        return 2

    checks = []

    p = reports / "docx-audit.json"
    rr = run_script("docx_audit.py", args.input, "--json-out", p)
    checks.append({"name": "结构与版式", "report": str(p), "run": rr, "result": load_json(p)})

    p = reports / "automatic-numbering.json"
    rr = run_script("numbering_audit.py", args.input, "--json-out", p)
    data = load_json(p)
    if rr.get("returncode") not in (0, 2) or not isinstance(data, dict) or data.get("status") not in {"passed", "failed"}:
        print(json.dumps({"status": "failed", "error": "自动编号检查执行失败，不能跳过", "run": rr}, ensure_ascii=False))
        return 2
    checks.append({"name": "标题/题注/脚注自动编号", "report": str(p), "run": rr, "result": data})

    p = reports / "table-layout.json"
    rr = run_script("table_layout_audit.py", args.input, "--json-out", p)
    checks.append({"name": "表格列宽/协调性/合并候选", "report": str(p), "run": rr, "result": load_json(p)})

    p = reports / "punctuation.json"
    rr = run_script("contextual_punctuation.py", args.input, "--json-out", p)
    checks.append({"name": "中英文标点", "report": str(p), "run": rr, "result": load_json(p)})

    p = reports / "hard-text.json"
    rr = run_script("hard_text_audit.py", args.input, "--json-out", p)
    checks.append({"name": "字符级文本", "report": str(p), "run": rr, "result": load_json(p)})

    if template and template.is_file():
        p = reports / "template-usage.json"
        rr = run_script("template_usage_audit.py", template, args.input, "--json-out", p, "--text-rules", rules_file)
        data = load_json(p) or {}
        if not rr["ok"] or data.get("text_rules_sha256") != rules_hash:
            print(json.dumps({"status": "failed", "error": "文字检查未使用本次规则或执行失败", "run": rr}, ensure_ascii=False))
            return 2
        checks.append({"name": "模板实际使用", "report": str(p), "run": rr, "result": load_json(p)})

        p = reports / "template-conformance.json"
        rr = run_script("template_conformance.py", template, args.input, "--json-out", p)
        checks.append({"name": "模板结构一致性", "report": str(p), "run": rr, "result": load_json(p)})

    image_run = run_script("extract_docx_images.py", args.input, "--out-dir", images, allow=(0,))
    extracted = sorted(str(p) for p in images.glob("*") if p.is_file())

    issue_count = 0
    hard_error_count = 0
    review_count = 0
    for item in checks:
        data = item.get("result") or {}
        if isinstance(data.get("issues"), list):
            issue_count += len(data["issues"])
            for issue in data["issues"]:
                if issue.get("severity") == "error":
                    hard_error_count += 1
                elif issue.get("severity") == "review":
                    review_count += 1
        elif isinstance(data.get("issue_count"), int):
            issue_count += data["issue_count"]
        elif data.get("status") == "failed":
            issue_count += 1
            hard_error_count += 1

    result = {
        "status": "inspected",
        "input": str(args.input),
        "template": str(template) if template else None,
        "template_style_json": str(style_json) if style_json else None,
        "text_rules": str(rules_file),
        "text_rules_sha256": rules_hash,
        "issue_count": issue_count,
        "hard_error_count": hard_error_count,
        "semantic_review_count": review_count,
        "checks": checks,
        "extracted_images": extracted,
        "visual_review_required": True,
        "next": (
            "Agent 根据 inspection.json、table-layout.json 和页面/图片视觉检查形成完整问题清单后再进入修复。"
            "表格 review 级结果必须结合语义逐项处理，不能当作自动合并命令。"
        ),
    }
    out = args.work_dir / "inspection.json"
    out.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({
        "status": result["status"],
        "inspection": str(out),
        "issue_count": issue_count,
        "hard_error_count": hard_error_count,
        "semantic_review_count": review_count,
        "image_count": len(extracted),
    }, ensure_ascii=False, indent=2))
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
