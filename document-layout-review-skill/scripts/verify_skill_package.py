#!/usr/bin/env python3
"""Skill 发布自检：结构、引用、Python 语法和关键入口烟测。"""
from __future__ import annotations
import argparse, json, py_compile, re, subprocess, sys, tempfile
from pathlib import Path

REQUIRED = [
    "SKILL.md",
    "README.md",
    "assets/default-template.docx",
    "assets/default-template-style.json",
    "references/template-resolution.md",
    "references/repair-playbook.md",
    "references/acceptance-checklist.md",
    "references/table-rules.md",
    "references/flowchart-rules.md",
    "references/punctuation-rules.md",
    "references/whitespace-rules.md",
    "scripts/inspect_document.py",
    "scripts/review_pipeline.py",
    "scripts/render_docx.py",
    "scripts/docx_audit.py",
    "scripts/template_style_enforce.py",
    "scripts/table_layout_repair.py",
]

def run(cmd):
    return subprocess.run(cmd, capture_output=True, text=True, check=False)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("skill_dir", type=Path)
    ap.add_argument("--skip-smoke", action="store_true")
    args = ap.parse_args()
    root = args.skill_dir.resolve()
    errors = []

    for rel in REQUIRED:
        p = root / rel
        if not p.is_file() or p.stat().st_size == 0:
            errors.append(f"missing-or-empty:{rel}")

    skill = root / "SKILL.md"
    if skill.is_file():
        txt = skill.read_text(encoding="utf-8")
        fm = re.match(r"^---\n(.*?)\n---\n", txt, re.S)
        if not fm:
            errors.append("invalid-frontmatter")
        else:
            front = fm.group(1)
            if "name: document-layout-review" not in front:
                errors.append("wrong-skill-name")
            if "description:" not in front:
                errors.append("missing-description")
        for ref in re.findall(r"`((?:references|scripts|assets)/[^`]+)`", txt):
            if not (root / ref).exists():
                errors.append(f"broken-reference:{ref}")

    for p in (root / "scripts").glob("*.py"):
        try:
            py_compile.compile(str(p), doraise=True)
        except Exception as e:
            errors.append(f"python-compile:{p.name}:{e}")

    try:
        data = json.loads((root / "assets/default-template-style.json").read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            errors.append("default-style-json-not-object")
    except Exception as e:
        errors.append(f"default-style-json-invalid:{e}")

    smoke = {}
    if not errors and not args.skip_smoke:
        with tempfile.TemporaryDirectory(prefix="document-layout-review-skill-smoke-") as td:
            td = Path(td)
            inp = root / "assets/default-template.docx"

            cp = run([
                sys.executable, str(root / "scripts/inspect_document.py"),
                str(inp), "--work-dir", str(td / "inspect"),
            ])
            smoke["inspect_document"] = {"returncode": cp.returncode, "stdout": cp.stdout[-1000:], "stderr": cp.stderr[-1000:]}
            if cp.returncode != 0 or not (td / "inspect/inspection.json").is_file():
                errors.append("smoke-failed:inspect_document")

            cp = run([
                sys.executable, str(root / "scripts/review_pipeline.py"),
                str(inp), "--work-dir", str(td / "review"),
            ])
            smoke["review_pipeline"] = {"returncode": cp.returncode, "stdout": cp.stdout[-1500:], "stderr": cp.stderr[-1500:]}
            # return 0 means awaiting visual review; 2 is only acceptable when deterministic gates find real template issues,
            # but the workflow must still have produced candidate + manifest. Missing either is a broken pipeline.
            if not (td / "review/candidate.docx").is_file() or not (td / "review/review-manifest.json").is_file():
                errors.append("smoke-failed:review_pipeline-no-candidate-or-manifest")

    result = {
        "status": "passed" if not errors else "failed",
        "error_count": len(errors),
        "errors": errors,
        "file_count": sum(1 for p in root.rglob("*") if p.is_file()),
        "smoke": smoke,
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if not errors else 2

if __name__ == "__main__":
    raise SystemExit(main())
