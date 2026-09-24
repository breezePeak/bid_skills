#!/usr/bin/env python3
"""Run the scoped stability suite. This is not the repository's full regression suite."""
from pathlib import Path
import subprocess
import sys

ROOT=Path(__file__).resolve().parents[1]
if __name__=='__main__':
    suites=['test_block_progress.py','test_block_workflow_v2.py','test_block_final_flow.py','test_stability_revision.py']
    result=subprocess.run([sys.executable,'-m','pytest','-q',*[str(ROOT/'tests'/n) for n in suites]],cwd=ROOT,check=False)
    raise SystemExit(result.returncode)
