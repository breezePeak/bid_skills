#!/usr/bin/env python3
"""Replace one embedded media file inside a DOCX package without changing relationships.

Use after a flowchart has been semantically redrawn. The replacement must match the
original media extension unless --allow-format-change is explicitly used and the caller
also owns the related content-type update (not implemented here).
"""
from __future__ import annotations
import argparse, json, shutil, tempfile, zipfile
from pathlib import Path


def main() -> int:
    ap = argparse.ArgumentParser(description='Replace an embedded DOCX media part safely.')
    ap.add_argument('docx', type=Path)
    ap.add_argument('--media-path', required=True, help='e.g. word/media/image3.png from images-manifest.json')
    ap.add_argument('--replacement', type=Path, required=True)
    ap.add_argument('--out', type=Path, required=True)
    a = ap.parse_args()
    if not a.docx.is_file(): ap.error('DOCX does not exist')
    if not a.replacement.is_file(): ap.error('replacement does not exist')
    media = a.media_path.replace('\\','/').lstrip('/')
    if not media.startswith('word/media/'):
        ap.error('media-path must be under word/media/')
    old_ext = Path(media).suffix.lower()
    new_ext = a.replacement.suffix.lower()
    if old_ext != new_ext:
        ap.error(f'replacement extension must match original media extension: {old_ext} != {new_ext}')
    a.out.parent.mkdir(parents=True, exist_ok=True)
    replaced = False
    with zipfile.ZipFile(a.docx, 'r') as zin, zipfile.ZipFile(a.out, 'w') as zout:
        names = set(zin.namelist())
        if media not in names:
            ap.error(f'media part not found: {media}')
        for item in zin.infolist():
            data = a.replacement.read_bytes() if item.filename == media else zin.read(item.filename)
            if item.filename == media: replaced = True
            zout.writestr(item, data)
    print(json.dumps({'status':'replaced' if replaced else 'failed','source':str(a.docx),'media_path':media,'replacement':str(a.replacement),'output':str(a.out)},ensure_ascii=False,indent=2))
    return 0 if replaced else 2

if __name__ == '__main__':
    raise SystemExit(main())
