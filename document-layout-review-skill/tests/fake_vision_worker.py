#!/usr/bin/env python3
"""TEST DOUBLE ONLY. Returns scripted decisions; never claims to recognize images."""
import base64
import io
import json
import sys
import time
from PIL import Image

request = json.load(sys.stdin)
mode = sys.argv[1] if len(sys.argv) > 1 else 'pass'
if mode == 'timeout':
    time.sleep(5)
if mode == 'exit':
    raise SystemExit(7)
if mode == 'invalid-json':
    print('not json')
    raise SystemExit(0)
for item in request['images']:
    with Image.open(io.BytesIO(base64.b64decode(item['data_base64'], validate=True))) as image:
        image.load()
checks = ('text_inside_bounds', 'no_overlap', 'readable', 'not_clipped', 'connections_correct', 'no_embedded_caption')
rows = [{'view_id': ident, 'observation': 'TEST DOUBLE: scripted observation, not a model finding.',
         'checks': {k: 'pass' for k in checks}} for ident in request['required_view_ids']]
response = {'_test_double': True, 'request_id': request['request_id'], 'image_type': 'diagram',
    'observation': 'TEST DOUBLE: scripted fixture verdict only.', 'semantics_preserved': True,
    'page_match': True, 'views': rows, 'findings': []}
if mode in ('fail', 'contradictory') or mode == 'second-fail' and request['role'] == 'final-independent':
    target = next(r for r in rows if r['view_id'].endswith('/edge-right'))
    if mode != 'contradictory':
        target['checks']['text_inside_bounds'] = 'fail'
    response['findings'] = [{'view_id': target['view_id'], 'check': 'text_inside_bounds',
        'description': 'TEST DOUBLE: expected right-column overflow fixture.', 'bbox': [.1, .2, .9, .8]}]
if mode == 'uncertain':
    rows[0]['checks']['readable'] = 'uncertain'
if mode == 'na':
    rows[0]['checks']['text_inside_bounds'] = 'not_applicable'
    rows[0]['not_applicable_reasons'] = {'text_inside_bounds': 'TEST DOUBLE invalid exemption'}
if mode == 'missing':
    rows.pop()
if mode == 'no-observation':
    rows[0]['observation'] = ''
if mode == 'unlocated-fail':
    rows[0]['checks']['text_inside_bounds'] = 'fail'
if mode == 'wrong-id':
    response['request_id'] = 'wrong'
if mode == 'page-mismatch':
    response['page_match'] = False
if mode == 'semantic-mismatch':
    response['semantics_preserved'] = False
if mode == 'type-disagreement' and request['role'] == 'final-independent':
    response['image_type'] = 'photo'
print(json.dumps(response))
