#!/usr/bin/env python3
"""Optional Chat-Completions-compatible vision adapter; no default remote endpoint.

Set DLR_VISION_API_URL (full endpoint), DLR_VISION_MODEL, DLR_VISION_API_KEY.
Only point these at the user's already approved provider. A host subagent adapter
can replace this entire command while retaining the same stdin/stdout protocol.
"""
from __future__ import annotations
import json
import os
import sys
import urllib.error
import urllib.parse
import urllib.request


class NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        raise ValueError('Vision endpoint redirects are disabled.')


def main():
    url = os.environ.get('DLR_VISION_API_URL', '')
    model = os.environ.get('DLR_VISION_MODEL', '')
    key = os.environ.get('DLR_VISION_API_KEY', '')
    parsed = urllib.parse.urlparse(url)
    local = parsed.hostname in {'localhost', '127.0.0.1', '::1'}
    if not url or not model or (parsed.scheme != 'https' and not (parsed.scheme == 'http' and local)):
        raise ValueError('Configure the full approved HTTPS DLR_VISION_API_URL and DLR_VISION_MODEL; no provider is selected automatically.')
    if parsed.username or parsed.password:
        raise ValueError('Do not place credentials in the URL.')
    task = json.load(sys.stdin)
    metadata = {k: v for k, v in task.items() if k != 'images'}
    content = [{'type': 'text', 'text': json.dumps(metadata, ensure_ascii=False)}]
    for image in task['images']:
        content.append({'type': 'text', 'text': 'Image/view id: ' + image['id']})
        content.append({'type': 'image_url', 'image_url': {'url': 'data:image/png;base64,' + image['data_base64'], 'detail': 'high'}})
    payload = {'model': model, 'messages': [{'role': 'user', 'content': content}]}
    max_tokens = os.environ.get('DLR_VISION_MAX_TOKENS')
    if max_tokens:
        payload['max_completion_tokens'] = int(max_tokens)
    request = urllib.request.Request(url, data=json.dumps(payload).encode('utf-8'),
        headers={'Content-Type': 'application/json', **({'Authorization': 'Bearer ' + key} if key else {})}, method='POST')
    opener = urllib.request.build_opener(NoRedirect())
    try:
        with opener.open(request, timeout=150) as response:
            data = json.load(response)
    except urllib.error.HTTPError as exc:
        raise ValueError(f'Vision provider HTTP {exc.code}; no raw response/credential is logged.') from exc
    choice = data['choices'][0]
    if choice.get('finish_reason') not in (None, 'stop'):
        raise ValueError('Incomplete/refused vision response: ' + str(choice.get('finish_reason')))
    text = choice['message']['content']
    if not isinstance(text, str):
        raise ValueError('Vision provider returned no text JSON.')
    text = text.strip()
    if text.startswith('```') and text.endswith('```'):
        text = '\n'.join(text.splitlines()[1:-1])
    result = json.loads(text)
    print(json.dumps(result, ensure_ascii=False))
    return 0


if __name__ == '__main__':
    try:
        raise SystemExit(main())
    except Exception as exc:
        print('Vision worker failed: ' + str(exc), file=sys.stderr)
        raise SystemExit(2)
