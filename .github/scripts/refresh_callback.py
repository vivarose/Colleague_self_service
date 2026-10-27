import json
import os
import sys
from urllib.request import Request, urlopen


mode = sys.argv[1]
body = {
    'request_id': os.environ['REQUEST_ID'],
    'refresh_date': os.environ['REFRESH_DATE'],
    'run_id': os.environ['GITHUB_RUN_ID'],
}
if mode == 'complete':
    body['status'] = os.environ['REFRESH_RESULT']
request = Request(
    os.environ['REFRESH_SERVICE_URL'].rstrip('/') + '/' + mode,
    data=json.dumps(body).encode(),
    headers={
        'Authorization': 'Bearer ' + os.environ['REFRESH_CALLBACK_SECRET'],
        'Content-Type': 'application/json',
    },
    method='POST',
)
with urlopen(request, timeout=60) as response:
    result = json.load(response)
if mode == 'claim':
    with open(os.environ['GITHUB_OUTPUT'], 'a') as output:
        output.write(f'proceed={str(result.get("proceed") is True).lower()}\n')
