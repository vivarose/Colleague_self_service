import io
import json
import os
from pathlib import Path
import runpy
import tempfile
import unittest
from unittest.mock import patch


class RefreshCallbackTests(unittest.TestCase):
    def test_callbacks_identify_application_and_authenticate(self):
        script = Path(__file__).resolve().parents[1] / '.github/scripts/refresh_callback.py'
        for mode in ('claim', 'complete'):
            with self.subTest(mode=mode), tempfile.TemporaryDirectory() as directory:
                output = Path(directory) / 'output'
                environment = {
                    'REQUEST_ID': 'attempt',
                    'REFRESH_DATE': '2026-10-27',
                    'GITHUB_RUN_ID': '123',
                    'REFRESH_SERVICE_URL': 'https://worker.example/',
                    'REFRESH_CALLBACK_SECRET': 'test-secret',
                    'REFRESH_RESULT': 'succeeded',
                    'GITHUB_OUTPUT': str(output),
                }
                with patch.dict(os.environ, environment), \
                        patch('sys.argv', [str(script), mode]), \
                        patch('urllib.request.urlopen',
                              return_value=io.BytesIO(b'{"proceed":true}')) as urlopen:
                    runpy.run_path(str(script), run_name='__main__')
                request = urlopen.call_args.args[0]
                self.assertEqual(request.full_url, f'https://worker.example/{mode}')
                self.assertEqual(request.get_header('User-agent'), 'Hamilton-Catalog-Refresh')
                self.assertEqual(request.get_header('Authorization'), 'Bearer test-secret')
                self.assertEqual(urlopen.call_args.kwargs['timeout'], 60)
                body = json.loads(request.data)
                self.assertEqual(body['request_id'], 'attempt')
                self.assertEqual(body['run_id'], '123')
                if mode == 'claim':
                    self.assertEqual(output.read_text(), 'proceed=true\n')
                else:
                    self.assertEqual(body['status'], 'succeeded')
