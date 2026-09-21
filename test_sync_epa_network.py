import unittest
from unittest.mock import Mock, patch

import requests
from sync_epa import graph_request


class NetworkTests(unittest.TestCase):
    @patch('sync_epa.time.sleep')
    @patch('requests.request')
    def test_timeout_then_success(self, request, sleep):
        ok = Mock(status_code=200)
        request.side_effect = [requests.ReadTimeout('signed-secret-url'), ok]
        self.assertIs(graph_request('GET', 'https://example.test', stage='Descarga'), ok)
        self.assertEqual(request.call_count, 2)
        sleep.assert_called_once_with(2)

    @patch('sync_epa.time.sleep')
    @patch('requests.request')
    def test_exhausted_timeout_is_sanitized(self, request, sleep):
        request.side_effect = requests.ReadTimeout('signed-secret-url')
        with self.assertRaisesRegex(ValueError, '^Descarga: ReadTimeout tras 4 intentos.$'):
            graph_request('GET', 'https://example.test', stage='Descarga')
        self.assertEqual(request.call_count, 4)

    @patch('sync_epa.time.sleep')
    @patch('requests.request')
    def test_throttling_and_server_error(self, request, sleep):
        limited = Mock(status_code=429, headers={'Retry-After': '10'})
        failed = Mock(status_code=503, headers={})
        ok = Mock(status_code=200)
        request.side_effect = [limited, failed, ok]
        self.assertIs(graph_request('GET', 'https://example.test', stage='Listado'), ok)
        self.assertEqual(sleep.call_args_list[0].args, (10,))
        limited.close.assert_called_once()
        failed.close.assert_called_once()

    @patch('sync_epa.time.sleep')
    @patch('requests.request')
    def test_permission_error_does_not_retry(self, request, sleep):
        request.return_value = Mock(status_code=403, headers={})
        with self.assertRaisesRegex(ValueError, 'HTTP 403'):
            graph_request('GET', 'https://example.test', stage='Listado')
        request.assert_called_once()
        sleep.assert_not_called()


if __name__ == '__main__':
    unittest.main()
