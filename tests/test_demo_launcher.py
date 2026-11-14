"""Manual demo safeguards; all storage and source files live in temporary folders."""
from contextlib import ExitStack, closing
import importlib.util
import io
import json
from pathlib import Path
import sqlite3
import sys
import tempfile
import unittest
from unittest.mock import MagicMock, patch

APP = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(APP))
import server


def load_script(name, filename):
    spec = importlib.util.spec_from_file_location(name, APP / 'docs' / filename)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


run_demo = load_script('auditflow_test_run_demo', 'run_demo.py')
start_demo = load_script('auditflow_test_start_demo', 'start_demo.py')


class DemoStorageTests(unittest.TestCase):
    def setUp(self):
        self.stack = ExitStack()
        self.addCleanup(self.stack.close)
        self.root = Path(self.stack.enter_context(tempfile.TemporaryDirectory()))
        self.docs = self.root / 'docs'
        self.examples = self.docs / 'examples'
        self.examples.mkdir(parents=True)
        for name in ['balances.xlsx', 'details.xlsx', 'working.xlsx', 'template.docx']:
            (self.examples / name).write_bytes(b'synthetic-source')
        self.production_db = self.root / 'workspace.sqlite3'
        self.stack.enter_context(patch.multiple(
            server, DB=self.production_db, OUTPUT=self.root / 'outputs',
            RESEARCH=self.root / 'research', KNOWN={}, LIBRARY_ROOT=self.root / 'production_library',
        ))
        self.stack.enter_context(patch.object(run_demo, 'DOCS', self.docs))

    def test_demo_writes_leave_production_state_and_intake_unchanged(self):
        production_state = {'project': {'client': 'real-client'}, 'bindings': [{'id': 'keep'}]}
        production_intake = {'id': 'production-tb', 'lines': [{'amount': '123'}]}
        server.save_state(production_state)
        server.save_intake(production_intake)
        before = self.production_db.read_bytes()

        run_demo.configure()
        server.save_state({'project': {'client': 'demo'}, 'bindings': []})
        server.save_intake({'id': 'demo-tb', 'lines': []})
        (server.OUTPUT / 'preview.json').write_text('{}', encoding='utf-8')

        self.assertEqual(self.production_db.read_bytes(), before)
        with closing(sqlite3.connect(self.production_db)) as connection:
            self.assertEqual(json.loads(connection.execute('SELECT payload FROM state').fetchone()[0]), production_state)
            self.assertEqual(json.loads(connection.execute('SELECT payload FROM intake').fetchone()[0]), production_intake)
        self.assertEqual(server.DB, self.docs / 'demo_workspace.sqlite3')
        self.assertEqual(server.OUTPUT, self.docs / 'demo_outputs')
        self.assertEqual(server.RESEARCH, self.docs / 'demo_reading')
        self.assertTrue(server.DB.is_file())
        self.assertTrue((self.docs / 'demo_outputs' / 'preview.json').is_file())

    def test_only_direct_example_workbooks_and_docx_are_registered(self):
        outside = self.root / 'client.xlsx'
        outside.write_bytes(b'private-client-source')
        nested = self.examples / 'nested'
        nested.mkdir()
        (nested / 'hidden.xlsx').write_bytes(b'nested-source')
        (self.examples / 'readme.txt').write_text('not an accounting source', encoding='utf-8')

        run_demo.configure()
        self.assertEqual(len(server.KNOWN), 4)
        self.assertTrue(all(Path(path).parent == self.examples.resolve() for path in server.KNOWN))
        self.assertTrue(all(Path(path).suffix in {'.xlsx', '.docx'} for path in server.KNOWN))
        self.assertEqual(server.source_path(str(self.examples / 'balances.xlsx'), ['.xlsx']), self.examples / 'balances.xlsx')
        with self.assertRaisesRegex(ValueError, '已登记'):
            server.source_path(str(outside), ['.xlsx'])
        with self.assertRaisesRegex(ValueError, '已登记'):
            server.source_path(str(nested / 'hidden.xlsx'), ['.xlsx'])

    def test_missing_example_prevents_demo_start(self):
        (self.examples / 'template.docx').unlink()
        with self.assertRaisesRegex(RuntimeError, '演示资料不完整'):
            run_demo.configure()


class DemoLauncherTests(unittest.TestCase):
    def test_free_port_is_not_ready(self):
        with patch.object(start_demo.socket, 'create_connection', side_effect=OSError('not listening')):
            with patch.object(start_demo.OPENER, 'open') as request:
                self.assertFalse(start_demo.is_ready())
                request.assert_not_called()

    def test_other_program_on_port_is_refused(self):
        for payload in [b'{"application":"other"}', b'not JSON']:
            with self.subTest(payload=payload):
                response = MagicMock()
                response.__enter__.return_value = io.BytesIO(payload)
                with patch.object(start_demo.socket, 'create_connection', return_value=MagicMock()):
                    with patch.object(start_demo.OPENER, 'open', return_value=response):
                        with self.assertRaisesRegex(RuntimeError, '其他程序占用'):
                            start_demo.is_ready()

    def test_ready_check_requires_demo_environment_marker(self):
        response = MagicMock()
        response.__enter__.return_value = io.BytesIO(b'{"environment":"auditflow_demo"}')
        with patch.object(start_demo.socket, 'create_connection', return_value=MagicMock()):
            with patch.object(start_demo.OPENER, 'open', return_value=response) as request:
                self.assertTrue(start_demo.is_ready())
                request.assert_called_once_with(start_demo.URL + '/api/session', timeout=2)

    def test_existing_demo_is_reused_without_starting_another_process(self):
        with patch.object(start_demo, 'is_ready', return_value=True) as ready:
            with patch.object(start_demo.subprocess, 'Popen') as launch:
                with patch.object(start_demo.webbrowser, 'open', return_value=True) as browser:
                    start_demo.main()
        ready.assert_called_once_with()
        launch.assert_not_called()
        browser.assert_called_once_with(start_demo.URL, new=2)


if __name__ == '__main__':
    unittest.main()
