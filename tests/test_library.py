"""Document classification, change tracking and source integration on temporary files."""
from pathlib import Path
import json
import os
import shutil
import tempfile
import threading
import unittest
from unittest.mock import patch
from urllib.request import Request, urlopen
from urllib.error import HTTPError
from urllib.parse import quote
from openpyxl import Workbook

import server
from library import Library, fingerprint


class LibraryTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.base = Path(self.temp.name)
        self.store = Library(self.base/'客户资料', self.base/'workspace.sqlite3')

    def write(self, name, text='sample'):
        path = self.store.root/name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(text, encoding='utf-8')
        return path

    def item(self, path):
        return next(r for r in self.store.snapshot()['files'] if r['path'] == str(path))

    def test_nested_packages_classify_without_moving_or_modifying_sources(self):
        paths = [self.write('Client A/管理账.txt'), self.write('Client A/固定资产折旧表.txt'),
                 self.write('Client B/租赁合同.txt'), self.write('Client B/mystery.txt')]
        before = {p: fingerprint(p) for p in paths}
        self.write('Client A/~$temp.xlsx')
        self.store.scan()
        self.assertEqual(len(self.store.snapshot()['files']), 4)
        self.assertEqual([self.item(p)['category'] for p in paths], ['management', 'assets', 'contracts', 'other'])
        self.assertEqual(self.item(paths[0])['package'], 'Client A')
        self.assertEqual(before, {p: fingerprint(p) for p in paths})

    def test_nondescript_workbook_uses_sheet_names(self):
        path = self.store.root/'附件001.xlsx'
        book = Workbook(); book.active.title = 'Fixed assets'; book.save(path); book.close()
        self.store.scan()
        self.assertEqual(self.item(path)['category'], 'assets')
        self.assertTrue(self.item(path)['can_use_tb'])

    def test_manual_classification_status_and_confirmation_survive_restart(self):
        path = self.write('附件.txt'); self.store.scan()
        self.store.update({'path': str(path), 'sha256': self.item(path)['sha256'], 'category': 'contracts', 'status': 'done'})
        self.store = Library(self.store.root, self.store.db); self.store.scan()
        item = self.item(path)
        self.assertEqual((item['category'], item['status'], item['change']), ('contracts', 'done', ''))
        self.assertTrue(item['category_confirmed'])

    def test_content_change_resets_done_and_invalidates_old_confirmation(self):
        path = self.write('Management accounts.txt'); self.store.scan()
        old = self.item(path)['sha256']
        self.store.update({'path': str(path), 'sha256': old, 'category': 'management', 'status': 'done'})
        path.write_text('new year new data', encoding='utf-8'); self.store.scan()
        item = self.item(path)
        self.assertEqual((item['status'], item['change'], item['category']), ('pending', 'updated', 'management'))
        self.assertFalse(item['category_confirmed'])
        with self.assertRaisesRegex(ValueError, '变更'):
            self.store.update({'path': str(path), 'sha256': old, 'status': 'done'})

    def test_unscanned_update_cannot_be_marked_done_using_old_screen(self):
        path = self.write('TB.txt'); self.store.scan(); old = self.item(path)['sha256']
        path.write_text('different', encoding='utf-8')
        with self.assertRaisesRegex(ValueError, '重新扫描'):
            self.store.update({'path': str(path), 'sha256': old, 'status': 'done'})

    def test_same_content_timestamp_change_preserves_status(self):
        path = self.write('TB.txt'); self.store.scan()
        self.store.update({'path': str(path), 'sha256': self.item(path)['sha256'], 'status': 'done'})
        stamp = path.stat(); os.utime(path, ns=(stamp.st_atime_ns, stamp.st_mtime_ns+1000000000))
        self.store.scan()
        self.assertEqual(self.item(path)['status'], 'done')
        self.assertEqual(self.item(path)['change'], '')

    def test_removed_and_restored_files_keep_history(self):
        path = self.write('TB.txt'); self.store.scan(); sha = self.item(path)['sha256']
        self.store.update({'path': str(path), 'sha256': sha, 'status': 'done'})
        path.unlink(); self.store.scan()
        self.assertEqual(self.item(path)['availability'], 'missing')
        self.assertEqual(self.item(path)['sha256'], sha)
        self.write('TB.txt'); self.store.scan()
        self.assertEqual((self.item(path)['availability'], self.item(path)['status']), ('present', 'done'))

    def test_ambiguous_and_unreadable_content_does_not_claim_confirmed(self):
        path = self.write('合同与固定资产.pdf', 'not a PDF'); self.store.scan()
        item = self.item(path)
        self.assertEqual(item['category'], 'other')
        self.assertTrue(item['read_warning'])
        self.assertFalse(item['category_confirmed'])

    def test_catalogue_failure_does_not_mark_every_source_removed(self):
        path = self.write('TB.txt'); self.store.scan()
        with patch('library.os.walk', side_effect=OSError('folder unavailable')):
            self.store.scan()
        self.assertEqual(self.item(path)['availability'], 'present')
        self.assertTrue(self.store.snapshot()['errors'])

    def test_path_escape_and_invalid_status_are_rejected(self):
        outside = self.base/'outside.txt'; outside.write_text('outside')
        with self.assertRaisesRegex(ValueError, '资料库'):
            self.store.resolve(str(outside))
        path = self.write('TB.txt'); self.store.scan()
        with self.assertRaisesRegex(ValueError, '状态'):
            self.store.update({'path': str(path), 'sha256': self.item(path)['sha256'], 'status': 'audited'})


class LibraryHttpTests(unittest.TestCase):
    def test_new_library_source_is_available_to_workbook_and_intake_without_inventory_edit(self):
        with tempfile.TemporaryDirectory() as folder:
            base = Path(folder); root = base/'客户资料'; root.mkdir()
            source = root/'Management accounts.xlsx'
            shutil.copy2(Path(server.APP)/'docs/examples/AuditFlow_Demo_Management.xlsx', source)
            with patch.multiple(server, DB=base/'workspace.sqlite3', LIBRARY_ROOT=root, KNOWN={}, RESEARCH=base):
                store = server.document_library(); store.scan()
                with patch.object(store, 'request_scan'):
                    http = server.ThreadingHTTPServer(('127.0.0.1',0), server.Handler)
                    thread = threading.Thread(target=http.serve_forever, daemon=True); thread.start()
                    try:
                        url = f'http://127.0.0.1:{http.server_port}'
                        def get(route):
                            with urlopen(url+route) as response: return json.load(response)
                        listing = get('/api/library')
                        self.assertEqual(listing['files'][0]['category'], 'management')
                        self.assertEqual(get('/api/files?kind=workbook')['files'][0]['path'], str(source))
                        book = get('/api/workbook?path='+quote(str(source)))
                        self.assertIn('ClosingBalances', [s['name'] for s in book['sheets']])
                        spec={'path':str(source),'sheet':'ClosingBalances','first_row':2,'last_row':4,
                              'columns':{'code':'A','name':'B','amount':'C'},'amount_mode':'signed','sign_convention':'debit_positive',
                              'source_role':'closing_balances','currency_basis':'functional','unit':1,
                              'context':{'entity':'AuditFlow Demo HK Limited','period_start':'2025-04-01','period_end':'2026-03-31','currency':'HKD'}}
                        req=Request(url+'/api/intake/preview',data=json.dumps(spec).encode(),headers={'X-Audit-Token':server.TOKEN})
                        with urlopen(req) as response: preview=json.load(response)
                        self.assertEqual(len(preview['lines']),3)
                        self.assertFalse(any(l['error'] for l in preview['lines']))
                        with self.assertRaises(HTTPError) as exc:
                            urlopen(Request(url+'/api/library/update',data=b'{}'))
                        self.assertEqual(exc.exception.code,403)
                        outside=base/'outside.xlsx'; outside.write_bytes(b'outside')
                        with self.assertRaisesRegex(ValueError,'已登记'):
                            server.source_path(str(outside))
                    finally:
                        http.shutdown(); http.server_close(); thread.join(timeout=3)


if __name__ == '__main__': unittest.main()
