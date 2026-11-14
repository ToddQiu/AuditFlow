"""Local HTTP boundary and export status tests; no customer inputs are edited."""
import json
from pathlib import Path
import sys, tempfile, threading, unittest
from urllib.request import Request, urlopen
from urllib.error import HTTPError
from urllib.parse import quote
from unittest.mock import patch
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
import server

class HttpTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.http=server.ThreadingHTTPServer(('127.0.0.1',0),server.Handler)
        cls.thread=threading.Thread(target=cls.http.serve_forever,daemon=True)
        cls.thread.start(); cls.base=f'http://127.0.0.1:{cls.http.server_port}'
    @classmethod
    def tearDownClass(cls):
        cls.http.shutdown();cls.http.server_close();cls.thread.join(timeout=3)
    def test_unicode_download_filename_is_decoded_without_path_escape(self):
        with tempfile.TemporaryDirectory(dir=Path(__file__).parent) as directory:
            target=Path(directory).resolve()
            self.assertTrue(target.is_relative_to(Path(__file__).parent.resolve()))
            name='填列清单_测试.csv'; (target/name).write_bytes(b'year,amount\n2026,100\n')
            with patch.object(server,'OUTPUT',target):
                self.assertEqual(urlopen(self.base+'/download/'+quote(name)).read(),b'year,amount\n2026,100\n')
                with self.assertRaises(HTTPError) as context:
                    urlopen(self.base+'/download/'+quote('..\\server.py'))
                self.assertEqual(context.exception.code,400)
    def test_json_source_record_download_preserves_file_bytes(self):
        with tempfile.TemporaryDirectory(dir=Path(__file__).parent) as directory:
            target=Path(directory).resolve()
            name='来源与明细_测试.json'
            expected=json.dumps({'entity':'AuditFlow 示例','amount':'1000.25','details':[{'cell':'C2'}]},ensure_ascii=False,indent=2).encode('utf-8')
            (target/name).write_bytes(expected)
            with patch.object(server,'OUTPUT',target):
                response=urlopen(self.base+'/download/'+quote(name))
                downloaded=response.read()
                self.assertEqual(downloaded,expected)
                self.assertEqual(json.loads(downloaded)['amount'],'1000.25')
                self.assertIn('application/json',response.headers.get('Content-Type',''))
    def test_nonlocal_host_rejected_before_sources_are_read(self):
        with patch.object(server,'state') as state:
            with self.assertRaises(HTTPError) as context:
                urlopen(Request(self.base+'/api/session',headers={'Host':'external.example'}))
            self.assertEqual(context.exception.code,403);state.assert_not_called()
    def test_writer_rejection_cannot_be_reported_as_a_successful_download(self):
        project={'current_year':2026,'comparative_year':2025,'currency':'HKD','unit':1,
                 'period_start':'2025-04-01','period_end':'2026-03-31',
                 'source_path':'fixture.xlsx','template_path':'fixture.docx','source_confirmed':True}
        with patch.object(server,'state',return_value={'project':project,'bindings':[{}]}), \
             patch.object(server,'confirmed_source'), \
             patch.object(server,'binding_preview',return_value={'changes':[],'rejects':[]}), \
             patch.object(server,'source_path',return_value=Path('fixture.docx')), \
             patch('templates.export_docx',return_value={'written':False,'rejects':[{'reason':'changed-template'}]}):
            request=Request(self.base+'/api/export',data=b'{}',headers={'Content-Type':'application/json','X-Audit-Token':server.TOKEN})
            with self.assertRaises(HTTPError) as context:urlopen(request)
            self.assertEqual(context.exception.code,400)
            self.assertIn('changed-template',json.load(context.exception)['error'])
    def test_missing_token_cannot_modify_project(self):
        with patch.object(server,'save_state') as save:
            with self.assertRaises(HTTPError) as context:urlopen(Request(self.base+'/api/project',data=b'{}'))
            self.assertEqual(context.exception.code,403);save.assert_not_called()

    def test_front_half_round_trip_keeps_fs_state_and_original_sources_separate(self):
        from engine import SourceBook
        def cell(address,value,kind='n'):
            return {'address':address,'value':value,'type':kind}
        source=SourceBook.from_data({'kind':'workbook','sheets':[
            {'name':'Balances','cells':[cell('A1','Code','s'),cell('B1','Name','s'),cell('C1','Closing','s'),
                 cell('A2','1002','s'),cell('B2','Bank','s'),cell('C2','100.005'),
                 cell('A3','3001','s'),cell('B3','Capital','s'),cell('C3','-100.005')]},
            {'name':'BankDetail','cells':[cell('B1','Name','s'),cell('C1','Amount','s'),
                 cell('B2','Real bank balance','s'),cell('C2','100.005')]}]})
        entry={'path':'synthetic.xlsx','sha256':'synthetic-hash','book':source}
        spec={'path':'synthetic.xlsx','sheet':'Balances','first_row':2,'last_row':3,
              'columns':{'code':'A','name':'B','amount':'C'},'amount_mode':'signed','sign_convention':'debit_positive',
              'source_role':'closing_balances','currency_basis':'functional','unit':1,
              'context':{'entity':'Test HK Limited','period_start':'2025-04-01','period_end':'2026-03-31','currency':'HKD'}}
        def post(route,payload):
            request=Request(self.base+route,data=json.dumps(payload).encode(),
                headers={'Content-Type':'application/json','X-Audit-Token':server.TOKEN})
            return json.load(urlopen(request))
        with tempfile.TemporaryDirectory(dir=Path(__file__).parent) as directory:
            folder=Path(directory)
            with patch.object(server,'DB',folder/'isolated.sqlite3'),patch.object(server,'OUTPUT',folder), \
                 patch.object(server,'book',return_value=entry),patch.object(server,'intake_hash',return_value='synthetic-hash'):
                before={'project':{'example':'FS state retained'},'bindings':[{'id':'existing-FS-binding'}]}
                server.save_state(before)
                preview=post('/api/intake/preview',spec)
                imported=post('/api/intake/import',{**spec,'sha256':preview['sha256'],'selected_rows':[2,3],
                    'metadata_confirmed':True,'selection_confirmed':True})
                mapping=[]
                for line in imported['draft']['lines']:
                    code=line['source_code']
                    mapping.append({'row_id':line['row_id'],'account_code':code,'account_name':line['name'],
                        'lead_code':'K101' if code=='1002' else 'O101','lead_title':'Cash' if code=='1002' else 'Capital'})
                mapped=post('/api/intake/map',{'mappings':mapping})
                self.assertEqual(mapped['tb']['difference'],'0.000')
                detail_spec={**spec,'sheet':'BankDetail','last_row':2,'source_role':'support_details','columns':{'name':'B','amount':'C'}}
                supported=post('/api/intake/support',{'account_code':'1002','label':'Client bank schedule','spec':detail_spec,
                    'sha256':'synthetic-hash','selected_rows':[2],'metadata_confirmed':True,'selection_confirmed':True})
                bank=next(b for b in supported['breakdowns'] if b['lead_code']=='K101')
                self.assertEqual(bank['difference'],'0.000')
                exported=post('/api/intake/export',{})
                self.assertIn('100.005',urlopen(self.base+quote(exported['url'])).read().decode('utf-8-sig'))
                self.assertEqual(server.state(),before)
                self.assertIsNone(post('/api/intake/reset',{'confirm':True})['draft'])
                self.assertEqual(server.state(),before)

    def test_intake_mutations_also_require_session_token(self):
        with patch.object(server,'save_intake') as save:
            for endpoint in ['import','map','support','reset','export']:
                with self.subTest(endpoint=endpoint), self.assertRaises(HTTPError) as context:
                    urlopen(Request(self.base+'/api/intake/'+endpoint,data=b'{}'))
                self.assertEqual(context.exception.code,403)
            save.assert_not_called()

    def test_csv_client_labels_are_not_active_excel_formulas(self):
        for label in ['=HYPERLINK("unsafe")','+cmd','-command','@SUM(A1:A2)','  =1+1']:
            self.assertTrue(server.csv_text(label).startswith("'"))
        self.assertEqual(server.csv_text('Cash at bank'),'Cash at bank')

if __name__=='__main__': unittest.main()
