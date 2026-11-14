"""Local-only audit data workspace. All input files are read-only."""
from __future__ import annotations
import argparse, csv, hashlib, io, json, mimetypes, os, secrets, sqlite3, subprocess, sys, threading, time, uuid
from datetime import date
from collections import Counter, OrderedDict
from contextlib import contextmanager
from decimal import Decimal, InvalidOperation
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse, unquote

APP = Path(__file__).resolve().parent
ROOT = Path(r'E:\CodeX\Acenda Caeht')
RESEARCH = APP / 'research'
DATA = Path(os.environ['LOCALAPPDATA']) / 'AuditFlow' / 'data' if getattr(sys, 'frozen', False) else APP
OUTPUT = DATA / 'outputs'
OUTPUT.mkdir(parents=True, exist_ok=True)
TOKEN = secrets.token_urlsafe(32)
DB = DATA / 'workspace.sqlite3'
CACHE = OrderedDict()
LOCK = threading.RLock()
KNOWN = {str(Path(i['path']).resolve()):i for i in (json.loads((RESEARCH/'inventory.json').read_text(encoding='utf-8')) if (RESEARCH/'inventory.json').exists() else [])}
LIBRARY_ROOT = DATA / '客户资料'
LIBRARIES = {}

def document_library():
    from library import Library
    key = (str(LIBRARY_ROOT.resolve()), str(DB.resolve()))
    with LOCK:
        if key not in LIBRARIES:
            LIBRARIES[key] = Library(LIBRARY_ROOT, DB)
        return LIBRARIES[key]

def dumps(value):
    return json.dumps(value, ensure_ascii=False, default=lambda v:str(v) if isinstance(v,Decimal) else None)

def csv_text(value):
    value=str(value or '')
    return "'"+value if value.lstrip().startswith(('=','+','-','@')) else value

@contextmanager
def connection():
    con=sqlite3.connect(DB)
    try:
        con.execute('CREATE TABLE IF NOT EXISTS state (id INTEGER PRIMARY KEY CHECK(id=1), payload TEXT NOT NULL)')
        with con:
            yield con
    finally:
        con.close()

def state():
    with connection() as con:
        row=con.execute('SELECT payload FROM state WHERE id=1').fetchone()
    return json.loads(row[0]) if row else {'project':{'current_year':2026,'comparative_year':2025,
              'currency':'HKD','unit':1,'source_confirmed':False},'bindings':[]}

def save_state(payload):
    with connection() as con:
        con.execute('INSERT OR REPLACE INTO state VALUES(1,?)',(dumps(payload),))

def intake_draft():
    with connection() as con:
        con.execute('CREATE TABLE IF NOT EXISTS intake (id INTEGER PRIMARY KEY CHECK(id=1), payload TEXT NOT NULL)')
        row=con.execute('SELECT payload FROM intake WHERE id=1').fetchone()
    return json.loads(row[0]) if row else None

def save_intake(draft):
    with connection() as con:
        con.execute('CREATE TABLE IF NOT EXISTS intake (id INTEGER PRIMARY KEY CHECK(id=1), payload TEXT NOT NULL)')
        if draft is None: con.execute('DELETE FROM intake WHERE id=1')
        else: con.execute('INSERT OR REPLACE INTO intake VALUES(1,?)',(dumps(draft),))

def intake_hash(path):
    return sha256(source_path(path,['.xlsx','.xls','.xlsm']))

def intake_status():
    from intake import summarize
    return summarize(intake_draft(),hash_reader=intake_hash)

def source_path(raw, extensions=None):
    path=Path(raw)
    if not path.is_absolute(): path=ROOT/path
    path=path.resolve()
    if str(path) not in KNOWN and not path.is_relative_to(LIBRARY_ROOT.resolve()): raise ValueError('请选择已登记的来源文件，或放入客户资料库。')
    if extensions and path.suffix.lower() not in extensions: raise ValueError('此文件格式不适用于当前操作。')
    if not path.is_file(): raise ValueError('来源文件已移动，请重新登记。')
    return path

def sha256(path):
    h=hashlib.sha256()
    with path.open('rb') as f:
        for block in iter(lambda:f.read(1024*1024),b''): h.update(block)
    return h.hexdigest()

def book(path):
    from engine import SourceBook
    p=source_path(path,['.xlsx','.xls','.xlsm']); stamp=(p.stat().st_mtime_ns,p.stat().st_size)
    before=sha256(p)
    with LOCK:
        item=CACHE.get(str(p))
        if item and item['stamp']==stamp and item['sha256']==before:
            CACHE.move_to_end(str(p)); return item
    obj=SourceBook(str(p)); after=sha256(p)
    if before!=after: raise ValueError('读取期间来源文件正在保存，请保存完成后重试。')
    item={'book':obj,'stamp':stamp,'sha256':after,'path':str(p)}
    with LOCK:
        CACHE[str(p)]=item; CACHE.move_to_end(str(p))
        while len(CACHE)>3: CACHE.popitem(last=False)
    return item

def reading():
    latest={}
    for path in RESEARCH.glob('read_*.jsonl'):
        for line in path.read_text(encoding='utf-8').splitlines():
            try:
                record=json.loads(line); latest[record['path']]=record
            except ValueError: pass
    return list(latest.values())

def files(kind,q=''):
    allowed={'.xlsx','.xls','.xlsm'} if kind=='workbook' else {'.docx','.doc','.wps'}
    statuses={x['path']:x for x in reading()}
    output=[]
    for key,item in KNOWN.items():
        rel=item['relative']; low=rel.lower()
        if item['ext'] not in allowed or any(t in low for t in ['.edge','payroll','.codex_tmp']): continue
        if q and q.lower() not in low: continue
        status=statuses.get(item['path'],{})
        output.append({'path':key,'name':Path(key).name,'relative':rel,'size':item['size'],
                       'read_status':status.get('status','pending'),'error':status.get('error')})
    catalogue=document_library(); catalogue.request_scan()
    for item in catalogue.snapshot()['files']:
        if item['availability']!='present' or item['ext'] not in allowed or q and q.lower() not in item['relative'].lower(): continue
        output.append({**item,'relative':'客户资料库 / '+item['relative'],'read_status':'待选择取数范围'})
    output.sort(key=lambda x:(-1 if x['path'].startswith(str(LIBRARY_ROOT.resolve())) else 0 if 'aska2026.9.25\\aska amend\\fs_working' in x['relative'].lower() else
                              1 if any(t in x['relative'].lower() for t in ['fs_','afs','annual report','chung shun']) else 2,x['relative']))
    return output

def source_amount(payload, *, preview=False, entry=None):
    from engine import format_money, FormulaReview
    entry=entry or book(payload['source_path']); refs=payload.get('refs',[])
    if not refs or len(refs)>200: raise ValueError('请选择 1 至 200 个来源单元格。')
    amount=Decimal(0); evidence=[]; keys=set(); has_unverified=False; has_formula=False
    confirm_saved=payload.get('confirm_saved_values') is True
    for ref in refs:
        sheet,address=ref['sheet'],ref['address'].replace('$','').upper()
        if (sheet,address) in keys: raise ValueError('同一个来源单元格被重复加入。')
        keys.add((sheet,address)); cell=entry['book'].cell(sheet,address)
        numeric=cell.get('numeric')
        if numeric is None: raise ValueError(f'{sheet}!{address} 不是可填列的金额，空值或错误不会当作零。')
        legacy=entry['book'].data.get('formula_values')=='legacy_formula_text_unavailable'
        verification={'status':'pass','basis':'numeric_source_value',
                      'reason':'已读取原始数值；无公式需要重算，这不代表金额已经审计。'}
        if legacy:
            verification={'status':'needs_review','basis':'legacy_saved_value_only',
                          'reason':'旧 XLS 无法取得公式文本，不能判断此格是常数或公式结果。','confirmed_saved_values':confirm_saved}
            has_unverified=True
            if not confirm_saved and not preview: raise ValueError('旧 XLS 仅有保存值，请核实后明确确认，或先在 Excel 另存 XLSX。')
        elif cell.get('formula'):
            has_formula=True
            try:
                calculated=entry['book'].evaluate(sheet,address)
            except FormulaReview as exc:
                verification={'status':'needs_review','basis':'saved_value_only','reason':str(exc),
                              'confirmed_saved_values':confirm_saved}
                has_unverified=True
                if not confirm_saved and not preview:
                    raise ValueError(f'{sheet}!{address} 目前只能读取保存值，未独立重算（{exc}）。核实后请明确勾选确认使用保存值；该结果仍是待复核的填列预览。') from exc
            else:
                if calculated!=numeric:
                    raise ValueError(f'{sheet}!{address} 保存值 {numeric} 与独立计算值 {calculated} 不一致；拒绝绑定或填列，请先核实并重新保存来源文件。')
                verification={'status':'pass','basis':'independently_evaluated_formula',
                              'calculated_value':str(calculated),'saved_value':str(numeric),
                              'reason':'支持的公式独立计算与保存值一致；仍需核对来源及分类完整性。'}
        factor=Decimal(str(ref.get('coefficient',1)))
        if factor not in [Decimal(1),Decimal(-1)]: raise ValueError('符号只能选择原方向或反向。')
        amount+=numeric*factor
        evidence.append({'sheet':sheet,'address':address,'coefficient':str(factor),
                         'value':str(numeric),'formula':cell.get('formula'),'verification':verification})
    divisor=Decimal(str(payload.get('unit_divisor',1)))
    if divisor not in [Decimal(1),Decimal(1000),Decimal(1000000)]: raise ValueError('单位换算请选择元、千元或百万元。')
    amount/=divisor
    precision=Decimal(str(payload.get('places',0)))
    if precision not in [Decimal(0),Decimal(2)]: raise ValueError('请选择整数报告或两位小数 working。')
    places=int(precision)
    basis=('explicitly_confirmed_saved_values' if confirm_saved else 'unconfirmed_saved_value_preview') if has_unverified else 'independently_verified_saved_values' if has_formula else 'numeric_source_values'
    return {'value':format_money(amount,places),'exact_amount':str(amount),
            'source':{'path':entry['path'],'sha256':entry['sha256'],'refs':evidence,
                      'unit_divisor':str(divisor),'value_basis':basis,
                      'verification_status':'needs_review' if has_unverified else 'pass',
                      'confirm_saved_values':confirm_saved},'places':places,
            'needs_review':has_unverified,
            'warning':('部分来源采用您明确确认的保存值，未独立重算，输出只能作为待复核预览。' if confirm_saved else
                       '金额仅用于预览，部分来源保存值未独立重算；绑定前须明确确认采用保存值。') if has_unverified else ''}

def validate_project(project):
    year=int(project.get('current_year',0)); prior=int(project.get('comparative_year',0))
    if not 2000<=year<=2100 or not 1900<=prior<year: raise ValueError('请明确本期和比较年度。')
    if project.get('currency') not in ['HKD','RMB','USD','EUR','GBP','OTHER']: raise ValueError('请选择币种。')
    if project.get('unit') not in [1,1000,1000000]: raise ValueError('请确认来源金额单位。')
    if project.get('period_start') or project.get('period_end') or project.get('source_confirmed'):
        try:
            beginning=date.fromisoformat(project.get('period_start','')); ending=date.fromisoformat(project.get('period_end',''))
        except (ValueError,TypeError): raise ValueError('确认来源前，请填写实际报告起止日期。')
        if beginning>ending: raise ValueError('报告起始日不能晚于期末日。')
    return year,prior

def project_context(project):
    return {k:project.get(k) for k in ['source_path','template_path','current_year','comparative_year','period_start','period_end','currency','unit','current_column']}

def confirmed_source(project):
    if not project.get('source_confirmed') or not project.get('confirmed_sha256'):
        raise ValueError('请先确认来源版本、期间、币种、单位及当年金额列。')
    if project['confirmed_sha256']!=sha256(source_path(project['source_path'])):
        raise ValueError('来源文件已改变，请重新确认版本后再绑定。')

def binding_preview(current):
    from templates import preview_bindings
    project=current['project']; year,prior=validate_project(project)
    template=source_path(project.get('template_path',''),['.docx'])
    binding_list=current.get('bindings',[])
    if not binding_list: raise ValueError('先绑定一个当年金额，再查看修改预览。')
    rejects=[]
    template_hash=sha256(template); source_hashes={}
    for b in binding_list:
        if b.get('project_context')!=project_context(project):
            rejects.append({'reason':'期间、币种、单位或金额列已改变，需要重新绑定。','binding':b})
        source_file=b['source']['path']
        if source_file not in source_hashes: source_hashes[source_file]=sha256(source_path(source_file))
        if b.get('source',{}).get('sha256')!=source_hashes[source_file]:
            rejects.append({'reason':'来源文件已改变，需要重新核对绑定。','binding':b})
        if b.get('template_sha256')!=template_hash:
            rejects.append({'reason':'模板文件已改变，需要重新定位。','binding':b})
    preview=preview_bindings(str(template),binding_list,year,prior)
    preview.setdefault('rejects',[]).extend(rejects)
    return preview

class Handler(BaseHTTPRequestHandler):
    def log_message(self,*args): pass
    def valid_host(self):
        return self.headers.get('Host') in {f'127.0.0.1:{self.server.server_port}',f'localhost:{self.server.server_port}'}
    def send(self,code,payload,content_type='application/json; charset=utf-8'):
        raw=payload if isinstance(payload,bytes) else dumps(payload).encode('utf-8') if content_type.startswith('application/json') else payload
        self.send_response(code); self.send_header('Content-Type',content_type)
        self.send_header('Content-Length',str(len(raw))); self.send_header('Cache-Control','no-store')
        self.send_header('X-Content-Type-Options','nosniff'); self.end_headers(); self.wfile.write(raw)
    def do_GET(self):
        try:
            if not self.valid_host(): return self.send(403,{'error':'只接受本机地址访问。'})
            url=urlparse(self.path); p=unquote(url.path); query={k:v[0] for k,v in parse_qs(url.query).items()}
            if p=='/api/session': return self.send(200,{'application':'HK_Audit_MVP','token':TOKEN,'state':state()})
            if p=='/api/library':
                catalogue=document_library(); catalogue.request_scan()
                return self.send(200,catalogue.snapshot())
            if p=='/api/intake/state': return self.send(200,intake_status())
            if p=='/api/intake/evidence':
                path=RESEARCH/'front_half_evidence.json'
                return self.send(200,json.loads(path.read_text(encoding='utf-8')) if path.exists() else {})
            if p=='/api/status':
                rows=reading(); counts=Counter(x['status'] for x in rows)
                return self.send(200,{'total_files':len(KNOWN),'processed':len(rows),'counts':dict(counts),
                      'formats':dict(Counter(x['ext'] for x in rows)),
                      'scanned_pdfs':sum(bool(x.get('scan_pages')) for x in rows),
                      'scope':'来源只读；文件读取不代表金额已获复核确认。',
                      'limitations':'扫描件需要视觉核对；旧 DOC 的静态提取不能确认修订视图或排版。'})
            if p=='/api/files': return self.send(200,{'files':files(query.get('kind','workbook'),query.get('q',''))})
            if p=='/api/workbook':
                from engine import inspect_workbook
                entry=book(query['path']); obj=entry['book']
                summary=[{'name':s['name'],'cells':len(s['cells']),'formulas':s.get('formulas'),
                           'errors':s.get('errors',[]),'state':s.get('state','visible')} for s in obj.data['sheets']]
                return self.send(200,{'path':entry['path'],'sha256':entry['sha256'],'sheets':summary,
                      'diagnostics':inspect_workbook(obj),'basis':obj.data.get('formula_values')})
            if p=='/api/sheet':
                obj=book(query['path'])['book']; start=max(1,int(query.get('start',1))); end=min(start+79,int(query.get('end',start+49)))
                return self.send(200,{'sheet':query['sheet'],'rows':obj.rows(query['sheet'],start,end)})
            if p=='/api/cell':
                obj=book(query['path'])['book']
                return self.send(200,{'cell':obj.cell(query['sheet'],query['address']),
                                      'trace':obj.trace(query['sheet'],query['address'],recursive=True)})
            if p=='/api/checks': return self.send(200,{'checks':book(query['path'])['book'].checks(query['sheet'])})
            if p=='/api/template':
                from templates import inspect_template
                path=source_path(query['path'],['.docx','.doc','.wps'])
                if path.suffix.lower()!='.docx':
                    from research.read_corpus import read_file
                    doc=read_file(path)
                    return self.send(200,{'kind':'legacy','text':doc.get('text',''),
                                          'needs_docx_template':True,'reason':'旧 DOC 可只读学习；填列请使用另存的 DOCX 模板。'})
                return self.send(200,inspect_template(str(path)))
            if p=='/api/reading':
                rows=reading(); filter_by=query.get('status'); q=query.get('q','').lower()
                rows=[x for x in rows if (not filter_by or x['status']==filter_by) and (not q or q in x['relative'].lower())]
                offset=max(0,int(query.get('offset',0)))
                return self.send(200,{'total':len(rows),'rows':rows[offset:offset+100]})
            if p=='/api/examples':
                path=RESEARCH/'regression_cases.json'
                return self.send(200,{'cases':json.loads(path.read_text(encoding='utf-8')) if path.exists() else []})
            if p.startswith('/download/'):
                name=p.rsplit('/',1)[-1]; path=OUTPUT/name
                if path.parent.resolve()!=OUTPUT.resolve() or not path.is_file(): raise ValueError('找不到输出文件。')
                raw=path.read_bytes(); return self.send(200,raw,mimetypes.guess_type(name)[0] or 'application/octet-stream')
            relative=p.lstrip('/') or 'index.html'; target=(APP/'web'/relative).resolve()
            if not target.is_relative_to((APP/'web').resolve()) or not target.is_file(): return self.send(404,{'error':'页面不存在。'})
            return self.send(200,target.read_bytes(),(mimetypes.guess_type(str(target))[0] or 'text/plain')+'; charset=utf-8')
        except Exception as exc: self.send(400,{'error':str(exc)})
    def do_POST(self):
        try:
            if not self.valid_host(): return self.send(403,{'error':'只接受本机地址访问。'})
            if self.headers.get('X-Audit-Token')!=TOKEN: return self.send(403,{'error':'页面会话已失效，请刷新。'})
            origin=self.headers.get('Origin')
            if origin and urlparse(origin).netloc!=self.headers.get('Host'): return self.send(403,{'error':'请求来源不匹配。'})
            length=int(self.headers.get('Content-Length',0))
            if length>2_000_000: raise ValueError('请求过大。')
            payload=json.loads(self.rfile.read(length) or b'{}'); p=urlparse(self.path).path
            if p.startswith('/api/library/'):
                catalogue=document_library()
                if p=='/api/library/scan':
                    catalogue.request_scan(force=True)
                    return self.send(200,catalogue.snapshot())
                if p=='/api/library/update': return self.send(200,catalogue.update(payload))
                if p=='/api/library/open-folder':
                    subprocess.Popen(['explorer.exe',str(catalogue.root)])
                    return self.send(200,{'opened':True})
                if p=='/api/library/open':
                    from library import OPENABLE
                    target=catalogue.resolve(payload['path'])
                    if payload.get('reveal') is True:
                        subprocess.Popen(['explorer.exe','/select,',str(target)])
                    else:
                        if target.suffix.lower() not in OPENABLE: raise ValueError('此类型请使用「所在文件夹」查看，不直接运行。')
                        os.startfile(str(target))
                    return self.send(200,{'opened':True})
                raise ValueError('资料库操作不存在。')
            if p.startswith('/api/intake/'):
                from intake import preview_import, import_rows, apply_mappings, add_support, summarize
                arguments={'get_book':book,'amount_reader':source_amount,'hash_reader':intake_hash}
                if p=='/api/intake/preview':
                    return self.send(200,preview_import(payload,get_book=book,amount_reader=source_amount))
                with LOCK:
                    draft=intake_draft()
                    if p=='/api/intake/import':
                        draft=import_rows(draft,payload,**arguments)
                    elif p=='/api/intake/map':
                        if summarize(draft,hash_reader=intake_hash)['stale_sources']:
                            raise ValueError('来源已改变，请先重新导入客户资料。')
                        draft=apply_mappings(draft,payload)
                    elif p=='/api/intake/support':
                        draft=add_support(draft,payload,**arguments)
                    elif p=='/api/intake/reset':
                        if payload.get('confirm') is not True: raise ValueError('请明确确认重建前半程草稿。')
                        draft=None
                    elif p=='/api/intake/export':
                        report=summarize(draft,hash_reader=intake_hash)
                        if draft is None: raise ValueError('先建立管理账 TB 草稿。')
                        if report['stale_sources']: raise ValueError('来源已改变，不导出旧草稿。')
                        name='管理账_TB草稿_'+uuid.uuid4().hex[:8]+'.csv'
                        stream=io.StringIO(); writer=csv.writer(stream)
                        writer.writerow(['实体','起始日','期末日','币种','金额单位','科目代码','名称','借方','贷方','净额（借为正）','底稿编号','底稿标题','状态','来源行ID'])
                        ctx=draft['context']
                        for account in report['tb']['accounts']:
                            writer.writerow([csv_text(ctx['entity']),ctx['period_start'],ctx['period_end'],ctx['currency'],'元',
                                  csv_text(account['account_code']),csv_text(account['account_name']),account['debit'],account['credit'],account['amount'],
                                  csv_text(account['lead_code']),csv_text(account['lead_title']),'管理账草稿／未审定',';'.join(account['source_rows'])])
                        writer.writerow(['借贷差额（未分摊）','','','','','','','','',report['tb']['difference']])
                        writer.writerow(['未映射来源行','','','','','','','','','','','',';'.join(report['tb']['unmapped'])])
                        (OUTPUT/name).write_text(stream.getvalue(),encoding='utf-8-sig')
                        record=name[:-4]+'_来源与明细.json'
                        (OUTPUT/record).write_text(dumps(report),encoding='utf-8')
                        return self.send(200,{'url':'/download/'+name,'source_record':'/download/'+record,
                            'message':'已导出管理账 TB 草稿及来源／明细核对记录；未写入客户 working，尚未包括审计调整及正式审定。'})
                    else: raise ValueError('前半程操作不存在。')
                    save_intake(draft)
                    return self.send(200,summarize(draft,hash_reader=intake_hash))
            if p=='/api/project':
                current=state(); project=payload; validate_project(project)
                if project.get('source_path'): source_path(project['source_path'],['.xlsx','.xls','.xlsm'])
                if project.get('template_path'): source_path(project['template_path'],['.docx','.doc','.wps'])
                if project_context(current['project'])!=project_context(project):
                    current['bindings']=[]
                if project.get('source_confirmed'):
                    if not project.get('source_path') or not str(project.get('current_column','')).strip():
                        raise ValueError('确认来源时必须指定文件及当年金额列，例如 A3.1!Y。')
                    project['confirmed_sha256']=book(project['source_path'])['sha256']
                else: project.pop('confirmed_sha256',None)
                current['project']=project; save_state(current); return self.send(200,current)
            if p=='/api/amount': return self.send(200,source_amount(payload,preview=True))
            if p=='/api/manual-check':
                from engine import sum_check, rounding_check, cashflow_check
                obj=book(payload['source_path'])['book']
                def get(ref): return obj.cell(ref['sheet'],ref['address'])
                if payload['kind']=='cashflow':
                    fields=['opening','operating','investing','financing','fx','closing','bs_cash']
                    return self.send(200,{'check':cashflow_check(*(get(payload[k]) for k in fields))})
                values=[get(ref) for ref in payload.get('refs',[])]
                return self.send(200,{'check':rounding_check(values) if payload['kind']=='rounding' else sum_check(values,get(payload['total']))})
            if p=='/api/bind':
                current=state(); year,prior=validate_project(current['project'])
                if int(payload.get('year',0))!=year: raise ValueError('只能绑定当前年度，比较年度不会修改。')
                confirmed_source(current['project'])
                if payload['source_path']!=current['project'].get('source_path'): raise ValueError('绑定来源与项目确认版本不同。')
                result=source_amount(payload); template=source_path(current['project'].get('template_path',''),['.docx'])
                locator={k:payload[k] for k in ['table','row','column','paragraph','segment'] if k in payload}
                binding={**locator,**result,'id':uuid.uuid4().hex,'year':year,'expected_old':payload['expected_old'],
                         'template_sha256':sha256(template),'confirmed_column_role':'current_year',
                         'project_context':project_context(current['project']),
                         'confirm_year_like_amount':bool(payload.get('confirm_year_like_amount'))}
                trial={**current,'bindings':current['bindings']+[binding]}
                preview=binding_preview(trial)
                if preview.get('rejects'): raise ValueError('绑定未通过：'+dumps(preview['rejects']))
                current['bindings'].append(binding); save_state(current); return self.send(200,{'state':current,'preview':preview})
            if p=='/api/unbind':
                current=state(); current['bindings']=[b for b in current['bindings'] if b['id']!=payload['id']]
                save_state(current); return self.send(200,current)
            if p=='/api/preview': return self.send(200,binding_preview(state()))
            if p=='/api/export':
                from templates import export_docx
                current=state(); project=current['project']; year,prior=validate_project(project)
                confirmed_source(project)
                preview=binding_preview(current)
                if preview.get('rejects'): raise ValueError('存在未通过的绑定，请先处理。')
                template=source_path(project['template_path'],['.docx'])
                name=f'{template.stem}_填列预览_{year}_{uuid.uuid4().hex[:6]}.docx'; output=OUTPUT/name
                result=export_docx(str(template),str(output),current['bindings'],year,prior,author='Todd')
                if not result.get('written') or result.get('rejects') or not output.is_file():
                    raise ValueError('Word 预览未生成：'+dumps(result.get('rejects') or result))
                report=OUTPUT/(output.stem+'_来源记录.json')
                report.write_text(dumps({'project':project,'bindings':current['bindings'],'result':result}),encoding='utf-8')
                return self.send(200,{'file':name,'url':'/download/'+name,'source_record':'/download/'+report.name,
                     'message':'已生成带 Todd 修订的独立预览副本；原始文件未修改。'})
            if p=='/api/csv':
                current=state(); stream=io.StringIO(); writer=csv.writer(stream)
                writer.writerow(['year','target','old','new','source_file','source_cells','source_hash'])
                for b in current['bindings']:
                    target=f"table {b['table']} row {b['row']} col {b['column']}" if 'table' in b else f"paragraph {b['paragraph']} segment {b['segment']}"
                    writer.writerow([b['year'],target,b['expected_old'],b['value'],b['source']['path'],
                                      '; '.join(f"{r['sheet']}!{r['address']}" for r in b['source']['refs']),b['source']['sha256']])
                name='填列清单_'+uuid.uuid4().hex[:6]+'.csv'; (OUTPUT/name).write_text(stream.getvalue(),encoding='utf-8-sig')
                return self.send(200,{'url':'/download/'+name})
            raise ValueError('操作不存在。')
        except Exception as exc: self.send(400,{'error':str(exc)})

def main():
    parser=argparse.ArgumentParser(); parser.add_argument('--port',type=int,default=8776); args=parser.parse_args()
    server=ThreadingHTTPServer(('127.0.0.1',args.port),Handler)
    if sys.stdout is not None:
        print(f'AuditFlow http://127.0.0.1:{args.port}',flush=True)
    server.serve_forever()

if __name__=='__main__': main()
