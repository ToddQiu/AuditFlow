"""Read-only corpus extraction. Never saves or executes a source document."""
import argparse, collections, gzip, hashlib, io, json, logging, posixpath
import struct, sys, warnings, zipfile, os, uuid, re, threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
import xml.etree.ElementTree as ET
from contextlib import ExitStack

BASE = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE / '_deps'))
sys.stdout.reconfigure(encoding='utf-8')
logging.getLogger('pypdf').setLevel(logging.ERROR)
warnings.filterwarnings('ignore')
S = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
MAX_XML_BYTES = 80 * 1024 * 1024
HASH_LOCKS = {}
HASH_LOCKS_GUARD = threading.Lock()
PDF_LOCK = threading.Lock()

def large_workbook_metadata(path):
    """Do not expand enormous workbook XML into a Python object graph.

    The full package is still hashed by the batch reader. This result is
    explicitly metadata-only and must never count as a completed cell review.
    """
    path = Path(path)
    if path.suffix.lower() not in ['.xlsx', '.xls', '.xlsm', '.xltx']:
        return None
    with path.open('rb') as stream:
        if stream.read(4) != b'PK\x03\x04': return None
    with zipfile.ZipFile(path) as package:
        xml_members = [member for member in package.infolist()
                       if member.filename.startswith('xl/') and member.filename.endswith('.xml')]
        oversized = [member for member in xml_members if member.file_size > MAX_XML_BYTES]
        if not oversized: return None
        # These small manifest parts name all sheets without loading sheet XML.
        info = {member.filename: member for member in package.infolist()}
        root = ET.fromstring(package.read('xl/workbook.xml'))
        rels = {e.get('Id'): e.get('Target') for e in ET.fromstring(package.read('xl/_rels/workbook.xml.rels'))}
        sheets = []
        for item in root.find(S+'sheets'):
            target = rels[item.get(R+'id')]
            target = target.lstrip('/') if target.startswith('/') else posixpath.normpath('xl/'+target)
            member = info[target]
            sheets.append({'name': item.get('name'), 'state': item.get('state', 'visible'),
                           'xml_member': target, 'xml_uncompressed_bytes': member.file_size,
                           'xml_compressed_bytes': member.compress_size,
                           'read_status': 'needs_streaming_review', 'cells': [], 'formulas': None,
                           'missing_formula_cache': None, 'errors': []})
        return {'kind': 'workbook', 'method': 'large_workbook_sheet_metadata_only',
                'status': 'needs_streaming_review', 'metadata_only': True,
                'reason': 'Workbook XML exceeds 80 MiB. Full cell extraction requires a dedicated streaming reader; no amount cells were read.',
                'xml_size_limit_bytes': MAX_XML_BYTES,
                'xml_uncompressed_bytes': sum(member.file_size for member in xml_members),
                'largest_xml_uncompressed_bytes': max(member.file_size for member in xml_members),
                'oversized_xml_members': [{'name': member.filename, 'uncompressed_bytes': member.file_size,
                                           'compressed_bytes': member.compress_size} for member in oversized],
                'sheets': sheets, 'external_links': [],
                'formula_values': 'not_read_metadata_only',
                'limitation': 'Sheet inventory only. Missing cells are not zero and this file has not been read in full.'}

def digest(path):
    h = hashlib.sha256()
    with open(path, 'rb') as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b''): h.update(block)
    return h.hexdigest()

def workbook(path):
    deferred = large_workbook_metadata(path)
    if deferred is not None: return deferred
    with zipfile.ZipFile(path) as z, ExitStack() as streams:
        strings = []
        if 'xl/sharedStrings.xml' in z.namelist():
            string_root = None
            for event, item in ET.iterparse(streams.enter_context(z.open('xl/sharedStrings.xml')), events=['start', 'end']):
                if event == 'start':
                    if item.tag == S+'sst': string_root = item
                    continue
                if item.tag == S+'si':
                    strings.append(''.join(t.text or '' for t in item.iter(S+'t')))
                    item.clear()
                    if string_root is not None: string_root.remove(item)
        root = ET.fromstring(z.read('xl/workbook.xml'))
        rels = {e.get('Id'):e.get('Target') for e in ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
        sheets = []
        for item in root.find(S+'sheets'):
            target = rels[item.get(R+'id')]
            target = target.lstrip('/') if target.startswith('/') else posixpath.normpath('xl/'+target)
            cells, errors, formulas, missing = [], [], 0, 0
            sheet_data = None
            for event, cell in ET.iterparse(streams.enter_context(z.open(target)), events=['start', 'end']):
                if event == 'start':
                    if cell.tag == S+'sheetData': sheet_data = cell
                    continue
                if cell.tag == S+'row':
                    cell.clear()
                    if sheet_data is not None: sheet_data.remove(cell)
                    continue
                if cell.tag != S+'c': continue
                addr, kind = cell.get('r'), cell.get('t','n')
                v, f = cell.find(S+'v'), cell.find(S+'f')
                value = None if v is None else v.text
                if kind == 's' and value is not None: value = strings[int(value)]
                elif kind == 'inlineStr': value = ''.join(t.text or '' for t in cell.iter(S+'t'))
                formula = None
                if f is not None:
                    formula = {'text':f.text or '', 'type':f.get('t'), 'shared':f.get('si'), 'range':f.get('ref')}
                    formulas += 1
                    if value is None: missing += 1
                if value is not None or formula is not None:
                    cells.append({'address':addr,'type':kind,'value':value,'formula':formula,'style':cell.get('s')})
                if kind == 'e': errors.append({'address':addr,'value':value})
                cell.clear()
            sheets.append({'name':item.get('name'),'state':item.get('state','visible'),'cells':cells,
                           'formulas':formulas,'missing_formula_cache':missing,'errors':errors})
            masters={c['formula']['shared']:c for c in cells if c.get('formula') and c['formula']['type']=='shared' and c['formula']['text']}
            for c in cells:
                f=c.get('formula')
                if f and f['type']=='shared' and not f['text']:
                    master=masters.get(f['shared'])
                    if master:
                        from openpyxl.formula.translate import Translator
                        try:
                            f['text']=Translator('='+master['formula']['text'],origin=master['address']).translate_formula(c['address'])[1:]
                            f['translated_from']=master['address']
                        except Exception as exc:
                            # A formula-translation limitation must not discard
                            # the workbook's original numeric cache and labels.
                            # No replacement formula/result is manufactured.
                            f['unresolved_shared']=True
                            f['translation_error']=f'{type(exc).__name__}: {exc}'
                    else: f['unresolved_shared']=True
        links = [n for n in z.namelist() if n.startswith('xl/externalLinks/') and n.endswith('.rels')]
        targets = []
        for name in links:
            for e in ET.fromstring(z.read(name)): targets.append(e.get('Target'))
        return {'kind':'workbook','method':'ooxml_all_populated_cells','sheets':sheets,
                'external_links':targets,'formula_values':'saved_cache_not_recalculated'}

def old_workbook(path):
    import xlrd
    wb = xlrd.open_workbook(str(path), on_demand=True)
    sheets=[]
    for sheet in wb.sheets():
        cells=[]
        for r in range(sheet.nrows):
            for c in range(sheet.ncols):
                cell=sheet.cell(r,c)
                if cell.ctype in [0,6]: continue
                col=''; n=c+1
                while n: n,k=divmod(n-1,26); col=chr(65+k)+col
                value=cell.value
                if cell.ctype==xlrd.XL_CELL_NUMBER: value=str(value)
                cells.append({'address':f'{col}{r+1}','type':'n' if cell.ctype==2 else 's',
                              'value':value,'formula':None,'style':None})
        sheets.append({'name':sheet.name,'cells':cells,'formulas':None,'errors':[]})
    wb.release_resources()
    return {'kind':'workbook','method':'xls_cached_values_only','sheets':sheets,
            'external_links':[],'formula_values':'legacy_formula_text_unavailable'}

def docx(path):
    with zipfile.ZipFile(path) as z:
        root=ET.fromstring(z.read('word/document.xml'))
        paragraphs=[]
        def text(node):
            result=[]
            def walk(e):
                if e.tag in [W+'del',W+'moveFrom']: return
                if e.tag==W+'t': result.append(e.text or '')
                elif e.tag==W+'tab': result.append('\t')
                elif e.tag in [W+'br',W+'cr']: result.append('\n')
                else:
                    for ch in e: walk(ch)
            walk(node)
            return ''.join(result)
        for p in root.iter(W+'p'): paragraphs.append(text(p))
        tables=[]
        for i,t in enumerate(root.iter(W+'tbl')):
            tables.append({'index':i,'rows':[[text(c) for c in row.findall(W+'tc')] for row in t.findall(W+'tr')]})
        revisions=collections.Counter(e.get(W+'author','') for e in root.iter() if e.tag in [W+'ins',W+'del'])
        return {'kind':'document','method':'docx_current_text_excluding_deletions','paragraphs':paragraphs,
                'tables':tables,'revision_authors':dict(revisions)}

def old_doc(path):
    import olefile
    with olefile.OleFileIO(str(path)) as ole:
        if not ole.exists('WordDocument'): raise ValueError('OLE is not a Word document, possibly encrypted Office')
        data=ole.openstream('WordDocument').read()
        flags=struct.unpack_from('<H',data,10)[0]
        if flags & 0x8100: raise ValueError('Encrypted legacy Word document')
        offset=32; cw=struct.unpack_from('<H',data,offset)[0]; offset+=2+cw*2
        clw=struct.unpack_from('<H',data,offset)[0]; offset+=2+clw*4
        pairs=struct.unpack_from('<H',data,offset)[0]; offset+=2
        if pairs<34: raise ValueError('Unsupported pre-Word97 FIB')
        fc,lcb=struct.unpack_from('<II',data,offset+33*8)
        table=ole.openstream('1Table' if flags & 0x200 else '0Table').read()[fc:fc+lcb]
        offset=0
        while offset<len(table) and table[offset]==1:
            offset+=3+struct.unpack_from('<H',table,offset+1)[0]
        if offset>=len(table) or table[offset]!=2: raise ValueError('Word piece table not found')
        length=struct.unpack_from('<I',table,offset+1)[0]; plc=table[offset+5:offset+5+length]
        count=(length-4)//12; cp=struct.unpack_from('<'+'I'*(count+1),plc,0); texts=[]
        for i in range(count):
            encoded=struct.unpack_from('<I',plc,4*(count+1)+i*8+2)[0]
            compressed=bool(encoded & 0x40000000); start=encoded & 0x3fffffff
            chars=cp[i+1]-cp[i]
            if compressed: start//=2
            raw=data[start:start+chars*(1 if compressed else 2)]
            texts.append(raw.decode('cp1252' if compressed else 'utf-16le',errors='replace'))
        return {'kind':'document','method':'legacy_word_piece_table_text_only',
                'text':''.join(texts),'tables':[],'limitation':'Legacy revisions and layout require Word confirmation'}

def read_file(path):
    path=Path(path)
    with path.open('rb') as f: magic=f.read(8)
    ext=path.suffix.lower()
    if magic.startswith(b'PK') and ext in ['.xlsx','.xls','.xlsm','.xltx']: return workbook(path)
    if ext=='.xls': return old_workbook(path)
    if magic.startswith(b'PK') and ext in ['.docx','.doc','.wps']: return docx(path)
    if ext in ['.doc','.wps']: return old_doc(path)
    if ext=='.pdf':
        # PyMuPDF is task-local in _deps. Its native API must not run in parallel
        # threads; the rest of the corpus reader can still use its worker pool.
        fallback_reason = None
        try:
            import pymupdf
            with PDF_LOCK, pymupdf.open(str(path)) as reader:
                if reader.needs_pass and not reader.authenticate(''):
                    raise ValueError('Encrypted PDF')
                pages = [page.get_text('text') or '' for page in reader]
            backend = 'pymupdf'
        except Exception as exc:
            fallback_reason = f'{type(exc).__name__}: {exc}'
            from pypdf import PdfReader
            try:
                reader = PdfReader(str(path))
                if reader.is_encrypted and not reader.decrypt(''):
                    raise ValueError('Encrypted PDF')
                pages = [page.extract_text() or '' for page in reader.pages]
            except Exception as fallback:
                raise ValueError(f'PDF text extraction failed; PyMuPDF: {fallback_reason}; pypdf: {type(fallback).__name__}: {fallback}') from fallback
            backend = 'pypdf'
        result = {'kind':'pdf','method':'pdf_text_layer','backend':backend,'pages':pages,
                  'scan_pages':[i+1 for i,t in enumerate(pages) if len(t.strip())<20],
                  'scan_pages_basis':'Fewer than 20 stripped text-layer characters; not an OCR result.'}
        if fallback_reason is not None: result['fallback_reason'] = fallback_reason
        return result
    if ext=='.zip':
        with zipfile.ZipFile(path) as z:
            return {'kind':'archive','method':'member_inventory_only','members':
                    [{'name':i.filename,'size':i.file_size,'encrypted':bool(i.flag_bits&1)} for i in z.infolist()]}
    if ext in ['.txt','.md','.csv','.json','.py','.ps1','.mjs','.js','.html','.xml','.cs']:
        raw=path.read_bytes()
        for enc in ['utf-8-sig','utf-16','gb18030','cp1252']:
            try: return {'kind':'text','method':'text_file','text':raw.decode(enc)}
            except UnicodeError: pass
    raise ValueError('No automatic text reader for this format')

def relevant(item):
    rel=item['relative'].lower()
    if any(x in rel for x in ['.edge-profile','.edge-tyc','.edge-profile-copy','node_modules','payroll -','payroll and expense']): return False
    if item['ext'] not in ['.xlsx','.xls','.xlsm','.docx','.doc','.wps','.pdf','.txt','.md','.csv','.py','.ps1','.mjs','.js','.html','.cs','.zip']: return False
    if '.codex_tmp' in rel and item['ext'] not in ['.py','.ps1','.mjs','.js']: return False
    return True

def main():
    parser=argparse.ArgumentParser(); parser.add_argument('--partition',default='core')
    parser.add_argument('--workers', type=int, default=2)
    parser.add_argument('--retry-unreadable', action='store_true',
                        help='Retry previously unreadable files after reader fixes; retain prior diagnostic rows')
    parser.add_argument('--retry-error-prefix', default='',
                        help='Limit a retry to one diagnosed error family, without retrying other known failures')
    args=parser.parse_args()
    if not 1 <= args.workers <= 4: parser.error('--workers must be between 1 and 4')
    items=json.loads((BASE/'inventory.json').read_text(encoding='utf-8'))
    def partition(item):
        if item['top']=='Aska-26中期': return 'aska'
        if item['top'] in ['天瑞文件-24.02.2026','Acenda Caeht']: return 'tianrui'
        return 'core'
    chosen=[i for i in items if partition(i)==args.partition and relevant(i)]
    corpus=BASE/'corpus'/args.partition; corpus.mkdir(parents=True,exist_ok=True)
    manifest=BASE/f'read_{args.partition}.jsonl'; seen={}; counts=collections.Counter()
    previous=[]
    if manifest.exists():
        # A terminated writer may leave one incomplete JSON line. Keep it as
        # diagnostic data, but ensure a resumed record starts on its own line.
        with manifest.open('rb') as stream:
            stream.seek(0, 2)
            if stream.tell():
                stream.seek(-1, 2)
                incomplete = stream.read(1) != b'\n'
            else: incomplete = False
        if incomplete:
            with manifest.open('ab') as stream: stream.write(b'\n')
        for line in manifest.read_text(encoding='utf-8').splitlines():
            try: previous.append(json.loads(line))
            except ValueError: pass
    # The append-only manifest retains failed attempts. Progress and resume
    # decisions use each source path's latest attempt, not duplicate history.
    previous=list({x['path']:x for x in previous}.values())
    completed={x['path'] for x in previous
               if not (args.retry_unreadable and x['status']=='unreadable'
                       and (not args.retry_error_prefix or x.get('error','').startswith(args.retry_error_prefix)))}
    for x in previous:
        if x['path'] in completed: counts[x['status']]+=1
        if x.get('sha256'): seen[x['sha256']]=x.get('same_as',x['relative'])
    pending=[x for x in chosen if x['path'] not in completed]
    def priority(item):
        rel=item['relative'].lower()
        if re.search(r'fs_|afs|annual report|financial|consol|group.tb|lead.ifrs|note\s*\d',rel): return 0
        if any(x in rel for x in ['working','底稿','合并','a3.1','a4.1']): return 1
        if item['ext'] in ['.xlsx','.xls','.doc','.docx','.py','.ps1','.mjs']: return 2
        return 3
    pending.sort(key=priority)
    def process(item):
            result=dict(item)
            try:
                sha=digest(item['path']); result['sha256']=sha
                deferred=large_workbook_metadata(item['path'])
                target=corpus/(sha+('.metadata.json.gz' if deferred is not None else '.json.gz'))
                with HASH_LOCKS_GUARD:
                    content_lock=HASH_LOCKS.setdefault(sha, threading.Lock())
                with content_lock:
                    if target.exists() and deferred is None:
                        if digest(item['path'])!=sha:
                            raise ValueError('Source changed while reading; retry after the source has finished saving')
                        result.update({'status':'duplicate','same_as':seen.get(sha,'same SHA256 content'),'corpus':str(target)})
                    else:
                        data=deferred if deferred is not None else read_file(item['path'])
                        if digest(item['path'])!=sha:
                            raise ValueError('Source changed while reading; retry after the source has finished saving')
                        if not target.exists():
                            temp=target.with_name(target.name+'.'+uuid.uuid4().hex+'.tmp')
                            with gzip.open(temp,'wt',encoding='utf-8',compresslevel=1) as f: json.dump(data,f,ensure_ascii=False)
                            os.replace(temp,target)
                        result.update({'status':data.get('status','read'),'method':data['method'],'kind':data['kind'],'corpus':str(target)})
                        if data['kind']=='workbook':
                            result['sheets']=[{'name':s['name'],'cells':len(s['cells']),'formulas':s.get('formulas'),
                                              'missing_formula_cache':s.get('missing_formula_cache'),'errors':s['errors'][:25],
                                              'unresolved_shared_formulas':sum(bool(c.get('formula',{}).get('unresolved_shared')) for c in s['cells'] if c.get('formula')),
                                              **{k:s[k] for k in ['xml_member','xml_uncompressed_bytes','xml_compressed_bytes','read_status'] if k in s}}
                                             for s in data['sheets']]
                        if data.get('metadata_only'):
                            result.update({k:data[k] for k in ['metadata_only','reason','xml_size_limit_bytes',
                                          'xml_uncompressed_bytes','largest_xml_uncompressed_bytes','oversized_xml_members','limitation']})
                        if data['kind']=='pdf': result.update({'pages':len(data['pages']),'scan_pages':data['scan_pages']})
            except Exception as e: result.update({'status':'unreadable','error':f'{type(e).__name__}: {e}'})
            return result
    with manifest.open('a',encoding='utf-8') as out, ThreadPoolExecutor(max_workers=args.workers) as pool:
        futures={pool.submit(process,item):item for item in pending}
        for n,future in enumerate(as_completed(futures),len(completed)):
            result=future.result(); counts[result['status']]+=1
            if result.get('sha256'): seen[result['sha256']]=result['relative']
            out.write(json.dumps(result,ensure_ascii=False)+'\n'); out.flush()
            if n%50==0: print(args.partition,n,'/',len(chosen),dict(counts),flush=True)
    print('DONE',args.partition,len(chosen),dict(counts),flush=True)

if __name__=='__main__': main()
