"""Local customer-document catalogue; source files stay in their original folders."""
from pathlib import Path
from datetime import datetime, timezone
from contextlib import closing
import hashlib
import json
import os
import re
import sqlite3
import threading
import time
import zipfile
from xml.etree import ElementTree as ET

CATEGORIES = {
    'management': '管理账／管理报表', 'tb': '科目余额表／TB', 'ledger': '总账／流水账',
    'assets': '固定资产／折旧', 'contracts': '合同／协议', 'bank': '银行／现金',
    'receivables': '应收／客户明细', 'payables': '应付／供应商明细',
    'income_expenses': '收入／费用／发票', 'tax': '税务资料', 'payroll': '工资／人事',
    'report': '财务报告／审计底稿', 'other': '其他／待分类',
}
STATUSES = {'pending': '待处理', 'in_progress': '处理中', 'waiting': '待补资料', 'done': '已处理'}
OPENABLE = {'.xlsx', '.xls', '.xlsm', '.csv', '.docx', '.doc', '.wps', '.pdf', '.txt', '.png', '.jpg', '.jpeg', '.tif', '.tiff'}
RULES = {
    'management': r'management\s*(?:accounts?|reports?)|管理[账帳帐]|管理报表|管理報表|closing\s*balances?',
    'tb': r'trial[ _-]*balance|(?:^|[\W_])(?:TB|ETB)(?:$|[\W_])|科目[余餘]额|科目餘額|余额表|餘額表',
    'ledger': r'general[ _-]*ledger|(?:^|[\W_])GL(?:$|[\W_])|总[账帳帐]|總[账帳帐]|明[细細][账帳帐]|流水[账帳帐]|序时|序時',
    'assets': r'fixed[ _-]*assets?|depreciation|固定[资資][产產]|折[旧舊]|ppe\s*(?:schedule|register)',
    'contracts': r'contracts?|agreements?|合同|[协協][议議]|租[约約]|貸款合約|合約',
    'bank': r'bank\s*(?:statements?|reconcil|confirm)|银行|銀行|cash\s*book|现金日记|現金日記',
    'receivables': r'receivables?|debtor|应收|應收|客[户戶]明[细細]',
    'payables': r'payables?|creditor|应付|應付|供[应應]商',
    'income_expenses': r'invoices?|revenue|sales\s*(?:details?|register)|expenses?|发票|發票|收入|[费費]用',
    'tax': r'profits?\s*tax|tax\s*(?:comp|return)|[税稅][务務]|[报報][税稅]|利得[税稅]|[税稅]款',
    'payroll': r'payroll|salar(?:y|ies)|工资|工資|薪[酬资資]|mpf|强积金|強積金',
    'report': r'financial\s*statements?|annual\s*report|audit\s*(?:working|report)|财务报告|財務報告|审计|審計|底稿|working|(?:^|[\W_])FS(?:$|[\W_])',
}


def now():
    return datetime.now(timezone.utc).isoformat(timespec='seconds')


def fingerprint(path):
    digest = hashlib.sha256()
    with path.open('rb') as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b''):
            digest.update(block)
    return digest.hexdigest()


def sample_text(path):
    """Bounded classification hints, never an accounting-data extraction."""
    ext = path.suffix.lower()
    if ext in {'.xlsx', '.xlsm', '.docx'}:
        with zipfile.ZipFile(path) as z:
            names = ['xl/workbook.xml', 'xl/sharedStrings.xml'] if ext != '.docx' else ['word/document.xml']
            texts = []
            for name in names:
                if name not in z.namelist():
                    continue
                with z.open(name) as stream:
                    raw = stream.read(200_000)
                if b'<!DOCTYPE' in raw.upper():
                    raise ValueError('不支持含 DTD 的文件')
                if name == 'xl/workbook.xml':
                    texts.extend(node.get('name', '') for node in ET.fromstring(raw).iter() if node.tag.endswith('}sheet'))
                else:
                    texts.extend(re.findall(r'<(?:\w+:)?t(?:\s[^>]*)?>([^<]*)', raw.decode('utf-8', errors='replace'))[:300])
            return ' '.join(texts)[:12000]
    if ext == '.xls':
        import xlrd
        book = xlrd.open_workbook(str(path), on_demand=True)
        try:
            return ' '.join(book.sheet_names())
        finally:
            book.release_resources()
    if ext == '.pdf':
        from pypdf import PdfReader
        with path.open('rb') as stream:
            reader = PdfReader(stream)
            if reader.is_encrypted and not reader.decrypt(''):
                raise ValueError('加密 PDF，需要先解密')
            return ' '.join((reader.pages[i].extract_text() or '')[:6000] for i in range(min(2, len(reader.pages))))
    if ext in {'.txt', '.csv'}:
        with path.open('rb') as stream:
            return stream.read(20000).decode('utf-8-sig', errors='replace')
    return ''


def classify(relative, content):
    name = Path(relative).stem
    for text, basis in [(name, '文件名'), (content, '表名／文字摘要'), (str(Path(relative).parent), '所在文件夹')]:
        hits = [key for key, rule in RULES.items() if re.search(rule, text, re.I)]
        if len(hits) == 1:
            return hits[0], f'{basis}匹配「{CATEGORIES[hits[0]]}」'
        if len(hits) > 1:
            return 'other', f'{basis}包含多个用途：' + '、'.join(CATEGORIES[key] for key in hits)
    return 'other', '未找到明确用途，请手动分类'


class Library:
    def __init__(self, root, db):
        self.root, self.db = Path(root).resolve(), Path(db)
        self.root.mkdir(parents=True, exist_ok=True)
        self.lock, self.scan_lock = threading.RLock(), threading.Lock()
        with closing(sqlite3.connect(self.db)) as con, con:
            con.execute('CREATE TABLE IF NOT EXISTS file_library (path TEXT PRIMARY KEY, payload TEXT NOT NULL)')
            self.records = {path: json.loads(payload) for path, payload in con.execute('SELECT path,payload FROM file_library')}
        self.scanning, self.last_scan, self.errors = False, None, []
        self.last_requested = 0

    def request_scan(self, force=False):
        with self.lock:
            if self.scanning or not force and time.monotonic() - self.last_requested < 8:
                return
            self.scanning = True
            self.last_requested = time.monotonic()
        threading.Thread(target=self.scan, daemon=True, name='AuditFlow-library').start()

    def scan(self):
        with self.scan_lock:
            with self.lock:
                self.scanning = True
                old = dict(self.records)
            seen, updates, errors = set(), {}, []
            try:
                if not self.root.is_dir():
                    raise ValueError('资料文件夹不存在，请恢复文件夹后重新扫描。')
                def walk_error(exc):
                    errors.append(str(exc))
                for folder, dirs, names in os.walk(self.root, followlinks=False, onerror=walk_error):
                    dirs[:] = [d for d in dirs if not d.startswith('.') and not (Path(folder)/d).is_symlink() and not ((Path(folder)/d).stat().st_file_attributes & 1024 if os.name == 'nt' else False)]
                    for name in names:
                        if name.startswith(('~$', '.', '~')) or name.lower() in {'thumbs.db', 'desktop.ini'}:
                            continue
                        path = (Path(folder)/name).resolve()
                        if not path.is_relative_to(self.root):
                            continue
                        key = str(path)
                        seen.add(key)
                        previous = old.get(key, {})
                        try:
                            stat = path.stat()
                            stamp = [stat.st_mtime_ns, stat.st_size]
                            if previous.get('stamp') == stamp and previous.get('availability') == 'present':
                                continue
                            sha = fingerprint(path)
                            warning = ''
                            try:
                                content = sample_text(path)
                                if not content:
                                    warning = '未提取到文字；分类仅参考文件名和路径，扫描件可打开原件查看。'
                            except Exception as exc:
                                content, warning = '', '文字预览未完成：' + str(exc)[:240]
                            if [path.stat().st_mtime_ns, path.stat().st_size] != stamp:
                                raise OSError('文件正在复制或保存，下一轮扫描重试')
                            relative = str(path.relative_to(self.root))
                            category, reason = classify(relative, content)
                            updates[key] = {'path': key, 'relative': relative, 'name': path.name, 'ext': path.suffix.lower(),
                                'package': path.relative_to(self.root).parts[0] if len(path.relative_to(self.root).parts) > 1 else '根目录',
                                'size': stat.st_size, 'modified': datetime.fromtimestamp(stat.st_mtime, timezone.utc).isoformat(),
                                'stamp': stamp, 'sha256': sha, 'category': category, 'reason': reason,
                                'snippet': content[:1200], 'read_warning': warning, 'availability': 'present', 'seen_at': now()}
                        except OSError as exc:
                            updates[key] = {**previous, 'path': key, 'relative': str(path.relative_to(self.root)),
                                'name': path.name, 'ext': path.suffix.lower(), 'availability': 'unreadable',
                                'read_warning': str(exc), 'stamp': None}
                with self.lock, closing(sqlite3.connect(self.db)) as con, con:
                    for key, item in updates.items():
                        current = self.records.get(key, {})
                        item = {**current, **item}
                        changed = current.get('sha256') != item.get('sha256')
                        if changed:
                            item['previous_sha256'] = current.get('sha256')
                            item['status'] = 'pending'
                        if current.get('category_source') == 'manual':
                            item.update(category=current['category'], reason='手动分类', category_source='manual')
                        else:
                            item['category_source'] = 'automatic'
                        item.setdefault('status', 'pending')
                        item.setdefault('category', 'other')
                        item.setdefault('first_seen', now())
                        self.records[key] = item
                    if not errors:
                        for key, item in self.records.items():
                            if key not in seen:
                                self.records[key] = {**item, 'availability': 'missing'}
                    con.executemany('INSERT OR REPLACE INTO file_library VALUES (?,?)',
                        [(key, json.dumps(item, ensure_ascii=False)) for key, item in self.records.items()])
                    self.last_scan = now()
            except Exception as exc:
                errors.append(str(exc))
            finally:
                with self.lock:
                    self.scanning, self.errors = False, errors

    def snapshot(self):
        with self.lock:
            rows = []
            for item in self.records.values():
                sha = item.get('sha256')
                change = ('updated' if item.get('previous_sha256') else 'new') if sha and item.get('ack_sha256') != sha else ''
                rows.append({**item, 'change': change, 'category_confirmed': bool(sha and item.get('category_confirmed_sha256') == sha),
                             'can_use_tb': item['availability'] == 'present' and item['ext'] in {'.xlsx', '.xls', '.xlsm'}})
            rows.sort(key=lambda r: (r['availability'] == 'missing', not bool(r['change']), r['relative'].lower()))
            return {'root': str(self.root), 'files': rows, 'categories': CATEGORIES, 'statuses': STATUSES,
                    'scanning': self.scanning, 'last_scan': self.last_scan, 'errors': list(self.errors)}

    def update(self, payload):
        key = str(Path(payload['path']).resolve())
        with self.lock:
            if key not in self.records:
                raise ValueError('资料不存在，请刷新资料库。')
            item = dict(self.records[key])
            if item['availability'] != 'present' or not item.get('sha256') or payload.get('sha256') != item['sha256']:
                raise ValueError('资料已变更或不可读取，请刷新后重新确认。')
            path = self.resolve(key)
            if fingerprint(path) != item['sha256']:
                raise ValueError('资料已更新，请重新扫描后再保存状态。')
            if 'category' in payload:
                if payload['category'] not in CATEGORIES:
                    raise ValueError('未知资料类别。')
                item.update(category=payload['category'], category_source='manual', reason='手动分类', category_confirmed_sha256=item['sha256'])
            if 'status' in payload:
                if payload['status'] not in STATUSES:
                    raise ValueError('未知处理状态。')
                item.update(status=payload['status'], ack_sha256=item['sha256'])
            with closing(sqlite3.connect(self.db)) as con, con:
                con.execute('INSERT OR REPLACE INTO file_library VALUES (?,?)', (key, json.dumps(item, ensure_ascii=False)))
            self.records[key] = item
        return self.snapshot()

    def resolve(self, raw):
        path = Path(raw).resolve()
        if not path.is_relative_to(self.root) or not path.is_file():
            raise ValueError('只能打开资料库中的现有文件。')
        return path
