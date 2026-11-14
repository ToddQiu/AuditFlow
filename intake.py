"""Read-only customer balance imports, explicit TB mapping and detail tie-outs.

The TB here remains an unadjusted management draft. No audit adjustment,
foreign-exchange rate, missing ledger period or balancing account is invented.
"""
from __future__ import annotations
import calendar
import copy
import hashlib
import re
import uuid
from collections import defaultdict
from datetime import date
from decimal import Decimal

from engine import _dates_in_text, _decimal

CONTROL = re.compile(r'^(?:grand\s+total|sub\s*total|total(?:\b|\s)|合计|总计|小计|本期合计|本年累计)', re.I)
RATIO = re.compile(r'%|百分比|比率|毛利率|净利率|\bratio\b|\bmargin\s*%\b', re.I)


def context(raw):
    entity = str(raw.get('entity', '')).strip()
    if not entity or len(entity) > 300:
        raise ValueError('请填写本次客户实体，单体和集团不能混在同一份 TB。')
    try:
        start, end = date.fromisoformat(raw['period_start']), date.fromisoformat(raw['period_end'])
    except (KeyError, TypeError, ValueError):
        raise ValueError('请填写实际报告起止日期。')
    if start > end:
        raise ValueError('报告起始日不能晚于期末日。')
    currency = raw.get('currency')
    if currency not in ['HKD', 'RMB', 'USD', 'EUR', 'GBP', 'OTHER']:
        raise ValueError('请明确本位／取数币种，初版不自动换汇。')
    return {'entity': entity, 'period_start': str(start), 'period_end': str(end), 'currency': currency}


def _column(value):
    value = str(value or '').strip().upper()
    if not re.fullmatch(r'[A-Z]{1,3}', value):
        raise ValueError('列字段填写 Excel 列字母，例如 A 或 K。')
    number = 0
    for c in value:
        number = number * 26 + ord(c) - 64
    if number > 16384:
        raise ValueError('列字段超出 Excel 范围。')
    return value


def _blank(cell):
    return cell.get('status') == 'missing' or cell.get('status') == 'text' and not str(cell.get('value', '')).strip()


def _text(book, sheet, column, row):
    cell = book.cell(sheet, column + str(row))
    if cell.get('status') == 'error':
        raise ValueError(f'{sheet}!{column}{row} 是错误值。')
    return str(cell.get('value') if cell.get('value') is not None else '').strip()


def preview_import(spec, *, get_book, amount_reader):
    ctx = context(spec.get('context', {}))
    role = spec.get('source_role')
    if role not in ['closing_balances', 'support_details']:
        raise ValueError('TB 需要期末余额，不能把 ledger 发生额直接当期末；缺期初或期间资料须先补齐。')
    if spec.get('currency_basis') not in ['functional', 'original']:
        raise ValueError('请选择所取金额属于本位币还是原币，不能混取重复计量层。')
    unit = spec.get('unit')
    if isinstance(unit, bool) or unit not in [1, 1000, 1000000]:
        raise ValueError('请明确来源金额单位。')
    mode = spec.get('amount_mode')
    if mode not in ['signed', 'debit_credit']:
        raise ValueError('请选择借方为正的净额或借贷两栏。')
    sign = spec.get('sign_convention')
    if sign not in ['debit_positive', 'credit_positive']:
        raise ValueError('请明确净额的借贷正负方向。')
    credit_sign = spec.get('credit_sign', 'positive')
    if credit_sign not in ['positive', 'negative']:
        raise ValueError('请明确贷方栏本身为正数还是负数。')
    first, last = spec.get('first_row'), spec.get('last_row')
    if any(isinstance(x, bool) or not isinstance(x, int) for x in (first, last)) or not 1 <= first <= last <= 1048576 or last-first >= 3000:
        raise ValueError('请选择最多 3,000 行的明确来源范围；大表可分批导入。')
    columns = {k: _column(v) for k, v in spec.get('columns', {}).items() if v}
    fields = ['name', 'amount'] if mode == 'signed' else ['name', 'debit', 'credit']
    if not all(k in columns for k in fields):
        raise ValueError('请填写名称列及所需金额列。科目代码列可留空，之后明确映射 TB。')
    if len(set(columns.values())) != len(columns):
        raise ValueError('不同字段不能使用同一个来源列。')
    entry = get_book(spec['path'])
    obj, sheet = entry['book'], spec['sheet']
    if sheet not in obj.sheets:
        raise ValueError('来源工作表不存在。')
    if role == 'closing_balances' and re.search(r'general\s*ledger|journal|序时簿|总账流水', sheet, re.I):
        raise ValueError('所选表是总账流水，须先取得期初和完整报告期间并形成余额，不能直接当期末 TB。')
    headers = obj.rows(sheet, 1, min(10, max(1, first-1)))
    if first > 1:
        date_headers = obj.rows(sheet, 1, min(10, first-1))
        full_dates = set()
        for header in date_headers:
            for cell in header['cells'].values():
                text = str(cell.get('value') or '')
                full_dates.update(_dates_in_text(text))
                for year, month in re.findall(r'(20\d{2})年\s*(\d{1,2})月(?!\s*\d+日)', text):
                    if 1 <= int(month) <= 12:
                        full_dates.add(date(int(year), int(month), calendar.monthrange(int(year), int(month))[1]))
        if full_dates and date.fromisoformat(ctx['period_end']) not in full_dates:
            raise ValueError('来源表头显示的日期与本次期末不一致：' + ', '.join(map(str, sorted(full_dates))))
        amount_columns = [columns[k] for k in ['amount', 'debit', 'credit'] if k in columns]
        for col in amount_columns:
            if role == 'closing_balances':
                labels = [_text(obj, sheet, col, r) for r in range(1, min(first, 11))]
                if any(re.search(r'发生额|本期借方累计|本期贷方累计|本年借方累计|本年贷方累计|turnover|movement|transactions', label, re.I) for label in labels):
                    raise ValueError(f'{col} 列标示发生额，不能当作期末余额。')
            years = {_text(obj, sheet, col, r) for r in range(1, min(first, 11))}
            years = {int(v) for v in years if re.fullmatch(r'20\d{2}', v)}
            if years and date.fromisoformat(ctx['period_end']).year not in years:
                raise ValueError(f'{col} 列年头属于 {sorted(years)}，不能取作本期。')
    lines = []
    for row in range(first, last+1):
        row_id = hashlib.sha256(f"{entry['sha256']}|{sheet}|{row}".encode()).hexdigest()[:24]
        line = {'row': row, 'row_id': row_id, 'source_code': '', 'name': '', 'amount': None,
                'refs': [], 'needs_review': False, 'error': '', 'warnings': [], 'parent_children': []}
        try:
            line['source_code'] = _text(obj, sheet, columns['code'], row) if 'code' in columns else ''
            line['name'] = _text(obj, sheet, columns['name'], row)
            if not line['name']:
                raise ValueError('科目／明细名称为空；请排除空行或补齐真实名称。')
            if CONTROL.search(line['name']) or RATIO.search(line['name']):
                raise ValueError('这是合计、控制或比率行，不能与明细重复计入 TB／breakdown。')
            if re.fullmatch(r'20\d{2}', line['name']):
                raise ValueError('这是年度表头，不能作为科目。')
            if 'currency' in columns:
                row_currency = _text(obj, sheet, columns['currency'], row).upper()
                line['source_currency'] = row_currency
                if spec['currency_basis'] == 'original' and row_currency != ctx['currency']:
                    raise ValueError('原币列币种为空或不同，不能与本次取数币种混用。')
                if spec['currency_basis'] == 'functional' and row_currency and row_currency != ctx['currency']:
                    line['warnings'].append('该行标示其他原币；需确认所选金额栏确为本位币金额。')
            refs, blanks = [], []
            if mode == 'signed':
                refs = [{'sheet': sheet, 'address': columns['amount']+str(row), 'coefficient': 1 if sign == 'debit_positive' else -1}]
            else:
                for key, factor in [('debit', 1), ('credit', -1 if credit_sign == 'positive' else 1)]:
                    address = columns[key]+str(row)
                    cell = obj.cell(sheet, address)
                    if _blank(cell):
                        blanks.append(address)
                        continue
                    if key == 'credit' and cell.get('numeric') is not None:
                        credit = _decimal(cell['numeric'])
                        if credit_sign == 'positive' and credit < 0 or credit_sign == 'negative' and credit > 0:
                            raise ValueError('贷方栏的实际正负与所选约定不一致；不得自动 ABS。')
                    refs.append({'sheet': sheet, 'address': address, 'coefficient': factor})
                if not refs:
                    raise ValueError('借贷两边均为空，缺资料不能当作零。')
                if blanks and spec.get('blank_sides_are_zero') is not True:
                    raise ValueError('一边借／贷为空；如客户格式以空白表示另一边为零，请先明确确认该格式。')
            result = amount_reader({'source_path': entry['path'], 'refs': refs, 'unit_divisor': 1,
                                    'places': 2, 'confirm_saved_values': spec.get('confirm_saved_values') is True},
                                   preview=True, entry=entry)
            line['amount'] = str(_decimal(result['exact_amount']) * Decimal(unit))
            line['refs'] = result['source']['refs']
            line['needs_review'] = result['needs_review']
            line['verification_basis'] = result['source']['value_basis']
            line['explicit_blank_zero_cells'] = blanks
        except (ValueError, KeyError) as exc:
            line['error'] = str(exc)
        lines.append(line)
    codes = {line['source_code'] for line in lines if line['source_code'] and not line['error']}
    for line in lines:
        code = line['source_code']
        if code:
            line['parent_children'] = sorted(c for c in codes if c != code and c.startswith(code))
            if line['parent_children']:
                line['warnings'].append('代码疑似包含父子层级，请核实后只选择一层明细。')
    return {'sha256': entry['sha256'], 'path': entry['path'], 'sheet': sheet, 'context': ctx,
            'source_unit': unit, 'headers': headers, 'lines': lines, 'issues': [],
            'verification_status': 'needs_review' if any(l['needs_review'] for l in lines) else 'source_values_only'}


def _fresh(draft, hash_reader):
    stale = []
    if draft:
        for source in draft.get('imports', []) + draft.get('detail_imports', []):
            try:
                if hash_reader(source['path']) != source['sha256']:
                    stale.append({'path': source['path'], 'reason': '来源内容已改变，请重新导入并核对映射。'})
            except (OSError, ValueError) as exc:
                stale.append({'path': source['path'], 'reason': str(exc)})
    return stale


def _select(spec, preview):
    if spec.get('metadata_confirmed') is not True or spec.get('selection_confirmed') is not True:
        raise ValueError('请明确确认实体、期间、币种、单位、正负方向，以及只选择一层真实明细。')
    if spec.get('sha256') != preview['sha256']:
        raise ValueError('预览后来源文件已改变，请重新预览。')
    selected = spec.get('selected_rows', [])
    if not selected or any(isinstance(r, bool) or not isinstance(r, int) for r in selected) or len(set(selected)) != len(selected):
        raise ValueError('请选择不重复的来源明细行。')
    by_row = {line['row']: line for line in preview['lines']}
    if any(r not in by_row for r in selected):
        raise ValueError('选定行不在已预览的范围内。')
    lines = [by_row[r] for r in selected]
    for line in lines:
        if line['error']:
            raise ValueError(f"第 {line['row']} 行：{line['error']}")
        if line['needs_review'] and spec.get('confirm_saved_values') is not True:
            raise ValueError('部分保存值未独立重算，请先核实，再明确确认使用保存值。')
    selected_codes = {l['source_code'] for l in lines if l['source_code']}
    for line in lines:
        if selected_codes.intersection(line['parent_children']):
            raise ValueError('父科目与子科目被同时选中，会重复计算；请选择一层明细。')
    return copy.deepcopy(lines)


def _source_record(spec, preview):
    return {'id': uuid.uuid4().hex, 'path': preview['path'], 'sha256': preview['sha256'],
            'sheet': preview['sheet'], 'context': preview['context'], 'source_unit': spec['unit'],
            'spec': copy.deepcopy(spec), 'selected_rows': list(spec['selected_rows'])}


def import_rows(draft, spec, *, get_book, amount_reader, hash_reader):
    if spec.get('source_role') != 'closing_balances':
        raise ValueError('TB 只导入期末余额，不将发生额直接当期末。')
    if _fresh(draft, hash_reader):
        raise ValueError('已有来源已改变，旧 TB 草稿需要重新建立。')
    preview = preview_import(spec, get_book=get_book, amount_reader=amount_reader)
    lines = _select(spec, preview)
    draft = copy.deepcopy(draft) if draft else {'id': uuid.uuid4().hex, 'context': preview['context'],
                        'imports': [], 'lines': [], 'mappings': [], 'details': [], 'detail_imports': []}
    if draft['context'] != preview['context']:
        raise ValueError('实体、报告期间或币种与已有 TB 不一致；请使用同一项目资料。')
    existing = {l['row_id'] for l in draft['lines']}
    if existing.intersection(l['row_id'] for l in lines):
        raise ValueError('同一来源行已经导入，不能重复计入 TB。')
    existing_codes = {l['source_code'] for l in draft['lines'] if l.get('source_code') and
                      l.get('source_sha256') == preview['sha256'] and l.get('source_sheet') == preview['sheet']}
    for line in lines:
        code = line['source_code']
        if code and any(old != code and (old.startswith(code) or code.startswith(old)) for old in existing_codes):
            raise ValueError('已有导入含疑似父／子科目，不能分批重复计入同一层级余额。')
    source = _source_record(spec, preview)
    for line in lines:
        line['import_id'] = source['id']
        line['source_path'], line['source_sha256'], line['source_sheet'] = preview['path'], preview['sha256'], preview['sheet']
    draft['imports'].append(source)
    draft['lines'].extend(lines)
    if hash_reader(preview['path']) != preview['sha256']:
        raise ValueError('读取期间来源正在保存，请保存完成后重试。')
    return draft


def apply_mappings(draft, payload):
    if not draft:
        raise ValueError('先导入客户余额明细。')
    by_id = {line['row_id']: line for line in draft['lines']}
    output, seen, names, leads = [], set(), {}, {}
    for raw in payload.get('mappings', []):
        row_id = raw.get('row_id')
        if row_id not in by_id or row_id in seen:
            raise ValueError('映射行不存在或重复。')
        seen.add(row_id)
        code = str(raw.get('account_code', '')).strip()
        name = str(raw.get('account_name', '')).strip()
        lead = str(raw.get('lead_code', '')).strip()
        title = str(raw.get('lead_title', '')).strip()
        if not code:
            continue
        if not name or len(code) > 100 or len(name) > 300 or len(lead) > 100 or len(title) > 300:
            raise ValueError('请明确 TB 科目名称，字段长度不超过合理范围。')
        if bool(lead) != bool(title):
            raise ValueError('底稿编号与实际标题须同时填写，不凭编号猜类别。')
        identity = (name, lead, title)
        if code in names and names[code] != identity:
            raise ValueError('同一 TB 科目的名称或底稿映射不一致。')
        if lead in leads and lead and leads[lead] != title:
            raise ValueError('同一底稿编号存在不同标题，请核实项目模板。')
        names[code], leads[lead] = identity, title
        output.append({'row_id': row_id, 'account_code': code, 'account_name': name,
                       'lead_code': lead, 'lead_title': title})
    draft = copy.deepcopy(draft)
    draft['mappings'] = output
    return draft


def add_support(draft, payload, *, get_book, amount_reader, hash_reader):
    if not draft:
        raise ValueError('先建立 TB 草稿及科目映射。')
    if _fresh(draft, hash_reader):
        raise ValueError('来源已改变，请重新建立 TB／明细。')
    code, label = str(payload.get('account_code', '')).strip(), str(payload.get('label', '')).strip()
    if not label or len(label) > 300:
        raise ValueError('请明确客户明细的内容，例如设备清单原价或贷款合同余额。')
    if code not in {m['account_code'] for m in draft['mappings']}:
        raise ValueError('所选 TB 科目不存在，请先完成映射。')
    spec = copy.deepcopy(payload.get('spec', {}))
    for key in ['sha256', 'selected_rows', 'metadata_confirmed', 'selection_confirmed', 'confirm_saved_values']:
        if key in payload:
            spec[key] = payload[key]
    if spec.get('source_role') != 'support_details':
        raise ValueError('请选择客户真实明细来源，TB 科目汇总不能冒充明细 breakdown。')
    preview = preview_import(spec, get_book=get_book, amount_reader=amount_reader)
    if draft['context'] != preview['context']:
        raise ValueError('明细实体、期间或币种与 TB 不一致。')
    lines = _select(spec, preview)
    occupied = {l['row_id'] for l in draft['lines'] + draft['details']}
    if occupied.intersection(l['row_id'] for l in lines):
        raise ValueError('不能重复使用 TB 余额行或已导入明细来证明 breakdown。')
    record = _source_record(spec, preview)
    draft = copy.deepcopy(draft)
    for line in lines:
        line.update(account_code=code, detail_label=label, import_id=record['id'],
                    source_path=preview['path'], source_sha256=preview['sha256'], source_sheet=preview['sheet'])
    draft['details'].extend(lines)
    draft['detail_imports'].append(record)
    if hash_reader(preview['path']) != preview['sha256']:
        raise ValueError('读取期间明细来源正在保存，请保存完成后重试。')
    return draft


def summarize(draft, hash_reader=None):
    stale = _fresh(draft, hash_reader) if hash_reader else []
    tb = {'accounts': [], 'total_debit': '0', 'total_credit': '0', 'difference': '0',
          'unmapped': [], 'unassigned_leads': [], 'status': 'needs_review', 'issues': []}
    if not draft:
        tb['issues'] = ['尚未导入客户余额。']
        return {'draft': None, 'tb': tb, 'breakdowns': [], 'stale_sources': stale}
    mappings = {m['row_id']: m for m in draft['mappings']}
    accounts = {}
    for line in draft['lines']:
        m = mappings.get(line['row_id'])
        if not m:
            tb['unmapped'].append(line['row_id'])
            continue
        code = m['account_code']
        account = accounts.setdefault(code, {**{k: m[k] for k in ['account_code', 'account_name', 'lead_code', 'lead_title']},
                'management_amount': Decimal(0), 'source_rows': [], 'needs_review': False})
        account['management_amount'] += _decimal(line['amount'])
        account['source_rows'].append(line['row_id'])
        account['needs_review'] |= line['needs_review']
    debit = credit = Decimal(0)
    groups = {}
    detail_by_account = defaultdict(list)
    for line in draft['details']:
        detail_by_account[line['account_code']].append(line)
    for account in accounts.values():
        amount = account['management_amount']
        account.update(amount=str(amount), debit=str(max(amount, Decimal(0))), credit=str(max(-amount, Decimal(0))))
        account['management_amount'] = str(amount)
        debit += _decimal(account['debit'])
        credit += _decimal(account['credit'])
        if not account['lead_code']:
            tb['unassigned_leads'].append(account['account_code'])
            continue
        group = groups.setdefault(account['lead_code'], {'lead_code': account['lead_code'], 'lead_title': account['lead_title'],
            'tb_total': Decimal(0), 'account_lines': [], 'details': [], 'detail_total': Decimal(0),
            'missing_details': [], 'difference': None, 'status': 'needs_review', 'reason': ''})
        details = detail_by_account[account['account_code']]
        detail_total = sum((_decimal(l['amount']) for l in details), Decimal(0)) if details else None
        diff = detail_total - amount if detail_total is not None else None
        group['tb_total'] += amount
        group['account_lines'].append({**account, 'detail_total': str(detail_total) if detail_total is not None else None,
                                     'difference': str(diff) if diff is not None else None})
        if not details:
            group['missing_details'].append(account['account_code'])
        else:
            group['details'].extend(details)
            group['detail_total'] += detail_total
    tb['accounts'] = list(accounts.values())
    tb.update(total_debit=str(debit), total_credit=str(credit), difference=str(debit-credit))
    if tb['unmapped']:
        tb['issues'].append(f"还有 {len(tb['unmapped'])} 个客户余额行未映射 TB。")
    if debit != credit:
        tb['issues'].append('借贷尚未平衡；差额原样保留，不新增凑数科目。')
    if tb['unassigned_leads']:
        tb['issues'].append('部分 TB 科目还未指定底稿编号及实际标题。')
    if any(account['needs_review'] for account in accounts.values()):
        tb['issues'].append('部分 TB 金额采用已确认的保存值，未独立重算，仍需复核。')
    if stale:
        tb['issues'].append('来源文件已改变，旧金额和映射需重新核对。')
    tb['status'] = 'difference' if debit != credit else 'needs_review' if tb['issues'] else 'draft_balanced'
    tb['issues'].append('这是管理账 TB 草稿，尚未包括审计调整／审定及完整性复核。')
    for group in groups.values():
        per_account_differences = [a['difference'] for a in group['account_lines'] if a['difference'] is not None and _decimal(a['difference']) != 0]
        if group['missing_details']:
            group['status'] = 'needs_review'
            group['reason'] = '仅建立科目余额汇总；缺客户明细，不能视为完整 breakdown。'
            group['detail_total'] = str(group['detail_total']) if group['details'] else None
        else:
            group['difference'] = str(group['detail_total'] - group['tb_total'])
            group['detail_total'] = str(group['detail_total'])
            group['status'] = 'difference' if per_account_differences else 'tied_needs_review'
            group['reason'] = '按每个 TB 科目分别核对客户明细；合计对上仍需资产变动、合同及其他业务复核。'
        if stale:
            group['status'] = 'needs_review'
            group['reason'] = '来源已改变，不能继续使用旧核对结果。'
        elif any(a['needs_review'] for a in group['account_lines']) or any(l['needs_review'] for l in group['details']):
            if group['status'] != 'difference': group['status'] = 'needs_review'
            group['reason'] += ' 部分来源只有保存值，未独立重算。'
        group['tb_total'] = str(group['tb_total'])
    orphaned = [d['account_code'] for d in draft['details'] if d['account_code'] not in accounts]
    if orphaned:
        tb['status'] = 'needs_review'
        tb['issues'].append('旧明细对应的科目已被移除／改码，请重新核对。')
    return {'draft': draft, 'tb': tb, 'breakdowns': list(groups.values()), 'stale_sources': stale}
