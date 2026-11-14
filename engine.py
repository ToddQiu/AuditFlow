"""Read-only financial-data inspection. No workbook or report authoring.

Saved Excel results remain evidence, not proof of recalculation. Only the small
explicit formula subset below is independently evaluated; other formulas are
shown as saved values with a needs-review result.
"""
from __future__ import annotations

from decimal import Decimal, InvalidOperation, ROUND_HALF_UP, localcontext
from datetime import date
from pathlib import Path
import importlib.util
import re
from typing import Any

from openpyxl.formula import Tokenizer
from openpyxl.formula.translate import Translator
from openpyxl.utils.cell import range_boundaries, get_column_letter


def _decimal(value: Any) -> Decimal:
    if isinstance(value, dict):
        if value.get("status") in {"missing", "error", "text"}:
            raise ValueError("Missing, error or text is not a monetary amount")
        value = value.get("numeric", value.get("value"))
    if value is None or isinstance(value, bool) or value == "":
        raise ValueError("A blank amount is not zero")
    try:
        amount = value if isinstance(value, Decimal) else Decimal(str(value))
    except (InvalidOperation, ValueError, TypeError) as exc:
        raise ValueError(f"Not a finite monetary amount: {value!r}") from exc
    if not amount.is_finite():
        raise ValueError("A monetary amount must be finite")
    return amount


def round_money(value: Any, places: int = 0) -> Decimal:
    """Decimal rounding, half away from zero (Excel ROUND monetary convention)."""
    if isinstance(places, bool) or not isinstance(places, int) or abs(places) > 28:
        raise ValueError("places must be an integer between -28 and 28")
    amount = _decimal(value)
    with localcontext() as context:
        context.prec = max(50, len(amount.as_tuple().digits) + abs(places) + 5)
        result = amount.quantize(Decimal(1).scaleb(-places), rounding=ROUND_HALF_UP)
    return abs(result) if result == 0 else result


def format_money(value: Any, places: int = 0) -> str:
    """Report-style comma separators and parentheses; blank remains blank."""
    if value is None or value == "" or isinstance(value, dict) and value.get("status") == "missing":
        return ""
    rounded = round_money(value, places)
    number = f"{abs(rounded):,.{max(0, places)}f}"
    return f"({number})" if rounded < 0 else number


def _shown(value: Any) -> str:
    if isinstance(value, dict):
        value = value.get("value")
    return "" if value is None else str(value)


def _evidence(value: Any, label: str) -> dict:
    if isinstance(value, dict):
        return {"sheet": value.get("sheet", ""), "address": value.get("address", label), "value": _shown(value)}
    return {"sheet": "", "address": label, "value": _shown(value)}


def _check(title: str, status: str, difference: Any = "", evidence=None, reason="", **extra) -> dict:
    return {"title": title, "status": status, "difference": str(difference),
            "evidence": evidence or [], "reason": reason, **extra}


def sum_check(values, reported) -> dict:
    """Compare an explicitly selected set of amounts with its reported total."""
    values = list(values)
    evidence = [_evidence(value, f"item {i+1}") for i, value in enumerate(values)]
    evidence.append(_evidence(reported, "reported total"))
    try:
        if not values:
            raise ValueError("No source items selected")
        total = sum((_decimal(value) for value in values), Decimal(0))
        difference = total - _decimal(reported)
    except ValueError as exc:
        return _check("Selected-cell sum", "needs_review", evidence=evidence, reason=str(exc))
    return _check("Selected-cell sum", "pass" if difference == 0 else "difference", difference,
                  evidence, "Independent sum of the selected amounts; selection completeness still requires review.", total=str(total))


def rounding_check(values, places: int = 0) -> dict:
    """Distinguish sum-of-rounded items from a rounded exact total."""
    values = list(values)
    evidence = [_evidence(value, f"item {i+1}") for i, value in enumerate(values)]
    try:
        if not values:
            raise ValueError("No source items selected")
        amounts = [_decimal(value) for value in values]
        total = sum(amounts, Decimal(0))
        rounded_total = round_money(total, places)
        sum_of_rounded = sum((round_money(amount, places) for amount in amounts), Decimal(0))
    except ValueError as exc:
        return _check("Display rounding", "needs_review", evidence=evidence, reason=str(exc))
    difference = sum_of_rounded - rounded_total
    return _check("Display rounding", "pass" if difference == 0 else "difference", difference, evidence,
                  "No rounding allocation has been applied; a difference requires an approved display adjustment.",
                  exact_total=str(total), rounded_total=str(rounded_total), sum_of_rounded=str(sum_of_rounded))


def cashflow_check(opening, operating, investing, financing, fx, closing, bs_cash) -> dict:
    """Require both a cash-flow roll-forward and an independent BS cash tie."""
    labels = ("opening", "operating", "investing", "financing", "fx", "closing", "BS cash")
    inputs = (opening, operating, investing, financing, fx, closing, bs_cash)
    evidence = [_evidence(value, label) for value, label in zip(inputs, labels)]
    try:
        amounts = [_decimal(value) for value in inputs]
    except ValueError as exc:
        return _check("Cash flow closing and BS tie", "needs_review", evidence=evidence, reason=str(exc))
    o, op, inv, fin, exchange, close, bs = amounts
    calculated = o + op + inv + fin + exchange
    flow_diff, bs_diff = calculated - close, close - bs
    status = "pass" if flow_diff == 0 and bs_diff == 0 else "difference"
    return _check("Cash flow closing and BS tie", status, flow_diff if flow_diff else bs_diff, evidence,
                  "Opening + operating + investing + financing + FX must equal closing, and closing must agree to BS cash. This does not validate cash/non-cash classification.",
                  calculated_closing=str(calculated), flow_difference=str(flow_diff), bs_difference=str(bs_diff))


class FormulaReview(ValueError):
    """A formula cannot be independently calculated by this limited engine."""


def _dates_in_text(text: str) -> set[date]:
    """Parse explicit full dates only; never infer a month/day from a year."""
    found = set()
    months = {name: index for index, name in enumerate(
        ("jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"), 1)}
    month_pattern = r"Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:t(?:ember)?)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?"

    def add(year, month, day):
        try:
            found.add(date(int(year), int(month), int(day)))
        except ValueError:
            pass  # An impossible date is not evidence of the reporting period.

    for match in re.finditer(r"(?<!\d)((?:19|20)\d{2})\s*[-/.]\s*(\d{1,2})\s*[-/.]\s*(\d{1,2})(?!\d)", text):
        add(*match.groups())
    for match in re.finditer(r"(?<!\d)(\d{1,2})\s*[-/.]\s*(\d{1,2})\s*[-/.]\s*((?:19|20)\d{2})(?!\d)", text):
        day, month, year = match.groups()
        add(year, month, day)
    for match in re.finditer(r"(?<!\d)(\d{1,2})\s*(" + month_pattern + r")\.?\s*((?:19|20)\d{2})(?!\d)", text, re.I):
        day, month, year = match.groups()
        add(year, months[month[:3].lower()], day)
    for match in re.finditer(r"(" + month_pattern + r")\.?\s*(\d{1,2})\s*,?\s*((?:19|20)\d{2})(?!\d)", text, re.I):
        month, day, year = match.groups()
        add(year, months[month[:3].lower()], day)
    for match in re.finditer(r"(?<!\d)((?:19|20)\d{2})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日", text):
        add(*match.groups())
    return found


class _Range:
    def __init__(self, sheet, addresses):
        self.sheet, self.addresses = sheet, addresses


def _reader():
    # Load the existing read-only extractor without executing its CLI main().
    path = Path(__file__).parent / "research" / "read_corpus.py"
    spec = importlib.util.spec_from_file_location("hk_audit_corpus_reader", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class SourceBook:
    """A source workbook, with exact sheet/address access and cached evidence."""
    def __init__(self, path=None, *, data=None):
        self.path = str(Path(path).resolve()) if path is not None else "<in-memory test data>"
        self.data = data if data is not None else _reader().read_file(path)
        if self.data.get("kind") != "workbook":
            raise ValueError("SourceBook requires an Excel workbook")
        if self.data.get("metadata_only"):
            raise ValueError("大 XML 工作簿仅取得工作表清单，未读取金额；需要专门流式读取后才能作为数据来源 (metadata_only).")
        self.legacy_cached_values = self.data.get("formula_values") == "legacy_formula_text_unavailable"
        self.sheets = [sheet["name"] for sheet in self.data["sheets"]]
        if len(set(self.sheets)) != len(self.sheets):
            raise ValueError("Duplicate worksheet names")
        self._cells = {}
        self._by_sheet = {}
        for sheet in self.data["sheets"]:
            self._by_sheet[sheet["name"]] = {}
            anchors = {}
            for cell in sheet["cells"]:
                formula = cell.get("formula")
                if isinstance(formula, dict) and formula.get("type") == "shared" and formula.get("text"):
                    anchors[formula.get("shared")] = (cell["address"], formula["text"])
            for cell in sheet["cells"]:
                formula = cell.get("formula")
                if isinstance(formula, dict) and formula.get("type") == "shared" and not formula.get("text"):
                    anchor = anchors.get(formula.get("shared"))
                    formula["original_shared_text"] = formula.get("text", "")
                    if anchor:
                        try:
                            translated = Translator("=" + anchor[1].lstrip("="), origin=anchor[0]).translate_formula(cell["address"])
                            formula.update(text=translated.lstrip("="), shared_expanded=True)
                        except Exception as exc:
                            formula["translation_error"] = str(exc)
                    else:
                        formula["translation_error"] = "Shared-formula anchor missing"
                self._cells[(sheet["name"], cell["address"].upper())] = cell
                self._by_sheet[sheet["name"]][cell["address"].upper()] = cell
        self._memo = {}

    @classmethod
    def from_data(cls, data):
        return cls(data=data)

    def cell(self, sheet: str, address: str) -> dict:
        if sheet not in self.sheets:
            raise KeyError(f"Worksheet not found: {sheet}")
        address = str(address).upper().replace("$", "")
        if not re.fullmatch(r"[A-Z]{1,3}[1-9][0-9]*", address):
            raise ValueError(f"Use one exact A1-style cell address: {address}")
        try:
            bounds = range_boundaries(address)
            if bounds[0] > 16384 or bounds[1] > 1048576:
                raise ValueError("Cell is outside the Excel worksheet")
        except (TypeError, ValueError) as exc:
            raise ValueError(f"Invalid cell address: {address}") from exc
        source = self._cells.get((sheet, address), {"address": address, "value": None, "type": "n", "formula": None})
        result = dict(source, sheet=sheet, address=address, numeric=None)
        value, kind = source.get("value"), source.get("type", "n")
        if kind == "e":
            result["status"] = "error"
        elif value is None or value == "":
            result["status"] = "missing"
        elif kind in {"s", "str", "inlineStr", "b", "d"}:
            result["status"] = "text"
        else:
            try:
                result["numeric"] = _decimal(value)
                result["status"] = "number"
            except ValueError:
                result["status"] = "text"
        return result

    @staticmethod
    def _formula(cell):
        formula = cell.get("formula")
        if isinstance(formula, dict):
            return formula.get("text", "")
        return str(formula).lstrip("=") if formula else ""

    def rows(self, sheet: str, start: int = 1, end: int = 80) -> list[dict]:
        if sheet not in self.sheets:
            raise KeyError(f"Worksheet not found: {sheet}")
        if start < 1 or end < start or end > 1048576:
            raise ValueError("Invalid row range")
        grouped = {}
        for address, source in self._by_sheet[sheet].items():
            row = int(re.search(r"[0-9]+$", address).group())
            if start <= row <= end:
                grouped.setdefault(row, {})[address] = self.cell(sheet, address)
        result = []
        for row, cells in sorted(grouped.items()):
            cells = dict(sorted(cells.items(), key=lambda pair: range_boundaries(pair[0])[0]))
            labels = [str(cell["value"]) for cell in cells.values() if cell["status"] == "text" and not cell.get("formula")]
            result.append({"row": row, "label": " | ".join(labels), "cells": cells})
        return result

    def _reference(self, text: str, current_sheet: str) -> _Range:
        if "[" in text or "]" in text:
            raise FormulaReview(f"External workbook dependency: {text}")
        if "!" in text:
            sheet, address = text.rsplit("!", 1)
            if sheet.startswith("'") and sheet.endswith("'"):
                sheet = sheet[1:-1].replace("''", "'")
        else:
            sheet, address = current_sheet, text
        if sheet not in self.sheets:
            raise FormulaReview(f"Unresolved or 3-D worksheet reference: {sheet}")
        address = address.replace("$", "").upper()
        if not re.fullmatch(r"[A-Z]{1,3}[1-9][0-9]*(?::[A-Z]{1,3}[1-9][0-9]*)?", address):
            raise FormulaReview(f"Unsupported named, full-column or structured reference: {text}")
        left, top, right, bottom = range_boundaries(address)
        if right > 16384 or bottom > 1048576 or right < left or bottom < top:
            raise FormulaReview(f"Invalid Excel range: {text}")
        if (right-left+1)*(bottom-top+1) > 10000:
            raise FormulaReview("Range exceeds the safe inspection limit (10,000 cells)")
        return _Range(sheet, [f"{get_column_letter(column)}{row}"
                             for row in range(top, bottom+1) for column in range(left, right+1)])

    def references(self, sheet: str, address: str) -> list[dict]:
        cell = self.cell(sheet, address)
        text = self._formula(cell)
        result = []
        if not text:
            return result
        try:
            tokens = Tokenizer("=" + text).items
        except Exception as exc:
            return [{"reference": text, "status": "needs_review", "reason": str(exc)}]
        for token in tokens:
            if token.type == "OPERAND" and token.subtype == "RANGE":
                try:
                    ref = self._reference(token.value, sheet)
                    result.extend(dict(self.cell(ref.sheet, target), reference=token.value) for target in ref.addresses)
                except FormulaReview as exc:
                    result.append({"reference": token.value, "status": "needs_review", "reason": str(exc),
                                   "external": "[" in token.value, "sheet": "", "address": "", "value": None})
        return result

    def trace(self, sheet: str, address: str, recursive: bool = True, limit: int = 1000) -> list[dict]:
        """Trace actual formula dependencies, never formula-looking text cells."""
        self.cell(sheet, address)
        output, seen = [], set()
        def walk(name, target, ancestors):
            if len(ancestors) > 128:
                output.append({"sheet": name, "address": target, "value": None, "status": "needs_review",
                               "reason": "Dependency depth exceeds the safe inspection limit"})
                return
            for ref in self.references(name, target):
                if len(output) >= limit:
                    return
                key = (ref.get("sheet"), ref.get("address"))
                if ref.get("status") == "needs_review":
                    output.append(ref)
                elif key in ancestors:
                    output.append(dict(ref, status="needs_review", reason="Circular dependency"))
                elif key not in seen:
                    seen.add(key)
                    output.append(ref)
                    if recursive and ref.get("formula"):
                        walk(key[0], key[1], ancestors | {key})
        walk(sheet, address, {(sheet, address.upper().replace("$", ""))})
        return output

    def evaluate(self, sheet: str, address: str, _stack=frozenset()) -> Decimal:
        """Independently evaluate only refs, SUM, ROUND and scalar arithmetic."""
        cell = self.cell(sheet, address)
        if self.legacy_cached_values:
            raise FormulaReview("Legacy XLS formula text is unavailable; a saved numeric value cannot be independently verified as a constant or recalculated formula")
        key = (sheet, cell["address"])
        if key in _stack:
            raise FormulaReview("Circular dependency")
        if len(_stack) > 128:
            raise FormulaReview("Dependency depth exceeds the safe inspection limit")
        if key in self._memo:
            return self._memo[key]
        if cell["status"] == "error":
            raise FormulaReview(f"Saved formula error at {sheet}!{address}: {cell['value']}")
        if cell.get("formula"):
            text = self._formula(cell)
            if not text:
                raise FormulaReview("Formula text unavailable; only a saved value is available")
            formula = cell.get("formula")
            if isinstance(formula, dict) and formula.get("type") in {"array", "dataTable"}:
                raise FormulaReview("Array and data-table formulas are not independently evaluated")
            try:
                value = _FormulaParser(self, sheet, text, _stack | {key}).parse()
            except FormulaReview:
                raise
            except (ArithmeticError, ValueError, TypeError) as exc:
                raise FormulaReview(str(exc)) from exc
        else:
            if cell["status"] != "number":
                raise FormulaReview(f"No numeric source value at {sheet}!{address} ({cell['status']})")
            value = cell["numeric"]
        self._memo[key] = value
        return value

    def selected_sum(self, sheet: str, addresses, reported_address: str) -> dict:
        result = sum_check([self.cell(sheet, address) for address in addresses], self.cell(sheet, reported_address))
        if self.legacy_cached_values:
            if result["status"] == "pass": result["status"] = "needs_review"
            result["reason"] += " Legacy XLS saved values only; source formulas are unavailable and have not been independently verified."
            result["recalculated"] = False
        return result

    def _current_column(self, sheet, candidates):
        """Detect a presentation column, not an authoritative reporting period."""
        columns = {range_boundaries(address)[0] for address in candidates}
        dated = []
        for row in self.rows(sheet, 1, 20):
            for address, cell in row["cells"].items():
                column = range_boundaries(address)[0]
                if column not in columns or cell["value"] is None:
                    continue
                text = str(cell["value"])
                if cell["status"] == "number" and re.fullmatch(r"(?:19|20)\d{2}", text):
                    dated.append((int(text), column))
                elif cell["status"] == "text":
                    matches = re.findall(r"(?<!\d)(?:19|20)\d{2}(?!\d)", text)
                    dated.extend((int(year), column) for year in matches)
        if dated:
            latest = max(year for year, _ in dated)
            matching = {column for year, column in dated if year == latest}
            if len(matching) == 1:
                return matching.pop(), f"Detected {latest} header; confirm it is the selected reporting period."
            raise FormulaReview("Multiple columns share the latest period; select an explicit column")
        if len(columns) == 1:
            return columns.pop(), "Only one amount column found; verify its period and currency."
        raise FormulaReview("Current-period column could not be determined from explicit headers")

    def _currency(self, sheet, column):
        found = set()
        for row in self.rows(sheet, 1, 20):
            for address, cell in row["cells"].items():
                if range_boundaries(address)[0] != column or cell["status"] != "text":
                    continue
                text = str(cell["value"]).upper()
                for pattern, code in ((r"HK\$|HKD", "HKD"), (r"RMB|CNY", "RMB"), (r"US\$|USD", "USD")):
                    if re.search(pattern, text):
                        found.add((code, "1000" if re.search(r"000|千", text) else "1"))
        return next(iter(found)) if len(found) == 1 else None

    def _column_year(self, sheet, column):
        """Require an unambiguous year header; matching amounts cannot prove it."""
        years = set()
        for row in self.rows(sheet, 1, 20):
            # Financial amounts coincidentally equal to 2025/2026 are not headers.
            if re.search(r"cash|operating|investing|financing|余额|餘額|现金|現金", row["label"], re.I):
                continue
            for address, cell in row["cells"].items():
                if range_boundaries(address)[0] != column or cell["value"] is None:
                    continue
                text = str(cell["value"])
                if cell["status"] == "number" and re.fullmatch(r"(?:19|20)\d{2}", text):
                    years.add(int(text))
                elif cell["status"] == "text":
                    years.update(int(year) for year in re.findall(r"(?<!\d)(?:19|20)\d{2}(?!\d)", text))
        if not years:
            raise FormulaReview(f"Reporting period is not explicit for {sheet}!{get_column_letter(column)}")
        if len(years) != 1:
            raise FormulaReview(f"Conflicting reporting-period labels for {sheet}!{get_column_letter(column)}: {sorted(years)}")
        return next(iter(years))

    def _column_end_date(self, sheet, column, year):
        """Require one explicit full end date for the selected year/column."""
        dates = set()
        for row in self.rows(sheet, 1, 20):
            if re.search(r"cash|operating|investing|financing|余额|餘額|现金|現金", row["label"], re.I):
                continue
            for address, cell in row["cells"].items():
                if cell["status"] != "text" or cell.get("formula"):
                    continue
                text = str(cell["value"])
                selected_column = range_boundaries(address)[0] == column
                end_header = re.search(r"(?:year|period)\s*ended|as\s*(?:at|of)|期末|截至", text, re.I)
                if selected_column or end_header:
                    dates.update(_dates_in_text(text))
        matching = {value for value in dates if value.year == year}
        location = f"{sheet}!{get_column_letter(column)}"
        if not matching:
            detail = f"; other dates found: {sorted(value.isoformat() for value in dates)}" if dates else ""
            raise FormulaReview(f"A full reporting-period end date is not explicit for {location} ({year}){detail}; a year alone is insufficient")
        if len(matching) != 1:
            raise FormulaReview(f"Conflicting full reporting-period end dates for {location}: {sorted(value.isoformat() for value in matching)}")
        return next(iter(matching))

    def _cashflow_checks(self, sheet):
        if not re.search(r"cash\s*flow|\bCF\b|Y101|现金流|現金流", sheet, re.I):
            return []
        patterns = {
            "opening": r"cash.*(?:beginning|start)|as\s+at.*(?:20\d{2})|期初.*現?现金|期初.*現金",
            "operating": r"net\s+cash.*operating\s+activities|经营活动.*净|經營活動.*淨",
            "investing": r"net\s+cash.*investing\s+activities|投资活动.*净|投資活動.*淨",
            "financing": r"net\s+cash.*financing\s+activities|融资活动.*净|籌資活動.*淨",
            "fx": r"effect.*(?:foreign\s+exchange|exchange\s+rate)|汇率.*影响|匯率.*影響",
            "closing": r"cash.*(?:end\s+of|ending)|期末.*(?:现金|現金)",
        }
        detected = {}
        rows = self.rows(sheet, 1, 250)
        for row in rows:
            for key, pattern in patterns.items():
                if re.search(pattern, row["label"], re.I):
                    if key == "closing" and re.search(r"analysis|分析", row["label"], re.I):
                        continue
                    detected.setdefault(key, row)
        if "closing" not in detected:
            return []
        try:
            candidates = [a for a, c in detected["closing"]["cells"].items() if c["status"] == "number"]
            if not candidates:
                raise FormulaReview("Closing cash has no saved numeric amount")
            column, hint = self._current_column(sheet, candidates)
            cf_inputs = []
            for key in ("opening", "operating", "investing", "financing", "fx", "closing"):
                if key not in detected:
                    raise FormulaReview(f"Cash-flow {key} row is missing; no zero has been assumed")
                cf_inputs.append(self.cell(sheet, f"{get_column_letter(column)}{detected[key]['row']}"))
            bs_sheet = next((name for name in self.sheets if name.strip().upper() == "BS"), None)
            if bs_sheet is None:
                raise FormulaReview("A separate BS sheet is required for the closing-cash cross-check")
            cash_rows = [row for row in self.rows(bs_sheet, 1, 150)
                         if re.search(r"cash\s+and\s+(?:cash\s+equivalents|bank\s+balances)|(?:cash|bank)\s+balances|货币资金|現金及|现金及", row["label"], re.I)]
            if len(cash_rows) != 1:
                raise FormulaReview("BS cash row is missing or ambiguous; select it explicitly")
            bs_candidates = [a for a, c in cash_rows[0]["cells"].items() if c["status"] == "number"]
            bs_column, bs_hint = self._current_column(bs_sheet, bs_candidates)
            cf_year, bs_year = self._column_year(sheet, column), self._column_year(bs_sheet, bs_column)
            if cf_year != bs_year:
                raise FormulaReview(f"Cash flow period {cf_year} does not agree to BS period {bs_year}; matching cash amounts do not resolve the conflict")
            cf_date = self._column_end_date(sheet, column, cf_year)
            bs_date = self._column_end_date(bs_sheet, bs_column, bs_year)
            if cf_date != bs_date:
                raise FormulaReview(f"Cash flow end date {cf_date.isoformat()} does not agree to BS end date {bs_date.isoformat()}; matching years and cash amounts do not resolve the conflict")
            currency, bs_currency = self._currency(sheet, column), self._currency(bs_sheet, bs_column)
            if currency is None or bs_currency is None:
                raise FormulaReview("Cash flow / BS amount currency and unit must be explicit before comparing")
            if currency != bs_currency:
                raise FormulaReview("Cash flow and BS use different currency or display units; no conversion has been guessed")
            bs_cash = self.cell(bs_sheet, f"{get_column_letter(bs_column)}{cash_rows[0]['row']}")
            result = cashflow_check(*cf_inputs, bs_cash)
            result["period_end_date"] = cf_date.isoformat()
            result["reason"] += " Saved amounts only. " + hint + " " + bs_hint
            return [result]
        except (FormulaReview, ValueError) as exc:
            return [_check("Cash flow closing and BS tie", "needs_review", reason=str(exc))]

    def checks(self, sheet: str, limit: int | None = None) -> list[dict]:
        if sheet not in self.sheets:
            raise KeyError(f"Worksheet not found: {sheet}")
        if self.legacy_cached_values:
            return [_check(f"Legacy XLS saved values: {sheet}", "needs_review",
                           reason="Legacy XLS formula text is unavailable. Only saved numeric values were read; neither source constants nor formula calculations have been independently verified.", recalculated=False)]
        rows = self.rows(sheet, 1, 1048576)
        row_lookup = {row["row"]: row for row in rows}
        maximum = max(row_lookup, default=0)
        checks = self._cashflow_checks(sheet)
        for row in rows:
            for address, cell in row["cells"].items():
                if not cell.get("formula"):
                    continue
                text = self._formula(cell)
                evidence = [_evidence(cell, address)]
                is_control = bool(re.search(r"check|difference|control|差[額异異]|對數|对数", row["label"], re.I))
                # Common unnamed footer controls A-B; ordinary mid-table values
                # are not treated as zero-target controls.
                is_control |= (row["row"] >= maximum-2 and not row["label"] and
                               bool(re.fullmatch(r"\s*'?[^(),+*/]*!?\$?[A-Z]+\$?\d+\s*-\s*'?[^(),+*/]*!?\$?[A-Z]+\$?\d+\s*", text)))
                title = ("Zero-target control" if is_control else "Independent formula arithmetic") + f": {sheet}!{address}"
                try:
                    value = self.evaluate(sheet, address)
                    if is_control:
                        diff = value
                        reason = "Independently evaluated zero-target control; source selection still requires review."
                    elif cell["numeric"] is None:
                        checks.append(_check(title, "needs_review", evidence=evidence,
                                             reason="Calculated from supported formula, but saved numeric result is missing.", calculated=str(value)))
                        continue
                    else:
                        diff = value-cell["numeric"]
                        reason = "Independent supported-formula result compared with saved Excel result; not proof of mapping completeness."
                    checks.append(_check(title, "pass" if diff == 0 else "difference", diff, evidence, reason, calculated=str(value)))
                except FormulaReview as exc:
                    if is_control and cell["numeric"] is not None and cell["numeric"] != 0:
                        checks.append(_check(title, "difference", cell["numeric"], evidence,
                                             "Saved nonzero control requires investigation. Independent recalculation unavailable: " + str(exc), recalculated=False))
                    else:
                        checks.append(_check(title, "needs_review", evidence=evidence,
                                             reason="Saved value only; not independently recalculated: " + str(exc), recalculated=False))
                if limit is not None and len(checks) >= limit:
                    return checks
        return checks


class _FormulaParser:
    def __init__(self, book, sheet, text, stack):
        self.book, self.sheet, self.stack = book, sheet, stack
        try:
            self.tokens = [t for t in Tokenizer("="+text.lstrip("=")).items if t.type != "WHITE-SPACE"]
        except Exception as exc:
            raise FormulaReview(f"Cannot tokenize formula: {exc}") from exc
        self.position = 0

    def peek(self):
        return self.tokens[self.position] if self.position < len(self.tokens) else None

    def take(self):
        token = self.peek()
        if token is None:
            raise FormulaReview("Unexpected end of formula")
        self.position += 1
        return token

    def scalar(self, value):
        if isinstance(value, _Range):
            if len(value.addresses) != 1:
                raise FormulaReview("A multi-cell range is not a scalar amount")
            return self.book.evaluate(value.sheet, value.addresses[0], self.stack)
        return value

    def parse(self):
        if not self.tokens:
            raise FormulaReview("Empty formula")
        value = self.scalar(self.expression())
        if self.peek() is not None:
            raise FormulaReview("Unsupported formula syntax: " + self.peek().value)
        return _decimal(value)

    def expression(self, minimum=0):
        left = self.atom()
        while self.peek() is not None:
            token = self.peek()
            precedence = {"+": 1, "-": 1, "*": 2, "/": 2}.get(token.value)
            if token.type != "OPERATOR-INFIX" or precedence is None or precedence < minimum:
                break
            operator = self.take().value
            right = self.expression(precedence+1)
            a, b = self.scalar(left), self.scalar(right)
            if operator == "+": left = a+b
            elif operator == "-": left = a-b
            elif operator == "*": left = a*b
            else: left = a/b
        return left

    def atom(self):
        token = self.take()
        if token.type == "OPERATOR-PREFIX" and token.value in {"+", "-"}:
            amount = self.scalar(self.atom())
            return amount if token.value == "+" else -amount
        if token.type == "OPERAND":
            if token.subtype == "NUMBER": return _decimal(token.value)
            if token.subtype == "RANGE": return self.book._reference(token.value, self.sheet)
            raise FormulaReview("Unsupported non-numeric formula operand: " + token.value)
        if token.type == "PAREN" and token.subtype == "OPEN":
            value = self.expression()
            closing = self.take()
            if closing.type != "PAREN" or closing.subtype != "CLOSE":
                raise FormulaReview("Unbalanced parentheses")
            return value
        if token.type == "FUNC" and token.subtype == "OPEN":
            name = token.value[:-1].upper()
            if name not in {"SUM", "ROUND"}:
                raise FormulaReview(f"Unsupported function {name}; saved value is not recalculated")
            arguments = []
            if not (self.peek() and self.peek().type == "FUNC" and self.peek().subtype == "CLOSE"):
                while True:
                    arguments.append(self.expression())
                    if self.peek() and self.peek().type == "SEP" and self.peek().subtype == "ARG":
                        self.take()
                    else: break
            closing = self.take()
            if closing.type != "FUNC" or closing.subtype != "CLOSE":
                raise FormulaReview("Unbalanced function arguments")
            if name == "ROUND":
                if len(arguments) != 2:
                    raise FormulaReview("ROUND requires two arguments")
                places = self.scalar(arguments[1])
                if places != places.to_integral_value():
                    raise FormulaReview("Non-integral ROUND precision is unsupported")
                return round_money(self.scalar(arguments[0]), int(places))
            total, numeric_count = Decimal(0), 0
            for argument in arguments:
                if isinstance(argument, _Range):
                    for target in argument.addresses:
                        cell = self.book.cell(argument.sheet, target)
                        if not cell.get("formula") and cell["status"] in {"missing", "text"}:
                            continue  # Excel SUM ignores blank/text range cells.
                        total += self.book.evaluate(argument.sheet, target, self.stack)
                        numeric_count += 1
                else:
                    total += self.scalar(argument)
                    numeric_count += 1
            if numeric_count == 0:
                raise FormulaReview("SUM has no numeric source; blanks have not been treated as evidence of zero")
            return total
        raise FormulaReview("Unsupported formula token: " + token.value)


def inspect_workbook(book: SourceBook) -> dict:
    """Collect diagnostics without recalculating or altering the customer file."""
    errors, missing, external, unsupported, years, currencies = [], [], [], [], [], []
    year_pattern = re.compile(r"(?<!\d)(?:19|20)\d{2}(?!\d)")
    currency_pattern = re.compile(r"HK\$|HKD|RMB|CNY|US\$|USD|人民[币幣]|港[币幣]", re.I)
    for sheet in book.sheets:
        for row in book.rows(sheet, 1, 1048576):
            for address, cell in row["cells"].items():
                evidence = _evidence(cell, address)
                if cell["status"] == "error": errors.append(evidence)
                if cell.get("formula"):
                    if cell["numeric"] is None and cell["status"] == "missing": missing.append(evidence)
                    try:
                        tokens = Tokenizer("="+book._formula(cell)).items
                        functions = [t.value[:-1].upper() for t in tokens if t.type == "FUNC" and t.subtype == "OPEN"]
                        if any(name not in {"SUM", "ROUND"} for name in functions):
                            unsupported.append(dict(evidence, functions=functions, reason="Saved value only; unsupported calculation"))
                    except Exception as exc:
                        unsupported.append(dict(evidence, reason=str(exc)))
                    for ref in book.references(sheet, address):
                        if ref.get("external"):
                            external.append(dict(evidence, reference=ref["reference"], reason=ref["reason"]))
                elif row["row"] <= 20 and cell["value"] is not None:
                    text = str(cell["value"])
                    found = year_pattern.findall(text)
                    if found: years.append(dict(evidence, years=found))
                    found = currency_pattern.findall(text)
                    if found: currencies.append(dict(evidence, currencies=found, unit_hint="thousands" if re.search(r"000|千", text) else "units"))
    return {"path": book.path, "sheets": book.sheets,
            "formula_values": "legacy_formula_text_unavailable" if book.legacy_cached_values else "saved_cache_not_recalculated_except_supported_independent_checks",
            "review_warnings": ["Legacy XLS saved values only: formula text unavailable; independent formula verification is not possible."] if book.legacy_cached_values else [],
            "formula_errors": errors, "missing_formula_cache": missing, "external_dependencies": external,
            "external_link_files": book.data.get("external_links", []), "unsupported_formulas": unsupported,
            "year_labels": years, "currency_labels": currencies,
            "label_review": "Labels are hints, not automatic year/currency/classification decisions. Confirm current and comparative columns separately."}
