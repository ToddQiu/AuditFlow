"""Financial correctness and refusal-to-guess tests; no customer-file writes."""
from decimal import Decimal
from pathlib import Path
import sys
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from engine import (SourceBook, FormulaReview, round_money, format_money,
                    sum_check, rounding_check, cashflow_check, inspect_workbook)


def cell(address, value=None, formula=None, kind="n"):
    if isinstance(formula, str):
        formula = {"text": formula, "type": None, "shared": None, "range": None}
    return {"address": address, "value": value, "formula": formula, "type": kind}


def book(*sheets):
    return SourceBook.from_data({"kind": "workbook", "sheets": [
        {"name": name, "cells": cells} for name, cells in sheets], "external_links": []})


class MoneyTests(unittest.TestCase):
    def test_excel_half_away_from_zero_not_bankers_rounding(self):
        self.assertEqual(round_money("0.5"), Decimal("1"))
        self.assertEqual(round_money("-0.5"), Decimal("-1"))
        self.assertEqual(round_money("1.005", 2), Decimal("1.01"))
        self.assertEqual(round_money("-1.005", 2), Decimal("-1.01"))
        self.assertEqual(round_money("-0.004", 2), Decimal("0.00"))

    def test_blank_is_not_zero_and_nonfinite_is_rejected(self):
        for value in (None, "", True, "NaN", "Infinity"):
            with self.subTest(value=value), self.assertRaises(ValueError):
                round_money(value)
        self.assertEqual(round_money(0), Decimal(0))
        self.assertEqual(sum_check(["0", None], "0")["status"], "needs_review")
        self.assertEqual(sum_check(["0", "0"], "0")["status"], "pass")
        self.assertEqual(sum_check([], "0")["status"], "needs_review")

    def test_report_format_does_not_invent_missing_zero(self):
        self.assertEqual(format_money(None), "")
        self.assertEqual(format_money("-1234.5"), "(1,235)")
        self.assertEqual(format_money("8273470.49"), "8,273,470")
        self.assertEqual(format_money("-0.004", 2), "0.00")

    def test_tail_difference_is_reported_without_allocation(self):
        check = rounding_check(["0.49", "0.49"])
        self.assertEqual(check["status"], "difference")
        self.assertEqual(check["sum_of_rounded"], "0")
        self.assertEqual(check["rounded_total"], "1")
        self.assertEqual(check["difference"], "-1")
        self.assertEqual(check["exact_total"], "0.98")

    def test_selected_total_and_missing_reported_amount(self):
        self.assertEqual(sum_check(["11944", "1745", "362"], "14052")["difference"], "-1")
        self.assertEqual(sum_check(["10"], None)["status"], "needs_review")

    def test_cf_can_balance_internally_but_fail_the_bs_tie(self):
        good = cashflow_check("100", "40", "-20", "10", "-1", "129", "129")
        self.assertEqual(good["status"], "pass")
        # A lost operating line, even with the same closing statement figure.
        omitted = cashflow_check("100", "35", "-20", "10", "-1", "129", "129")
        self.assertEqual(omitted["flow_difference"], "-5")
        self.assertEqual(omitted["status"], "difference")
        wrong_bs = cashflow_check("100", "40", "-20", "10", "-1", "129", "128")
        self.assertEqual(wrong_bs["flow_difference"], "0")
        self.assertEqual(wrong_bs["bs_difference"], "1")
        self.assertEqual(wrong_bs["status"], "difference")
        self.assertEqual(cashflow_check("100", "40", None, "10", "-1", "129", "129")["status"], "needs_review")


class SourceTests(unittest.TestCase):
    def test_strict_source_identity_and_status(self):
        source = book(("PL", [cell("E8", "0"), cell("E9"), cell("E10", "#REF!", kind="e"), cell("A8", "123", kind="s")]))
        self.assertEqual(source.cell("PL", "$E$8")["numeric"], Decimal(0))
        self.assertEqual(source.cell("PL", "E9")["status"], "missing")
        self.assertEqual(source.cell("PL", "E10")["status"], "error")
        self.assertIsNone(source.cell("PL", "A8")["numeric"])
        self.assertEqual(source.cell("PL", "E11")["status"], "missing")
        with self.assertRaises(KeyError): source.cell("P/L", "E8")
        for address in ("E8:E9", "E0", "XFE1", "A1048577", "=E8"):
            with self.subTest(address=address), self.assertRaises(ValueError):
                source.cell("PL", address)

    def test_shared_formulas_translate_ranges_absolute_and_cross_sheet_refs(self):
        source = book(("PL", [
            cell("B2", "11", {"text": "SUM(A2:A3)+$C$1+'Other'!D2", "type": "shared", "shared": "7", "range": "B2:B3"}),
            cell("B3", "14", {"text": "", "type": "shared", "shared": "7", "range": None}),
            cell("A2", "1"), cell("A3", "2"), cell("A4", "3"), cell("C1", "5")]),
            ("Other", [cell("D2", "3"), cell("D3", "4")]))
        formula = source.cell("PL", "B3")["formula"]
        self.assertEqual(formula["text"], "SUM(A3:A4)+$C$1+'Other'!D3")
        self.assertTrue(formula["shared_expanded"])
        self.assertEqual(source.evaluate("PL", "B3"), Decimal(14))
        self.assertEqual({(r["sheet"], r["address"]) for r in source.references("PL", "B3")},
                         {("PL", "A3"), ("PL", "A4"), ("PL", "C1"), ("Other", "D3")})

    def test_missing_shared_anchor_cannot_pass_from_cache(self):
        source = book(("PL", [cell("B3", "0", {"text": "", "type": "shared", "shared": "7"})]))
        self.assertEqual(source.checks("PL")[0]["status"], "needs_review")

    def test_independent_formula_result_exposes_stale_excel_cache(self):
        source = book(("PL", [cell("A1", "1.005"), cell("A2", "2.005"),
                              cell("A3", "3.00", "ROUND(SUM(A1:A2),2)"),
                              cell("A4", "0", "A3-3.01")]))
        self.assertEqual(source.evaluate("PL", "A3"), Decimal("3.01"))
        first = source.checks("PL")[0]
        self.assertEqual(first["status"], "difference")
        self.assertEqual(first["difference"], "0.01")
        self.assertEqual(source.evaluate("PL", "A4"), Decimal(0))

    def test_sum_blanks_are_distinct_from_missing_explicit_amounts(self):
        source = book(("PL", [cell("A1", "2"), cell("A2"), cell("A3", "heading", kind="s"),
                              cell("B1", "2", "SUM(A1:A3)"), cell("B2", "0", "A2"),
                              cell("B3", "0", "SUM(C1:C3)")]))
        self.assertEqual(source.evaluate("PL", "B1"), Decimal(2))
        with self.assertRaises(FormulaReview): source.evaluate("PL", "B2")
        with self.assertRaises(FormulaReview): source.evaluate("PL", "B3")

    def test_unsupported_external_and_circular_formulas_never_get_cache_pass(self):
        source = book(("PL", [cell("A1", "0", "SUMIF(C1:C2,D1,E1:E2)"),
                              cell("A2", "0", "'[old.xlsx]TB'!B1"),
                              cell("B1", "0", "B2"), cell("B2", "0", "B1")]))
        checks = source.checks("PL")
        self.assertTrue(all(c["status"] == "needs_review" for c in checks))
        self.assertIn("Circular", str(source.trace("PL", "B1")))
        diagnostics = inspect_workbook(source)
        self.assertEqual(len(diagnostics["external_dependencies"]), 1)
        self.assertEqual(len(diagnostics["unsupported_formulas"]), 1)
        self.assertIn("old.xlsx", diagnostics["external_dependencies"][0]["reference"])

    def test_formula_looking_text_is_not_an_external_dependency(self):
        source = book(("PL", [cell("A1", "Example: '[fake.xlsx]TB'!A1", kind="s"),
                              cell("B1", "1")]))
        self.assertEqual(inspect_workbook(source)["external_dependencies"], [])

    def test_saved_nonzero_footer_control_is_not_silently_accepted(self):
        source = book(("BS", [cell("A50", "Difference", kind="s"),
                              cell("E50", "120000", "SUMIF(A:A,B1,C:C)-D1")]))
        check = source.checks("BS")[0]
        self.assertEqual(check["status"], "difference")
        self.assertEqual(check["difference"], "120000")
        self.assertFalse(check["recalculated"])

    def test_year_currency_hints_do_not_make_classification_decisions(self):
        source = book(("PL", [cell("A2", "Year ended 31 March 2026", kind="s"),
                              cell("E5", "2026"), cell("G5", "2025"),
                              cell("E6", "HK$", kind="s"), cell("G6", "RMB’000", kind="s")]))
        diagnostics = inspect_workbook(source)
        self.assertEqual(len(diagnostics["year_labels"]), 3)
        self.assertEqual(diagnostics["currency_labels"][1]["unit_hint"], "thousands")
        rows = source.rows("PL", 2, 6)
        self.assertEqual(rows[0]["row"], 2)
        self.assertIn("2026", rows[0]["label"])

    def test_y101_detects_current_column_and_crosschecks_separate_bs(self):
        source = book(("Y101 Cash Flow", [
            cell("F1", "Year ended 31 March 2026", kind="s"),
            cell("F5", "2026"), cell("G5", "2025"), cell("F6", "HK$", kind="s"),
            cell("A10", "Net cash used in operating activities", kind="s"), cell("F10", "40"),
            cell("A20", "Net cash used in investing activities", kind="s"), cell("F20", "-20"),
            cell("A30", "Net cash generated from financing activities", kind="s"), cell("F30", "10"),
            cell("A40", "Cash and cash equivalents at beginning of year", kind="s"), cell("F40", "100"),
            cell("A41", "Effect of foreign exchange rate changes", kind="s"), cell("F41", "-1"),
            cell("A42", "Cash and cash equivalents at end of year", kind="s"), cell("F42", "129"), cell("G42", "100")]),
            ("BS", [cell("D1", "As at 31 March 2026", kind="s"),
                    cell("D5", "2026"), cell("E5", "2025"), cell("D6", "HK$", kind="s"),
                    cell("A22", "Cash and bank balances", kind="s"), cell("D22", "128"), cell("E22", "100")]))
        check = source.checks("Y101 Cash Flow")[0]
        self.assertEqual(check["status"], "difference")
        self.assertEqual(check["flow_difference"], "0")
        self.assertEqual(check["bs_difference"], "1")
        self.assertEqual(check["evidence"][-1]["address"], "D22")

    def test_arithmetic_negative_refs_and_division(self):
        source = book(("PL", [cell("A1", "-1.005"), cell("B1", "2"),
                              cell("C1", "0", "ROUND((-A1*B1)/2,2)")]))
        self.assertEqual(source.evaluate("PL", "C1"), Decimal("1.01"))

    def test_cashflow_period_mismatch_or_unknown_cannot_pass_from_equal_cash(self):
        cf = [cell("F2", "2026"), cell("F3", "HK$", kind="s")]
        for row, label, amount in [(10, "Net cash from operating activities", "10"),
                                  (11, "Net cash from investing activities", "20"),
                                  (12, "Net cash from financing activities", "30"),
                                  (13, "Effect of foreign exchange rate changes", "0"),
                                  (14, "Cash at beginning of year", "40"),
                                  (15, "Cash at end of year", "100")]:
            cf.extend([cell(f"A{row}", label, kind="s"), cell(f"F{row}", amount)])
        for year in ("2025", None, "2026"):
            with self.subTest(year=year):
                bs = [cell("F3", "HK$", kind="s"), cell("A10", "Cash and bank balances", kind="s"), cell("F10", "100")]
                if year is not None: bs.append(cell("F2", year))
                source = book(("Y101 Cash Flow", cf), ("BS", bs))
                check = source.checks("Y101 Cash Flow")[0]
                self.assertEqual(check["status"], "needs_review")
                if year == "2025":
                    self.assertIn("2026", check["reason"])
                    self.assertIn("2025", check["reason"])

    def test_conflicting_period_labels_in_the_same_column_need_review(self):
        source = book(("BS", [cell("F2", "Year ended 2025", kind="s"), cell("F5", "2026")]))
        with self.assertRaisesRegex(FormulaReview, "Conflicting"):
            source._column_year("BS", 6)

    def test_cashflow_full_date_must_match_not_only_year_and_cash(self):
        def make_source(cf_date, bs_date, headers_in_amount_column=False):
            cf = [cell("F2", "2026"), cell("F3", "HK$", kind="s")]
            for row, label, amount in [(10, "Net cash from operating activities", "10"),
                                      (11, "Net cash from investing activities", "20"),
                                      (12, "Net cash from financing activities", "30"),
                                      (13, "Effect of foreign exchange rate changes", "0"),
                                      (14, "Cash at beginning of year", "40"),
                                      (15, "Cash at end of year", "100")]:
                cf.extend([cell(f"A{row}", label, kind="s"), cell(f"F{row}", amount)])
            bs = [cell("F2", "2026"), cell("F3", "HK$", kind="s"),
                  cell("A10", "Cash and bank balances", kind="s"), cell("F10", "100")]
            header = "F1" if headers_in_amount_column else "A1"
            if cf_date is not None: cf.append(cell(header, cf_date, kind="s"))
            if bs_date is not None: bs.append(cell(header, bs_date, kind="s"))
            return book(("Y101 Cash Flow", cf), ("BS", bs))

        for cf_date, bs_date in (("Yearended31March2026", "Asat2026-03-31"),
                                 ("Period ended 31 Mar. 2026", "As at 31.03.2026"),
                                 ("截至2026年3月31日", "As at 31 March 2026")):
            with self.subTest(cf_date=cf_date, bs_date=bs_date):
                check = make_source(cf_date, bs_date).checks("Y101 Cash Flow")[0]
                self.assertEqual(check["status"], "pass", check)
                self.assertEqual(check["period_end_date"], "2026-03-31")
        for cf_date, bs_date in (("Year ended 31 March 2026", "As at 30 June 2026"),
                                 (None, None),
                                 ("Year ended 31 March 2026", None),
                                 ("Year ended 31 March 2026", "As at 31 April 2026"),
                                 ("Year ended 31 March 2026 and 30 June 2026", "As at 31 March 2026"),
                                 ("Year ended 31 March 2025", "As at 31 March 2025")):
            with self.subTest(cf_date=cf_date, bs_date=bs_date):
                for amount_column in (False, True):
                    check = make_source(cf_date, bs_date, amount_column).checks("Y101 Cash Flow")[0]
                    self.assertEqual(check["status"], "needs_review", check)
                    self.assertNotIn("period_end_date", check)

    def test_metadata_only_workbook_cannot_masquerade_as_empty_verified_source(self):
        with self.assertRaisesRegex(ValueError, "未读取金额"):
            SourceBook.from_data({"kind": "workbook", "metadata_only": True,
                                  "status": "needs_streaming_review", "sheets": [{"name": "BS", "cells": []}]})

    def test_legacy_xls_saved_numeric_values_are_not_independently_verified(self):
        source = SourceBook.from_data({"kind": "workbook", "formula_values": "legacy_formula_text_unavailable",
                                      "sheets": [{"name": "PL", "cells": [cell("A1", "10"), cell("A2", "20"), cell("A3", "30")]}]})
        self.assertEqual(source.cell("PL", "A3")["numeric"], Decimal(30))
        with self.assertRaisesRegex(FormulaReview, "Legacy XLS"):
            source.evaluate("PL", "A3")
        checks = source.checks("PL")
        self.assertEqual(checks[0]["status"], "needs_review")
        self.assertFalse(checks[0]["recalculated"])
        self.assertEqual(source.selected_sum("PL", ["A1", "A2"], "A3")["status"], "needs_review")
        diagnostics = inspect_workbook(source)
        self.assertEqual(diagnostics["formula_values"], "legacy_formula_text_unavailable")
        self.assertTrue(diagnostics["review_warnings"])


if __name__ == "__main__":
    unittest.main()
