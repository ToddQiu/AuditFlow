"""Independent front-half accounting controls; all source cells are synthetic.

No source workbook, customer file, or application database is written here.
"""
from copy import deepcopy
from decimal import Decimal
from pathlib import Path
import sys
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from engine import SourceBook, FormulaReview
import intake


def cell(address, value=None, kind="n", formula=None):
    return {"address": address, "value": value, "type": kind,
            "formula": {"text": formula} if formula else None}


class FrontHalfTests(unittest.TestCase):
    def setUp(self):
        self.path = "synthetic-management.xlsx"
        self.hashes = {self.path: "sha-source-one", "synthetic-detail.xlsx": "sha-detail-one"}
        self.entries = {}
        self.install(self.path, [
            cell("A1", "Code", "s"), cell("B1", "Account", "s"), cell("C1", "Closing", "s"),
            cell("A2", "1002", "s"), cell("B2", "Bank", "s"), cell("C2", "100.005"),
            cell("A3", "3001", "s"), cell("B3", "Capital", "s"), cell("C3", "-100.005"),
        ])

    def install(self, path, cells):
        source = SourceBook.from_data({"kind": "workbook", "sheets": [{"name": "Data", "cells": cells}]})
        self.entries[path] = {"path": path, "sha256": self.hashes[path], "book": source}

    def get_book(self, path):
        return self.entries[path]

    def read_amount(self, payload, *, preview=False, entry=None):
        entry = entry or self.get_book(payload["source_path"])
        amount, refs, needs_review = Decimal(0), [], False
        for ref in payload["refs"]:
            c = entry["book"].cell(ref["sheet"], ref["address"])
            if c["numeric"] is None:
                raise ValueError("Missing/text/error is not a numeric amount")
            verification = {"status": "pass"}
            if c.get("formula"):
                try:
                    calculated = entry["book"].evaluate(ref["sheet"], ref["address"])
                except FormulaReview:
                    needs_review = True
                    verification = {"status": "needs_review"}
                    if not preview and payload.get("confirm_saved_values") is not True:
                        raise ValueError("Please explicitly confirm saved values")
                else:
                    if calculated != c["numeric"]:
                        raise ValueError("Formula cache does not match independent calculation")
            amount += c["numeric"] * Decimal(str(ref.get("coefficient", 1)))
            refs.append({**ref, "value": str(c["numeric"]), "verification": verification})
        amount /= Decimal(str(payload.get("unit_divisor", 1)))
        return {"exact_amount": str(amount), "value": str(amount), "needs_review": needs_review,
                "source": {"path": entry["path"], "sha256": entry["sha256"], "refs": refs,
                           "verification_status": "needs_review" if needs_review else "pass",
                           "value_basis": "saved_value_only" if needs_review else "numeric_source_values"}}

    def spec(self, **changes):
        result = {"path": self.path, "sheet": "Data", "first_row": 2, "last_row": 3,
                  "columns": {"code": "A", "name": "B", "amount": "C"},
                  "amount_mode": "signed", "sign_convention": "debit_positive",
                  "credit_sign": "positive", "source_role": "closing_balances",
                  "currency_basis": "functional", "unit": 1,
                  "context": {"entity": "Test HK Limited", "period_start": "2025-04-01",
                              "period_end": "2026-03-31", "currency": "HKD"}}
        result.update(changes)
        return result

    def preview(self, spec=None):
        return intake.preview_import(spec or self.spec(), get_book=self.get_book, amount_reader=self.read_amount)

    def confirmed_spec(self, spec=None, rows=None):
        result = deepcopy(spec or self.spec())
        result.update(sha256=self.hashes[result["path"]], selected_rows=rows or [2, 3],
                      selection_confirmed=True, metadata_confirmed=True, confirm_saved_values=False)
        return result

    def load(self, draft=None, spec=None, rows=None):
        return intake.import_rows(draft or {}, self.confirmed_spec(spec, rows),
                                  get_book=self.get_book, amount_reader=self.read_amount,
                                  hash_reader=self.hashes.__getitem__)

    def summarize(self, draft):
        return intake.summarize(draft, hash_reader=self.hashes.__getitem__)

    def mapped(self, draft=None, mapping_codes=None):
        draft = draft or self.load()
        codes = mapping_codes or ["1002", "3001"]
        mappings = [{"row_id": line["row_id"], "account_code": code,
                     "account_name": "Bank" if code == "1002" else "Capital",
                     "lead_code": "E101" if code == "1002" else "M101",
                     "lead_title": "Bank balances" if code == "1002" else "Capital balances"}
                    for line, code in zip(draft["lines"], codes)]
        return intake.apply_mappings(draft, {"mappings": mappings})

    def support(self, draft, amounts, account="1002", *, path="synthetic-detail.xlsx", **spec_changes):
        self.hashes.setdefault(path, "sha-" + path)
        self.install(path, [item for row, amount in enumerate(amounts, 2) for item in
                           (cell(f"B{row}", f"Real detail {row}", "s"), cell(f"C{row}", amount))])
        spec = self.spec(path=path, first_row=2, last_row=len(amounts)+1,
                         source_role="support_details", columns={"name": "B", "amount": "C"},
                         **spec_changes)
        spec = self.confirmed_spec(spec, list(range(2, len(amounts)+2)))
        return intake.add_support(draft, {"account_code": account, "label": "Customer real detail", "spec": spec},
                                  get_book=self.get_book, amount_reader=self.read_amount,
                                  hash_reader=self.hashes.__getitem__)

    def test_preview_retains_decimal_precision_and_cell_provenance(self):
        result = self.preview()
        self.assertIn("100.005", str(result))
        self.assertIn("C2", str(result))
        self.assertIn(self.hashes[self.path], str(result))

    def test_import_requires_literal_selection_and_metadata_confirmation(self):
        for flag in ("selection_confirmed", "metadata_confirmed"):
            for value in (False, "true", 1, None):
                spec = self.confirmed_spec()
                spec[flag] = value
                with self.subTest(flag=flag, value=value), self.assertRaises(ValueError):
                    intake.import_rows({}, spec, get_book=self.get_book, amount_reader=self.read_amount,
                                       hash_reader=self.hashes.__getitem__)

    def test_only_actual_closing_balances_can_be_imported_to_tb(self):
        for role in ("ledger_movements", "general_ledger", "opening_balances", "support_details"):
            with self.subTest(role=role), self.assertRaises(ValueError):
                self.load(spec=self.spec(source_role=role))

    def test_wrong_source_sha_is_rejected_even_with_confirmation(self):
        spec = self.confirmed_spec()
        spec["sha256"] = "a-different-preview-version"
        with self.assertRaises(ValueError):
            intake.import_rows({}, spec, get_book=self.get_book, amount_reader=self.read_amount,
                               hash_reader=self.hashes.__getitem__)

    def test_repeated_source_rows_do_not_double_count(self):
        original = self.load()
        with self.assertRaises(ValueError):
            self.load(draft=original)

    def test_parent_and_child_accounts_cannot_both_enter_additive_tb(self):
        self.install(self.path, [cell("A2", "1002", "s"), cell("B2", "Bank total", "s"), cell("C2", "100"),
                                 cell("A3", "100201", "s"), cell("B3", "Bank detail", "s"), cell("C3", "100")])
        with self.assertRaises(ValueError):
            self.load()

    def test_blank_account_or_formula_error_is_not_imported_as_zero(self):
        for cells in ([cell("C2", "0")],
                      [cell("A2", "1002", "s"), cell("B2", "Bank", "s"), cell("C2", "#REF!", "e")]):
            self.install(self.path, cells)
            with self.subTest(cells=cells), self.assertRaises(ValueError):
                self.load(spec=self.spec(first_row=2, last_row=2), rows=[2])

    def test_total_row_is_rejected_as_an_account(self):
        self.install(self.path, [cell("A2", "TOTAL", "s"), cell("B2", "Grand total", "s"), cell("C2", "100")])
        with self.assertRaises(ValueError):
            self.load(spec=self.spec(last_row=2), rows=[2])

    def test_month_year_and_balance_dates_cannot_silently_change_context(self):
        original = self.load()
        for field, value in (("entity", "Other Limited"), ("currency", "USD"),
                             ("period_start", "2026-01-01"), ("period_end", "2025-12-31")):
            spec = self.spec()
            spec["context"][field] = value
            with self.subTest(field=field), self.assertRaises(ValueError):
                self.load(draft=original, spec=spec)

    def test_units_are_explicit_not_silently_accepted(self):
        for unit in (0, -1, 100, "1000", True):
            with self.subTest(unit=unit), self.assertRaises(ValueError):
                self.load(spec=self.spec(unit=unit))

    def test_original_currency_amount_cannot_mix_a_different_currency(self):
        self.install(self.path, [cell("A2", "1002", "s"), cell("B2", "Bank", "s"),
                                 cell("C2", "100"), cell("D2", "USD", "s")])
        spec = self.spec(last_row=2, currency_basis="original",
                         columns={"code": "A", "name": "B", "amount": "C", "currency": "D"})
        with self.assertRaises(ValueError):
            self.load(spec=spec, rows=[2])

    def test_independently_wrong_formula_cache_is_refused(self):
        self.install(self.path, [cell("A2", "1002", "s"), cell("B2", "Bank", "s"),
                                 cell("D2", "2"), cell("E2", "3"), cell("C2", "100", formula="SUM(D2:E2)")])
        with self.assertRaises(ValueError):
            self.load(spec=self.spec(last_row=2), rows=[2])

    def test_source_changes_are_reported_without_mutating_draft(self):
        draft = self.load()
        before = deepcopy(draft)
        self.hashes[self.path] = "source-changed-after-import"
        result = self.summarize(draft)
        self.assertTrue(result["stale_sources"])
        self.assertEqual(draft, before)
        with self.assertRaises(ValueError):
            self.load(draft=draft)

    def test_blank_debit_credit_side_needs_explicit_policy(self):
        self.install(self.path, [cell("A2", "1002", "s"), cell("B2", "Bank", "s"), cell("C2", "100"),
                                 cell("A3", "3001", "s"), cell("B3", "Capital", "s"), cell("D3", "100")])
        spec = self.spec(amount_mode="debit_credit", columns={"code": "A", "name": "B", "debit": "C", "credit": "D"})
        with self.assertRaises(ValueError):
            self.load(spec=spec)
        draft = self.load(spec={**spec, "blank_sides_are_zero": True})
        self.assertIn("100", str(draft))

    def test_both_debit_credit_sides_blank_are_not_assumed_zero(self):
        self.install(self.path, [cell("A2", "1002", "s"), cell("B2", "Bank", "s")])
        spec = self.spec(last_row=2, amount_mode="debit_credit", blank_sides_are_zero=True,
                         columns={"code": "A", "name": "B", "debit": "C", "credit": "D"})
        with self.assertRaises(ValueError):
            self.load(spec=spec, rows=[2])

    def test_selection_outside_declared_rows_is_refused(self):
        for rows in ([1, 2], [2, 4], [2, 2], ["2", 3], []):
            spec = self.confirmed_spec()
            spec["selected_rows"] = rows
            with self.subTest(rows=rows), self.assertRaises(ValueError):
                intake.import_rows({}, spec, get_book=self.get_book, amount_reader=self.read_amount,
                                   hash_reader=self.hashes.__getitem__)

    def test_unmapped_balanced_source_does_not_mean_tb_is_finished(self):
        result = self.summarize(self.load())
        self.assertEqual(len(result["tb"]["unmapped"]), 2)
        self.assertEqual(result["tb"]["status"], "needs_review")
        self.assertFalse(result["breakdowns"])

    def test_balanced_tb_remains_an_unadjusted_draft_without_synthetic_journals(self):
        result = self.summarize(self.mapped())
        self.assertEqual(Decimal(result["tb"]["difference"]), Decimal(0))
        self.assertEqual(Decimal(result["tb"]["total_debit"]), Decimal("100.005"))
        self.assertEqual(Decimal(result["tb"]["total_credit"]), Decimal("100.005"))
        self.assertEqual(result["tb"]["status"], "draft_balanced")
        self.assertTrue(all(b["status"] == "needs_review" for b in result["breakdowns"]))
        self.assertTrue(all(b["detail_total"] is None for b in result["breakdowns"]))
        self.assertNotIn("journals", result["draft"])

    def test_difference_of_one_cent_is_preserved_not_forced_to_zero(self):
        self.install(self.path, [cell("A2", "1002", "s"), cell("B2", "Bank", "s"), cell("C2", "100.01"),
                                 cell("A3", "3001", "s"), cell("B3", "Capital", "s"), cell("C3", "-100")])
        result = self.summarize(self.mapped())
        self.assertEqual(Decimal(result["tb"]["difference"]), Decimal("0.01"))
        self.assertEqual(result["tb"]["status"], "difference")
        self.assertEqual(len(result["tb"]["accounts"]), 2)

    def test_breakdown_retains_real_detail_rows_and_exact_difference(self):
        draft = self.support(self.mapped(), ["40.005", "60"])
        result = self.summarize(draft)
        bank = next(b for b in result["breakdowns"] if b["lead_code"] == "E101")
        self.assertEqual(len(bank["details"]), 2)
        self.assertEqual(Decimal(bank["detail_total"]), Decimal("100.005"))
        self.assertEqual(Decimal(bank["difference"]), Decimal(0))
        self.assertEqual(bank["status"], "tied_needs_review")
        self.assertEqual(bank["details"][0]["source_path"], "synthetic-detail.xlsx")
        self.assertIn("C2", str(bank["details"][0]["refs"]))

    def test_missing_detail_is_not_created_by_subtracting_actual_detail_from_tb(self):
        draft = self.support(self.mapped(), ["80"])
        result = self.summarize(draft)
        bank = next(b for b in result["breakdowns"] if b["lead_code"] == "E101")
        self.assertEqual(Decimal(bank["difference"]), Decimal("-20.005"))
        self.assertEqual(bank["status"], "difference")
        self.assertEqual(len(bank["details"]), 1)
        self.assertEqual(Decimal(bank["details"][0]["amount"]), Decimal("80"))

    def test_tb_summary_cannot_be_reused_as_its_own_breakdown_support(self):
        draft = self.mapped()
        spec = self.confirmed_spec(self.spec(source_role="support_details", last_row=2), [2])
        with self.assertRaises(ValueError):
            intake.add_support(draft, {"account_code": "1002", "label": "Not real detail", "spec": spec},
                               get_book=self.get_book, amount_reader=self.read_amount,
                               hash_reader=self.hashes.__getitem__)

    def test_detail_source_role_and_context_must_match(self):
        draft = self.mapped()
        self.hashes["synthetic-detail.xlsx"] = "different-detail-content"
        self.install("synthetic-detail.xlsx", [cell("B2", "Actual bank", "s"), cell("C2", "100.005")])
        for role, currency in (("closing_balances", "HKD"), ("support_details", "USD")):
            spec = self.spec(path="synthetic-detail.xlsx", last_row=2,
                             source_role=role, columns={"name": "B", "amount": "C"})
            spec["context"]["currency"] = currency
            spec = self.confirmed_spec(spec, [2])
            with self.subTest(role=role, currency=currency), self.assertRaises(ValueError):
                intake.add_support(draft, {"account_code": "1002", "label": "Real detail", "spec": spec},
                                   get_book=self.get_book, amount_reader=self.read_amount,
                                   hash_reader=self.hashes.__getitem__)

    def test_unknown_or_inconsistent_tb_mapping_cannot_silently_replace_accounts(self):
        draft = self.load()
        for mappings in ([{"row_id": "unknown", "account_code": "1002", "account_name": "Bank"}],
                         [{"row_id": l["row_id"], "account_code": "1002", "account_name": n,
                           "lead_code": "E101", "lead_title": "Banks"}
                          for l, n in zip(draft["lines"], ("First name", "Different name"))]):
            with self.subTest(mappings=mappings), self.assertRaises(ValueError):
                intake.apply_mappings(draft, {"mappings": mappings})

    def test_credit_positive_and_credit_negative_conventions_do_not_use_abs(self):
        self.install(self.path, [cell("A2", "1002", "s"), cell("B2", "Bank", "s"), cell("C2", "100"), cell("D2", "0"),
                                 cell("A3", "3001", "s"), cell("B3", "Capital", "s"), cell("C3", "0"), cell("D3", "-100")])
        spec = self.spec(amount_mode="debit_credit", credit_sign="negative",
                         columns={"code": "A", "name": "B", "debit": "C", "credit": "D"})
        draft = self.load(spec=spec)
        amounts = [Decimal(l["amount"]) for l in draft["lines"]]
        self.assertEqual(amounts, [Decimal(100), Decimal(-100)])
        with self.assertRaises(ValueError):
            self.load(spec={**spec, "credit_sign": "positive"})

    def test_header_closing_period_and_amount_column_year_are_independently_checked(self):
        for value in ("As at 31 December 2025", "2025"):
            self.install(self.path, [cell("C1", value, "s"), cell("A2", "1002", "s"),
                                     cell("B2", "Bank", "s"), cell("C2", "100")])
            with self.subTest(value=value), self.assertRaises(ValueError):
                self.load(spec=self.spec(last_row=2), rows=[2])

    def test_declared_closing_amount_cannot_read_a_turnover_column(self):
        self.install(self.path, [cell("C1", "本年累计发生额", "s"), cell("A2", "1002", "s"),
                                 cell("B2", "Bank", "s"), cell("C2", "100")])
        with self.assertRaises(ValueError):
            self.load(spec=self.spec(last_row=2), rows=[2])

    def test_saved_formula_confirmation_never_changes_its_review_status(self):
        self.install(self.path, [cell("A2", "1002", "s"), cell("B2", "Bank", "s"),
                                 cell("C2", "100", formula="SUMIF(E:E,F1,G:G)")])
        spec = self.confirmed_spec(self.spec(last_row=2), [2])
        for confirmation in (False, "true", 1):
            spec["confirm_saved_values"] = confirmation
            with self.subTest(confirmation=confirmation), self.assertRaises(ValueError):
                intake.import_rows(None, spec, get_book=self.get_book, amount_reader=self.read_amount,
                                   hash_reader=self.hashes.__getitem__)
        spec["confirm_saved_values"] = True
        draft = intake.import_rows(None, spec, get_book=self.get_book, amount_reader=self.read_amount,
                                   hash_reader=self.hashes.__getitem__)
        self.assertTrue(draft["lines"][0]["needs_review"])

    def test_preview_source_copy_and_existing_batch_parent_do_not_evade_duplicates(self):
        original = self.load(spec=self.spec(last_row=2), rows=[2])
        copied = "synthetic-copy.xlsx"
        self.hashes[copied] = self.hashes[self.path]
        self.entries[copied] = {**self.entries[self.path], "path": copied}
        with self.assertRaises(ValueError):
            self.load(draft=original, spec=self.spec(path=copied, last_row=2), rows=[2])
        self.install(self.path, [cell("A2", "1002", "s"), cell("B2", "Bank total", "s"), cell("C2", "100"),
                                 cell("A3", "100201", "s"), cell("B3", "Bank detail", "s"), cell("C3", "100")])
        original = self.load(spec=self.spec(last_row=2), rows=[2])
        with self.assertRaises(ValueError):
            self.load(draft=original, spec=self.spec(first_row=3, last_row=3), rows=[3])

    def test_opposite_account_differences_do_not_pass_because_lead_total_is_zero(self):
        draft = self.load()
        mappings = [{"row_id": line["row_id"], "account_code": code, "account_name": name,
                     "lead_code": "E101", "lead_title": "Actual project grouping"}
                    for line, code, name in zip(draft["lines"], ("1002", "3001"), ("Bank", "Capital"))]
        draft = intake.apply_mappings(draft, {"mappings": mappings})
        draft = self.support(draft, ["110.005"], path="bank-details.xlsx")
        draft = self.support(draft, ["-110.005"], account="3001", path="capital-details.xlsx")
        group = self.summarize(draft)["breakdowns"][0]
        self.assertEqual(Decimal(group["difference"]), Decimal(0))
        self.assertEqual(group["status"], "difference")
        self.assertEqual([Decimal(a["difference"]) for a in group["account_lines"]],
                         [Decimal(10), Decimal(-10)])

    def test_source_changing_during_import_is_rejected_without_partial_commit(self):
        empty_draft = {}
        before = deepcopy(empty_draft)
        def changing_reader(payload, **kwargs):
            result = self.read_amount(payload, **kwargs)
            self.hashes[self.path] = "changed-during-read"
            return result
        with self.assertRaises(ValueError):
            intake.import_rows(empty_draft, self.confirmed_spec(), get_book=self.get_book,
                               amount_reader=changing_reader, hash_reader=self.hashes.__getitem__)
        self.assertEqual(empty_draft, before)

    def test_detail_changing_during_import_is_rejected_without_partial_commit(self):
        draft = self.mapped()
        before = deepcopy(draft)
        path = "synthetic-detail.xlsx"
        self.install(path, [cell("B2", "Actual bank detail", "s"), cell("C2", "100.005")])
        spec = self.confirmed_spec(self.spec(path=path, last_row=2, source_role="support_details",
                                             columns={"name": "B", "amount": "C"}), [2])
        def changing_reader(payload, **kwargs):
            result = self.read_amount(payload, **kwargs)
            if payload["source_path"] == path:
                self.hashes[path] = "changed-during-detail-read"
            return result
        with self.assertRaises(ValueError):
            intake.add_support(draft, {"account_code": "1002", "label": "Real detail", "spec": spec},
                               get_book=self.get_book, amount_reader=changing_reader,
                               hash_reader=self.hashes.__getitem__)
        self.assertEqual(draft, before)

    def test_saved_unverified_formula_keeps_whole_tb_needs_review_even_if_balanced(self):
        self.install(self.path, [cell("A2", "1002", "s"), cell("B2", "Bank", "s"),
                                 cell("C2", "100.005", formula="SUMIF(E:E,F1,G:G)"),
                                 cell("A3", "3001", "s"), cell("B3", "Capital", "s"), cell("C3", "-100.005")])
        spec = self.confirmed_spec()
        spec["confirm_saved_values"] = True
        draft = intake.import_rows(None, spec, get_book=self.get_book, amount_reader=self.read_amount,
                                   hash_reader=self.hashes.__getitem__)
        result = self.summarize(self.mapped(draft))
        self.assertEqual(Decimal(result["tb"]["difference"]), Decimal(0))
        self.assertEqual(result["tb"]["status"], "needs_review")
        self.assertTrue(result["tb"]["accounts"][0]["needs_review"])

    def test_empty_tb_has_no_readiness_even_though_its_numeric_totals_are_zero(self):
        result = self.summarize(None)
        self.assertEqual(result["tb"]["status"], "needs_review")
        self.assertTrue(result["tb"]["issues"])
        self.assertFalse(result["tb"]["accounts"])
        self.assertFalse(result["breakdowns"])

    def test_zero_balanced_accounts_still_require_real_breakdown_detail(self):
        self.install(self.path, [cell("A2", "1002", "s"), cell("B2", "Bank", "s"), cell("C2", "0"),
                                 cell("A3", "3001", "s"), cell("B3", "Capital", "s"), cell("C3", "0")])
        result = self.summarize(self.mapped())
        self.assertEqual(result["tb"]["status"], "draft_balanced")
        self.assertTrue(all(g["status"] == "needs_review" and g["detail_total"] is None
                            for g in result["breakdowns"]))


if __name__ == "__main__":
    unittest.main()
