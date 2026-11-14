"""Source-binding safeguards using mocks; no customer outputs or DB writes."""
from decimal import Decimal
from pathlib import Path
import sys
import unittest
from unittest.mock import Mock, patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import server
from engine import FormulaReview


class SourceAmountTests(unittest.TestCase):
    def payload(self, **overrides):
        return {"source_path": "fixture.xlsx", "refs": [{"sheet": "A3.1", "address": "F68"}],
                "unit_divisor": 1, "places": 0, **overrides}

    def fixture(self, value="100", *, calculated="100", formula=True, error=None):
        source = Mock()
        source.data={}
        source.cell.return_value = {"sheet": "A3.1", "address": "F68", "value": value,
                                    "numeric": None if value is None else Decimal(value),
                                    "status": "missing" if value is None else "number",
                                    "formula": {"text": "SUM(A1:A2)"} if formula else None}
        if error:
            source.evaluate.side_effect = FormulaReview(error)
        else:
            source.evaluate.return_value = Decimal(calculated)
        return source, {"book": source, "path": "fixture.xlsx", "sha256": "fixture-hash"}

    def test_stale_cache_rejected_even_when_saved_values_confirmation_is_true(self):
        source, entry = self.fixture("100", calculated="30")
        with patch.object(server, "book", return_value=entry):
            for confirm in (False, True):
                with self.subTest(confirm=confirm), self.assertRaisesRegex(ValueError, "保存值100|保存值 100"):
                    server.source_amount(self.payload(confirm_saved_values=confirm))
        source.evaluate.assert_called_with("A3.1", "F68")

    def test_missing_source_rejected_before_recalculation_or_zero_substitution(self):
        source, entry = self.fixture(None)
        with patch.object(server, "book", return_value=entry), self.assertRaisesRegex(ValueError, "空值"):
            server.source_amount(self.payload(confirm_saved_values=True))
        source.evaluate.assert_not_called()

    def test_unsupported_or_external_formula_requires_explicit_boolean_confirmation(self):
        source, entry = self.fixture("100", error="External workbook dependency: [old.xlsx]TB!A1")
        with patch.object(server, "book", return_value=entry):
            for confirm in (None, False, "true", 1):
                with self.subTest(confirm=confirm), self.assertRaisesRegex(ValueError, "明确确认|明确.*保存值"):
                    server.source_amount(self.payload(confirm_saved_values=confirm))
            result = server.source_amount(self.payload(confirm_saved_values=True))
        self.assertTrue(result["needs_review"])
        self.assertEqual(result["value"], "100")
        self.assertEqual(result["source"]["verification_status"], "needs_review")
        self.assertEqual(result["source"]["value_basis"], "explicitly_confirmed_saved_values")
        evidence = result["source"]["refs"][0]
        self.assertEqual(evidence["verification"]["status"], "needs_review")
        self.assertIn("old.xlsx", evidence["verification"]["reason"])
        self.assertEqual(result["source"]["sha256"], "fixture-hash")

    def test_independent_formula_verified_and_evidence_retained(self):
        source, entry = self.fixture("30", calculated="30")
        with patch.object(server, "book", return_value=entry):
            result = server.source_amount(self.payload())
        self.assertFalse(result["needs_review"])
        evidence = result["source"]["refs"][0]
        self.assertEqual(evidence["verification"]["calculated_value"], "30")
        self.assertEqual(evidence["verification"]["basis"], "independently_evaluated_formula")
        self.assertEqual(evidence["formula"]["text"], "SUM(A1:A2)")

    def test_decimal_unit_conversion_and_precision_are_not_float_rounding(self):
        source, entry = self.fixture("1005", formula=False)
        with patch.object(server, "book", return_value=entry):
            result = server.source_amount(self.payload(unit_divisor=Decimal("1000"), places=Decimal("2")))
            negative = server.source_amount(self.payload(unit_divisor=1000, places=2,
                refs=[{"sheet": "A3.1", "address": "$F$68", "coefficient": -1}]))
            with self.assertRaisesRegex(ValueError, "两位小数"):
                server.source_amount(self.payload(places=2.9))
        source.evaluate.assert_not_called()
        self.assertEqual(result["exact_amount"], "1.005")
        self.assertEqual(result["value"], "1.01")
        self.assertEqual(negative["value"], "(1.01)")
        self.assertEqual(result["source"]["value_basis"], "numeric_source_values")
        self.assertEqual(negative["source"]["refs"][0]["address"], "F68")

    def test_zero_is_valid_but_duplicate_refs_cannot_double_count(self):
        source, entry = self.fixture("0", formula=False)
        with patch.object(server, "book", return_value=entry):
            self.assertEqual(server.source_amount(self.payload())["value"], "0")
            with self.assertRaisesRegex(ValueError, "重复"):
                server.source_amount(self.payload(refs=[
                    {"sheet": "A3.1", "address": "F68"}, {"sheet": "A3.1", "address": "$F$68"}]))

    def test_unconfirmed_unsupported_preview_can_display_amount_but_cannot_bind(self):
        source, entry = self.fixture("100", error="Unsupported function SUMIF")
        with patch.object(server, "book", return_value=entry):
            result = server.source_amount(self.payload(), preview=True)
            with self.assertRaisesRegex(ValueError, "明确"):
                server.source_amount(self.payload())
            # A client flag cannot select the server's privileged preview path.
            with self.assertRaisesRegex(ValueError, "明确"):
                server.source_amount(self.payload(preview=True))
        self.assertEqual(result["value"], "100")
        self.assertTrue(result["needs_review"])
        self.assertFalse(result["source"]["confirm_saved_values"])
        self.assertEqual(result["source"]["value_basis"], "unconfirmed_saved_value_preview")
        self.assertIn("绑定前", result["warning"])

    def test_preview_cannot_bypass_supported_formula_cache_mismatch(self):
        source, entry = self.fixture("100", calculated="30")
        with patch.object(server, "book", return_value=entry), self.assertRaisesRegex(ValueError, "不一致"):
            server.source_amount(self.payload(), preview=True)

    def test_legacy_xls_amount_is_not_certified_as_an_original_constant(self):
        source,entry=self.fixture('100',formula=False)
        source.data={'formula_values':'legacy_formula_text_unavailable'}
        with patch.object(server,'book',return_value=entry):
            with self.assertRaisesRegex(ValueError,'仅有保存值'): server.source_amount(self.payload())
            preview=server.source_amount(self.payload(),preview=True)
            confirmed=server.source_amount(self.payload(confirm_saved_values=True))
        self.assertTrue(preview['needs_review']);self.assertTrue(confirmed['needs_review'])
        self.assertEqual(preview['source']['refs'][0]['verification']['basis'],'legacy_saved_value_only')
        source.evaluate.assert_not_called()


class ProjectMetadataTests(unittest.TestCase):
    def project(self):
        return {"source_path": "fixture.xlsx", "template_path": "template.docx", "current_year": 2026,
                "comparative_year": 2025, "currency": "HKD", "unit": 1, "current_column": "A3.1!F",
                "period_start":"2025-04-01","period_end":"2026-03-31",
                "source_confirmed": True, "confirmed_sha256": "old-hash"}

    def test_metadata_snapshot_detects_changed_year_currency_unit_or_column(self):
        project = self.project()
        original = server.project_context(project)
        for field, value in (("current_year", 2027), ("currency", "RMB"), ("unit", 1000), ("current_column", "A3.1!G"),("period_end","2026-06-30")):
            with self.subTest(field=field):
                modified = {**project, field: value}
                self.assertNotEqual(server.project_context(modified), original)
        self.assertNotIn("source_confirmed", original)

    def test_changed_source_hash_invalidates_earlier_confirmation(self):
        with patch.object(server, "source_path", return_value=Path("fixture.xlsx")), \
             patch.object(server, "sha256", return_value="changed-hash"), \
             self.assertRaisesRegex(ValueError, "已改变"):
            server.confirmed_source(self.project())

    def test_confirmation_is_not_inferred_from_a_valid_year(self):
        project = {**self.project(), "source_confirmed": False}
        with self.assertRaisesRegex(ValueError, "先确认"):
            server.confirmed_source(project)
        self.assertEqual(server.validate_project(self.project()), (2026, 2025))

    def test_source_confirmation_requires_exact_period_range(self):
        with self.assertRaisesRegex(ValueError,'起止日期'): server.validate_project({**self.project(),'period_end':''})
        with self.assertRaisesRegex(ValueError,'晚于'): server.validate_project({**self.project(),'period_start':'2026-04-01'})


if __name__ == "__main__":
    unittest.main()
