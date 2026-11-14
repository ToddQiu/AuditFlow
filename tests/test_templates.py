"""Synthetic OOXML tests only; never edit client reports or workbooks."""
import hashlib
import sys
import tempfile
import unittest
from unittest.mock import patch
import zipfile
from pathlib import Path

from lxml import etree as ET

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import templates as m


def element(name, **attrs):
    return ET.Element(m._tag(name), {m._tag(k): str(v) for k, v in attrs.items()})


def run(text, style=None):
    r = element("r")
    if style:
        props = ET.SubElement(r, m._tag("rPr"))
        ET.SubElement(props, m._tag(style))
    for index, value in enumerate(text.split("\t")):
        if index:
            ET.SubElement(r, m._tag("tab"))
        if value:
            t = ET.SubElement(r, m._tag("t"))
            t.text = value
    return r


def paragraph(text="", style=None):
    p = element("p")
    props = ET.SubElement(p, m._tag("pPr"))
    ET.SubElement(props, m._tag("jc"), {m._tag("val"): "right"})
    if text or style:
        p.append(run(text, style))
    return p


def table(rows):
    tbl = element("tbl")
    for texts in rows:
        tr = ET.SubElement(tbl, m._tag("tr"))
        for text in texts:
            tc = ET.SubElement(tr, m._tag("tc"))
            pr = ET.SubElement(tc, m._tag("tcPr"))
            ET.SubElement(pr, m._tag("tcW"), {m._tag("w"): "1500", m._tag("type"): "dxa"})
            tc.append(paragraph(text))
    return tbl


def revision(text, author, ident, deleted=False):
    rev = element("del" if deleted else "ins", id=ident, author=author,
                  date="2025-08-31T03:04:05Z")
    r = run(text)
    if deleted:
        for t in r.findall("w:t", m.NS):
            t.tag = m._tag("delText")
    rev.append(r)
    return rev


class TemplatesTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.directory = Path(self.tmp.name)
        self.path = self.directory / "template.docx"
        self.output = self.directory / "result.docx"
        self.document = element("document")
        self.body = ET.SubElement(self.document, m._tag("body"))
        self.body.append(paragraph("\tNotes\t2026\t2025"))
        self.statement = paragraph()
        self.statement.extend([run("Loss before tax\t6\t"), run("1,", "b"),
                               run("234", "i"), run("\t(5,678)")])
        self.body.append(self.statement)
        self.tbl = table([["", "Notes", "2026", "2025"],
                          ["Cash", "12", "8,273,470", "3,855,362"],
                          ["Blank", "13", "", "-"]])
        self.body.append(self.tbl)
        self.body.append(paragraph("\tNotes\t2026\t2025"))
        self.body.append(paragraph("Other\t7\t2,000\t800"))
        self.other = paragraph()
        self.other.append(revision("Reviewed", "Alice", 80))
        self.other.append(revision("Old", "Codex", 81, True))
        self.body.append(self.other)
        self.write()

    def tearDown(self):
        self.tmp.cleanup()

    def write(self, settings=True):
        self.parts = {
            "[Content_Types].xml": b'<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>',
            "_rels/.rels": b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
            "word/_rels/document.xml.rels": b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>',
            "word/document.xml": m._serialized(self.document),
            "word/styles.xml": m._serialized(element("styles")),
            "docProps/core.xml": b'<coreProperties xmlns="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"/>',
            "word/comments.xml": b'<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="100" w:author="Codex"/></w:comments>',
        }
        if settings:
            self.parts["word/settings.xml"] = m._serialized(element("settings"))
        with zipfile.ZipFile(self.path, "w", zipfile.ZIP_DEFLATED) as z:
            for name, raw in self.parts.items():
                z.writestr(name, raw)

    def binding(self, **changes):
        b = {"table": 0, "row": 1, "column": 2, "expected_old": "8,273,470",
             "value": 8273471, "year": 2026,
             "source": {"path": "source.xlsx", "sheet": "A3.1", "cell": "F68"}}
        b.update(changes)
        return b

    def paragraph_binding(self, **changes):
        b = {"paragraph": 1, "segment": 2, "expected_old": "1,234", "value": 5678,
             "year": 2026, "source": {"path": "source.xlsx", "sheet": "PL", "cell": "F4"}}
        b.update(changes)
        return b

    def preview(self, *bindings):
        return m.preview_bindings(self.path, list(bindings), 2026, 2025)

    def export(self, *bindings):
        return m.export_docx(self.path, self.output, list(bindings), 2026, 2025)

    def document_after(self):
        with zipfile.ZipFile(self.output) as z:
            return m._parse(z.read("word/document.xml"))

    def assertRejected(self, binding, reason):
        result = self.preview(binding)
        self.assertEqual(result["rejects"][0]["reason"], reason)

    def test_inspection_indices_include_table_paragraphs(self):
        result = m.inspect_template(self.path)
        self.assertEqual(result["kind"], "docx")
        self.assertEqual(result["paragraphs"][1]["segments"],
                         ["Loss before tax", "6", "1,234", "(5,678)"])
        self.assertFalse(result["paragraphs"][2]["top_level"])
        self.assertEqual(result["paragraphs"][15]["text"], "Other\t7\t2,000\t800")
        self.assertEqual(result["tables"][0]["rows"][1][2], "8,273,470")
        self.assertEqual(result["sha256"], hashlib.sha256(self.path.read_bytes()).hexdigest())

    def test_preview_is_read_only_and_has_context(self):
        original = self.path.read_bytes()
        result = self.preview(self.binding())
        self.assertFalse(result["rejects"])
        self.assertEqual(result["changes"][0]["new"], "8,273,471")
        self.assertIn("Cash", result["changes"][0]["context"])
        self.assertEqual(original, self.path.read_bytes())

    def test_table_export_preserves_comparative_properties_and_package_parts(self):
        before = self.path.read_bytes()
        cell = self.tbl.findall("w:tr", m.NS)[1].findall("w:tc", m.NS)[2]
        original_properties = ET.tostring(cell.find("w:tcPr", m.NS))
        result = self.export(self.binding())
        self.assertTrue(result["written"])
        document = self.document_after()
        cells = document.find("w:body/w:tbl", m.NS).findall("w:tr", m.NS)[1].findall("w:tc", m.NS)
        self.assertEqual(m._cell_text(cells[2]), "8,273,471")
        self.assertEqual(m._cell_text(cells[3]), "3,855,362")
        self.assertEqual(ET.tostring(cells[2].find("w:tcPr", m.NS)), original_properties)
        self.assertEqual(cells[2].find(".//w:delText", m.NS).text, "8,273,470")
        self.assertEqual(self.path.read_bytes(), before)
        with zipfile.ZipFile(self.output) as z:
            for name, raw in self.parts.items():
                if name not in {"word/document.xml", "word/settings.xml"}:
                    self.assertEqual(z.read(name), raw, name)
            settings = m._parse(z.read("word/settings.xml"))
            self.assertEqual(settings.find("w:trackRevisions", m.NS).get(m._tag("val")), "true")

    def test_multirun_number_retains_styles_and_paragraph_properties(self):
        properties = ET.tostring(self.statement.find("w:pPr", m.NS))
        result = self.export(self.paragraph_binding())
        self.assertTrue(result["written"])
        p = self.document_after().find("w:body", m.NS).findall("w:p", m.NS)[1]
        self.assertEqual(m._text(p), "Loss before tax\t6\t5,678\t(5,678)")
        self.assertEqual(ET.tostring(p.find("w:pPr", m.NS)), properties)
        inserted = p.find("w:ins", m.NS).findall("w:r", m.NS)
        self.assertEqual([m._run_text(r) for r in inserted], ["5,", "678"])
        self.assertIsNotNone(inserted[0].find("w:rPr/w:b", m.NS))
        self.assertIsNotNone(inserted[1].find("w:rPr/w:i", m.NS))

    def test_existing_reviewers_dates_and_unique_revision_ids(self):
        self.export(self.binding())
        document = self.document_after()
        revs = [e for e in document.iter() if ET.QName(e).localname in m.REVISION_NAMES]
        prior = {e.get(m._tag("id")): e for e in revs}
        self.assertEqual(prior["80"].get(m._tag("author")), "Alice")
        self.assertEqual(prior["81"].get(m._tag("author")), "Todd")
        self.assertEqual(prior["81"].get(m._tag("date")), "2025-08-31T03:04:05Z")
        ids = [e.get(m._tag("id")) for e in revs]
        self.assertEqual(len(ids), len(set(ids)))
        self.assertTrue(all(e.get(m._tag("author")) == "Todd" for e in revs if int(e.get(m._tag("id"))) > 100))

    def test_comparative_table_column_cannot_be_claimed_as_current_year(self):
        self.assertRejected(self.binding(column=3, expected_old="3,855,362"), "comparative_column_protected")

    def test_comparative_paragraph_segment_cannot_be_claimed_as_current_year(self):
        self.assertRejected(self.paragraph_binding(segment=3, expected_old="(5,678)"), "comparative_column_protected")

    def test_wrong_binding_year_rejected(self):
        self.assertRejected(self.binding(year=2025), "comparative_or_wrong_year_protected")

    def test_no_year_header_means_no_write(self):
        self.tbl.findall("w:tr", m.NS)[0].findall("w:tc", m.NS)[2].find("w:p/w:r/w:t", m.NS).text = "Amount"
        self.write()
        self.assertRejected(self.binding(), "current_year_column_not_identified")

    def test_prior_year_template_does_not_automatically_roll_forward(self):
        header = self.tbl.findall("w:tr", m.NS)[0].findall("w:tc", m.NS)
        header[2].find("w:p/w:r/w:t", m.NS).text = "2025"
        header[3].find("w:p/w:r/w:t", m.NS).text = "2024"
        self.write()
        self.assertRejected(self.binding(), "comparative_column_protected")

    def test_year_headers_and_note_numbers_are_not_amounts(self):
        self.assertRejected(self.binding(row=0, expected_old="2026", confirm_year_like_amount=True), "year_header_not_amount")
        self.assertRejected(self.binding(column=1, expected_old="12"), "note_column_not_amount")
        self.assertRejected(self.paragraph_binding(segment=1, expected_old="6"), "note_column_not_amount")

    def test_year_like_amount_needs_confirmation_but_can_be_filled(self):
        binding = self.paragraph_binding(paragraph=15, expected_old="2,000", value=2101)
        self.assertRejected(binding, "ambiguous_year_like_amount")
        binding["confirm_year_like_amount"] = True
        self.assertFalse(self.preview(binding)["rejects"])

    def test_empty_table_cell_explicit_insertion_keeps_properties(self):
        binding = self.binding(row=2, expected_old="", value=0)
        result = self.export(binding)
        self.assertTrue(result["written"])
        cell = self.document_after().find("w:body/w:tbl", m.NS).findall("w:tr", m.NS)[2].findall("w:tc", m.NS)[2]
        self.assertEqual(m._cell_text(cell), "0")
        self.assertIsNotNone(cell.find("w:tcPr/w:tcW", m.NS))
        self.assertIsNone(cell.find(".//w:del", m.NS))
        self.assertEqual(cell.find(".//w:ins", m.NS).get(m._tag("author")), "Todd")

    def test_empty_tab_segment_inserts_at_exact_position(self):
        self.body.insert(2, paragraph("Blank\t8\t\t90"))
        self.write()
        result = self.export(self.paragraph_binding(paragraph=2, expected_old="", value=-45))
        self.assertTrue(result["written"])
        p = self.document_after().find("w:body", m.NS).findall("w:p", m.NS)[2]
        self.assertEqual(m._text(p), "Blank\t8\t(45)\t90")

    def test_missing_value_is_not_zero(self):
        for value in (None, "", True, float("nan")):
            with self.subTest(value=value):
                self.assertTrue(self.preview(self.binding(value=value))["rejects"])

    def test_expected_old_and_provenance_required(self):
        self.assertRejected(self.binding(expected_old="8,273,469"), "expected_old_mismatch")
        self.assertRejected(self.binding(source={}), "missing_source_provenance")
        binding = self.binding()
        del binding["expected_old"]
        self.assertRejected(binding, "missing_expected_old")

    def test_nonnumeric_targets_and_values_rejected(self):
        self.assertRejected(self.binding(column=0, expected_old="Cash"), "current_year_column_not_identified")
        self.assertRejected(self.binding(value="Amount 10"), "value_not_numeric")

    def test_duplicate_targets_and_invalid_indices_rejected(self):
        result = self.preview(self.binding(), self.binding())
        self.assertEqual(result["rejects"][0]["reason"], "duplicate_target_binding")
        for index in (-1, True, "0", 10000):
            with self.subTest(index=index):
                self.assertTrue(self.preview(self.binding(table=index))["rejects"])

    def test_existing_target_revisions_rejected_without_acceptance(self):
        p = self.tbl.findall("w:tr", m.NS)[1].findall("w:tc", m.NS)[2].find("w:p", m.NS)
        p.append(revision(" extra", "Codex", 101))
        self.write()
        self.assertRejected(self.binding(), "target_has_existing_revisions")

    def test_nested_tables_and_table_paragraph_bindings_rejected(self):
        cell = self.tbl.findall("w:tr", m.NS)[2].findall("w:tc", m.NS)[2]
        cell.append(table([["", "2026", "2025"], ["Nested", "10", "20"]]))
        self.write()
        self.assertTrue(m.inspect_template(self.path)["tables"][1]["nested"])
        self.assertRejected(self.binding(table=1, column=1, expected_old="10"), "nested_table_not_supported")
        self.assertRejected(self.paragraph_binding(paragraph=8), "only_top_level_tab_paragraphs_supported")

    def test_fields_and_merged_cells_rejected(self):
        cell = self.tbl.findall("w:tr", m.NS)[1].findall("w:tc", m.NS)[2]
        ET.SubElement(cell.find("w:p/w:r", m.NS), m._tag("fldChar"), {m._tag("fldCharType"): "begin"})
        self.write()
        self.assertRejected(self.binding(), "unsupported_target_markup")
        ET.SubElement(cell.find("w:tcPr", m.NS), m._tag("gridSpan"), {m._tag("val"): "2"})
        self.write()
        self.assertRejected(self.binding(), "merged_cell_not_supported")

    def test_all_or_nothing_export_and_no_source_overwrite(self):
        result = self.export(self.binding(), self.paragraph_binding(expected_old="wrong"))
        self.assertFalse(result["written"])
        self.assertFalse(self.output.exists())
        before = self.path.read_bytes()
        result = m.export_docx(self.path, self.path, [self.binding()], 2026, 2025)
        self.assertFalse(result["written"])
        self.assertEqual(result["rejects"][0]["reason"], "source_overwrite_prohibited")
        self.assertEqual(before, self.path.read_bytes())

    def test_legacy_doc_export_needs_docx(self):
        legacy = self.directory / "legacy.doc"
        legacy.write_bytes(b"Not a docx; no conversion should be attempted")
        result = m.export_docx(legacy, self.output, [], 2026, 2025)
        self.assertTrue(result["needs_docx_template"])
        self.assertFalse(result["written"])
        self.assertFalse(self.output.exists())

    def test_missing_settings_rejected_without_output(self):
        self.write(settings=False)
        result = self.export(self.binding())
        self.assertFalse(result["written"])
        self.assertEqual(result["rejects"][0]["reason"], "template_missing_settings_part")
        self.assertFalse(self.output.exists())

    def test_currency_on_year_header_supported(self):
        p = self.tbl.findall("w:tr", m.NS)[0].findall("w:tc", m.NS)[2].find("w:p", m.NS)
        p.append(run(" HK$"))
        self.write()
        self.assertFalse(self.preview(self.binding())["rejects"])

    def test_two_current_year_segments_in_same_paragraph(self):
        self.body[0] = paragraph("\tNotes\t2026\t2026\t2025")
        self.body[1] = paragraph("Separate currencies\t4\t1,234\t5,678\t900")
        self.write()
        result = self.export(self.paragraph_binding(value=4321),
                             self.paragraph_binding(segment=3, expected_old="5,678", value=8765))
        self.assertTrue(result["written"])
        p = self.document_after().find("w:body", m.NS).findall("w:p", m.NS)[1]
        self.assertEqual(m._text(p), "Separate currencies\t4\t4,321\t8,765\t900")

    def test_closest_repeated_header_protects_later_prior_year_section(self):
        self.tbl.extend(table([["", "Notes", "2025", "2024"],
                               ["Historical", "10", "123", "456"]]).findall("w:tr", m.NS))
        self.write()
        self.assertRejected(self.binding(row=4, expected_old="123"), "comparative_column_protected")

    def test_table_year_like_amount_requires_explicit_confirmation(self):
        cell = self.tbl.findall("w:tr", m.NS)[1].findall("w:tc", m.NS)[2]
        cell.find("w:p/w:r/w:t", m.NS).text = "2,000"
        self.write()
        self.assertRejected(self.binding(expected_old="2,000"), "ambiguous_year_like_amount")
        self.assertFalse(self.preview(self.binding(expected_old="2,000", confirm_year_like_amount=True))["rejects"])

    def test_merged_neighbor_or_year_header_cannot_shift_column_role(self):
        first = self.tbl.findall("w:tr", m.NS)[1].findall("w:tc", m.NS)[0]
        ET.SubElement(first.find("w:tcPr", m.NS), m._tag("gridSpan"), {m._tag("val"): "2"})
        self.write()
        self.assertRejected(self.binding(), "merged_row_not_supported")
        first.find("w:tcPr", m.NS).remove(first.find("w:tcPr/w:gridSpan", m.NS))
        header = self.tbl.findall("w:tr", m.NS)[0].findall("w:tc", m.NS)[0]
        ET.SubElement(header.find("w:tcPr", m.NS), m._tag("gridSpan"), {m._tag("val"): "2"})
        self.write()
        self.assertRejected(self.binding(), "merged_header_not_supported")

    def test_changed_row_column_count_is_not_assumed_to_align(self):
        row = self.tbl.findall("w:tr", m.NS)[1]
        row.remove(row.findall("w:tc", m.NS)[0])
        self.write()
        self.assertRejected(self.binding(column=2, expected_old="3,855,362"), "column_layout_not_identified")

    def test_track_revisions_uses_settings_schema_order(self):
        settings = element("settings")
        settings.extend([element("zoom", percent=100), element("compat"), element("rsids")])
        self.parts["word/settings.xml"] = m._serialized(settings)
        with zipfile.ZipFile(self.path, "w") as z:
            for name, raw in self.parts.items():
                z.writestr(name, raw)
        self.export(self.binding())
        with zipfile.ZipFile(self.output) as z:
            actual = m._parse(z.read("word/settings.xml"))
        self.assertEqual([ET.QName(child).localname for child in actual],
                         ["zoom", "trackRevisions", "compat", "rsids"])

    def test_load_uses_same_snapshot_for_hash_and_parsed_content(self):
        original = self.path.read_bytes()
        real_read = Path.read_bytes
        changed = False

        def read_and_change(path):
            nonlocal changed
            raw = real_read(path)
            if path == self.path and not changed:
                changed = True
                cell = self.tbl.findall("w:tr", m.NS)[1].findall("w:tc", m.NS)[2]
                cell.find("w:p/w:r/w:t", m.NS).text = "777"
                self.write()
            return raw

        with patch.object(Path, "read_bytes", read_and_change):
            result = m.inspect_template(self.path)
        self.assertEqual(result["sha256"], hashlib.sha256(original).hexdigest())
        self.assertEqual(result["tables"][0]["rows"][1][2], "8,273,470")
        self.assertEqual(m.inspect_template(self.path)["tables"][0]["rows"][1][2], "777")

    def test_change_between_preview_and_export_snapshot_rejects(self):
        real_preview = m.preview_bindings

        def preview_and_change(*args, **kwargs):
            result = real_preview(*args, **kwargs)
            cell = self.tbl.findall("w:tr", m.NS)[1].findall("w:tc", m.NS)[2]
            cell.find("w:p/w:r/w:t", m.NS).text = "777"
            self.write()
            return result

        with patch.object(m, "preview_bindings", side_effect=preview_and_change):
            result = self.export(self.binding(value=999))
        self.assertFalse(result["written"])
        self.assertEqual(result["rejects"][0]["reason"], "template_changed_during_export")
        self.assertFalse(self.output.exists())
        self.assertEqual(m.inspect_template(self.path)["tables"][0]["rows"][1][2], "777")

    def test_change_during_patch_does_not_commit_output_and_cleans_temp(self):
        real_replace = m._replace

        def replace_and_change(*args, **kwargs):
            real_replace(*args, **kwargs)
            cell = self.tbl.findall("w:tr", m.NS)[1].findall("w:tc", m.NS)[3]
            cell.find("w:p/w:r/w:t", m.NS).text = "111"
            self.write()

        with patch.object(m, "_replace", side_effect=replace_and_change):
            result = self.export(self.binding())
        self.assertFalse(result["written"])
        self.assertEqual(result["rejects"][0]["reason"], "template_changed_during_export")
        self.assertFalse(self.output.exists())
        self.assertEqual(list(self.directory.glob("*.docx")), [self.path])

    def test_binding_hash_rejects_template_change_even_when_target_matches(self):
        digest = hashlib.sha256(self.path.read_bytes()).hexdigest()
        cell = self.tbl.findall("w:tr", m.NS)[1].findall("w:tc", m.NS)[3]
        cell.find("w:p/w:r/w:t", m.NS).text = "111"
        self.write()
        self.assertRejected(self.binding(template_sha256=digest), "template_changed_since_binding")


if __name__ == "__main__":
    unittest.main()
