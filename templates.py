"""Explicit numeric Word bindings. Source packages are never overwritten.

Indices are zero-based. Paragraph indices include every w:p below w:body,
including table paragraphs. Table indices include every w:tbl in body order;
nested tables are inspectable but cannot be edited. Only direct-body,
tab-delimited paragraphs and single-paragraph table cells can be bound.

This module deliberately does not infer mappings or convert legacy .doc.
Exports remain subject to render and visual verification before delivery.
"""
from __future__ import annotations

import copy
import hashlib
import importlib.util
import io
import os
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

from lxml import etree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML = "http://www.w3.org/XML/1998/namespace"
NS = {"w": W}
REVISION_NAMES = {
    "ins", "del", "moveFrom", "moveTo", "rPrChange", "pPrChange",
    "tblPrChange", "tblGridChange", "trPrChange", "tcPrChange",
    "sectPrChange", "numberingChange", "cellIns", "cellDel", "cellMerge",
}
NUMERIC = re.compile(r"^[+-]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?$")
NOTE_HEADERS = {"note", "notes", "note no", "notes no", "附注", "附註"}
CURRENCY_HEADERS = {"", "hk$", "hkd", "rmb", "rmb'000", "us$", "usd", "£", "$"}
SETTINGS_BEFORE_TRACK = set("writeProtection view zoom removePersonalInformation removeDateAndTime "
    "displayBackgroundShape printPostScriptOverText printFractionalCharacterWidth printFormsData "
    "embedTrueTypeFonts embedSystemFonts saveSubsetFonts saveFormsData mirrorMargins alignBordersAndEdges "
    "bordersDoNotSurroundHeader bordersDoNotSurroundFooter gutterAtTop hideSpellingErrors hideGrammaticalErrors "
    "activeWritingStyle proofState formsDesign attachedTemplate linkStyles stylePaneFormatFilter "
    "stylePaneSortMethod documentType mailMerge revisionView".split())


def _tag(name: str) -> str:
    return f"{{{W}}}{name}"


def _parse(raw: bytes):
    if b"<!DOCTYPE" in raw.upper():
        raise ValueError("DTD is not supported in a Word template")
    return ET.fromstring(raw, ET.XMLParser(resolve_entities=False, no_network=True))


def _serialized(root) -> bytes:
    return ET.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)


def _numeric(text: str) -> bool:
    value = text.strip()
    if value in {"-", "–", "—"}:
        return True
    if value.startswith("(") and value.endswith(")"):
        value = value[1:-1]
        if value.startswith(("+", "-")):
            return False
    return bool(NUMERIC.fullmatch(value))


def _year_like(text: str) -> bool:
    value = text.strip().replace(",", "")
    return value.isdigit() and 1900 <= int(value) <= 2100


def _text(node) -> str:
    """Current-view text; deleted/moved-from text is not accepted."""
    if node.tag in {_tag("del"), _tag("moveFrom")}:
        return ""
    if node.tag == _tag("t"):
        return node.text or ""
    if node.tag == _tag("tab"):
        return "\t"
    if node.tag in {_tag("br"), _tag("cr")}:
        return "\n"
    return "".join(_text(child) for child in node)


def _cell_text(cell) -> str:
    return "\n".join(_text(p) for p in cell.findall("w:p", NS))


def _load(path):
    path = Path(path)
    raw = path.read_bytes()
    with zipfile.ZipFile(io.BytesIO(raw)) as package:
        if "word/document.xml" not in package.namelist():
            raise ValueError("Not a Word OOXML document")
        root = _parse(package.read("word/document.xml"))
    body = root.find("w:body", NS)
    if body is None:
        raise ValueError("Word document has no body")
    return root, body, hashlib.sha256(raw).hexdigest()


def _legacy_template(path) -> dict:
    # Reuse the task's read-only Word binary piece-table reader. No Office COM.
    reader_path = Path(__file__).parent / "research" / "read_corpus.py"
    report = {"kind": "doc", "sha256": hashlib.sha256(Path(path).read_bytes()).hexdigest(),
              "paragraphs": [], "tables": [], "needs_docx_template": True,
              "read_only": True, "index_contract": "legacy text only; not bindable"}
    try:
        spec = importlib.util.spec_from_file_location("_hk_read_corpus", reader_path)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        result = module.old_doc(Path(path))
        text = result.get("text", "")
        for index, line in enumerate(text.splitlines()):
            segments = line.split("\t")
            report["paragraphs"].append({"index": index, "text": line,
                                        "segments": segments,
                                        "numeric_segments": [i for i, s in enumerate(segments) if _numeric(s)]})
    except Exception as error:
        report["read_error"] = str(error)
    return report


def inspect_template(path) -> dict:
    """Inspect current text and stable XML locators without executing Word."""
    if Path(path).suffix.lower() == ".doc":
        return _legacy_template(path)
    if Path(path).suffix.lower() != ".docx":
        raise ValueError("Template must be .docx or a read-only legacy .doc")
    root, body, digest = _load(path)
    paragraphs = []
    for index, p in enumerate(body.iter(_tag("p"))):
        text = _text(p)
        segments = text.split("\t")
        tracked = any(ET.QName(e).localname in REVISION_NAMES for e in p.iter())
        paragraphs.append({"index": index, "text": text, "segments": segments,
                           "numeric_segments": [i for i, s in enumerate(segments) if _numeric(s)],
                           "top_level": p.getparent() is body,
                           "has_revisions": tracked,
                           "bindable": p.getparent() is body and "\t" in text and not tracked})
    tables = []
    for index, table in enumerate(body.iter(_tag("tbl"))):
        tables.append({"index": index,
                       "nested": any(a.tag == _tag("tbl") for a in table.iterancestors()),
                       "rows": [[_cell_text(cell) for cell in row.findall("w:tc", NS)]
                                for row in table.findall("w:tr", NS)]})
    return {"kind": "docx", "paragraphs": paragraphs, "tables": tables, "sha256": digest,
            "needs_docx_template": False,
            "index_contract": "zero-based; paragraph indices count all body w:p; table indices count all body w:tbl"}


def _index(binding: dict, key: str) -> int:
    value = binding[key]
    if isinstance(value, bool) or not isinstance(value, int) or value < 0:
        raise ValueError(f"invalid_{key}_index")
    return value


def _header(text: str) -> str:
    return re.sub(r"[\s.:：]+", " ", text.strip().lower()).strip()


def _column_year(text: str):
    """Only a standalone year, optionally followed by a currency caption."""
    tokens = text.strip().split()
    if not tokens or not re.fullmatch(r"(?:19|20|21)\d{2}", tokens[0]):
        return None
    if any(_header(token) not in CURRENCY_HEADERS for token in tokens[1:]):
        return None
    return tokens[0]


def _year_header(parts):
    return (any(_column_year(part) for part in parts) and
            all(_column_year(part) or _header(part) in NOTE_HEADERS | CURRENCY_HEADERS
                for part in parts))


def _column_role(header_rows, column, current_year, comparative_year, width=None):
    for parts in header_rows:
        if column < len(parts) and _header(parts[column]) in NOTE_HEADERS:
            raise ValueError("note_column_not_amount")
        if _year_header(parts):
            if width is not None and len(parts) != width:
                raise ValueError("column_layout_not_identified")
            year = _column_year(parts[column]) if column < len(parts) else None
            if year == str(comparative_year):
                raise ValueError("comparative_column_protected")
            if year != str(current_year):
                raise ValueError("current_year_column_not_identified")
            return
    raise ValueError("current_year_column_not_identified")


def _revision_in(node) -> bool:
    return any(ET.QName(e).localname in REVISION_NAMES for e in node.iter())


def _simple_paragraph(p):
    """Reject fields, hyperlinks, drawings, content controls and complex markup."""
    allowed_markers = {"pPr", "r", "bookmarkStart", "bookmarkEnd", "proofErr"}
    for child in p:
        if ET.QName(child).localname not in allowed_markers:
            raise ValueError("unsupported_target_markup")
        if child.tag == _tag("r"):
            for content in child:
                if content.tag not in {_tag("rPr"), _tag("t"), _tag("tab")}:
                    raise ValueError("unsupported_target_markup")


def _locate(body, binding, current_year=None, comparative_year=None):
    if "table" in binding:
        if "paragraph" in binding or "segment" in binding:
            raise ValueError("ambiguous_locator")
        ti, ri, ci = (_index(binding, name) for name in ("table", "row", "column"))
        tables = list(body.iter(_tag("tbl")))
        table = tables[ti]
        if any(a.tag == _tag("tbl") for a in table.iterancestors()):
            raise ValueError("nested_table_not_supported")
        rows = table.findall("w:tr", NS)
        row_cells = rows[ri].findall("w:tc", NS)
        cell = row_cells[ci]
        if cell.find(".//w:tbl", NS) is not None:
            raise ValueError("nested_table_not_supported")
        if _revision_in(cell):
            raise ValueError("target_has_existing_revisions")
        props = cell.find("w:tcPr", NS)
        if props is not None and (props.find("w:vMerge", NS) is not None or
                                  props.find("w:gridSpan", NS) is not None):
            raise ValueError("merged_cell_not_supported")
        if any(c.find("w:tcPr/w:vMerge", NS) is not None or
               c.find("w:tcPr/w:gridSpan", NS) is not None for c in row_cells):
            raise ValueError("merged_row_not_supported")
        if rows[ri].find("w:trPr/w:gridBefore", NS) is not None:
            raise ValueError("column_layout_not_identified")
        paragraphs = cell.findall("w:p", NS)
        nonempty = [p for p in paragraphs if _text(p).strip()]
        if not nonempty and len(paragraphs) == 1:
            p = paragraphs[0]
        elif len(nonempty) != 1:
            raise ValueError("target_not_single_numeric_paragraph")
        else:
            p = nonempty[0]
        old = _text(p)
        row_text = [_cell_text(c) for c in rows[ri].findall("w:tc", NS)]
        if _year_header(row_text):
            raise ValueError("year_header_not_amount")
        header_rows = []
        for candidate in reversed(rows[:ri]):
            cells = candidate.findall("w:tc", NS)
            texts = [_cell_text(c) for c in cells]
            if _year_header(texts) and (any(c.find("w:tcPr/w:vMerge", NS) is not None or
                    c.find("w:tcPr/w:gridSpan", NS) is not None for c in cells) or
                    candidate.find("w:trPr/w:gridBefore", NS) is not None):
                raise ValueError("merged_header_not_supported")
            header_rows.append(texts)
            if _year_header(texts):
                break
        _column_role(header_rows, ci, current_year, comparative_year, len(row_cells))
        location = {"table": ti, "row": ri, "column": ci}
        start, end = 0, len(old)
        context = " | ".join(row_text)
    elif "paragraph" in binding:
        if any(k in binding for k in ("row", "column")):
            raise ValueError("ambiguous_locator")
        pi, si = (_index(binding, name) for name in ("paragraph", "segment"))
        paragraphs = list(body.iter(_tag("p")))
        p = paragraphs[pi]
        if p.getparent() is not body:
            raise ValueError("only_top_level_tab_paragraphs_supported")
        if _revision_in(p):
            raise ValueError("target_has_existing_revisions")
        text = _text(p)
        if "\t" not in text:
            raise ValueError("paragraph_not_tab_delimited")
        segments = text.split("\t")
        old = segments[si]
        if _year_header(segments):
            raise ValueError("year_header_not_amount")
        _column_role((_text(candidate).split("\t")
                      for candidate in reversed(paragraphs[:pi])
                      if candidate.getparent() is body and
                      len(_text(candidate).split("\t")) == len(segments)),
                     si, current_year, comparative_year)
        start = sum(len(t) + 1 for t in segments[:si])
        end = start + len(old)
        location = {"paragraph": pi, "segment": si}
        context = text
    else:
        raise ValueError("missing_locator")
    _simple_paragraph(p)
    return p, start, end, old, location, context


def _new_text(value, old: str) -> str:
    if value is None or isinstance(value, bool):
        raise ValueError("missing_numeric_value")
    if isinstance(value, str):
        result = value.strip()
        if not _numeric(result):
            raise ValueError("value_not_numeric")
    else:
        try:
            amount = Decimal(str(value))
        except (InvalidOperation, ValueError):
            raise ValueError("value_not_numeric") from None
        if not amount.is_finite():
            raise ValueError("value_not_finite")
        result = format(abs(amount), ",f")
        if "." in result:
            result = result.rstrip("0").rstrip(".")
        if amount < 0:
            result = f"({result})"
    leading = old[:len(old) - len(old.lstrip())]
    trailing = old[len(old.rstrip()):] if old.strip() else ""
    return leading + result + trailing


def preview_bindings(path, bindings, current_year, comparative_year) -> dict:
    """Validate every binding and return an explicit, non-mutating change set."""
    if Path(path).suffix.lower() == ".doc":
        return {"kind": "doc", "changes": [], "rejects": [{"reason": "needs_docx_template"}],
                "needs_docx_template": True}
    _, body, digest = _load(path)
    changes, rejects, used = [], [], set()
    if str(current_year) == str(comparative_year):
        return {"kind": "docx", "sha256": digest, "changes": [],
                "rejects": [{"reason": "current_and_comparative_year_must_differ"}]}
    for index, binding in enumerate(bindings):
        try:
            if not isinstance(binding, dict):
                raise ValueError("binding_not_object")
            if str(binding.get("year")) != str(current_year):
                raise ValueError("comparative_or_wrong_year_protected")
            if "template_sha256" in binding and binding["template_sha256"] != digest:
                raise ValueError("template_changed_since_binding")
            source = binding.get("source")
            if not isinstance(source, dict) or not source or not any(source.values()):
                raise ValueError("missing_source_provenance")
            if "expected_old" not in binding:
                raise ValueError("missing_expected_old")
            p, start, end, old, location, context = _locate(body, binding, current_year, comparative_year)
            locator_key = tuple(sorted(location.items()))
            if locator_key in used:
                raise ValueError("duplicate_target_binding")
            used.add(locator_key)
            if old.strip() != str(binding["expected_old"]).strip():
                raise ValueError("expected_old_mismatch")
            if old.strip() and not _numeric(old):
                raise ValueError("target_not_pure_numeric")
            if _year_like(old) and binding.get("confirm_year_like_amount") is not True:
                raise ValueError("ambiguous_year_like_amount")
            new = _new_text(binding.get("value"), old)
            if old == new:
                continue
            changes.append({"binding_index": index, "location": location, "old": old, "new": new,
                            "year": str(current_year), "source": copy.deepcopy(source), "context": context})
        except (ValueError, KeyError, IndexError, TypeError) as error:
            rejects.append({"binding_index": index, "reason": str(error),
                            "binding": copy.deepcopy(binding)})
    return {"kind": "docx", "sha256": digest, "changes": changes, "rejects": rejects,
            "needs_docx_template": False, "requires_visual_verification": True}


def _run_text(run) -> str:
    return "".join("\t" if c.tag == _tag("tab") else (c.text or "")
                   for c in run if c.tag in {_tag("t"), _tag("tab")})


def _slice_run(run, start: int, end: int, deletion: bool = False):
    result = ET.Element(run.tag, attrib=dict(run.attrib))
    props = run.find("w:rPr", NS)
    if props is not None:
        result.append(copy.deepcopy(props))
    offset = 0
    for child in run:
        if child.tag == _tag("rPr"):
            continue
        text = "\t" if child.tag == _tag("tab") else (child.text or "")
        left, right = max(start - offset, 0), min(end - offset, len(text))
        if left < right:
            fragment = copy.deepcopy(child)
            if fragment.tag != _tag("tab"):
                fragment.text = text[left:right]
                if deletion:
                    fragment.tag = _tag("delText")
                fragment.set(f"{{{XML}}}space", "preserve")
            result.append(fragment)
        offset += len(text)
    return result


def _insertion(new, styles, author, date, next_id):
    inserted = ET.Element(_tag("ins"), attrib={_tag("id"): str(next_id[0]),
                         _tag("author"): author, _tag("date"): date})
    next_id[0] += 1
    offset = 0
    for index, (style, width) in enumerate(styles):
        fragment = new[offset:] if index == len(styles)-1 else new[offset:offset+width]
        offset += len(fragment)
        if not fragment:
            continue
        new_run = ET.Element(_tag("r"), attrib=dict(style.attrib))
        props = style.find("w:rPr", NS)
        if props is not None:
            new_run.append(copy.deepcopy(props))
        value = ET.SubElement(new_run, _tag("t"))
        value.set(f"{{{XML}}}space", "preserve")
        value.text = fragment
        inserted.append(new_run)
    return inserted


def _replace(p, start, end, new, author, date, next_id):
    runs = []
    offset = 0
    for child in p:
        if child.tag == _tag("r"):
            text = _run_text(child)
            runs.append((child, offset, offset + len(text)))
            offset += len(text)
    touched = [(r, a, b) for r, a, b in runs if a < end and b > start]
    if start == end:
        # An explicitly bound blank is an insertion, never an inferred zero.
        surrounding = next(((r, a, b) for r, a, b in runs if a <= start <= b), None)
        if surrounding is None:
            style = ET.Element(_tag("r"))
            p.append(_insertion(new, [(style, len(new))], author, date, next_id))
            return
        run, a, b = surrounding
        at = p.index(run)
        items = []
        if start > a:
            items.append(_slice_run(run, 0, start-a))
        items.append(_insertion(new, [(run, len(new))], author, date, next_id))
        if start < b:
            items.append(_slice_run(run, start-a, b-a))
        p.remove(run)
        for item in reversed(items):
            p.insert(at, item)
        return
    if not touched:
        raise ValueError("numeric_target_has_no_direct_runs")
    styles = [(run, min(end, b)-max(start, a)) for run, a, b in touched]
    replacements = {}
    for run, a, b in touched:
        items = []
        if a < start:
            items.append(_slice_run(run, 0, start-a))
        deleted = ET.Element(_tag("del"), attrib={_tag("id"): str(next_id[0]),
                            _tag("author"): author, _tag("date"): date})
        next_id[0] += 1
        deleted.append(_slice_run(run, max(start-a, 0), min(end-a, b-a), deletion=True))
        items.append(deleted)
        if run is touched[-1][0]:
            items.append(_insertion(new, styles, author, date, next_id))
        if b > end:
            items.append(_slice_run(run, end-a, b-a))
        replacements[run] = items
    for run in list(p):
        if run in replacements:
            at = p.index(run)
            p.remove(run)
            for item in reversed(replacements[run]):
                p.insert(at, item)


def export_docx(template_path, output_path, bindings, current_year, comparative_year, author="Todd") -> dict:
    """Export all-or-nothing. Reject unsafe targets; preserve original revisions."""
    bindings = copy.deepcopy(bindings)
    preview = preview_bindings(template_path, bindings, current_year, comparative_year)
    preview.update({"written": False, "author": author})
    if preview["rejects"] or not preview["changes"]:
        return preview
    template_path, output_path = Path(template_path), Path(output_path)
    if template_path.suffix.lower() != ".docx" or output_path.suffix.lower() != ".docx":
        preview["rejects"].append({"reason": "needs_docx_template_and_output"})
        return preview
    if template_path.resolve() == output_path.resolve():
        preview["rejects"].append({"reason": "source_overwrite_prohibited"})
        return preview
    if not isinstance(author, str) or not author.strip():
        preview["rejects"].append({"reason": "missing_revision_author"})
        return preview
    snapshot = template_path.read_bytes()
    if hashlib.sha256(snapshot).hexdigest() != preview["sha256"]:
        preview["rejects"].append({"reason": "template_changed_during_export"})
        return preview
    with zipfile.ZipFile(io.BytesIO(snapshot)) as source:
        entries = source.infolist()
        data = {item.filename: source.read(item.filename) for item in entries}
    xml_roots, largest_id = {}, -1
    for name, raw in data.items():
        if name.startswith("word/") and name.endswith(".xml"):
            root = _parse(raw)
            xml_roots[name] = root
            for node in root.iter():
                node_id = node.get(_tag("id"))
                if node_id and node_id.isdigit():
                    largest_id = max(largest_id, int(node_id))
    document = xml_roots["word/document.xml"]
    body = document.find("w:body", NS)
    counter = [largest_id + 1]
    date = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    # Descending offsets make multiple tab-segment bindings independent.
    patches = []
    for change in preview["changes"]:
        binding = bindings[change["binding_index"]]
        p, start, end, old, _, _ = _locate(body, binding, current_year, comparative_year)
        if old != change["old"] or old.strip() != str(binding["expected_old"]).strip():
            preview["rejects"].append({"binding_index": change["binding_index"],
                                      "reason": "template_changed_during_export"})
            return preview
        patches.append((p, start, end, change["new"]))
    patches.sort(key=lambda item: item[1], reverse=True)
    for p, start, end, new in patches:
        _replace(p, start, end, new, author, date, counter)
    renamed = []
    for name, root in xml_roots.items():
        changed = name == "word/document.xml"
        for node in root.iter():
            if ET.QName(node).localname in REVISION_NAMES:
                old_author = node.get(_tag("author"), "")
                if old_author.lower().startswith("codex"):
                    node.set(_tag("author"), "Todd")
                    renamed.append({"part": name, "id": node.get(_tag("id")), "old_author": old_author})
                    changed = True
        if changed:
            data[name] = _serialized(root)
    settings_name = "word/settings.xml"
    if settings_name in xml_roots:
        settings = xml_roots[settings_name]
    else:
        # A complete template normally has settings. Do not manufacture an
        # unregistered package part without relationships/content type wiring.
        preview["rejects"].append({"reason": "template_missing_settings_part"})
        return preview
    tracker = settings.find("w:trackRevisions", NS)
    if tracker is None:
        tracker = ET.Element(_tag("trackRevisions"))
        at = next((index for index, child in enumerate(settings)
                   if ET.QName(child).localname not in SETTINGS_BEFORE_TRACK), len(settings))
        settings.insert(at, tracker)
    tracker.set(_tag("val"), "true")
    data[settings_name] = _serialized(settings)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = None
    try:
        with tempfile.NamedTemporaryFile(dir=output_path.parent, suffix=".docx", delete=False) as stream:
            temp_path = Path(stream.name)
        with zipfile.ZipFile(temp_path, "w") as package:
            for item in entries:
                package.writestr(item, data[item.filename])
        try:
            source_unchanged = hashlib.sha256(template_path.read_bytes()).hexdigest() == preview["sha256"]
        except OSError:
            source_unchanged = False
        if not source_unchanged:
            preview["rejects"].append({"reason": "template_changed_during_export"})
            return preview
        os.replace(temp_path, output_path)
    finally:
        if temp_path is not None and temp_path.exists():
            temp_path.unlink()
    preview.update({"written": True, "output_path": str(output_path),
                    "renamed_revisions": renamed,
                    "sha256_output": hashlib.sha256(output_path.read_bytes()).hexdigest()})
    return preview
