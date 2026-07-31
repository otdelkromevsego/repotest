#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Export XLSX/XLSM/PDF files into repository-friendly JSON/JSONL snapshots.

Design principles:
- Originals remain the source of truth.
- JSONL is sharded for readable diffs and streaming.
- Formula cached values are preserved; formulas are written to a sidecar.
- XLSM VBA source is extracted through LibreOffice when available.
"""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import html
import json
import math
import os
import posixpath
import re
import shutil
import subprocess
import tempfile
import unicodedata
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, Iterator, List, Optional, Tuple
from xml.etree import ElementTree as ET

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CORE = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
NS_DC = "http://purl.org/dc/elements/1.1/"
NS_DCTERMS = "http://purl.org/dc/terms/"
NS_EP = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"

MAIN = f"{{{NS_MAIN}}}"
REL = f"{{{NS_REL}}}"
PKG_REL = f"{{{NS_PKG_REL}}}"

CYR_MAP = str.maketrans({
    "А":"A","Б":"B","В":"V","Г":"G","Д":"D","Е":"E","Ё":"E","Ж":"Zh","З":"Z","И":"I","Й":"Y","К":"K","Л":"L","М":"M","Н":"N","О":"O","П":"P","Р":"R","С":"S","Т":"T","У":"U","Ф":"F","Х":"Kh","Ц":"Ts","Ч":"Ch","Ш":"Sh","Щ":"Shch","Ъ":"","Ы":"Y","Ь":"","Э":"E","Ю":"Yu","Я":"Ya",
    "а":"a","б":"b","в":"v","г":"g","д":"d","е":"e","ё":"e","ж":"zh","з":"z","и":"i","й":"y","к":"k","л":"l","м":"m","н":"n","о":"o","п":"p","р":"r","с":"s","т":"t","у":"u","ф":"f","х":"kh","ц":"ts","ч":"ch","ш":"sh","щ":"shch","ъ":"","ы":"y","ь":"","э":"e","ю":"yu","я":"ya",
})

BUILTIN_DATE_FMT_IDS = set(range(14, 23)) | set(range(27, 37)) | set(range(45, 48)) | set(range(50, 59))


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def json_dump(path: Path, data: Any, *, indent: int = 2) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="\n") as f:
        json.dump(data, f, ensure_ascii=False, indent=indent, sort_keys=False)
        f.write("\n")


def slugify(text: str, max_len: int = 100) -> str:
    text = text.translate(CYR_MAP)
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"[^A-Za-z0-9]+", "-", text).strip("-").lower()
    return (text[:max_len].rstrip("-") or "source")


def col_to_num(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + ord(ch.upper()) - 64
    return n


def num_to_col(n: int) -> str:
    out = []
    while n:
        n, r = divmod(n - 1, 26)
        out.append(chr(65 + r))
    return "".join(reversed(out)) or "A"


def split_address(address: str) -> Tuple[int, int]:
    m = re.match(r"^\$?([A-Z]+)\$?(\d+)$", address or "")
    if not m:
        return 0, 0
    return int(m.group(2)), col_to_num(m.group(1))


def clean_xml_target(base_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(base_part), target))


def read_relationships(z: zipfile.ZipFile, rels_part: str) -> Dict[str, Dict[str, str]]:
    if rels_part not in z.namelist():
        return {}
    root = ET.fromstring(z.read(rels_part))
    out = {}
    for rel in root:
        out[rel.attrib.get("Id", "")] = {
            "target": rel.attrib.get("Target", ""),
            "type": rel.attrib.get("Type", ""),
            "target_mode": rel.attrib.get("TargetMode", "Internal"),
        }
    return out


def read_shared_strings(z: zipfile.ZipFile) -> List[str]:
    part = "xl/sharedStrings.xml"
    if part not in z.namelist():
        return []
    strings: List[str] = []
    root = ET.fromstring(z.read(part))
    for si in root.findall(f"{MAIN}si"):
        texts = [t.text or "" for t in si.iter(f"{MAIN}t")]
        strings.append("".join(texts))
    return strings


def read_styles(z: zipfile.ZipFile) -> Tuple[List[Dict[str, Any]], Dict[int, str]]:
    part = "xl/styles.xml"
    if part not in z.namelist():
        return [], {}
    root = ET.fromstring(z.read(part))
    custom: Dict[int, str] = {}
    numfmts = root.find(f"{MAIN}numFmts")
    if numfmts is not None:
        for item in numfmts.findall(f"{MAIN}numFmt"):
            try:
                custom[int(item.attrib["numFmtId"])] = item.attrib.get("formatCode", "")
            except Exception:
                pass
    xfs: List[Dict[str, Any]] = []
    cellxfs = root.find(f"{MAIN}cellXfs")
    if cellxfs is not None:
        for i, xf in enumerate(cellxfs.findall(f"{MAIN}xf")):
            num_fmt_id = int(xf.attrib.get("numFmtId", "0"))
            xfs.append({
                "style_index": i,
                "num_fmt_id": num_fmt_id,
                "num_fmt_code": custom.get(num_fmt_id),
                "font_id": int(xf.attrib.get("fontId", "0")),
                "fill_id": int(xf.attrib.get("fillId", "0")),
                "border_id": int(xf.attrib.get("borderId", "0")),
                "xf_id": int(xf.attrib.get("xfId", "0")),
                "apply_number_format": xf.attrib.get("applyNumberFormat") == "1",
            })
    return xfs, custom


def is_date_format(num_fmt_id: int, fmt: Optional[str]) -> bool:
    if num_fmt_id in BUILTIN_DATE_FMT_IDS:
        return True
    if not fmt:
        return False
    s = fmt.lower()
    s = re.sub(r'"[^"]*"', "", s)
    s = re.sub(r"\\.", "", s)
    s = re.sub(r"\[[^\]]*\]", "", s)
    s = re.sub(r"_[^ ]|\*.", "", s)
    return bool(re.search(r"(^|[^a-z])[ymdhis]+([^a-z]|$)", s))


def excel_serial_to_iso(value: float, date1904: bool) -> str:
    base = dt.datetime(1904, 1, 1) if date1904 else dt.datetime(1899, 12, 30)
    try:
        result = base + dt.timedelta(days=float(value))
    except Exception:
        return str(value)
    if result.time() == dt.time(0, 0):
        return result.date().isoformat()
    if result.microsecond:
        return result.isoformat(timespec="milliseconds")
    return result.isoformat(timespec="seconds")


def parse_number(text: Optional[str]) -> Optional[Any]:
    if text is None or text == "":
        return None
    try:
        n = float(text)
        if math.isfinite(n) and n.is_integer() and not re.search(r"[.eE]", text):
            return int(n)
        return n
    except Exception:
        return text


def decode_cell(c: ET.Element, shared: List[str], styles: List[Dict[str, Any]], date1904: bool) -> Dict[str, Any]:
    address = c.attrib.get("r", "")
    row, col = split_address(address)
    cell_type = c.attrib.get("t", "n")
    style_idx = int(c.attrib.get("s", "0")) if c.attrib.get("s", "0").isdigit() else 0
    v_el = c.find(f"{MAIN}v")
    raw = v_el.text if v_el is not None else None
    value: Any = None

    if cell_type == "s":
        try:
            value = shared[int(raw)] if raw is not None else None
        except Exception:
            value = raw
    elif cell_type == "inlineStr":
        is_el = c.find(f"{MAIN}is")
        value = "".join((t.text or "") for t in is_el.iter(f"{MAIN}t")) if is_el is not None else None
    elif cell_type == "b":
        value = raw == "1"
    elif cell_type == "e":
        value = raw
    elif cell_type in ("str", "d"):
        value = raw
    else:
        value = parse_number(raw)
        if isinstance(value, (int, float)) and style_idx < len(styles):
            st = styles[style_idx]
            if is_date_format(st["num_fmt_id"], st.get("num_fmt_code")):
                value = excel_serial_to_iso(float(value), date1904)

    f_el = c.find(f"{MAIN}f")
    formula = None
    formula_attrs = None
    if f_el is not None:
        formula = f_el.text
        formula_attrs = dict(f_el.attrib)

    return {
        "address": address,
        "row": row,
        "column": col,
        "column_letter": num_to_col(col),
        "value": value,
        "raw_value": raw,
        "cell_type": cell_type,
        "style_index": style_idx,
        "formula": formula,
        "formula_attributes": formula_attrs,
    }


def iter_sheet_rows(z: zipfile.ZipFile, sheet_part: str, shared: List[str], styles: List[Dict[str, Any]], date1904: bool) -> Iterator[Tuple[int, Dict[int, Dict[str, Any]], Dict[str, Any]]]:
    with z.open(sheet_part) as f:
        context = ET.iterparse(f, events=("end",))
        for event, elem in context:
            if elem.tag == f"{MAIN}row":
                row_num = int(elem.attrib.get("r", "0") or "0")
                cells: Dict[int, Dict[str, Any]] = {}
                for c in elem.findall(f"{MAIN}c"):
                    cd = decode_cell(c, shared, styles, date1904)
                    cells[cd["column"]] = cd
                row_props = {
                    "hidden": elem.attrib.get("hidden") == "1",
                    "height": float(elem.attrib["ht"]) if "ht" in elem.attrib else None,
                    "outline_level": int(elem.attrib.get("outlineLevel", "0")),
                    "collapsed": elem.attrib.get("collapsed") == "1",
                }
                yield row_num, cells, row_props
                elem.clear()


def cell_record_value(cell: Dict[str, Any]) -> Any:
    if cell.get("formula") is not None or cell.get("formula_attributes") is not None:
        if cell.get("value") is not None:
            return cell["value"]
        attrs = cell.get("formula_attributes") or {}
        payload: Dict[str, Any] = {"_formula": cell.get("formula")}
        if attrs:
            payload["_formula_attributes"] = attrs
        payload["_value"] = None
        return payload
    return cell.get("value")


def detect_header(first_rows: List[Tuple[int, Dict[int, Dict[str, Any]], Dict[str, Any]]], max_col: int) -> Optional[int]:
    best: Optional[Tuple[float, int]] = None
    for row_num, cells, _ in first_rows:
        vals = [c.get("value") for c in cells.values() if c.get("value") not in (None, "")]
        if len(vals) < 2:
            continue
        strings = sum(isinstance(v, str) and not re.fullmatch(r"[-+]?\d+(?:[.,]\d+)?", v.strip()) for v in vals)
        numerics = sum(isinstance(v, (int, float)) and not isinstance(v, bool) for v in vals)
        formula_count = sum(c.get("formula") is not None or c.get("formula_attributes") is not None for c in cells.values())
        score = len(vals) + 0.75 * strings - 0.5 * numerics - 0.4 * formula_count - 0.01 * row_num
        if strings / max(1, len(vals)) < 0.45:
            score -= 3
        if best is None or score > best[0]:
            best = (score, row_num)
    if best is None:
        return None
    return best[1]


def normalize_headers(header_cells: Dict[int, Dict[str, Any]], max_col: int) -> Tuple[Dict[int, str], List[Dict[str, Any]]]:
    keys: Dict[int, str] = {}
    inventory: List[Dict[str, Any]] = []
    used: Dict[str, int] = {}
    for col in range(1, max_col + 1):
        raw = header_cells.get(col, {}).get("value") if col in header_cells else None
        base = str(raw).strip().replace("\r", " ").replace("\n", " ") if raw not in (None, "") else f"_col_{num_to_col(col)}"
        base = re.sub(r"\s+", " ", base)
        used[base] = used.get(base, 0) + 1
        key = base if used[base] == 1 else f"{base}__{used[base]}"
        keys[col] = key
        inventory.append({"column": col, "letter": num_to_col(col), "source_header": raw, "json_key": key})
    return keys, inventory


@dataclass
class ShardInfo:
    file: str
    rows: int
    first_row: Optional[int]
    last_row: Optional[int]
    bytes: int
    sha256: str


class JsonlShardWriter:
    def __init__(self, folder: Path, prefix: str, max_bytes: int = 5_000_000, max_rows: int = 5000):
        self.folder = folder
        self.prefix = prefix
        self.max_bytes = max_bytes
        self.max_rows = max_rows
        self.folder.mkdir(parents=True, exist_ok=True)
        self.part = 0
        self.handle = None
        self.current_path: Optional[Path] = None
        self.current_bytes = 0
        self.current_rows = 0
        self.first_row = None
        self.last_row = None
        self.parts: List[ShardInfo] = []

    def _open(self) -> None:
        self.part += 1
        self.current_path = self.folder / f"{self.prefix}.part-{self.part:05d}.jsonl"
        self.handle = self.current_path.open("w", encoding="utf-8", newline="\n")
        self.current_bytes = 0
        self.current_rows = 0
        self.first_row = None
        self.last_row = None

    def _close(self) -> None:
        if self.handle is None or self.current_path is None:
            return
        self.handle.close()
        self.parts.append(ShardInfo(
            file=self.current_path.name,
            rows=self.current_rows,
            first_row=self.first_row,
            last_row=self.last_row,
            bytes=self.current_path.stat().st_size,
            sha256=sha256_file(self.current_path),
        ))
        self.handle = None
        self.current_path = None

    def write(self, obj: Dict[str, Any], row_num: Optional[int] = None) -> None:
        line = json.dumps(obj, ensure_ascii=False, separators=(",", ":")) + "\n"
        size = len(line.encode("utf-8"))
        if self.handle is None:
            self._open()
        elif self.current_rows > 0 and (self.current_bytes + size > self.max_bytes or self.current_rows >= self.max_rows):
            self._close()
            self._open()
        assert self.handle is not None
        self.handle.write(line)
        self.current_bytes += size
        self.current_rows += 1
        if row_num is not None:
            if self.first_row is None:
                self.first_row = row_num
            self.last_row = row_num

    def close(self) -> List[Dict[str, Any]]:
        self._close()
        return [s.__dict__ for s in self.parts]


def parse_dimension(ref: Optional[str]) -> Tuple[int, int]:
    if not ref:
        return 0, 0
    last = ref.split(":")[-1]
    return split_address(last)


def extract_sheet_static_meta(z: zipfile.ZipFile, sheet_part: str) -> Dict[str, Any]:
    root = ET.fromstring(z.read(sheet_part))
    dim_el = root.find(f"{MAIN}dimension")
    dimension = dim_el.attrib.get("ref") if dim_el is not None else None
    merges_el = root.find(f"{MAIN}mergeCells")
    merges = [m.attrib.get("ref") for m in merges_el.findall(f"{MAIN}mergeCell")] if merges_el is not None else []
    cols_out = []
    cols_el = root.find(f"{MAIN}cols")
    if cols_el is not None:
        for col in cols_el.findall(f"{MAIN}col"):
            cols_out.append({
                "min": int(col.attrib.get("min", "0")),
                "max": int(col.attrib.get("max", "0")),
                "width": float(col.attrib["width"]) if "width" in col.attrib else None,
                "hidden": col.attrib.get("hidden") == "1",
                "style": int(col.attrib.get("style", "0")),
                "outline_level": int(col.attrib.get("outlineLevel", "0")),
                "collapsed": col.attrib.get("collapsed") == "1",
            })
    pane = None
    pane_el = root.find(f"{MAIN}sheetViews/{MAIN}sheetView/{MAIN}pane")
    if pane_el is not None:
        pane = dict(pane_el.attrib)
    auto_filter_el = root.find(f"{MAIN}autoFilter")
    auto_filter = auto_filter_el.attrib.get("ref") if auto_filter_el is not None else None

    rels_part = posixpath.join(posixpath.dirname(sheet_part), "_rels", posixpath.basename(sheet_part) + ".rels")
    rels = read_relationships(z, rels_part)
    hyperlinks = []
    hyperlinks_el = root.find(f"{MAIN}hyperlinks")
    if hyperlinks_el is not None:
        for h in hyperlinks_el.findall(f"{MAIN}hyperlink"):
            rid = h.attrib.get(f"{REL}id")
            rel = rels.get(rid or "", {})
            hyperlinks.append({
                "ref": h.attrib.get("ref"),
                "display": h.attrib.get("display"),
                "location": h.attrib.get("location"),
                "tooltip": h.attrib.get("tooltip"),
                "target": rel.get("target"),
                "target_mode": rel.get("target_mode"),
            })

    tables = []
    table_parts = root.find(f"{MAIN}tableParts")
    if table_parts is not None:
        for tp in table_parts.findall(f"{MAIN}tablePart"):
            rid = tp.attrib.get(f"{REL}id")
            rel = rels.get(rid or "")
            if not rel:
                continue
            table_part = clean_xml_target(sheet_part, rel["target"])
            if table_part in z.namelist():
                tr = ET.fromstring(z.read(table_part))
                tables.append({
                    "name": tr.attrib.get("name"),
                    "display_name": tr.attrib.get("displayName"),
                    "ref": tr.attrib.get("ref"),
                    "totals_row_shown": tr.attrib.get("totalsRowShown"),
                    "columns": [c.attrib.get("name") for c in tr.findall(f"{MAIN}tableColumns/{MAIN}tableColumn")],
                })

    drawings = []
    for rid, rel in rels.items():
        if rel["type"].endswith("/drawing") or rel["type"].endswith("/legacyDrawing"):
            drawings.append({"relationship_id": rid, "target": clean_xml_target(sheet_part, rel["target"]), "type": rel["type"]})

    return {
        "dimension": dimension,
        "merged_ranges": merges,
        "columns": cols_out,
        "freeze_pane": pane,
        "auto_filter": auto_filter,
        "hyperlinks": hyperlinks,
        "tables": tables,
        "drawings": drawings,
    }


def read_doc_properties(z: zipfile.ZipFile) -> Dict[str, Any]:
    out: Dict[str, Any] = {}
    if "docProps/core.xml" in z.namelist():
        root = ET.fromstring(z.read("docProps/core.xml"))
        for child in root:
            key = child.tag.split("}")[-1]
            out[key] = child.text
    if "docProps/app.xml" in z.namelist():
        root = ET.fromstring(z.read("docProps/app.xml"))
        app = {}
        for child in root:
            key = child.tag.split("}")[-1]
            if len(child) == 0:
                app[key] = child.text
        out["application"] = app
    return out


def export_workbook(source: Path, out_dir: Path) -> Dict[str, Any]:
    slug = slugify(source.stem)
    target = out_dir / slug
    target.mkdir(parents=True, exist_ok=True)
    source_hash = sha256_file(source)

    with zipfile.ZipFile(source) as z:
        names = set(z.namelist())
        wb_root = ET.fromstring(z.read("xl/workbook.xml"))
        wb_rels = read_relationships(z, "xl/_rels/workbook.xml.rels")
        shared = read_shared_strings(z)
        styles, custom_formats = read_styles(z)
        workbook_pr = wb_root.find(f"{MAIN}workbookPr")
        date1904 = workbook_pr is not None and workbook_pr.attrib.get("date1904") == "1"
        book_views = wb_root.find(f"{MAIN}bookViews/{MAIN}workbookView")
        active_tab = int(book_views.attrib.get("activeTab", "0")) if book_views is not None else 0
        calc_pr = wb_root.find(f"{MAIN}calcPr")
        defined_names = []
        dn_root = wb_root.find(f"{MAIN}definedNames")
        if dn_root is not None:
            for dn in dn_root.findall(f"{MAIN}definedName"):
                defined_names.append({
                    "name": dn.attrib.get("name"),
                    "local_sheet_id": int(dn.attrib["localSheetId"]) if "localSheetId" in dn.attrib else None,
                    "hidden": dn.attrib.get("hidden") == "1",
                    "text": dn.text,
                })

        sheets_meta = []
        sheets_el = wb_root.find(f"{MAIN}sheets")
        if sheets_el is None:
            raise ValueError(f"No sheets in {source}")

        for idx, sheet in enumerate(sheets_el.findall(f"{MAIN}sheet"), start=1):
            name = sheet.attrib.get("name", f"Sheet{idx}")
            rid = sheet.attrib.get(f"{REL}id", "")
            rel = wb_rels.get(rid)
            if not rel:
                continue
            sheet_part = clean_xml_target("xl/workbook.xml", rel["target"])
            static = extract_sheet_static_meta(z, sheet_part)
            dim_row, dim_col = parse_dimension(static.get("dimension"))

            first_rows = []
            max_observed_col = 0
            for row_num, cells, props in iter_sheet_rows(z, sheet_part, shared, styles, date1904):
                if cells:
                    max_observed_col = max(max_observed_col, max(cells))
                if row_num <= 30:
                    first_rows.append((row_num, cells, props))
                elif len(first_rows) >= 30:
                    break
            max_col = max(dim_col, max_observed_col, 1)
            header_row = detect_header(first_rows, max_col)
            header_cells: Dict[int, Dict[str, Any]] = {}
            for rn, cells, _ in first_rows:
                if rn == header_row:
                    header_cells = cells
                    break
            if header_row is not None:
                header_map, header_inventory = normalize_headers(header_cells, max_col)
            else:
                header_map = {c: num_to_col(c) for c in range(1, max_col + 1)}
                header_inventory = [{"column": c, "letter": num_to_col(c), "source_header": None, "json_key": num_to_col(c)} for c in range(1, max_col + 1)]

            sheet_slug = f"{idx:02d}-{slugify(name, 60)}"
            sheet_dir = target / "sheets" / sheet_slug
            records_writer = JsonlShardWriter(sheet_dir, "records", max_bytes=5_000_000, max_rows=max(500, min(5000, 50000 // max_col)))
            formulas_writer = JsonlShardWriter(sheet_dir, "formulas", max_bytes=5_000_000, max_rows=10000)
            preamble = []
            preview = []
            formula_count = 0
            cell_count = 0
            nonempty_rows = 0
            row_properties = []
            shared_formula_masters: Dict[str, Dict[str, Any]] = {}

            for row_num, cells, props in iter_sheet_rows(z, sheet_part, shared, styles, date1904):
                if not cells:
                    continue
                nonempty_rows += 1
                cell_count += len(cells)
                if any(v not in (None, False, 0) for v in props.values()):
                    rp = {"row": row_num, **props}
                    row_properties.append(rp)

                raw_row_map = {num_to_col(c): cell_record_value(cd) for c, cd in sorted(cells.items())}
                if header_row is not None and row_num < header_row:
                    preamble.append({"_row": row_num, "cells": raw_row_map})
                    continue
                if header_row is not None and row_num == header_row:
                    continue

                record: Dict[str, Any] = {"_row": row_num}
                for col, cd in sorted(cells.items()):
                    key = header_map.get(col, num_to_col(col))
                    record[key] = cell_record_value(cd)
                    if cd.get("formula") is not None or cd.get("formula_attributes") is not None:
                        formula_count += 1
                        attrs = cd.get("formula_attributes") or {}
                        shared_index = attrs.get("si") if attrs.get("t") == "shared" else None
                        if shared_index is not None and cd.get("formula"):
                            shared_formula_masters[str(shared_index)] = {
                                "master_cell": cd["address"],
                                "master_formula": cd["formula"],
                                "range": attrs.get("ref"),
                            }
                        formula_item = {
                            "address": cd["address"],
                            "row": row_num,
                            "column": col,
                            "column_letter": num_to_col(col),
                            "formula": cd.get("formula"),
                            "formula_attributes": attrs,
                            "cached_value": cd.get("value"),
                            "raw_cached_value": cd.get("raw_value"),
                            "style_index": cd.get("style_index"),
                        }
                        formulas_writer.write(formula_item, row_num=row_num)
                records_writer.write(record, row_num=row_num)
                if len(preview) < 50:
                    preview.append(record)

            record_parts = records_writer.close()
            formula_parts = formulas_writer.close()

            # Add shared master references to formula metadata after streaming.
            formula_summary = {
                "formula_count": formula_count,
                "shared_formula_masters": shared_formula_masters,
                "parts": formula_parts,
                "note": "For shared formulas Excel may store the full formula only in the master cell; dependent cells reference formula_attributes.si.",
            }
            json_dump(sheet_dir / "formula_index.json", formula_summary)
            json_dump(sheet_dir / "preview.json", {
                "sheet": name,
                "header_row": header_row,
                "preamble": preamble,
                "first_records": preview,
            })

            sheet_meta = {
                "sheet_index": idx,
                "sheet_name": name,
                "sheet_state": sheet.attrib.get("state", "visible"),
                "sheet_id": sheet.attrib.get("sheetId"),
                "source_part": sheet_part,
                **static,
                "header_row": header_row,
                "columns_normalized": header_inventory,
                "preamble": preamble,
                "row_properties": row_properties,
                "nonempty_row_count": nonempty_rows,
                "cell_count": cell_count,
                "formula_count": formula_count,
                "record_parts": record_parts,
                "formula_parts": formula_parts,
            }
            json_dump(sheet_dir / "sheet.meta.json", sheet_meta)
            sheets_meta.append({
                "index": idx,
                "name": name,
                "state": sheet.attrib.get("state", "visible"),
                "dimension": static.get("dimension"),
                "header_row": header_row,
                "nonempty_rows": nonempty_rows,
                "cells": cell_count,
                "formulas": formula_count,
                "path": f"sheets/{sheet_slug}/sheet.meta.json",
            })

        json_dump(target / "styles.json", {
            "cell_styles": styles,
            "custom_number_formats": {str(k): v for k, v in custom_formats.items()},
        })

        has_vba = "xl/vbaProject.bin" in names
        vba_hash = None
        if has_vba:
            vba_bytes = z.read("xl/vbaProject.bin")
            vba_hash = hashlib.sha256(vba_bytes).hexdigest()
            (target / "vba").mkdir(exist_ok=True)
            (target / "vba" / "vbaProject.bin.sha256").write_text(vba_hash + "\n", encoding="ascii")

        wb_meta = {
            "format_version": "1.0",
            "source_file": source.name,
            "source_size_bytes": source.stat().st_size,
            "source_sha256": source_hash,
            "source_type": source.suffix.lower().lstrip("."),
            "generated_at_utc": dt.datetime.now(dt.timezone.utc).isoformat(),
            "date_system": "1904" if date1904 else "1900",
            "active_sheet_index_zero_based": active_tab,
            "calculation_properties": dict(calc_pr.attrib) if calc_pr is not None else {},
            "defined_names": defined_names,
            "document_properties": read_doc_properties(z),
            "has_vba": has_vba,
            "vba_project_sha256": vba_hash,
            "sheets": sheets_meta,
            "limitations": [
                "Cached formula results are read from the source file and are not recalculated during export.",
                "Formatting is summarized; drawings, buttons and embedded objects are inventoried but not reconstructed in JSON.",
                "The original XLSX/XLSM remains the authoritative file, especially for VBA, styles and interactive controls.",
            ],
        }
        json_dump(target / "workbook.json", wb_meta)

    return {
        "source": source.name,
        "slug": slug,
        "type": source.suffix.lower().lstrip("."),
        "sha256": source_hash,
        "size_bytes": source.stat().st_size,
        "output": slug,
        "sheets": len(sheets_meta),
        "has_vba": has_vba,
    }


def extract_vba_with_libreoffice(source: Path, workbook_out_dir: Path) -> Dict[str, Any]:
    libreoffice = shutil.which("libreoffice") or shutil.which("soffice")
    if not libreoffice:
        return {"status": "skipped", "reason": "LibreOffice not found"}
    with tempfile.TemporaryDirectory(prefix="vba_extract_") as tmp:
        tmp_path = Path(tmp)
        proc = subprocess.run(
            [libreoffice, "--headless", "--convert-to", "ods", "--outdir", str(tmp_path), str(source)],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            timeout=180,
        )
        ods_files = list(tmp_path.glob("*.ods"))
        if proc.returncode != 0 or not ods_files:
            return {"status": "failed", "output": proc.stdout.strip()}
        ods = ods_files[0]
        modules = []
        vba_dir = workbook_out_dir / "vba" / "source"
        vba_dir.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(ods) as z:
            module_parts = sorted(n for n in z.namelist() if n.startswith("Basic/VBAProject/") and n.endswith(".xml") and not n.endswith("script-lb.xml"))
            for part in module_parts:
                xml = z.read(part).decode("utf-8", errors="replace")
                start = xml.find(">") + 1
                end = xml.rfind("</script:module>")
                source_text = html.unescape(xml[start:end]) if end > start else html.unescape(xml)
                name = Path(part).stem
                module_type_match = re.search(r"Attribute VBA_ModuleType=([A-Za-z]+)", source_text)
                module_type = module_type_match.group(1) if module_type_match else "VBAModule"
                ext = ".bas" if module_type == "VBAModule" else ".cls"
                out_file = vba_dir / f"{name}{ext}"
                out_file.write_text(source_text.strip() + "\n", encoding="utf-8", newline="\n")
                procedures = []
                for m in re.finditer(r"(?im)^\s*(Public\s+|Private\s+|Friend\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_][A-Za-z0-9_]*)", source_text):
                    procedures.append({"kind": m.group(2), "name": m.group(3), "visibility": (m.group(1) or "").strip() or "default"})
                modules.append({
                    "module": name,
                    "module_type": module_type,
                    "file": f"source/{out_file.name}",
                    "lines": source_text.count("\n") + 1,
                    "sha256": sha256_file(out_file),
                    "procedures": procedures,
                })
        inventory = {
            "status": "ok",
            "method": "LibreOffice XLSM-to-ODS VBA import, then XML source extraction",
            "warning": "This text export is for review and version control. The original XLSM remains authoritative and should be used for execution.",
            "libreoffice_output": proc.stdout.strip(),
            "modules": modules,
        }
        json_dump(workbook_out_dir / "vba" / "inventory.json", inventory)
        return inventory


def flatten_outline(items: Any) -> List[Dict[str, Any]]:
    out: List[Dict[str, Any]] = []
    def walk(seq: Any, level: int = 0) -> None:
        if not isinstance(seq, list):
            return
        for item in seq:
            if isinstance(item, list):
                walk(item, level + 1)
            else:
                title = getattr(item, "title", None) or str(item)
                page_number = None
                try:
                    page_number = item.page_number  # rarely available
                except Exception:
                    pass
                out.append({"level": level, "title": title, "page": page_number})
    walk(items)
    return out


def export_pdf(source: Path, out_dir: Path) -> Dict[str, Any]:
    import fitz  # PyMuPDF

    slug = slugify(source.stem)
    target = out_dir / slug
    pages_dir = target / "pages"
    pages_dir.mkdir(parents=True, exist_ok=True)
    doc = fitz.open(source)
    pages_index = []
    detected_sections = []
    section_re = re.compile(r"^\s*((?:\d+\.)+\s+.+|(?:\d+(?:\.\d+)+\.)\s+.+|Приложение.*)$", re.IGNORECASE)
    for i, page in enumerate(doc, start=1):
        text = page.get_text("text", sort=True)
        blocks_raw = page.get_text("blocks", sort=True)
        blocks = []
        for b in blocks_raw:
            x0, y0, x1, y1, block_text, block_no, block_type = b[:7]
            if not str(block_text).strip():
                continue
            blocks.append({
                "bbox": [round(float(x0), 2), round(float(y0), 2), round(float(x1), 2), round(float(y1), 2)],
                "text": block_text,
                "block_no": int(block_no),
                "block_type": int(block_type),
            })
        for line in text.splitlines():
            if section_re.match(line) and len(line.strip()) < 180:
                detected_sections.append({"page": i, "title": line.strip()})
        page_data = {
            "page": i,
            "width_points": round(page.rect.width, 2),
            "height_points": round(page.rect.height, 2),
            "rotation": page.rotation,
            "text": text,
            "blocks": blocks,
            "links": page.get_links(),
        }
        page_file = pages_dir / f"page-{i:03d}.json"
        json_dump(page_file, page_data)
        pages_index.append({
            "page": i,
            "file": f"pages/{page_file.name}",
            "text_chars": len(text),
            "blocks": len(blocks),
            "sha256": sha256_file(page_file),
        })
    meta = doc.metadata or {}
    toc = [{"level": level, "title": title, "page": page_num} for level, title, page_num, *rest in doc.get_toc(simple=True)]
    document = {
        "format_version": "1.0",
        "source_file": source.name,
        "source_size_bytes": source.stat().st_size,
        "source_sha256": sha256_file(source),
        "source_type": "pdf",
        "generated_at_utc": dt.datetime.now(dt.timezone.utc).isoformat(),
        "page_count": doc.page_count,
        "metadata": meta,
        "table_of_contents": toc,
        "detected_sections": detected_sections,
        "pages": pages_index,
        "limitations": [
            "PDF text and layout blocks are extracted, but complex tables are not guaranteed to be reconstructed as semantic rows and columns.",
            "The original PDF remains authoritative for visual layout, diagrams and pagination.",
        ],
    }
    json_dump(target / "document.json", document)
    doc.close()
    return {
        "source": source.name,
        "slug": slug,
        "type": "pdf",
        "sha256": document["source_sha256"],
        "size_bytes": source.stat().st_size,
        "output": slug,
        "pages": document["page_count"],
    }


def write_readme(root: Path, manifest: Dict[str, Any]) -> None:
    readme = f"""# Repository-friendly export of WB files

This package contains generated, text-readable snapshots of {len(manifest['files'])} source files.

## Recommended source-of-truth model

1. Keep the original `.xlsx`, `.xlsm` and `.pdf` files under `data/source/`.
2. Commit the generated JSON/JSONL under `data/json/`.
3. Treat `manifest.json` and SHA-256 hashes as the link between originals and generated snapshots.
4. Regenerate snapshots with `scripts/export_office_to_repo_json.py` whenever a source file changes.
5. Never replace the original `.xlsm` with JSON: VBA, buttons, drawings, styles and Excel behavior are not fully round-trippable.

## Why JSONL instead of one huge JSON file

Large WB reports contain hundreds of thousands of cells. One monolithic JSON file is difficult to diff, review and stream. Therefore each sheet is exported as sharded `records.part-XXXXX.jsonl` files, one logical row per line. Small `preview.json` files make quick review easy in GitHub.

## Workbook export layout

- `workbook.json` - workbook-level metadata, source checksum, sheet index, defined names and calculation properties.
- `styles.json` - style indexes and number formats used to decode dates.
- `sheets/<sheet>/sheet.meta.json` - dimensions, merged ranges, header mapping, tables, hyperlinks, row/column properties and shard index.
- `sheets/<sheet>/records.part-*.jsonl` - values using readable header names and original row number `_row`.
- `sheets/<sheet>/formulas.part-*.jsonl` - formulas, cached results and shared-formula attributes.
- `sheets/<sheet>/preview.json` - preamble plus first 50 records.
- `vba/source/*` - readable VBA modules extracted from the XLSM through LibreOffice, when available.

## PDF export layout

- `document.json` - metadata, section index and page index.
- `pages/page-XXX.json` - page text, positioned text blocks and links.

## Important limitations

- Formula values are cached values stored in the source workbook; this export does not recalculate Excel formulas.
- JSON is an audit/read layer, not a replacement for the working Excel model.
- VBA text extracted through LibreOffice is intended for review and diffs; run macros only from the original XLSM in Microsoft Excel.
- PDF tables are preserved as text/layout blocks, not guaranteed relational tables.

## Suggested repository layout

```text
data/
  source/                 # original XLSX/XLSM/PDF, authoritative
  json/                   # contents of this package's data/json
scripts/
  export_office_to_repo_json.py
manifest.json
```

## Validation

Every source and generated shard has a SHA-256 checksum. `validation_report.json` records parse and count checks performed after export.
"""
    (root / "README.md").write_text(readme, encoding="utf-8", newline="\n")


def validate_export(root: Path, manifest: Dict[str, Any]) -> Dict[str, Any]:
    errors = []
    checked_json = 0
    checked_jsonl = 0
    jsonl_rows = 0
    for path in root.rglob("*.json"):
        try:
            with path.open("r", encoding="utf-8") as f:
                json.load(f)
            checked_json += 1
        except Exception as e:
            errors.append({"file": str(path.relative_to(root)), "error": str(e)})
    for path in root.rglob("*.jsonl"):
        try:
            with path.open("r", encoding="utf-8") as f:
                for line_no, line in enumerate(f, start=1):
                    if line.strip():
                        json.loads(line)
                        jsonl_rows += 1
            checked_jsonl += 1
        except Exception as e:
            errors.append({"file": str(path.relative_to(root)), "error": f"line {line_no}: {e}"})
    return {
        "status": "ok" if not errors else "failed",
        "checked_json_files": checked_json,
        "checked_jsonl_files": checked_jsonl,
        "checked_jsonl_rows": jsonl_rows,
        "errors": errors,
        "validated_at_utc": dt.datetime.now(dt.timezone.utc).isoformat(),
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("inputs", nargs="+", help="XLSX/XLSM/PDF files")
    parser.add_argument("--output", required=True, help="Output package directory")
    args = parser.parse_args()

    root = Path(args.output).resolve()
    data_json = root / "data" / "json"
    scripts_dir = root / "scripts"
    data_json.mkdir(parents=True, exist_ok=True)
    scripts_dir.mkdir(parents=True, exist_ok=True)

    entries = []
    for raw in args.inputs:
        source = Path(raw).resolve()
        suffix = source.suffix.lower()
        if suffix in (".xlsx", ".xlsm"):
            entry = export_workbook(source, data_json)
            if suffix == ".xlsm":
                vba_result = extract_vba_with_libreoffice(source, data_json / entry["output"])
                entry["vba_extraction_status"] = vba_result.get("status")
                entry["vba_modules"] = len(vba_result.get("modules", []))
            entries.append(entry)
        elif suffix == ".pdf":
            entries.append(export_pdf(source, data_json))
        else:
            entries.append({"source": source.name, "type": suffix.lstrip("."), "status": "unsupported"})

    manifest = {
        "format_version": "1.0",
        "generated_at_utc": dt.datetime.now(dt.timezone.utc).isoformat(),
        "generator": "scripts/export_office_to_repo_json.py",
        "files": entries,
    }
    json_dump(root / "manifest.json", manifest)
    write_readme(root, manifest)
    shutil.copy2(Path(__file__), scripts_dir / "export_office_to_repo_json.py")
    (root / "requirements.txt").write_text("PyMuPDF>=1.24\n", encoding="utf-8")
    (root / ".gitattributes").write_text("*.xlsx binary\n*.xlsm binary\n*.pdf binary\n*.json text eol=lf\n*.jsonl text eol=lf\n*.bas text eol=crlf\n*.cls text eol=crlf\n", encoding="utf-8")
    (root / ".gitignore").write_text("__pycache__/\n*.pyc\n.tmp/\n", encoding="utf-8")

    validation = validate_export(root, manifest)
    json_dump(root / "validation_report.json", validation)
    if validation["status"] != "ok":
        print(json.dumps(validation, ensure_ascii=False, indent=2))
        return 2
    print(json.dumps({"manifest": manifest, "validation": validation}, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
