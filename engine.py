# -*- coding: utf-8 -*-
"""Pivot Topos vs BC, compare Total (qty) / Unit, export result workbook."""

from __future__ import annotations

import re
import shutil
import unicodedata
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Callable, Iterable

import openpyxl
from openpyxl.styles import Border, Font, Side
from openpyxl.utils import get_column_letter

ENGINE_VERSION = "2026-09-25.6"
COLUMN_LAYOUT = (
    "Ma san pham | Ten san pham | Topos SL | Topos Unit | BC SL | BC Unit | "
    "Chenh lech SL | Status | Ma bill Topos lech"
)

QTY_TOLERANCE = 1e-6
OUTPUT_SUFFIX = " - PivotCompare v5.xlsx"
RESULT_SHEET = "Ket qua lech ma SP"
TOPOS_SHEET = "Topos"
BC_SHEET = "BC"
SUMMARY_HEADERS = (
    "TransferFromCode",
    "TransferToCode",
    "Date",
    "Ma san pham",
    "Ten san pham",
    "Topos SL",
    "Topos Unit",
    "BC SL",
    "BC Unit",
    "Chenh lech SL",
    "Status",
    "Ma bill Topos lech",
)
WEB_PREVIEW_ROWS = 300

_THIN = Side(style="thin", color="808080")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_HEADER_FONT = Font(bold=True)


class CompareError(Exception):
    """User-facing validation or processing error."""


# normalized header -> list of acceptable norm keys (first match wins)
TOPOS_ALIASES: dict[str, tuple[str, ...]] = {
    "transferfromcode": ("transferfromcode", "transferfrom", "fromcode"),
    "transfertocode": ("transfertocode", "transferto", "tocode"),
    "date": ("ngay", "date", "postingdate"),
    "product": ("masanpham", "masp", "productcode", "itemno"),
    "qty": ("soluong", "quantity", "qty"),
    "unit": ("donvi", "unit", "unitofmeasure", "uom"),
    "bill": ("mabill", "bill", "billno", "documentno"),
    "name": ("tensanpham", "description", "itemname"),
}

BC_ALIASES: dict[str, tuple[str, ...]] = {
    "transferfromcode": ("transferfromcode", "transferfrom", "fromcode"),
    "transfertocode": ("transfertocode", "transferto", "tocode"),
    "date": ("postingdate", "ngay", "date"),
    "itemno": ("itemno", "no", "itemnumber"),
    "itemref": ("itemreferenceno", "itemreference", "itemrefno", "referenceno"),
    "qty": ("quantity", "soluong", "qty"),
    "unit": ("unitofmeasure", "unitofmeasurecode", "donvi", "unit", "uom"),
    "extdoc": ("externaldocumentno", "externaldocno", "extdocno"),
    "name": ("description", "description2", "tensanpham"),
}

def normalize_header(value) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = text.replace("\u0111", "d").replace("\u0110", "d")
    text = unicodedata.normalize("NFD", text)
    text = "".join(c for c in text if unicodedata.category(c) != "Mn")
    return re.sub(r"[^a-z0-9]+", "", text)


def normalize_text(value) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = text.replace("\u0111", "d").replace("\u0110", "d")
    text = unicodedata.normalize("NFD", text)
    text = "".join(c for c in text if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", text)


def normalize_unit(value) -> str:
    u = normalize_text(value).upper()
    aliases = {
        "KG": "KG",
        "KILOGRAM": "KG",
        "KILOGRAMS": "KG",
        "G": "G",
        "VI": "VI",
        "VỈ": "VI",
        "CAI": "CAI",
        "CHIEC": "CAI",
        "CHIẾC": "CAI",
        "THUNG": "THUNG",
        "THÙNG": "THUNG",
        "HOP": "HOP",
        "HỘP": "HOP",
        "LIT": "LIT",
        "L": "LIT",
    }
    return aliases.get(u, u)


def parse_date(value) -> date | None:
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value).strip()
    if text.upper() in ("NULL", "NONE", "NAN"):
        return None
    if "T" in text:
        text = text.split("T", 1)[0]
    if " " in text:
        text = text.split(" ", 1)[0]
    for fmt in (
        "%d/%m/%Y",
        "%d-%m-%Y",
        "%Y-%m-%d",
        "%m/%d/%Y",
        "%Y/%m/%d",
        "%d.%m.%Y",
    ):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    m = re.match(r"^(\d{4})(\d{2})(\d{2})$", text)
    if m:
        return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return None


def to_float(value) -> float:
    if value is None or value == "":
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return 0.0
    try:
        return float(text)
    except ValueError:
        return 0.0


def resolve_columns(
    headers: Iterable,
    alias_map: dict[str, tuple[str, ...]],
    optional: frozenset[str] | None = None,
) -> dict[str, int]:
    norm_index = {normalize_header(h): i for i, h in enumerate(headers)}
    resolved: dict[str, int] = {}
    optional = optional or frozenset()
    for role, keys in alias_map.items():
        for key in keys:
            if key in norm_index:
                resolved[role] = norm_index[key]
                break
        if role not in resolved and role not in optional:
            raise CompareError(
                f"Thieu cot '{role}'. Header co san: "
                + ", ".join(str(h) for h in headers if h is not None)
            )
    return resolved


BC_OPTIONAL = frozenset({"itemref"})


def _display_unit(raw) -> str:
    if raw is None:
        return ""
    return str(raw).strip()


def extract_bill_from_extdoc(value) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if not text:
        return ""
    if "-" in text:
        return text.split("-", 1)[0].strip()
    return text


@dataclass
class Agg:
    qty: float = 0.0
    unit: str = ""
    units_seen: set[str] = field(default_factory=set)

    def add(self, qty: float, unit: str) -> None:
        self.qty += qty
        u = _display_unit(unit)
        if u:
            self.units_seen.add(u)
            if not self.unit:
                self.unit = u


def _agg_display_unit(agg: Agg) -> str:
    if agg.unit:
        return _display_unit(agg.unit)
    if agg.units_seen:
        return sorted(agg.units_seen)[0]
    return ""


def _product_unit_map_from_topos(rows: list[tuple], cols: dict[str, int]) -> dict[str, str]:
    counts: dict[str, Counter[str]] = defaultdict(Counter)
    for row in rows:
        product = str(row[cols["product"]]).strip()
        u = _display_unit(row[cols["unit"]])
        if product and u:
            counts[product][u] += 1
    return {p: c.most_common(1)[0][0] for p, c in counts.items()}


def _product_unit_map_from_bc(
    rows: list[tuple],
    cols: dict[str, int],
    item_to_topos: dict[str, str],
) -> dict[str, str]:
    counts: dict[str, Counter[str]] = defaultdict(Counter)
    for row in rows:
        item = str(row[cols["itemno"]]).strip()
        product = item_to_topos.get(item, item)
        u = _display_unit(row[cols["unit"]])
        if product and u:
            counts[product][u] += 1
    return {p: c.most_common(1)[0][0] for p, c in counts.items()}


def _units_from_bills(
    pivot_key: tuple,
    bill_pivot: dict[tuple, Agg],
) -> str:
    units: list[str] = []
    for (from_c, to_c, d, product, _bill), agg in bill_pivot.items():
        if (from_c, to_c, d, product) != pivot_key:
            continue
        u = _agg_display_unit(agg)
        if u:
            units.append(u)
    if not units:
        return ""
    return Counter(units).most_common(1)[0][0]


def _resolve_pair_units(
    key: tuple,
    a: Agg,
    b: Agg,
    tp_bills: dict[tuple, Agg],
    bc_bills: dict[tuple, Agg],
    tp_product_units: dict[str, str],
    bc_product_units: dict[str, str],
) -> tuple[str, str]:
    tp_u = _agg_display_unit(a)
    bc_u = _agg_display_unit(b)
    if not tp_u:
        tp_u = _units_from_bills(key, tp_bills)
    if not bc_u:
        bc_u = _units_from_bills(key, bc_bills)
    product = str(key[3])
    if not tp_u and product:
        tp_u = tp_product_units.get(product, "")
    if not bc_u and product:
        bc_u = bc_product_units.get(product, "")
    # Phia khong phat sinh SL tren khoa pivot: hien don vi tham chieu tu phia con lai.
    if not tp_u and bc_u and a.qty == 0:
        tp_u = bc_u
    if not bc_u and tp_u and b.qty == 0:
        bc_u = tp_u
    if not tp_u and bc_u:
        tp_u = bc_u
    if not bc_u and tp_u:
        bc_u = tp_u
    if not tp_u and not bc_u:
        tp_u = bc_u = "-"
    return tp_u, bc_u


def _qty_mismatch(a: float, b: float) -> bool:
    return abs(a - b) > QTY_TOLERANCE


def _unit_mismatch(a: str, b: str) -> bool:
    ua, ub = normalize_unit(a), normalize_unit(b)
    if not ua and not ub:
        return False
    if not ua or not ub:
        return True
    return ua != ub


def _mismatch_status(tp_qty: float, tp_unit: str, bc_qty: float, bc_unit: str) -> str:
    qty_bad = _qty_mismatch(tp_qty, bc_qty)
    unit_bad = _unit_mismatch(tp_unit, bc_unit)
    if qty_bad and unit_bad:
        return "Lech ca SL va don vi"
    if qty_bad:
        return "Lech so luong"
    if unit_bad:
        return "Lech don vi"
    return ""


def _build_product_names_topos(rows: list[tuple], cols: dict[str, int]) -> dict[str, str]:
    out: dict[str, str] = {}
    for row in rows:
        code = str(row[cols["product"]]).strip() if row[cols["product"]] is not None else ""
        if not code or code in out:
            continue
        raw = row[cols["name"]]
        if raw is not None and str(raw).strip():
            out[code] = str(raw).strip()
    return out


def _build_product_names_bc(
    rows: list[tuple],
    cols: dict[str, int],
    item_to_topos: dict[str, str],
) -> dict[str, str]:
    out: dict[str, str] = {}
    for row in rows:
        item = str(row[cols["itemno"]]).strip() if row[cols["itemno"]] is not None else ""
        if not item:
            continue
        code = item_to_topos.get(item, item)
        if code in out:
            continue
        raw = row[cols["name"]]
        if raw is not None and str(raw).strip():
            out[code] = str(raw).strip()
    return out


def _merge_product_names(topos: dict[str, str], bc: dict[str, str]) -> dict[str, str]:
    merged = dict(topos)
    for code, name in bc.items():
        merged.setdefault(code, name)
    return merged


def _build_name_to_topos(rows: list[tuple], name_idx: int, product_idx: int) -> dict[str, str]:
    out: dict[str, str] = {}
    for row in rows:
        name = normalize_text(row[name_idx])
        code = str(row[product_idx]).strip() if row[product_idx] is not None else ""
        if name and code:
            out[name] = code
    return out


def _build_item_to_topos_from_rows(
    bc_rows: list[tuple],
    bc_cols: dict[str, int],
    name_to_topos: dict[str, str],
) -> dict[str, str]:
    item_to_topos: dict[str, str] = {}
    ref_idx = bc_cols.get("itemref")
    name_idx = bc_cols["name"]
    item_idx = bc_cols["itemno"]

    for row in bc_rows:
        item = str(row[item_idx]).strip() if row[item_idx] is not None else ""
        if not item:
            continue
        if item in item_to_topos:
            continue
        if ref_idx is not None:
            ref = row[ref_idx]
            if ref is not None and str(ref).strip():
                item_to_topos[item] = str(ref).strip()
                continue
        name = normalize_text(row[name_idx])
        if name in name_to_topos:
            item_to_topos[item] = name_to_topos[name]

    return item_to_topos


def _read_sheet_rows(wb, sheet_name: str) -> list[tuple]:
    if sheet_name not in wb.sheetnames:
        raise CompareError(f"Khong tim thay sheet '{sheet_name}'. Co: {', '.join(wb.sheetnames)}")
    ws = wb[sheet_name]
    return list(ws.iter_rows(values_only=True))


def pivot_topos(rows: list[tuple], log: Callable[[str], None]) -> tuple[dict, dict, list[tuple], dict[str, int]]:
    if not rows:
        raise CompareError("Sheet Topos trong.")
    headers = rows[0]
    cols = resolve_columns(headers, TOPOS_ALIASES)
    data = rows[1:]
    pivot: dict[tuple, Agg] = defaultdict(Agg)
    bill_pivot: dict[tuple, Agg] = defaultdict(Agg)

    skipped_date = 0
    for row in data:
        d = parse_date(row[cols["date"]])
        if d is None:
            skipped_date += 1
            continue
        product = str(row[cols["product"]]).strip()
        from_c = str(row[cols["transferfromcode"]]).strip()
        to_c = str(row[cols["transfertocode"]]).strip()
        qty = to_float(row[cols["qty"]])
        unit = str(row[cols["unit"]] or "")
        bill = str(row[cols["bill"]] or "").strip()
        key = (from_c, to_c, d, product)
        pivot[key].add(qty, unit)
        if bill:
            bill_key = (from_c, to_c, d, product, bill)
            bill_pivot[bill_key].add(qty, unit)

    if skipped_date:
        log(f"Topos: bo qua {skipped_date} dong khong doc duoc ngay.")
    log(f"Topos pivot: {len(pivot)} khoa.")
    return pivot, bill_pivot, data, cols


def pivot_bc(
    rows: list[tuple],
    item_to_topos: dict[str, str],
    log: Callable[[str], None],
) -> tuple[dict, dict, dict[str, int]]:
    if not rows:
        raise CompareError("Sheet BC trong.")
    headers = rows[0]
    cols = resolve_columns(headers, BC_ALIASES, BC_OPTIONAL)
    data = rows[1:]
    pivot: dict[tuple, Agg] = defaultdict(Agg)
    bill_pivot: dict[tuple, Agg] = defaultdict(Agg)
    skipped_date = 0
    unmapped = 0

    for row in data:
        d = parse_date(row[cols["date"]])
        if d is None:
            skipped_date += 1
            continue
        item = str(row[cols["itemno"]]).strip()
        product = item_to_topos.get(item, item)
        if item and product == item and item_to_topos and item not in item_to_topos:
            unmapped += 1
        from_c = str(row[cols["transferfromcode"]]).strip()
        to_c = str(row[cols["transfertocode"]]).strip()
        qty = to_float(row[cols["qty"]])
        unit = str(row[cols["unit"]] or "")
        key = (from_c, to_c, d, product)
        pivot[key].add(qty, unit)
        bill = extract_bill_from_extdoc(row[cols["extdoc"]])
        if bill:
            bill_key = (from_c, to_c, d, product, bill)
            bill_pivot[bill_key].add(qty, unit)

    if skipped_date:
        log(f"BC: bo qua {skipped_date} dong khong doc duoc PostingDate.")
    if unmapped:
        log(f"BC: {unmapped} dong ItemNo chua map sang Ma san pham Topos (dung ItemNo de so).")
    log(f"BC pivot: {len(pivot)} khoa.")
    return pivot, bill_pivot, cols


def _pivot_row(
    key: tuple,
    tp: dict[tuple, Agg],
    bc: dict[tuple, Agg],
    tp_bills: dict[tuple, Agg],
    bc_bills: dict[tuple, Agg],
    tp_product_units: dict[str, str],
    bc_product_units: dict[str, str],
    product_names: dict[str, str],
) -> tuple:
    a = tp.get(key, Agg())
    b = bc.get(key, Agg())
    tp_u, bc_u = _resolve_pair_units(
        key, a, b, tp_bills, bc_bills, tp_product_units, bc_product_units
    )
    status = _mismatch_status(a.qty, tp_u, b.qty, bc_u) or "Khop"
    product_code = str(key[3])
    product_name = product_names.get(product_code, "")
    return (
        key[0],
        key[1],
        key[2],
        product_code,
        product_name,
        a.qty,
        tp_u,
        b.qty,
        bc_u,
        a.qty - b.qty,
        status,
    )


def build_lech_bills_by_pivot_key(
    tp_bills: dict[tuple, Agg],
    bc_bills: dict[tuple, Agg],
) -> dict[tuple, str]:
    """Map pivot key (from, to, date, product) -> Topos bill codes lech SL/Unit."""
    bills: dict[tuple, set[str]] = defaultdict(set)
    for key in set(tp_bills) | set(bc_bills):
        from_c, to_c, d, product, bill = key
        if not bill:
            continue
        a = tp_bills.get(key, Agg())
        b = bc_bills.get(key, Agg())
        tp_u = _agg_display_unit(a)
        bc_u = _agg_display_unit(b)
        if _qty_mismatch(a.qty, b.qty) or _unit_mismatch(tp_u, bc_u):
            bills[(from_c, to_c, d, product)].add(bill)
    return {k: "; ".join(sorted(v)) for k, v in bills.items()}


def build_mismatch_pivot_rows(
    tp: dict[tuple, Agg],
    bc: dict[tuple, Agg],
    tp_bills: dict[tuple, Agg],
    bc_bills: dict[tuple, Agg],
    lech_bills: dict[tuple, str],
    tp_product_units: dict[str, str],
    bc_product_units: dict[str, str],
    product_names: dict[str, str],
) -> list[tuple]:
    keys = set(tp) | set(bc)
    rows: list[tuple] = []
    for key in sorted(keys, key=lambda k: (k[2], k[0], k[1], k[3])):
        base = _pivot_row(
            key,
            tp,
            bc,
            tp_bills,
            bc_bills,
            tp_product_units,
            bc_product_units,
            product_names,
        )
        if base[10] == "Khop":
            continue
        bill_str = lech_bills.get(key, "")
        rows.append(base + (bill_str,))
    return rows


def _format_date(d: date) -> str:
    return d.strftime("%d/%m/%Y")


def _preview_product_rows(rows: list[tuple], limit: int = WEB_PREVIEW_ROWS) -> list[dict]:
    out: list[dict] = []
    for row in rows[:limit]:
        out.append(
            {
                "TransferFromCode": row[0],
                "TransferToCode": row[1],
                "Date": _format_date(row[2]) if isinstance(row[2], date) else row[2],
                "Ma san pham": row[3],
                "Ten san pham": row[4],
                "Topos SL": row[5],
                "Topos Unit": row[6],
                "BC SL": row[7],
                "BC Unit": row[8],
                "Chenh lech SL": row[9],
                "Status": row[10],
                "Ma bill Topos lech": row[11],
            }
        )
    return out


def _autofit_columns(ws, max_col: int, max_row: int) -> None:
    for col in range(1, max_col + 1):
        letter = get_column_letter(col)
        width = 12
        for row in range(1, max_row + 1):
            val = ws.cell(row, col).value
            if val is None:
                continue
            width = max(width, min(48, len(str(val)) + 2))
        ws.column_dimensions[letter].width = width


def _write_data_table(
    ws,
    start_row: int,
    headers: tuple[str, ...],
    rows: list[tuple],
    date_cols: frozenset[int],
    unit_cols: frozenset[int],
    qty_cols: frozenset[int],
) -> int:
    for c, h in enumerate(headers, 1):
        cell = ws.cell(start_row, c, h)
        cell.font = _HEADER_FONT
        cell.border = _BORDER
    for r_idx, row in enumerate(rows, start_row + 1):
        for c_idx, val in enumerate(row, 1):
            if c_idx in date_cols and isinstance(val, date):
                out = _format_date(val)
            elif c_idx in unit_cols:
                out = _display_unit(val) or ""
            else:
                out = val
            cell = ws.cell(r_idx, c_idx, out)
            cell.border = _BORDER
            if c_idx in qty_cols:
                cell.number_format = "0.########"
    return start_row + len(rows)


def write_result_sheet(wb, mismatch_rows: list[tuple]) -> None:
    for name in list(wb.sheetnames):
        if name == RESULT_SHEET or name.startswith("Ket qua"):
            del wb[name]
    ws = wb.create_sheet(RESULT_SHEET)

    ws.cell(1, 1, f"PivotCompare engine {ENGINE_VERSION}").font = _HEADER_FONT
    ws.cell(1, 2, COLUMN_LAYOUT)

    ws.cell(2, 1, "Pivot Topos vs BC (chi dong lech)").font = _HEADER_FONT
    last_row = _write_data_table(
        ws,
        3,
        SUMMARY_HEADERS,
        mismatch_rows,
        date_cols=frozenset({3}),
        unit_cols=frozenset({7, 9}),
        qty_cols=frozenset({6, 8, 10}),
    )

    _autofit_columns(ws, len(SUMMARY_HEADERS), last_row)
    ws.freeze_panes = "A4"


def compare_workbook(
    source: Path,
    output: Path,
    log: Callable[[str], None] | None = None,
) -> dict:
    _log = log or (lambda _m: None)
    source = Path(source)
    output = Path(output)
    if not source.is_file():
        raise CompareError(f"Khong tim thay file: {source}")

    wb = openpyxl.load_workbook(source, read_only=True, data_only=True)
    topos_rows = _read_sheet_rows(wb, TOPOS_SHEET)
    bc_rows = _read_sheet_rows(wb, BC_SHEET)
    wb.close()

    tp_pivot, tp_bill_pivot, topos_data, tp_cols = pivot_topos(topos_rows, _log)
    name_to_topos = _build_name_to_topos(topos_data, tp_cols["name"], tp_cols["product"])
    bc_header_cols = resolve_columns(bc_rows[0], BC_ALIASES, BC_OPTIONAL)
    item_to_topos = _build_item_to_topos_from_rows(
        bc_rows[1:], bc_header_cols, name_to_topos
    )
    _log(f"Anh xa ItemNo -> Ma SP: {len(item_to_topos)} ma.")
    tp_product_units = _product_unit_map_from_topos(topos_data, tp_cols)
    bc_product_units = _product_unit_map_from_bc(bc_rows[1:], bc_header_cols, item_to_topos)
    product_names = _merge_product_names(
        _build_product_names_topos(topos_data, tp_cols),
        _build_product_names_bc(bc_rows[1:], bc_header_cols, item_to_topos),
    )

    bc_pivot, bc_bill_pivot, _ = pivot_bc(bc_rows, item_to_topos, _log)
    lech_bills = build_lech_bills_by_pivot_key(tp_bill_pivot, bc_bill_pivot)
    mismatches = build_mismatch_pivot_rows(
        tp_pivot,
        bc_pivot,
        tp_bill_pivot,
        bc_bill_pivot,
        lech_bills,
        tp_product_units,
        bc_product_units,
        product_names,
    )
    _log(f"Dong pivot lech (khong gom dong Khop): {len(mismatches)}.")
    with_bill = sum(1 for m in mismatches if m[11])
    _log(f"Dong lech co ma bill Topos: {with_bill}.")

    shutil.copy2(source, output)
    out_wb = openpyxl.load_workbook(output)
    write_result_sheet(out_wb, mismatches)
    out_wb.save(output)
    out_wb.close()

    return {
        "OutputPath": str(output),
        "EngineVersion": ENGINE_VERSION,
        "ResultHeaders": list(SUMMARY_HEADERS),
        "MismatchRows": len(mismatches),
        "MismatchWithBill": with_bill,
        "PreviewProducts": _preview_product_rows(mismatches),
    }
