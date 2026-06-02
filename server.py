from datetime import datetime, timezone, timedelta
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from io import BytesIO
import base64
import json
import mimetypes
import os
import re
import traceback
import xml.etree.ElementTree as ET
import zipfile


HOST = os.environ.get("HOST", "0.0.0.0")
PORT = int(os.environ.get("PORT", "8788"))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PUBLIC_DIR = os.path.join(BASE_DIR, "public")
DATA_DIR = os.path.join(BASE_DIR, "data")
STATE_PATH = os.path.join(DATA_DIR, "current_state.json")

NS_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NS_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
NS_PACKAGE_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"

TARGET_SHEETS = ("ID003", "ID004", "ID005", "ID006")
EXPENSE_SHEET = "TOTALE SPESE Sandu Francesco"

SPLITS = {
    "ID003": {"francesco": 0.20, "sandu": 0.05, "total": 0.25, "airbnb_auto": True},
    "ID004": {"francesco": 0.15, "sandu": 0.10, "total": 0.25, "airbnb_auto": True},
    "ID005": {"francesco": 0.05, "sandu": 0.05, "total": 0.10, "airbnb_auto": False},
    "ID006": {"francesco": 0.10, "sandu": 0.05, "total": 0.15, "airbnb_auto": False},
}


def col_to_num(col):
    total = 0
    for ch in str(col).upper():
        total = total * 26 + ord(ch) - 64
    return total


def num_to_col(num):
    col = ""
    while num:
        num, rem = divmod(num - 1, 26)
        col = chr(65 + rem) + col
    return col


def split_cell_ref(ref):
    match = re.match(r"([A-Z]+)(\d+)", ref or "")
    if not match:
        return None, None
    return match.group(1), int(match.group(2))


def excel_date(value):
    if value in (None, ""):
        return None
    try:
        serial = float(value)
    except (TypeError, ValueError):
        return value
    return (datetime(1899, 12, 30) + timedelta(days=serial)).date().isoformat()


def money(value):
    if value in (None, ""):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    text = re.sub(r"[^\d,\.\-]", "", text)
    if not text:
        return 0.0
    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".")
    elif "," in text:
        text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return 0.0


def parse_iso_date(value):
    if not value:
        return None
    try:
        return datetime.fromisoformat(str(value)).date()
    except ValueError:
        return None


def previous_month_end(today=None):
    base = today or datetime.now().date()
    return base.replace(day=1) - timedelta(days=1)


def norm(value):
    return re.sub(r"[^a-z0-9]+", "", str(value or "").strip().lower())


def display(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def is_empty_row(row):
    return not any(str(value or "").strip() for value in row.values())


def find_col(headers, candidates):
    normalized = {col: norm(value) for col, value in headers.items()}
    wanted = [norm(candidate) for candidate in candidates]
    for candidate in wanted:
        for col, value in normalized.items():
            if value == candidate:
                return col
    for candidate in wanted:
        for col, value in normalized.items():
            if candidate and candidate in value:
                return col
    return None


def formula_value(formula):
    if not formula:
        return None
    expr = formula.lstrip("=")
    if not re.fullmatch(r"[0-9\.\+\-\*/\(\) ]+", expr):
        return None
    try:
        return float(eval(expr, {"__builtins__": {}}, {}))
    except Exception:
        return None


class XlsxReader:
    def __init__(self, raw):
        self.zip = zipfile.ZipFile(BytesIO(raw))
        self.shared_strings = self._shared_strings()
        self.date_style_ids = self._date_style_ids()
        self.sheets = self._sheets()

    def _xml(self, name):
        return ET.fromstring(self.zip.read(name))

    def _shared_strings(self):
        try:
            root = self._xml("xl/sharedStrings.xml")
        except KeyError:
            return []
        return ["".join(t.text or "" for t in item.iter(f"{NS_MAIN}t")) for item in root.findall(f"{NS_MAIN}si")]

    def _date_style_ids(self):
        try:
            root = self._xml("xl/styles.xml")
        except KeyError:
            return set()
        date_ids = {14, 15, 16, 17, 22, 27, 30, 36, 45, 46, 47, 50, 57}
        custom = set()
        num_fmts = root.find(f"{NS_MAIN}numFmts")
        if num_fmts is not None:
            for fmt in num_fmts.findall(f"{NS_MAIN}numFmt"):
                code = fmt.attrib.get("formatCode", "").lower()
                if any(token in code for token in ("yy", "dd", "mm")):
                    custom.add(int(fmt.attrib.get("numFmtId", 0)))
        styles = set()
        cell_xfs = root.find(f"{NS_MAIN}cellXfs")
        if cell_xfs is not None:
            for idx, xf in enumerate(cell_xfs.findall(f"{NS_MAIN}xf")):
                fmt_id = int(xf.attrib.get("numFmtId", 0))
                if fmt_id in date_ids or fmt_id in custom:
                    styles.add(str(idx))
        return styles

    def _sheets(self):
        workbook = self._xml("xl/workbook.xml")
        rels = self._xml("xl/_rels/workbook.xml.rels")
        rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels.findall(f"{NS_PACKAGE_REL}Relationship")}
        sheets = []
        for sheet in workbook.findall(f".//{NS_MAIN}sheet"):
            target = rel_map[sheet.attrib[f"{NS_REL}id"]]
            sheets.append({"name": sheet.attrib["name"], "path": "xl/" + target.lstrip("/")})
        return sheets

    def cell_value(self, cell):
        cell_type = cell.attrib.get("t")
        style_id = cell.attrib.get("s")
        formula = cell.find(f"{NS_MAIN}f")
        value = cell.find(f"{NS_MAIN}v")
        text = value.text if value is not None else None
        if text is None and formula is not None:
            return formula_value(formula.text)
        if cell_type == "s":
            return self.shared_strings[int(text)] if text is not None else None
        if cell_type == "inlineStr":
            return "".join(t.text or "" for t in cell.iter(f"{NS_MAIN}t"))
        if style_id in self.date_style_ids:
            return excel_date(text)
        if text is None:
            return None
        try:
            value = float(text)
            return int(value) if value.is_integer() else value
        except ValueError:
            return text

    def rows(self, sheet_path):
        root = self._xml(sheet_path)
        rows = {}
        for cell in root.findall(f".//{NS_MAIN}c"):
            col, row_num = split_cell_ref(cell.attrib.get("r"))
            if not col:
                continue
            rows.setdefault(row_num, {})[col] = self.cell_value(cell)
        return rows


def sheet_by_name(reader, target):
    wanted = norm(target)
    target_id = norm(target[:5])
    for sheet in reader.sheets:
        if norm(sheet["name"]) == wanted:
            return sheet
    if target_id:
        for sheet in reader.sheets:
            if norm(sheet["name"]).startswith(target_id):
                return sheet
    return None


def sheet_label(sheet_name, sheet_id):
    text = str(sheet_name or sheet_id).strip()
    if text.startswith(sheet_id):
        text = text[len(sheet_id):].strip("_ -")
    text = text.replace("_", " ").replace("-", " ")
    text = re.sub(r"\s+", " ", text).strip()
    return f"{sheet_id} - {text}" if text else sheet_id


def find_booking_header(rows):
    best = None
    best_score = 0
    for row_num, row in rows.items():
        values = [norm(value) for value in row.values()]
        score = 0
        if any("datain" in value or "checkin" in value or value == "in" for value in values):
            score += 2
        if any("dataout" in value or "checkout" in value or value == "out" for value in values):
            score += 2
        if any("portale" in value or "canale" in value or "booking" in value or "airbnb" in value for value in values):
            score += 1
        if any("pulizie" in value for value in values):
            score += 1
        if any("ota" in value or "commission" in value for value in values):
            score += 1
        if any("francesco" in value for value in values):
            score += 1
        if score > best_score:
            best = row_num
            best_score = score
    return best if best_score >= 3 else None


def channel_name(value):
    text = str(value or "").strip()
    key = norm(text)
    if "airbnb" in key:
        return "Airbnb"
    if "booking" in key:
        return "Booking"
    if "dirett" in key or "direct" in key or "privat" in key:
        return "Dirette"
    return text or "Non indicato"


def parse_apartment_sheet(sheet_name, rows, source_sheet_name=None):
    label = sheet_label(source_sheet_name, sheet_name)
    header_rows = []
    for row_num in sorted(rows):
        row = rows[row_num]
        if (
            find_col(row, ["Data IN", "Check in", "Check-in", "Arrivo"])
            and find_col(row, ["Data OUT", "Check out", "Check-out", "Partenza"])
            and find_col(row, ["Portale", "Canale", "Sito", "OTA", "Provenienza"])
        ):
            header_rows.append(row_num)
    if not header_rows:
        return {"id": sheet_name, "label": label, "sourceSheet": source_sheet_name or sheet_name, "bookings": [], "summary": {}, "warnings": ["Intestazione prenotazioni non trovata."]}

    split = SPLITS[sheet_name]
    bookings = []
    warnings = []

    for idx, header_row in enumerate(header_rows):
        next_header = header_rows[idx + 1] if idx + 1 < len(header_rows) else None
        headers = rows[header_row]
        col_in = find_col(headers, ["Data IN", "Check in", "Check-in", "Arrivo"])
        col_out = find_col(headers, ["Data OUT", "Check out", "Check-out", "Partenza"])
        col_channel = find_col(headers, ["Portale", "Canale", "Sito", "OTA", "Provenienza"])
        col_gross = find_col(headers, ["Lordo", "Totale", "Importo", "Incasso", "Prezzo", "Pagamento"])
        col_ota = find_col(headers, ["% OTA", "Commissioni OTA", "Commissione OTA", "Fee OTA", "OTA"])
        col_cleaning = find_col(headers, ["Pulizie", "Pulizia", "Cleaning"])
        col_net = find_col(headers, ["Netto proprieta", "Netto proprietà", "Netto Prop", "Proprietà"])
        col_francesco = (
            find_col(headers, ["Gestione PM - Francesco", "Gestione Francesco"])
            if sheet_name == "ID005"
            else find_col(headers, ["Francesco"])
        )
        col_sandu = None if sheet_name == "ID005" else find_col(headers, ["Sandu"])

        if not col_net:
            warnings.append(f"Riga {header_row}: colonna netto proprieta non trovata, usato Lordo - OTA - Pulizie quando possibile.")

        for row_num in sorted(rows):
            if row_num <= header_row or (next_header and row_num >= next_header):
                continue
            row = rows[row_num]
            if is_empty_row(row):
                continue
            check_in = row.get(col_in) if col_in else ""
            check_out = row.get(col_out) if col_out else ""
            channel = channel_name(row.get(col_channel) if col_channel else "")
            if channel not in {"Airbnb", "Booking", "Dirette"}:
                continue
            gross = money(row.get(col_gross) if col_gross else 0)
            ota = money(row.get(col_ota) if col_ota else 0)
            cleaning = money(row.get(col_cleaning) if col_cleaning else 0)
            net_owner = money(row.get(col_net) if col_net else 0)
            if not net_owner and gross:
                net_owner = gross - ota - cleaning
            francesco = money(row.get(col_francesco) if col_francesco else 0)
            if sheet_name == "ID005" and col_francesco:
                francesco = francesco / 2
                sandu = francesco
            else:
                sandu = money(row.get(col_sandu) if col_sandu else 0)
            if not francesco and not sandu and net_owner and not (sheet_name == "ID005" and col_francesco):
                francesco = net_owner * split["francesco"]
                sandu = net_owner * split["sandu"]
            if not any([gross, ota, cleaning, net_owner, francesco, sandu]) and not str(check_in or "").strip():
                continue
            booking = {
                "sheet": sheet_name,
                "apartmentLabel": label,
                "row": row_num,
                "checkIn": display(check_in),
                "checkOut": display(check_out),
                "channel": channel,
                "gross": gross,
                "ota": ota,
                "cleaning": cleaning,
                "netOwner": net_owner,
                "francesco": francesco,
                "sandu": sandu,
                "sanduDueByFrancesco": 0 if channel == "Airbnb" and split["airbnb_auto"] else sandu,
                "notes": "",
            }
            bookings.append(booking)

    summary = {}
    for booking in bookings:
        channel = booking["channel"]
        item = summary.setdefault(channel, {"gross": 0, "netOwner": 0, "francesco": 0, "sandu": 0, "sanduDueByFrancesco": 0, "count": 0})
        item["count"] += 1
        for key in ("gross", "netOwner", "francesco", "sandu", "sanduDueByFrancesco"):
            item[key] += booking[key]

    return {"id": sheet_name, "label": label, "sourceSheet": source_sheet_name or sheet_name, "bookings": bookings, "summary": summary, "warnings": warnings}


def range_rows(rows, start_col, end_col):
    start = col_to_num(start_col)
    end = col_to_num(end_col)
    output = []
    for row_num in sorted(rows):
        row = rows[row_num]
        selected = {}
        for col_num in range(start, end + 1):
            col = num_to_col(col_num)
            selected[col] = display(row.get(col, ""))
        if any(str(value or "").strip() for value in selected.values()):
            output.append({"row": row_num, "cells": selected})
    return output


def parse_expenses(rows):
    raw_sections = {
        "francescoCosts": {
            "title": "ADS e spese appartamenti sostenute da Francesco",
            "range": "AH:AR",
            "rows": range_rows(rows, "AH", "AR"),
        },
    }

    def title_info(value):
        text = str(value or "").strip()
        match = re.search(r"(20\d{2})", text)
        if not text or not match:
            return None
        upper = text.upper()
        if "TOTALE SPESE" in upper:
            return {"title": text, "year": int(match.group(1)), "kind": "TOTAL"}
        if "ADS" in upper or "SPESE" in upper:
            clean = re.sub(r"\s*-\s*20\d{2}\s*$", "", text, flags=re.IGNORECASE).strip()
            kind = "ADS" if "ADS" in upper else "APPARTAMENTO"
            name = clean.replace("SPESE", "").replace("ADS", "ADS").strip(" -")
            return {"title": text, "year": int(match.group(1)), "kind": kind, "name": name or clean}
        return None

    def parse_side(first_col, desc_col, total_col, share_col, note_col):
        groups = []
        current = None
        for row_num in sorted(rows):
            row = rows[row_num]
            info = title_info(row.get(first_col))
            if info and info["kind"] != "TOTAL":
                current = {
                    "title": info["title"],
                    "name": info.get("name") or info["title"],
                    "year": info["year"],
                    "kind": info["kind"],
                    "items": [],
                    "total": 0,
                    "share": 0,
                }
                groups.append(current)
                continue
            if current is None:
                continue
            first = str(row.get(first_col) or "").strip()
            if not first or norm(first) == "cosa":
                continue
            if norm(first) == "totale":
                current["total"] = money(row.get(total_col)) or current["total"]
                current["share"] = money(row.get(share_col)) or current["share"]
                current = None
                continue
            total = money(row.get(total_col))
            share = money(row.get(share_col))
            description = str(row.get(desc_col) or "").strip()
            note = str(row.get(note_col) or "").strip()
            if not description and not total and not share:
                continue
            item = {
                "row": row_num,
                "category": first,
                "description": description,
                "total": total,
                "share": share,
                "note": note,
            }
            current["items"].append(item)
            current["total"] += total
            current["share"] += share
        return groups

    groups = (
        parse_side("AH", "AI", "AJ", "AK", "AL")
        + parse_side("AN", "AO", "AP", "AQ", "AR")
    )
    groups.sort(key=lambda item: (item["year"], 0 if item["kind"] == "APPARTAMENTO" else 1, item["name"]))

    items = []
    totals = {}
    for group in groups:
        key = f"{group['year']}:{group['kind']}:{group['name']}"
        totals[key] = {
            "title": group["title"],
            "year": group["year"],
            "kind": group["kind"],
            "name": group["name"],
            "total": group["total"],
            "share": group["share"],
            "count": len(group["items"]),
        }
        for item in group["items"]:
            copied = dict(item)
            copied.update({
                "section": key,
                "year": group["year"],
                "kind": group["kind"],
                "group": group["name"],
                "francescoShare": item["share"],
                "sanduShare": item["share"],
            })
            items.append(copied)
    return {"sections": raw_sections, "groups": groups, "items": items, "totals": totals}


def parse_due_blocks(rows, side):
    if side == "francesco_to_sandu":
        cols = {"head": "B", "category": "B", "description": "C", "checkIn": "D", "checkOut": "E", "amount": "F", "note": "G", "paid": "H"}
    else:
        cols = {"head": "T", "category": "T", "description": "U", "checkIn": "V", "checkOut": "W", "amount": "X", "note": "Y", "paid": "Z"}
    blocks = []
    current = None
    for row_num in sorted(rows):
        row = rows[row_num]
        head = str(row.get(cols["head"]) or "").strip()
        match = re.search(r"(20\d{2})", head)
        if match and "SPESE" in head.upper():
            current = {
                "year": int(match.group(1)),
                "title": head,
                "side": side,
                "items": [],
                "total": 0,
            }
            blocks.append(current)
            continue
        if current is None:
            continue
        category = str(row.get(cols["category"]) or "").strip()
        if not category or norm(category) == "cosa":
            continue
        amount = money(row.get(cols["amount"]))
        if norm(category) == "totale":
            current["total"] = amount or current["total"]
            current = None
            continue
        description = str(row.get(cols["description"]) or "").strip()
        note = str(row.get(cols["note"]) or "").strip()
        if not amount and not description and not note:
            continue
        item = {
            "row": row_num,
            "year": current["year"],
            "side": side,
            "category": category,
            "description": description,
            "checkIn": display(row.get(cols["checkIn"])),
            "checkOut": display(row.get(cols["checkOut"])),
            "amount": amount,
            "note": note,
            "paid": str(row.get(cols["paid"]) or "").strip(),
        }
        current["items"].append(item)
        current["total"] += amount
    return blocks


def parse_payments(rows):
    payments = []
    year_cols = {}
    for col in ("AC", "AD"):
        year = money(rows.get(2, {}).get(col))
        if year:
            year_cols[col] = int(year)
    for row_num in sorted(rows):
        label = str(rows[row_num].get("AB") or "").strip()
        normalized = norm(label)
        if "pagamentofrancescosandu" in normalized:
            side = "francesco_to_sandu"
        elif "pagamentosandufrancesco" in normalized:
            side = "sandu_to_francesco"
        else:
            continue
        for col, year in year_cols.items():
            amount = money(rows[row_num].get(col))
            if amount:
                payments.append({
                    "row": row_num,
                    "year": year,
                    "side": side,
                    "amount": amount,
                    "label": label,
                })
    return payments


def add_settlement_bucket(years, year, side, group, amount):
    year_key = str(year)
    entry = years.setdefault(year_key, {
        "year": year,
        "francescoToSandu": {},
        "sanduToFrancesco": {},
        "francescoToSanduTotal": 0,
        "sanduToFrancescoTotal": 0,
        "delta": 0,
    })
    target_key = "francescoToSandu" if side == "francesco_to_sandu" else "sanduToFrancesco"
    total_key = "francescoToSanduTotal" if side == "francesco_to_sandu" else "sanduToFrancescoTotal"
    entry[target_key][group] = entry[target_key].get(group, 0) + amount
    entry[total_key] += amount


def add_personal_bucket(years, year, person, group, amount):
    year_key = str(year)
    entry = years.setdefault(year_key, {
        "year": year,
        "francesco": {},
        "sandu": {},
        "francescoTotal": 0,
        "sanduTotal": 0,
    })
    total_key = "francescoTotal" if person == "francesco" else "sanduTotal"
    entry[person][group] = entry[person].get(group, 0) + amount
    entry[total_key] += amount


def finalize_personal_years(years):
    overall = {"francesco": 0, "sandu": 0}
    for entry in years.values():
        overall["francesco"] += entry["francescoTotal"]
        overall["sandu"] += entry["sanduTotal"]
    return {
        "overall": overall,
        "years": [years[key] for key in sorted(years.keys(), reverse=True)],
    }


def parse_workbook(raw, filename):
    reader = XlsxReader(raw)
    apartments = []
    all_bookings = []
    warnings = []
    for sheet_name in TARGET_SHEETS:
        sheet = sheet_by_name(reader, sheet_name)
        if not sheet:
            warnings.append(f"Foglio {sheet_name} non trovato.")
            apartments.append({"id": sheet_name, "label": sheet_name, "sourceSheet": sheet_name, "bookings": [], "summary": {}, "warnings": ["Foglio non trovato."]})
            continue
        parsed = parse_apartment_sheet(sheet_name, reader.rows(sheet["path"]), sheet["name"])
        apartments.append(parsed)
        all_bookings.extend(parsed["bookings"])
        warnings.extend([f"{sheet_name}: {warning}" for warning in parsed["warnings"]])

    expense_sheet = sheet_by_name(reader, EXPENSE_SHEET)
    expense_rows = reader.rows(expense_sheet["path"]) if expense_sheet else {}
    expenses = parse_expenses(expense_rows) if expense_sheet else {"sections": {}, "groups": [], "items": [], "totals": {}}
    due_to_sandu_blocks = parse_due_blocks(expense_rows, "francesco_to_sandu") if expense_sheet else []
    due_to_francesco_blocks = parse_due_blocks(expense_rows, "sandu_to_francesco") if expense_sheet else []
    payments = parse_payments(expense_rows) if expense_sheet else []
    if not expense_sheet:
        warnings.append(f"Foglio {EXPENSE_SHEET} non trovato.")

    channels = {}
    for booking in all_bookings:
        bucket = channels.setdefault(booking["channel"], {"gross": 0, "netOwner": 0, "francesco": 0, "sandu": 0, "sanduDueByFrancesco": 0, "count": 0})
        bucket["count"] += 1
        for key in ("gross", "netOwner", "francesco", "sandu", "sanduDueByFrancesco"):
            bucket[key] += booking[key]

    cutoff_date = previous_month_end()
    settlement_years = {}
    projection_years = {}
    personal_years = {}
    personal_projection_years = {}
    for booking in all_bookings:
        check_in = parse_iso_date(booking.get("checkIn"))
        amount = money(booking.get("sanduDueByFrancesco"))
        if check_in:
            if amount:
                target = settlement_years if check_in <= cutoff_date else projection_years
                add_settlement_bucket(target, check_in.year, "francesco_to_sandu", "Utili appartamenti", amount)
            personal_target = personal_years if check_in <= cutoff_date else personal_projection_years
            add_personal_bucket(personal_target, check_in.year, "francesco", "Quote prenotazioni", money(booking.get("francesco")))
            add_personal_bucket(personal_target, check_in.year, "sandu", "Quote prenotazioni", money(booking.get("sandu")))
    for block in due_to_sandu_blocks:
        for item in block["items"]:
            add_settlement_bucket(settlement_years, item["year"], "francesco_to_sandu", item["category"] or "Sanremo", item["amount"])
    for item in expenses["items"]:
        group = "Varie" if norm(item.get("category")) == "varie" else ("ADS" if item.get("kind") == "ADS" else "Spese appartamenti")
        add_settlement_bucket(settlement_years, item["year"], "sandu_to_francesco", group, money(item.get("share")))
        add_personal_bucket(personal_years, item["year"], "francesco", group, -money(item.get("share")))
        add_personal_bucket(personal_years, item["year"], "sandu", group, -money(item.get("share")))
    for block in due_to_francesco_blocks:
        for item in block["items"]:
            add_settlement_bucket(settlement_years, item["year"], "sandu_to_francesco", "Tassa soggiorno Sanremo", item["amount"])
    for payment in payments:
        add_settlement_bucket(
            settlement_years,
            payment["year"],
            payment["side"],
            "Pagamenti avvenuti",
            -payment["amount"],
        )

    overall = {
        "francescoToSandu": 0,
        "sanduToFrancesco": 0,
        "delta": 0,
    }
    for entry in settlement_years.values():
        entry["delta"] = entry["francescoToSanduTotal"] - entry["sanduToFrancescoTotal"]
        overall["francescoToSandu"] += entry["francescoToSanduTotal"]
        overall["sanduToFrancesco"] += entry["sanduToFrancescoTotal"]
    overall["delta"] = overall["francescoToSandu"] - overall["sanduToFrancesco"]
    projection_overall = {"francescoToSandu": 0, "sanduToFrancesco": 0, "delta": 0}
    for entry in projection_years.values():
        entry["delta"] = entry["francescoToSanduTotal"] - entry["sanduToFrancescoTotal"]
        projection_overall["francescoToSandu"] += entry["francescoToSanduTotal"]
        projection_overall["sanduToFrancesco"] += entry["sanduToFrancescoTotal"]
    projection_overall["delta"] = projection_overall["francescoToSandu"] - projection_overall["sanduToFrancesco"]
    personal_final = finalize_personal_years(personal_years)
    personal_projection_final = finalize_personal_years(personal_projection_years)

    settlement = {
        "overall": overall,
        "cutoffDate": cutoff_date.isoformat(),
        "years": [settlement_years[key] for key in sorted(settlement_years.keys(), reverse=True)],
        "projection": {
            "overall": projection_overall,
            "years": [projection_years[key] for key in sorted(projection_years.keys(), reverse=True)],
            "note": "Proiezione su prenotazioni future: valori da confermare quando i soggiorni saranno conclusi o modificati.",
        },
        "personalProgress": {
            "overall": personal_final["overall"],
            "years": personal_final["years"],
            "projection": {
                "overall": personal_projection_final["overall"],
                "years": personal_projection_final["years"],
                "note": "Proiezione personale sulle prenotazioni future, senza spese non ancora inserite.",
            },
        },
        "payments": payments,
        "dueToSanduBlocks": due_to_sandu_blocks,
        "dueToFrancescoBlocks": due_to_francesco_blocks,
        "francescoToSanduProfit": sum(booking["sanduDueByFrancesco"] for booking in all_bookings),
        "francescoToSanduCleaning": sum(block["total"] for block in due_to_sandu_blocks),
        "sanduToFrancescoExpenses": sum(item.get("sanduShare", 0) for item in expenses["items"]),
        "netFrancescoPaysSandu": overall["delta"],
    }

    uploaded_at = datetime.now(timezone.utc).isoformat(timespec="seconds").replace("+00:00", "Z")
    return {
        "filename": filename,
        "uploadedAt": uploaded_at,
        "apartments": apartments,
        "bookings": all_bookings,
        "channels": channels,
        "expenses": expenses,
        "settlement": settlement,
        "warnings": warnings,
        "rules": SPLITS,
    }


def load_state():
    if not os.path.exists(STATE_PATH):
        return None
    with open(STATE_PATH, "r", encoding="utf-8") as handle:
        return json.load(handle)


def save_state(state):
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(STATE_PATH, "w", encoding="utf-8") as handle:
        json.dump(state, handle, ensure_ascii=False, indent=2)


def json_response(handler, payload, status=200):
    raw = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    handler.send_response(status)
    handler.send_header("content-type", "application/json; charset=utf-8")
    handler.send_header("content-length", str(len(raw)))
    handler.end_headers()
    handler.wfile.write(raw)


class Handler(BaseHTTPRequestHandler):
    def do_GET(self):
        if self.path.startswith("/api/current"):
            json_response(self, {"state": load_state()})
            return
        path = self.path.split("?", 1)[0]
        if path == "/":
            path = "/index.html"
        file_path = os.path.abspath(os.path.join(PUBLIC_DIR, path.lstrip("/")))
        if not file_path.startswith(PUBLIC_DIR) or not os.path.exists(file_path):
            self.send_error(404)
            return
        with open(file_path, "rb") as handle:
            raw = handle.read()
        self.send_response(200)
        self.send_header("content-type", mimetypes.guess_type(file_path)[0] or "application/octet-stream")
        self.send_header("content-length", str(len(raw)))
        self.end_headers()
        self.wfile.write(raw)

    def do_POST(self):
        try:
            if not self.path.startswith("/api/upload"):
                self.send_error(404)
                return
            length = int(self.headers.get("content-length", "0"))
            payload = json.loads(self.rfile.read(length).decode("utf-8"))
            filename = str(payload.get("filename") or "workbook.xlsx")
            if not filename.lower().endswith(".xlsx"):
                json_response(self, {"error": "Carica un file .xlsx. I vecchi .xls vanno prima salvati come .xlsx da Excel o Numbers."}, 400)
                return
            raw = base64.b64decode(payload.get("contentBase64") or "")
            state = parse_workbook(raw, filename)
            save_state(state)
            json_response(self, state)
        except zipfile.BadZipFile:
            json_response(self, {"error": "Il file non sembra un .xlsx valido."}, 400)
        except Exception as error:
            traceback.print_exc()
            json_response(self, {"error": f"Errore durante la lettura del file: {error}"}, 500)

    def log_message(self, fmt, *args):
        print("%s - %s" % (self.address_string(), fmt % args))


if __name__ == "__main__":
    os.makedirs(DATA_DIR, exist_ok=True)
    print(f"Rendiconto Sandu Francesco su http://127.0.0.1:{PORT}")
    if HOST == "0.0.0.0":
        print("Per altre postazioni sulla stessa rete usa http://IP_DEL_MAC:%s" % PORT)
    ThreadingHTTPServer((HOST, PORT), Handler).serve_forever()
