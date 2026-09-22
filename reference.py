"""Read a district-level attendance summary supplied for comparison."""

from __future__ import annotations

import io
import re
import zipfile
from collections import Counter
from xml.etree import ElementTree as ET

NS = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
BUCKETS = ("1-30", "31-35", "36-40", "41-45", "46-50", "51-55",
           "56-60", "61-65", "66-70", "71+")
DISPLAY_BUCKETS = BUCKETS[1:]


def _column_number(address: str) -> int:
    number = 0
    for char in re.match(r"[A-Z]+", address).group():
        number = number * 26 + ord(char) - ord("A") + 1
    return number


def _text(cell: ET.Element, strings: list[str]) -> str:
    value = cell.find("x:v", NS)
    if cell.get("t") == "s" and value is not None:
        return strings[int(value.text or "0")]
    if cell.get("t") == "inlineStr":
        inline = cell.find("x:is", NS)
        return "".join(inline.itertext()) if inline is not None else ""
    return value.text or "" if value is not None else ""


def _bucket_header(value: str) -> str | None:
    key = " ".join(value.lower().replace("ё", "е").split())
    numbers = re.findall(r"\d+", key)
    if len(numbers) >= 2:
        candidate = f"{numbers[0]}-{numbers[1]}"
        if candidate in BUCKETS:
            return candidate
    if numbers and int(numbers[0]) >= 70 and ("ошган" in key or "+" in key):
        return "71+"
    return None


def parse_reference_xlsx(data: bytes, normalize_name) -> dict:
    """Return summary counts; reference cells are data, never executable instructions."""
    if len(data) > 10 * 1024 * 1024:
        raise ValueError("Solishtirish Excel fayli 10 MB dan oshmasin.")
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            sheet_names = sorted(
                (name for name in archive.namelist()
                 if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", name)),
                key=lambda name: int(re.search(r"sheet(\d+)\.xml", name).group(1)),
            )
            if not sheet_names:
                raise ValueError("Excel faylida varaq topilmadi.")
            if archive.getinfo(sheet_names[0]).file_size > 20 * 1024 * 1024:
                raise ValueError("Excel varag'i juda katta.")
            strings = []
            if "xl/sharedStrings.xml" in archive.namelist():
                if archive.getinfo("xl/sharedStrings.xml").file_size > 10 * 1024 * 1024:
                    raise ValueError("Excel matnlari juda katta.")
                shared = ET.fromstring(archive.read("xl/sharedStrings.xml"))
                strings = ["".join(item.itertext()) for item in shared.findall("x:si", NS)]
            root = ET.fromstring(archive.read(sheet_names[0]))
    except (zipfile.BadZipFile, KeyError, ET.ParseError, IndexError) as exc:
        raise ValueError("Solishtirish uchun haqiqiy .xlsx fayl kerak.") from exc

    rows = {}
    for row in root.findall("x:sheetData/x:row", NS):
        values = {}
        for cell in row.findall("x:c", NS):
            address = cell.get("r", "")
            if re.match(r"^[A-Z]+\d+$", address):
                values[_column_number(address)] = _text(cell, strings)
        rows[int(row.get("r"))] = values

    header_row = next((number for number, row in sorted(rows.items())
                       if normalize_name(row.get(1, "")) == "райгаз"
                       or ("хгт" in normalize_name(row.get(1, ""))
                           and any(word in normalize_name(row.get(1, ""))
                                   for word in ("булими", "бўлими")))), None)
    if header_row is None:
        raise ValueError("Solishtirish faylida 'Райгаз' yoki 'ХГТ булими' sarlavhasi topilmadi.")
    header = rows[header_row]
    columns = {}
    for column, value in header.items():
        bucket = _bucket_header(value)
        if bucket:
            if bucket in columns:
                raise ValueError(f"Solishtirish faylida {bucket} ustuni takrorlangan.")
            columns[bucket] = column
    missing = [bucket for bucket in DISPLAY_BUCKETS if bucket not in columns]
    if missing:
        raise ValueError("Solishtirish faylida oraliq ustunlar yetishmaydi: " + ", ".join(missing))
    buckets = tuple(bucket for bucket in BUCKETS if bucket in columns)
    total_column = next((column for column, value in header.items()
                         if normalize_name(value) in ("жами", "итого", "total")), None)
    special_columns = [(column, value) for column, value in sorted(header.items())
                       if column > max(columns.values()) and column != total_column and value.strip()]

    districts = {}
    names = {}
    special_values = {}
    total_row = None
    for number, row in sorted(rows.items()):
        if number <= header_row:
            continue
        name = row.get(1, "").strip()
        if normalize_name(name) in ("жами", "итого", "total"):
            total_row = row
            break
        if not name:
            continue
        key = normalize_name(name)
        if key in districts:
            raise ValueError(f"Solishtirish faylida hudud takrorlangan: {name}")
        counts = Counter()
        for bucket, column in columns.items():
            raw = str(row.get(column, "")).strip()
            try:
                amount = int(raw)
            except ValueError as exc:
                raise ValueError(f"Solishtirish fayli {number}-qator, {bucket}: butun son kerak.") from exc
            if amount < 0:
                raise ValueError(f"Solishtirish fayli {number}-qator, {bucket}: manfiy son mumkin emas.")
            counts[bucket] = amount
        if total_column is not None:
            raw_total = str(row.get(total_column, "")).replace(" ", "").strip()
            if not raw_total.isdigit() or int(raw_total) != sum(counts.values()):
                raise ValueError(f"Solishtirish fayli {number}-qator: Жами ustuni oraliqlar yig'indisiga teng emas.")
        extras = []
        for column, label in special_columns:
            raw = str(row.get(column, "")).replace(" ", "").strip()
            if raw and (not raw.isdigit() or int(raw) < 0):
                raise ValueError(f"Solishtirish fayli {number}-qator, {label}: butun son kerak.")
            extras.append(int(raw) if raw else None)
        districts[key] = counts
        names[key] = name
        special_values[key] = extras
    if not districts:
        raise ValueError("Solishtirish faylida hudud ma'lumotlari yo'q.")

    totals = Counter({bucket: sum(row[bucket] for row in districts.values()) for bucket in buckets})
    if total_row:
        for bucket, column in columns.items():
            cached = str(total_row.get(column, "")).strip()
            if cached and cached.isdigit() and int(cached) != totals[bucket]:
                raise ValueError(f"Solishtirish faylida {bucket} jami hududlar yig'indisiga teng emas.")
        if total_column is not None:
            cached = str(total_row.get(total_column, "")).replace(" ", "").strip()
            if not cached.isdigit() or int(cached) != sum(totals.values()):
                raise ValueError("Solishtirish faylida umumiy Жами hududlar yig'indisiga teng emas.")
        for index, (column, label) in enumerate(special_columns):
            cached = str(total_row.get(column, "")).replace(" ", "").strip()
            expected = sum(values[index] or 0 for values in special_values.values())
            if cached and (not cached.isdigit() or int(cached) != expected):
                raise ValueError(f"Solishtirish faylida {label} jami hududlar yig'indisiga teng emas.")

    title = rows.get(1, {}).get(9, "").strip() if header_row > 1 else ""
    return {"title": title, "districts": districts, "names": names,
            "totals": totals, "buckets": buckets,
            "first_header": header.get(1, "Райгаз"),
            "bucket_headers": {bucket: header[columns[bucket]] for bucket in buckets},
            "special_headers": [label for _, label in special_columns],
            "special_values": special_values,
            "has_total_column": total_column is not None}
