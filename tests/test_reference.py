import io
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import xlsxwriter

import app
from reference import BUCKETS, DISPLAY_BUCKETS, parse_reference_xlsx

ROOT = Path(__file__).resolve().parents[1]


def reference_fixture(first_bucket=0, extra_51=1, count_31=1):
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output)
    sheet = workbook.add_worksheet("Export detailed")
    sheet.write(0, 8, "21-Сентябрь Давомилик")
    sheet.write(1, 0, "ХГТ булими")
    for col, bucket in enumerate(BUCKETS, 1):
        sheet.write(1, col, "70 кундан ошган" if bucket == "71+" else f"{bucket} кун")
    sheet.write(1, 11, "умуман газ олмаганлар")
    sheet.write(1, 12, "Муддатида алмаштиришга эҳтиёжи мавжуд эмас")
    sheet.write(2, 0, "ТЕСТ РАЙГАЗ")
    for col, bucket in enumerate(BUCKETS, 1):
        sheet.write_number(2, col, first_bucket if bucket == "1-30" else
                           (count_31 if bucket == "31-35" else extra_51 if bucket == "51-55" else 0))
    sheet.write_number(2, 11, 0)
    sheet.write_number(2, 12, 0)
    workbook.close()
    return output.getvalue()


def screenshot_fixture(extra_51=1, wrong_total=False):
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output)
    sheet = workbook.add_worksheet("Sheet1")
    sheet.write_row(0, 0, ["Райгаз", *DISPLAY_BUCKETS, "Жами"])
    counts = [1 if bucket == "31-35" else extra_51 if bucket == "51-55" else 0
              for bucket in DISPLAY_BUCKETS]
    sheet.write_row(1, 0, ["ТЕСТ РАЙГАЗ", *counts, sum(counts) + int(wrong_total)])
    sheet.write_row(2, 0, ["ЖАМИ", *counts, sum(counts)])
    workbook.close()
    return output.getvalue()


def sheet_cell(workbook_bytes, sheet_number, address):
    ns = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with zipfile.ZipFile(io.BytesIO(workbook_bytes)) as archive:
        root = ET.fromstring(archive.read(f"xl/worksheets/sheet{sheet_number}.xml"))
    cell = root.find(f".//x:c[@r='{address}']", ns)
    return cell.find("x:v", ns).text if cell is not None and cell.find("x:v", ns) is not None else None


def sheet_text(workbook_bytes, sheet_number, address):
    ns = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with zipfile.ZipFile(io.BytesIO(workbook_bytes)) as archive:
        root = ET.fromstring(archive.read(f"xl/worksheets/sheet{sheet_number}.xml"))
    cell = root.find(f".//x:c[@r='{address}']", ns)
    return "".join(cell.itertext()) if cell is not None else None


class ReferenceTests(unittest.TestCase):
    def test_reference_fixture_parses_district_buckets(self):
        parsed = parse_reference_xlsx(reference_fixture(), app.header_key)
        self.assertEqual(len(parsed["districts"]), 1)
        self.assertEqual(parsed["totals"]["31-35"], 1)
        self.assertEqual(parsed["totals"]["51-55"], 1)
        self.assertEqual(len(parsed["special_headers"]), 2)

    def test_screenshot_layout_parses_without_1_30(self):
        parsed = parse_reference_xlsx(screenshot_fixture(), app.header_key)
        self.assertEqual(parsed["buckets"], DISPLAY_BUCKETS)
        self.assertEqual(parsed["totals"]["31-35"], 1)
        self.assertEqual(parsed["totals"]["51-55"], 1)
        self.assertEqual(parsed["title"], "")

    def test_screenshot_layout_checks_row_total(self):
        with self.assertRaisesRegex(ValueError, "ustuni oraliqlar"):
            parse_reference_xlsx(screenshot_fixture(wrong_total=True), app.header_key)

    def test_upload_comparison_identifies_one_changed_bucket(self):
        csv_data = (ROOT / "NAMUNA.csv").read_bytes()
        for changed_count, expected_differences in ((1, 0), (2, 1)):
            with self.subTest(changed_count=changed_count):
                result = app.app.test_client().post(
                    "/generate",
                    data={"files": (io.BytesIO(csv_data), "NAMUNA.csv"),
                          "reference": (io.BytesIO(reference_fixture(extra_51=changed_count)), "etalon.xlsx")},
                )
                self.assertEqual(result.status_code, 200, result.get_data(as_text=True) if result.is_json else "")
                self.assertEqual(int(result.headers["X-Mismatch-Count"]), expected_differences)
                with zipfile.ZipFile(io.BytesIO(result.data)) as archive:
                    names = archive.read("xl/workbook.xml").decode("utf-8")
                self.assertIn("Давомилик", names)
                self.assertIn("Солиштириш", names)
                self.assertEqual(sheet_cell(result.data, 2, "B2"), "1")
                self.assertEqual(sheet_cell(result.data, 2, "K2"), "2")
                if changed_count == 2:
                    self.assertEqual(sheet_text(result.data, 3, "G3"), "2 -1")  # 2 etalon, 1 yangi
                    self.assertEqual(sheet_cell(result.data, 3, "L3"), "0")
                    self.assertEqual(sheet_cell(result.data, 3, "M3"), "0")
                    with zipfile.ZipFile(io.BytesIO(result.data)) as archive:
                        styles = archive.read("xl/styles.xml")
                        comparison_xml = archive.read("xl/worksheets/sheet3.xml")
                    self.assertIn(b"FFA6A6A6", styles)  # 31-40 day columns
                    self.assertIn(b"FFFFFF00", styles)  # 66+ day columns
                    self.assertIn(b"FFE60000", comparison_xml)  # inline signed change

    def test_screenshot_layout_upload_compares_as_matrix(self):
        result = app.app.test_client().post(
            "/generate",
            data={"files": (io.BytesIO((ROOT / "NAMUNA.csv").read_bytes()), "NAMUNA.csv"),
                  "reference": (io.BytesIO(screenshot_fixture(extra_51=2)), "current.xlsx")},
        )
        self.assertEqual(result.status_code, 200, result.get_data(as_text=True) if result.is_json else "")
        self.assertEqual(result.headers["X-Mismatch-Count"], "1")
        self.assertEqual(sheet_cell(result.data, 2, "B2"), "1")
        self.assertEqual(sheet_text(result.data, 3, "F3"), "2 -1")  # etalon minus yangi amount
        self.assertEqual(sheet_text(result.data, 3, "K3"), "3 -2")
        self.assertEqual(sheet_cell(result.data, 4, "F21"), "1")
        self.assertEqual(sheet_cell(result.data, 4, "K10"), "3")
        self.assertEqual(sheet_cell(result.data, 4, "K16"), "2")
        self.assertEqual(sheet_cell(result.data, 4, "K22"), "1")

    def test_reference_6841_new_zero_shows_no_red_subtraction(self):
        sample_lines = (ROOT / "NAMUNA.csv").read_bytes().splitlines()
        csv_data = b"\n".join((sample_lines[0], sample_lines[2])) + b"\n"
        result = app.app.test_client().post(
            "/generate",
            data={"files": (io.BytesIO(csv_data), "one.csv"),
                  "reference": (io.BytesIO(reference_fixture(count_31=6841)), "etalon.xlsx")},
        )
        self.assertEqual(result.status_code, 200, result.get_data(as_text=True) if result.is_json else "")
        self.assertEqual(sheet_cell(result.data, 3, "C3"), "6841")
        self.assertEqual(sheet_cell(result.data, 4, "C21"), "6841")

    def test_new_amount_is_always_red_subtraction(self):
        result = app.app.test_client().post(
            "/generate",
            data={"files": (io.BytesIO((ROOT / "NAMUNA.csv").read_bytes()), "NAMUNA.csv"),
                  "reference": (io.BytesIO(screenshot_fixture(extra_51=0)), "current.xlsx")},
        )
        self.assertEqual(result.status_code, 200, result.get_data(as_text=True) if result.is_json else "")
        self.assertEqual(sheet_text(result.data, 3, "F3"), "0 -1")
        self.assertEqual(sheet_text(result.data, 3, "K3"), "1 -2")
        self.assertEqual(sheet_cell(result.data, 4, "F21"), "-1")

    def test_pakhtaobod_5829_minus_1383_display(self):
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output)
        sheet = workbook.add_worksheet("Etalon")
        sheet.write_row(0, 0, ["Райгаз", *DISPLAY_BUCKETS, "Жами"])
        values = [0, 5829, 0, 0, 0, 0, 0, 0, 0]
        sheet.write_row(1, 0, ["Пахтаобод туман ГАЗ", *values, sum(values)])
        workbook.close()
        csv_data = ("Вилоят;Райгаз;Маҳалла;Абонент код;Абонент;Еҳтиёж;Сўнги реализация;Бугунги реализация\n"
                    + "\n".join(f"АНДИЖОН;Пахтаобод туман ГАЗ;ТЕСТ;{i};ТЕСТ;1;2026-08-14;2026-09-22"
                                for i in range(1, 1384)) + "\n").encode()
        result = app.app.test_client().post(
            "/generate",
            data={"files": (io.BytesIO(csv_data), "pakhtaobod.csv"),
                  "reference": (io.BytesIO(output.getvalue()), "etalon.xlsx")},
        )
        self.assertEqual(result.status_code, 200, result.get_data(as_text=True) if result.is_json else "")
        self.assertEqual(sheet_text(result.data, 3, "C3"), "5829 -1383")
        self.assertEqual(sheet_cell(result.data, 4, "C21"), "4446")

    def test_day_zero_is_excluded_from_reference_bucket(self):
        csv_data = ("Вилоят;Райгаз;Маҳалла;Абонент код;Абонент;Еҳтиёж;Сўнги реализация;Бугунги реализация\n"
                    "НАМАНГАН;ТЕСТ РАЙГАЗ;ТЕСТ МФЙ;100001;ТЕСТ;1;2026-09-21;2026-09-21\n").encode()
        result = app.app.test_client().post("/generate", data={"files": (io.BytesIO(csv_data), "zero.csv")})
        self.assertEqual(result.status_code, 200)
        self.assertEqual(sheet_cell(result.data, 2, "B2"), "0")
        self.assertEqual(sheet_cell(result.data, 1, "B12"), "1")


if __name__ == "__main__":
    unittest.main()
