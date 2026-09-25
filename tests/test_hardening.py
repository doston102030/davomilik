import csv
import io
import unittest

import xlsxwriter

import app
from reference import DISPLAY_BUCKETS, parse_reference_xlsx


class HardeningTests(unittest.TestCase):
    """Twenty strict checks for input integrity, arithmetic, UI, and privacy."""

    @staticmethod
    def csv_bytes(rows, headers=None, delimiter=";"):
        output = io.StringIO()
        writer = csv.writer(output, delimiter=delimiter, lineterminator="\n")
        writer.writerow(headers or app.REQUIRED_HEADERS)
        writer.writerows(rows)
        return output.getvalue().encode("utf-8")

    @staticmethod
    def row(code="1001", last="2026-08-01", snapshot="2026-09-21",
            raygaz="A RAYGAZ", mahalla="A MFY", subscriber="ABONENT", need="1"):
        return ["NAMANGAN", raygaz, mahalla, code, subscriber, need, last, snapshot]

    @staticmethod
    def gas_xlsx(headers=None, values=None):
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {"in_memory": True})
        sheet = workbook.add_worksheet("Hisobot")
        sheet.write_row(0, 0, headers or app.GAS_SALES_HEADERS)
        sheet.write_row(1, 0, values or ["A GAZ", "INSPEKTOR", 100, 80, 20])
        workbook.close()
        return output.getvalue()

    @staticmethod
    def reference_xlsx(blank_total_bucket=False):
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {"in_memory": True})
        # The first sheet deliberately looks like an incomplete reference header.
        workbook.add_worksheet("Izoh").write(0, 0, "Райгаз")
        sheet = workbook.add_worksheet("Etalon")
        counts = list(range(1, len(DISPLAY_BUCKETS) + 1))
        sheet.write_row(0, 0, ["Райгаз", *DISPLAY_BUCKETS, "Жами"])
        sheet.write_row(1, 0, ["A RAYGAZ", *counts, sum(counts)])
        totals = counts.copy()
        if blank_total_bucket:
            totals[0] = ""
        sheet.write_row(2, 0, ["ЖАМИ", *totals, sum(counts)])
        workbook.close()
        return output.getvalue()

    def parse_one(self, data, name="source.csv"):
        return app.parse_files([app.MemoryUpload(name, data)])[0]

    # 1
    def test_security_headers_disable_caching_framing_and_sniffing(self):
        response = app.app.test_client().get("/")
        self.assertEqual(response.headers["Cache-Control"], "no-store, max-age=0")
        self.assertEqual(response.headers["X-Frame-Options"], "DENY")
        self.assertEqual(response.headers["X-Content-Type-Options"], "nosniff")
        self.assertIn("frame-ancestors 'none'", response.headers["Content-Security-Policy"])

    # 2
    def test_pages_include_keyboard_live_region_and_reduced_motion_support(self):
        client = app.app.test_client()
        for path in ("/", "/sotuvlar"):
            page = client.get(path).get_data(as_text=True)
            with self.subTest(path=path):
                self.assertIn('role="status" aria-live="polite"', page)
                self.assertIn("prefers-reduced-motion:reduce", page)
                self.assertIn('aria-current="page"', page)
                self.assertIn('role="button"', page)
                self.assertIn('id="toastRegion"', page)
                self.assertIn("playTone", page)
                self.assertIn("showToast", page)

    # 3
    def test_csv_decodes_utf8_bom_and_cp1251(self):
        text = "Вилоят;Райгаз\n"
        self.assertEqual(app.decode_csv(b"\xef\xbb\xbf" + text.encode()), text)
        self.assertEqual(app.decode_csv(text.encode("cp1251")), text)

    # 4
    def test_csv_rejects_nul_bytes(self):
        with self.assertRaisesRegex(ValueError, "NUL"):
            app.decode_csv(b"a;b\x00c")

    # 5
    def test_semicolon_comma_and_tab_delimiters_are_detected(self):
        for delimiter in (";", ",", "\t"):
            data = self.csv_bytes([self.row()], delimiter=delimiter)
            with self.subTest(delimiter=repr(delimiter)):
                rows = self.parse_one(data)
                self.assertEqual(len(rows), 1)

    # 6
    def test_all_documented_date_formats_are_accepted(self):
        values = ("2026-09-23", "23.09.2026", "23/09/2026", "2026/09/23",
                  "23-09-2026", "2026.09.23", "2026-09-23T00:00:00")
        for value in values:
            with self.subTest(value=value):
                self.assertEqual(app.parse_date(value, "x.csv", 2, "sana").date().isoformat(), "2026-09-23")

    # 7
    def test_invalid_calendar_date_is_rejected_with_location(self):
        with self.assertRaisesRegex(ValueError, "x.csv: 7-qator"):
            app.parse_date("31.02.2026", "x.csv", 7, "sana")

    # 8
    def test_interval_boundaries_and_negative_days_are_strict(self):
        expected = {0: "0-30 (НАЗОРАТ)", 30: "0-30 (НАЗОРАТ)", 31: "31-35",
                    35: "31-35", 36: "36-40", 70: "66-70", 71: "71+"}
        self.assertEqual({day: app.interval_for(day) for day in expected}, expected)
        with self.assertRaisesRegex(ValueError, "Manfiy"):
            app.interval_for(-1)

    # 9
    def test_missing_required_column_is_rejected(self):
        data = self.csv_bytes([self.row()[:-1]], headers=app.REQUIRED_HEADERS[:-1])
        with self.assertRaisesRegex(ValueError, "kerakli ustun"):
            self.parse_one(data)

    # 10
    def test_normalized_duplicate_header_is_rejected(self):
        headers = [*app.REQUIRED_HEADERS, "ВИЛОЯТ"]
        data = self.csv_bytes([[*self.row(), "NAMANGAN"]], headers=headers)
        with self.assertRaisesRegex(ValueError, "bir necha marta"):
            self.parse_one(data)

    # 11
    def test_unquoted_extra_csv_column_is_rejected(self):
        text = ";".join(app.REQUIRED_HEADERS) + "\n" + ";".join([*self.row(), "ORTIQCHA"]) + "\n"
        with self.assertRaisesRegex(ValueError, "sarlavha soniga mos emas"):
            self.parse_one(text.encode())

    # 12
    def test_blank_required_value_is_rejected(self):
        row = self.row()
        row[4] = ""
        with self.assertRaisesRegex(ValueError, "bo'sh majburiy maydon"):
            self.parse_one(self.csv_bytes([row]))

    # 13
    def test_mixed_snapshot_dates_are_rejected(self):
        data = self.csv_bytes([
            self.row(code="1", snapshot="2026-09-21"),
            self.row(code="2", snapshot="2026-09-22"),
        ])
        with self.assertRaisesRegex(ValueError, "barcha fayllarda bir xil emas"):
            self.parse_one(data)

    # 14
    def test_byte_identical_source_files_are_not_double_counted(self):
        data = self.csv_bytes([self.row()])
        uploads = [app.MemoryUpload("a.csv", data), app.MemoryUpload("b.csv", data)]
        with self.assertRaisesRegex(ValueError, "aynan bir xil"):
            app.parse_files(uploads)

    # 15
    def test_sales_need_total_uses_exact_decimal_arithmetic(self):
        old = self.parse_one(self.csv_bytes([
            self.row(code="1", need="0.1"), self.row(code="2", need="0.2")
        ]), "old.csv")
        new = self.parse_one(self.csv_bytes([
            self.row(code="1", last="2026-09-21", snapshot="2026-09-22", need="0.1"),
            self.row(code="2", last="2026-09-21", snapshot="2026-09-22", need="0.2"),
        ]), "new.csv")
        result = app.compare_sales_rows(old, new)
        self.assertEqual(result["summary"][0]["need_total"], "0.3")

    # 16
    def test_changed_subscriber_metadata_is_sent_to_control(self):
        old = self.parse_one(self.csv_bytes([self.row()]), "old.csv")
        new = self.parse_one(self.csv_bytes([
            self.row(snapshot="2026-09-22", mahalla="B MFY", need="2")
        ]), "new.csv")
        result = app.compare_sales_rows(old, new)
        self.assertEqual(result["total_sales"], 0)
        self.assertIn("mahalla", result["controls"][0]["status"])
        self.assertIn("ehtiyoj", result["controls"][0]["status"])

    # 17
    def test_gas_report_rejects_duplicate_canonical_headers(self):
        data = self.gas_xlsx(
            ["Учреждение", "Инспектор", "Принял", "Қабул қилинди", "Реализовал", "Вернул"],
            ["A GAZ", "INSPEKTOR", 100, 100, 80, 20],
        )
        with self.assertRaisesRegex(ValueError, "ustun takrorlangan"):
            app.parse_gas_sales_xlsx(data, "gas.xlsx")

    # 18
    def test_gas_report_rejects_nonfinite_and_impossible_numbers(self):
        cases = [(["A GAZ", "INSPEKTOR", "NaN", 1, 0], "chekli son"),
                 (["A GAZ", "INSPEKTOR", 10, 11, 0], "dan katta")]
        for values, message in cases:
            with self.subTest(values=values):
                with self.assertRaisesRegex(ValueError, message):
                    app.parse_gas_sales_xlsx(self.gas_xlsx(values=values), "gas.xlsx")

    # 19
    def test_duplicate_gas_reports_are_not_double_counted(self):
        data = self.gas_xlsx()
        reports = [app.parse_gas_sales_xlsx(data, "a.xlsx"),
                   app.parse_gas_sales_xlsx(data, "b.xlsx")]
        with self.assertRaisesRegex(ValueError, "aynan bir xil"):
            app.ensure_unique_reports(reports, "Hisobotlar")

    # 20
    def test_reference_searches_later_sheets_and_requires_complete_totals(self):
        parsed = parse_reference_xlsx(self.reference_xlsx(), app.header_key)
        self.assertEqual(parsed["totals"]["31-35"], 1)
        with self.assertRaisesRegex(ValueError, "umumiy jami"):
            parse_reference_xlsx(self.reference_xlsx(blank_total_bucket=True), app.header_key)


if __name__ == "__main__":
    unittest.main()
