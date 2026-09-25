import io
import subprocess
import sys
import unittest
import zipfile
from pathlib import Path

import app


ROOT = Path(__file__).resolve().parents[1]


class AppTests(unittest.TestCase):
    @staticmethod
    def zip_bytes(entries):
        output = io.BytesIO()
        with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
            for name, data in entries.items():
                archive.writestr(name, data)
        return output.getvalue()

    @staticmethod
    def snapshot_csv(snapshot_date, rows):
        lines = ["Вилоят;Райгаз;Маҳалла;Абонент код;Абонент;Еҳтиёж;Сўнги реализация;Бугунги реализация"]
        for code, last_sale, raygaz, mahalla, subscriber, need in rows:
            lines.append(
                f"НАМАНГАН;{raygaz};{mahalla};{code};{subscriber};{need};{last_sale};{snapshot_date}"
            )
        return ("\n".join(lines) + "\n").encode("utf-8")

    @staticmethod
    def snapshot_xlsx(snapshot_date, rows):
        output = io.BytesIO()
        workbook = app.xlsxwriter.Workbook(output, {"in_memory": True})
        sheet = workbook.add_worksheet("Holat")
        date_format = workbook.add_format({"num_format": "yyyy-mm-dd"})
        sheet.write(0, 0, "KUNLIK HOLAT")
        sheet.write_row(2, 0, app.REQUIRED_HEADERS)
        for index, (code, last_sale, raygaz, mahalla, subscriber, need) in enumerate(rows, start=3):
            sheet.write_row(index, 0, ["НАМАНГАН", raygaz, mahalla, code, subscriber, need])
            sheet.write_datetime(index, 6, app.datetime.strptime(last_sale, "%Y-%m-%d"), date_format)
            sheet.write_datetime(index, 7, app.datetime.strptime(snapshot_date, "%Y-%m-%d"), date_format)
        workbook.close()
        return output.getvalue()

    @staticmethod
    def egaz_sales_csv(rows):
        lines = ["Sana;Raygaz;Ariza №;Mahalla;Summa"]
        lines.extend(";".join(map(str, row)) for row in rows)
        return ("\n".join(lines) + "\n").encode("utf-8")

    @staticmethod
    def gnp_orders_xlsx(rows):
        output = io.BytesIO()
        workbook = app.xlsxwriter.Workbook(output, {"in_memory": True})
        sheet = workbook.add_worksheet("Buyurtmalar")
        sheet.write_row(0, 0, [
            "Ariza №", "Qabul qiluvchi (RAYGAZ)", "Summa", "Ballonlar soni",
            "O'tkazilgan ballonlar", "Yaratilgan",
        ])
        for row_index, row in enumerate(rows, start=1):
            sheet.write_row(row_index, 0, row)
        workbook.close()
        return output.getvalue()

    @staticmethod
    def gas_sales_xlsx(rows=None, period="1-31 Август холатига"):
        if rows is None:
            rows = [
                ("А туман ГАЗ", "BIRINCHI INSPEKTOR", 100, 80, 20),
                ("А туман ГАЗ", "IKKINCHI INSPEKTOR", 200, 150, 50),
            ]
        output = io.BytesIO()
        workbook = app.xlsxwriter.Workbook(output, {"in_memory": True})
        sheet = workbook.add_worksheet("Hisobot")
        sheet.write(3, 2, period)
        sheet.write_row(4, 0, [
            "Шахар-туман номи ва Чилангарларни Ф.И.Ш", "",
            "Е-газ дастурида қабул қилинган газ", "Е-газ дастурида сотилган газ",
            "Е-газ дастурида сотилмаган газ", "%",
        ])
        sheet.write_row(5, 0, ["Учреждение", "Инспектор", "Принял", "Реализовал", "Вернул", ""])
        row_index = 6
        district_totals = {}
        for organization, inspector, accepted, sold, returned in rows:
            sheet.write_row(row_index, 0, [
                organization, inspector, accepted, sold, returned,
                sold * 100 / accepted if accepted else 0,
            ])
            totals = district_totals.setdefault(organization, [0, 0, 0])
            totals[0] += accepted
            totals[1] += sold
            totals[2] += returned
            row_index += 1
        for organization, (accepted, sold, returned) in district_totals.items():
            sheet.write_row(row_index, 0, [
                organization, f"--- {organization} ЖАМИ ---", accepted, sold, returned,
                sold * 100 / accepted if accepted else 0,
            ])
            row_index += 1
        accepted = sum(row[2] for row in rows)
        sold = sum(row[3] for row in rows)
        returned = sum(row[4] for row in rows)
        sheet.write_row(row_index, 0, ["Жами ", "", accepted, sold, returned,
                                      sold * 100 / accepted if accepted else 0])
        workbook.close()
        return output.getvalue()

    def test_serverless_starts_without_writable_project_directory(self):
        code = (
            "import os, pathlib\n"
            "os.environ['VERCEL'] = '1'\n"
            "def read_only(*args, **kwargs): raise OSError('read-only filesystem')\n"
            "pathlib.Path.mkdir = read_only\n"
            "import app\n"
            "assert app.app.test_client().get('/').status_code == 200\n"
        )
        result = subprocess.run(
            [sys.executable, "-c", code], cwd=ROOT, capture_output=True, text=True
        )
        self.assertEqual(result.returncode, 0, result.stderr)

    def test_hosted_sample_download(self):
        client = app.app.test_client()
        page = client.get("/", base_url="https://davomilik.vercel.app")
        self.assertEqual(page.status_code, 200)
        self.assertIn("Vercel serverida qayta ishlanadi", page.get_data(as_text=True))

        sample = (ROOT / "NAMUNA.csv").read_bytes()
        result = client.post(
            "/generate",
            data={"files": (io.BytesIO(sample), "NAMUNA.csv")},
            base_url="https://davomilik.vercel.app",
        )
        self.assertEqual(result.status_code, 200, result.get_data(as_text=True) if result.is_json else "")
        self.assertEqual(result.headers["X-Total-Rows"], "2")
        with zipfile.ZipFile(io.BytesIO(result.data)) as workbook:
            self.assertIn("xl/workbook.xml", workbook.namelist())
            self.assertIn("Свод".encode(), workbook.read("xl/workbook.xml"))

    def test_hosted_upload_limit_returns_413(self):
        result = app.app.test_client().post(
            "/generate",
            data={"files": (io.BytesIO(b"x" * (5 * 1024 * 1024)), "large.csv")},
            base_url="https://davomilik.vercel.app",
        )
        self.assertEqual(result.status_code, 413)
        result.request.input_stream.close()
        result.close()

    def test_zip_with_nested_csv_is_accepted(self):
        sample = (ROOT / "NAMUNA.csv").read_bytes()
        zipped = self.zip_bytes({"kunlik/NAMUNA.csv": sample, "izoh.txt": b"ignored"})
        result = app.app.test_client().post(
            "/generate",
            data={"files": (io.BytesIO(zipped), "kunlik_fayllar.zip")},
        )
        self.assertEqual(result.status_code, 200, result.get_data(as_text=True) if result.is_json else "")
        self.assertEqual(result.headers["X-Total-Rows"], "2")

    def test_zip_without_csv_returns_clear_error(self):
        zipped = self.zip_bytes({"izoh.txt": b"CSV yo'q"})
        result = app.app.test_client().post(
            "/generate",
            data={"files": (io.BytesIO(zipped), "empty.zip")},
        )
        self.assertEqual(result.status_code, 400)
        self.assertIn("ZIP ichida CSV yoki XLSX fayl topilmadi", result.get_json()["error"])

    def test_expected_validation_errors_are_logged_without_tracebacks(self):
        with self.assertLogs(level="WARNING") as captured:
            result = app.app.test_client().post("/egaz/generate", data={})

        self.assertEqual(result.status_code, 400)
        joined = "\n".join(captured.output)
        self.assertIn("E-GAZ request rejected", joined)
        self.assertNotIn("Traceback", joined)

    def test_upload_page_accepts_zip_and_has_remove_controls(self):
        page = app.app.test_client().get("/").get_data(as_text=True)
        self.assertIn('accept=".csv,.xlsx,.zip', page)
        self.assertIn("Fayl yoki papkani shu yerga tashlang", page)
        self.assertIn('id="filePick"', page)
        self.assertIn('id="referencePick"', page)
        self.assertIn('id="toastRegion"', page)
        self.assertIn("window.AudioContext", page)
        self.assertIn("Olib tashlash", page)
        self.assertIn('id="folderInput"', page)
        self.assertIn("webkitdirectory", page)
        self.assertIn("collectDroppedFiles", page)
        self.assertIn('id="referenceDrop"', page)
        self.assertIn("referenceDrop.ondrop", page)

    def test_pages_hide_redundant_helper_text(self):
        client = app.app.test_client()
        main_page = client.get("/").get_data(as_text=True)
        sales_page = client.get("/sotuvlar").get_data(as_text=True)
        egaz_page = client.get("/egaz").get_data(as_text=True)

        self.assertNotIn('<div class="trust-chip">', main_page)
        self.assertNotIn("Manba fayllarni yuklang", main_page)
        self.assertNotIn("Abonent holati fayllari", main_page)
        self.assertNotIn(".xlsx faylni tanlang yoki shu blok ustiga tashlang", main_page)
        self.assertNotIn("Manba CSV fayllar o'zgartirilmaydi", main_page)
        self.assertNotIn("Qat'iy nazorat:", main_page)

        self.assertNotIn('<div class="mode-chip">', sales_page)
        self.assertNotIn("Avval eski, keyin yangi holat", sales_page)
        self.assertNotIn("Oldingi kunning to‘liq", sales_page)
        self.assertNotIn("Keyingi kunning to‘liq", sales_page)

        self.assertNotIn('<div class="mode-chip">', egaz_page)
        self.assertNotIn("E-GAZ Billing’dan olingan XLSX faylni yuklang", egaz_page)

    def test_generate_accepts_inspector_gas_sales_xlsx(self):
        result = app.app.test_client().post(
            "/generate",
            data={"files": (io.BytesIO(self.gas_sales_xlsx()), "avgust.xlsx")},
        )
        self.assertEqual(
            result.status_code,
            200,
            result.get_data(as_text=True) if result.is_json else "",
        )
        self.assertEqual(result.headers["X-Report-Type"], "gas-sales")
        self.assertEqual(result.headers["X-Total-Rows"], "2")
        self.assertEqual(result.headers["X-Raygaz-Count"], "1")
        self.assertEqual(result.headers["X-Inspector-Count"], "2")
        self.assertEqual(result.headers["X-Accepted-Total"], "300")
        self.assertEqual(result.headers["X-Sold-Total"], "230")
        self.assertEqual(result.headers["X-Returned-Total"], "70")
        with zipfile.ZipFile(io.BytesIO(result.data)) as workbook:
            names = workbook.read("xl/workbook.xml")
            self.assertIn("Свод".encode(), names)
            self.assertIn("Райгаз".encode(), names)
            self.assertIn("Чилангарлар".encode(), names)
            self.assertIn("Назорат".encode(), names)

    def test_gas_sales_parser_calculates_unsold_from_accepted_minus_sold(self):
        source = self.gas_sales_xlsx([
            ("Андижон туман ГАЗ", "INSPEKTOR", 3063, 2374, 0),
        ], "Yangi holat")
        report = app.parse_gas_sales_xlsx(source, "yangi.xlsx")
        self.assertIsNotNone(report)
        self.assertEqual(report["rows"][0]["returned"], 689)

    def test_sales_page_is_separate_and_linked_from_main_page(self):
        client = app.app.test_client()
        main_page = client.get("/").get_data(as_text=True)
        sales_page = client.get("/sotuvlar").get_data(as_text=True)
        self.assertIn('href="/sotuvlar"', main_page)
        self.assertIn("Kechagi va yangi holatni solishtirish", sales_page)
        self.assertIn('id="oldDrop"', sales_page)
        self.assertIn('id="newDrop"', sales_page)
        self.assertIn(".csv,.xlsx,.zip", sales_page)
        self.assertIn('id="summaryHead"', sales_page)
        self.assertIn('id="summarySection"', sales_page)
        self.assertIn("summarySection').hidden=true", sales_page)
        self.assertIn('id="detailsTableWrap"', sales_page)
        self.assertIn('id="detailsTable"', sales_page)
        self.assertIn("classList.add('gas-combined-wrap')", sales_page)
        self.assertIn("classList.add('gas-combined-table')", sales_page)
        self.assertIn('id="downloadExcel"', sales_page)
        self.assertIn('id="controlSection"', sales_page)
        self.assertIn("controlSection').hidden=true", sales_page)

    def test_all_pages_use_the_same_header_width(self):
        client = app.app.test_client()
        for path in ("/", "/sotuvlar", "/egaz"):
            with self.subTest(path=path):
                page = client.get(path).get_data(as_text=True)
                header_rules = [
                    block.split("}", 1)[0]
                    for block in page.split(".top-inner{")[1:]
                ]
                self.assertTrue(any("max-width:1180px" in rule for rule in header_rules))
                self.assertFalse(any("max-width:1500px" in rule for rule in header_rules))

    def test_navigation_buttons_are_separate_and_color_coded(self):
        client = app.app.test_client()
        for path in ("/", "/sotuvlar", "/egaz"):
            with self.subTest(path=path):
                page = client.get(path).get_data(as_text=True)
                self.assertIn("gap:12px", page)
                self.assertIn("background:#2563eb", page)
                self.assertIn("background:#d97706", page)
                self.assertIn("background:#15803d", page)
                self.assertIn("border-color:#fff", page)

    def test_gnp_and_mfy_svody_are_one_page(self):
        client = app.app.test_client()
        response = client.get("/mfy-svod", follow_redirects=False)
        self.assertEqual(response.status_code, 302)
        self.assertEqual(response.headers["Location"], "/gnp-taqqoslash")

        page = client.get("/gnp-taqqoslash").get_data(as_text=True)
        self.assertIn("GNP + E-GAZ + MFY SVOD", page)
        self.assertIn("GNP + MFY SVOD", page)
        self.assertIn("GNP buyurtmalari", page)
        self.assertNotIn('href="/mfy-svod"', page)

    def test_upload_panels_use_compact_responsive_layout(self):
        client = app.app.test_client()
        main_page = client.get("/").get_data(as_text=True)
        sales_page = client.get("/sotuvlar").get_data(as_text=True)
        egaz_page = client.get("/egaz").get_data(as_text=True)

        self.assertEqual(main_page.count('class="drop-symbol"'), 1)
        self.assertEqual(sales_page.count('class="drop-symbol"'), 2)
        self.assertEqual(egaz_page.count('class="drop-symbol"'), 1)
        self.assertIn("grid-template-columns:48px minmax(0,1fr) auto", main_page)
        self.assertIn("grid-template-columns:44px minmax(0,1fr) auto", sales_page)
        self.assertIn("#newDrop .drop-symbol", sales_page)
        self.assertIn("display:flex;align-items:center;justify-content:center", egaz_page)
        self.assertIn(".drop-meta{justify-content:center}", egaz_page)
        self.assertIn("justify-content:space-between", main_page)

    def test_upload_sections_have_distinct_color_accents(self):
        client = app.app.test_client()
        main_page = client.get("/").get_data(as_text=True)
        sales_page = client.get("/sotuvlar").get_data(as_text=True)
        egaz_page = client.get("/egaz").get_data(as_text=True)

        self.assertIn("border-color:#74acd2", main_page)
        self.assertIn("border-left:4px solid #d99a12", main_page)
        self.assertIn(".snapshot:first-child{border-color:#b8d8ee", sales_page)
        self.assertIn(".snapshot:last-child{border-color:#b7dfcc", sales_page)
        self.assertIn("#oldDrop{border-color:#79add1", sales_page)
        self.assertIn("#newDrop{border-color:#75b997", sales_page)
        self.assertIn("border:2px dashed #69b690", egaz_page)

    def test_egaz_page_is_linked_and_generates_picture_style_workbook(self):
        client = app.app.test_client()
        main_page = client.get("/").get_data(as_text=True)
        sales_page = client.get("/sotuvlar").get_data(as_text=True)
        egaz_page = client.get("/egaz").get_data(as_text=True)
        self.assertIn('href="/egaz"', main_page)
        self.assertIn('href="/egaz"', sales_page)
        self.assertIn("E-GAZ hisobotini tayyorlash", egaz_page)
        self.assertIn('id="drop"', egaz_page)
        self.assertIn('id="generate"', egaz_page)

        source = self.gas_sales_xlsx([
            ("Рђ С‚СѓРјР°РЅ Р“РђР—", "BIRINCHI INSPEKTOR", 100, 80, 0),
            ("Рђ С‚СѓРјР°РЅ Р“РђР—", "IKKINCHI INSPEKTOR", 200, 150, 0),
        ])
        result = client.post(
            "/egaz/generate",
            data={"file": (
                io.BytesIO(source),
                "E-GAZ Billing Hisobot - 2026-09-23T120816.895.xlsx",
            )},
        )
        self.assertEqual(result.status_code, 200, result.get_data(as_text=True) if result.is_json else "")
        self.assertEqual(result.headers["X-Output-Name"], "E-GAZ_1-23_Sentyabr_holatiga.xlsx")
        self.assertEqual(result.headers["X-Egaz-Rows"], "2")
        self.assertEqual(result.headers["X-Egaz-Accepted"], "300")
        self.assertEqual(result.headers["X-Egaz-Sold"], "230")
        self.assertEqual(result.headers["X-Egaz-Unsold"], "70")
        with zipfile.ZipFile(io.BytesIO(result.data)) as workbook:
            self.assertIn("1-23 Сентябрь".encode(), workbook.read("xl/workbook.xml"))
            strings = workbook.read("xl/sharedStrings.xml")
            self.assertIn("М А Ъ Л У М О Т".encode(), strings)
            self.assertIn("Жами".encode(), strings)

    def test_egaz_generate_accepts_readable_period_filename(self):
        client = app.app.test_client()
        source = self.gas_sales_xlsx()

        for filename in ("1-31 Avgust.xlsx", "1-31 Avgyst.xlsx"):
            with self.subTest(filename=filename):
                result = client.post(
                    "/egaz/generate",
                    data={"file": (io.BytesIO(source), filename)},
                )
                self.assertEqual(
                    result.status_code,
                    200,
                    result.get_data(as_text=True) if result.is_json else "",
                )
                self.assertEqual(
                    result.headers["X-Output-Name"],
                    "E-GAZ_1-31_Avgust_holatiga.xlsx",
                )
                cyrillic_period = "1-31 \u0410\u0432\u0433\u0443\u0441\u0442 \u04b3\u043e\u043b\u0430\u0442\u0438\u0433\u0430"
                with zipfile.ZipFile(io.BytesIO(result.data)) as workbook:
                    self.assertIn(
                        "1-31 \u0410\u0432\u0433\u0443\u0441\u0442".encode(),
                        workbook.read("xl/workbook.xml"),
                    )
                    self.assertIn(
                        cyrillic_period.encode(),
                        workbook.read("xl/sharedStrings.xml"),
                    )
                    self.assertNotIn(
                        "1-31 Avgust holatiga".encode(),
                        workbook.read("xl/sharedStrings.xml"),
                    )

    def test_egaz_page_explains_both_filename_formats(self):
        page = app.app.test_client().get("/egaz").get_data(as_text=True)
        self.assertIn("YYYY-MM-DD", page)
        self.assertIn("1-31 Avgust.xlsx", page)

    def test_egaz_period_filename_variants_and_invalid_ranges(self):
        cyrillic_suffix = "\u04b3\u043e\u043b\u0430\u0442\u0438\u0433\u0430"
        cases = {
            "1-29  May.xlsx": (
                f"1-29 {app.EGAZ_MONTHS_CYRILLIC[5]} {cyrillic_suffix}",
                "E-GAZ_1-29_May_holatiga.xlsx",
            ),
            "1 - 30 sentabr holatiga.xlsx": (
                f"1-30 {app.EGAZ_MONTHS_CYRILLIC[9]} {cyrillic_suffix}",
                "E-GAZ_1-30_Sentyabr_holatiga.xlsx",
            ),
            "1_31 август.xlsx": (
                f"1-31 {app.EGAZ_MONTHS_CYRILLIC[8]} {cyrillic_suffix}",
                "E-GAZ_1-31_Avgust_holatiga.xlsx",
            ),
        }
        for filename, (title, output_name) in cases.items():
            with self.subTest(filename=filename):
                period = app.egaz_period_from_filename(filename)
                self.assertEqual(period["title"], title)
                self.assertEqual(period["filename"], output_name)

        for filename in ("1-31 Aprel.xlsx", "hisobot.xlsx"):
            with self.subTest(filename=filename):
                with self.assertRaises(ValueError):
                    app.egaz_period_from_filename(filename)

    def test_sales_comparison_accepts_inspector_gas_sales_xlsx(self):
        old_xlsx = self.gas_sales_xlsx([
            ("А туман ГАЗ", "BIRINCHI INSPEKTOR", 100, 80, 20),
            ("А туман ГАЗ", "IKKINCHI INSPEKTOR", 200, 150, 50),
        ], "1-30 Август холатига")
        new_xlsx = self.gas_sales_xlsx([
            ("А туман ГАЗ", "BIRINCHI INSPEKTOR", 110, 90, 20),
            ("А туман ГАЗ", "IKKINCHI INSPEKTOR", 200, 150, 50),
            ("Б туман ГАЗ", "UCHINCHI INSPEKTOR", 50, 20, 30),
        ], "1-31 Август холатига")
        result = app.app.test_client().post(
            "/sotuvlar/compare",
            data={
                "old_files": (io.BytesIO(old_xlsx), "old.xlsx"),
                "new_files": (io.BytesIO(new_xlsx), "new.xlsx"),
            },
        )
        self.assertEqual(result.status_code, 200, result.get_data(as_text=True))
        data = result.get_json()
        self.assertEqual(data["report_type"], "gas-sales-summary")
        self.assertEqual(data["old_period"], "1-30 Август холатига")
        self.assertEqual(data["new_period"], "1-31 Август холатига")
        self.assertEqual(data["accepted_delta"], 60)
        self.assertEqual(data["sold_delta"], 30)
        self.assertEqual(data["returned_delta"], 30)
        self.assertEqual(data["changed_count"], 2)
        self.assertEqual(data["raygaz_count"], 2)
        self.assertEqual(len(data["summary"]), 2)
        summary = {item["raygaz"]: item for item in data["summary"]}
        self.assertEqual(summary["А туман ГАЗ"]["returned_delta"], 0)
        self.assertEqual(summary["Б туман ГАЗ"]["old_returned"], 0)
        self.assertEqual(summary["Б туман ГАЗ"]["new_returned"], 30)
        self.assertEqual(summary["Б туман ГАЗ"]["returned_delta"], 30)
        comparison_rows = data["comparison_rows"]
        self.assertEqual(comparison_rows[0]["inspector"], "BIRINCHI INSPEKTOR")
        self.assertEqual(comparison_rows[0]["accepted"], 100)
        self.assertEqual(comparison_rows[0]["sold"], 80)
        self.assertEqual(comparison_rows[0]["returned"], 20)
        a_total = next(
            item for item in comparison_rows
            if item["is_total"] and item["organization"] == "А туман ГАЗ"
        )
        self.assertEqual(a_total["returned"], 70)
        self.assertEqual(a_total["returned_delta"], 0)
        grand_total = next(item for item in comparison_rows if item["is_grand_total"])
        self.assertEqual(grand_total["accepted"], 300)
        self.assertEqual(grand_total["sold"], 230)
        self.assertEqual(grand_total["returned"], 70)
        self.assertEqual(grand_total["returned_delta"], 30)
        self.assertEqual({item["inspector"] for item in data["details"]}, {
            "BIRINCHI INSPEKTOR", "UCHINCHI INSPEKTOR",
        })
        self.assertEqual(data["controls"], [{
            "status": "Yangi holatda qo‘shilgan",
            "organization": "Б туман ГАЗ",
            "inspector": "UCHINCHI INSPEKTOR",
        }])

    def test_gas_sales_comparison_xlsx_keeps_yesterday_rows_and_writes_delta(self):
        old_xlsx = self.gas_sales_xlsx([
            ("Андижон шахар ГАЗ", "INSPEKTOR", 1400, 0, 1400),
        ], "1-30 Август холатига")
        new_xlsx = self.gas_sales_xlsx([
            ("Андижон шахар ГАЗ", "INSPEKTOR", 1000, 0, 0),
        ], "1-31 Август холатига")
        result = app.app.test_client().post(
            "/sotuvlar/compare-xlsx",
            data={
                "old_files": (io.BytesIO(old_xlsx), "old.xlsx"),
                "new_files": (io.BytesIO(new_xlsx), "new.xlsx"),
            },
        )
        self.assertEqual(
            result.status_code,
            200,
            result.get_data(as_text=True) if result.is_json else "",
        )
        self.assertTrue(result.headers["X-Output-Name"].startswith("QABUL_SOTUV_SOLISHTIRISH_"))
        with zipfile.ZipFile(io.BytesIO(result.data)) as workbook:
            self.assertIn("Солиштириш".encode(), workbook.read("xl/workbook.xml"))
            shared = app.ET.fromstring(workbook.read("xl/sharedStrings.xml"))
            shared_strings = [
                "".join(item.itertext()) for item in shared.findall("x:si", app.XLSX_NS)
            ]
            sheet = app.ET.fromstring(workbook.read("xl/worksheets/sheet1.xml"))
            period_cell = sheet.find(".//x:c[@r='C4']", app.XLSX_NS)
            self.assertEqual(
                shared_strings[int(period_cell.find("x:v", app.XLSX_NS).text)],
                "1-30 Август холатига",
            )
            cell_values = {
                cell.get("r"): cell.find("x:v", app.XLSX_NS).text
                for cell in sheet.findall(".//x:c", app.XLSX_NS)
                if cell.find("x:v", app.XLSX_NS) is not None
            }
            self.assertEqual(cell_values["E7"], "1400")
            self.assertEqual(cell_values["E8"], "1400")
            self.assertEqual(cell_values["G8"], "-400")
            self.assertEqual(cell_values["C9"], "1400")
            self.assertEqual(cell_values["E9"], "1400")
            self.assertEqual(cell_values["G9"], "-400")
            column_g_values = [
                cell.find("x:v", app.XLSX_NS).text
                for cell in sheet.findall(".//x:c", app.XLSX_NS)
                if cell.get("r", "").startswith("G") and cell.find("x:v", app.XLSX_NS) is not None
            ]
            column_e_values = [
                cell.find("x:v", app.XLSX_NS).text
                for cell in sheet.findall(".//x:c", app.XLSX_NS)
                if cell.get("r", "").startswith("E") and cell.find("x:v", app.XLSX_NS) is not None
            ]
            self.assertIn("1400", column_e_values)
            self.assertIn("-400", column_g_values)

    def test_sales_comparison_reports_confirmed_sales_and_control_differences(self):
        old_csv = self.snapshot_csv("2026-09-21", [
            ("1001", "2026-08-01", "A RAYGAZ", "A MFY", "BIRINCHI", "2"),
            ("1002", "2026-08-10", "A RAYGAZ", "B MFY", "IKKINCHI", "1"),
            ("1003", "2026-09-01", "B RAYGAZ", "C MFY", "UCHINCHI", "1"),
        ])
        new_csv = self.snapshot_csv("2026-09-22", [
            ("1001", "2026-09-21", "A RAYGAZ", "A MFY", "BIRINCHI", "2"),
            ("1002", "2026-08-10", "A RAYGAZ", "B MFY", "IKKINCHI", "1"),
            ("1004", "2026-09-02", "C RAYGAZ", "D MFY", "TORTINCHI", "1"),
        ])
        result = app.app.test_client().post(
            "/sotuvlar/compare",
            data={
                "old_files": (io.BytesIO(old_csv), "old.csv"),
                "new_files": (io.BytesIO(new_csv), "new.csv"),
            },
        )
        self.assertEqual(result.status_code, 200, result.get_data(as_text=True))
        data = result.get_json()
        self.assertEqual(data["old_date"], "2026-09-21")
        self.assertEqual(data["new_date"], "2026-09-22")
        self.assertEqual(data["total_sales"], 1)
        self.assertEqual(data["missing_count"], 1)
        self.assertEqual(data["added_count"], 1)
        self.assertEqual(data["details"][0]["code"], "1001")
        self.assertEqual(data["summary"], [{"raygaz": "A RAYGAZ", "sales_count": 1, "need_total": "2"}])
        self.assertEqual({item["code"] for item in data["controls"]}, {"1003", "1004"})

    def test_sales_comparison_rejects_swapped_snapshot_dates(self):
        old_csv = self.snapshot_csv("2026-09-22", [
            ("1001", "2026-08-01", "A RAYGAZ", "A MFY", "BIRINCHI", "1"),
        ])
        new_csv = self.snapshot_csv("2026-09-21", [
            ("1001", "2026-08-01", "A RAYGAZ", "A MFY", "BIRINCHI", "1"),
        ])
        result = app.app.test_client().post(
            "/sotuvlar/compare",
            data={
                "old_files": (io.BytesIO(old_csv), "old.csv"),
                "new_files": (io.BytesIO(new_csv), "new.csv"),
            },
        )
        self.assertEqual(result.status_code, 400)
        self.assertIn("keyin bo‘lishi kerak", result.get_json()["error"])

    def test_sales_comparison_rejects_duplicate_subscriber_code(self):
        old_csv = self.snapshot_csv("2026-09-21", [
            ("1001", "2026-08-01", "A RAYGAZ", "A MFY", "BIRINCHI", "1"),
            ("1001", "2026-08-02", "A RAYGAZ", "B MFY", "BIRINCHI", "1"),
        ])
        new_csv = self.snapshot_csv("2026-09-22", [
            ("1001", "2026-09-21", "A RAYGAZ", "A MFY", "BIRINCHI", "1"),
        ])
        result = app.app.test_client().post(
            "/sotuvlar/compare",
            data={
                "old_files": (io.BytesIO(old_csv), "old.csv"),
                "new_files": (io.BytesIO(new_csv), "new.csv"),
            },
        )
        self.assertEqual(result.status_code, 400)
        self.assertIn("abonent kodi takrorlangan", result.get_json()["error"])

    def test_sales_comparison_accepts_excel_xlsx_files(self):
        old_xlsx = self.snapshot_xlsx("2026-09-21", [
            ("2001", "2026-08-01", "EXCEL RAYGAZ", "EXCEL MFY", "EXCEL ABONENT", "2"),
        ])
        new_xlsx = self.snapshot_xlsx("2026-09-22", [
            ("2001", "2026-09-22", "EXCEL RAYGAZ", "EXCEL MFY", "EXCEL ABONENT", "2"),
        ])
        result = app.app.test_client().post(
            "/sotuvlar/compare",
            data={
                "old_files": (io.BytesIO(old_xlsx), "old.xlsx"),
                "new_files": (io.BytesIO(new_xlsx), "new.xlsx"),
            },
        )
        self.assertEqual(result.status_code, 200, result.get_data(as_text=True))
        data = result.get_json()
        self.assertEqual(data["total_sales"], 1)
        self.assertEqual(data["details"][0]["code"], "2001")

    def test_gnp_comparison_report_includes_same_day_checks_and_all_mfy_summaries(self):
        csv_data = self.egaz_sales_csv([
            ("2026-09-24", "A RAYGAZ", "1001", "Yangiobod MFY", "40000"),
            ("2026-09-24", "A RAYGAZ", "1001", "Yangiobod MFY", "40000"),
            ("2026-09-25", "B RAYGAZ", "1002", "B MFY", "40000"),
            ("2026-09-24", "A RAYGAZ", "1004", "Yangiobod MFY", "40000"),
            ("2026-09-24", "B RAYGAZ", "1004", "B MFY", "40000"),
        ])
        gnp_data = self.gnp_orders_xlsx([
            ("1001", "A RAYGAZ", 80000, 2, 2, "24.09.2026"),
            ("1002", "B RAYGAZ", 40000, 1, 1, "24.09.2026"),
            ("1003", "C RAYGAZ", 40000, 1, 1, "24.09.2026"),
            ("1004", "A RAYGAZ", 80000, 2, 2, "24.09.2026"),
            ("1005", "D RAYGAZ", 0, 0, 0, "24.09.2026"),
        ])
        response = app.app.test_client().post(
            "/gnp-taqqoslash/generate",
            data={
                "csv_file": [
                    (io.BytesIO(self.egaz_sales_csv([
                        ("2026-09-24", "A RAYGAZ", "1001", "Yangiobod MFY", "40000"),
                        ("2026-09-24", "A RAYGAZ", "1001", "Yangiobod MFY", "40000"),
                        ("2026-09-25", "B RAYGAZ", "1002", "B MFY", "40000"),
                    ])), "district_a.csv"),
                    (io.BytesIO(self.egaz_sales_csv([
                        ("2026-09-24", "A RAYGAZ", "1004", "Yangiobod MFY", "40000"),
                        ("2026-09-24", "B RAYGAZ", "1004", "B MFY", "40000"),
                    ])), "district_b.csv"),
                ],
                "gnp_file": (io.BytesIO(gnp_data), "gnp.xlsx"),
            },
        )
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.headers["X-Mismatches"], "3")
        self.assertEqual(response.headers["X-Matched"], "1")
        self.assertEqual(response.headers["X-Zero-Transfer"], "1")
        self.assertNotIn("X-Source-Amount", response.headers)
        self.assertEqual(response.headers["X-Source-Rows"], "5")

        records = app.parse_egaz_sales_csv(csv_data, "egaz.csv")
        sales = app.aggregate_egaz_sales(records)
        orders = app.parse_gnp_orders_xlsx(gnp_data, "gnp.xlsx")
        comparison = app.compare_egaz_with_gnp(sales, orders)
        by_number = {row["order_no"]: row for row in comparison["orders"]}
        self.assertEqual(by_number["1001"]["status"], "MOS")
        self.assertEqual(by_number["1001"]["csv_mahallas"], ["Yangiobod MFY"])
        self.assertEqual(by_number["1001"]["mahalla_match"], "GNPda MFY yo'q - E-GAZ MFYlari ko'rsatildi")
        self.assertEqual(by_number["1002"]["status"], "SANA MOS EMAS")
        self.assertEqual(by_number["1002"]["other_day_sales"], 1)
        self.assertEqual(by_number["1003"]["status"], "E-GAZDA TOPILMADI")
        self.assertEqual(by_number["1004"]["status"], "TUMAN TAFOVUTI")
        self.assertEqual(by_number["1005"]["status"], "0 SOTUV - E-GAZDA YO'Q")
        self.assertEqual(comparison["matched"], 1)
        self.assertEqual(comparison["mismatches"], 3)
        self.assertEqual(comparison["zero_transfer_no_csv"], 1)
        self.assertEqual(comparison["same_day_found"], 2)

        with zipfile.ZipFile(io.BytesIO(response.data)) as archive:
            workbook_xml = app.ET.fromstring(archive.read("xl/workbook.xml"))
            sheet_names = {
                node.get("name") for node in workbook_xml.findall("x:sheets/x:sheet", app.XLSX_NS)
            }
            shared = "".join(
                "".join(node.itertext())
                for node in app.ET.fromstring(archive.read("xl/sharedStrings.xml")).findall("x:si", app.XLSX_NS)
            )
        self.assertTrue({"Umumiy", "Arizalar", "Tafovutlar", "0 sotuv", "E-GAZ tumanlar", "MFYlar", "A RAYGAZ", "B RAYGAZ"}.issubset(sheet_names))
        self.assertIn("E-GAZ CSV MFY(lar)i", shared)
        self.assertIn("MFY tekshiruvi", shared)
        self.assertIn("Soni farqi", shared)
        self.assertIn("E-GAZdagi noyob arizalar", shared)
        self.assertIn("GNPda topilmadi (ariza+sana)", shared)
        self.assertIn("MOS arizalar", shared)
        self.assertIn("GNP faylida MFY ustuni bo'lmasa", shared)
        self.assertIn("Yangiobod MFY", shared)
        self.assertIn("B MFY", shared)
        self.assertIn("district_a.csv", shared)
        self.assertIn("district_b.csv", shared)
        self.assertNotIn("Summa farqi", shared)
        self.assertNotIn("Jami summa", shared)

    def test_gnp_comparison_never_marks_zero_transfer_without_egaz_row_as_matched(self):
        csv_data = self.egaz_sales_csv([
            ("2026-09-23", "A RAYGAZ", "3001", "A MFY", "40000"),
        ])
        gnp_data = self.gnp_orders_xlsx([
            ("3001", "A RAYGAZ", 40000, 1, 1, "24.09.2026"),
            ("3002", "B RAYGAZ", 0, 0, 0, "24.09.2026"),
        ])
        sales = app.aggregate_egaz_sales(app.parse_egaz_sales_csv(csv_data, "egaz.csv"))
        comparison = app.compare_egaz_with_gnp(sales, app.parse_gnp_orders_xlsx(gnp_data, "gnp.xlsx"))
        by_number = {row["order_no"]: row for row in comparison["orders"]}
        self.assertEqual(by_number["3001"]["status"], "SANA MOS EMAS")
        self.assertEqual(by_number["3002"]["status"], "0 SOTUV - E-GAZDA YO'Q")
        self.assertEqual(comparison["matched"], 0)
        self.assertEqual(comparison["mismatches"], 1)
        self.assertEqual(comparison["zero_transfer_no_csv"], 1)
        _, summary = app.build_gnp_comparison_xlsx(sales, comparison, "egaz.csv", "gnp.xlsx")
        self.assertTrue(summary["date_status"].startswith("YO'Q"))

    def test_gnp_count_report_accepts_egaz_csv_without_sum_column(self):
        csv_data = (
            "Sana;Raygaz;Ariza №;Mahalla\n"
            "2026-09-24;A RAYGAZ;2001;A MFY\n"
        ).encode("utf-8")
        gnp_data = self.gnp_orders_xlsx([
            ("2001", "A RAYGAZ", 40000, 1, 1, "24.09.2026"),
        ])
        response = app.app.test_client().post(
            "/gnp-taqqoslash/generate",
            data={
                "csv_file": (io.BytesIO(csv_data), "egaz_no_sum.csv"),
                "gnp_file": (io.BytesIO(gnp_data), "gnp.xlsx"),
            },
        )
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.headers["X-Source-Rows"], "1")
        self.assertEqual(response.headers["X-Matched"], "1")

    def test_egaz_csv_portal_no_data_response_gets_a_clear_error(self):
        with self.assertRaisesRegex(ValueError, "E-GAZ portali.*Ma'lumot topilmadi"):
            app.parse_egaz_sales_csv(b"\xef\xbb\xbfMa'lumot topilmadi\n", "egaz.csv")

    def test_gnp_comparison_accepts_zip_with_district_csvs(self):
        header = "Sana;Raygaz;Ariza №;Mahalla;Summa\n"
        sales_zip = self.zip_bytes({
            "Andijon/north.csv": (header + "2026-09-24;A RAYGAZ;1001;A MFY;40000\n").encode("utf-8"),
            "Andijon/south.csv": (header + "2026-09-24;A RAYGAZ;1001;B MFY;40000\n").encode("utf-8"),
        })
        gnp_data = self.gnp_orders_xlsx([
            ("1001", "A RAYGAZ", 80000, 2, 2, "24.09.2026"),
        ])
        response = app.app.test_client().post(
            "/gnp-taqqoslash/generate",
            data={
                "csv_file": (io.BytesIO(sales_zip), "all_districts.zip"),
                "gnp_file": (io.BytesIO(gnp_data), "gnp.xlsx"),
            },
        )
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.headers["X-Source-Rows"], "2")
        self.assertEqual(response.headers["X-Matched"], "1")
        self.assertNotIn("X-Source-Amount", response.headers)


if __name__ == "__main__":
    unittest.main()
