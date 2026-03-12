from django.test import TestCase, Client
from django.urls import reverse
from .utils import SOFormatter, ExcelParser, OrderRow, DumpExporter, GTMassAutomation
import pandas as pd
import io

class UtilsTestCase(TestCase):
    def test_so_formatter(self):
        self.assertEqual(SOFormatter.from_filename("SOGTM5985.xlsx"), "SO/GTM/5985")
        self.assertEqual(SOFormatter.from_filename("SOGTM5985"), "SO/GTM/5985")
        self.assertIsNone(SOFormatter.from_filename("NoNumbersHere.xlsx"))

    def test_clean_qty(self):
        self.assertEqual(ExcelParser._clean_qty("1,000"), 1000)
        self.assertEqual(ExcelParser._clean_qty("-"), 0)
        self.assertEqual(ExcelParser._clean_qty(""), 0)
        self.assertEqual(ExcelParser._clean_qty(None), 0)
        self.assertEqual(ExcelParser._clean_qty(15.5), 15)

    def test_excel_parser(self):
        # Create a mock Excel file in memory
        df = pd.DataFrame({
            "Ignore": ["A", "B"],
            "BC Code": [200453.0, 200173],
            "Order Qty": ["1,000", "-"]
        })

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        output.seek(0)

        parser = ExcelParser()
        rows = parser.parse(output, "SOGTM5985.xlsx")

        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0].item_no, "200453")
        self.assertEqual(rows[0].qty, 1000)
        self.assertEqual(rows[0].so_number, "SO/GTM/5985")

    def test_dump_exporter(self):
        exporter = DumpExporter()
        rows = [OrderRow(so_number="SO/GTM/5985", item_no="200453", qty=1000)]
        output = exporter.export_to_memory(rows)
        self.assertIsNotNone(output)

        # Read it back
        df = pd.read_excel(output)
        self.assertEqual(len(df), 1)
        self.assertEqual(df.iloc[0]["SO Number"], "SO/GTM/5985")
        self.assertEqual(str(df.iloc[0]["Item No"]), "200453")
        self.assertEqual(df.iloc[0]["Qty"], 1000)

class ViewsTestCase(TestCase):
    def setUp(self):
        self.client = Client()

    def test_index_view(self):
        response = self.client.get(reverse('index'))
        self.assertEqual(response.status_code, 200)
        self.assertTemplateUsed(response, 'offline/index.html')

    def test_process_files_view_no_files(self):
        response = self.client.post(reverse('process_files'))
        self.assertEqual(response.status_code, 400)
        self.assertJSONEqual(response.content, {"error": "No files selected"})
