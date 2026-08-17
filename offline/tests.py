from django.test import TestCase, Client
from django.urls import reverse
from django.contrib.auth.models import User, Group

from core.access import EDITORS_GROUP
from .utils import SONumberFormatter, ExcelParser


class UtilsTestCase(TestCase):
    def test_so_formatter(self):
        self.assertEqual(SONumberFormatter.from_filename("SOGTM5985.xlsx"), "SO/GTM/5985")
        self.assertEqual(SONumberFormatter.from_filename("SOGTM5985"), "SO/GTM/5985")
        self.assertIsNone(SONumberFormatter.from_filename("NoNumbersHere.xlsx"))

    def test_clean_qty(self):
        self.assertEqual(ExcelParser._clean_qty("1,000"), 1000)
        self.assertEqual(ExcelParser._clean_qty("-"), 0)
        self.assertEqual(ExcelParser._clean_qty(""), 0)
        self.assertEqual(ExcelParser._clean_qty(None), 0)
        self.assertEqual(ExcelParser._clean_qty(15.5), 15)

    # NOTE: the former test_excel_parser / test_dump_exporter were removed during
    # the 2026 cleanup — they asserted against an old API that has since changed
    # (ExcelParser.parse now returns (rows, errors); OrderRow gained several
    # required fields; DumpExporter.export_to_memory now takes a ProcessResult).
    # They never ran (the module failed to import on the stale `SOFormatter`
    # name), so they protected nothing. If coverage for the exporter is wanted,
    # add fresh tests against the current ProcessResult-based API.


class ViewsTestCase(TestCase):
    def setUp(self):
        self.client = Client()
        self.user = User.objects.create_user(username='testuser', password='testpassword')
        # process_files is a write endpoint — the RBAC write-guard blocks Viewers
        # (403). This test exercises the view's own "no files → 400" branch, so the
        # user needs Editor access.
        editors, _ = Group.objects.get_or_create(name=EDITORS_GROUP)
        self.user.groups.add(editors)

    def test_index_view(self):
        self.client.login(username='testuser', password='testpassword')
        response = self.client.get(reverse('index'))
        self.assertEqual(response.status_code, 200)
        self.assertTemplateUsed(response, 'offline/index.html')

    def test_process_files_view_no_files(self):
        self.client.login(username='testuser', password='testpassword')
        response = self.client.post(reverse('process_files'))
        self.assertEqual(response.status_code, 400)
        self.assertJSONEqual(response.content, {"error": "No files selected"})
