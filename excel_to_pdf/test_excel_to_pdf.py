import unittest
import os
import tempfile
from pathlib import Path

# Import the functions
from excel_to_pdf import get_excel_files_from_folder, get_pdf_path

class TestExcelToPdf(unittest.TestCase):

    def test_get_pdf_path(self):
        excel_path = "C:\\path\\to\\file.xlsx"
        pdf_path = get_pdf_path(excel_path)
        self.assertEqual(pdf_path, "C:\\path\\to\\file.pdf")

        # Test with forward slashes
        excel_path = "/folder/test.xls"
        pdf_path = get_pdf_path(excel_path)
        expected = str(Path("/folder/test.pdf"))
        self.assertEqual(pdf_path, expected)

    def test_get_excel_files_from_folder_invalid(self):
        files = get_excel_files_from_folder("/invalid/path")
        self.assertEqual(files, [])

    def test_get_excel_files_from_folder_empty(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            files = get_excel_files_from_folder(temp_dir)
            self.assertEqual(files, [])