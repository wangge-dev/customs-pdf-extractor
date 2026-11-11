"""End-to-end tests use generated documents only, never business records."""
import subprocess
import sys
import tempfile
import unittest
import os
from pathlib import Path
from unittest.mock import patch

from openpyxl import load_workbook
from pdfminer.pdfdocument import PDFPasswordIncorrect

from pdf_inventory import export_text_directory, inspect_directory, inspect_pdf, write_inventory


def write_pdf(path, objects):
    data = bytearray(b"%PDF-1.4\n")
    offsets = [0]
    for number, obj in enumerate(objects, 1):
        offsets.append(len(data))
        data.extend(f"{number} 0 obj\n".encode() + obj + b"\nendobj\n")
    start = len(data)
    data.extend(f"xref\n0 {len(offsets)}\n0000000000 65535 f \n".encode())
    for offset in offsets[1:]:
        data.extend(f"{offset:010d} 00000 n \n".encode())
    data.extend(f"trailer\n<< /Size {len(offsets)} /Root 1 0 R >>\nstartxref\n{start}\n%%EOF\n".encode())
    path.write_bytes(data)


def synthetic_pdf(path):
    """Generate a two-page PDF: one searchable page and one blank page."""
    stream = b"BT /F1 18 Tf 72 720 Td (Hello PDF) Tj ET"
    write_pdf(path, [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R 4 0 R] /Count 2 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 5 0 R >> >> /Contents 6 0 R >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] >>",
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
        b"<< /Length %d >>\nstream\n" % len(stream) + stream + b"\nendstream",
    ])


def image_pdf(path, width=550, height=700, label=b"1"):
    """A real embedded image plus searchable text, with configurable image size."""
    stream = (f"q {width} 0 0 {height} 30 40 cm /Im1 Do Q ".encode()
              + b"BT /F1 10 Tf 30 760 Td (" + label + b") Tj ET")
    write_pdf(path, [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> /XObject << /Im1 5 0 R >> >> /Contents 6 0 R >>",
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
        b"<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length 3 >>\nstream\n\x00\x00\x00\nendstream",
        b"<< /Length %d >>\nstream\n" % len(stream) + stream + b"\nendstream",
    ])


class InventoryTests(unittest.TestCase):
    def test_mixed_pdf_corrupt_recursive_and_excel(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = root / "input"
            source.mkdir()
            (source / "nested").mkdir()
            synthetic_pdf(source / "=formula.pdf")
            (source / "nested" / "broken.PDF").write_bytes(b"not a pdf")
            files, pages = inspect_directory(source, recursive=True)
            self.assertEqual(len(files), 2)
            good = next(row for row in files if row["文件名"] == "=formula.pdf")
            broken = next(row for row in files if row["文件名"] == "broken.PDF")
            self.assertEqual((good["页数"], good["有文字页数"], good["无文字页数"], good["可提取文字数"]), (2, 1, 1, 8))
            self.assertEqual(good["状态"], "存在无文字页")
            self.assertEqual(broken["状态"], "无法打开")
            self.assertEqual(len(pages), 2)
            self.assertEqual(pages[0]["宽度(mm)"], 215.9)
            output = root / "inventory.xlsx"
            write_inventory(output, files, pages)
            workbook = load_workbook(output, data_only=False)
            self.assertEqual(workbook.sheetnames, ["文件清单", "页面明细"])
            self.assertEqual(workbook["文件清单"]["A2"].value, "=formula.pdf")
            self.assertEqual(workbook["文件清单"]["A2"].data_type, "s")
            workbook.close()
            with self.assertRaises(FileExistsError):
                write_inventory(output, files, pages)
            self.assertEqual(len(inspect_directory(source)[0]), 1)

    def test_password_protected_is_not_reported_as_empty(self):
        with patch("pdf_inventory.pdfplumber.open", side_effect=PDFPasswordIncorrect):
            row, pages = inspect_pdf("protected.pdf")
        self.assertEqual(row["状态"], "密码保护")
        self.assertIsNone(row["页数"])
        self.assertEqual(pages, [])

    def test_large_image_with_searchable_page_number_is_flagged(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            image_pdf(root / "large.pdf")
            image_pdf(root / "small.pdf", width=55, height=70)
            image_pdf(root / "text.pdf", label=b"1234567890" * 4)
            large, pages = inspect_pdf(root / "large.pdf")
            self.assertEqual(large["有文字页数"], 1)
            self.assertEqual(large["状态"], "疑似图片主体，需核对")
            self.assertEqual(large["疑似图片主体页数"], 1)
            self.assertEqual(pages[0]["文字层"], "有")
            self.assertEqual(pages[0]["页面提示"], "疑似图片主体，需核对")
            self.assertEqual(inspect_pdf(root / "small.pdf")[0]["状态"], "正常")
            self.assertEqual(inspect_pdf(root / "text.pdf")[0]["状态"], "正常")
            output = root / "check.xlsx"
            write_inventory(output, [large], pages)
            workbook = load_workbook(output, read_only=True)
            self.assertEqual(workbook["文件清单"]["C2"].value, "疑似图片主体，需核对")
            self.assertEqual(workbook["页面明细"]["H2"].value, "疑似图片主体，需核对")
            workbook.close()

    def test_command_line(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = root / "input"
            source.mkdir()
            synthetic_pdf(source / "demo.pdf")
            result = subprocess.run([sys.executable, "-m", "pdf_inventory", str(source), str(root / "result.xlsx"),
                                     "--text-dir", str(root / "text")], capture_output=True, text=True,
                                    encoding="utf-8", env={**os.environ, "PYTHONUTF8": "1"})
            self.assertEqual(result.returncode, 0, result.stderr)
            workbook = load_workbook(root / "result.xlsx", read_only=True)
            self.assertEqual(workbook["文件清单"]["D2"].value, 2)
            self.assertEqual(workbook["文本导出"]["C2"].value, "存在无文字页")
            self.assertTrue(workbook["文本导出"]["G2"].value.endswith("demo.pdf.txt"))
            workbook.close()
            text = (root / "text" / "demo.pdf.txt").read_text(encoding="utf-8")
            self.assertIn("第 1 页", text)
            self.assertIn("Hello PDF", text)
            self.assertIn("第 2 页", text)
            self.assertIn("未进行 OCR", text)
            again = subprocess.run([sys.executable, "-m", "pdf_inventory", str(source), str(root / "again.xlsx"),
                                    "--text-dir", str(root / "text")], capture_output=True, text=True,
                                   encoding="utf-8", env={**os.environ, "PYTHONUTF8": "1"})
            self.assertEqual(again.returncode, 0, again.stderr)
            workbook = load_workbook(root / "again.xlsx", read_only=True)
            self.assertEqual(workbook["文本导出"]["C2"].value, "未覆盖已有文件")
            workbook.close()

    def test_text_export_same_names_and_no_overwrite(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = root / "input"
            for name in ("one", "two"):
                (source / name).mkdir(parents=True)
                synthetic_pdf(source / name / "same.pdf")
            output = root / "texts"
            result = export_text_directory(source, output, recursive=True)
            self.assertEqual(len(result), 2)
            self.assertTrue((output / "one" / "same.pdf.txt").is_file())
            self.assertTrue((output / "two" / "same.pdf.txt").is_file())
            self.assertTrue(all(item["无文字页数"] == 1 for item in result))
            second = export_text_directory(source, output, recursive=True)
            self.assertTrue(all(item["状态"] == "未覆盖已有文件" for item in second))


if __name__ == "__main__":
    unittest.main()
