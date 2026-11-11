"""End-to-end tests using synthetic PDFs and no business data."""
import json
import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook, Workbook

from field_extractor import _append, extract_fields, load_template


def make_pdf(path, page_lines):
    # Minimal PDF generated in-process; no external PDF creation dependency.
    objects = [b"<< /Type /Catalog /Pages 2 0 R >>", None,
               b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"]
    page_ids = []
    for lines in page_lines:
        page_id, stream_id = len(objects) + 1, len(objects) + 2
        page_ids.append(page_id)
        objects.append(f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 3 0 R >> >> /Contents {stream_id} 0 R >>".encode())
        commands = ["BT /F1 12 Tf 50 730 Td 20 TL"]
        for line in lines:
            escaped = line.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")
            commands.append(f"({escaped}) Tj T*")
        commands.append("ET")
        stream = "\n".join(commands).encode("ascii")
        objects.append(f"<< /Length {len(stream)} >>\nstream\n".encode() + stream + b"\nendstream")
    kids = " ".join(f"{number} 0 R" for number in page_ids)
    objects[1] = f"<< /Type /Pages /Count {len(page_ids)} /Kids [{kids}] >>".encode()
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


class FieldExtractionTests(unittest.TestCase):
    def test_example_amount_does_not_silently_truncate_thousands(self):
        template = Path(__file__).resolve().parents[1] / '示例字段模板.json'
        amount = next((pattern, group) for name, pattern, group in load_template(template) if name == '参考金额')
        pattern, group = amount
        self.assertEqual(pattern.search('参考金额：1,234.56').group(group), '1,234.56')
        self.assertEqual(pattern.search('参考金额：-100.50').group(group), '-100.50')
        self.assertIsNone(pattern.search('参考金额：1,23.45'))

    def test_batch_and_review_workbook(self):
        with tempfile.TemporaryDirectory() as temp:
            base = Path(temp)
            source, destination = base / "input", base / "output"
            (source / "a").mkdir(parents=True)
            (source / "b").mkdir()
            destination.mkdir()
            make_pdf(source / "a" / "same.pdf", [["Invoice: ABC", "Invoice: DEF", "Value: =1+1"],
                                                    ["Invoice: GHI"], []])
            make_pdf(source / "b" / "same.pdf", [["Invoice: XYZ", "Value: 42"]])
            make_pdf(source / "=danger.pdf", [["Invoice: QRS"]])
            (source / "corrupt.pdf").write_bytes(b"not a pdf")
            template = base / "template.json"
            template.write_text(json.dumps({"fields": [
                {"name": "invoice", "pattern": r"Invoice:\s*(\S+)", "group": 1},
                {"name": "value", "pattern": r"Value:\s*(\S+)", "group": 1},
                {"name": "absent", "pattern": "DOESNOTEXIST"}]}), encoding="utf-8")
            progress = []
            result = extract_fields(source, destination, template, progress.append)
            self.assertEqual(result["pdf_count"], 4)
            self.assertEqual(result["match_count"], 7)
            self.assertEqual(len(progress), 4)
            workbook = load_workbook(result["workbook"], data_only=False)
            try:
                sheet = workbook["匹配结果"]
                rows = list(sheet.iter_rows(min_row=2, values_only=True))
                self.assertEqual({r[1] for r in rows if r[0] == str(Path("a") / "same.pdf") and r[2] == "invoice"}, {1, 2})
                self.assertEqual([r[5] for r in rows if r[0] == str(Path("a") / "same.pdf") and r[2] == "invoice"], [1, 2, 3])
                self.assertTrue(any(r[0] == str(Path("b") / "same.pdf") for r in rows))
                cells = [cell for row in sheet.iter_rows() for cell in row if cell.value in ("=1+1", "=danger.pdf")]
                self.assertEqual(len(cells), 2)
                self.assertTrue(all(cell.data_type == "s" for cell in cells))
                problems = list(workbook["待核对"].iter_rows(min_row=2, values_only=True))
                self.assertTrue(any(r[3] == "无可提取文字" and r[1] == 3 for r in problems))
                self.assertTrue(any(r[3] == "多次匹配" and r[2] == "invoice" for r in problems))
                self.assertTrue(any(r[3] == "无法打开" and r[0] == "corrupt.pdf" for r in problems))
                self.assertTrue(any(r[3] == "未找到（检索不完整）" and r[0] == "corrupt.pdf" for r in problems))
                self.assertEqual(result["warning_count"], len(problems))
            finally:
                workbook.close()
            self.assertNotEqual(result["output_dir"], extract_fields(source, destination, template)["output_dir"])

    def test_invalid_template_and_no_silent_excel_truncation(self):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "template.json"
            path.write_text(json.dumps({"fields": [{"name": "x", "pattern": "(abc)", "group": 2}]}))
            with self.assertRaisesRegex(ValueError, "group 不存在"):
                load_template(path)
            book = Workbook()
            with self.assertRaisesRegex(ValueError, "32767"):
                _append(book.active, ["a" * 32768])
            book.close()


if __name__ == "__main__":
    unittest.main()
