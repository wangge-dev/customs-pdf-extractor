"""Use small generated vector PDFs, never real customs or business documents."""
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import load_workbook
import pdf_tables


def _pdf(path, streams):
    # Minimal valid PDF: Helvetica text and vector table lines.
    objects = [b'<< /Type /Catalog /Pages 2 0 R >>', None,
               b'<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>']
    refs = []
    for stream in streams:
        page_id = len(objects) + 1
        refs.append(page_id)
        objects.append((f'<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 700] '
                        f'/Resources << /Font << /F1 3 0 R >> >> /Contents {page_id + 1} 0 R >>').encode())
        raw = stream.encode('ascii')
        objects.append(b'<< /Length ' + str(len(raw)).encode() + b' >>\nstream\n' + raw + b'\nendstream')
    objects[1] = ('<< /Type /Pages /Count %d /Kids [%s] >>' %
                  (len(refs), ' '.join(f'{ref} 0 R' for ref in refs))).encode()
    data = bytearray(b'%PDF-1.4\n')
    offsets = [0]
    for index, item in enumerate(objects, 1):
        offsets.append(len(data))
        data.extend(f'{index} 0 obj\n'.encode() + item + b'\nendobj\n')
    xref = len(data)
    data.extend(f'xref\n0 {len(offsets)}\n0000000000 65535 f \n'.encode())
    for offset in offsets[1:]:
        data.extend(f'{offset:010d} 00000 n \n'.encode())
    data.extend(f'trailer\n<< /Size {len(offsets)} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n'.encode())
    path.write_bytes(data)


GRID = '''
0.8 w
40 500 m 240 500 l S
40 530 m 240 530 l S
40 560 m 240 560 l S
40 500 m 40 560 l S
140 500 m 140 560 l S
240 500 m 240 560 l S
BT /F1 11 Tf 50 540 Td (=2+2) Tj ET
BT /F1 11 Tf 150 540 Td (apple) Tj ET
BT /F1 11 Tf 50 510 Td (second) Tj ET
BT /F1 11 Tf 150 510 Td (42) Tj ET
'''


class TablesTest(unittest.TestCase):
    def test_real_extraction_multiple_pages_same_name_and_corrupt(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source, output = root / 'source', root / 'output'
            (source / 'a').mkdir(parents=True)
            (source / 'b').mkdir()
            _pdf(source / 'a' / 'same.pdf', [GRID, 'BT /F1 11 Tf 50 540 Td (   ) Tj ET'])
            _pdf(source / 'b' / 'same.PDF', [GRID])
            (source / 'broken.pdf').write_bytes(b'not a PDF')
            progress = []
            result = pdf_tables.run_tables(source, output, progress.append)
            self.assertEqual((result['pdf_count'], result['page_count'], result['table_count'], result['cell_count']), (3, 3, 2, 8))
            self.assertGreaterEqual(result['issue_count'], 2)
            self.assertEqual(len(progress), 3)
            wb = load_workbook(result['workbook'], data_only=False)
            rows = list(wb['单元格明细'].iter_rows(min_row=2, values_only=False))
            self.assertEqual({row[0].value for row in rows}, {'a/same.pdf', 'b/same.PDF'})
            self.assertEqual(rows[0][5].value, '=2+2')
            self.assertEqual(rows[0][5].data_type, 's')
            self.assertEqual([(c[1].value, c[2].value, c[3].value, c[4].value) for c in rows[:4]],
                             [(1, 1, 1, 1), (1, 1, 1, 2), (1, 1, 2, 1), (1, 1, 2, 2)])
            statuses = list(wb['逐页状态'].values)
            self.assertTrue(any(row[0] == 'a/same.pdf' and row[1] == 2 and '无可提取文字层' in row[5] for row in statuses[1:]))
            files = list(wb['文件状态'].values)
            self.assertTrue(any(row[0] == 'broken.pdf' and '文件处理失败' in row[4] for row in files[1:]))
            wb.close()

    def test_output_inside_input_is_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / 'source'
            source.mkdir()
            with self.assertRaisesRegex(ValueError, '结果保存目录'):
                pdf_tables.run_tables(source, source / 'result')
            self.assertFalse((source / 'result').exists())

    def test_empty_input_is_rejected_without_creating_output(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / 'source'
            source.mkdir()
            with self.assertRaisesRegex(ValueError, '未找到 PDF'):
                pdf_tables.run_tables(source, root / 'result')
            self.assertFalse((root / 'result').exists())

    def test_page_failure_is_reported(self):
        class BrokenPage:
            @property
            def chars(self):
                raise ValueError('invalid page')
            def close(self):
                pass

        class FakePdf:
            pages = [BrokenPage()]
            def __enter__(self):
                return self
            def __exit__(self, *args):
                pass

        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / 'source'
            source.mkdir()
            (source / 'brokenpage.pdf').touch()
            with patch('pdf_tables.pdfplumber.open', return_value=FakePdf()):
                result = pdf_tables.run_tables(source, root / 'output')
            wb = load_workbook(result['workbook'], data_only=True)
            self.assertEqual(wb['逐页状态']['E2'].value, '处理失败')
            self.assertIn('invalid page', wb['逐页状态']['F2'].value)
            wb.close()

    def test_partial_page_failure_does_not_export_incomplete_tables(self):
        class Table:
            def __init__(self, should_fail):
                self.should_fail = should_fail
            def extract(self):
                if self.should_fail:
                    raise ValueError('table extraction failed')
                return [['partial']]

        class Page:
            chars = [{'text': 'text'}]
            def find_tables(self):
                return [Table(False), Table(True)]
            def close(self):
                pass

        class Pdf:
            pages = [Page()]
            def __enter__(self):
                return self
            def __exit__(self, *args):
                pass

        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / 'source'
            source.mkdir()
            (source / 'partial.pdf').touch()
            with patch('pdf_tables.pdfplumber.open', return_value=Pdf()):
                result = pdf_tables.run_tables(source, root / 'output')
            self.assertEqual((result['table_count'], result['cell_count']), (0, 0))
            wb = load_workbook(result['workbook'], data_only=True)
            self.assertEqual(wb['单元格明细'].max_row, 1)
            self.assertEqual(wb['逐页状态']['E2'].value, '处理失败')
            wb.close()


if __name__ == '__main__':
    unittest.main()
