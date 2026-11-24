"""Exercise scanned PDF OCR orchestration with synthetic image PDFs."""
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from openpyxl import load_workbook
from PIL import Image
import pypdfium2 as pdfium

from scan_ocr import create_ocr_engine, recognize_image, render_page, run_scan_ocr


class StubEngine:
    def __init__(self):
        self.calls = 0

    def __call__(self, image, **kwargs):
        self.calls += 1
        if self.calls == 1:
            return SimpleNamespace(txts=('=2+2',), scores=(0.76,),
                                   boxes=[[[5, 6], [90, 6], [90, 28], [5, 28]]])
        if self.calls == 2:
            return SimpleNamespace(txts=None, scores=None, boxes=None)
        raise ValueError('synthetic OCR failure')


class ScanOcrTests(unittest.TestCase):
    def test_text_output_failure_counts_unprocessed_pages_and_continues(self):
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            source, output = base / 'input', base / 'output'
            source.mkdir()
            output.mkdir()
            image = Image.new('RGB', (100, 140), 'white')
            image.save(source / 'a.pdf', save_all=True, append_images=[image], resolution=72)
            image.save(source / 'b.pdf', resolution=72)
            image.close()
            original_open = Path.open

            def fail_one_text(path, *args, **kwargs):
                if path.name == 'a.pdf.txt':
                    raise PermissionError('simulated write error')
                return original_open(path, *args, **kwargs)

            with patch.object(Path, 'open', fail_one_text):
                result = run_scan_ocr(source, output, engine=StubEngine())
            self.assertEqual((result['pdf_count'], result['page_count'], result['ocr_count']), (2, 3, 1))
            self.assertEqual(result['issue_count'], 2)  # output failure plus low-score b.pdf
            book = load_workbook(result['workbook'], data_only=True)
            try:
                rows = {row[0]: row for row in list(book['文件状态'].values)[1:]}
                self.assertEqual(rows['a.pdf'][1:4], (2, 0, '需核对'))
                self.assertIn('逐页文本输出失败', rows['a.pdf'][4])
                self.assertIn('2 页未处理', rows['a.pdf'][4])
                self.assertEqual(rows['b.pdf'][1:3], (1, 1))
                page_rows = list(book['页面状态'].values)[1:]
                self.assertEqual([(row[0], row[1]) for row in page_rows], [('b.pdf', 1)])
            finally:
                book.close()
            self.assertTrue((result['text_dir'] / 'b.pdf.txt').is_file())

    def test_batch_preserves_page_failure_empty_page_and_formula_text(self):
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            source, output = base / 'input', base / 'output'
            source.mkdir()
            output.mkdir()
            image = Image.new('RGB', (100, 140), 'white')
            image.save(source / 'image.pdf', save_all=True, append_images=[image, image], resolution=72)
            image.close()
            (source / 'corrupt.pdf').write_bytes(b'broken')
            result = run_scan_ocr(source, output, engine=StubEngine())
            self.assertEqual((result['pdf_count'], result['page_count'], result['ocr_count']), (2, 3, 2))
            self.assertGreaterEqual(result['issue_count'], 4)
            book = load_workbook(result['workbook'], data_only=False)
            try:
                self.assertEqual(book['识别明细']['D2'].value, '=2+2')
                self.assertEqual(book['识别明细']['D2'].data_type, 's')
                self.assertIn('5.0,6.0', book['识别明细']['F2'].value)
                page_rows = list(book['页面状态'].values)[1:]
                self.assertIn('低分', page_rows[0][5])
                self.assertIn('未识别出文字', page_rows[1][5])
                self.assertEqual(page_rows[2][2], '识别失败')
                files = list(book['文件状态'].values)[1:]
                self.assertTrue(any(row[0] == 'corrupt.pdf' and '文件打开' in row[4] for row in files))
            finally:
                book.close()
            text = (result['text_dir'] / 'image.pdf.txt').read_text(encoding='utf-8')
            self.assertIn('本页未识别出文字', text)
            self.assertIn('第 3 页', text)
            self.assertIn('本页识别失败', text)

    def test_render_is_capped_and_output_cannot_be_inside_source(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            image = Image.new('RGB', (100, 140), 'white')
            image.save(root / 'sample.pdf', resolution=72)
            image.close()
            with pdfium.PdfDocument(root / 'sample.pdf') as document:
                rendered, info = render_page(document, 0, dpi=200, max_pixels=2000)
                self.assertLessEqual(rendered.width * rendered.height, 2000)
                self.assertTrue(info['limited'])
                rendered.close()
            with self.assertRaisesRegex(ValueError, '不能位于'):
                run_scan_ocr(root, root, engine=StubEngine())

    def test_model_missing_and_invalid_ocr_output(self):
        with tempfile.TemporaryDirectory() as tmp:
            with self.assertRaisesRegex(FileNotFoundError, '不会自动下载'):
                create_ocr_engine(tmp)
        class BadEngine:
            def __call__(self, *args, **kwargs):
                return SimpleNamespace(txts=('word',), scores=(), boxes=())
        with Image.new('RGB', (10, 10), 'white') as image:
            with self.assertRaisesRegex(ValueError, '数量不一致'):
                recognize_image(image, BadEngine())


if __name__ == '__main__':
    unittest.main()
