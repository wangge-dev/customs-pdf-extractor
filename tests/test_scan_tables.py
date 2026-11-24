"""扫描表格测试只使用运行时生成的虚构图片和 PDF。"""
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import load_workbook
from PIL import Image, ImageDraw

from scan_tables import assign_text, detect_grid_tables, run_scan_tables


def grid_image(two=False, broken=False):
    image = Image.new('RGB', (800, 950), 'white')
    draw = ImageDraw.Draw(image)
    # Each independent rectangle is a 2x2 table.
    for top in ([80, 500] if two else [80]):
        for y in (top, top + 100, top + 200):
            draw.line((80, y, 600, y), fill='black', width=3)
        for x in (80, 340, 600):
            if broken and x == 340:
                draw.line((x, top, x, top + 85), fill='black', width=3)
                draw.line((x, top + 125, x, top + 200), fill='black', width=3)
            else:
                draw.line((x, top, x, top + 200), fill='black', width=3)
    return image


def line(text, rect, score=.99):
    x0, y0, x1, y1 = rect
    return {'text': text, 'score': score, 'box': [[x0, y0], [x1, y0], [x1, y1], [x0, y1]]}


class ScanTablesTest(unittest.TestCase):
    def test_horizontal_only_candidate_is_exported_as_inferred(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source, output = root / 'input', root / 'output'
            source.mkdir()
            output.mkdir()
            image = Image.new('RGB', (800, 950), 'white')
            draw = ImageDraw.Draw(image)
            for y in (80, 180, 280):
                draw.line((80, y, 600, y), fill='black', width=3)
            image.save(source / 'horizontal.pdf', resolution=200)
            records = [line('A', (100, 100, 180, 130)), line('B', (380, 100, 440, 130)),
                       line('C', (100, 200, 180, 230)), line('D', (380, 200, 440, 230))]
            with patch('scan_ocr.recognize_image', return_value=records):
                result = run_scan_tables(source, output, engine=object(), dpi=200)
            self.assertEqual((result['table_count'], result['cell_count']), (1, 4))
            book = load_workbook(result['workbook'], data_only=True)
            try:
                cells = list(book['单元格明细'].values)[1:]
                self.assertEqual([row[5] for row in cells], ['A', 'B', 'C', 'D'])
                self.assertTrue(all('列按文字位置推测' in row[12] for row in cells))
                self.assertIn('列按文字位置推测', book['逐页状态']['G2'].value)
            finally:
                book.close()

    def test_many_page_warnings_preserve_all_page_details(self):
        class ManyPages:
            def __init__(self, path):
                pass
            def __len__(self):
                return 1000
            def close(self):
                pass

        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source, output = root / 'input', root / 'output'
            source.mkdir()
            output.mkdir()
            (source / 'many.pdf').touch()
            def blank_page(*args, **kwargs):
                return Image.new('RGB', (40, 40), 'white'), {'limited': False, 'dpi': 200}
            with patch('scan_tables.pdfium.PdfDocument', ManyPages), \
                 patch('scan_ocr.render_page', side_effect=blank_page), \
                 patch('scan_ocr.recognize_image', return_value=[]), \
                 patch('scan_tables.detect_grid_tables', return_value=[]), \
                 patch('scan_table_fallback.detect_horizontal_tables', return_value=[]):
                result = run_scan_tables(source, output, engine=object())
            self.assertEqual(result['page_count'], 1000)
            self.assertEqual(result['warning_count'], 1000)
            book = load_workbook(result['workbook'], data_only=True)
            try:
                self.assertEqual(book['文件状态']['E2'].value, '异常较多，请查看逐页状态')
                pages = book['逐页状态']
                self.assertEqual(pages.max_row, 1001)
                self.assertIn('整页未识别出文字', pages['G2'].value)
                self.assertEqual(pages['B1001'].value, 1000)
                self.assertIn('整页未识别出文字', pages['G1001'].value)
            finally:
                book.close()

    def test_two_separate_grids_and_empty_cells(self):
        tables = detect_grid_tables(grid_image(two=True))
        self.assertEqual(len(tables), 2)
        self.assertEqual([len(table['cells']) for table in tables], [4, 4])
        outside = assign_text(tables, [line('A', (100, 100, 180, 130)),
                                       line('B', (380, 210, 440, 240)),
                                       line('C', (120, 520, 190, 550)),
                                       line('outside', (20, 20, 100, 40))])
        self.assertEqual(len(outside), 1)
        self.assertEqual(tables[0]['cells'][0]['texts'][0]['text'], 'A')
        self.assertEqual(tables[0]['cells'][3]['texts'][0]['text'], 'B')
        self.assertEqual(tables[1]['cells'][0]['texts'][0]['text'], 'C')
        self.assertEqual(tables[0]['cells'][1]['texts'], [])

    def test_crossing_and_broken_separator_are_flagged(self):
        tables = detect_grid_tables(grid_image(broken=True))
        self.assertEqual(len(tables), 1)
        self.assertTrue(tables[0]['uncertain'])
        assign_text(tables, [line('across', (300, 110, 370, 140))])
        self.assertTrue(any(item['crossing'] for cell in tables[0]['cells'] for item in cell['texts']))

    def test_no_grid_is_not_reported_as_table(self):
        self.assertEqual(detect_grid_tables(Image.new('RGB', (800, 950), 'white')), [])

    def test_same_cell_text_is_sorted_by_page_position(self):
        tables = detect_grid_tables(grid_image())
        assign_text(tables, [line('second', (110, 140, 210, 160)), line('first', (100, 100, 200, 120))])
        entries = sorted(tables[0]['cells'][0]['texts'], key=lambda item: (item['rect'][1], item['rect'][0]))
        self.assertEqual([item['text'] for item in entries], ['first', 'second'])

    def test_invalid_late_cell_does_not_leave_partial_page(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source, output = root / 'input', root / 'output'
            source.mkdir()
            output.mkdir()
            grid_image().save(source / 'sample.pdf', resolution=200)
            lines = [line('valid', (100, 100, 180, 130)), line('bad\x01', (400, 100, 480, 130))]
            with patch('scan_ocr.recognize_image', return_value=lines):
                result = run_scan_tables(source, output, engine=object(), dpi=200)
            book = load_workbook(result['workbook'], data_only=True)
            try:
                self.assertEqual(book['单元格明细'].max_row, 1)
                self.assertEqual(book['逐页状态']['F2'].value, '处理失败')
            finally:
                book.close()

    def test_multipage_scanned_pdf_outputs_cells_and_failures(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source, output = root / 'input', root / 'output'
            source.mkdir()
            output.mkdir()
            # PIL writes image-only pages, so there is no pre-existing text layer.
            grid_image().save(source / 'sample.pdf', save_all=True,
                              append_images=[Image.new('RGB', (800, 950), 'white')], resolution=200)
            (source / 'bad.pdf').write_bytes(b'not a PDF')

            def fake_recognize(image, engine):
                # PDF is rendered at the same 200 DPI as the synthetic image.
                return [line('=2+2', (100, 100, 180, 130), .8)] if image.getpixel((80, 80))[0] < 200 else []

            with patch('scan_ocr.recognize_image', side_effect=fake_recognize):
                result = run_scan_tables(source, output, engine=object(), dpi=200)
            self.assertEqual(result['pdf_count'], 2)
            self.assertEqual(result['page_count'], 2)
            self.assertEqual(result['table_count'], 1)
            book = load_workbook(result['workbook'], data_only=False)
            try:
                rows = list(book['单元格明细'].iter_rows(min_row=2))
                self.assertEqual(len(rows), 4)
                self.assertEqual(rows[0][5].value, '=2+2')
                self.assertEqual(rows[0][5].data_type, 's')
                self.assertTrue(any('未识别到文字' in str(row[12].value) for row in rows))
                page_rows = list(book['逐页状态'].values)
                self.assertTrue(any(row[1] == 2 and '未识别出文字' in str(row[6]) for row in page_rows[1:]))
                self.assertEqual(page_rows[2][5], '未识别文字')
                file_rows = list(book['文件状态'].values)
                self.assertTrue(any(row[0] == 'bad.pdf' and '文件读取失败' in str(row[4]) for row in file_rows[1:]))
            finally:
                book.close()


if __name__ == '__main__':
    unittest.main()
