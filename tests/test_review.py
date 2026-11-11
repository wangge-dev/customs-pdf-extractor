"""Synthetic regression checks: no real declarations or OCR models required."""
import tempfile
import unittest
from pathlib import Path

from backend import _merge_file
from exporter import _html_report
from extract_customs import HEADER_COLUMNS


class AuditSummaryTests(unittest.TestCase):
    def test_invalid_numeric_ocr_is_incomplete_even_with_score(self):
        header = {field: '示例' for field in HEADER_COLUMNS}
        header.update({'文件名': 'example.pdf', '境内发货人18位代码': '123456789012345678',
                       '生产销售单位18位代码': '123456789012345678',
                       '件数': 1, '毛重(千克)': 2, '净重(千克)': 1})
        rows = []
        for field in HEADER_COLUMNS[1:]:
            value = header[field]
            rows.append({'文件名': 'example.pdf', '页码': 1, '项号': None,
                         '字段': field, '提取值': value, 'OCR读数': str(value),
                         'OCR分数': .99, '核对状态': '两路一致', '原因': ''})
        count = next(row for row in rows if row['字段'] == '件数')
        count.update({'OCR读数': 'I', 'OCR分数': .4, '核对状态': '待核对'})
        combined = {'audit': [], 'headers': [], 'details': [], 'rules': [], 'summaries': []}
        _, incomplete = _merge_file({'headers': [header], 'details': [], 'evidence': [],
                                     'audit': rows, 'rules': []}, 'example.pdf', '', combined)
        self.assertTrue(incomplete)
        self.assertEqual(combined['summaries'][0]['应核对字段'], 13)
        self.assertEqual(combined['summaries'][0]['已核对字段'], 12)

    def test_offline_review_escapes_input_and_adds_filters(self):
        row = {'核对编号': 1, '文件名': '<script>示例</script>.pdf', '页码': 1,
               '项号': 1, '字段': '总价', '提取值': '2', 'OCR读数': 'Z',
               'OCR分数': .5, '核对状态': '待核对', '原因': '示例', '截图': ''}
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'review.html'
            _html_report(path, {'audit': [row], 'headers': [], 'details': []})
            page = path.read_text(encoding='utf-8')
        self.assertIn('&lt;script&gt;示例&lt;/script&gt;.pdf', page)
        self.assertEqual(page.count('<script>'), 1)
        self.assertIn('id="search"', page)
        self.assertIn('id="status"', page)
        self.assertIn('id="field"', page)
        self.assertIn('id="field-1"', page)


if __name__ == '__main__':
    unittest.main()
