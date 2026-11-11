"""CLI integration of the general modes with generated documents."""
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch

from openpyxl import load_workbook
from general_backend import run_mode
from test_pdf_tables import _pdf, GRID


class GeneralModesTest(unittest.TestCase):
    def test_failed_text_export_is_in_summary(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source, output = root / 'source', root / 'output'
            source.mkdir()
            output.mkdir()
            _pdf(source / 'example.pdf', [GRID])
            failed = [{'文件名': 'example.pdf', '相对路径': 'example.pdf', '状态': '无法导出',
                       '页数': 1, '输出文件': '', '说明': 'write failed'}]
            with patch('pdf_inventory.export_text_directory', return_value=failed):
                info = run_mode('PDF 体检及文本导出', source, output)
            self.assertEqual(info['未导出文件数'], 1)
            self.assertEqual(info['待查看文件数'], 1)
            wb = load_workbook(info['主要结果'], data_only=True)
            self.assertTrue(any('无法导出' in str(row) for row in wb['文本导出'].values))
            wb.close()

    def test_cli_inventory_tables_and_fields(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source, output = root / 'source', root / 'output'
            source.mkdir()
            output.mkdir()
            _pdf(source / 'example.pdf', [GRID, ''])
            (source / 'broken.pdf').write_bytes(b'not a PDF')
            template = root / 'fields.json'
            template.write_text(json.dumps({'fields': [
                {'name': 'item', 'pattern': r'(apple)', 'group': 1},
                {'name': 'missing', 'pattern': 'not-present'},
            ]}), encoding='utf-8')
            env = dict(os.environ, PYTHONIOENCODING='utf-8')
            entry = Path(__file__).resolve().parents[1] / 'main.py'
            for mode, extra in (('--inventory', []), ('--tables', []), ('--fields', [str(template)])):
                proc = subprocess.run([sys.executable, str(entry), mode, str(source), str(output), *extra],
                                      capture_output=True, text=True, encoding='utf-8', env=env, timeout=90)
                self.assertEqual(proc.returncode, 0, proc.stderr)
                self.assertIn('PDF总数：2', proc.stdout)
            books = {p.name: p for p in output.rglob('*.xlsx')}
            self.assertEqual(set(books), {'PDF文件体检.xlsx', 'PDF表格提取结果.xlsx', '通用字段提取.xlsx'})
            wb = load_workbook(books['PDF文件体检.xlsx'], data_only=True)
            self.assertIn('文本导出', wb.sheetnames)
            self.assertTrue(any('存在无文字页' in str(row) for row in wb['文件清单'].values))
            wb.close()
            text = next(output.rglob('example.pdf.txt')).read_text(encoding='utf-8')
            self.assertIn('第 2 页', text)
            self.assertIn('未进行 OCR', text)
            wb = load_workbook(books['PDF表格提取结果.xlsx'], data_only=False)
            self.assertTrue(any(row[5] == 'apple' for row in list(wb['单元格明细'].values)[1:]))
            wb.close()
            wb = load_workbook(books['通用字段提取.xlsx'], data_only=True)
            self.assertTrue(any(row[3] == 'apple' for row in list(wb['匹配结果'].values)[1:]))
            self.assertTrue(any(row[3] == '未找到（检索不完整）' for row in list(wb['待核对'].values)[1:]))
            wb.close()


if __name__ == '__main__':
    unittest.main()
