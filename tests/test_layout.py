"""Printed page positions in these fixtures are synthetic PDF points."""
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock, patch

from audit_customs import field_specs
from backend import _merge_file
from extract_customs import COLS, HEADER_COLUMNS, parse_document, read_header, read_row


def word(text, x, y, width=12):
    return dict(text=text, x0=x, x1=x + width, top=y, bottom=y + 7)


def row(number, y, quantity):
    return [word(str(number), 10, y, 5), word('1234567890', 35, y, 45),
            word('坐垫', 85, y), word('9千克', 402, y), word(quantity, 410, y + 20),
            word('5.02', 510, y), word('20.08', 509, y + 10), word('美元', 509, y + 20),
            word('中国', 602, y), word('德国', 655, y), word('(12345)杭州', 722, y)]


def page(words, mark):
    obj = SimpleNamespace(width=842, height=595, edges=[], images=[], annots=[],
                          chars=[None] * 70, words=words, search=lambda _: [])
    obj.extract_text = lambda: mark
    return obj


def parse(pages):
    context = MagicMock()
    context.__enter__.return_value.pages = pages
    with patch('extract_customs.pdfplumber.open', return_value=context), patch(
            'extract_customs.normalized_words', side_effect=lambda p: p.words):
        return parse_document(Path('fixture.pdf'))


class LayoutTests(unittest.TestCase):
    def setUp(self):
        xs = [10, 45, 130, 380, 480, 595, 645, 710, 795]
        self.head = [word(text, x, 80) for text, x in zip(COLS, xs)]
        self.first = page(self.head + row(1, 100, '4条'), '页码/页数:1/2')

    def test_bill_cell_preserves_its_own_slash(self):
        header = dict.fromkeys(HEADER_COLUMNS)
        words = [word('运输工具名称及航次号', 260, 10, 80), word('提运单号', 400, 10, 45),
                 word('/', 261, 21), word('AB/123', 401, 21, 40)]
        obj = page(words, '')
        obj.edges = [dict(orientation='v', x0=x, top=5, bottom=35) for x in (250, 395, 520)]
        read_header(obj, words, header, lambda *_: None)
        self.assertEqual(header['提运单号'], 'AB/123')

    def test_duplicate_item_has_distinct_source_on_same_page(self):
        doc = parse([page(self.head + row(1, 100, '4条') + row(1, 150, '8条'), '页码/页数:1/1')])
        specs = field_specs(None, dict(headers=[doc['header']], details=doc['details'], evidence=doc['evidence']))
        quantities = [s for s in specs if s['field'] == '数量']
        self.assertEqual([(s['position'], s['value']) for s in quantities], [(100, '4条'), (150, '8条')])
        self.assertNotEqual(quantities[0]['bbox'], quantities[1]['bbox'])
        self.assertTrue(any(i['字段'] == '项号列异常' for i in doc['issues']))

    def test_wide_right_aligned_total_and_empty_last_quantity(self):
        body = [v for v in row(1, 100, '4条') if v['text'] not in ('4条', '20.08')]
        body.append(word('12345.67', 418, 110, 110))
        regions = dict(zip(COLS, zip([0, 26, 80, 370, 480, 595, 645, 710, 795],
                                     [26, 80, 370, 480, 595, 645, 710, 795, 842])))
        item, _ = read_row(body, body[0], [body[1]], regions, 'fixture.pdf', 1, lambda *_: None)
        self.assertIsNone(item['数量'])
        self.assertEqual(item['商品名称'], '坐垫')
        self.assertEqual(item['总价'], 12345.67)

    def test_headerless_page_uses_row_and_document_evidence(self):
        for printed in ('页码/页数:2/2', ''):
            with self.subTest(printed=printed):
                doc = parse([self.first, page(row(2, 42, '6套'), printed)])
                self.assertEqual([r['项号'] for r in doc['details']], [1, 2])
                self.assertEqual(doc['details'][1]['数量'], '6套')

    def test_attachment_and_broken_continuation_report_page_and_count(self):
        for words in ([word('装箱单', 200, 10)] + row(2, 42, '6条'), row(7, 42, '6条')):
            with self.subTest(first_word=words[0]['text']):
                doc = parse([self.first, page(words, '页码/页数:2/2')])
                combined = dict(audit=[], headers=[], details=[], rules=[], summaries=[])
                data = dict(doc, headers=[doc['header']], page_count=doc['pages'])
                _merge_file(data, 'fixture.pdf', '', combined)
                self.assertEqual([r['项号'] for r in doc['details']], [1])
                note = combined['summaries'][0]['完整性说明']
                self.assertIn('第 2 页未确认', note)
                self.assertIn('已导出 1 项', note)
