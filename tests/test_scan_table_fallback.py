"""Horizontal-line table candidates use only generated images and fictional text."""
import unittest
from PIL import Image, ImageDraw

from scan_table_fallback import detect_horizontal_tables, NOTICE
from scan_tables import assign_text


def record(text, x0, y0, x1, y1):
    return {'text': text, 'score': .95, 'box': [[x0, y0], [x1, y0], [x1, y1], [x0, y1]]}


def image_and_records(tops=(80,), broken=False, columns=3, rows=2):
    image = Image.new('RGB', (900, 900), 'white')
    draw = ImageDraw.Draw(image)
    records = []
    for t, top in enumerate(tops):
        for y in range(top, top + 90 * (rows+1), 90):
            if broken:
                draw.line((80, y, 405, y), fill='black', width=3)
                draw.line((420, y, 800, y), fill='black', width=3)
            else:
                draw.line((80, y, 800, y), fill='black', width=3)
        for row in range(rows):
            for col in range(columns):
                x = 115 + col * 210
                y = top + 22 + row * 90
                records.append(record(f'{t}-{row}-{col}', x, y, x+60, y+35))
    return image, records


class HorizontalFallbackTests(unittest.TestCase):
    def test_two_tables_remain_separate_and_records_are_assigned(self):
        image, records = image_and_records(tops=(60, 480))
        tables = detect_horizontal_tables(image, records)
        self.assertEqual(len(tables), 2)
        self.assertEqual([len(t['cells']) for t in tables], [6, 6])
        self.assertTrue(all(t['notice'] == NOTICE and t['inferred_columns'] for t in tables))
        self.assertTrue(all(c['notice'] == NOTICE for t in tables for c in t['cells']))
        self.assertEqual(assign_text(tables, records), [])
        self.assertEqual(sum(len(c['texts']) for t in tables for c in t['cells']), len(records))

    def test_adjacent_tables_with_a_blank_band_remain_separate(self):
        image, records = image_and_records(tops=(60, 330))
        tables = detect_horizontal_tables(image, records)
        self.assertEqual(len(tables), 2)
        self.assertEqual([len(table['cells']) for table in tables], [6, 6])

    def test_side_by_side_tables_are_not_merged(self):
        image = Image.new('RGB', (1000, 500), 'white')
        draw = ImageDraw.Draw(image)
        records = []
        for start in (50, 560):
            for y in (80, 160, 240):
                draw.line((start, y, start+360, y), fill='black', width=3)
            for y in (105, 185):
                for x in (start+25, start+195):
                    records.append(record(f'{start}-{x}-{y}', x, y, x+55, y+30))
        tables = detect_horizontal_tables(image, records)
        self.assertEqual(len(tables), 2)
        self.assertEqual([len(table['cells']) for table in tables], [4, 4])

    def test_one_row_and_single_column_are_not_tables(self):
        image, lines = image_and_records(rows=1)
        self.assertEqual(detect_horizontal_tables(image, lines), [])
        image, lines = image_and_records(columns=1)
        self.assertEqual(detect_horizontal_tables(image, lines), [])

    def test_broken_rule_and_missing_or_crossing_cell(self):
        image, records = image_and_records(rows=3, broken=True)
        records.pop(4)  # An empty cell must stay visible.
        records.append(record('crossing', 220, 290, 280, 320))
        tables = detect_horizontal_tables(image, records)
        self.assertEqual(len(tables), 1)
        self.assertEqual(len(tables[0]['cells']), 9)
        unassigned = assign_text(tables, records)
        assigned = sum(len(c['texts']) for c in tables[0]['cells'])
        self.assertEqual(assigned + len(unassigned), len(records))
        self.assertTrue(any(not c['texts'] for c in tables[0]['cells']))
        self.assertTrue(any(item['crossing'] for c in tables[0]['cells'] for item in c['texts']) or unassigned)


if __name__ == '__main__':
    unittest.main()
