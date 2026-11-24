"""仅横线表格的保守候选：列由 OCR 位置推测，必须逐格核对。"""
from __future__ import annotations

import math
from statistics import median

import cv2
import numpy as np


NOTICE = '缺少竖线，列按文字位置推测，须核对'


def _rect(record):
    try:
        points = record['box']
        if len(points) != 4 or not str(record.get('text') or '').strip():
            return None
        xs = [float(p[0]) for p in points]
        ys = [float(p[1]) for p in points]
        x0, x1, y0, y1 = min(xs), max(xs), min(ys), max(ys)
        if not all(math.isfinite(v) for v in (x0, x1, y0, y1)) or x1 <= x0 or y1 <= y0:
            return None
        return x0, y0, x1, y1
    except (KeyError, TypeError, ValueError, IndexError):
        return None


def _horizontal_lines(image):
    gray = cv2.cvtColor(np.asarray(image.convert('RGB')), cv2.COLOR_RGB2GRAY)
    height, width = gray.shape
    binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU)[1]
    # Small closing reconnects scan gaps, but does not bridge unrelated distant rules.
    gap = max(3, min(61, width // 40))
    closed = cv2.morphologyEx(binary, cv2.MORPH_CLOSE, cv2.getStructuringElement(cv2.MORPH_RECT, (gap, 1)))
    kernel = max(15, min(80, width // 40))
    mask = cv2.morphologyEx(closed, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (kernel, 1)))
    _, _, stats, _ = cv2.connectedComponentsWithStats(mask, 8)
    result = []
    for x, y, w, h, area in stats[1:]:
        x, y, w, h, area = map(int, (x, y, w, h, area))
        if w >= max(60, width * .18) and h <= max(8, height * .012) and area >= w:
            result.append((x, y + h // 2, x + w))
    return sorted(result, key=lambda line: (line[1], line[0]))


def _groups(lines):
    """Group rules with comparable horizontal extent; split large vertical gaps."""
    groups = []
    for line in lines:
        x0, y, x1 = line
        choices = []
        for group in groups:
            last = group[-1]
            overlap = max(0, min(x1, last[2]) - max(x0, last[0]))
            ratio = overlap / max(1, min(x1 - x0, last[2] - last[0]))
            if y > last[1] + 4 and ratio >= .85 and max(x1-x0, last[2]-last[0]) <= 1.3 * min(x1-x0, last[2]-last[0]):
                choices.append((y - last[1], group))
        if choices:
            min(choices, key=lambda pair: pair[0])[1].append(line)
        else:
            groups.append([line])
    parts = []
    for group in groups:
        if len(group) < 3:
            continue
        gaps = [b[1] - a[1] for a, b in zip(group, group[1:])]
        typical = median(gaps)
        current = [group[0]]
        for gap, line in zip(gaps, group[1:]):
            if gap > max(typical * 2.0, typical + 20):
                if len(current) >= 3:
                    parts.append(current)
                current = []
            current.append(line)
        if len(current) >= 3:
            parts.append(current)
    return parts


def _split_empty_bands(group, rectangles):
    """不跨越没有文字的间隔连接两个横线表；空白行也可能被拆开。"""
    parts, current = [], [group[0]]
    for before, after in zip(group, group[1:]):
        left, right = max(before[0], after[0]), min(before[2], after[2])
        has_text = any(left <= (r[0]+r[2])/2 <= right and before[1] < (r[1]+r[3])/2 < after[1]
                       for r in rectangles)
        if not has_text:
            if len(current) >= 3:
                parts.append(current)
            current = [after]
        else:
            current.append(after)
    if len(current) >= 3:
        parts.append(current)
    return parts


def detect_horizontal_tables(image, records):
    """OCR + ≥3 horizontal rules + two aligned multi-column rows → candidates.

    Input is RGB PIL and records from scan_ocr.recognize_image; the format of
    returned table/cell dictionaries matches scan_tables.assign_text().
    Column edges are inferred, not observed. All output is review-required.
    Returns [] when evidence is insufficient, not evidence that no table exists.
    """
    if image.width < 20 or image.height < 20:
        return []
    rectangles = [rect for item in records if (rect := _rect(item)) is not None]
    tables = []
    groups = [part for group in _groups(_horizontal_lines(image))
              for part in _split_empty_bands(group, rectangles)]
    for group in groups:
        ys = [line[1] for line in group]
        left = max(line[0] for line in group)
        right = min(line[2] for line in group)
        if right <= left:
            continue
        rows = [[] for _ in range(len(ys)-1)]
        for rect in rectangles:
            x0, y0, x1, y1 = rect
            cx, cy = (x0+x1)/2, (y0+y1)/2
            if not left <= cx <= right:
                continue
            for row, (top, bottom) in enumerate(zip(ys, ys[1:])):
                if top < cy < bottom:
                    rows[row].append(rect)
                    break
        heights = [box[3] - box[1] for row in rows for box in row]
        if not heights:
            continue
        tolerance = max(8.0, min((right-left)*.06, median(heights)*1.3))
        clusters = []
        for row_index, row in enumerate(rows):
            for box in sorted(row, key=lambda r: r[0]):
                center = (box[0]+box[2])/2
                candidates = [(abs(center - item['center']), item) for item in clusters
                              if row_index not in item['rows'] and abs(center - item['center']) <= tolerance]
                if candidates:
                    cluster = min(candidates, key=lambda item: item[0])[1]
                    cluster['boxes'].append(box)
                    cluster['rows'].add(row_index)
                    cluster['center'] = sum((b[0]+b[2])/2 for b in cluster['boxes']) / len(cluster['boxes'])
                else:
                    clusters.append({'center': center, 'rows': {row_index}, 'boxes': [box]})
        clusters = sorted((item for item in clusters if len(item['rows']) >= 2), key=lambda item: item['center'])
        if len(clusters) < 2:
            continue
        aligned_rows = sum(sum(i in cluster['rows'] for cluster in clusters) >= 2 for i in range(len(rows)))
        if aligned_rows < 2:
            continue
        edges = [left]
        for a, b in zip(clusters, clusters[1:]):
            a_right = max(box[2] for box in a['boxes'])
            b_left = min(box[0] for box in b['boxes'])
            boundary = (a_right + b_left)/2 if a_right < b_left else (a['center'] + b['center'])/2
            edges.append(int(round(boundary)))
        edges.append(right)
        if any(b-a < 4 for a, b in zip(edges, edges[1:])):
            continue
        cells = [{'row': row+1, 'col': col+1, 'bbox': (x0, y0, x1, y1),
                  'lines_incomplete': True, 'texts': [], 'inferred_columns': True, 'notice': NOTICE}
                 for row, (y0, y1) in enumerate(zip(ys, ys[1:]))
                 for col, (x0, x1) in enumerate(zip(edges, edges[1:]))]
        tables.append({'bbox': (left, ys[0], right, ys[-1]), 'cells': cells,
                       'uncertain': True, 'inferred_columns': True, 'notice': NOTICE})
    return sorted(tables, key=lambda table: (table['bbox'][1], table['bbox'][0]))
