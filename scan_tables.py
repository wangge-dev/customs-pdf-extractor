"""扫描 PDF 有线表格识别；结果必须与原图核对。

逐页 OCR 与图像渲染由 scan_ocr 提供。本模块用 OpenCV 寻找彼此独立的
线框，将 OCR 文本按图像位置分配到单元格。不支持无框表格的可靠还原，
断线、印章、倾斜、合并单元格和复杂版式均可能造成错提或漏提。
"""
from __future__ import annotations

import math
import tempfile
from pathlib import Path

import cv2
import numpy as np
import pypdfium2 as pdfium
from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


def _positions(values, threshold):
    """将同一条粗线的相邻像素合并为一个中心位置。"""
    indexes = np.flatnonzero(values >= threshold)
    if not len(indexes):
        return []
    groups = np.split(indexes, np.where(np.diff(indexes) > 2)[0] + 1)
    return [int(round(float(group.mean()))) for group in groups]


def _edge_fraction(mask, horizontal, coord, start, end):
    """按线段统计覆盖率，空段视为无法确认边界。"""
    if end <= start:
        return 0.0
    h, w = mask.shape
    if horizontal:
        segment = mask[max(0, coord - 3):min(h, coord + 4), max(0, start):min(w, end)]
        return float(np.any(segment, axis=0).mean()) if segment.size else 0.0
    segment = mask[max(0, start):min(h, end), max(0, coord - 3):min(w, coord + 4)]
    return float(np.any(segment, axis=1).mean()) if segment.size else 0.0


def detect_grid_tables(image):
    """返回按阅读顺序排列的独立网格表及单元格像素坐标。

    输入为 RGB PIL 图像。不会将不相连的网格强行合并；没检出并不代表
    文档没有表格。检测到的边线不完整时返回提示，不能证明表格准确。
    """
    array = np.asarray(image.convert('RGB'))
    gray = cv2.cvtColor(array, cv2.COLOR_RGB2GRAY)
    h, w = gray.shape
    if h < 20 or w < 20:
        return []
    binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU)[1]
    # Open removes small characters; short closing reconnects minor scanning gaps.
    hkernel = cv2.getStructuringElement(cv2.MORPH_RECT, (max(20, min(100, w // 50)), 1))
    vkernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(20, min(100, h // 60))))
    horizontal = cv2.morphologyEx(binary, cv2.MORPH_OPEN, hkernel)
    vertical = cv2.morphologyEx(binary, cv2.MORPH_OPEN, vkernel)
    combined = cv2.bitwise_or(horizontal, vertical)
    combined = cv2.dilate(combined, cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3)), iterations=1)
    count, labels, stats, _ = cv2.connectedComponentsWithStats(combined, 8)
    tables = []
    for index in range(1, count):
        x, y, width, height, area = map(int, stats[index])
        if width < 36 or height < 36 or area < 80:
            continue
        crop_h = horizontal[y:y + height, x:x + width]
        crop_v = vertical[y:y + height, x:x + width]
        ys = [y + v for v in _positions(np.count_nonzero(crop_h, axis=1), max(12, width * .28))]
        xs = [x + v for v in _positions(np.count_nonzero(crop_v, axis=0), max(12, height * .28))]
        if len(xs) < 2 or len(ys) < 2:
            continue
        # A page frame, seal or random lines can resemble a table. All are review candidates.
        cells = []
        uncertain = False
        for row, (top, bottom) in enumerate(zip(ys, ys[1:]), 1):
            for col, (left, right) in enumerate(zip(xs, xs[1:]), 1):
                if right - left < 4 or bottom - top < 4:
                    uncertain = True
                    continue
                fractions = (
                    _edge_fraction(horizontal, True, top, left, right),
                    _edge_fraction(horizontal, True, bottom, left, right),
                    _edge_fraction(vertical, False, left, top, bottom),
                    _edge_fraction(vertical, False, right, top, bottom),
                )
                incomplete = min(fractions) < .90
                uncertain |= incomplete
                cells.append({'row': row, 'col': col, 'bbox': (left, top, right, bottom),
                              'lines_incomplete': incomplete, 'texts': []})
        if cells:
            tables.append({'bbox': (xs[0], ys[0], xs[-1], ys[-1]), 'cells': cells,
                           'uncertain': uncertain})
    tables.sort(key=lambda item: (item['bbox'][1], item['bbox'][0]))
    return tables


def _ocr_rect(record):
    points = record.get('box') or []
    if len(points) < 2:
        return None
    try:
        x = [float(point[0]) for point in points]
        y = [float(point[1]) for point in points]
        rect = min(x), min(y), max(x), max(y)
        return rect if all(math.isfinite(v) for v in rect) and rect[2] > rect[0] and rect[3] > rect[1] else None
    except (TypeError, ValueError, IndexError):
        return None


def assign_text(tables, records):
    """按最大重叠分配 OCR 结果，跨格或低重叠必须标注。"""
    unassigned = []
    for record in records:
        rect = _ocr_rect(record)
        text = str(record.get('text') or '')
        if not rect or not text.strip():
            unassigned.append((record, '文字为空或位置不可用'))
            continue
        x0, y0, x1, y1 = rect
        area = (x1 - x0) * (y1 - y0)
        candidates = []
        for table_no, table in enumerate(tables, 1):
            for cell in table['cells']:
                left, top, right, bottom = cell['bbox']
                overlap = max(0, min(x1, right) - max(x0, left)) * max(0, min(y1, bottom) - max(y0, top))
                if overlap / area >= .12:
                    candidates.append((overlap / area, table_no, cell))
        candidates.sort(key=lambda item: item[0], reverse=True)
        if not candidates or candidates[0][0] < .5:
            unassigned.append((record, '未匹配到单元格；可能是表格外文字或跨格文字'))
            continue
        fraction, _, cell = candidates[0]
        crossing = fraction < .8 or any(score >= .2 for score, _, _ in candidates[1:])
        cell['texts'].append({'text': text, 'score': record.get('score'), 'rect': rect, 'crossing': crossing})
    return unassigned


def _sheet(book, title, headers, widths):
    sheet = book.create_sheet(title)
    sheet.append(headers)
    for col, width in enumerate(widths, 1):
        sheet.column_dimensions[get_column_letter(col)].width = width
    for cell in sheet[1]:
        cell.font = Font(name='Microsoft YaHei', size=12, bold=True, color='1F1E1C')
        cell.fill = PatternFill('solid', fgColor='F0EDE6')
        cell.alignment = Alignment(vertical='center')
    sheet.freeze_panes = 'A2'
    return sheet


def _append(sheet, values):
    """原文写成文本；异常字符或过长时明确报错，不静默丢失数据。"""
    if sheet._current_row >= 1048576:
        raise ValueError('超过 Excel 工作表行数上限，请拆分批次')
    row = []
    for value in values:
        if isinstance(value, str) and (len(value) > 32767 or ILLEGAL_CHARACTERS_RE.search(value)):
            raise ValueError('原文超过 Excel 单元格长度或含非法控制字符，请检查源文件')
        if isinstance(value, str):
            cell = WriteOnlyCell(sheet, value=value)
            cell.data_type = 's'
            row.append(cell)
        else:
            row.append(value)
    sheet.append(row)


def run_scan_tables(input_dir, output_dir, engine=None, progress=None, dpi=200, model_dir=None):
    """批量识别扫描 PDF 有线表格，创建新结果目录并返回结果信息。"""
    from scan_ocr import create_ocr_engine, recognize_image, render_page

    source, output = Path(input_dir).expanduser().resolve(), Path(output_dir).expanduser().resolve()
    if not source.is_dir() or not output.is_dir():
        raise ValueError('PDF 输入文件夹和结果保存文件夹必须已经存在')
    if output == source or output.is_relative_to(source):
        raise ValueError('结果保存文件夹不能位于 PDF 输入文件夹中')
    pdfs = sorted((p for p in source.rglob('*') if p.is_file() and p.suffix.lower() == '.pdf'),
                  key=lambda p: str(p.relative_to(source)).casefold())
    if not pdfs:
        raise ValueError('输入文件夹及子文件夹没有 PDF')
    if not 72 <= int(dpi) <= 400:
        raise ValueError('渲染分辨率需在 72 到 400 DPI 之间')
    if engine is None:
        engine = create_ocr_engine(model_dir)
    target = Path(tempfile.mkdtemp(prefix='扫描表格识别_', dir=output))
    workbook_path = target / '扫描表格识别.xlsx'
    book = Workbook()
    book.remove(book.active)
    cells = _sheet(book, '单元格明细', ['相对路径', '页码', '表号', '行号', '列号', '文字', '最低识别分数', '左上X', '左上Y', '右下X', '右下Y', '状态', '说明'], [42, 10, 10, 10, 10, 48, 18, 12, 12, 12, 12, 18, 55])
    pages = _sheet(book, '逐页状态', ['相对路径', '页码', '表格数', '识别文字数', '未分配文字数', '状态', '说明', '实际分辨率(DPI)', '图像宽(像素)', '图像高(像素)'], [42, 10, 12, 18, 20, 18, 72, 22, 18, 18])
    files = _sheet(book, '文件状态', ['相对路径', '页数', '表格数', '状态', '说明'], [42, 10, 12, 18, 72])
    unmatched = _sheet(book, '未分配文字', ['相对路径', '页码', '文字', '识别分数', '位置', '说明'], [42, 10, 48, 18, 30, 68])
    page_count = table_count = cell_count = warning_count = 0
    try:
        for n, path in enumerate(pdfs, 1):
            rel = path.relative_to(source).as_posix()
            file_pages = file_tables = 0
            problems = []
            try:
                doc = pdfium.PdfDocument(str(path))
                try:
                    for page_index in range(len(doc)):
                        file_pages += 1
                        page_count += 1
                        page_no = page_index + 1
                        image = None
                        try:
                            image, metadata = render_page(doc, page_index, dpi=dpi)
                            records = recognize_image(image, engine)
                            tables = detect_grid_tables(image)
                            if not tables:
                                from scan_table_fallback import detect_horizontal_tables
                                tables = detect_horizontal_tables(image, records)
                            unassigned = assign_text(tables, records)
                            notes = []
                            if metadata.get('limited'):
                                notes.append('页面较大，渲染分辨率已降低；请检查小字')
                            if not records:
                                notes.append('整页未识别出文字，不能确认表格内容')
                            if not tables:
                                notes.append('未检测到线框表格；无框、断线或复杂表格可能漏检')
                            if any(table.get('inferred_columns') for table in tables):
                                notes.append('缺少竖线，列按文字位置推测，须核对')
                            if unassigned:
                                notes.append(f'{len(unassigned)} 条文字未匹配单元格，可能在表外或被漏提')
                            if any(table['uncertain'] for table in tables):
                                notes.append('部分表格边线不完整，可能存在断线或合并单元格')
                            cross = any(item['crossing'] for table in tables for cell in table['cells'] for item in cell['texts'])
                            if cross:
                                notes.append('部分文字跨越单元格，请检查原图')
                            if any(not cell['texts'] for table in tables for cell in table['cells']):
                                notes.append('存在未识别文字的单元格，可能为空格或漏识别')
                            low = any(isinstance(item['score'], (float, int)) and item['score'] < .8
                                      for table in tables for cell in table['cells'] for item in cell['texts'])
                            if low:
                                notes.append('存在识别分数低于 0.8 的文字，分数不是准确率')
                            if any(item['score'] is None for table in tables for cell in table['cells'] for item in cell['texts']):
                                notes.append('部分文字没有识别分数')
                            pending_cells = []
                            for table_no, table in enumerate(tables, 1):
                                for cell in table['cells']:
                                    entries = sorted(cell['texts'], key=lambda item: (item['rect'][1], item['rect'][0]))
                                    text = '\n'.join(item['text'] for item in entries)
                                    scores = [float(item['score']) for item in entries if isinstance(item['score'], (float, int)) and math.isfinite(item['score'])]
                                    reasons = []
                                    if cell.get('inferred_columns'):
                                        reasons.append('缺少竖线，列按文字位置推测，须核对')
                                    if cell['lines_incomplete']:
                                        reasons.append('边线不完整，可能断线或合并单元格')
                                    if any(item['crossing'] for item in entries):
                                        reasons.append('文字可能跨格')
                                    if not entries:
                                        reasons.append('未识别到文字，可能为空格或漏识别')
                                    if scores and min(scores) < .8:
                                        reasons.append('识别分数低于 0.8')
                                    if any(item['score'] is None for item in entries):
                                        reasons.append('识别分数不可用')
                                    pending_cells.append([rel, page_no, table_no, cell['row'], cell['col'], text,
                                                          min(scores) if scores else None, *cell['bbox'], '需核对', '；'.join(reasons)])
                            pending_unmatched = []
                            for item, reason in unassigned:
                                pending_unmatched.append([rel, page_no, str(item.get('text') or ''), item.get('score'),
                                                          str(_ocr_rect(item) or ''), reason])
                            # A failed cell must not leave misleading partial data for the page.
                            for sheet, rows in ((cells, pending_cells), (unmatched, pending_unmatched)):
                                if sheet._current_row + len(rows) > 1048576:
                                    raise ValueError('超过 Excel 工作表行数上限，请拆分批次')
                                for row in rows:
                                    for value in row:
                                        if isinstance(value, str) and (len(value) > 32767 or ILLEGAL_CHARACTERS_RE.search(value)):
                                            raise ValueError('原文超过 Excel 单元格长度或含非法控制字符')
                            for row in pending_cells:
                                _append(cells, row)
                            for row in pending_unmatched:
                                _append(unmatched, row)
                            cell_count += len(pending_cells)
                            file_tables += len(tables)
                            table_count += len(tables)
                            warning_count += bool(notes)
                            if notes:
                                problems.append(f'第 {page_no} 页：' + '；'.join(notes))
                            _append(pages, [rel, page_no, len(tables), len(records), len(unassigned),
                                            '未识别文字' if not records else '需核对',
                                            '；'.join(notes) or '已识别，请逐格核对原图',
                                            metadata.get('dpi'), image.width, image.height])
                        except Exception as exc:
                            warning_count += 1
                            msg = f'第 {page_no} 页处理失败：{type(exc).__name__}: {exc}'
                            problems.append(msg)
                            _append(pages, [rel, page_no, 0, 0, 0, '处理失败', msg])
                        finally:
                            if image is not None:
                                image.close()
                finally:
                    doc.close()
            except Exception as exc:
                warning_count += 1
                problems.append(f'文件读取失败：{type(exc).__name__}: {exc}')
            if not file_pages and not problems:
                warning_count += 1
                problems.append('PDF 没有页面')
            summary = '；'.join(problems) or '请对照原文件核对表格'
            if len(summary) > 32767:
                summary = '异常较多，请查看逐页状态'
            _append(files, [rel, file_pages, file_tables, '需核对', summary])
            if progress:
                progress(f'已处理 {n}/{len(pdfs)}：{rel}')
        for sheet in (cells, pages, files, unmatched):
            sheet.auto_filter.ref = sheet.dimensions
        book.save(workbook_path)
    finally:
        book.close()
    return {'output_dir': target, 'workbook': workbook_path, 'pdf_count': len(pdfs),
            'page_count': page_count, 'table_count': table_count, 'cell_count': cell_count,
            'warning_count': warning_count}
