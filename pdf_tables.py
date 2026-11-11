"""通用 PDF 表格提取：逐页输出位置、状态和原文，不自动运行 OCR。

run_tables(input_dir, output_dir, progress_callback=None) 返回独立结果目录、
工作簿路径及统计数字。PDF 的版式不同，结果必须与原文件核对。
"""

from datetime import datetime
from pathlib import Path
import tempfile

import pdfplumber
from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter


def _within(path, root):
    try:
        path.relative_to(root)
        return True
    except ValueError:
        return False


def _safe_text(value):
    """Excel 无法存储非法控制字符或超过 32767 字符的单元格。"""
    if value is None:
        return '', False
    value = str(value)
    clean = ILLEGAL_CHARACTERS_RE.sub('', value)
    changed = clean != value or len(clean) > 32767
    return clean[:32767], changed


def _append(ws, values):
    """将任何 PDF 文本作为文本写入，禁止被 Excel 当作公式执行。"""
    row = []
    for value in values:
        if isinstance(value, str):
            cell = WriteOnlyCell(ws, value=value)
            cell.data_type = 's'
            row.append(cell)
        else:
            row.append(value)
    ws.append(row)


def _sheet(wb, name, headings, widths):
    ws = wb.create_sheet(name)
    _append(ws, headings)
    for i, cell in enumerate(ws[1], 1):
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor='4472C4')
        ws.column_dimensions[get_column_letter(i)].width = widths[i - 1]
    ws.freeze_panes = 'A2'
    return ws


def run_tables(input_dir, output_dir, progress_callback=None):
    """递归读取 PDF 并生成一份单元格明细与逐页、逐文件状态工作簿。

    参数均可为 str 或 Path；progress_callback 接受一条状态文字。输出目录
    不可位于输入目录中。返回字典的路径值为 Path；计数值为 int。
    """
    source = Path(input_dir).expanduser().resolve()
    output = Path(output_dir).expanduser().resolve()
    if not source.is_dir():
        raise ValueError('PDF 输入目录不存在或不是文件夹')
    if _within(output, source):
        raise ValueError('结果保存目录不能是输入目录或其子目录，请选择其他文件夹')
    pdfs = sorted((p for p in source.rglob('*') if p.is_file() and p.suffix.lower() == '.pdf'),
                  key=lambda p: str(p.relative_to(source)).casefold())
    if not pdfs:
        raise ValueError('输入文件夹及其子文件夹中未找到 PDF 文件')
    output.mkdir(parents=True, exist_ok=True)
    result_dir = Path(tempfile.mkdtemp(prefix='PDF表格提取_' + datetime.now().strftime('%Y%m%d_%H%M%S_'), dir=output))
    workbook_path = result_dir / 'PDF表格提取结果.xlsx'
    wb = Workbook()
    wb.remove(wb.active)
    cells = _sheet(wb, '单元格明细', ['相对路径', '页码', '表号', '行号', '列号', '单元格文本', '备注'], [45, 10, 10, 10, 10, 60, 42])
    pages = _sheet(wb, '逐页状态', ['相对路径', '页码', '表格数', '文字层', '状态', '说明'], [45, 10, 12, 12, 20, 65])
    files = _sheet(wb, '文件状态', ['相对路径', '页数', '表格数', '状态', '说明'], [45, 10, 12, 20, 65])
    page_count = table_count = cell_count = issue_count = 0

    for index, path in enumerate(pdfs, 1):
        relative = path.relative_to(source).as_posix()
        file_pages = file_tables = 0
        problems = []
        try:
            with pdfplumber.open(path) as pdf:
                for page_no, page in enumerate(pdf.pages, 1):
                    file_pages += 1
                    page_count += 1
                    try:
                        has_text = bool(''.join(str(char.get('text') or '') for char in page.chars).strip())
                        found = page.find_tables()
                        current_tables = len(found)
                        truncated = False
                        page_cells = []
                        for table_no, table in enumerate(found, 1):
                            for row_no, row in enumerate(table.extract(), 1):
                                for column_no, value in enumerate(row, 1):
                                    value, changed = _safe_text(value)
                                    truncated |= changed
                                    page_cells.append([relative, page_no, table_no, row_no, column_no, value,
                                                       '文本含 Excel 不支持的字符或超过长度限制，已截断或移除，请查原文' if changed else ''])
                        for cell in page_cells:
                            _append(cells, cell)
                        file_tables += current_tables
                        table_count += current_tables
                        cell_count += len(page_cells)
                        notes = []
                        if not has_text:
                            notes.append('无可提取文字层，可能是扫描件；本功能未执行 OCR')
                        if not current_tables:
                            notes.append('未检测到表格；无边框或复杂表格可能漏检')
                        if truncated:
                            notes.append('有单元格超过 Excel 限制或含不支持的控制字符')
                        status = '需核对' if notes else '已提取（需核对原文）'
                        if notes:
                            issue_count += 1
                            problems.append(f'第 {page_no} 页：' + '；'.join(notes))
                        _append(pages, [relative, page_no, current_tables, '有' if has_text else '无', status, '；'.join(notes)])
                    except Exception as exc:
                        issue_count += 1
                        message = f'{type(exc).__name__}: {exc}'
                        problems.append(f'第 {page_no} 页处理失败：{message}')
                        _append(pages, [relative, page_no, 0, '未知', '处理失败', message])
                    finally:
                        page.close()
        except Exception as exc:
            issue_count += 1
            problems.append(f'文件处理失败：{type(exc).__name__}: {exc}')
        if file_pages == 0 and not problems:
            issue_count += 1
            problems.append('PDF 没有可处理的页面')
        _append(files, [relative, file_pages, file_tables, '需核对' if problems else '已提取（需核对原文）', '\n'.join(problems)[:32767]])
        if progress_callback:
            progress_callback(f'已处理 {index}/{len(pdfs)}：{relative}')

    for ws in (cells, pages, files):
        ws.auto_filter.ref = ws.dimensions
    wb.save(workbook_path)
    return {'result_dir': result_dir, 'workbook': workbook_path, 'pdf_count': len(pdfs),
            'page_count': page_count, 'table_count': table_count, 'cell_count': cell_count,
            'issue_count': issue_count}
