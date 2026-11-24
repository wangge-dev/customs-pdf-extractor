"""扫描 PDF 的离线逐页 OCR：保留位置、分数、文本和失败状态。"""
from __future__ import annotations

import ctypes
import math
import os
import sys
import tempfile
from pathlib import Path

import pypdfium2 as pdfium
from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
RESOURCE_ROOT = Path(getattr(sys, '_MEIPASS', ROOT))
_RUNTIME = None


def create_ocr_engine(model_dir=None):
    """只使用安装包内的检测模型和本地识别模型，不触发模型下载。"""
    global _RUNTIME
    runtime = RESOURCE_ROOT / 'runtime' / 'msvcp140.dll'
    if sys.platform == 'win32' and runtime.is_file():
        _RUNTIME = ctypes.WinDLL(str(runtime))
    import rapidocr
    from rapidocr import RapidOCR, OCRVersion, ModelType, EngineType
    detection = Path(rapidocr.__file__).resolve().parent / 'models' / 'PP-OCRv6_det_small.onnx'
    recognition = Path(model_dir or RESOURCE_ROOT / 'ocr_models') / 'PP-OCRv6_rec_medium.onnx'
    missing = [str(path) for path in (detection, recognition) if not path.is_file()]
    if missing:
        raise FileNotFoundError('缺少本地 OCR 模型，请按说明准备；运行时不会自动下载：' + '；'.join(missing))
    fonts = [Path(os.environ.get('WINDIR', 'C:/Windows')) / 'Fonts' / name
             for name in ('msyh.ttc', 'simsun.ttc', 'arial.ttf')]
    fonts.extend((Path('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'),
                  Path('/System/Library/Fonts/Helvetica.ttc')))
    font = next((item for item in fonts if item.is_file()), None)
    if font is None:
        raise FileNotFoundError('未找到可用字体，无法初始化 OCR。')
    engine = RapidOCR(params={
        'Global.use_det': True, 'Global.use_cls': False, 'Global.use_rec': True,
        'Global.text_score': 0.0, 'Global.max_side_len': 4000, 'Global.font_path': str(font),
        'Det.model_path': str(detection), 'Rec.model_path': str(recognition),
        'Rec.ocr_version': OCRVersion.PPOCRV6, 'Rec.model_type': ModelType.MEDIUM,
        'Rec.engine_type': EngineType.ONNXRUNTIME,
        'EngineConfig.onnxruntime.intra_op_num_threads': 2,
        'EngineConfig.onnxruntime.inter_op_num_threads': 1,
    })
    engine.scan_engine_label = 'RapidOCR PP-OCRv6 small 检测 / medium 识别'
    return engine


def render_page(document, page_index, dpi=200, max_pixels=20_000_000):
    """逐页渲染为独立 RGB 图片，并返回实际 DPI 和是否限缩。"""
    if not math.isfinite(dpi) or not math.isfinite(max_pixels) or dpi <= 0 or max_pixels < 1:
        raise ValueError('渲染 DPI 和最大像素数必须大于零。')
    page = document[page_index]
    bitmap = None
    try:
        width, height = page.get_size()
        if not (math.isfinite(width) and math.isfinite(height)) or width <= 0 or height <= 0:
            raise ValueError('页面尺寸无效。')
        requested = dpi / 72
        scale = min(requested, math.sqrt(max_pixels / (width * height)))
        # PDFium rounds dimensions up; keep the actual allocation within the cap.
        while math.ceil(width * scale) * math.ceil(height * scale) > max_pixels:
            scale *= 0.999
        bitmap = page.render(scale=scale)
        image = bitmap.to_pil().convert('RGB').copy()
        return image, {'dpi': round(scale * 72, 2), 'limited': scale < requested * 0.9999}
    finally:
        if bitmap is not None:
            bitmap.close()
        page.close()


def recognize_image(image, engine):
    """检测并识别整页，返回左上原点的像素四边形，不自行筛去低分。"""
    result = engine(image, use_det=True, use_cls=False, use_rec=True, text_score=0.0)
    texts = result.txts if result.txts is not None else ()
    scores = result.scores if result.scores is not None else ()
    boxes = result.boxes if result.boxes is not None else ()
    if not (len(texts) == len(scores) == len(boxes)):
        raise ValueError('OCR 返回的文字、分数和位置数量不一致。')
    lines = []
    for text, score, box in zip(texts, scores, boxes):
        points = [[float(point[0]), float(point[1])] for point in box]
        if len(points) != 4 or not all(math.isfinite(v) for point in points for v in point):
            raise ValueError('OCR 返回的位置无效。')
        score = float(score) if score is not None else None
        if score is not None and (not math.isfinite(score) or not 0 <= score <= 1):
            score = None
        lines.append({'text': str(text), 'score': score, 'box': points})
    return lines


def _append(sheet, values):
    """始终以字符串写文本；Excel 不支持的内容由调用方显式标记。"""
    row = []
    for value in values:
        if isinstance(value, str):
            if len(value) > 32767 or ILLEGAL_CHARACTERS_RE.search(value):
                raise ValueError('结果含 Excel 无法保存的文本。')
            cell = WriteOnlyCell(sheet, value=value)
            cell.data_type = 's'
            row.append(cell)
        else:
            row.append(value)
    sheet.append(row)


def _sheet(book, title, headings, widths):
    sheet = book.create_sheet(title)
    _append(sheet, headings)
    for i, cell in enumerate(sheet[1], 1):
        cell.font = Font(bold=True, color='1F1E1C')
        cell.fill = PatternFill('solid', fgColor='F0EDE6')
        sheet.column_dimensions[get_column_letter(i)].width = widths[i - 1]
    sheet.freeze_panes = 'A2'
    return sheet


def run_scan_ocr(input_dir, output_dir, progress_callback=None, *, engine=None, dpi=200, max_pixels=20_000_000):
    """识别每一页原图；失败逐页/文件记录，扫描结果始终须复核。"""
    source, output = Path(input_dir).expanduser().resolve(), Path(output_dir).expanduser().resolve()
    if not source.is_dir() or not output.is_dir():
        raise ValueError('PDF 输入文件夹和结果保存文件夹必须已经存在。')
    if output == source or output.is_relative_to(source):
        raise ValueError('结果保存文件夹不能位于 PDF 输入文件夹中。')
    if not math.isfinite(dpi) or not math.isfinite(max_pixels) or dpi <= 0 or max_pixels < 1:
        raise ValueError('渲染 DPI 和最大像素数必须大于零。')
    pdfs = sorted((p for p in source.rglob('*') if p.is_file() and p.suffix.lower() == '.pdf'),
                  key=lambda p: str(p.relative_to(source)).casefold())
    if not pdfs:
        raise ValueError('输入文件夹及其子文件夹中未找到 PDF。')
    # Initialization is deliberately before creating output directories.
    engine = engine if engine is not None else create_ocr_engine()
    engine_label = getattr(engine, 'scan_engine_label', type(engine).__name__)
    target = Path(tempfile.mkdtemp(prefix='扫描PDF文字识别_', dir=output))
    text_dir = target / '逐页识别文本'
    book = Workbook()
    book.remove(book.active)
    details = _sheet(book, '识别明细', ['相对路径', '页码', '顺序', '识别文字', '分数', '四角像素坐标', '状态', '识别引擎'],
                     [40, 10, 10, 55, 12, 48, 28, 45])
    pages = _sheet(book, '页面状态', ['相对路径', '页码', '状态', '文字行数', '实际DPI', '说明'],
                   [40, 10, 25, 14, 14, 60])
    files = _sheet(book, '文件状态', ['相对路径', '页数', 'OCR完成页数', '状态', '说明', '文本文件'],
                   [40, 10, 12, 25, 60, 50])
    page_count = ocr_count = issue_count = 0
    try:
        for index, path in enumerate(pdfs, 1):
            rel = path.relative_to(source).as_posix()
            text_path = text_dir / path.relative_to(source).parent / (path.name + '.txt')
            file_pages = file_ocr = file_attempted = 0
            errors = []
            try:
                with pdfium.PdfDocument(path) as document:
                    file_pages = len(document)
                    page_count += file_pages
                    if not file_pages:
                        errors.append('没有页面')
                        issue_count += 1
                    text_path.parent.mkdir(parents=True, exist_ok=True)
                    with text_path.open('x', encoding='utf-8', newline='\n') as text_file:
                        for page_index in range(file_pages):
                            page_no = page_index + 1
                            text_file.write(f'===== 第 {page_no} 页（整页图像 OCR）=====\n')
                            file_attempted += 1
                            image = None
                            try:
                                image, info = render_page(document, page_index, dpi=dpi, max_pixels=max_pixels)
                                try:
                                    lines = recognize_image(image, engine)
                                finally:
                                    image.close()
                                notes = []
                                if not lines or not any(line['text'].strip() for line in lines):
                                    notes.append('未识别出文字，可能是空白页、图片内容未识别或版面异常')
                                    text_file.write('[本页未识别出文字，请核对原页。]\n')
                                elif any(not line['text'].strip() for line in lines):
                                    notes.append('部分检测框未识别出文字')
                                if info['limited']:
                                    notes.append('页面较大，渲染分辨率已降低，可能漏字')
                                if any(line['score'] is None or line['score'] < 0.8 for line in lines):
                                    notes.append('含低分或无分数的文字；分数不是准确率')
                                for order, line in enumerate(lines, 1):
                                    text = line['text']
                                    text_file.write(text + '\n')
                                    detail_note = '需对照原页核对'
                                    if len(text) > 32767 or ILLEGAL_CHARACTERS_RE.search(text):
                                        text = '[内容超出 Excel 限制或含不支持字符，请查看逐页文本]'
                                        detail_note = '原文保留在逐页文本，须查看'
                                        if '部分文字不能写入 Excel，原文保留在文本文件' not in notes:
                                            notes.append('部分文字不能写入 Excel，原文保留在文本文件')
                                    coords = '; '.join(f'{x:.1f},{y:.1f}' for x, y in line['box'])
                                    _append(details, [rel, page_no, order, text, line['score'], coords, detail_note, engine_label])
                                state = '需核对' if notes else '已识别（需核对原文）'
                                _append(pages, [rel, page_no, state, len(lines), info['dpi'], '；'.join(notes)])
                                ocr_count += 1
                                file_ocr += 1
                                if notes:
                                    issue_count += 1
                                    errors.append(f'第 {page_no} 页：' + '；'.join(notes))
                            except Exception as exc:
                                issue_count += 1
                                error = f'{type(exc).__name__}: {exc}'
                                errors.append(f'第 {page_no} 页识别失败：{error}')
                                text_file.write('[本页识别失败，请查看页面状态]\n')
                                _append(pages, [rel, page_no, '识别失败', None, None, error])
                            text_file.write('\n')
            except Exception as exc:
                issue_count += 1
                if file_pages:
                    errors.append(f'文件读取或逐页文本输出失败：{type(exc).__name__}: {exc}；'
                                  f'{file_pages - file_attempted} 页未处理')
                else:
                    errors.append(f'文件打开或输出失败（可能为密码保护）：{type(exc).__name__}: {exc}')
            # Keep summaries short; each detailed failure is retained in the page sheet.
            summary = '；'.join(errors)
            if len(summary) > 32767 or ILLEGAL_CHARACTERS_RE.search(summary):
                summary = '部分异常说明无法写入单元格，请查看页面状态。'
            _append(files, [rel, file_pages, file_ocr, '需核对' if errors else '已识别（需核对原文）',
                            summary, str(text_path) if text_path.is_file() else ''])
            if progress_callback:
                progress_callback(f'已处理 {index}/{len(pdfs)}：{rel}')
        for sheet in (details, pages, files):
            sheet.auto_filter.ref = sheet.dimensions
        workbook = target / '扫描PDF文字识别.xlsx'
        book.save(workbook)
    finally:
        book.close()
    return {'result_dir': target, 'workbook': workbook, 'text_dir': text_dir,
            'pdf_count': len(pdfs), 'page_count': page_count, 'ocr_count': ocr_count,
            'issue_count': issue_count}
