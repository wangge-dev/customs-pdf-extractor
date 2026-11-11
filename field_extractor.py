"""Extract user-specified fields from PDF text layers using local regex templates.

This is not OCR: image-only, empty and unreadable pages require manual review.
Use only templates supplied by trusted users; pathological regular expressions can
take a very long time to run in Python's standard regular-expression engine.
"""
from __future__ import annotations

import json
import re
import tempfile
from pathlib import Path

import pdfplumber
from openpyxl import Workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter


def load_template(template_path):
    """Validate JSON and return ordered (name, compiled pattern, group) tuples."""
    path = Path(template_path)
    try:
        data = json.loads(path.read_text(encoding="utf-8-sig"))
    except (OSError, ValueError) as exc:
        raise ValueError(f"无法读取 JSON 模板 {path}: {exc}") from exc
    fields = data.get("fields") if isinstance(data, dict) else None
    if not isinstance(fields, list) or not fields:
        raise ValueError("模板必须包含非空 fields 数组")
    rules, names = [], set()
    for i, field in enumerate(fields, 1):
        if not isinstance(field, dict):
            raise ValueError(f"第 {i} 项字段必须是对象")
        name, pattern, group = field.get("name"), field.get("pattern"), field.get("group", 0)
        if not isinstance(name, str) or not name.strip() or name in names:
            raise ValueError(f"第 {i} 项字段名称为空或重复")
        if not isinstance(pattern, str) or not pattern:
            raise ValueError(f"字段 {name} 的 pattern 必须是非空字符串")
        if not isinstance(group, (int, str)) or isinstance(group, bool) or (isinstance(group, int) and group < 0):
            raise ValueError(f"字段 {name} 的 group 必须是非负整数或捕获组名称")
        try:
            compiled = re.compile(pattern, re.MULTILINE)
        except re.error as exc:
            raise ValueError(f"字段 {name} 的正则表达式无效：{exc}") from exc
        if (isinstance(group, int) and group > compiled.groups) or (isinstance(group, str) and group not in compiled.groupindex):
            raise ValueError(f"字段 {name} 的 group 不存在：{group}")
        names.add(name)
        rules.append((name, compiled, group))
    return rules


def _append(sheet, values):
    """Keep strings as text and refuse values Excel would silently truncate."""
    if sheet.max_row >= 1048576:
        raise ValueError(f"工作表 {sheet.title} 超出 Excel 行数上限；请缩小本次处理范围")
    for value in values:
        if isinstance(value, str):
            if len(value) > 32767:
                raise ValueError(f"工作表 {sheet.title} 中有文本超过 Excel 单元格 32767 字符上限；请调整模板")
            if ILLEGAL_CHARACTERS_RE.search(value):
                raise ValueError(f"工作表 {sheet.title} 中有 Excel 不支持的控制字符；请检查源文件或模板")
    sheet.append(values)
    for cell in sheet[sheet.max_row]:
        if isinstance(cell.value, str):
            cell.data_type = "s"


def extract_fields(input_dir, output_parent, template_path, progress=None):
    """Process all PDFs recursively, returning result paths and counts.

    progress, if supplied, is called with a single human-readable status string.
    Each call creates a distinct output directory without overwriting past output.
    """
    source, parent = Path(input_dir).resolve(), Path(output_parent).resolve()
    if not source.is_dir():
        raise NotADirectoryError(f"PDF 文件夹不存在：{source}")
    if not parent.is_dir():
        raise NotADirectoryError(f"结果保存文件夹不存在：{parent}")
    rules = load_template(template_path)
    pdfs = sorted((p for p in source.rglob("*") if p.is_file() and p.suffix.lower() == ".pdf"),
                  key=lambda p: str(p.relative_to(source)).casefold())
    if not pdfs:
        raise ValueError(f"文件夹内未找到 PDF：{source}")
    result_dir = Path(tempfile.mkdtemp(prefix="通用字段提取_", dir=parent))
    output = result_dir / "通用字段提取.xlsx"
    book = Workbook()
    matches = book.active
    matches.title = "匹配结果"
    issues = book.create_sheet("待核对")
    _append(matches, ["相对路径", "页码", "字段", "值", "上下文", "文件内字段匹配序号"])
    _append(issues, ["相对路径", "页码", "字段", "问题", "说明"])
    match_count = warning_count = 0

    def warn(rel, page, name, problem, detail):
        nonlocal warning_count
        _append(issues, [rel, page, name, problem, detail])
        warning_count += 1

    try:
        for index, path in enumerate(pdfs, 1):
            rel = str(path.relative_to(source))
            counts = {name: 0 for name, _, _ in rules}
            incomplete = False
            try:
                pdf = pdfplumber.open(path)
            except Exception as exc:
                warn(rel, "", "", "无法打开", f"{type(exc).__name__}: {exc}")
                incomplete = True
                pdf = None
            if pdf is not None:
                with pdf:
                    try:
                        pages = pdf.pages
                    except Exception as exc:
                        warn(rel, "", "", "无法读取页面", f"{type(exc).__name__}: {exc}")
                        incomplete = True
                        pages = []
                    if not pages and not incomplete:
                        warn(rel, "", "", "没有页面", "无法检索字段")
                        incomplete = True
                    for page_num, page in enumerate(pages, 1):
                        try:
                            content = page.extract_text() or ""
                        except Exception as exc:
                            warn(rel, page_num, "", "页面文字读取失败", f"{type(exc).__name__}: {exc}")
                            incomplete = True
                            continue
                        if not content.strip():
                            warn(rel, page_num, "", "无可提取文字", "可能为空白页、扫描页或文字层不受支持；此功能未执行 OCR")
                            incomplete = True
                            continue
                        for name, pattern, group in rules:
                            for found in pattern.finditer(content):
                                value = found.group(group)
                                counts[name] += 1
                                match_count += 1
                                context = content[max(0, found.start() - 35):min(len(content), found.end() + 35)]
                                _append(matches, [rel, page_num, name, value if value is not None else "", context, counts[name]])
                                if not value:
                                    warn(rel, page_num, name, "匹配值为空", "请检查捕获组和源文档")
            for name, _, _ in rules:
                n = counts[name]
                if n == 0:
                    warn(rel, "", name, "未找到" if not incomplete else "未找到（检索不完整）",
                         "未检索到匹配值；请检查模板和原文" if not incomplete else "有页面不可检索，不能确认该字段不存在")
                elif n > 1:
                    warn(rel, "", name, "多次匹配", f"共 {n} 个匹配，请逐项核对")
            if progress:
                progress(f"已处理 {index}/{len(pdfs)}：{rel}")
        for sheet in (matches, issues):
            sheet.freeze_panes = "B2"
            sheet.auto_filter.ref = sheet.dimensions
            for cell in sheet[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="4472C4")
            for col in range(1, sheet.max_column + 1):
                sheet.column_dimensions[get_column_letter(col)].width = 45 if col in (1, 5) else 20
        book.save(output)
    finally:
        book.close()
    return {"output_dir": result_dir, "workbook": output, "pdf_count": len(pdfs),
            "match_count": match_count, "warning_count": warning_count}
