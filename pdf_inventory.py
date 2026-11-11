"""Offline PDF inventory. Text-layer detection is not OCR or correctness verification."""

from __future__ import annotations

import argparse
from pathlib import Path

import pdfplumber
from pdfminer.pdfdocument import PDFPasswordIncorrect
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter


FILE_COLUMNS = ("文件名", "相对路径", "状态", "页数", "有文字页数", "无文字页数", "提取失败页数", "可提取文字数", "说明", "疑似图片主体页数")
PAGE_COLUMNS = ("相对路径", "页码", "宽度(mm)", "高度(mm)", "可提取文字数", "文字层", "说明", "页面提示")
EXPORT_COLUMNS = ("文件名", "相对路径", "状态", "页数", "无文字页数", "提取失败页数", "输出文件", "说明")


def _large_image_fraction(page) -> float:
    """Largest single image's clipped bounding-box area / page area, as a hint only."""
    area = page.width * page.height
    if area <= 0:
        return 0.0
    fractions = []
    for image in page.images:
        x0 = max(0.0, float(image["x0"]))
        x1 = min(page.width, float(image["x1"]))
        top = max(0.0, float(image["top"]))
        bottom = min(page.height, float(image["bottom"]))
        fractions.append(max(0.0, x1 - x0) * max(0.0, bottom - top) / area)
    return max(fractions, default=0.0)


def inspect_pdf(path: Path | str, relative_path: str | None = None) -> tuple[dict, list[dict]]:
    """Inspect one PDF; return a file record and page records without modifying the PDF.

    The text count excludes whitespace. A missing text layer may mean an image-only
    page, a blank page, or unsupported text; it does not establish scan status.
    """
    path = Path(path)
    rel = relative_path if relative_path is not None else path.name
    row = dict(zip(FILE_COLUMNS, (path.name, rel, "正常", None, 0, 0, 0, 0, "", 0)))
    pages = []
    image_check_errors = 0
    try:
        with pdfplumber.open(path) as pdf:
            row["页数"] = len(pdf.pages)
            for number, page in enumerate(pdf.pages, 1):
                entry = dict(zip(PAGE_COLUMNS, (rel, number, round(page.width * 25.4 / 72, 1),
                                                round(page.height * 25.4 / 72, 1), None, "未知", "", "")))
                try:
                    count = sum(not char.isspace() for char in (page.extract_text() or ""))
                    entry["可提取文字数"] = count
                    entry["文字层"] = "有" if count else "无"
                    row["有文字页数" if count else "无文字页数"] += 1
                    row["可提取文字数"] += count
                except Exception as exc:
                    row["提取失败页数"] += 1
                    entry["说明"] = f"文字提取失败：{type(exc).__name__}"
                if entry["可提取文字数"] is not None:
                    try:
                        # A printed page number may be searchable even when the document is an image.
                        # This intentionally only checks a single large embedded image, not OCR quality.
                        if entry["可提取文字数"] <= 30 and _large_image_fraction(page) >= 0.60:
                            entry["页面提示"] = "疑似图片主体，需核对"
                            row["疑似图片主体页数"] += 1
                    except Exception as exc:
                        image_check_errors += 1
                        entry["说明"] = f"图片检测失败：{type(exc).__name__}"
                pages.append(entry)
            if row["提取失败页数"]:
                row["状态"] = "部分页面读取失败"
            elif image_check_errors:
                row["状态"] = "部分页面图片检测失败"
            elif row["疑似图片主体页数"]:
                row["状态"] = "疑似图片主体，需核对"
            elif row["无文字页数"]:
                row["状态"] = "存在无文字页"
            elif not row["页数"]:
                row["状态"] = "无页面"
            notes = []
            if row["疑似图片主体页数"]:
                notes.append(f"{row['疑似图片主体页数']} 页疑似图片主体；仅为提示，请核对原件。")
            if image_check_errors:
                notes.append(f"{image_check_errors} 页图片检测失败。")
            row["说明"] = " ".join(notes)
    except PDFPasswordIncorrect:
        row.update({"状态": "密码保护", "说明": "无法使用空密码打开；需先提供可读取的 PDF"})
    except Exception as exc:
        row.update({"状态": "无法打开", "说明": f"{type(exc).__name__}: {exc}"})
    return row, pages


def _pdf_paths(source: Path | str, recursive: bool) -> tuple[Path, list[Path]]:
    source = Path(source)
    if not source.is_dir():
        raise NotADirectoryError(f"PDF 文件夹不存在：{source}")
    paths = sorted((p for p in (source.rglob("*") if recursive else source.iterdir())
                    if p.is_file() and p.suffix.lower() == ".pdf"), key=lambda p: str(p).casefold())
    return source, paths


def inspect_directory(source: Path | str, recursive: bool = False, progress=None) -> tuple[list[dict], list[dict]]:
    """Inspect PDFs in a directory; optional progress receives (completed, total, path)."""
    source, paths = _pdf_paths(source, recursive)
    files, pages = [], []
    for index, path in enumerate(paths, 1):
        file_row, page_rows = inspect_pdf(path, str(path.relative_to(source)))
        files.append(file_row)
        pages.extend(page_rows)
        if progress:
            progress(index, len(paths), path)
    return files, pages


def export_text_directory(source: Path | str, output_dir: Path | str, recursive: bool = False, progress=None) -> list[dict]:
    """Export text by page; no OCR, no overwriting, subdirectories are retained.

    Failures are returned as individual records, so other PDFs continue processing.
    """
    source, paths = _pdf_paths(source, recursive)
    output_dir = Path(output_dir)
    results = []
    for index, path in enumerate(paths, 1):
        relative = path.relative_to(source)
        destination = output_dir / relative.parent / (relative.name + ".txt")
        row = {"文件名": path.name, "相对路径": str(relative), "状态": "完成", "页数": 0,
               "无文字页数": 0, "提取失败页数": 0, "输出文件": "", "说明": ""}
        try:
            sections = []
            with pdfplumber.open(path) as pdf:
                row["页数"] = len(pdf.pages)
                for number, page in enumerate(pdf.pages, 1):
                    try:
                        text = page.extract_text() or ""
                        if not text.strip():
                            row["无文字页数"] += 1
                            text = "[本页未提取到文字，未进行 OCR；可能为空白页、图片页或文字无法提取。]"
                    except Exception as exc:
                        row["提取失败页数"] += 1
                        text = f"[本页文字提取失败：{type(exc).__name__}；未进行 OCR。]"
                    sections.append(f"===== 第 {number} 页 =====\n{text}\n")
            destination.parent.mkdir(parents=True, exist_ok=True)
            with destination.open("x", encoding="utf-8", newline="\n") as target:
                target.write("\n".join(sections))
            row["输出文件"] = str(destination)
            if row["提取失败页数"]:
                row["状态"] = "部分页面读取失败"
            elif row["无文字页数"]:
                row["状态"] = "存在无文字页"
        except FileExistsError:
            row.update({"状态": "未覆盖已有文件", "说明": str(destination)})
        except PDFPasswordIncorrect:
            row.update({"状态": "密码保护", "说明": "无法使用空密码打开；需先提供可读取的 PDF"})
        except Exception as exc:
            row.update({"状态": "无法导出", "说明": f"{type(exc).__name__}: {exc}"})
        results.append(row)
        if progress:
            progress(index, len(paths), path)
    return results


def write_inventory(output: Path | str, files: list[dict], pages: list[dict], exports: list[dict] | None = None) -> Path:
    """Write an Excel inventory and optional export statuses, without overwriting."""
    output = Path(output)
    if output.exists():
        raise FileExistsError(f"输出文件已存在：{output}")
    output.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    sheets = [(workbook.active, "文件清单", FILE_COLUMNS, files),
              (workbook.create_sheet(), "页面明细", PAGE_COLUMNS, pages)]
    if exports is not None:
        sheets.append((workbook.create_sheet(), "文本导出", EXPORT_COLUMNS, exports))
    for sheet, title, columns, rows in sheets:
        sheet.title = title
        sheet.append(columns)
        for row in rows:
            sheet.append([row.get(column) for column in columns])
            # PDF names can begin with '='; preserve as text, never as an Excel formula.
            for cell in sheet[sheet.max_row]:
                if isinstance(cell.value, str):
                    cell.data_type = "s"
        sheet.freeze_panes = "B2"
        sheet.auto_filter.ref = sheet.dimensions
        for cell in sheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="4472C4")
        for index, column in enumerate(columns, 1):
            sheet.column_dimensions[get_column_letter(index)].width = 45 if column in ("相对路径", "输出文件", "说明") else 20
    workbook.save(output)
    workbook.close()
    return output


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="离线生成 PDF 文件和页面体检清单，不对文件进行 OCR。")
    parser.add_argument("source", type=Path, help="待检查 PDF 的文件夹")
    parser.add_argument("output", type=Path, help="输出 Excel 路径（不能覆盖已有文件）")
    parser.add_argument("--recursive", action="store_true", help="包含子文件夹")
    parser.add_argument("--text-dir", type=Path, help="可选：逐页导出文字到此文件夹；不进行 OCR")
    args = parser.parse_args(argv)
    if args.output.suffix.lower() != ".xlsx":
        parser.error("输出文件必须以 .xlsx 结尾")
    try:
        if args.output.exists():
            raise FileExistsError(f"输出文件已存在：{args.output}")
        files, pages = inspect_directory(args.source, recursive=args.recursive,
                                         progress=lambda done, total, path: print(f"{done}/{total}: {path.name}"))
        texts = None
        if args.text_dir:
            texts = export_text_directory(args.source, args.text_dir, recursive=args.recursive)
            for row in texts:
                print(f"文字导出：{row['相对路径']} — {row['状态']}")
        write_inventory(args.output, files, pages, texts)
    except (OSError, ValueError) as exc:
        parser.exit(1, f"体检未完成：{exc}\n")
    print(f"完成：{len(files)} 份文件，{len(pages)} 页；结果：{args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
