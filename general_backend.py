"""面向界面和命令行的模式适配；通用功能与报关单规则相互独立。"""
from __future__ import annotations

import tempfile
import time
from pathlib import Path


def run_mode(mode, input_dir, output_dir, rules_path='', progress=None):
    progress = progress or (lambda message: None)
    if mode == '报关单提取':
        from backend import run_batch
        return run_batch(input_dir, output_dir, progress)

    source = Path(input_dir).expanduser().resolve()
    destination = Path(output_dir).expanduser().resolve()
    if not source.is_dir() or not destination.is_dir():
        raise ValueError('PDF 文件夹和结果保存文件夹必须已经存在。')
    if destination == source or destination.is_relative_to(source):
        raise ValueError('结果保存文件夹不能位于 PDF 输入文件夹中，请选择其他位置。')

    if mode == 'PDF 体检及文本导出':
        from pdf_inventory import inspect_directory, write_inventory, export_text_directory
        progress('正在检查 PDF 页数与文字层……')
        files, pages = inspect_directory(source, recursive=True,
                                         progress=lambda n, total, path: progress(f'已检查 {n}/{total}：{path}'))
        if not files:
            raise ValueError('所选文件夹及其子文件夹中没有找到 PDF。')
        target = Path(tempfile.mkdtemp(prefix='PDF体检_' + time.strftime('%Y%m%d_%H%M%S') + '_', dir=destination))
        workbook = target / 'PDF文件体检.xlsx'
        progress('正在逐页导出文字层；无文字层的页面不会自动 OCR。')
        exports = export_text_directory(source, target / '逐页文本', recursive=True,
                                        progress=lambda n, total, path: progress(f'已导出 {n}/{total}：{path}'))
        write_inventory(workbook, files, pages, exports=exports)
        review_paths = {row['相对路径'] for row in files if row.get('状态') != '正常'}
        review_paths.update(row['相对路径'] for row in exports if row.get('状态') != '完成')
        return {'结果文件夹': str(target), 'PDF总数': len(files), '主要结果': str(workbook),
                '文本目录': str(target / '逐页文本'),
                '文本导出文件数': sum(bool(row.get('输出文件')) for row in exports),
                '未导出文件数': sum(not row.get('输出文件') for row in exports),
                '待查看文件数': len(review_paths)}

    if mode == 'PDF 表格提取':
        from pdf_tables import run_tables
        result = run_tables(source, destination, progress_callback=progress)
        return {'结果文件夹': str(result['result_dir']), '主要结果': str(result['workbook']),
                'PDF总数': result['pdf_count'], '表格数': result['table_count'],
                '单元格数': result['cell_count'], '异常提示数': result['issue_count']}

    if mode == '自定义字段提取':
        from field_extractor import extract_fields
        result = extract_fields(source, destination, rules_path, progress=progress)
        return {'结果文件夹': str(result['output_dir']), '主要结果': str(result['workbook']),
                'PDF总数': result['pdf_count'], '匹配数': result['match_count'],
                '异常提示数': result['warning_count']}
    raise ValueError(f'未知处理模式：{mode}')
