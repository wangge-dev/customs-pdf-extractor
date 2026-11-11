"""离线批量处理：每个文件独立提取、OCR 核对，最后生成 Excel。"""
from __future__ import annotations

import collections
import os
import tempfile
import time
from concurrent.futures import ProcessPoolExecutor, as_completed
from pathlib import Path

from audit_customs import audit, make_engine, normalize
from extract_customs import DETAIL_COLUMNS, HEADER_COLUMNS, extract_batch
from exporter import export_all

_ENGINE = None


def _worker(path, relative):
    global _ENGINE
    try:
        if _ENGINE is None:
            os.environ['CUSTOMS_OCR_THREADS'] = '2'
            _ENGINE = make_engine('v6_medium')
        result = audit(Path(path), 'v6_medium', engine=_ENGINE, relative_name=relative)
        return {'ok': True, 'data': result}
    except Exception as exc:
        return {'ok': False, 'error': f'{type(exc).__name__}: {exc}'}


def _fallback(path, relative, error):
    """保留文字层读取结果，但把未运行的 OCR 明确标出。"""
    data = extract_batch(Path(path))
    # A file may disappear after the initial scan, for example if a network
    # share disconnects. Preserve one visible placeholder and continue the
    # rest of the batch instead of indexing an empty headers list.
    if not data.get('headers'):
        data['headers'] = [{field: relative if field == '文件名' else None for field in HEADER_COLUMNS}]
    for name in ('headers', 'details', 'evidence', 'issues'):
        for row in data[name]:
            row['文件名'] = relative
    data['audit'] = []
    data['rules'] = [{'文件名': relative, '检查项': 'OCR 执行', '页码': None,
                      '项号': None, '实际结果': error, '参考值': 'OCR 完成',
                      '状态': '待核对', '说明': 'OCR 未完成，请查看原 PDF。'}]
    return data


def _merge_file(data, relative, ocr_error, combined):
    header = data['headers'][0]
    details = data.get('details', [])
    evidence = data.get('evidence', [])
    expected = [(None, 1, field, header.get(field)) for field in HEADER_COLUMNS[1:]]
    if len(details) != len(evidence):
        ocr_error = (ocr_error + '；' if ocr_error else '') + '明细与来源定位数量不一致'
    for index, row in enumerate(details):
        e = evidence[index] if index < len(evidence) else {}
        page = e.get('页码')
        if e.get('项号') != row.get('项号'):
            ocr_error = (ocr_error + '；' if ocr_error else '') + '明细来源项号不一致'
        expected.extend((row.get('项号'), page, field, row.get(field)) for field in DETAIL_COLUMNS[1:])

    records = list(data.get('audit', []))
    expected_keys = collections.Counter((i, p, f) for i, p, f, _ in expected)
    actual_keys = collections.Counter((r.get('项号'), r.get('页码'), r.get('字段')) for r in records)
    coverage_bad = expected_keys != actual_keys
    if coverage_bad and not ocr_error:
        ocr_error = '部分字段的图像核对未完成或定位不一致'
    missing = expected_keys - actual_keys
    for item, page, field, value in expected:
        key = item, page, field
        if missing[key]:
            records.append({'文件名': relative, '页码': page, '项号': item, '字段': field,
                            '提取值': value, 'OCR读数': '', 'OCR分数': None,
                            '核对状态': '未核对', '原因': ocr_error or '未收到 OCR 结果',
                            '人工确认值': None, '复核意见': None, '截图': ''})
            missing[key] -= 1
    for row in records:
        row['文件名'] = relative
        row['核对编号'] = len(combined['audit']) + 1
        combined['audit'].append(row)
    for key in ('headers', 'details', 'rules'):
        for row in data.get(key, []):
            row['文件名'] = relative
            combined[key].append(row)
    rules = data.get('rules', [])
    matched = [r for r in records if normalize(r['字段'], r['提取值']) is not None and
               normalize(r['字段'], r['提取值']) == normalize(r['字段'], r['OCR读数'])]
    different = len(records) - len(matched)
    low_same = sum(r['核对状态'] != '两路一致' for r in matched)
    rule_issues = sum(r.get('状态') != '相符' for r in rules)
    needs_review = bool(different or low_same or rule_issues or coverage_bad or ocr_error or not details)
    combined['summaries'].append({
        '文件名': relative, '页数': data.get('page_count'), '商品项数': len(details),
        '应核对字段': len(expected), '已核对字段': sum(r.get('OCR分数') is not None and normalize(r['字段'], r.get('OCR读数')) is not None for r in records),
        '读数一致': len(matched), '读数不同或缺失': different, '低分但读数一致': low_same,
        '规则待核对': rule_issues, '文件状态': '待核对' if needs_review else '机器一致，待抽查',
        'OCR错误': ocr_error,
    })
    # A worker can finish normally while individual crops have no OCR result,
    # or its field coverage does not match the extracted rows.  Those files
    # must count as incomplete as well as a worker that raised an exception.
    ocr_incomplete = bool(ocr_error or coverage_bad or any(
        r.get('OCR分数') is None or normalize(r['字段'], r.get('OCR读数')) is None for r in records))
    return needs_review, ocr_incomplete


def run_batch(input_dir, output_dir, progress_callback=lambda _: None, workers=None):
    """处理文件夹（含子文件夹）。每次新建结果目录，不覆盖已有文件。"""
    source = Path(input_dir).expanduser().resolve()
    destination = Path(output_dir).expanduser().resolve()
    if not source.is_dir() or not destination.is_dir():
        raise ValueError('PDF 文件夹和结果保存文件夹必须已经存在。')
    files = sorted((p for p in source.rglob('*') if p.is_file() and p.suffix.lower() == '.pdf'),
                   key=lambda p: str(p.relative_to(source)).casefold())
    if not files:
        raise ValueError('所选文件夹及其子文件夹中没有找到 PDF，请先解压压缩包。')
    count = len(files)
    preferred = min(4, max(1, (os.cpu_count() or 2) // 2), count)
    workers = preferred if workers is None else max(1, min(int(workers), count, 4))
    target = Path(tempfile.mkdtemp(prefix='报关单提取_' + time.strftime('%Y%m%d_%H%M%S') + '_', dir=str(destination)))
    progress_callback(f'找到 {count} 份 PDF，使用 {workers} 个工作进程。结果目录：{target}')
    relative = [str(p.relative_to(source)) for p in files]
    results = [None] * count
    completed = 0
    # Even a one-file run exercises the same packaged child-process path.
    with ProcessPoolExecutor(max_workers=workers) as pool:
        futures = {pool.submit(_worker, str(p), rel): (idx, p, rel)
                   for idx, (p, rel) in enumerate(zip(files, relative))}
        for future in as_completed(futures):
            idx, path, rel = futures[future]
            try:
                result = future.result()
            except Exception as exc:
                result = {'ok': False, 'error': f'工作进程异常：{type(exc).__name__}: {exc}'}
            results[idx] = result
            completed += 1
            progress_callback(f'已处理 {completed}/{count}：{rel}' + ('' if result['ok'] else '（OCR 未完成，保留待核对记录）'))

    combined = {'header_columns': HEADER_COLUMNS, 'detail_columns': DETAIL_COLUMNS,
                'headers': [], 'details': [], 'audit': [], 'rules': [], 'summaries': []}
    ocr_failures = 0
    for path, rel, result in zip(files, relative, results):
        if result and result['ok']:
            data, error = result['data'], ''
        else:
            error = (result or {}).get('error', '未获得处理结果')
            data = _fallback(path, rel, error)
        _, incomplete = _merge_file(data, rel, error, combined)
        if incomplete:
            ocr_failures += 1
    progress_callback('正在生成两份业务表和核对报告……')
    output = export_all(combined, target)
    review = sum(x['文件状态'] == '待核对' for x in combined['summaries'])
    info = {'结果文件夹': str(target), 'PDF总数': count, '商品明细行数': len(combined['details']),
            '待核对文件数': review, 'OCR未完成文件数': ocr_failures,
            '收发件人表': str(output['headers']), '商品明细表': str(output['details']),
            '核对报告': str(output['report']), '原图核对页': str(output['images'])}
    progress_callback(f'完成：{count} 份 PDF，{len(combined["details"])} 条明细；{review} 份需要核对。')
    # Usable by operators and by unattended, isolated end-to-end verification.
    (target / '本次运行结果.txt').write_text('\n'.join(f'{k}：{v}' for k, v in info.items()), encoding='utf-8-sig')
    return info
