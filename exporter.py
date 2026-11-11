"""将报关单提取结果导出为标准 Excel 和可离线使用的核对页面。

export_all(data, output_dir) 接受 audit_customs 结构的合并数据；output_dir
由调用方为本次运行独立创建。返回各结果文件的 Path。
"""
from pathlib import Path
import html
import ntpath

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


HEADER_COLUMNS = ['文件名', '境内发货人', '境内发货人18位代码', '境外收货人', '生产销售单位', '生产销售单位18位代码', '合同协议号', '提运单号', '贸易国', '运抵国', '离境口岸', '件数', '毛重(千克)', '净重(千克)']
DETAIL_COLUMNS = ['文件名', '项号', '商品编号', '商品名称', '数量', '总价', '原产国', '最终目的国', '境内货源地']
SUMMARY_COLUMNS = ['文件名', '页数', '商品项数', '应核对字段', '已核对字段', '读数一致', '读数不同或缺失', '低分但读数一致', '规则待核对', '文件状态', 'OCR错误']
AUDIT_COLUMNS = ['核对编号', '文件名', '页码', '项号', '字段', '提取值', 'OCR读数', 'OCR分数', '核对状态', '原因', '人工确认值', '复核意见']
RULE_COLUMNS = ['文件名', '检查项', '页码', '项号', '实际结果', '参考值', '状态', '说明']
CODE_COLUMNS = {'境内发货人18位代码', '生产销售单位18位代码', '商品编号', '合同协议号', '提运单号', '文件名', '数量'}


def _filename(value):
    # Keep relative subfolders so identically named PDFs remain distinguishable.
    name = str(value or '')
    return name[:-4] if name.lower().endswith('.pdf') else name


def _cell_value(value):
    if isinstance(value, (str, int, float, bool)) or value is None:
        return value
    return str(value)


def _sheet(wb, name, columns, rows, widths=None, business=False):
    ws = wb.create_sheet(name)
    ws.sheet_view.showGridLines = False
    ws.append(columns)
    for row in rows:
        values = []
        for key in columns:
            value = row.get(key)
            if business and key == '文件名':
                value = _filename(value)
            if key in CODE_COLUMNS and value is not None:
                value = str(value)
            values.append(_cell_value(value))
        ws.append(values)
    for cell in ws[1]:
        cell.fill = PatternFill('solid', fgColor='4472C4')
        cell.font = Font(name='微软雅黑', bold=True, color='FFFFFF', size=10)
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.row_dimensions[1].height = 32
    for index, title in enumerate(columns, 1):
        ws.column_dimensions[get_column_letter(index)].width = (widths or {}).get(title, min(38, max(13, len(title) * 2 + 5)))
    for row in ws.iter_rows(min_row=2):
        for cell, key in zip(row, columns):
            cell.font = Font(name='微软雅黑', size=10)
            cell.alignment = Alignment(vertical='center')
            # Treat extracted PDF text as data even when it starts with '='.
            if isinstance(cell.value, str):
                cell.data_type = 's'
            if key in CODE_COLUMNS:
                cell.number_format = '@'
            elif key in {'总价', '毛重(千克)', '净重(千克)'}:
                cell.number_format = '#,##0.00'
            elif key == 'OCR分数':
                cell.number_format = '0.0000'
            if key in {'人工确认值', '复核意见'}:
                cell.fill = PatternFill('solid', fgColor='FFF4CE')
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions
    return ws


def _workbook(path, name, columns, rows, widths):
    wb = Workbook()
    wb.remove(wb.active)
    _sheet(wb, name, columns, rows, widths, business=True)
    wb.save(path)


def _report(path, data):
    wb = Workbook()
    wb.remove(wb.active)
    audit = data.get('audit', [])
    widths = {'文件名': 27, '字段': 25, '提取值': 38, 'OCR读数': 38, '原因': 55, '说明': 55, '人工确认值': 28, '复核意见': 40, 'OCR错误': 50}
    _sheet(wb, '文件总览', SUMMARY_COLUMNS, data.get('summaries', []), widths)
    _sheet(wb, '优先核对', AUDIT_COLUMNS, [r for r in audit if r.get('核对状态') != '两路一致'], widths)
    _sheet(wb, '全部字段核对', AUDIT_COLUMNS, audit, widths)
    _sheet(wb, '规则核对', RULE_COLUMNS, data.get('rules', []), widths)
    instructions = [
        {'项目': '业务结果', '说明': '两份业务表按标准列顺序导出；本工作簿保存核对结果。'},
        {'项目': '优先核对', '说明': '先看待核对字段，用核对编号在配套 HTML 中查找截图；OCR 低分不一定代表错误。'},
        {'项目': '两路一致', '说明': '程序分别读取 PDF 文字和对应截图；相同表示两次读出的值一样，但如果取错了同一行或漏掉整行，仍可能出错。请核对总项数，并对照原 PDF 抽查关键字段。'},
        {'项目': '填写结果', '说明': '黄色列可填写人工确认值和复核意见；填写后不会自动回写业务表，也不会同步其他工作表。'},
        {'项目': '修正业务数据', '说明': '确认错误后，请在对应业务 Excel 另行修改；保留提取记录和原图以便追溯。'},
        {'项目': '识别限制', '说明': '图像核对依赖文字层定位；扫描件、复杂版式、手写修改可能无法正确识别，需结合原 PDF 核对。'},
    ]
    ws = _sheet(wb, '使用说明', ['项目', '说明'], instructions, {'项目': 20, '说明': 110})
    for row in ws.iter_rows(min_row=2):
        row[1].alignment = Alignment(vertical='center', wrap_text=True)
        ws.row_dimensions[row[0].row].height = 34
    wb.save(path)


def _html_report(path, data):
    esc = lambda value: html.escape('' if value is None else str(value))
    audit = data.get('audit', [])
    review = [r for r in audit if r.get('核对状态') != '两路一致']
    content = ['<!doctype html><html lang="zh-CN"><meta charset="utf-8"><title>报关单逐字段原图核对</title>',
               '<style>body{font:16px Microsoft YaHei,Arial;margin:32px;background:#f3f6fa;color:#172b4d}article{background:white;padding:18px;margin:16px 0;border-left:5px solid #9aa8b5}article.review{border-color:#dc762a}article[hidden]{display:none}img{max-width:100%;border:1px solid #ddd;display:block;margin-top:12px}p{line-height:1.7}.meta{color:#546477}.filters{display:flex;flex-wrap:wrap;gap:12px;align-items:end;background:#fff;padding:14px}.filters label{display:flex;flex-direction:column;gap:4px}.filters input,.filters select,.filters button{font:inherit;padding:7px;border:1px solid #8295ad;border-radius:4px}</style>',
               '<h1>报关单逐字段原图核对</h1>',
               f'<p>表头结果 {len(data.get("headers", []))} 行，商品明细 {len(data.get("details", []))} 条；{len(audit)} 个核对字段，{len(review)} 个待核对。具体失败文件请看核对报告“文件总览”。</p>',
               '<p>程序分别读取 PDF 文字和对应截图；两次读数相同，表示读到的内容一样。但如果两次都取错了同一行，或漏掉整行，仍可能出错。请核对总项数，并对照原 PDF 抽查关键字段。OCR 分数不是正确率。可根据核对编号搜索截图；在核对报告填写的意见不会自动修改业务表。</p>',
               '<p>截图内嵌在本页面，可离线查看。优先显示待核对字段。</p>',
               '<section class="filters" aria-label="筛选核对项目"><label>搜索文件名、字段或编号<input id="search" type="search" placeholder="输入关键词"></label><label>核对状态<select id="status"><option value="all">全部</option><option value="review">待核对或未核对</option><option value="matched">两路一致</option></select></label><label>字段<select id="field"><option value="">全部字段</option>' + ''.join(f'<option value="{esc(field)}">{esc(field)}</option>' for field in sorted({str(r.get('字段')) for r in audit if r.get('字段')})) + '</select></label><button id="clear" type="button">清除筛选</button><span id="visible-count" role="status"></span></section>']
    for record in sorted(audit, key=lambda r: (r.get('核对状态') == '两路一致', r.get('核对编号') or 0)):
        number = esc(record.get('核对编号'))
        state = record.get('核对状态', '未核对')
        label = '表头' if record.get('项号') is None else '项号 ' + esc(record.get('项号'))
        cls = 'review' if state != '两路一致' else ''
        content.append(f'<article id="field-{number}" class="{cls}" data-field="{esc(record.get("字段"))}"><h3>#{number} · {esc(record.get("字段"))} · {esc(state)}</h3><p class="meta">{esc(record.get("文件名"))} · 第 {esc(record.get("页码"))} 页 · {label}</p><p>提取值：<strong>{esc(record.get("提取值"))}</strong><br>图像 OCR：<strong>{esc(record.get("OCR读数"))}</strong><br>OCR 分数：{esc(record.get("OCR分数"))}<br>{esc(record.get("原因"))}</p>')
        image = record.get('截图')
        if image:
            content.append(f'<img alt="该字段的 PDF 原图截图" src="data:image/png;base64,{esc(image)}">')
        else:
            content.append('<p>未生成截图，请直接查看原 PDF。</p>')
        content.append('</article>')
    content.append('''<script>
    (() => {
      const search = document.getElementById('search');
      const status = document.getElementById('status');
      const field = document.getElementById('field');
      const count = document.getElementById('visible-count');
      const cards = Array.from(document.querySelectorAll('article'));
      function filter() {
        const query = search.value.trim().toLocaleLowerCase();
        let visible = 0;
        for (const card of cards) {
          const review = card.classList.contains('review');
          const matches = (!query || card.textContent.toLocaleLowerCase().includes(query)) &&
            (status.value === 'all' || (status.value === 'review' ? review : !review)) &&
            (!field.value || card.dataset.field === field.value);
          card.hidden = !matches;
          if (matches) visible++;
        }
        count.textContent = `显示 ${visible} / ${cards.length} 项`;
      }
      search.addEventListener('input', filter);
      status.addEventListener('change', filter);
      field.addEventListener('change', filter);
      document.getElementById('clear').addEventListener('click', () => {
        search.value = ''; status.value = 'all'; field.value = ''; filter();
      });
      filter();
    })();
    </script></html>''')
    path.write_text('\n'.join(content), encoding='utf-8')


def export_all(data, output_dir):
    """导出三个 Excel 与离线 HTML；已有同名文件时报错，防止覆盖。"""
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    targets = {
        'headers': output_dir / '预录单收发件人信息.xlsx',
        'details': output_dir / '预录单商品明细提取.xlsx',
        'report': output_dir / '提取核对报告.xlsx',
        'images': output_dir / '逐字段原图核对.html',
    }
    existing = [str(path) for path in targets.values() if path.exists()]
    if existing:
        raise FileExistsError('结果文件已存在，请选择新的结果目录：' + '、'.join(existing))
    _workbook(targets['headers'], '收发件人信息', HEADER_COLUMNS, data.get('headers', []), {'文件名': 25, '境内发货人': 32, '境外收货人': 52, '生产销售单位': 32})
    _workbook(targets['details'], '商品明细', DETAIL_COLUMNS, data.get('details', []), {'文件名': 25, '商品名称': 20, '境内货源地': 20})
    _report(targets['report'], data)
    _html_report(targets['images'], data)
    return targets
