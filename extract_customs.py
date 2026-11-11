"""Local text-layer prototype for the supplied Chinese customs declaration layout.

Usage: python extract_customs.py INPUT_PDF_OR_FOLDER OUTPUT_JSON
No OCR, network calls, source modification, or value propagation between rows.
"""
import argparse
import json
import re
from pathlib import Path
import pdfplumber

DETAIL_COLUMNS = ['文件名','项号','商品编号','商品名称','数量','总价','原产国','最终目的国','境内货源地']
HEADER_COLUMNS = ['文件名','境内发货人','境内发货人18位代码','境外收货人','生产销售单位','生产销售单位18位代码','合同协议号','提运单号','贸易国','运抵国','离境口岸','件数','毛重(千克)','净重(千克)']

def plain(s):
    # Only strip numeric codes. Parenthesized place names can be meaningful.
    return re.sub(r'[（(]\s*\d+\s*[）)]', '', s).strip()

def lines(words):
    groups = []
    for w in sorted(words, key=lambda w: (w['top'], w['x0'])):
        if not groups or abs(w['top']-groups[-1][0]['top']) > 2.5:
            groups.append([w])
        else:
            groups[-1].append(w)
    return [' '.join(w['text'] for w in sorted(g, key=lambda w:w['x0'])) for g in groups]

def region(words, left, right, top, bottom):
    return [w for w in words if left <= w['x0'] < right and top <= w['top'] < bottom]

def detail_region(words, field, left, right, top, bottom):
    selected = region(words, left, right, top, bottom)
    if field == '商品编号':
        # The name may begin before the original code column's right edge.
        selected = [w for w in selected if re.fullmatch(r'\d{10}', w['text'])]
    elif field == '境内货源地':
        # A right-aligned tax-treatment value starts inside the old fixed
        # origin rectangle in one broker's layout. Its right edge belongs to
        # the next (征免) column; do not merge it into the place name.
        selected = [w for w in selected if not (w['x0'] >= 745 and w['x1'] >= 780)]
    return selected

def normalized_words(page):
    scale = 842/page.width
    result = []
    for word in page.extract_words(x_tolerance=2, y_tolerance=2, return_chars=True):
        chars = word.pop('chars')
        # Some providers put the product name immediately next to the ten
        # digits. Split only when individual PDF characters prove the boundary.
        text = word['text']
        if (58 <= word['x0']*scale < 108 and re.fullmatch(r'\d{10}\D.+|\d{10}\D', text)
                and len(chars) >= 11 and ''.join(c['text'] for c in chars) == text
                and all(len(c['text']) == 1 for c in chars)):
            for part in (chars[:10], chars[10:]):
                result.append(dict(word, text=''.join(c['text'] for c in part),
                    x0=min(c['x0'] for c in part)*scale, x1=max(c['x1'] for c in part)*scale,
                    top=min(c['top'] for c in part)*scale, bottom=max(c['bottom'] for c in part)*scale))
        else:
            result.append(dict(word, x0=word['x0']*scale, x1=word['x1']*scale,
                 top=word['top']*scale, bottom=word['bottom']*scale))
    return result

def parse_document(path, filename=None):
    name = filename or path.name
    header = {k:None for k in HEADER_COLUMNS}
    header['文件名'] = name
    details, issues, evidence = [], [], []
    def issue(page, item, field, reason):
        issues.append(dict(文件名=name, 页码=page, 项号=item, 字段=field, 原因=reason))
    with pdfplumber.open(path) as pdf:
        page_count = len(pdf.pages)
        for pn, page in enumerate(pdf.pages, 1):
            words = normalized_words(page)
            if len(page.chars) < 50:
                issue(pn, None, '整页', '无足够文字层，需 OCR 或人工核对；此原型未执行 OCR')
                continue
            if page.annots:
                issue(pn, None, '整页', '存在 PDF 批注，文字层可能不包含修改内容，请核对可见页面')
            large_images = [im for im in page.images if im['width']*im['height'] > page.width*page.height*0.1]
            if large_images:
                issue(pn, None, '整页', '存在较大图片；文字层与图像可能不同，请核对扫描或手写修改')
            if pn == 1:
                fields = [
                    ('境内发货人','境内发货人',37,255),
                    ('境外收货人','境外收货人',37,255),
                    ('生产销售单位','生产销售单位',37,255),
                    ('合同协议号','合同协议号',37,255),
                    ('提运单号','提运单号',516,805),
                    ('贸易国','贸易国',255,383),
                    ('运抵国','运抵国',383,516),
                    ('离境口岸','离境口岸',637,805),
                    ('件数','件数',255,301),
                    ('毛重(千克)','毛重',301,383),
                    ('净重(千克)','净重',383,464),
                ]
                for field, label, left, right in fields:
                    anchors = [w for w in words if w['text'].startswith(label) and left<=w['x0']<right]
                    if len(anchors) != 1:
                        issue(pn, None, field, '字段标签未唯一定位，留空')
                        continue
                    y = anchors[0]['top']
                    vals = lines(region(words,left,right,y+5,y+20))
                    value = ' '.join(vals)
                    header[field] = value or None
                    if not value:
                        issue(pn, None, field, '原文值为空或未识别')
                    if field in ('境内发货人','生产销售单位'):
                        code_text = ' '.join(lines(region(words,left,right,y-2,y+5)))
                        codes = re.findall(r'(?<![A-Z0-9])[A-Z0-9]{18}(?![A-Z0-9])', code_text)
                        header[field+'18位代码'] = codes[0] if len(codes)==1 else None
                        if len(codes)!=1:
                            issue(pn,None,field+'18位代码','未找到唯一18位代码')
                for field in ('件数','毛重(千克)','净重(千克)'):
                    value = header[field]
                    if value and re.fullmatch(r'\d+(?:\.\d+)?',value):
                        header[field] = int(value) if field=='件数' else float(value)
                    else:
                        header[field] = None
                        issue(pn,None,field,'不是明确的数值，留空')
            table_heads = [w for w in words if w['text']=='项号']
            if len(table_heads)!=1:
                issue(pn,None,'未提取页面','未找到唯一报关单商品表头，本页可能是附页或其他版式；本页未自动提取，请确认是否包含应提取的商品')
                continue
            table_top = table_heads[0]['top']
            starts = [w for w in words if 35<=w['x0']<58 and w['top']>table_top and re.fullmatch(r'\d+',w['text'])]
            starts.sort(key=lambda w:w['top'])
            if not starts:
                issue(pn,None,'商品表','未发现商品行，请确认是否空白续页或版式不匹配')
            for k, start in enumerate(starts):
                item, y = int(start['text']), start['top']
                bottom = min(y+31, starts[k+1]['top']-1) if k+1<len(starts) else y+31
                row = {c:None for c in DETAIL_COLUMNS}
                row.update(文件名=name,项号=item)
                raw = {}
                for field,left,right in [('商品编号',58,108),('商品名称',104,334),('数量',334,421),('总价',421,504),('原产国',504,568),('最终目的国',568,660),('境内货源地',660,763)]:
                    cell_words = detail_region(words,field,left,right,y-1,bottom)
                    raw[field] = lines(cell_words)
                    if field=='数量':
                        # The third VISUAL line is around +22 pt; a blank second line stays blank.
                        selected = lines([w for w in cell_words if 17<=w['top']-y<29])
                        value = ''.join(selected).replace(' ','')
                        if len(selected)==1 and re.fullmatch(r'\d+(?:\.\d+)?[^\d\s.]+',value):
                            row[field]=value
                        else:
                            issue(pn,item,field,'第三行数量/单位缺失或不唯一，留空；不回退到重量')
                    elif field=='总价':
                        selected = lines([w for w in cell_words if 6<=w['top']-y<17])
                        value = ''.join(selected).replace(',','').replace(' ','')
                        if len(selected)==1 and re.fullmatch(r'\d+(?:\.\d+)?',value):
                            row[field]=float(value)
                        else:
                            issue(pn,item,field,'第二行总价缺失或不唯一，留空')
                    else:
                        selected=lines([w for w in cell_words if abs(w['top']-y)<3])
                        value=' '.join(selected)
                        if field in ('原产国','最终目的国','境内货源地'):
                            value=plain(value)
                        if field=='商品编号' and not re.fullmatch(r'\d{10}',value):
                            value=''
                        row[field]=value or None
                        if not value:
                            issue(pn,item,field,'首行值缺失或格式不符，留空')
                details.append(row)
                evidence.append(dict(文件名=name,页码=pn,项号=item,原文=raw))
        actual = [d['项号'] for d in details]
        if actual != list(range(1,len(actual)+1)):
            issue(None,None,'项号','项号存在重复、缺项或未从1开始，请核对所有页')
        for field in HEADER_COLUMNS:
            if header[field] is None and not any(i['字段']==field for i in issues):
                issue(1,None,field,'表头字段未取得，留空')
    return dict(header=header,details=details,issues=issues,evidence=evidence,pages=page_count)

def extract_batch(source):
    files = [source] if source.is_file() else sorted((p for p in source.rglob('*') if p.suffix.lower()=='.pdf'), key=lambda p:str(p).lower())
    result = dict(header_columns=HEADER_COLUMNS,detail_columns=DETAIL_COLUMNS,headers=[],details=[],issues=[],evidence=[],file_count=len(files),page_count=0)
    for path in files:
        name = path.name if source.is_file() else str(path.relative_to(source))
        try:
            data=parse_document(path,name)
            result['headers'].append(data['header'])
            result['page_count']+=data['pages']
            for key in ('details','issues','evidence'):
                result[key].extend(data[key])
        except Exception as exc:
            result['headers'].append({k:name if k=='文件名' else None for k in HEADER_COLUMNS})
            result['issues'].append(dict(文件名=name,页码=None,项号=None,字段='整份文件',原因=f'{type(exc).__name__}: {exc}'))
    return result

if __name__=='__main__':
    parser=argparse.ArgumentParser()
    parser.add_argument('source',type=Path)
    parser.add_argument('output',type=Path)
    args=parser.parse_args()
    if not args.source.exists():
        parser.error('输入路径不存在')
    data=extract_batch(args.source)
    args.output.parent.mkdir(parents=True,exist_ok=True)
    args.output.write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps({k:len(data[k]) if isinstance(data[k],list) else data[k] for k in ('file_count','page_count','headers','details','issues')},ensure_ascii=False))
