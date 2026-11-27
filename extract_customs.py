"""Read declaration cells in page coordinates and retain their source positions."""
import argparse
import json
import re
from pathlib import Path

import pdfplumber

DETAIL_COLUMNS = ['文件名','项号','商品编号','商品名称','数量','总价','原产国','最终目的国','境内货源地']
HEADER_COLUMNS = ['文件名','境内发货人','境内发货人18位代码','境外收货人','生产销售单位','生产销售单位18位代码','合同协议号','提运单号','贸易国','运抵国','离境口岸','件数','毛重(千克)','净重(千克)']
HEADERS = {'境内发货人':'境内发货人', '境外收货人':'境外收货人', '生产销售单位':'生产销售单位',
           '合同协议号':'合同协议号','提运单号':'提运单号','贸易国':'贸易国','运抵国':'运抵国','离境口岸':'离境口岸',
           '件数':'件数','毛重(千克)':'毛重','净重(千克)':'净重'}
COLS = ('项号','商品编号','商品名称','数量','总价','原产国','最终目的国','境内货源地','征免')
LABELS = (('项号','序号'), ('商品编号','商品编码'), ('商品名称','品名'),
          ('数量','成交数量'), ('单价','总价'), ('原产国','原产地'), ('最终目的','目的国'),
          ('境内货源地','货源地'), ('征免','征税'))
NUMBER = r'\d+(?:\.\d+)?'
CODE = r'(?<![A-Z0-9])[A-Z0-9]{18}(?![A-Z0-9])'


def plain(s):
    return re.sub(r'[（(]\s*\d+\s*[）)]', '', s).strip()


def center(w):
    return (w['x0'] + w['x1']) / 2


def bounds(words):
    if words:
        return (min(w['x0'] for w in words)-1, min(w['top'] for w in words)-1,
                max(w['x1'] for w in words)+1, max(w['bottom'] for w in words)+1)


def text_line(words):
    return ' '.join(w['text'] for w in sorted(words, key=lambda w:w['x0']))


def groups(words):
    result = []
    for word in sorted(words, key=lambda w:(w['top'], w['x0'])):
        if not result or abs(word['top'] - result[-1][0]['top']) > 2.5:
            result.append([])
        result[-1].append(word)
    return result


def normalized_words(page):
    scale = 842 / page.width
    result = []
    words = page.extract_words(x_tolerance=2, y_tolerance=2, return_chars=True)
    short_numbers = [w for w in words if re.fullmatch(r'\d{1,4}',w['text'])]
    for word in words:
        chars = word.pop('chars')
        parts = [chars]
        if (re.fullmatch(r'\d{10}[^\d\s].*', word['text']) and len(chars)>10
            and all(len(c['text'])==1 for c in chars)
            and any(abs(w['top']-word['top'])*scale<3 and 0<=word['x0']-w['x1']<chars[9]['x1']-chars[0]['x0']
                    for w in short_numbers)):
            parts = [chars[:10], chars[10:]]
        for part in parts:
            result.append(dict(text=''.join(c['text'] for c in part),
                               x0=min(c['x0'] for c in part)*scale, x1=max(c['x1'] for c in part)*scale,
                               top=min(c['top'] for c in part)*scale, bottom=max(c['bottom'] for c in part)*scale))
    return result


def read_header(page, words, header, issue):
    scale = 842 / page.width
    boxes = {}
    code_hits = page.search(CODE) or []
    for field,label in HEADERS.items():
        anchors = [w for w in words if w['text'].startswith(label)]
        if len(anchors)!=1:
            issue(1,None,field,'未找到唯一的字段标签')
            continue
        a = anchors[0]
        x, y = center(a), (a['top']+a['bottom'])/2
        xs = [e['x0']*scale for e in page.edges if e.get('orientation')=='v'
              and e['top']*scale<=y<=e['bottom']*scale]
        ys = [e['top']*scale for e in page.edges if e.get('orientation')=='h'
              and e['x0']*scale<=x<=e['x1']*scale and e['top']*scale>a['bottom']]
        left = max((v for v in xs if v<=a['x0']+1), default=a['x0']-4)
        right = min((v for v in xs if v>a['x1']-1), default=842)
        # Labels in the same band also bound the value when grid lines are absent.
        neighbours = [w['x0']-2 for w in words if abs(w['top']-a['top'])<=2.5
                      and w['x0']>a['x1']+2 and re.match(r'[\u4e00-\u9fff]',w['text'])]
        right = min([right]+neighbours)
        bottom = min(ys,default=a['bottom']+15)
        vals = [w for w in words if left<=center(w)<right and a['top']+4<w['top']<bottom-1]
        value = ' '.join(text_line(g) for g in groups(vals))
        boxes[field] = bounds(vals)
        if field in ('件数','毛重(千克)','净重(千克)'):
            pat = r'\d+' if field=='件数' else NUMBER
            number = value.replace(',','')
            header[field] = (int(number) if field=='件数' else float(number)) if re.fullmatch(pat,number) else None
        else:
            header[field] = value or None
        if header[field] is None:
            issue(1,None,field,'单元格为空或值无法唯一确定')
        if field in ('境内发货人','生产销售单位'):
            key = field+'18位代码'
            hits = [h for h in code_hits if abs(h['top']*scale-a['top'])<5
                    and left<=center(h)*scale<right]
            if len(hits)==1:
                h = hits[0]
                header[key] = h['text']
                # Codes are often printed directly inside parentheses. Use the
                # glyph edges horizontally so the crop contains the 18 characters.
                boxes[key] = (h['x0']*scale,h['top']*scale-1,h['x1']*scale,h['bottom']*scale+1)
            else:
                issue(1,None,key,'未找到唯一18位代码')
    return boxes


def table_layout(words):
    for head in (w for w in words if w['text'] in LABELS[0]):
        band = [w for w in words if abs(w['top']-head['top'])<4]
        cells = []
        for aliases in LABELS:
            matched = [w for w in band if w['text'].startswith(aliases)]
            if not matched and aliases==LABELS[-1]:
                # The treatment header is optional. Use a visible treatment
                # value to locate the right edge of the origin column.
                markers = [w for w in words if w['top']>head['bottom']
                           and w['x0']>cells[-1]['x0'] and w['text'] in ('照章征税','全免','全征')]
                matched = markers[:1]
            if len(matched)!=1:
                break
            cells.append(matched[0])
        if len(cells)==len(COLS) and all(a['x0']<b['x0'] for a,b in zip(cells,cells[1:])):
            return dict(zip(COLS,cells)), head['bottom']
    return None, None


def column_regions(layout):
    # Number, description and right-aligned amounts need different boundaries.
    first = layout['项号']
    code = layout['商品编号']
    qty, price = layout['数量'],layout['总价']
    cuts = [first['x0']-20, (center(first)+center(code))/2, code['x1']+6,
            qty['x0']-14, price['x0']-(price['x1']-price['x0'])*0.12,
            layout['原产国']['x0']-4,layout['最终目的国']['x0']-5,
            layout['境内货源地']['x0']-3,layout['征免']['x0']-7,842]
    return dict(zip(COLS,zip(cuts,cuts[1:])))


def page_mark(page):
    mark = re.search(r'页码\s*/\s*页数\s*[:：]?\s*(\d+)\s*/\s*(\d+)',page.extract_text() or '')
    return tuple(map(int,mark.groups())) if mark else None


def row_candidates(words, regions, top):
    floor = min((w['top'] for w in words if w['top']>top and
                 w['text'].startswith(('特殊关系确认','报关人员','总价格：','总价格:'))),default=float('inf'))
    left, right = regions['项号']
    code_left, code_right = regions['商品编号']
    body = [w for w in words if top < w['top'] < floor]
    items = sorted([w for w in body if left<=center(w)<right and re.fullmatch(r'\d{1,4}',w['text'])],key=lambda w:w['top'])
    codes = [w for w in body if code_left<=center(w)<code_right and re.fullmatch(r'\d{10}',w['text'])]
    return floor, items, codes


def read_row(body, item, code, regions, filename, pn, issue):
    y = item['top']
    row = dict.fromkeys(DETAIL_COLUMNS)
    row.update(文件名=filename,项号=int(item['text']))
    boxes = {'项号':bounds([item]),'商品编号':bounds(code)}
    raw = {}
    row['商品编号'] = code[0]['text'] if len(code)==1 else None
    first = [w for w in body if abs(w['top']-y)<3]
    name_left = code[0]['x1']-0.5 if len(code)==1 else regions['商品名称'][0]
    def in_column(word, field):
        left,right = regions[field]
        # Right-aligned quantities and prices share an end point even when
        # their left edge crosses into the previous column.
        x = center(word)
        if field in ('数量','总价') and re.fullmatch(r'[\d,.]+(?:[^\d\s|]+)?', word['text']):
            x = word['x1']-0.1
        return left<=x<right
    price_bands = groups([w for w in body if in_column(w,'总价')])
    currency = [b for b in price_bands if not re.search(r'\d',text_line(b)) and b[0]['top']>y+4]
    amounts = [b for b in price_bands if b[0]['top']>y+3
               and (not currency or b[0]['top']<currency[0][0]['top']-2)
               and re.fullmatch(NUMBER,text_line(b).replace(' ','').replace(',',''))]
    # Missing currency can still leave an unambiguous third slot. Estimate its
    # baseline from the two printed price levels, never from the last quantity.
    quantity_y = None
    if len(currency)==1:
        quantity_y = currency[0][0]['top']
    elif not currency and len(amounts)==1 and any(abs(b[0]['top']-y)<3 for b in price_bands):
        quantity_y = y+2*(amounts[0][0]['top']-y)
    for field in DETAIL_COLUMNS[3:]:
        left,right = regions[field]
        if field=='商品名称':
            selected = [w for w in first if w['x0']>=name_left and center(w)<right and not in_column(w,'数量')]
        else:
            selected = [w for w in body if in_column(w,field)]
        bands = groups(selected)
        raw[field] = [text_line(b) for b in bands]
        chosen = [w for w in selected if abs(w['top']-y)<3]
        if field in ('数量','总价'):
            # The price column carries unit price / total / currency. Their actual
            # baselines determine the slots, even when a statutory quantity is blank.
            if field=='数量':
                chosen = [w for w in selected if quantity_y is not None and abs(w['top']-quantity_y)<3]
                if quantity_y is None and len(bands)==3 and abs(bands[0][0]['top']-y)<3 and all(
                    re.fullmatch(NUMBER+r'[^\d\s.]+',text_line(b).replace(' ','')) for b in bands):
                    chosen = bands[2]
                value = text_line(chosen).replace(' ','')
                if re.fullmatch(NUMBER+r'[^\d\s.]+',value):
                    row[field] = value
            else:
                chosen = amounts[0] if len(amounts)==1 else []
                if chosen:
                    row[field] = float(text_line(chosen).replace(' ','').replace(',',''))
        else:
            value = text_line(chosen)
            row[field] = (plain(value) if field in ('原产国','最终目的国','境内货源地') else value) or None
        boxes[field] = bounds(chosen)
    for field in DETAIL_COLUMNS[2:]:
        if row[field] is None:
            issue(pn,row['项号'],field,'原始行中未找到唯一符合条件的值')
    return row, dict(文件名=filename,页码=pn,项号=row['项号'],行位置=y,原文=raw,字段位置=boxes)


def parse_document(path, filename=None):
    name = filename or path.name
    header = dict.fromkeys(HEADER_COLUMNS)
    header['文件名'] = name
    details, issues, evidence, page_info = [], [], [], []
    header_boxes = {}
    def issue(page,item,field,reason):
        issues.append(dict(文件名=name,页码=page,项号=item,字段=field,原因=reason))
    previous = None
    declared_total = None
    with pdfplumber.open(path) as pdf:
        page_count = len(pdf.pages)
        for pn,page in enumerate(pdf.pages,1):
            words = normalized_words(page)
            mark = page_mark(page)
            if pn==1:
                declared_total = mark[1] if mark and mark[0]==1 else None
                header_boxes = read_header(page,words,header,issue)
            if len(page.chars)<50:
                issue(pn,None,'整页','文字层不足，需核对原图')
                previous = None
                continue
            if page.annots:
                issue(pn,None,'整页','存在 PDF 批注，请核对可见页面')
            if any(im['width']*im['height']>page.width*page.height*.1 for im in page.images):
                issue(pn,None,'整页','存在较大图片，请核对可见页面')
            layout, top = table_layout(words)
            continuation = layout is None
            if declared_total and pn>declared_total:
                issue(pn,None,'未提取页面','已超出首页标示的报关单页数，请确认附页内容')
                previous = None
                continue
            if continuation:
                total = declared_total or (mark[1] if mark else None)
                if not previous or not total or pn>total or (mark and mark!=(pn,total)):
                    issue(pn,None,'未提取页面','缺少可验证的报关单列名或续页关系')
                    previous = None
                    continue
                layout, prior_item, prior_code_x = previous
                top = 0
            regions = column_regions(layout)
            floor, starts, codes = row_candidates(words,regions,top)
            def paired(start):
                return [w for w in codes if abs(w['top']-start['top'])<3]
            if continuation:
                numbers = [int(w['text']) for w in starts]
                title = ''.join(w['text'] for w in sorted(words,key=lambda w:(w['top'],w['x0']))
                                if starts and w['top']<starts[0]['top']-3)
                attachment = re.search(r'装箱单|发票|合同清单|申报要素|PACKINGLIST|INVOICE',title,re.I)
                aligned = [s for s in starts if len(paired(s))==1 and abs(paired(s)[0]['x0']-prior_code_x)<10]
                # Require a connection to the previous table and a populated
                # price column; coincident page numbers alone do not identify it.
                price_left,price_right = regions['总价']
                prices = [w for w in words if price_left<=w['x1']-0.1<price_right
                          and re.fullmatch(NUMBER,w['text'].replace(',',''))]
                supported = [s for s in aligned if any(abs(w['top']-s['top'])<3 for w in prices)]
                if (attachment or not numbers or numbers[0]!=prior_item+1 or not supported
                    or (not mark and (len(supported)!=len(starts) or len(codes)!=len(starts)
                        or numbers!=list(range(prior_item+1,prior_item+1+len(starts)))))):
                    issue(pn,None,'未提取页面','续页的项号、编码或列位置未能相互印证')
                    previous = None
                    continue
            if not starts or len(codes)!=len(starts) or any(len(paired(s))!=1 for s in starts):
                issue(pn,None,'项号列异常','项号与同一行十位商品编号未能逐一对应，明细可能不完整')
            page_info.append(dict(文件名=name,页码=pn,印刷页码=mark,编码范围=regions['商品编号'],表起点=top,表终点=min(floor,page.height*842/page.width)))
            for index,start in enumerate(starts):
                if continuation and start not in supported:
                    issue(pn,int(start['text']),'项号列异常','续页中该行的项号、编码与价格位置未确认，已保留其余可定位的行')
                    continue
                bottom = starts[index+1]['top']-2.5 if index+1<len(starts) else floor
                # A horizontal row divider can separate a final item from the footer.
                x=center(start)*page.width/842
                dividers = [e['top']*842/page.width for e in page.edges if e.get('orientation')=='h'
                            and e['x0']<=x<=e['x1'] and e['width']>page.width*.5
                            and e['top']*842/page.width>start['bottom']+2]
                bottom = min(bottom,min(dividers,default=bottom))
                body = [w for w in words if start['top']-2.8<=w['top']<bottom]
                row, ev = read_row(body,start,paired(start),regions,name,pn,issue)
                details.append(row)
                evidence.append(ev)
            previous = (layout,int(starts[-1]['text']),paired(starts[-1])[0]['x0']) if starts and len(paired(starts[-1]))==1 else None
    for expected,ev in enumerate(evidence,1):
        if ev['项号']!=expected:
            issue(ev['页码'],ev['项号'],'项号列异常','原件项号重复、不连续或起点有误，须核对原始行')
    if declared_total and page_count<declared_total:
        for missing_page in range(page_count+1,declared_total+1):
            issue(missing_page,None,'未提取页面',f'原件标示共 {declared_total} 页，实际仅收到 {page_count} 页')
    for field in HEADER_COLUMNS[1:]:
        if header[field] is None and not any(i['字段']==field for i in issues):
            issue(1,None,field,'未取得表头字段')
    return dict(header=header,details=details,issues=issues,evidence=evidence,pages=page_count,
                header_boxes=header_boxes,page_info=page_info)

def extract_batch(source):
    files = [source] if source.is_file() else sorted((p for p in source.rglob('*') if p.suffix.lower()=='.pdf'), key=lambda p:str(p).lower())
    result = dict(header_columns=HEADER_COLUMNS,detail_columns=DETAIL_COLUMNS,headers=[],details=[],issues=[],evidence=[],header_boxes={},page_info=[],file_count=len(files),page_count=0)
    for path in files:
        name = path.name if source.is_file() else str(path.relative_to(source))
        try:
            data=parse_document(path,name)
            result['headers'].append(data['header'])
            result['header_boxes'][name] = data['header_boxes']
            result['page_count']+=data['pages']
            for key in ('details','issues','evidence','page_info'):
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
