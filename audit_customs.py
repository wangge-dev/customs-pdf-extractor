"""Field-by-field visual OCR comparison and non-blocking consistency checks.

This sample-layout prototype uses text positions to locate visual crops. The
OCR reading is independent of the text value, but location errors can be shared.
Matching readings are not proof of correctness and never mark manual approval.
"""
import argparse
import base64
import html
import io
import json
import os
import re
import sys
import time
import unicodedata
import ctypes
from decimal import Decimal, InvalidOperation
from pathlib import Path

ROOT=Path(__file__).parent
RESOURCE_ROOT=Path(getattr(sys, '_MEIPASS', ROOT))
# Load the application's own runtime in this process, without changing Windows.
_runtime_path=RESOURCE_ROOT/'runtime'/'msvcp140.dll'
_msvc=ctypes.WinDLL(str(_runtime_path)) if sys.platform=='win32' and _runtime_path.exists() else None
import pdfplumber
from PIL import ImageOps
from extract_customs import extract_batch, normalized_words, plain, center, HEADER_COLUMNS, DETAIL_COLUMNS

def field_specs(pdf, data):
    header = data['headers'][0]
    boxes = data.get('header_boxes', {}).get(header['文件名'], {})
    specs = [dict(page=1,item=None,position=None,field=field,value=header[field],bbox=boxes.get(field))
             for field in HEADER_COLUMNS[1:]]
    for row, source in zip(data['details'],data['evidence'],strict=True):
        for field in DETAIL_COLUMNS[1:]:
            specs.append(dict(page=source['页码'],item=row['项号'],position=source['行位置'],
                              field=field,value=row[field],bbox=source['字段位置'].get(field)))
    return specs

def normalize(field,value):
    if value is None: return None
    text=unicodedata.normalize('NFKC',str(value))
    text=re.sub(r'\s+','',text)
    if field=='项号': return text if re.fullmatch(r'[0-9]{1,4}',text) else None
    if field in ('总价','件数','毛重(千克)','净重(千克)'):
        try: return Decimal(text.replace(',',''))
        except InvalidOperation: return None
    if field in ('原产国','最终目的国','境内货源地','贸易国','运抵国','离境口岸'):
        text=plain(text)
    if field.endswith('18位代码'):
        return text if re.fullmatch(r'[A-Z0-9]{18}',text) else None
    if field=='商品编号': return text if re.fullmatch(r'\d{10}',text) else None
    return text or None

def compare(field,extracted,ocr,score):
    a,b=normalize(field,extracted),normalize(field,ocr)
    if a is None: return '待核对','提取值缺失或格式不符'
    if b is None: return '待核对','OCR 未取得有效值'
    if a!=b: return '待核对','文字层与页面 OCR 不一致'
    threshold=0.98 if field in ('商品编号','数量','总价','项号','件数','毛重(千克)','净重(千克)','合同协议号','提运单号') or field.endswith('18位代码') else 0.95
    if score is None or score<threshold:
        return '待核对',f'读数相同，但 OCR 分数低于暂定阈值 {threshold}'
    return '两路一致','仍需抽查；一致不等于人工确认'

def consistency_checks(pdf,data):
    rules=[]
    def add(name,observed,expected,ok,remark='',page=None,item=None):
        rules.append(dict(检查项=name,页码=page,项号=item,实际结果=str(observed),参考值=str(expected),状态='相符' if ok else '待核对',说明=remark))
    pages=data['page_info']
    page_marks=[p['印刷页码'] for p in pages]
    total=page_marks[0][1] if page_marks and page_marks[0] else 0
    observed=[p['页码'] for p in pages]
    complete = bool(total and observed==list(range(1,total+1)) and
                    all(not mark or tuple(mark)==(p['页码'],total) for p,mark in zip(pages,page_marks)))
    add('报关单页码完整性',page_marks,f'1 至 {total} 页',complete,
        '结合印刷页码和已定位明细页；原件缺少页码标识的页面另行提示')
    for p,mark in zip(pages,page_marks):
        if not mark:
            add('本页未标示印刷页码',p['页码'],'请核对原件顺序',False,page=p['页码'])
    items=[r['项号'] for r in data['details']]
    add('项号连续且不重复',items,f'1 至 {len(items)}',bool(items) and items==list(range(1,len(items)+1)))
    codes=0
    for entry in pages:
        left,right=entry['编码范围']
        codes+=sum(bool(re.fullmatch(r'\d{10}',w['text'])) and left<=center(w)<right
                   and entry['表起点']<w['top']<entry['表终点']
                   for w in normalized_words(pdf.pages[entry['页码']-1]))
    add('商品编码数量与提取行数',len(items),codes,len(items)==codes and bool(items),
        '从已定位的页面另数十位编码；未确认的页单列，不计为完成')
    header=data['headers'][0]
    gross,net=header['毛重(千克)'],header['净重(千克)']
    add('毛重不小于净重',gross,net,gross is not None and net is not None and gross>=net)
    weights=[]
    for row,e in zip(data['details'],data['evidence'],strict=True):
        raw=e['原文']
        quantity=re.fullmatch(r'(\d+(?:\.\d+)?)(.+)',row['数量'] or '')
        price=(raw.get('总价') or [''])[0].replace(' ','')
        try:
            q=Decimal(quantity[1]) if quantity else None
            p=Decimal(price)
            total=Decimal(str(row['总价']))
            calc=q*p if q is not None else None
            digits=len(price.split('.')[1]) if '.' in price else 0
            tolerance=max(Decimal('0.01'),abs(q)*Decimal(5).scaleb(-digits-1)+Decimal('0.005')) if q is not None else None
            ok=calc is not None and abs(calc-total)<=tolerance
            add('单价乘成交数量与总价',calc,total,ok,f'允许印刷单价舍入差；容差 {tolerance}。须确认单价对应第三行数量单位。',e['页码'],e['项号'])
        except (InvalidOperation,TypeError):
            add('单价乘成交数量与总价','无法计算','完整的单价、数量、总价',False,'缺失或非数字，不能当成0',e['页码'],e['项号'])
        # Legal quantities vary by commodity: kilograms may be the first or
        # second visual line. Only use a unique explicit kilogram value.
        weight_matches=[m for text in raw.get('数量',[]) if (m:=re.fullmatch(r'(\d+(?:\.\d+)?)千克',text.replace(' ','')))]
        if len(weight_matches)==1: weights.append(Decimal(weight_matches[0][1]))
    complete=len(weights)==len(items) and bool(items) and net is not None
    add('明细千克数合计与净重',sum(weights) if complete else '无法计算',net,
        complete and abs(sum(weights)-Decimal(str(net)))<=Decimal('0.01'),
        '仅在每项均有唯一千克值时计算；千克可出现在第一行或第二行；不把成交数量之和与包装件数比较')
    for issue in data['issues']:
        add('提取异常：'+issue['字段'],issue['原因'],'字段完整',False,page=issue['页码'],item=issue['项号'])
    return rules

def make_engine(model):
    from rapidocr import RapidOCR,OCRVersion,ModelType,EngineType
    if model!='v6_medium':raise ValueError('本工具仅附带 PP-OCRv6 medium 模型')
    version,size=OCRVersion.PPOCRV6,ModelType.MEDIUM
    model_path=RESOURCE_ROOT/'ocr_models'/'PP-OCRv6_rec_medium.onnx'
    if not model_path.is_file():raise FileNotFoundError('缺少 OCR 模型文件，请重新解压完整的工具文件夹。')
    fonts=Path(os.environ.get('WINDIR',r'C:\Windows'))/'Fonts'
    font=next((fonts/f for f in ('msyh.ttc','simsun.ttc','arial.ttf') if (fonts/f).is_file()),None)
    if font is None:raise FileNotFoundError('未找到 Windows 字体，无法初始化识别。')
    return RapidOCR(params={
        'Global.use_det':False,'Global.use_cls':False,'Global.use_rec':True,
        'Global.text_score':0.0,'Global.model_root_dir':str(RESOURCE_ROOT/'ocr_models'),
        'Global.font_path':str(font),'Rec.model_path':str(model_path),
        'Rec.ocr_version':version,'Rec.model_type':size,'Rec.engine_type':EngineType.ONNXRUNTIME,
        'EngineConfig.onnxruntime.intra_op_num_threads':int(os.environ.get('CUSTOMS_OCR_THREADS','4')),'EngineConfig.onnxruntime.inter_op_num_threads':1})

def audit(source,model,output=None,engine=None,relative_name=None):
    data=extract_batch(source)
    if data['file_count']!=1: raise ValueError('This audit CLI currently accepts one PDF; batch extractor is separate.')
    engine=engine or make_engine(model)
    records=[]
    started=time.perf_counter()
    with pdfplumber.open(source) as pdf:
        specs=field_specs(pdf,data)
        invalid_items={(i['页码'],i['项号']) for i in data['issues'] if i['字段']=='项号列异常'}
        pages={pn:page.to_image(resolution=300).original.convert('RGB') for pn,page in enumerate(pdf.pages,1)}
        for n,spec in enumerate(specs,1):
            bbox=spec['bbox']; raw=''; score=None; image64=''; error=''
            if bbox:
                im=pages[spec['page']]
                scale=im.width/842
                crop=im.crop(tuple(round(v*scale) for v in bbox))
                crop=ImageOps.expand(crop,border=8,fill='white')
                buf=io.BytesIO(); crop.save(buf,format='PNG')
                image64=base64.b64encode(buf.getvalue()).decode('ascii')
                try:
                    result=engine(crop,use_det=False,use_cls=False,use_rec=True)
                    raw=' '.join(result.txts or [])
                    score=float(min(result.scores)) if result.scores else None
                except Exception as exc:
                    error=f'{type(exc).__name__}: {exc}'
            status,reason=compare(spec['field'],spec['value'],raw,score)
            if spec['field']=='项号' and (spec['page'],spec['item']) in invalid_items:
                status,reason='待核对','项号列异常，请对照该物理页的原始行'
            if error: reason='OCR 执行失败：'+error
            records.append(dict(核对编号=n,文件名=source.name,页码=spec['page'],项号=spec['item'],行位置=spec['position'],字段=spec['field'],提取值=spec['value'],OCR读数=raw,OCR分数=score,核对状态=status,原因=reason,人工确认值=None,复核意见=None,截图=image64))
            if n%20==0 or n==len(specs): print(f'{model}: {n}/{len(specs)}',flush=True)
        rules=consistency_checks(pdf,data)
    filename=relative_name or source.name
    for collection in (data['headers'],data['details'],data['issues'],data['evidence'],records,rules):
        for row in collection:row['文件名']=filename
    data.update(audit=records,rules=rules,audit_model=model,audit_seconds=round(time.perf_counter()-started,2))
    if output:output.write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(dict(model=model,fields=len(records),agree=sum(r['核对状态']=='两路一致' for r in records),review=sum(r['核对状态']=='待核对' for r in records),rule_review=sum(r['状态']=='待核对' for r in rules),seconds=data['audit_seconds']),ensure_ascii=False))
    return data

def audit_batch(source,model,output):
    if source.is_file():return audit(source,model,output)
    files=sorted((p for p in source.rglob('*') if p.suffix.lower()=='.pdf'),key=lambda p:str(p).lower())
    if not files:raise ValueError('输入目录未找到 PDF')
    from extract_customs import HEADER_COLUMNS,DETAIL_COLUMNS
    combined=dict(header_columns=HEADER_COLUMNS,detail_columns=DETAIL_COLUMNS,headers=[],details=[],issues=[],evidence=[],audit=[],rules=[],file_count=len(files),page_count=0,audit_model=model,audit_seconds=0)
    engine=make_engine(model)
    for source_file in files:
        relative=str(source_file.relative_to(source))
        try:
            data=audit(source_file,model,engine=engine,relative_name=relative)
            combined['page_count']+=data['page_count'];combined['audit_seconds']+=data['audit_seconds']
            for k in ('headers','details','issues','evidence','audit','rules'):combined[k].extend(data[k])
        except Exception as exc:
            reason=f'{type(exc).__name__}: {exc}'
            combined['headers'].append({k:relative if k=='文件名' else None for k in HEADER_COLUMNS})
            combined['issues'].append(dict(文件名=relative,页码=None,项号=None,字段='整份文件',原因=reason))
            combined['rules'].append(dict(文件名=relative,检查项='整份文件处理',页码=None,项号=None,实际结果=reason,参考值='提取与图像核对完成',状态='待核对',说明='处理失败，未标为一致'))
    for i,row in enumerate(combined['audit'],1):row['核对编号']=i
    output.write_text(json.dumps(combined,ensure_ascii=False,indent=2),encoding='utf-8')
    return combined

def report(data,path):
    e=lambda v:html.escape('' if v is None else str(v))
    review=[r for r in data['audit'] if r['核对状态']!='两路一致']
    parts=['<!doctype html><meta charset="utf-8"><title>报关单逐字段核对</title><style>body{font:16px Microsoft YaHei,Arial;margin:32px;background:#f3f6fa;color:#172b4d}h1{font-size:26px}article{background:white;padding:18px;margin:16px 0;border-left:5px solid #9aa8b5;border-radius:5px}article.review{border-color:#dc762a}img{max-width:100%;border:1px solid #ddd;display:block;margin-top:12px}table{border-collapse:collapse;background:white}td,th{padding:10px;text-align:left;border-bottom:1px solid #ddd}p{line-height:1.7}.meta{color:#546477}a{color:#175b9c}</style>',
      '<h1>报关单逐字段核对</h1>',
      f'<p>{len(data["headers"])} 份文件，{data["page_count"]} 页，{len(data["details"])} 项商品；核对 {len(data["audit"])} 个内容字段，其中 {len(review)} 个待核对。文件名来自文件系统，不做 OCR。</p>',
      f'<p>OCR：{e(data["audit_model"])}，通过 RapidOCR / ONNX Runtime 在本机运行。OCR 分数不是正确率。编码、数量、金额等暂用 0.98 阈值，其他文字暂用 0.95；这些阈值尚未由真实批次校准。</p>',
      '<p>文字层用于定位截图，图像识别独立读取内容；两者仍可能共同定位错行。一致字段也需要抽查。人工确认值及复核意见请填写在配套 Excel 的“字段核对”表，原始提取值不自动覆盖。</p>',
      '<p>本页内嵌截图，可离线打开。以下优先展示待核对字段，再展示全部其余字段。</p>']
    parts.append('<h2>规则核对</h2><table><tr><th>检查项</th><th>项号</th><th>实际</th><th>参考</th><th>状态</th></tr>')
    for r in data['rules']:
        parts.append('<tr>'+''.join(f'<td>{e(r[k])}</td>' for k in ['检查项','项号','实际结果','参考值','状态'])+'</tr>')
    parts.append('</table><h2>字段与原图</h2>')
    for r in sorted(data['audit'],key=lambda r:(r['核对状态']=='两路一致',r['核对编号'])):
        parts.append(f'<article id="field-{r["核对编号"]}" class="{"review" if r["核对状态"]!="两路一致" else ""}"><h3>#{r["核对编号"]} · {e(r["字段"])} · {e(r["核对状态"])}</h3><p class="meta">{e(r["文件名"])} · 第 {r["页码"]} 页 · {"表头" if r["项号"] is None else "项号 "+str(r["项号"])}</p><p>文字层：<strong>{e(r["提取值"])}</strong><br>图像 OCR：<strong>{e(r["OCR读数"])}</strong><br>OCR 分数：{e(round(r["OCR分数"],4) if r["OCR分数"] is not None else None)}<br>{e(r["原因"])}</p>')
        if r['截图']: parts.append(f'<img alt="该字段的 PDF 原图截图" src="data:image/png;base64,{r["截图"]}">')
        parts.append('</article>')
    path.write_text('\n'.join(parts),encoding='utf-8')

if __name__=='__main__':
    p=argparse.ArgumentParser(); p.add_argument('source',type=Path); p.add_argument('output',type=Path); p.add_argument('--model',choices=['v6_medium'],default='v6_medium'); p.add_argument('--report',type=Path)
    a=p.parse_args(); data=audit_batch(a.source,a.model,a.output)
    if a.report: report(data,a.report)
