"""Local experiment: CC BY 4.0 public synthetic invoices and an original Chinese scan.

Source: https://huggingface.co/datasets/alamgirqazi/invoice-ocr-synthetic
Authors: Qazi et al., InvoiceOCR-Synth. Do not interpret as a production benchmark.
"""
import json
import time
from pathlib import Path
from urllib.request import Request, urlopen

from PIL import Image, ImageDraw, ImageFont, ImageFilter


BASE = Path(__file__).parent / 'ocr_benchmark_public'
BASE.mkdir(exist_ok=True)
DOCS = BASE / 'pdf'
DOCS.mkdir(exist_ok=True)


def download_public():
    from urllib.parse import urlencode
    rows = []
    for number in (0, 1, 2, 1876, 1877, 1878):
        q = urlencode({'dataset':'alamgirqazi/invoice-ocr-synthetic','config':'default',
                       'split':'train','offset':number,'length':1})
        with urlopen(Request('https://datasets-server.huggingface.co/rows?' + q,
                             headers={'User-Agent':'OCR-experiment'}), timeout=45) as f:
            record = json.load(f)['rows'][0]['row']
        image_path = BASE / f'public_{number}_{record["quality"]}.jpg'
        with urlopen(Request(record['image']['src'],headers={'User-Agent':'OCR-experiment'}), timeout=45) as f:
            image_path.write_bytes(f.read())
        with Image.open(image_path) as image:
            image.convert('RGB').save(DOCS / f'public_{number}_{record["quality"]}.pdf', 'PDF', resolution=150)
        record.pop('image')
        rows.append({'row':number, 'image':str(image_path), 'truth':record})
    (BASE / 'public_truth.json').write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding='utf-8')
    return rows


def make_chinese():
    image=Image.new('RGB',(1654,2339),'white')
    draw=ImageDraw.Draw(image)
    font=ImageFont.truetype('C:/Windows/Fonts/msyh.ttc',36)
    title=ImageFont.truetype('C:/Windows/Fonts/msyh.ttc',48)
    draw.text((100,90),'虚构示例采购单',font=title,fill='black')
    lines=['单据编号：TEST-2026-1130','发货单位：示例贸易有限公司','收货单位：测试商业有限公司',
           '制单日期：2026-11-24','币种：人民币','商品名称              数量             单价             金额',
           '棉质布料              12 米           15.50            186.00',
           '包装纸箱              25 个            2.20             55.00',
           '订单合计：241.00 元']
    for i,line in enumerate(lines):
        y=190+i*80
        draw.text((100,y),line,font=font,fill='black')
        if i in (4,5,6,7):
            draw.line((90,y+61,1510,y+61),fill='black',width=2)
    expected=['TEST-2026-1130','示例贸易有限公司','测试商业有限公司','2026-11-24','人民币',
              '棉质布料','12','15.50','186.00','包装纸箱','25','2.20','55.00','241.00']
    image=image.filter(ImageFilter.GaussianBlur(0.3))
    image.save(BASE/'chinese_original.png')
    image.save(DOCS/'chinese_original.pdf','PDF',resolution=150)
    (BASE/'chinese_truth.json').write_text(json.dumps(expected,ensure_ascii=False,indent=2),encoding='utf-8')
    return expected


if __name__ == '__main__':
    started=time.monotonic()
    rows=download_public()
    expected=make_chinese()
    print(f'Prepared {len(rows)} public scans and 1 original Chinese scan in {time.monotonic()-started:.1f}s')
