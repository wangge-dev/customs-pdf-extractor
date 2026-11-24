"""Compare same-page OCR using identical detector and two local recognition models."""
import json
import os
import re
import time
import ctypes
from pathlib import Path

if os.environ.get('OCR_MSVC_DLL'):
    ctypes.WinDLL(os.environ['OCR_MSVC_DLL'])

import rapidocr
from rapidocr import RapidOCR, OCRVersion, ModelType

ROOT = Path(__file__).parent
BASE = ROOT / 'ocr_benchmark_public'
DETECTION = Path(rapidocr.__file__).parent / 'models/PP-OCRv6_det_small.onnx'


def normal(value):
    return re.sub(r'[^0-9a-z\u4e00-\u9fff]', '', str(value).lower())


def init(name):
    if name == 'v6_medium':
        path, version, size = ROOT/'ocr_models/PP-OCRv6_rec_medium.onnx', OCRVersion.PPOCRV6, ModelType.MEDIUM
    else:
        path, version, size = ROOT/'ocr_models/ch_PP-OCRv5_rec_server.onnx', OCRVersion.PPOCRV5, ModelType.SERVER
    return RapidOCR(params={'Global.use_cls':False, 'Global.text_score':0.0, 'Global.max_side_len':1400,
                            'Global.font_path':'C:/Windows/Fonts/msyh.ttc',
                            'Det.model_path':str(DETECTION), 'Rec.model_path':str(path),
                            'Rec.ocr_version':version, 'Rec.model_type':size,
                            'EngineConfig.onnxruntime.intra_op_num_threads':4,
                            'EngineConfig.onnxruntime.inter_op_num_threads':1})


def assess(text, truths):
    output = normal(text)
    return [{'label':value,'found':normal(value) in output} for value in truths]


def main():
    records=json.loads((BASE/'public_truth.json').read_text(encoding='utf-8'))
    chinese=json.loads((BASE/'chinese_truth.json').read_text(encoding='utf-8'))
    sources=[]
    for row in records:
        if row['row'] not in (0, 1876):
            continue
        truth=row['truth']
        # Restrict to ground-truth content that can be visually checked later.
        targets=[truth.get('document_number'), truth.get('supplier_name')]
        targets += [item['description'] for item in truth.get('line_items',[])[:3]]
        sources.append({'sample':f"public_{row['row']}_{truth['quality']}",
                        'path':row['image'],'targets':[x for x in targets if x]})
    sources.insert(0,{'sample':'chinese_original','path':str(BASE/'chinese_original.png'),'targets':chinese})
    results=[]
    for model in ('v5_server','v6_medium'):
        engine=init(model)
        for src in sources:
            started=time.monotonic()
            output=engine(src['path'])
            elapsed=time.monotonic()-started
            pieces=list(output.txts) if output and output.txts else []
            hits=assess('\n'.join(pieces),src['targets'])
            results.append({'model':model,'sample':src['sample'],'seconds':round(elapsed,2),
                            'lines':len(pieces),'fields':hits,'recognized':pieces})
            (BASE/'ocr_compare.json').write_text(json.dumps(results,ensure_ascii=False,indent=2),encoding='utf-8')
            print(model,src['sample'],'hits',sum(x['found'] for x in hits),'/',len(hits),
                  'lines',len(pieces),'secs',round(elapsed,2),flush=True)
    (BASE/'ocr_compare.json').write_text(json.dumps(results,ensure_ascii=False,indent=2),encoding='utf-8')


if __name__ == '__main__':
    main()
