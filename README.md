# 报关单 PDF 批量提取与核对

本项目从适配版式的中国海关出口报关单 PDF 提取表头和商品明细，生成两份 Excel、核对工作簿和离线原图页面。处理在本机完成；普通使用者可以运行打包后的 Windows 程序，不需要 Codex。

## 已有功能

- 批量读取文件夹及子文件夹中的 PDF，支持多页；每次运行创建新的结果文件夹。
- 提取文件名、商品编号、商品名称、数量、总价、原产国、最终目的国、境内货源地，以及发收货人、合同号、重量等表头字段。
- PDF 文字与字段截图 OCR 读数对照，并列出未读出、低分、读数不同和规则异常；可在离线页面查看字段原图。
- 核对页面支持按文件名、内容、字段和状态筛选；处理完成后可从主界面直接打开核对报告或原图页面。
- 区分子文件夹中的同名 PDF；数字代码按文本保存。处理异常的文件会在报告中提示。

## 使用边界

本项目针对已见过的报关单样式，依赖 PDF 文字层及字段位置；不保证适配所有货代的版式、纯扫描件、手写修改或图片遮挡。截图 OCR 也使用文字位置，取错位置或遗漏整行时，即使两路读数一致也可能出错。请检查文件数、页数、商品项数，逐项复核待核对字段，并对照原 PDF 检查关键数据。OCR 分数不是准确率。

核对报告的黄色列可记录人工确认值和意见，**不会自动更改业务 Excel**；确认错误后需要另行修改业务表。

## Windows 从源码运行

需要 64 位 Windows、Python 3.12（包含 Tkinter）和联网安装依赖。以下命令在项目根目录的 PowerShell 运行：

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
New-Item -ItemType Directory -Force -Path ocr_models | Out-Null
curl.exe -L --fail --output ocr_models/PP-OCRv6_rec_medium.onnx "https://www.modelscope.cn/models/RapidAI/RapidOCR/resolve/v3.10.0/onnx/PP-OCRv6/rec/PP-OCRv6_rec_medium.onnx"
.\.venv\Scripts\python.exe main.py
```

程序不包含模型，上述地址列在 [RapidOCR 模型清单](https://github.com/RapidAI/RapidOCR/blob/main/python/rapidocr/default_models.yaml) 的 ONNX 识别模型项；运行时使用项目内 `ocr_models/PP-OCRv6_rec_medium.onnx`。安装依赖和下载模型需要网络；准备好后提取和核对在本机运行。缺失模型时 OCR 无法完成，结果报告会标记待核对。

打开后选择 PDF 文件夹及已有的结果保存文件夹，点击“开始提取”。完成后先查看“提取核对报告.xlsx”，再通过核对编号在“逐字段原图核对.html”查看原图。

也可在已有的两个文件夹上使用命令行，结果文件夹会在输出目录下新建：

```powershell
.\.venv\Scripts\python.exe main.py --cli "PDF文件夹路径" "结果保存文件夹路径" 2
```

本机需要兼容的 Microsoft Visual C++ 运行库；如 ONNX Runtime 导入时遇到 DLL 错误，请按 [Microsoft 文档](https://learn.microsoft.com/cpp/windows/latest-supported-vc-redist)安装相应的受支持运行库。

可用虚构数据运行基础回归测试，无需模型或业务文件：

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

## 构建 Windows 程序

已安装 `requirements.txt`、下载模型后，另外准备合法取得的兼容 Microsoft VC++ 运行库文件到 `runtime/`，至少包含 `msvcp140.dll`，然后运行：

```powershell
powershell -ExecutionPolicy Bypass -File .\build.ps1 -Python .\.venv\Scripts\python.exe
```

也可使用 `-MsvcDir` 指向本机准备好的运行库文件夹。构建结果位于项目上一级 `portable_build/时间-编号/dist/报关单PDF批量提取/`；转交时必须保留整个文件夹，包括 `_internal`。构建并不代表已验证其他电脑的兼容性，请用实际样本和目标 Windows 环境测试。

## 许可与数据

本项目源码使用 [MIT 许可证](LICENSE)。第三方 Python 包、OCR 模型和 Microsoft 运行库各自适用其独立条款，不因本项目使用 MIT 而改变。本仓库不分发模型或 Microsoft 运行库；如需重新分发它们或程序包，请先核实相关上游条款与通知要求。[PaddlePaddle 官方模型卡](https://huggingface.co/PaddlePaddle/PP-OCRv6_medium_rec_onnx)标注 Apache-2.0，但 RapidAI 发布的转换文件与该模型卡所指文件来自不同发布方，不能仅据此推断转换件的再分发授权。

真实报关单、企业名称、合同号、Excel、HTML 原图页面、运行日志和截图都可能包含敏感信息。提交 issue、样本或代码前请使用虚构数据，并检查文档、嵌入图片和元数据；不要上传生产资料。详见 [CONTRIBUTING.md](CONTRIBUTING.md)。

## English

This Windows-oriented prototype extracts header and line-item fields from supported, text-based Chinese customs declaration PDFs. It exports two Excel workbooks, an audit workbook, and an offline field-image review page. OCR reads image crops located using the PDF text layer; matching readings do not prove correctness. Scanned documents, altered images, obscured values, and unfamiliar layouts require review against the original.

For source setup, use the PowerShell commands above. The RapidOCR ONNX model is downloaded separately and is not included in this repository. Third-party packages, model files, and Microsoft runtimes have separate licenses; verify applicable terms before redistribution. Never upload real customs documents, outputs, logs, or identifiable screenshots; use synthetic fixtures.
