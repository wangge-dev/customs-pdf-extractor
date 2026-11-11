# 本地 PDF 批处理与核对工具

面向日常文档整理的离线工具：批量检查 PDF、按页导出文本、提取表格和自定义字段，并保留来源页码与异常信息。另提供针对已适配版式的中国海关出口报关单提取与原图核对。准备好依赖后，处理在本机完成，不需要 Codex。

## 功能

### 通用 PDF

- **文件体检与逐页文本**：递归列出文件、页数、页面尺寸、可提取文字量，以及无文字层、密码保护或读取失败的情况；逐页导出文本并保留目录结构。
- **表格提取**：逐页识别表格，导出文件相对路径、页码、表号、行列号及单元格文字，提供页面和文件状态供复查。
- **自定义字段**：用 JSON 定义字段和正则表达式，将匹配结果、页码、附近原文及异常导出 Excel；适合对文字层有稳定标识的不同单据试用。
- 同名文件以相对路径区分，结果放在新建文件夹，不覆盖之前的批次；文字作为 Excel 文本保存，避免把内容当作公式执行。

通用模式直接读取 PDF 文字层，**不对扫描件自动 OCR**。无文字层可能是扫描页、空白页或文字无法提取；体检提示不能判定具体原因。表格、字段规则可能漏项或误匹配，请按文件、页码和原文核对。

体检会对单张图片占页较大、可提取文字很少的页面提示“疑似图片主体”。这只是辅助筛查，图片分块、图层或少量文字等情况可能漏报或误报；有文字层也不代表正文可提取。

字段规则使用 Python 正则表达式，示例见 [示例字段模板.json](示例字段模板.json)：`name` 是输出字段名，`pattern` 是匹配规则，`group` 是取值的捕获组编号或名称，省略时取完整匹配。示例只作演示，需要按实际单据设计并对照原文；规则按页执行，不跨页匹配。建议只使用可信规则，复杂正则可能耗时很长。

### 报关单专用模式

- 批量读取文件夹及子文件夹中的 PDF，支持多页；每次运行创建新的结果文件夹。
- 提取文件名、商品编号、商品名称、数量、总价、原产国、最终目的国、境内货源地，以及发收货人、合同号、重量等表头字段。
- PDF 文字与字段截图 OCR 读数对照，并列出未读出、低分、读数不同和规则异常；可在离线页面查看字段原图。
- 核对页面支持按文件名、内容、字段和状态筛选；处理完成后可从主界面直接打开核对报告或原图页面。
- 区分子文件夹中的同名 PDF；数字代码按文本保存。处理异常的文件会在报告中提示。

## 报关单模式的使用边界

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

打开后选择 PDF 文件夹及已有的结果保存文件夹，选择处理模式并点击“开始提取”。字段提取模式还需选择 JSON 规则文件；可从仓库中的示例复制一份再修改。通用模式的输出位置不要选在 PDF 输入目录内部。

报关单模式完成后先查看“提取核对报告.xlsx”，再通过核对编号在“逐字段原图核对.html”查看原图。通用表格为方便追溯采用逐单元格长表，保留页码与表号，不保证还原复杂合并单元格或所有无框表格。

也可在已有的两个文件夹上使用命令行，结果文件夹会在输出目录下新建：

```powershell
.\.venv\Scripts\python.exe main.py --cli "PDF文件夹路径" "结果保存文件夹路径" 2
.\.venv\Scripts\python.exe main.py --inventory "PDF文件夹路径" "结果保存文件夹路径"
.\.venv\Scripts\python.exe main.py --tables "PDF文件夹路径" "结果保存文件夹路径"
.\.venv\Scripts\python.exe main.py --fields "PDF文件夹路径" "结果保存文件夹路径" "字段规则.json"
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

也可使用 `-MsvcDir` 指向本机准备好的运行库文件夹。构建结果位于项目上一级 `portable_build/时间-编号/dist/本地PDF批处理/`；转交时必须保留整个文件夹，包括 `_internal`。构建并不代表已验证其他电脑的兼容性，请用实际样本和目标 Windows 环境测试。

## 许可与数据

本项目源码使用 [MIT 许可证](LICENSE)。第三方 Python 包、OCR 模型和 Microsoft 运行库各自适用其独立条款，不因本项目使用 MIT 而改变。本仓库不分发模型或 Microsoft 运行库；如需重新分发它们或程序包，请先核实相关上游条款与通知要求。[PaddlePaddle 官方模型卡](https://huggingface.co/PaddlePaddle/PP-OCRv6_medium_rec_onnx)标注 Apache-2.0，但 RapidAI 发布的转换文件与该模型卡所指文件来自不同发布方，不能仅据此推断转换件的再分发授权。

真实报关单、企业名称、合同号、Excel、HTML 原图页面、运行日志和截图都可能包含敏感信息。提交 issue、样本或代码前请使用虚构数据，并检查文档、嵌入图片和元数据；不要上传生产资料。详见 [CONTRIBUTING.md](CONTRIBUTING.md)。

## English

This Windows-oriented, offline toolkit inventories PDFs, exports page text, extracts tables and configurable regex fields, and reports provenance and exceptions. These general modes require an extractable text layer and do not OCR image-only scans. A separate customs mode extracts fields from supported Chinese customs declarations and provides field-image review. OCR reads image crops located using the PDF text layer; matching readings do not prove correctness. Scans, obscured values, and unfamiliar layouts require review against the original.

For source setup, use the PowerShell commands above. The RapidOCR ONNX model is downloaded separately and is not included in this repository. Third-party packages, model files, and Microsoft runtimes have separate licenses; verify applicable terms before redistribution. Never upload real customs documents, outputs, logs, or identifiable screenshots; use synthetic fixtures.
