@echo off
setlocal
pushd "%~dp0" || exit /b 1
py -3.12-64 -m venv .venv
if errorlevel 1 (
  echo Install 64-bit Python 3.12 with the launcher and Tk from https://www.python.org/downloads/windows/
  goto :failed
)
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto :failed
if exist "ocr_models\PP-OCRv6_rec_medium.onnx" goto :ready
if not exist ocr_models mkdir ocr_models
if errorlevel 1 goto :failed
set "model=ocr_models\PP-OCRv6_rec_medium.onnx"
set "part=%model%.%RANDOM%-%RANDOM%.part"
curl.exe --fail --location --retry 2 --output "%part%" "https://www.modelscope.cn/models/RapidAI/RapidOCR/resolve/v3.10.0/onnx/PP-OCRv6/rec/PP-OCRv6_rec_medium.onnx"
if errorlevel 1 (
  if exist "%part%" del /q "%part%"
  goto :failed
)
move /y "%part%" "%model%" >nul
if errorlevel 1 goto :failed
:ready
".venv\Scripts\python.exe" -c "from scan_ocr import create_ocr_engine; create_ocr_engine()"
if errorlevel 1 (
  echo OCR could not start. Check the model and the Microsoft Visual C++ x64 runtime.
  echo https://learn.microsoft.com/cpp/windows/latest-supported-vc-redist
  goto :failed
)
echo Setup complete. Open Start-Windows.cmd to run.
popd
exit /b 0
:failed
echo Setup did not finish. See README.md and retry Setup-Windows.cmd.
if "%~1"=="" pause
popd
exit /b 1
