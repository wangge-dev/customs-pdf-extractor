"""图形入口；--cli 用于本机或封装后的端到端验证。"""
import multiprocessing
import os
import sys

if __name__ == '__main__':
    # Windowed executables do not have an attached console.
    if sys.stdout is None:
        sys.stdout = open(os.devnull, 'w', encoding='utf-8')
    if sys.stderr is None:
        sys.stderr = open(os.devnull, 'w', encoding='utf-8')
    multiprocessing.freeze_support()
    if len(sys.argv) > 1 and sys.argv[1] == '--cli':
        from backend import run_batch
        if len(sys.argv) not in (4, 5):
            raise SystemExit('用法：程序.exe --cli PDF文件夹 结果文件夹 [工作进程数]')
        run_batch(sys.argv[2], sys.argv[3],
                  progress_callback=lambda message: print(message, flush=True),
                  workers=int(sys.argv[4]) if len(sys.argv) == 5 else None)
    else:
        from gui import Application
        Application().mainloop()
