"""图形入口及离线批处理命令行。"""
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
    elif len(sys.argv) > 1 and sys.argv[1] in ('--inventory', '--tables', '--fields'):
        from general_backend import run_mode
        command = sys.argv[1]
        if len(sys.argv) != (5 if command == '--fields' else 4):
            raise SystemExit('用法：main.py --inventory|--tables 输入文件夹 输出文件夹；'
                             '--fields 输入文件夹 输出文件夹 规则.json')
        mode = {'--inventory': 'PDF 体检及文本导出', '--tables': 'PDF 表格提取',
                '--fields': '自定义字段提取'}[command]
        result = run_mode(mode, sys.argv[2], sys.argv[3], sys.argv[4] if command == '--fields' else '',
                          progress=lambda message: print(message, flush=True))
        for key, value in result.items():
            print(f'{key}：{value}', flush=True)
    else:
        from gui import Application
        Application().mainloop()
