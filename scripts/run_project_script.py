# -*- coding: utf-8 -*-
"""在 Windows 上跑项目脚本的壳子（模板）。

为什么需要它：本机实测把中文路径当命令行参数传给 python 会偶发编码错乱——
`C:\\Agent\\用研工作\\...` 被按 GBK 解成乱码，脚本抛 FileNotFoundError，报错信息里的
路径也是乱的，很容易误判成文件不存在。同样的调用换个时机又能跑通。

用法：把 PROJ 改成项目根目录（**写在源码里**，别走命令行参数），存成 ASCII 路径的
文件（例如 `C:\\Users\\<你>\\.workbuddy\\tmp\\run.py`），然后：

    python run.py polish_md.py

argv 里只出现 ASCII 的脚本名，中文路径由源码里的字面量拼出来（源码本身按 UTF-8 存）。
stdout / stderr 一并收进 LOG，因为 PowerShell 直接捕获 python 的输出经常是空的。
"""
import sys, os, io, runpy, contextlib, traceback

# ↓↓↓ 这两行按项目改 ↓↓↓
PROJ = r"C:\path\to\project"                      # 项目根目录
SUB = os.path.join(".workbuddy", "build")         # 脚本所在子目录
# ↑↑↑ 这两行按项目改 ↑↑↑

LOG = os.path.join(os.path.dirname(os.path.abspath(__file__)), "run_log.txt")
name = sys.argv[1] if len(sys.argv) > 1 else ""
if not name:
    sys.exit("用法：python run.py <脚本名，ASCII>")
target = os.path.join(PROJ, SUB, name)

buf = io.StringIO()
err = ""
try:
    with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(buf):
        runpy.run_path(target, run_name="__main__")
except Exception:
    err = traceback.format_exc()

body = "target=%s\nexit=%s\n---- output ----\n%s" % (target, "FAIL" if err else "OK", buf.getvalue())
if err:
    body += "\n---- traceback ----\n" + err
with open(LOG, "w", encoding="utf-8") as fh:
    fh.write(body)

print("日志写到:", LOG)
