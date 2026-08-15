# -*- coding: utf-8 -*-
"""
程序入口（明文，仅几行，无业务逻辑，泄露无妨）。
sferzaibo.pyd 为 Nuitka 编译的机器码（混淆加密后的核心代码）。
第三方库从 exe 同级的 libs\\ 加载，不打包进 exe。
所有运行日志（含报错）自动写入 exe 同级的「运行日志.log」。
"""
import os
import sys
import logging
import traceback
import threading

APP_DIR = os.path.dirname(os.path.abspath(sys.executable))
LOG_FILE = os.path.join(APP_DIR, "运行日志.log")


# ================= 统一日志（时间戳 + 全部报错写入 运行日志.log） =================

def _setup_logging():
    fmt = logging.Formatter("%(asctime)s.%(msecs)03d %(message)s", datefmt="%H:%M:%S")
    fh = logging.FileHandler(LOG_FILE, mode="a", encoding="utf-8")
    fh.setFormatter(fmt)
    root = logging.getLogger()
    root.setLevel(logging.INFO)
    root.addHandler(fh)
    logging.raiseExceptions = False  # 防止日志自身异常导致递归

    # windowed 模式下 stdout/stderr 无控制台：重定向到日志，print/traceback 不丢失
    class _Tee:
        def __init__(self, tag):
            self.tag = tag

        def write(self, msg):
            m = msg.strip()
            if m:
                logging.info("[%s] %s", self.tag, m)

        def flush(self):
            pass

        def isatty(self):
            return False

    sys.stdout = _Tee("信息")
    sys.stderr = _Tee("错误")

    # 未捕获异常钩子：主线程 + 子线程的异常都写进日志
    def _hook(etype, value, tb):
        logging.error("未捕获异常:\n%s",
                      "".join(traceback.format_exception(etype, value, tb)))

    sys.excepthook = _hook
    try:
        threading.excepthook = lambda args: _hook(
            args.exc_type, args.exc_value, args.exc_traceback)
    except AttributeError:
        pass


_setup_logging()
logging.info("启动: %s", sys.version)


# ================= 外部依赖目录注入 =================
#   libs       -> 第三方库（pandas/numpy/cv2 等）
#   libs\Lib   -> 标准库兜底（platform 等）
#   libs\DLLs  -> 标准库 C 扩展兜底（_socket/_ssl 等）
#   app        -> （备用）外部机器码模块目录
for sub in ("libs", os.path.join("libs", "Lib"), os.path.join("libs", "DLLs"), "app"):
    p = os.path.join(APP_DIR, sub)
    if os.path.isdir(p) and p not in sys.path:
        sys.path.insert(0, p)

# ================= 导入核心模块 =================
try:
    logging.info("路径: %s", sys.executable)
    logging.info("导入模块...")
    import sferzaibo  # Nuitka 编译的机器码模块
    logging.info("模块导入完成")
    logging.info("获取机器码...")
    _mid = sferzaibo.get_machine_id()
    logging.info("机器码: %s", _mid)
except Exception:
    logging.error("启动失败:\n%s", traceback.format_exc())
    try:
        with open(os.path.join(APP_DIR, "error.log"), "w", encoding="utf-8") as f:
            traceback.print_exc(file=f)
    except Exception:
        pass
    raise

# 诊断：设置 DSH_PRELOAD_CV2=1 时，启动即预加载 cv2 并写日志（验证 cv2 DLL，默认关闭）
if os.environ.get("DSH_PRELOAD_CV2"):
    try:
        sferzaibo._import_cv2()
        logging.info("cv2 预加载完成")
    except Exception:
        logging.error("cv2 预加载失败:\n%s", traceback.format_exc())

if __name__ == "__main__":
    # 高 DPI 适配必须在任何 Tk 窗口创建之前调用
    try:
        sferzaibo._enable_high_dpi()
    except Exception:
        pass

    logging.info("验证授权...")
    try:
        sferzaibo.check_license()
        logging.info("授权通过")
    except SystemExit:
        logging.error("授权验证未通过（已弹窗提示），程序退出")
        raise

    logging.info("启动主界面...")
    try:
        app = sferzaibo.App()
        logging.info("主界面已启动")
        app.mainloop()
    except Exception:
        logging.error("主界面异常:\n%s", traceback.format_exc())
        raise
    logging.info("程序退出")
