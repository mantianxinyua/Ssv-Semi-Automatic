# -*- coding: utf-8 -*-
"""把本机已安装的第三方依赖拷贝到发布用的 libs/ 目录（免重新下载）"""
import importlib
import os
import shutil
import glob
import site

DST = r"E:\源码\报告\源码\libs"
PKGS = ["pandas", "numpy", "dateutil", "pytz", "tzdata", "six",
        "openpyxl", "et_xmlfile", "PIL", "cv2"]

sp = site.getsitepackages()[0]
os.makedirs(DST, exist_ok=True)

for p in PKGS:
    try:
        mod = importlib.import_module(p)
        src_dir = os.path.dirname(mod.__file__)
        base = os.path.basename(src_dir)
        if os.path.isdir(src_dir):
            shutil.copytree(src_dir, os.path.join(DST, base), dirs_exist_ok=True)
            print(f"[OK] {base}")
        for di in glob.glob(os.path.join(sp, base + "-*.dist-info")):
            shutil.copytree(di, os.path.join(DST, os.path.basename(di)), dirs_exist_ok=True)
    except Exception as e:
        print(f"[ERR] {p}: {e}")

print("done")
