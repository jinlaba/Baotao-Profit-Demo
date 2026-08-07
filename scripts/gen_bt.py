# -*- coding: utf-8 -*-
"""读取 _data_bt.json + bt_tpl.html 模板，生成 宝淘淘日报月报看板.html
用法: python gen_bt.py  （与 extract_bt.py、bt_tpl.html 放在同一目录）"""
import json
import datetime
import os

BASE = os.path.dirname(os.path.abspath(__file__))
TPL = os.path.join(BASE, "bt_tpl.html")
JSON = os.path.join(BASE, "_data_bt.json")
OUT = os.path.join(BASE, "宝淘淘日报月报看板.html")

with open(TPL, encoding="utf-8") as f:
    html = f.read()

with open(JSON, encoding="utf-8") as f:
    data = json.load(f)

html = html.replace("%%DATA%%", json.dumps(data, ensure_ascii=False, separators=(",", ":")))
html = html.replace("%%RANGE_MIN%%", data["range"]["min"])
html = html.replace("%%RANGE_MAX%%", data["range"]["max"])
html = html.replace("%%GENTIME%%", datetime.datetime.now().strftime("%Y-%m-%d %H:%M"))

with open(OUT, "w", encoding="utf-8") as f:
    f.write(html)

print("生成完成:", OUT)
print("HTML KB:", round(os.path.getsize(OUT) / 1024, 1))
