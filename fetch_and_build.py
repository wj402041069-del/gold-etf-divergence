#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""自动抓取 SPDR Gold Shares 官方历史数据，更新 index.html 内的数据块与 gld_data.csv。

由 GitHub Actions 定时调用，也可本机直接运行：
    python fetch_and_build.py

所有数据来自发行人公开披露文件。该文件声明禁止再分发，请务必把仓库设为 private，
不要将含数据的 index.html 发布到公开站点。数据仅用于个人分析。
"""
import csv
import datetime
import io
import json
import os
import re
import sys
import time
import urllib.request

ARCHIVE_URL = (
    "https://api.spdrgoldshares.com/api/v1/historical-archive"
    "?product=gld&exchange=NYSE&lang=en"
)
SHEET = "US GLD Historical Archive"
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120 Safari/537.36"
START = "/* __GLD_DATA_START__ */"
END = "/* __GLD_DATA_END__ */"
BASE = os.path.dirname(os.path.abspath(__file__))


def download(retries=3):
    last = None
    for i in range(retries):
        try:
            req = urllib.request.Request(ARCHIVE_URL, headers={"User-Agent": UA, "Accept": "*/*"})
            with urllib.request.urlopen(req, timeout=90) as resp:
                data = resp.read()
            if data[:2] != b"PK":
                raise ValueError("返回内容不是 xlsx")
            return data
        except Exception as e:
            last = e
            print("下载第 %d 次失败: %s" % (i + 1, e), flush=True)
            time.sleep(3 * (i + 1))
    raise RuntimeError("下载失败: %s" % last)


def parse(raw):
    import openpyxl

    wb = openpyxl.load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
    ws = wb[SHEET]
    header = [str(c.value).strip() if c.value is not None else "" for c in next(ws.iter_rows(min_row=1, max_row=1))]

    def find(*keys):
        for i, h in enumerate(header):
            if all(k in h.lower() for k in keys):
                return i
        return None

    di, pi, ti = find("date"), find("closing", "price"), find("tonnes", "gold")
    if di is None or pi is None or ti is None:
        raise RuntimeError("表头未识别: %s" % header)

    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        try:
            d = datetime.datetime.strptime(str(r[di]).strip(), "%d-%b-%Y").date()
            p, t = float(r[pi]), float(r[ti])
        except (ValueError, TypeError, IndexError):
            continue  # "US Holiday" 等非交易行
        if p > 0 and t > 0:
            rows.append((d, round(p, 2), round(t, 2)))
    rows.sort()
    if len(rows) < 100:
        raise RuntimeError("有效数据过少: %d" % len(rows))
    return rows


def compact(rows):
    return [[d.isoformat(), p, t] for d, p, t in rows]


def update_html(data, updated):
    path = os.path.join(BASE, "index.html")
    html = open(path, encoding="utf-8").read()
    if START not in html or END not in html:
        raise RuntimeError("index.html 缺少数据块标记，请先放入模板")
    block = (
        START + "\n"
        + "window.GLD_DATA = " + json.dumps(data, separators=(",", ":")) + ";\n"
        + "window.GLD_UPDATED = " + json.dumps(updated) + ";\n"
        + END
    )
    new, n = re.subn(re.escape(START) + r"[\s\S]*?" + re.escape(END), lambda m: block, html, count=1)
    if n != 1:
        raise RuntimeError("数据块替换失败")
    # 固定用 LF 写回，避免 Windows(CRLF) 与 Actions/Linux(LF) 之间产生整文件 diff
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write(new)
    return len(new.encode("utf-8"))


def write_csv(rows):
    path = os.path.join(BASE, "gld_data.csv")
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["date", "price", "tonnes"])
        for d, p, t in rows:
            w.writerow([d.isoformat(), p, t])
    return path


def main():
    print("抓取 SPDR 官方历史数据 ...", flush=True)
    rows = parse(download())
    data = compact(rows)
    # 只记录「数据截止日」，不要记录脚本运行时间：
    # 否则官方没有新数据时也会产生 diff，导致每天空提交、仓库持续膨胀。
    updated = rows[-1][0].isoformat()
    size = update_html(data, updated)
    csv_path = write_csv(rows)
    print("有效交易日 %d，区间 %s ~ %s" % (len(rows), rows[0][0], rows[-1][0]))
    print("最新: 收盘 %.2f USD / 持仓 %.2f 吨" % (rows[-1][1], rows[-1][2]))
    print("数据截止 %s，index.html %d bytes，CSV %s" % (updated, size, csv_path))


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("ERROR: %s" % e, file=sys.stderr)
        sys.exit(1)
