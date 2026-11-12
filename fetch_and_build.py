#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""自动抓取 SPDR Gold Shares 官方历史数据，更新页面数据块、CSV 与给 AI 的文本摘要。

由 GitHub Actions 定时调用，也可本机直接运行：
    python fetch_and_build.py

产出三个同源同日文件：index.html（内嵌数据块，GitHub Pages 直接托管）、
gld_data.csv、llms.txt。只有数据真的变化才会产生 diff，保证无新数据时零提交。

数据来自发行人公开披露文件，本仓库只做衍生指标展示，不收录原始 XLSX；
原始文件含使用限制，如权利方有异议请联系仓库所有者移除。仅为图表分析，不构成投资建议。
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


PAGES = "https://wj402041069-del.github.io/gold-etf-divergence/"
REPO_URL = "https://github.com/wj402041069-del/gold-etf-divergence"


def weekly(rows):
    """与 index.html aggregate(rows,'weekly') 一致：周一为 key，同周取最后一个交易日。"""
    groups = {}
    for r in rows:
        key = r[0] - datetime.timedelta(days=r[0].weekday())
        groups[key] = r
    return [groups[k] for k in sorted(groups)]


def roll_z(vals, n):
    """与 index.html rollZ 一致：总体标准差（除以 n，不是 n-1）。"""
    out = [None] * len(vals)
    for i in range(n - 1, len(vals)):
        w = vals[i - n + 1:i + 1]
        m = sum(w) / n
        sd = (sum((x - m) ** 2 for x in w) / n) ** 0.5
        out[i] = (vals[i] - m) / sd if sd > 0 else 0.0
    return out


def write_llms(rows):
    """给 AI / 大模型读的纯文本摘要，与页面同源同日、同口径，自动生成，不要手改。"""
    wk = weekly(rows)
    win = 52
    zp = roll_z([r[1] for r in wk], win)
    zt = roll_z([r[2] for r in wk], win)
    D = [None if zp[i] is None or zt[i] is None else zt[i] - zp[i]
         for i in range(len(wk))]
    valid = [v for v in D if v is not None]
    cur = D[-1]
    pct = (sum(1 for v in valid if v <= cur) / len(valid) * 100) if cur is not None and valid else 0.0
    last = rows[-1]
    prices = [r[1] for r in rows]
    pmin, pmax = min(prices), max(prices)
    txt = """# 黄金 ETF 价格 x 持仓量背离（GLD）- 给 AI 的文本摘要

本文件由 fetch_and_build.py 自动生成，与 index.html / gld_data.csv 同源同日、同口径。
请勿手改；改文案请改 fetch_and_build.py 的 write_llms()。

## 数据快照
- 数据截止日: {end}
- 覆盖区间: {start} ~ {end}
- 有效交易日: {nd}
- 周频聚合: {nw} 期（滚动窗口 {win} 期）
- 最新 GLD 收盘: {price:.2f} USD
- 最新实物持仓: {tonnes:.2f} 吨

## 最新读数（周频 / 全历史口径）
- 背离强度 D = Z(持仓) - Z(价格) = {d:+.2f}
- 历史分位 = {pct:.1f}%（历史上约 {rest:.1f}% 的时间比现在更背离）
- 读法: D > 0 表示持仓相对价格偏强，常见于价格下跌时仍有逆势增持；
  D > 2 才算历史罕见，当前属于「偏强，值得留意」，不是买卖信号。

## 口径（改动前必读）
- 价格与实物持仓吨数来自同一份 SPDR Gold Shares 官方 Historical Archive，
  同一交易日、同一标的，同源同频，不存在跨源时点错配。
- 用滚动 Z-score（总体标准差，除以 n），不是全程标准化；缩放视图只改可视区间，不改统计口径。
- 持仓用实物「吨数」，不用 AUM：金价跌 10% 会让 AUM 自动缩水，即使一克都没被赎回。
- 散点用「持仓变化」而非「持仓水平」，避免共同趋势造成伪相关。
- 频率切换会改变读数（日 252 / 周 52 / 月 36），跨频率的数值不可直接互比。
- 交易日口径：只用 Historical Archive 的 Tonnes of Gold 列，不混用按结算日的金条清单。

## 参照图（位于图 1 与图 2 之间）
- 画的是 GLD 每份份额的**原始收盘价（USD/份）**，不是 Z-score、不是现货金价 USD/oz，不做对数或归一化。
  全历史价格区间 {pmin:.2f} ~ {pmax:.2f} USD；数据末期 {price:.2f} USD。
- 单轴，不叠加持仓吨数第二条线，避免双轴截断刻度的视觉误导；持仓真实吨数在悬停浮层里给出。
- 纵轴默认按当前视图 min/max 自适应并在图上标出范围，可一键切「从 0 起」看长期真实幅度。
- 与图 1 的 |D| > 1.5σ 浅色竖带共用 strongSegs()，位置完全一致，方便对照「背离高发时金价在什么位置」。
- 与图 1 / 图 2 共用 axisX() / bindChart() / showHover() 和同一个 S.view；三图 L=52 / R=18，
  hover 竖线像素级对齐，滚轮缩放 / 拖拽平移 / 双击重置 / 导航条框选全部同步。
- 图 1 里同色的橙色线是它的滚动 Z-score。判断背离用图 1 / 图 2，确认当时绝对价位用参照图，两者不可混读。

## 图 3 四象限散点怎么读（最容易看懵的一张）
- 一个点 = 某一期相对上一期的变化，不是价格水平，也不是持仓总量。
- 横轴 x = (本期价格 / 上期价格 - 1) * 100，单位 %；竖线左边 = 价跌，右边 = 价涨。
- 纵轴 y = 本期持仓吨数 - 上期持仓吨数，单位吨；横线上边 = 增持，下边 = 减持。
- 蓝色左上角 = 价跌 + 增持 = 越跌越买（本页重点）；灰点 = 追涨 / 跌了赎回 / 逢高兑现。
- 橙色虚线 = 最小二乘拟合趋势线，越往右上倾斜说明价涨时持仓也倾向增加，对应相关系数 r。
- 显示按 2-98 分位裁剪极端值，但统计仍用视图内全样本，因此图上可见点数可能略少于卡片里的期数。
- 卡片「价跌期 / 其中增持」就是这张图的统计版：左半边的点数，及其中同时落在上半边的点数。
- 紫圈 + 紫色细线 = 当前视图末尾 N 期（默认 10，可切 5 / 20 / 关闭）的时间轨迹，末端空心圈标「最新」。
  填充色仍表示象限（蓝 = 越跌越买），紫圈只表示「最近」，两个维度分开，避免颜色语义打架。
- 散点图本身丢掉了时间维度；沿这条短轨迹按时间看，才能判断最近是往左上角（越跌越买）还是别处移动。
  取的是当前视图（导航条蓝框右端）末尾的点，不是导航条上画的 D 序列。

## 已知限制
- GLD 是单只基金，不代表全球黄金 ETF 总持仓；与 WGC Goldhub 月度口径覆盖范围不同，不可互推。
- 参照图是 GLD 份额价格（USD/份），大致对应十分之一盎司且逐年扣管理费，不是现货金价 USD/oz，不能换算比价。
- 相关系数只说明同向程度，不说明因果；D 与「背离」都不是涨跌预测。
- 2004-2005 年建仓期吨数快速爬升，属结构性变化，不宜与成熟期直接比较。
- 原始数据版权归发行人所有；本仓库只做衍生分析展示，不收录原始 XLSX。

## 链接
- 在线看板: {pages}
- 本摘要: {pages}llms.txt
- 源码 / 手动刷新: {repo}（Actions -> Refresh GLD divergence data -> Run workflow）
- 逐日数据: {pages}gld_data.csv

仅为图表分析，不构成投资建议。
""".format(end=last[0].isoformat(), start=rows[0][0].isoformat(),
           nd=len(rows), nw=len(wk), win=win, price=last[1], tonnes=last[2],
           pmin=pmin, pmax=pmax, d=cur, pct=pct, rest=100 - pct,
           pages=PAGES, repo=REPO_URL)
    path = os.path.join(BASE, "llms.txt")
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write(txt)
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
    llms_path = write_llms(rows)
    print("有效交易日 %d，区间 %s ~ %s" % (len(rows), rows[0][0], rows[-1][0]))
    print("最新: 收盘 %.2f USD / 持仓 %.2f 吨" % (rows[-1][1], rows[-1][2]))
    print("数据截止 %s，index.html %d bytes" % (updated, size))
    print("CSV %s / 文本摘要 %s" % (csv_path, llms_path))


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("ERROR: %s" % e, file=sys.stderr)
        sys.exit(1)
