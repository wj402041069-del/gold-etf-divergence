#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""自动抓取 SPDR GLD 与 iShares SLV 官方历史数据，更新页面数据块、CSV 与给 AI 的文本摘要。

由 GitHub Actions 定时调用，也可本机直接运行：
    python fetch_and_build.py

产出同源同日文件：index.html（内嵌 GLD / SLV 两个数据块，GitHub Pages 直接托管）、
gld_data.csv、slv_data.csv、llms.txt。只有数据真的变化才会产生 diff，保证无新数据时零提交。

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
S_START = "/* __SLV_DATA_START__ */"
S_END = "/* __SLV_DATA_END__ */"
SLV_URL = (
    "https://www.blackrock.com/varnish-api/blk-one01-product-data/"
    "product-data/api/v1/get-fund-document?appType=PRODUCT_PAGE&appSubType=ISHARES"
    "&targetSite=us-ishares&locale=en_US&portfolioId=239855&component=fundDownload"
    "&userType=individual"
)
SLV_SHEET = "Historical"
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
        w = csv.writer(f, lineterminator="\n")
        w.writerow(["date", "price", "tonnes"])
        for d, p, t in rows:
            w.writerow([d.isoformat(), p, t])
    return path


def download_slv(retries=3):
    """iShares Silver Trust 官方 fund 下载（SpreadsheetML XML，非 xlsx）。"""
    last = None
    for i in range(retries):
        try:
            req = urllib.request.Request(SLV_URL, headers={
                "User-Agent": UA, "Accept": "*/*",
                "Referer": "https://www.ishares.com/us/products/239855/ishares-silver-trust"})
            with urllib.request.urlopen(req, timeout=90) as resp:
                data = resp.read()
            if not data.startswith(b"<?xml"):
                raise ValueError("返回内容不是 SpreadsheetML")
            return data
        except Exception as e:
            last = e
            print("SLV 下载第 %d 次失败: %s" % (i + 1, e), flush=True)
            time.sleep(3 * (i + 1))
    raise RuntimeError("SLV 下载失败: %s" % last)


def parse_slv(raw):
    """Historical 表只有 As Of / NAV per Share / Ex-Dividends / Shares Outstanding。
    持仓用份额数（换算成百万份）代理实物吨数，价格用 NAV/份代理收盘价。"""
    import xml.etree.ElementTree as ET

    text = raw.decode("utf-8", "replace")
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", text)
    # 官方文件 HRef 属性里有未转义的裸 &，先规范化，否则 XML 解析失败
    text = re.sub(r"&(?!(?:amp|lt|gt|quot|apos|#[0-9]+|#x[0-9a-fA-F]+);)", "&amp;", text)
    ns = "{urn:schemas-microsoft-com:office:spreadsheet}"
    root = ET.fromstring(text)
    ws = None
    for w in root.findall(ns + "Worksheet"):
        if w.get(ns + "Name") == SLV_SHEET:
            ws = w
            break
    if ws is None:
        raise RuntimeError("SLV 文件缺少 Historical 工作表")
    months = {m: i + 1 for i, m in enumerate(
        ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
         "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"])}
    rows = []
    for row in ws.find(ns + "Table").findall(ns + "Row")[1:]:
        vals = []
        for c in row.findall(ns + "Cell"):
            d = c.find(ns + "Data")
            vals.append("" if d is None or d.text is None else d.text.strip())
        if len(vals) < 4:
            continue
        m = re.match(r"^([A-Za-z]{3}) (\d{1,2}), (\d{4})$", vals[0])
        if not m or m.group(1) not in months:
            continue
        try:
            nav = float(vals[1].replace(",", ""))
            shares = float(vals[3].replace(",", ""))
        except ValueError:
            continue
        if nav > 0 and shares > 0:
            rows.append((datetime.date(int(m.group(3)), months[m.group(1)], int(m.group(2))),
                         round(nav, 6), round(shares / 1e6, 4)))
    rows.sort()
    if len(rows) < 100:
        raise RuntimeError("SLV 有效数据过少: %d" % len(rows))
    return rows


def compact_slv(rows):
    return [[d.isoformat(), p, t] for d, p, t in rows]


def update_slv_html(data, updated):
    path = os.path.join(BASE, "index.html")
    html = open(path, encoding="utf-8").read()
    if S_START not in html or S_END not in html:
        raise RuntimeError("index.html 缺少 SLV 数据块标记")
    block = (S_START + "\n"
             + "window.SLV_DATA = " + json.dumps(data, separators=(",", ":")) + ";\n"
             + "window.SLV_UPDATED = " + json.dumps(updated) + ";\n"
             + S_END)
    new, n = re.subn(re.escape(S_START) + r"[\s\S]*?" + re.escape(S_END),
                     lambda m: block, html, count=1)
    if n != 1:
        raise RuntimeError("SLV 数据块替换失败")
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write(new)
    return len(new.encode("utf-8"))


def write_slv_csv(rows):
    path = os.path.join(BASE, "slv_data.csv")
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, lineterminator="\n")
        w.writerow(["date", "nav_per_share", "shares_million"])
        for d, p, t in rows:
            w.writerow([d.isoformat(), p, t])
    return path


def read_slv_csv():
    """白银下载失败时沿用上一次快照，保证黄金主线不被阻断。"""
    path = os.path.join(BASE, "slv_data.csv")
    if not os.path.exists(path):
        return []
    out = []
    for r in list(csv.reader(open(path, encoding="utf-8")))[1:]:
        try:
            out.append((datetime.date.fromisoformat(r[0]), float(r[1]), float(r[2])))
        except (ValueError, IndexError):
            continue
    return out


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


def week_key(d):
    return d - datetime.timedelta(days=d.weekday())


def d_series(rows, win=52):
    """在品种自己的连续周序列上算 Z-score 与 D，不受跨品种对齐影响。"""
    wk = weekly(rows)
    zp = roll_z([r[1] for r in wk], win)
    zt = roll_z([r[2] for r in wk], win)
    D = [None if zp[i] is None or zt[i] is None else zt[i] - zp[i]
         for i in range(len(wk))]
    valid = [v for v in D if v is not None]
    cur = D[-1] if D else None
    pct = (sum(1 for v in valid if v <= cur) / len(valid) * 100) if cur is not None and valid else 0.0
    return wk, D, cur, pct


def corr_d(rows_g, rows_s, win=52):
    """按共同周键对齐两条 D 序列，算相关系数与共同周数。"""
    wg, dg, _, _ = d_series(rows_g, win)
    ws, ds, _, _ = d_series(rows_s, win)
    mg = {week_key(r[0]): i for i, r in enumerate(wg)}
    ms = {week_key(r[0]): i for i, r in enumerate(ws)}
    pairs = [(dg[mg[k]], ds[ms[k]]) for k in sorted(set(mg) & set(ms))]
    pairs = [(x, y) for x, y in pairs if x is not None and y is not None]
    if len(pairs) < 3:
        return 0.0, 0
    mx = sum(x for x, _ in pairs) / len(pairs)
    my = sum(y for _, y in pairs) / len(pairs)
    nu = sum((x - mx) * (y - my) for x, y in pairs)
    dx = sum((x - mx) ** 2 for x, _ in pairs) ** 0.5
    dy = sum((y - my) ** 2 for _, y in pairs) ** 0.5
    return (nu / (dx * dy) if dx * dy else 0.0), len(pairs)


def write_llms(rows, slv_rows=None):
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
    s = list(slv_rows or [])
    if s:
        s_wk, _, s_cur, s_pct = d_series(s, win)
        s_last = s[-1]
        s_start, s_end = s[0][0].isoformat(), s_last[0].isoformat()
        s_nd, s_nw = len(s), len(s_wk)
        s_price, s_shares, s_d = s_last[1], s_last[2], s_cur
        corr, n_common = corr_d(rows, s, win)
    else:
        s_start = s_end = "—"
        s_nd = s_nw = n_common = 0
        s_price = s_shares = s_d = s_pct = 0.0
        corr = 0.0
    txt = """# 黄金 / 白银 ETF 价格 x 持仓背离 - 给 AI 的文本摘要

本文件由 fetch_and_build.py 自动生成，与 index.html / gld_data.csv / slv_data.csv 同源同日、同口径。
请勿手改；改文案请改 fetch_and_build.py 的 write_llms()。

## 数据快照
- 数据截止日: {end}
- 覆盖区间: {start} ~ {end}
- 有效交易日: {nd}
- 周频聚合: {nw} 期（滚动窗口 {win} 期）
- 最新 GLD 收盘: {price:.2f} USD
- 最新实物持仓: {tonnes:.2f} 吨
- 白银有效交易日: {s_nd}，周频 {s_nw} 期；白银数据截止 {s_end}

## 黄金最新读数（周频 / 全历史口径）
- 背离强度 D = Z(持仓) - Z(价格) = {d:+.2f}
- 历史分位 = {pct:.1f}%（历史上约 {rest:.1f}% 的时间比现在更背离）
- 读法: D > 0 表示持仓相对价格偏强，常见于价格下跌时仍有逆势增持；
  D > 2 才算历史罕见，当前属于「偏强，值得留意」，不是买卖信号。

## 白银与金银对照（顶部「品种」切换）
- 页面顶部可切 黄金 GLD / 白银 SLV / 金银对照，三者共用同一根时间轴 S.view 与同一套
  hover / 滚轮缩放 / 拖拽平移 / 双击重置 / 导航条框选联动，切品种不会重置视图。
- 白银数据源是 iShares Silver Trust (SLV) 官方 fund 下载（BlackRock get-fund-document 端点）。
  官方 Historical 表是 As Of / NAV per Share / Ex-Dividends / Shares Outstanding 四列（Ex-Dividends 基本为空），
  **没有逐日实物吨数，也没有市场收盘价**；产品页显示的吨数是当前快照，不是历史序列。
- 白银采用**代理口径**：持仓 = Shares Outstanding（换算成百万份），价格 = NAV per Share。
  两者来自同一份官方文件、同一交易日，同源同频；份额数不受银价影响，不是 AUM。
  实物吨数 = 份额数 × 每份含银量，每份含银量只因管理费缓慢摊薄（年化约 -0.5%），
  比申赎波动小两个数量级，52 周滚动 Z-score 基本吸收该慢趋势。**但它不是吨数，页面已明确标注**。
- 白银历史 {s_start} ~ {s_end}（{s_nd} 个交易日 / 周频 {s_nw} 期），比黄金晚约 1.5 年。
  所有模式共用**黄金聚合序列作 master axis**（周频 {nw} 期），白银按同一个周键 / 月键对齐进来；
  前 74 周白银无数据就留空、不出 Z，切品种不重置 S.view。对照图里白银线早期留空，
  那是数据缺口不是 D = 0；白银单品种视图把时间拉到 2006-04 之前时，参照图 / 图 1 也会同样留空。
- 当前读数（周频 / 各自全历史口径）：黄金 D {d:+.2f}（分位 {pct:.1f}%）；
  白银 D {s_d:+.2f}（分位 {s_pct:.1f}%）。白银最新 NAV/份 {s_price:.4f} USD /
  流通份额 {s_shares:.2f} 百万份（数据截止 {s_end}）。
- 金银 D 都是标准差单位，可放同一张对照图比较：共同 {n_common} 周相关系数 r = {corr:.2f}。
  **绝对价格、绝对持仓水平不可跨品种比较**；白银 D 的历史分布比黄金宽，
  所以「+2 属罕见」这个经验阈值对白银要更谨慎地解释。

## 口径（改动前必读）
- 价格与实物持仓吨数来自同一份 SPDR Gold Shares 官方 Historical Archive，
  同一交易日、同一标的，同源同频，不存在跨源时点错配。
- 用滚动 Z-score（总体标准差，除以 n），不是全程标准化；缩放视图只改可视区间，不改统计口径。
- 持仓用实物「吨数」，不用 AUM：金价跌 10% 会让 AUM 自动缩水，即使一克都没被赎回。
- 散点用「持仓变化」而非「持仓水平」，避免共同趋势造成伪相关。
- 频率切换会改变读数（日 252 / 周 52 / 月 36），跨频率的数值不可直接互比。
- 交易日口径：只用 Historical Archive 的 Tonnes of Gold 列，不混用按结算日的金条清单。
- **金银口径铁律**：黄金用「实物吨数 + 市场收盘价」，白银用「流通份额（百万份）代理 + NAV/份」。
  白银官方表缺逐日吨数与收盘价，绝不能把份额数标成吨、把 NAV 标成收盘价；
  绝对价格 / 绝对持仓水平不可跨品种比较，只有标准化后的 Z、D 与变化率可比。
  白银 rollZ 遇到含 null 的窗口直接返回 null，不用 0 / 前值填充；变化量用「上一期有效值」计算。

## 参照图（位于图 1 与图 2 之间）
- 黄金模式画 GLD 每份份额的**原始收盘价（USD/份）**，不是 Z-score、不是现货金价 USD/oz，不做对数或归一化；
  切到白银模式画的是 **SLV 的 NAV/份**，不是市场收盘价、不是现货银价 USD/oz，页面已明确标注。
  全历史价格区间 {pmin:.2f} ~ {pmax:.2f} USD；数据末期 {price:.2f} USD。
- 单轴，不叠加持仓第二条线，避免双轴截断刻度的视觉误导；持仓真实值（黄金吨 / 白银百万份）在悬停浮层里给出。
- 纵轴默认按当前视图 min/max 自适应并在图上标出范围，可一键切「从 0 起」看长期真实幅度。
- 与图 1 的 |D| > 1.5σ 浅色竖带共用 strongSegs()，位置完全一致，方便对照「背离高发时价格处在什么位置」。
- 与图 1 / 图 2 / 图 2b 共用 axisX() / bindChart() / showHover() 和同一个 S.view；四图 L=52 / R=18，
  hover 竖线像素级对齐，滚轮缩放 / 拖拽平移 / 双击重置 / 导航条框选全部同步。
- 图 1 里橙色的「价格 Z」线是它的滚动 Z-score（白银模式参照图用银灰线区分品种，两条线颜色不同但数据同源）。判断背离用图 1 / 图 2，确认当时绝对价位用参照图，两者不可混读。

## 图 2b 金银 D 对照（仅「金银对照」模式显示）
- 金色线 = 黄金 D，银灰线 = 白银 D；两条都是标准差单位，可直接同图比高低与方向。
- 右上角 r 是**当前视图**内两条 D 的 Pearson 相关系数；全样本共同 {n_common} 周 r = {corr:.2f}。
- 2006-04 之前只有黄金线，是数据缺口；两条线在滚动窗口未填满时都不出值，
  所以黄金线也晚于页面起点约 52 期，这不是数据丢失。
- 切到「金银对照」时，上方图 1 / 参照图 / 图 2 / 图 3 与统计卡片仍显示**黄金**（主品种）；
  要看白银自己的全套图请切「白银 SLV」。三种模式共用同一视图，切品种不重置缩放。

## 图 3 四象限散点怎么读（最容易看懵的一张）
- 一个点 = 某一期相对上一期的变化，不是价格水平，也不是持仓总量。
- 横轴 x = (本期价格 / 上期价格 - 1) * 100，单位 %；竖线左边 = 价跌，右边 = 价涨。
- 纵轴 y = 本期持仓 - 上期持仓；黄金单位是吨，白银单位是百万份（份额代理口径），
  横线上边 = 增持，下边 = 减持。切品种时轴标签与悬停浮层单位会一起变。
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
- SLV 是单只基金，不代表全球白银 ETF 总持仓；其份额数代理口径与黄金的实物吨数口径不同，绝对值不可比。
- 参照图是份额价格（黄金为 GLD 收盘价、白银为 NAV/份），不是现货金价 / 银价 USD/oz，不能换算比价。
- 相关系数只说明同向程度，不说明因果；D 与「背离」都不是涨跌预测。
- 2004-2005 年建仓期吨数快速爬升，属结构性变化，不宜与成熟期直接比较。
- 原始数据版权归发行人所有；本仓库只做衍生分析展示，不收录原始 XLSX。

## 离线兜底导入（顶部「手动导入当前品种 CSV」）
- 这个控件导入的是**当前品种**的数据，三个模式该用的文件不一样，页面会把说明与链接直接显示在控件下方：
  * 黄金 GLD → `gld_data.csv`（三列 `date,price,tonnes`；GLD 收盘价 USD / 实物吨数）。在线：{pages}gld_data.csv
  * 白银 SLV → `slv_data.csv`（三列 `date,nav_per_share,shares_million`；NAV/份 / 流通份额百万份，**不是吨、不是收盘价**）。在线：{pages}slv_data.csv
  * 金银对照 → 主品种是黄金，在这里导入 `gld_data.csv` 只替换上方全套图与图 2b 的金色 D 线；
    要换白银那条线，先切到「白银 SLV」导入 `slv_data.csv`，再切回「金银对照」。**两份数据各自独立，需要分别导入一次**。
- 页面链接同时给相对路径（`gld_data.csv` / `slv_data.csv`，离线双击本地 index.html 时指向同目录文件）
  与完整在线副本（Pages 同源，可直接下载）。导入只替换对应品种的内存序列，不影响另一个品种。
- 导入不写回仓库、不上传；`build()` 会重算并把视图重置为默认近 3 年，刷新页面即恢复官方快照。
- CSV 需三列、日期 `YYYY-MM-DD`、两列数值都 > 0；表头不含 date/price/tonnes 时按第 1/2/3 列依次当作 日期/价格/持仓。

## 链接
- 在线看板: {pages}
- 看板中心（聚合入口，本页为第 8 个板块「黄金ETF背离」）: https://wj-dashboards.pages.dev/
- 本摘要: {pages}llms.txt
- 源码 / 手动刷新: {repo}（Actions -> Refresh GLD + SLV divergence data -> Run workflow）
- 黄金逐日数据: {pages}gld_data.csv
- 白银逐日数据: {pages}slv_data.csv

仅为图表分析，不构成投资建议。
""".format(end=last[0].isoformat(), start=rows[0][0].isoformat(),
           nd=len(rows), nw=len(wk), win=win, price=last[1], tonnes=last[2],
           pmin=pmin, pmax=pmax, d=cur, pct=pct, rest=100 - pct,
           s_start=s_start, s_end=s_end, s_nd=s_nd, s_nw=s_nw,
           s_price=s_price, s_shares=s_shares, s_d=s_d, s_pct=s_pct,
           corr=corr, n_common=n_common, pages=PAGES, repo=REPO_URL)
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

    # 白银是独立数据源：抓不到或解析失败时沿用上一次快照，
    # 既不阻断黄金主线，也不会因为错误数据制造额外提交。
    try:
        print("抓取 iShares SLV 官方历史数据 ...", flush=True)
        s_rows = parse_slv(download_slv())
        size = update_slv_html(compact_slv(s_rows), s_rows[-1][0].isoformat())
        s_csv = write_slv_csv(s_rows)
        print("SLV %d 行，%s ~ %s，CSV %s" % (
            len(s_rows), s_rows[0][0], s_rows[-1][0], s_csv), flush=True)
    except Exception as e:
        print("WARN: 白银更新跳过，沿用上次快照: %s" % e, flush=True)
        s_rows = read_slv_csv()

    llms_path = write_llms(rows, s_rows)
    print("GLD 有效交易日 %d，区间 %s ~ %s" % (len(rows), rows[0][0], rows[-1][0]))
    print("最新黄金: 收盘 %.2f USD / 持仓 %.2f 吨" % (rows[-1][1], rows[-1][2]))
    if s_rows:
        print("最新白银: NAV %.4f USD / 份额 %.2f 百万份" % (s_rows[-1][1], s_rows[-1][2]))
    print("数据截止 %s，index.html %d bytes" % (updated, size))
    print("CSV %s / 文本摘要 %s" % (csv_path, llms_path))


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("ERROR: %s" % e, file=sys.stderr)
        sys.exit(1)
