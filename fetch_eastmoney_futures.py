# -*- coding: utf-8 -*-
"""
东方财富期货 - 交易规则与保证金数据抓取脚本（Excel + HTML 双输出）

抓取三个数据源：
  1. 保证金比例表     https://www.eastmoneyfutures.com/pages/service/jyts.html#jyrl
                     （静态HTML表格）
  2. 品种及交易规则表 https://www.eastmoneyfutures.com/pages/service/jygz.html
                     （后端API JSON，接口 /emfApi/pzjy/getPZJYInfo）
  3. 品种单手保证金一览表 https://qhweb.eastmoney.com/bzj/low
                     （静态HTML表格）

输出（当前文件夹）：
  - 东方财富期货交易规则保证金<YYYY-MM-DD>.xlsx
    工作表：保证金比例 / 交易规则 / 保证金
  - 东方财富期货交易规则保证金<YYYY-MM-DD>.html
    三表切换 + 搜索的美化网页版

Excel 样式：仿宋 13 号、行高 55 磅、冻结首行、单元格自动换行。

依赖：requests, beautifulsoup4, openpyxl, lxml
运行：python fetch_eastmoney_futures.py
"""

import os
import re
import json
from datetime import datetime

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# 常量
# ---------------------------------------------------------------------------
URL_MARGIN_RATIO = "https://www.eastmoneyfutures.com/pages/service/jyts.html#jyrl"
URL_TRADE_RULES = "https://www.eastmoneyfutures.com/emfApi/pzjy/getPZJYInfo"  # jygz.html 页面实际调用的接口
URL_MARGIN_LIST = "https://qhweb.eastmoney.com/bzj/low"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
    "Accept-Language": "zh-CN,zh;q=0.9",
}

OUT_DIR = os.path.dirname(os.path.abspath(__file__))  # 脚本所在目录 = 当前文件夹
DATE_TODAY = datetime.now().strftime("%Y-%m-%d")
OUT_XLSX = os.path.join(OUT_DIR, f"东方财富期货交易规则保证金{DATE_TODAY}.xlsx")
OUT_HTML = os.path.join(OUT_DIR, f"东方财富期货交易规则保证金{DATE_TODAY}.html")


# ---------------------------------------------------------------------------
# 通用工具
# ---------------------------------------------------------------------------
def fetch(url, referer=None, timeout=30, encoding=None):
    """请求网页/接口，返回解码后的文本。

    encoding: 指定解码方式；为 None 时尝试从响应头推断，失败则回退 UTF-8。
              对 JSON 接口应显式传 "utf-8"，避免 apparent_encoding 误判。
    """
    headers = dict(HEADERS)
    if referer:
        headers["Referer"] = referer
    resp = requests.get(url, headers=headers, timeout=timeout)
    resp.raise_for_status()
    if encoding:
        resp.encoding = encoding
    else:
        resp.encoding = resp.apparent_encoding or "utf-8"
    return resp.text


def to_number(value):
    """
    将字符串尽量转为数值：
      "¥3883.2" -> 3883.2 ; "16%" -> 0.16 ; "43%" -> 0.43
    无法转换（如 "±20%"）时原样返回字符串。
    """
    if value is None:
        return ""
    text = str(value).strip()
    if not text:
        return ""
    m = re.fullmatch(r"[¥￥]\s*([+-]?[\d,]+(?:\.\d+)?)", text)
    if m:
        return float(m.group(1).replace(",", ""))
    m = re.fullmatch(r"([+-]?[\d.]+)\s*%", text)
    if m:
        return float(m.group(1)) / 100.0
    return text


# ---------------------------------------------------------------------------
# 1. 保证金比例表（jyts.html 静态表格）
# ---------------------------------------------------------------------------
def parse_margin_ratio(html):
    """解析 jyts.html 中 id='jyrl_tb' 的表格。

    结构：表头在一个 tbody 中，数据在 class='tbox' data-key='交易所' 的 tbody 中，
    每个数据行 = 品种/代码/主力合约/每日涨跌幅度/公司保证金比例/特殊合约参数调整/备注。
    """
    soup = BeautifulSoup(html, "lxml")
    table = soup.find("table", id="jyrl_tb")
    if table is None:
        raise RuntimeError("未在 jyts.html 中找到 id='jyrl_tb' 的表格，页面结构可能已变更")

    rows = []
    for tb in table.find_all("tbody"):
        classes = tb.get("class") or []
        if "tbox" not in classes:
            continue
        exchange = (tb.get("data-key") or "").strip()
        for tr in tb.find_all("tr"):
            tds = tr.find_all("td")
            # 组首行因 rowspan 带交易所单元格共 8 列，其余行 7 列
            if len(tds) == 8:
                cells = tds[1:]
            elif len(tds) == 7:
                cells = tds
            else:
                continue
            rows.append([
                exchange,
                cells[0].get_text(strip=True),   # 品种
                cells[1].get_text(strip=True),   # 代码
                cells[2].get_text(strip=True),   # 主力合约
                cells[3].get_text(strip=True),   # 每日涨跌幅度
                cells[4].get_text(strip=True),   # 公司保证金比例
                cells[5].get_text(strip=True),   # 特殊合约参数调整
                cells[6].get_text(strip=True),   # 备注
            ])
    return rows


# ---------------------------------------------------------------------------
# 2. 品种及交易规则表（jygz.html 后端接口 /emfApi/pzjy/getPZJYInfo）
# ---------------------------------------------------------------------------
def parse_trade_rules(api_json):
    """
    接口返回结构：
      {"code":"000","data":{交易所:{最后交易日:{最后交割日期:[品种对象,...]},...},...}}
    品种对象字段：exchange / varietiesCN / varietiesEN / minimumPrice / jgfy(交割月份)
                  / zhjyr(最后交易日) / zhjgrq(最后交割日期) / jydw(交易单位)
    展平为行：[交易所, 交易品种, 最小变动价位, 交割月份, 最后交易日, 最后交割日期, 交易单位]
    交易品种列与网页一致，为 "品种中文+代码"（如 燃料油FU）。
    """
    if not isinstance(api_json, dict):
        raise RuntimeError("交易规则接口返回格式异常")
    if str(api_json.get("code")) != "000":
        raise RuntimeError(f"交易规则接口返回错误码: {api_json.get('code')}")

    data = api_json.get("data") or {}
    rows = []
    for exchange, zhjyr_map in data.items():
        if not isinstance(zhjyr_map, dict):
            continue
        for zhjyr, zhjgrq_map in zhjyr_map.items():
            if not isinstance(zhjgrq_map, dict):
                continue
            for zhjgrq, varieties in zhjgrq_map.items():
                if not isinstance(varieties, list):
                    continue
                for item in varieties:
                    if not isinstance(item, dict):
                        continue
                    rows.append([
                        item.get("exchange") or exchange,
                        (item.get("varietiesCN", "") or "") + (item.get("varietiesEN", "") or ""),
                        item.get("minimumPrice", ""),
                        item.get("jgfy", ""),
                        item.get("zhjyr", ""),
                        item.get("zhjgrq", ""),
                        item.get("jydw", ""),
                    ])
    return rows


# ---------------------------------------------------------------------------
# 3. 品种单手保证金一览表（qhweb.eastmoney.com/bzj/low 静态表格）
# ---------------------------------------------------------------------------
def parse_margin_list(html):
    """
    页面为服务端渲染的 Vue 表格：
      <div class="t_tr" ...>
        <span>品种</span><span>合约代码</span><span>¥3883.2</span><span>16%</span><span>备注</span>
      </div>
    同时提取 <div class="update">更新于 ...</div> 与 <div class="note_border">注：...</div>。
    """
    soup = BeautifulSoup(html, "lxml")

    # 更新时间
    update_el = soup.select_one("div.update")
    update_time = update_el.get_text(strip=True) if update_el else ""

    # 底部注释
    note_el = soup.select_one("div.note_border")
    note = note_el.get_text(strip=True) if note_el else ""

    # 数据行：每个 t_tr 有 5 个 span，前 4 个必须有内容，第 3 个为金额
    rows = []
    for tr in soup.select("div.t_body div.t_tr"):
        spans = tr.find_all("span")
        if len(spans) < 5:
            continue
        cells = [s.get_text(strip=True) for s in spans[:5]]
        if not cells[0] or not cells[1] or not re.search(r"[¥￥]", cells[2]) or not cells[3]:
            continue
        rows.append(cells)  # [品种, 合约代码, 保证金(元), 保证金率, 备注]

    return rows, update_time, note


# ---------------------------------------------------------------------------
# Excel 写入（样式：仿宋13号 / 行高55磅 / 冻结首行 / 自动换行）
# ---------------------------------------------------------------------------
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(name="仿宋", bold=True, color="FFFFFF", size=13)
BODY_FONT = Font(name="仿宋", size=13)
HEADER_ALIGN = Alignment(vertical="center", horizontal="center", wrap_text=True)
BODY_ALIGN = Alignment(vertical="center", wrap_text=True)


def write_sheet(ws, headers, rows, number_formats=None, extra_rows=None, row_height=55):
    """
    写入一个工作表：首行表头（冻结 + 深蓝底白字），数据行 55 磅高、仿宋13号、自动换行。
    number_formats: dict {列号(0起): openpyxl数字格式}，仅对数值单元格生效。
    extra_rows: 追加在数据之后的行（如注释说明），第一列为文本。
    row_height: 数据行行高（磅），默认 55。
    """
    # 表头
    for c, title in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=c, value=title)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = HEADER_ALIGN
        cell.border = BORDER
    ws.row_dimensions[1].height = row_height

    # 数据
    for r, row in enumerate(rows, start=2):
        ws.row_dimensions[r].height = row_height
        for c, val in enumerate(row, start=1):
            cell = ws.cell(row=r, column=c, value=val)
            cell.font = BODY_FONT
            cell.alignment = BODY_ALIGN
            cell.border = BORDER
            if number_formats and (c - 1) in number_formats and isinstance(val, (int, float)):
                cell.number_format = number_formats[c - 1]

    # 追加说明行（合并 A 列到最后列）
    if extra_rows:
        last_col = len(headers)
        for r, text in enumerate(extra_rows, start=ws.max_row + 1):
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=last_col)
            cell = ws.cell(row=r, column=1, value=text)
            cell.font = Font(name="仿宋", italic=True, size=11, color="808080")
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            ws.row_dimensions[r].height = 45

    # 列宽自适应（按内容长度，限制最大 45）
    for c in range(1, len(headers) + 1):
        width = max(len(str(headers[c - 1])), 8)
        for row in rows:
            v = row[c - 1]
            width = max(width, len(str(v)) if not isinstance(v, float) else 10)
        ws.column_dimensions[get_column_letter(c)].width = min(width + 4, 45)

    ws.freeze_panes = "A2"  # 冻结首行


# ---------------------------------------------------------------------------
# HTML 生成（三表切换 + 搜索的美化网页版）
# ---------------------------------------------------------------------------
def build_html(ratio_rows, rule_rows, margin_rows, update_time, note, fetch_time):
    """把三个表的数据内嵌为静态 HTML（数据为抓取快照，双击即可打开）。"""
    # 保证金表数据转成便于 JS 展示的数值
    margin_data = []
    for row in margin_rows:
        margin_data.append([row[0], row[1], to_number(row[2]), to_number(row[3]), row[4]])

    def js(obj):
        # 转义 < 防止内联 JSON 意外闭合脚本标签
        return json.dumps(obj, ensure_ascii=False).replace("<", "\\u003c")

    payload = js({
        "fetchTime": fetch_time,
        "updateTime": update_time,
        "ratio": {"headers": ["交易所", "品种", "代码", "主力合约", "每日涨跌幅度",
                              "公司保证金比例", "特殊合约参数调整", "备注"], "rows": ratio_rows},
        "rules": {"headers": ["交易所", "交易品种", "最小变动价位", "交割月份",
                              "最后交易日", "最后交割日期", "交易单位"], "rows": rule_rows},
        "margin": {"headers": ["品种", "合约代码", "保证金(元)", "保证金率", "备注"], "rows": margin_data},
    })

    html = HTML_TEMPLATE.replace("__PAYLOAD__", payload)
    return html


HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>东方财富期货 · 交易规则与保证金一览</title>
<style>
  :root {
    --navy: #10294b;
    --navy2: #1e4e79;
    --blue: #2f6db3;
    --gold: #c9a227;
    --bg: #eef2f7;
    --card: #ffffff;
    --line: #e2e8f0;
    --text: #24344d;
    --muted: #6b7a93;
  }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: "FangSong", "仿宋", "STFangsong", "SimSun", serif;
    background: linear-gradient(160deg, #dde7f3 0%, var(--bg) 40%);
    color: var(--text);
    min-height: 100vh;
  }
  .wrap { max-width: 1280px; margin: 0 auto; padding: 0 20px 40px; }

  /* 头部 */
  header {
    background: linear-gradient(120deg, var(--navy) 0%, var(--navy2) 55%, #2a5f9e 100%);
    color: #fff;
    padding: 30px 0 26px;
    box-shadow: 0 6px 18px rgba(16, 41, 75, .28);
    position: sticky; top: 0; z-index: 20;
  }
  header .wrap { padding-bottom: 0; }
  .head-row { display: flex; align-items: flex-end; justify-content: space-between; gap: 16px; flex-wrap: wrap; }
  .head-title { display: flex; align-items: center; gap: 14px; }
  .head-title .logo {
    width: 46px; height: 46px; border-radius: 12px;
    background: linear-gradient(135deg, var(--gold), #e8c94f);
    display: flex; align-items: center; justify-content: center;
    font-weight: 700; font-size: 22px; color: var(--navy);
    box-shadow: 0 3px 8px rgba(0,0,0,.25);
  }
  .head-title h1 { font-size: 26px; letter-spacing: 1px; font-weight: 700; }
  .head-title p { font-size: 13px; opacity: .82; margin-top: 4px; letter-spacing: .5px; }
  .badges { display: flex; gap: 8px; flex-wrap: wrap; }
  .badge {
    background: rgba(255,255,255,.14); border: 1px solid rgba(255,255,255,.28);
    padding: 5px 12px; border-radius: 999px; font-size: 12.5px; backdrop-filter: blur(2px);
  }
  .badge b { color: #ffe9a3; font-weight: 600; }

  /* 控制条 */
  .controls {
    margin: 26px 0 18px; display: flex; align-items: center; justify-content: space-between;
    gap: 14px; flex-wrap: wrap;
  }
  .tabs { display: flex; gap: 8px; }
  .tab {
    font-family: inherit; font-size: 15px; cursor: pointer;
    padding: 9px 22px; border: 1px solid var(--line); border-radius: 999px;
    background: var(--card); color: var(--muted); transition: all .18s ease;
    box-shadow: 0 1px 3px rgba(16,41,75,.06);
  }
  .tab:hover { color: var(--navy2); border-color: var(--blue); }
  .tab.active {
    background: linear-gradient(120deg, var(--navy2), var(--blue));
    color: #fff; border-color: transparent; box-shadow: 0 4px 10px rgba(47,109,179,.35);
  }
  .search-box { position: relative; }
  .search-box input {
    font-family: inherit; font-size: 14px; width: 240px;
    padding: 9px 14px 9px 34px; border: 1px solid var(--line); border-radius: 999px;
    outline: none; background: var(--card); color: var(--text);
    transition: all .18s ease; box-shadow: 0 1px 3px rgba(16,41,75,.06);
  }
  .search-box input:focus { border-color: var(--blue); box-shadow: 0 0 0 3px rgba(47,109,179,.15); }
  .search-box .icon {
    position: absolute; left: 12px; top: 50%; transform: translateY(-50%);
    color: var(--muted); font-size: 14px; pointer-events: none;
  }

  /* 表格卡片 */
  .panel { display: none; }
  .panel.active { display: block; animation: fadeIn .22s ease; }
  @keyframes fadeIn { from { opacity: 0; transform: translateY(6px); } to { opacity: 1; transform: none; } }
  .card {
    background: var(--card); border-radius: 14px; overflow: hidden;
    box-shadow: 0 8px 24px rgba(16,41,75,.10); border: 1px solid var(--line);
  }
  .card-head {
    padding: 14px 20px; display: flex; align-items: center; justify-content: space-between;
    border-bottom: 1px solid var(--line); background: linear-gradient(180deg, #fbfdff, #f5f8fc);
  }
  .card-head h2 { font-size: 17px; color: var(--navy2); display: flex; align-items: center; gap: 10px; }
  .card-head h2::before {
    content: ""; width: 5px; height: 18px; border-radius: 3px;
    background: linear-gradient(180deg, var(--gold), var(--blue)); display: inline-block;
  }
  .card-head .count { font-size: 12.5px; color: var(--muted); }
  .card-head .count b { color: var(--navy2); font-size: 14px; }
  .table-scroll { overflow-x: auto; }
  table { border-collapse: separate; border-spacing: 0; width: 100%; font-size: 13.5px; }
  thead th {
    background: linear-gradient(180deg, var(--navy2), #2a5f9e);
    color: #fff; font-weight: 600; font-size: 14px;
    padding: 11px 12px; text-align: left; white-space: nowrap;
    border-bottom: 2px solid rgba(255,255,255,.25);
  }
  tbody td {
    padding: 9px 12px; border-bottom: 1px solid var(--line);
    vertical-align: middle; line-height: 1.45;
  }
  tbody tr:nth-child(even) { background: #f8fafd; }
  tbody tr:hover { background: #eaf1fa; }
  td.num, th.num { text-align: right; font-variant-numeric: tabular-nums; }
  .tag {
    display: inline-block; padding: 2px 8px; border-radius: 6px; font-size: 12px;
    background: #eef3fa; color: var(--navy2); border: 1px solid #dbe6f2; white-space: nowrap;
  }
  .tag.warn { background: #fdf3e3; color: #9a6a12; border-color: #f2dfb5; }
  .tag.special { background: #e8f5ee; color: #22724c; border-color: #cdeadd; }
  .empty { padding: 30px; text-align: center; color: var(--muted); font-size: 14px; }

  /* 页脚 */
  footer { margin-top: 26px; color: var(--muted); font-size: 12.5px; line-height: 1.7; }
  footer .note {
    background: #fff8e6; border: 1px solid #f0e0b0; border-left: 4px solid var(--gold);
    padding: 12px 16px; border-radius: 8px; margin-bottom: 12px;
  }
  footer .sources a { color: var(--blue); text-decoration: none; }
  footer .sources a:hover { text-decoration: underline; }
  footer .risk { color: #8a93a5; }

  @media (max-width: 760px) {
    header .head-row { flex-direction: column; align-items: flex-start; }
    .search-box input { width: 100%; }
    .controls { align-items: stretch; }
    .tabs { width: 100%; justify-content: space-between; }
    .tab { flex: 1; text-align: center; padding: 8px 6px; font-size: 13.5px; }
  }
</style>
</head>
<body>
<header>
  <div class="wrap">
    <div class="head-row">
      <div class="head-title">
        <div class="logo">东</div>
        <div>
          <h1>东方财富期货 · 交易规则与保证金一览</h1>
          <p>保证金比例表 · 品种及交易规则表 · 品种单手保证金一览表</p>
        </div>
      </div>
      <div class="badges">
        <span class="badge">数据抓取 <b id="fetchTime"></b></span>
        <span class="badge">行情更新 <b id="updateTime"></b></span>
      </div>
    </div>
  </div>
</header>

<div class="wrap">
  <div class="controls">
    <div class="tabs" id="tabs">
      <button class="tab active" data-tab="ratio">保证金比例</button>
      <button class="tab" data-tab="rules">交易规则</button>
      <button class="tab" data-tab="margin">单手保证金</button>
    </div>
    <div class="search-box">
      <span class="icon">&#128269;</span>
      <input id="search" type="text" placeholder="搜索品种 / 代码 / 交易所…">
    </div>
  </div>

  <div class="panel active" id="panel-ratio"></div>
  <div class="panel" id="panel-rules"></div>
  <div class="panel" id="panel-margin"></div>

  <footer>
    <div class="note" id="noteText"></div>
    <div class="sources">
      数据来源：
      <a href="https://www.eastmoneyfutures.com/pages/service/jyts.html#jyrl" target="_blank" rel="noopener">东方财富期货·保证金率</a> ·
      <a href="https://www.eastmoneyfutures.com/pages/service/jygz.html" target="_blank" rel="noopener">东方财富期货·品种及交易规则</a> ·
      <a href="https://qhweb.eastmoney.com/bzj/low" target="_blank" rel="noopener">东方财富期货·品种单手保证金一览表</a>
    </div>
    <div class="risk" style="margin-top:6px">免责声明：本页面数据仅供信息参考，不构成任何投资建议或承诺。保证金会随行情波动，交易所与期货公司也可能进行调整，具体以东方财富期货APP交易界面“保证金率”为准。</div>
  </footer>
</div>

<script>
(function () {
  var DATA = __PAYLOAD__;

  document.getElementById('fetchTime').textContent = DATA.fetchTime;
  document.getElementById('updateTime').textContent = DATA.updateTime;
  document.getElementById('noteText').textContent = DATA.updateTime + '；' + (DATA.ratio.rows.length + DATA.rules.rows.length + DATA.margin.rows.length) + ' 个品种数据条目，抓取自东方财富期货官网。';

  var panels = { ratio: 'panel-ratio', rules: 'panel-rules', margin: 'panel-margin' };
  var titles = { ratio: '保证金比例表', rules: '品种及交易规则表', margin: '品种单手保证金一览表' };
  var current = 'ratio';
  var keyword = '';

  function fmtMoney(v) { return '¥' + v.toLocaleString('zh-CN', {minimumFractionDigits: 1, maximumFractionDigits: 1}); }
  function fmtPct(v) { return (v * 100).toFixed(1) + '%'; }

  function buildPanel(key) {
    var conf = DATA[key];
    var el = document.getElementById(panels[key]);
    var rows = conf.rows.filter(function (row) {
      if (!keyword) return true;
      return row.some(function (cell) { return String(cell).toLowerCase().indexOf(keyword) > -1; });
    });
    var html = '';
    html += '<div class="card"><div class="card-head"><h2>' + titles[key] + '</h2>' +
            '<span class="count">共 <b>' + rows.length + '</b> 条' +
            (keyword ? '（筛选自 ' + conf.rows.length + ' 条）' : '') + '</span></div>' +
            '<div class="table-scroll"><table><thead><tr>';
    conf.headers.forEach(function (h, i) {
      html += '<th' + (key === 'margin' && (i === 2) ? ' class="num"' : '') + '>' + h + '</th>';
    });
    html += '</tr></thead><tbody>';
    if (!rows.length) { html += '<tr><td colspan="' + conf.headers.length + '" class="empty">未找到匹配记录</td></tr>'; }
    rows.forEach(function (row) {
      html += '<tr>';
      row.forEach(function (cell, i) {
        var cls = '';
        var text = cell;
        if (key === 'margin') {
          if (i === 2) { cls = 'num'; text = fmtMoney(cell); }
          if (i === 3) { cls = 'num'; text = fmtPct(cell); }
          if (i === 4 && cell) {
            var s = String(cell);
            var cls2 = s.indexOf('特殊') > -1 ? 'special' : (s.indexOf('不活跃') > -1 ? 'warn' : '');
            text = '<span class="tag ' + cls2 + '">' + s + '</span>';
          }
        }
        if (key === 'ratio' && i === 5 && typeof cell === 'string' && cell.indexOf('%') > -1 && cell.indexOf('±') === -1) {
          text = '<span class="tag">' + cell + '</span>';
        }
        html += '<td' + (cls ? ' class="' + cls + '"' : '') + '>' + text + '</td>';
      });
      html += '</tr>';
    });
    html += '</tbody></table></div></div>';
    el.innerHTML = html;
  }

  document.getElementById('tabs').addEventListener('click', function (e) {
    var btn = e.target.closest('.tab');
    if (!btn) return;
    document.querySelectorAll('.tab').forEach(function (t) { t.classList.remove('active'); });
    btn.classList.add('active');
    current = btn.getAttribute('data-tab');
    Object.keys(panels).forEach(function (k) {
      document.getElementById(panels[k]).classList.toggle('active', k === current);
    });
    buildPanel(current);
  });

  var timer = null;
  document.getElementById('search').addEventListener('input', function (e) {
    clearTimeout(timer);
    timer = setTimeout(function () {
      keyword = e.target.value.trim().toLowerCase();
      buildPanel(current);
    }, 150);
  });

  buildPanel('ratio');
})();
</script>
</body>
</html>
"""


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------
def main():
    print(f"[1/3] 抓取保证金比例表: {URL_MARGIN_RATIO}")
    html1 = fetch(URL_MARGIN_RATIO, referer="https://www.eastmoneyfutures.com/pages/service/index.html")
    ratio_rows = parse_margin_ratio(html1)
    print(f"      解析到 {len(ratio_rows)} 行")

    print(f"[2/3] 抓取品种及交易规则表: {URL_TRADE_RULES}")
    raw2 = fetch(URL_TRADE_RULES, referer="https://www.eastmoneyfutures.com/pages/service/jygz.html",
                 encoding="utf-8")
    try:
        api_json = json.loads(raw2)
    except json.JSONDecodeError:
        raise RuntimeError("交易规则接口未返回合法 JSON，可能已改版或需要额外参数")
    rule_rows = parse_trade_rules(api_json)
    print(f"      解析到 {len(rule_rows)} 行")

    print(f"[3/3] 抓取品种单手保证金一览表: {URL_MARGIN_LIST}")
    html3 = fetch(URL_MARGIN_LIST, referer="https://qhweb.eastmoney.com/bzj/low")
    margin_rows, update_time, note = parse_margin_list(html3)
    print(f"      解析到 {len(margin_rows)} 行, 更新时间: {update_time}")

    fetch_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # ---- Excel（仿宋13号 / 行高55磅 / 冻结首行 / 自动换行） ----
    wb = Workbook()

    ws1 = wb.active
    ws1.title = "保证金比例"
    write_sheet(
        ws1,
        ["交易所", "品种", "代码", "主力合约", "每日涨跌幅度", "公司保证金比例",
         "特殊合约参数调整", "备注"],
        ratio_rows,
    )

    ws2 = wb.create_sheet("交易规则")
    write_sheet(
        ws2,
        ["交易所", "交易品种", "最小变动价位", "交割月份", "最后交易日", "最后交割日期", "交易单位"],
        rule_rows,
    )

    ws3 = wb.create_sheet("保证金")
    margin_sorted = [[r[0], r[1], to_number(r[2]), to_number(r[3]), r[4]] for r in margin_rows]
    write_sheet(
        ws3,
        ["品种", "合约代码", "保证金(元)", "保证金率", "备注"],
        margin_sorted,
        number_formats={2: "¥#,##0.0", 3: "0.0%"},
        extra_rows=[note] if note else None,
    )

    wb.save(OUT_XLSX)
    print(f"\nExcel 已保存: {OUT_XLSX}")

    # ---- HTML（三表切换 + 搜索的美化网页版） ----
    html = build_html(ratio_rows, rule_rows, margin_rows, update_time, note, fetch_time)
    with open(OUT_HTML, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"HTML 已保存: {OUT_HTML}")
    print("工作表/页面: 保证金比例 / 交易规则 / 保证金")


if __name__ == "__main__":
    main()
