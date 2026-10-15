<p align="center"><img src="assets/logo.png" width="120" alt="eastmoney-futures-margin"></p>

# 东方财富期货 · 交易规则与保证金一览

抓取东方财富期货官网的五个数据源，生成 **Excel + HTML** 双输出（五表切换 + 搜索的美化网页版）。

## 数据源

| 表 | 来源 |
|----|------|
| 保证金比例表 | https://www.eastmoneyfutures.com/pages/service/jyts.html#jyrl |
| 品种及交易规则表 | https://www.eastmoneyfutures.com/pages/service/jygz.html （后端接口 `/emfApi/pzjy/getPZJYInfo`） |
| 品种单手保证金一览表 | https://qhweb.eastmoney.com/bzj/low |
| 期货手续费标准 | https://www.eastmoneyfutures.com/pages/service/sxf.html |
| 期权手续费标准 | https://www.eastmoneyfutures.com/pages/service/qqsxf.html |

## 使用

```bash
# 安装依赖
pip install requests beautifulsoup4 openpyxl lxml

# 运行（在当前目录生成 东方财富期货交易规则保证金<日期>.xlsx 和 .html）
python fetch_eastmoney_futures.py
```

## 产物

- `东方财富期货交易规则保证金<YYYY-MM-DD>.xlsx` — 五个工作表：保证金比例 / 交易规则 / 保证金 / 期货手续费 / 期权手续费
- `东方财富期货交易规则保证金<YYYY-MM-DD>.html` — 五表切换 + 搜索的网页版
- `index.html` — 最新一次抓取的快照（用于 GitHub Pages 直接访问）

## 各表字段

| 工作表 | 字段 |
|--------|------|
| 保证金比例 | 交易所 / 品种 / 代码 / 主力合约 / 每日涨跌幅度 / 公司保证金比例 / 特殊合约参数调整 / 备注 |
| 交易规则 | 交易所 / 交易品种 / 最小变动价位 / 交割月份 / 最后交易日 / 最后交割日期 / 交易单位 |
| 保证金 | 品种 / 合约代码 / 保证金(元) / 保证金率 / 备注 |
| 期货手续费 | 交易所 / 品种 / 代码 / 东方财富期货手续费标准 / 计费方式 / 平今·日内加收 |
| 期权手续费 | 交易所 / 品种 / 代码 / 东方财富期权手续费标准 / 计费方式 |

> 手续费表除官网原文标准外，程序另派生「计费方式」（按手收取 / 按成交额）与「平今·日内加收」两个标记列，便于筛选。

## 在线访问

本仓库通过 GitHub Pages 托管，最新快照可直接访问：

```
https://chenzhuanxin.github.io/eastmoney-futures-margin/
```

> 免责声明：数据仅供信息参考，不构成投资建议。保证金与手续费会随行情波动及交易所、期货公司政策调整，实际以中国期货市场监控中心交易结算单及东方财富期货 APP 交易界面为准。
