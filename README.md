# 东方财富期货 · 交易规则与保证金一览

抓取东方财富期货官网的三个数据源，生成 **Excel + HTML** 双输出（三表切换 + 搜索的美化网页版）。

## 数据源

| 表 | 来源 |
|----|------|
| 保证金比例表 | https://www.eastmoneyfutures.com/pages/service/jyts.html#jyrl |
| 品种及交易规则表 | https://www.eastmoneyfutures.com/pages/service/jygz.html （后端接口 `/emfApi/pzjy/getPZJYInfo`） |
| 品种单手保证金一览表 | https://qhweb.eastmoney.com/bzj/low |

## 使用

```bash
# 安装依赖
pip install requests beautifulsoup4 openpyxl lxml

# 运行（在当前目录生成 东方财富期货交易规则保证金<日期>.xlsx 和 .html）
python fetch_eastmoney_futures.py
```

## 产物

- `东方财富期货交易规则保证金<YYYY-MM-DD>.xlsx` — 三个工作表：保证金比例 / 交易规则 / 保证金
- `东方财富期货交易规则保证金<YYYY-MM-DD>.html` — 三表切换 + 搜索的网页版
- `index.html` — 最新一次抓取的快照（用于 GitHub Pages 直接访问）

## 在线访问

本仓库通过 GitHub Pages 托管，最新快照可直接访问：

```
https://chenzhuanxin.github.io/eastmoney-futures-margin/
```

> 免责声明：数据仅供信息参考，不构成投资建议。保证金会随行情波动，具体以东方财富期货交易界面为准。
