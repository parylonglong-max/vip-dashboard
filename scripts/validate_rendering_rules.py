#!/usr/bin/env python3
"""前端全局渲染规范静态门禁。"""
from pathlib import Path
import json
root=Path(__file__).resolve().parents[1]
app=(root/'app.js').read_text(encoding='utf-8')
css=(root/'styles.css').read_text(encoding='utf-8')
required_app=['normalizeAndSortTotalRows','function pricePctCell','function pricePpCell','function ratioPpCell',"cell.trend=n>0?'up':n<0?'down':null"]
for s in required_app:
    assert s in app, f'app.js 缺少规范实现: {s}'
for s in ['.excel-cell.trend-up','.excel-cell.trend-down','禁止按固定行号使用黄色']:
    assert s in css, f'styles.css 缺少规范实现: {s}'
assert '#fff7cc' not in css, '禁止遗留黄色整格高亮'
for p in [root/'data/excel_view.json',root/'data/adjustment_rate.json']:
    assert p.exists(), f'缺少数据文件: {p}'
print('RENDERING_RULES_GATE PASS')
