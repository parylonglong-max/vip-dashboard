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
# 历史月份列位门禁：外网YTD与单月字段不同，绝不能复用同一格式化索引。
data=json.loads((root/'data/excel_view.json').read_text(encoding='utf-8'))
section=next(s for s in data['sections'] if s['id']=='price_index_history')
header=next(r for r in section['rows'] if r['excelRow']==56)
fields=[c['text'] for c in header['cells']]
assert fields[12:23]==['综合得分','天猫价指','对标值','天猫降幅','抖音价指','对标值','抖音降幅','天猫外网加总','抖音外网加总','天猫权重','抖音权重'], '外网1月字段结构变化'
assert fields[23:34]==['综合得分','天猫价指','对标值','天猫降幅','抖音价指','对标值','抖音降幅','天猫外网加总','抖音外网加总','天猫权重','抖音权重'], '外网2月字段结构变化'
assert 'function priceIndexHistoryRows' in app and '天猫外网加总' in app, '缺少外网历史独立字段驱动渲染'
assert "[2,3,5,6,10,11]" in app and "[4,7]" in app, '外网单月比例/pp字段映射缺失'
assert "price_power_history" in app and "最新月份" in app, '五星价格力历史月份仍存在硬编码风险'
print('RENDERING_RULES_GATE PASS')
