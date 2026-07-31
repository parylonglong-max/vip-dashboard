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
# 品牌视角门禁：SN唯一、原始计数可重算、无分母不得判定未达标。
brand_path=root/'data/brand_adjustment_rate.json'
assert brand_path.exists(), '缺少品牌调价率数据'
brand=json.loads(brand_path.read_text(encoding='utf-8'))
assert brand['validation']['status']=='PASS'
sns=[x['sn'] for x in brand['brands']]
assert len(sns)==len(set(sns)), '品牌SN不唯一'
for x in brand['brands']:
    assert x['adjusted']<=x['denominator'], f"{x['sn']} 调价商品数大于价高商品数"
    if x['denominator']==0: assert x['rate'] is None, f"{x['sn']} 无分母时rate必须为null"
    else: assert abs(x['rate']-x['adjusted']/x['denominator'])<1e-12, f"{x['sn']} 调价率不可重算"
for token in ['viewMode','品牌视角','renderBrandAdjustmentPanel','renderBrandCalendar','brandSearchInput','输入品牌名称或品牌SN','本月价高数','调价日历']:
    assert token in app, f'品牌视角实现缺失: {token}'
# 品牌日历门禁：当月每日数据完整，月汇总必须等于日汇总；未来日期不得伪造为0。
for x in brand['brands']:
    assert x.get('daily') and len(x['daily']) in (28,29,30,31), f"{x['sn']} 缺少完整日历"
    assert x['denominator']==sum(z['denominator'] for z in x['daily']), f"{x['sn']} 月价高数不等于日汇总"
    assert x['adjusted']==sum(z['adjusted'] for z in x['daily']), f"{x['sn']} 月调价数不等于日汇总"
    assert all(z['adjusted']<=z['denominator'] for z in x['daily']), f"{x['sn']} 日调价数大于价高数"
    assert all(z['rate'] is None for z in x['daily'] if not z['has_data']), f"{x['sn']} 无明细日期不得显示0%"
print('RENDERING_RULES_GATE PASS')
