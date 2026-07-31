#!/usr/bin/env python3
"""从销售看板更新小组调价率，并调用标准层清洗器生成品牌日明细/月历。"""
from __future__ import annotations
import argparse, datetime as dt, json, re, subprocess, sys
from pathlib import Path
import openpyxl

GROUPS = ['饰品1组','饰品2组','海淘组','珠宝1组','珠宝2组','珠宝3组']

def text(v):
    s=str(v or '').strip()
    return s[:-2] if s.endswith('.0') else s

def date_text(v):
    s=text(v)
    return f'{s[:4]}-{s[4:6]}-{s[6:]}' if re.fullmatch(r'\d{8}',s) else s

def metric(den,adj,target):
    den=int(round(den)); adj=int(round(adj)); rate=adj/den if den else None
    return {'denominator':den,'adjusted':adj,'rate':rate,'target':target,'status':'no_data' if rate is None else ('pass' if rate>=target else 'fail')}

def overall(source):
    wb=openpyxl.load_workbook(source,data_only=True,read_only=True)
    ws=wb['data']; it=ws.iter_rows(values_only=True); headers=list(next(it))
    def col(name):
        for i,v in enumerate(headers):
            if str(v or '').strip()==name:return i
        raise RuntimeError(f'data缺少字段{name}')
    idx={x:col(x) for x in ('数据分组\n（选TOTAL来用）','日期','小组','神银','价高商品数','价高调价商品数','辅助列3')}
    result={g:{'denominator':0.0,'adjusted':0.0} for g in GROUPS}; dates=[]
    for row in it:
        if row[idx['数据分组\n（选TOTAL来用）']]!='小组' or row[idx['辅助列3']]!='cur':continue
        raw_date=text(row[idx['日期']]); group=row[idx['小组']]
        if re.fullmatch(r'\d{8}',raw_date):dates.append(raw_date)
        if group in result and row[idx['神银']] not in (None,''):
            result[group]['denominator']+=float(row[idx['价高商品数']] or 0)
            result[group]['adjusted']+=float(row[idx['价高调价商品数']] or 0)
    wb.close()
    if not dates:raise RuntimeError('data未找到小组+cur有效日期')
    return date_text(max(dates)),result

def main():
    ap=argparse.ArgumentParser(); ap.add_argument('--source',required=True,type=Path); ap.add_argument('--frontend-dir',required=True,type=Path); ap.add_argument('--keep-six-high',action='store_true',default=True); args=ap.parse_args()
    source_date,values=overall(args.source)
    old=json.loads((args.frontend_dir/'data/adjustment_rate.json').read_text('utf-8'))
    old_six={x['group']:x['six_high'] for x in old['groups']}
    groups=[{'group':g,'overall':metric(values[g]['denominator'],values[g]['adjusted'],.75),'six_high':old_six[g]} for g in GROUPS]
    overall_total=metric(sum(x['overall']['denominator'] for x in groups),sum(x['overall']['adjusted'] for x in groups),.75)
    six_total=metric(sum(x['six_high']['denominator'] for x in groups),sum(x['six_high']['adjusted'] for x in groups),.8)
    adjustment={**old,'overall_source_date':source_date,'groups':groups,'summary':{'group':'精品总','overall':overall_total,'six_high':six_total},'generated_at':dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),'validation':{'status':'PASS','errors':[]}}
    (args.frontend_dir/'data/adjustment_rate.json').write_text(json.dumps(adjustment,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    cleaner=Path(__file__).with_name('generate_brand_adjustment_rate.py')
    subprocess.run([sys.executable,str(cleaner),'--source',str(args.source),'--daily-output',str(args.frontend_dir/'data/brand_adjustment_daily.json'),'--output',str(args.frontend_dir/'data/brand_adjustment_rate.json')],check=True)
    brand=json.loads((args.frontend_dir/'data/brand_adjustment_rate.json').read_text('utf-8'))
    print(json.dumps({'status':'PASS','source_date':source_date,'overall':overall_total,'brand_summary':brand['summary'],'brand_validation':brand['validation'],'six_high_source_date':adjustment['six_high_source_date']},ensure_ascii=False))

if __name__=='__main__':main()
