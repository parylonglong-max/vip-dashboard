#!/usr/bin/env python3
"""从销售看板生成品牌月度调价率数据。

口径：
- 品牌目录与标签：brand_info，SN 为唯一键；
- 调价率：data 中 数据分组=TOTAL、辅助列3(BT)=cur、当月所有日期，按 SN 汇总 BK/BL；
- 比率 = SUM(价高调价商品数) / SUM(价高商品数)；分母为0时 rate=null。
注意：品牌 TOTAL 与小组汇总覆盖范围不同，不做跨粒度总量强行对平。
"""
from __future__ import annotations
import argparse, datetime as dt, json, re
from collections import defaultdict
from pathlib import Path
import openpyxl


def sn_text(v):
    if v is None: return ''
    s=str(v).strip()
    return s[:-2] if s.endswith('.0') else s

def date_text(v):
    s=sn_text(v)
    if re.fullmatch(r'\d{8}',s): return f'{s[:4]}-{s[4:6]}-{s[6:]}'
    return s

def main():
    ap=argparse.ArgumentParser()
    ap.add_argument('--source',required=True,type=Path)
    ap.add_argument('--output',required=True,type=Path)
    args=ap.parse_args()
    wb=openpyxl.load_workbook(args.source,data_only=True,read_only=True)
    for sheet in ('data','brand_info'):
        if sheet not in wb.sheetnames: raise RuntimeError(f'缺少子表: {sheet}')

    info=wb['brand_info']; ih=list(next(info.iter_rows(min_row=1,max_row=1,values_only=True)))
    def ix(headers,name):
        for i,h in enumerate(headers):
            if str(h or '').strip()==name: return i
        raise RuntimeError(f'缺少字段 {name}: {headers}')
    ii={x:ix(ih,x) for x in ('日期','品牌','sn','等级','小组','神银')}
    catalog={}; meta_conflicts=[]
    for row in info.iter_rows(min_row=2,values_only=True):
        sn=sn_text(row[ii['sn']])
        if not sn or sn in ('(NULL)','NULL'): continue
        item={'sn':sn,'brand':str(row[ii['品牌']] or '').strip(),'level':str(row[ii['等级']] or '').strip(),
              'group':str(row[ii['小组']] or '').strip(),'shenyin':str(row[ii['神银']] or '').strip(),
              'info_date':date_text(row[ii['日期']])}
        if sn in catalog:
            old=catalog[sn]
            if any(old[k]!=item[k] for k in ('brand','level','shenyin')):
                meta_conflicts.append({'sn':sn,'old':old,'new':item})
            # 同品牌跨组时保留多个小组，仅用于扩展，不影响当前标签。
            groups=sorted(set((old.get('groups') or [old.get('group')])+[item['group']]))
            old['groups']=[g for g in groups if g]
        else:
            item['groups']=[item['group']] if item['group'] else []
            catalog[sn]=item
    if meta_conflicts: raise RuntimeError(f'品牌元数据冲突: {meta_conflicts[:5]}')

    data=wb['data']; dh=list(next(data.iter_rows(min_row=1,max_row=1,values_only=True)))
    di={x:ix(dh,x) for x in ('数据分组\n（选TOTAL来用）','日期','sn','品牌','等级','小组','神银','价高商品数','价高调价商品数','辅助列3')}
    rows=list(data.iter_rows(min_row=2,values_only=True))
    valid_dates=[sn_text(r[di['日期']]) for r in rows if r[di['数据分组\n（选TOTAL来用）']]=='TOTAL' and r[di['辅助列3']]=='cur' and re.fullmatch(r'\d{8}',sn_text(r[di['日期']]))]
    if not valid_dates: raise RuntimeError('未找到 TOTAL + cur 的有效品牌日期')
    latest=max(valid_dates); month=latest[:6]
    agg=defaultdict(lambda:{'denominator':0.0,'adjusted':0.0,'row_count':0})
    for r in rows:
        sn=sn_text(r[di['sn']]); d=sn_text(r[di['日期']])
        if r[di['数据分组\n（选TOTAL来用）']]!='TOTAL' or r[di['辅助列3']]!='cur' or not d.startswith(month) or not sn or sn in ('(NULL)','NULL'):
            continue
        agg[sn]['denominator']+=float(r[di['价高商品数']] or 0)
        agg[sn]['adjusted']+=float(r[di['价高调价商品数']] or 0)
        agg[sn]['row_count']+=1
        if sn not in catalog:
            catalog[sn]={'sn':sn,'brand':str(r[di['品牌']] or '').strip(),'level':str(r[di['等级']] or '').strip(),
                         'group':str(r[di['小组']] or '').strip(),'groups':[str(r[di['小组']] or '').strip()],
                         'shenyin':str(r[di['神银']] or '').strip(),'info_date':date_text(latest)}
    brands=[]; errors=[]
    for sn,item in catalog.items():
        a=agg[sn]; den=int(round(a['denominator'])); adj=int(round(a['adjusted']))
        if adj>den: errors.append(f'{sn} 调价商品数大于价高商品数: {adj}>{den}')
        brands.append({**item,'date':date_text(latest),'month':f'{int(month[4:6])}月截至{int(latest[6:])}日',
                       'denominator':den,'adjusted':adj,'rate':adj/den if den else None,'row_count':a['row_count']})
    if errors: raise RuntimeError('; '.join(errors[:10]))
    brands.sort(key=lambda x:(-x['denominator'],x['brand'],x['sn']))
    payload={'title':'品牌调价率','source_date':date_text(latest),'source_month':month,
             'scope':'data分组=TOTAL；辅助列3=cur；本月按品牌SN汇总BK/BL',
             'coverage_note':'品牌TOTAL与小组汇总覆盖范围不同，不做跨粒度总量对平。',
             'brands':brands,'summary':{'catalog_count':len(brands),'metric_brand_count':sum(x['denominator']>0 for x in brands),
             'denominator':sum(x['denominator'] for x in brands),'adjusted':sum(x['adjusted'] for x in brands)},
             'validation':{'status':'PASS','errors':[]},'generated_at':dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
    args.output.parent.mkdir(parents=True,exist_ok=True)
    args.output.write_text(json.dumps(payload,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(payload['summary'],ensure_ascii=False))

if __name__=='__main__': main()
