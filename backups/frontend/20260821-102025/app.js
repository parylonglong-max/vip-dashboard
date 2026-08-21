/* 精品货价监控数据看板 — Mobile BI / No Pie, No Ring */
(function () {
  "use strict";

  var API_BASE = window.DASHBOARD_API_BASE || "http://127.0.0.1:8900";
  var FALLBACK_URL = "data/excel_view.json";
  var STATIC_PREVIEW_PASSWORD = "vip2026";
  var PERIODS = ["YTD", "1月", "2月", "3月", "4月", "5月", "6月", "7月"];

  var state = { token: null, data: null, activeTab: "sales", viewMode: "group", activeBrandTab: "brand-adjustment", selectedBrandSn: null, brandQuery: "", periods: {} };

  var $loginPage = document.getElementById("loginPage");
  var $dashboard = document.getElementById("dashboard");
  var $passwordInput = document.getElementById("passwordInput");
  var $loginBtn = document.getElementById("loginBtn");
  var $loginError = document.getElementById("loginError");
  var $logoutBtn = document.getElementById("logoutBtn");
  var $modulesContainer = document.getElementById("modulesContainer");
  var $navbarDate = document.getElementById("navbarDate");
  var $periodToggle = document.querySelector(".period-toggle");

  var TABS = [
    { id: "sales", label: "自营销售", sectionIds: ["self_sales_mtd", "self_sales_history"] },
    { id: "brand-tier", label: "品牌分层", sectionIds: [] },
    { id: "gross", label: "毛利", sectionIds: ["gross_profit"] },
    { id: "price", label: "外网价指", sectionIds: ["price_index_mtd", "price_index_history", "six_high_price_index"] },
    { id: "discount", label: "内网折扣", sectionIds: ["internal_discount"] },
    { id: "liugao", label: "六高", sectionIds: ["six_high"] },
    { id: "adjustment", label: "调价率", sectionIds: [] },
    { id: "quality", label: "优质款", sectionIds: ["quality_product_mtd", "quality_product_history"] },
    { id: "machine", label: "机采", sectionIds: ["machine_purchase_mtd", "machine_purchase_history"] },
    { id: "power", label: "五星价格力", sectionIds: ["price_power_mtd", "price_power_history"] },
    { id: "traffic", label: "流量趋势", sectionIds: ["traffic"] },
  ];
  var BRAND_TABS = [
    { id: "brand-adjustment", label: "调价率" },
    { id: "brand-price-index", label: "外网价指" },
    { id: "brand-sales-traffic", label: "品牌销售流量" },
  ];

  var PERIOD_CONFIG = {
    self_sales_history: { rowStart: 17, rowEnd: 23, periods: { "YTD": [0,1,2,3,4,5], "1月": [0,6,7,8,9,10], "2月": [0,11,12,13,14,15], "3月": [0,16,17,18,19,20], "4月": [0,21,22,23,24,25], "5月": [0,26,27,28,29,30], "6月": [0,31,32,33,34,35], "7月": [0,36,37,38,39,40] }, headers: function(p){ return p==="YTD" ? ["小组","YTD目标","YTD完成","YTD完成率","同期","业绩同比"] : ["小组",p+"目标",p+"完成",p+"完成率","同期","业绩同比"]; } },
    price_index_history: { rowStart: 56, rowEnd: 63, periods: { "YTD": [0,1,2,3,4,5,6,7,8,9,10,11], "1月": [0,12,13,14,15,16,17,18,19,20,21,22], "2月": [0,23,24,25,26,27,28,29,30,31,32,33], "3月": [0,34,35,36,37,38,39,40,41,42,43,44], "4月": [0,45,46,47,48,49,50,51,52,53,54,55], "5月": [0,56,57,58,59,60,61,62,63,64,65,66], "6月": [0,67,68,69,70,71,72,73,74,75,76,77], "7月": "MTD_SNAPSHOT" }, headers: function(p){ return ["小组",p+"综合得分",p+"天猫得分",p+"抖音得分",p+"天猫权重",p+"抖音权重",p+"天猫价指",p+"天猫目标",p+"天猫差值",p+"抖音价指",p+"抖音目标",p+"抖音差值"]; } },
    internal_discount: { rowStart: 68, rowEnd: 74, periods: { "MTD": [0,1,2,3], "YTD": [0,4,5,6], "1月": [0,7,8,9], "2月": [0,10,11,12], "3月": [0,13,14,15], "4月": [0,16,17,18], "5月": [0,19,20,21], "6月": [0,22,23,24], "7月": "MTD_SNAPSHOT" }, headers: function(p){ return ["小组",p+"去年内网价指",p+"今年内网价指",p+"系数差"]; }, periodList: ["MTD","YTD","1月","2月","3月","4月","5月","6月","7月"] },
    quality_product_history: { rowStart: 103, rowEnd: 109, periods: { "YTD": [0,1,2,3,4,5], "1月": [0,6,7,8,9,10], "2月": [0,11,12,13,14,15], "3月": [0,16,17,18,19,20], "4月": [0,21,22,23,24,25], "5月": [0,26,27,28,29,30], "6月": [0,31,32,33,34,35], "7月": "MTD_SNAPSHOT" }, headers: function(p){ return ["小组",p+"已引进",p+"未引进",p+"暂不引进",p+"总计",p+"引入率"]; } },
    machine_purchase_history: { rowStart: 123, rowEnd: 126, periods: { "YTD": [0,1,2,3,4,5], "1月": [0,6,7,8,9,10], "2月": [0,11,12,13,14,15], "3月": [0,16,17,18,19,20], "4月": [0,21,22,23,24,25], "5月": [0,26,27,28,29,30], "6月": [0,31,32,33,34,35], "7月": "MTD_SNAPSHOT" }, headers: function(p){ return p==="YTD" ? ["小组","YTD目标","YTD完成","YTD完成率","同期","业绩同比"] : ["小组",p+"目标",p+"完成",p+"完成率","同期","业绩同比"]; } },
    price_power_history: { rowStart: 139, rowEnd: 152, periods: { "YTD": null, "最新月份": null }, headers: function(p){ return ["时间","维度","指标","曝光","曝光占比","APP销售","APP销售占比","实际","目标","VS目标差距","完成率"]; }, periodList: ["YTD","最新月份","7月"] }
  };

  function apiFetch(path, options) {
    options = options || {}; options.headers = options.headers || {};
    if (state.token && state.token !== "static-preview") options.headers.Authorization = "Bearer " + state.token;
    return fetch(API_BASE + path, options).then(function (res) { if (!res.ok) return res.text().then(function (text) { throw new Error("HTTP " + res.status + " " + text); }); return res.json(); });
  }
  function loginByApi(password) { return apiFetch("/api/login", { method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify({password:password}) }).then(function(json){ state.token=json.token; return json; }); }
  function dataUrl(path){ return path+'?v='+encodeURIComponent(window.__DASHBOARD_RELEASE__||'202608211020'); }
  function loadExtraData(){ return Promise.all([fetch(dataUrl('data/traffic_uv.json')).then(function(r){return r.json();}).then(function(d){state.trafficData=d;}).catch(function(){}),fetch(dataUrl('data/traffic_flow.json')).then(function(r){return r.json();}).then(function(d){state.trafficFlowData=d;}).catch(function(){}),fetch(dataUrl('data/adjustment_rate.json')).then(function(r){return r.json();}).then(function(d){state.adjustmentData=d;}).catch(function(){}),fetch(dataUrl('data/brand_adjustment_rate.json')).then(function(r){return r.json();}).then(function(d){state.brandAdjustmentData=d;}).catch(function(){}),fetch(dataUrl('data/problem_brands.json')).then(function(r){return r.json();}).then(function(d){state.problemBrandsData=d;}).catch(function(){}),fetch(dataUrl('data/brand_price_index.json')).then(function(r){return r.json();}).then(function(d){state.brandPriceIndexData=d;}).catch(function(){}),fetch(dataUrl('data/brand_tier_mtd.json')).then(function(r){return r.json();}).then(function(d){state.brandTierData=d;}).catch(function(){})]); }
  function loadData() { return apiFetch("/api/excel_view").then(function(json){ state.data=json.data||json; return loadExtraData().then(function(){renderDashboard();}); }).catch(function(){ return fetch(FALLBACK_URL).then(function(res){ if(!res.ok) throw new Error("HTTP "+res.status); return res.json(); }).then(function(json){ state.data=json; return loadExtraData().then(function(){renderDashboard();}); }).catch(function(){ $modulesContainer.innerHTML='<div class="loading">数据加载失败，请稍后重试</div>'; }); }); }

  function escapeHtml(value){ return String(value==null?"":value).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"); }
  function parseNum(cell){ if(!cell) return null; var v=cell.raw; if(typeof v==="number") return v; var n=Number(v); return isNaN(n)?null:n; }
  function textWithUnit(cell){ if(!cell||!cell.text) return "—"; return cell.text + (cell.unit && cell.unit !== "%" ? cell.unit : ""); }
  function clamp(v,min,max){ return Math.max(min,Math.min(max,v)); }
  function getSection(id){ return (state.data.sections||[]).find(function(s){ return s.id===id; }); }
  // 全局渲染规范：仅“精品总”可使用总计样式，且必须置于所属表格最后。
  function normalizeGroupName(name){ return /^(总|精品|精品总计|精品总)$/.test(String(name||'').trim()) ? '精品总' : name; }
  function isTotalRow(row){ return !!(row&&(row.total||(row.cells&&row.cells[0]&&normalizeGroupName(row.cells[0].text)==='精品总'))); }
  function normalizeAndSortTotalRows(rows){
    var headers=[], data=[];
    (rows||[]).forEach(function(row){
      var cloned=Object.assign({},row,{cells:(row.cells||[]).map(function(cell){return Object.assign({},cell);})});
      var first=cloned.cells[0];
      if(first && !first.header && !(first.merge&&first.merge.covered) && normalizeGroupName(first.text)==='精品总'){
        first.text='精品总'; first.raw='精品总'; cloned.total=true;
      }
      if((cloned.cells||[]).some(function(cell){return cell&&cell.header;})) headers.push(cloned); else data.push(cloned);
    });
    return headers.concat(data.filter(function(row){return !isTotalRow(row);}),data.filter(isTotalRow));
  }
  function activePeriod(sectionId){ return state.periods[sectionId] || (PERIOD_CONFIG[sectionId] && PERIOD_CONFIG[sectionId].periodList ? PERIOD_CONFIG[sectionId].periodList[0] : "YTD"); }

  function salesMtdRecords(){ var section=getSection("self_sales_mtd"); if(!section) return []; return section.rows.filter(function(row){return row.excelRow>=6&&row.excelRow<=12;}).map(function(row){return {group:row.cells[0].text, actual:parseNum(row.cells[2]), deptRate:parseNum(row.cells[7]), groupRate:parseNum(row.cells[9]), row:row};}); }

  function renderSalesSummary(){
    var records=salesMtdRecords(); var total=records.find(function(r){return r.group==="总";})||records[records.length-1]; if(!total) return "";
    var groups=records.filter(function(r){return r.group!=="总";});
    var html='<section class="sales-hero neu-card no-circle"><div class="sales-hero-title"><span></span><b>自营销售达成率总览</b></div>';
    html+='<div class="summary-progress-grid">';
    html+=progressCard("部门阶段目标完成率", total.deptRate, textWithUnit(total.row.cells[2]), textWithUnit(total.row.cells[6]));
    html+=progressCard("小组阶段目标完成率", total.groupRate, textWithUnit(total.row.cells[2]), textWithUnit(total.row.cells[8]));
    html+='</div><div class="group-progress-list">';
    groups.forEach(function(r){ html+=progressCard(r.group, r.groupRate, textWithUnit(r.row.cells[2]), textWithUnit(r.row.cells[8])); });
    html+='</div></section>';
    return html;
  }
  function progressCard(label, rate, actual, target){ var pct=clamp((rate||0)*100,0,120); var cls=(rate>=1?'good':rate<0.8?'bad':'normal'); return '<div class="rate-card '+cls+'"><div class="rate-card__head"><b>'+escapeHtml(label)+'</b><span>'+((rate==null||isNaN(rate))?'—':(rate*100).toFixed(1)+'%')+'</span></div><div class="rate-card__bar"><i style="width:'+pct+'%"></i></div><div class="rate-card__money">'+escapeHtml(actual)+' / '+escapeHtml(target)+'</div></div>'; }

  function cellClass(cell,rowIndex,colIndex){ var classes=["excel-cell"]; var ctx=(cell&&cell.context||"")+" "+(cell&&cell.text||""); if(!cell||!cell.text) classes.push("is-blank"); if(cell&&cell.type==="number") classes.push("is-number"); if(rowIndex===0 || (cell&&cell.header)) classes.push("is-header"); if(colIndex===0) classes.push("is-row-label"); if(cell&&cell.merge&&cell.merge.colspan>1) classes.push("is-merged-head"); if(cell&&cell.sectionDivider) classes.push("section-divider"); if(cell&&cell.trend==="up") classes.push("trend-up"); if(cell&&cell.trend==="down") classes.push("trend-down"); if(cell&&cell.qualityFail) classes.push("quality-fail"); if(/同比/.test(ctx)) classes.push("is-yoy"); if(/完成率|达成率|阶段完成率|引入率|占比/.test(ctx)) classes.push("is-progress"); return classes.join(" "); }
  function renderEnhancedContent(cell){ if(!cell||!cell.text) return '<span class="blank-placeholder">—</span>'; var ctx=cell.context||""; var raw=parseNum(cell); var text=escapeHtml(cell.text); var unit=cell.unit&&cell.unit!=="%"?'<small>'+escapeHtml(cell.unit)+'</small>':""; if(/同比/.test(ctx)&&raw!=null){ var status=raw>=0?"up":"down"; return '<span class="yoy-pill '+status+'">'+(raw>=0?'↑':'↓')+' '+text+'</span>'; } if(/完成率|达成率|阶段完成率|引入率|占比/.test(ctx)&&raw!=null&&cell.unit==="%"){ var pct=raw>2?raw:raw*100; var cls=pct>=100?"good":pct<80?"bad":"normal"; var chip=cell.qualityFail?'<em class="status-chip">未达标</em>':''; return '<div class="progress-cell '+cls+'"><span>'+text+chip+'</span><div class="mini-progress"><i style="width:'+clamp(pct,0,120)+'%"></i></div></div>'; } return text+unit; }
  function renderCell(cell,rowIndex,colIndex){ if(cell&&cell.merge&&cell.merge.covered) return ''; var title=cell&&cell.coord&&cell.text?' title="'+escapeHtml(cell.coord+' '+(cell.raw==null?'':cell.raw))+'"':""; var span=''; if(cell&&cell.merge&&cell.merge.rowspan>1) span+=' rowspan="'+cell.merge.rowspan+'"'; if(cell&&cell.merge&&cell.merge.colspan>1) span+=' colspan="'+cell.merge.colspan+'"'; return '<td class="'+cellClass(cell,rowIndex,colIndex)+'"'+title+span+'>'+renderEnhancedContent(cell)+'</td>'; }

  function markHeaders(rows, count){ return (rows||[]).map(function(row,idx){ if(idx<count){ row={excelRow:row.excelRow,cells:(row.cells||[]).map(function(cell){ var c=Object.assign({},cell); c.header=true; return c; })}; } return row; }); }
  function headerRow(names){ return {excelRow:0,cells:names.map(function(name){return {text:name,raw:name,type:"text",header:true};})}; }
  function renderRows(rows, tableClass){ var normalized=normalizeAndSortTotalRows(rows); var html='<div class="excel-scroll"><table class="excel-table '+(tableClass||'')+'">'; normalized.forEach(function(row,rowIndex){ var total=isTotalRow(row)?" total-row":""; html+='<tr class="'+total+'" data-excel-row="'+row.excelRow+'">'; (row.cells||[]).forEach(function(cell,colIndex){html+=renderCell(cell,rowIndex,colIndex);}); html+='</tr>'; }); html+='</table></div>'; return html; }

  function salesMoneyDisplayCell(cell){
    if(!cell||typeof cell.raw!=="number") return cell;
    var n=cell.raw, abs=Math.abs(n);
    if(abs>=1e8){cell.text=(n/1e8).toFixed(2);cell.unit='亿';}
    else if(abs>=1e4){cell.text=(n/1e4).toFixed(1);cell.unit='万';}
    return cell;
  }
  function salesMtdTableSection(section){
    var rows=(section.rows||[]).filter(function(row){return row.excelRow>=5&&row.excelRow<=12;}).map(function(row){
      var cells=cloneCells(row.cells).slice(0,10);
      [1,2,4,6,8].forEach(function(i){salesMoneyDisplayCell(cells[i]);});
      return {excelRow:row.excelRow,cells:cells};
    });
    rows=markHeaders(rows,1);
    return '<div class="section-title"><span></span>自营销售 · MTD</div>'+renderRows(rows);
  }

  function blankCell(){ return {text:"",raw:null,type:"blank"}; }
  function cloneCells(cells){ return (cells||[]).map(function(c){return Object.assign({},c||blankCell());}); }
  function normalizeTotalCell(cell){ if(cell&&cell.text==='总'){cell.text='精品总';cell.raw='精品总';} return cell; }
  function historyMtdSnapshotRows(sectionId){
    if(sectionId==='self_sales_history'){
      var mtd=getSection('self_sales_mtd');
      var rows=(mtd&&mtd.rows||[]).filter(function(row){return row.excelRow>=6&&row.excelRow<=12;}).map(function(row){
        var cells=cloneCells(row.cells).slice(0,6); normalizeTotalCell(cells[0]); [1,2,4].forEach(function(i){salesMoneyDisplayCell(cells[i]);}); ratePctCell(cells[3]); growthPctCell(cells[5]); return {excelRow:row.excelRow,cells:cells};
      });
      return [headerRow(['小组','7月目标','7月完成','7月完成率','同期','业绩同比'])].concat(rows);
    }
    if(sectionId==='quality_product_history'){
      var q=getSection('quality_product_mtd');
      var qrows=(q&&q.rows||[]).filter(function(row){return row.excelRow>=93&&row.excelRow<=99;}).map(function(row){
        var cells=cloneCells(row.cells).slice(0,6); normalizeTotalCell(cells[0]); ratePctCell(cells[5]); return {excelRow:row.excelRow,cells:cells,total:normalizeGroupName(cells[0]&&cells[0].text)==='精品总'};
      });
      return [headerRow(['小组','7月已引进','7月未引进','7月暂不引进','7月总计','7月引入率'])].concat(qrows);
    }
    if(sectionId==='machine_purchase_history'){
      var m=getSection('machine_purchase_mtd');
      var mrows=(m&&m.rows||[]).filter(function(row){return row.excelRow>=116&&row.excelRow<=119;}).map(function(row){
        var cells=cloneCells(row.cells).slice(0,6); normalizeTotalCell(cells[0]); ratePctCell(cells[3]); growthPctCell(cells[5]); return {excelRow:row.excelRow,cells:cells,total:normalizeGroupName(cells[0]&&cells[0].text)==='精品总'};
      });
      return [headerRow(['小组','7月目标','7月完成','7月完成率','同期','业绩同比'])].concat(mrows);
    }
    return [];
  }
  function periodRows(section, sectionId){ var cfg=PERIOD_CONFIG[sectionId]; var period=activePeriod(sectionId); if(!cfg) return section.rows||[]; if(period==='7月'&&cfg.periods&&cfg.periods['7月']==='MTD_SNAPSHOT') return historyMtdSnapshotRows(sectionId); if(sectionId==='price_index_history') return priceIndexHistoryRows(section, period); if(sectionId==='price_power_history'){ return pricePowerPeriodRows(section, period); } var idx=cfg.periods[period]||cfg.periods.YTD; var startHeader=Math.max(0,cfg.rowStart-2); var rows=(section.rows||[]).filter(function(row){return row.excelRow>=startHeader&&row.excelRow<=cfg.rowEnd;}).map(function(row){var cells=idx.map(function(i){return Object.assign({},row.cells[i]||{text:"",type:"blank"});}); if((sectionId==='self_sales_history'||sectionId==='machine_purchase_history')&&row.excelRow>=cfg.rowStart){ratePctCell(cells[3]); growthPctCell(cells[5]);} if(sectionId==='quality_product_history'&&row.excelRow>=cfg.rowStart){ratePctCell(cells[5]);} return {excelRow:row.excelRow,cells:cells};}); return markHeaders(rows,2); }
  // 外网价指历史：YTD(11业务列)与单月(11业务列，但字段语义不同)必须分别切片，禁止用索引硬套。
  function priceIndexHistoryRows(section, period){
    if(period==='7月'){
      var mtd=getSection('price_index_mtd');
      var fields7=['小组','7月综合得分','7月天猫得分','7月抖音得分','7月天猫权重','7月抖音权重','7月天猫价指','7月天猫目标','7月天猫差值','7月抖音价指','7月抖音目标','7月抖音差值'];
      var rows7=(mtd&&mtd.rows||[]).filter(function(row){return row.excelRow>=46&&row.excelRow<=52;}).map(function(row){
        var keep=[0,1,2,3,4,5,6,7,8,10,11,12];
        var cells=keep.map(function(i){return Object.assign({},row.cells[i]||{text:'',type:'blank'});});
        if(cells[0]&&normalizeGroupName(cells[0].text)==='精品总'){cells[0].text='精品总';cells[0].raw='精品总';}
        [6,7,9,10].forEach(function(i){pricePctCell(cells[i]);}); [8,11].forEach(function(i){pricePpCell(cells[i]);});
        return {excelRow:row.excelRow,cells:cells};
      });
      return [headerRow(fields7)].concat(rows7);
    }
    var isYtd=period==='YTD';
    var start=isYtd?1:({"1月":12,"2月":23,"3月":34,"4月":45,"5月":56,"6月":67}[period]);
    var fields=isYtd?
      ['小组','综合得分','天猫得分','抖音得分','天猫权重','抖音权重','天猫价指','天猫目标','天猫降幅','抖音价指','抖音目标','抖音降幅']:
      ['小组','综合得分','天猫价指','对标值','天猫降幅','抖音价指','对标值','抖音降幅','天猫外网加总','抖音外网加总','天猫权重','抖音权重'];
    var rows=(section.rows||[]).filter(function(row){return row.excelRow>=57&&row.excelRow<=63;}).map(function(row){
      var cells=[Object.assign({},row.cells[0]||{text:'',type:'blank'})];
      for(var i=0;i<11;i++) cells.push(Object.assign({},row.cells[start+i]||{text:'',type:'blank'}));
      if(cells[0]&&normalizeGroupName(cells[0].text)==='精品总'){cells[0].text='精品总';cells[0].raw='精品总';}
      // 价指/对标是比例；降幅是源表已计算的 pp；外网加总是数量，必须保持数量格式。
      if(isYtd){ [6,7,9,10].forEach(function(i){pricePctCell(cells[i]);}); [8,11].forEach(function(i){pricePpCell(cells[i]);}); }
      else { [2,3,5,6,10,11].forEach(function(i){pricePctCell(cells[i]);}); [4,7].forEach(function(i){pricePpCell(cells[i]);}); }
      return {excelRow:row.excelRow,cells:cells};
    });
    return [headerRow(fields)].concat(rows);
  }
  function qtyDisplayCell(c){
    if(!c||typeof c.raw!=='number')return;
    var n=c.raw;
    if(Math.abs(n)>=1e8){c.text=(n/1e8).toFixed(2);c.unit='亿';}
    else {c.text=(n/1e4).toFixed(0);c.unit='万';}
  }
  function pricePowerPeriodRows(section, period){
    var headers=["时间","维度","曝光","曝光占比","APP销售","APP销售占比","实际","目标","VS目标差距","完成率"];
    if(period==='7月'){
      var mtd=getSection('price_power_mtd');
      var data7=(mtd&&mtd.rows||[]).filter(function(row){var label=row.cells&&row.cells[0]&&row.cells[0].text;return label==='四五星'||label==='大爆款';}).slice(0,2).map(function(row){
        var src=cloneCells(row.cells);
        var cells=[{text:'7月',raw:'202607',type:'text'},src[0]||blankCell(),src[7]||blankCell(),src[1]||blankCell(),src[8]||blankCell(),src[2]||blankCell(),src[3]||blankCell(),src[4]||blankCell(),src[5]||blankCell(),src[6]||blankCell()];
        return {excelRow:row.excelRow,cells:cells};
      });
      data7.forEach(function(row){ qtyDisplayCell(row.cells[2]); qtyDisplayCell(row.cells[4]); [3,5,6,7,9].forEach(function(i){ratePctCell(row.cells[i]);}); ppRatioCell(row.cells[8]); });
      return [headerRow(headers)].concat(data7);
    }
    // 仅展示源表已有实际数据的最新月份，禁止硬编码月份造成跨月后取错列。
    var monthRows=(section.rows||[]).filter(function(row){var c=row.cells&&row.cells[0]; return c&&/^20\d{4}$/.test(String(c.raw||''))&&row.cells[2]&&typeof row.cells[2].raw==='number'&&row.cells[6]&&typeof row.cells[6].raw==='number';});
    var latest=monthRows.map(function(row){return String(row.cells[0].raw);}).sort().pop();
    var wanted=period==='YTD'?'YTD':latest;
    var data=(section.rows||[]).filter(function(row){
      var first=row.cells&&row.cells[0];
      var raw=first&&first.raw;
      var text=first&&first.text;
      return wanted==='YTD'?text==='YTD':String(raw)===String(wanted);
    }).map(function(row){return {excelRow:row.excelRow,cells:row.cells.slice(0,10).map(function(c){return Object.assign({},c);})};});
    data.forEach(function(row){
      if(wanted!=='YTD'&&row.cells[0]&&row.cells[0].merge&&!row.cells[0].merge.covered){row.cells[0].text=wanted ? (wanted.slice(4,6).replace(/^0/,'')+'月') : '—';row.cells[0].raw=wanted||'';row.cells[0].unit='';}
      qtyDisplayCell(row.cells[2]); qtyDisplayCell(row.cells[4]);
      // 曝光/销售占比、实际/目标/完成率均为比例；VS目标差距为百分点差。
      [3,5,6,7,9].forEach(function(i){ratePctCell(row.cells[i]);});
      ppRatioCell(row.cells[8]);
    });
    return [headerRow(headers)].concat(data);
  }
  function renderPeriodFilter(sectionId){ var cfg=PERIOD_CONFIG[sectionId]; if(!cfg) return ""; var list=cfg.periodList||PERIODS; var current=activePeriod(sectionId); var html='<div class="filterbar">'; list.forEach(function(p){html+='<button class="filter-btn '+(current===p?'active':'')+'" data-section="'+sectionId+'" data-period="'+p+'">'+p+'</button>';}); html+='</div>'; return html; }

  function cloneCell(text){ return {text:text,raw:text,type:"text",unit:"",header:true}; }
  // 价格指数统一展示：数值转百分比并保留 1 位小数；差值统一 pp。
  function pricePctCell(cell){ if(!cell||typeof cell.raw!=="number") return cell; cell.text=(cell.raw*100).toFixed(1)+"%"; cell.unit=""; return cell; }
  // 通用比例与同比：源值为比例，统一保留1位百分比；同比保留正负号和颜色。
  function ratePctCell(cell){ if(!cell||typeof cell.raw!=="number") return cell; cell.text=(cell.raw*100).toFixed(1)+"%"; cell.unit=""; return cell; }
  function growthPctCell(cell){ if(!cell||typeof cell.raw!=="number") return cell; var n=cell.raw; cell.text=(n>0?'+':n<0?'-':'')+Math.abs(n*100).toFixed(1)+'%'; cell.unit=''; cell.trend=n>0?'up':n<0?'down':null; return cell; }
  function ppRatioCell(cell){ if(!cell||typeof cell.raw!=="number") return cell; var n=cell.raw; cell.text=(n>0?'+':n<0?'-':'')+Math.abs(n*100).toFixed(1)+'pp'; cell.unit=''; cell.trend=n>0?'up':n<0?'down':null; return cell; }
  // 外网价指的降幅/差值源值已是 pp；仅保留一位小数。
  function pricePpCell(cell){ if(!cell||typeof cell.raw!=="number") return cell; var n=cell.raw; cell.text=(n>0?"+":n<0?"-":"")+Math.abs(n).toFixed(1); cell.unit="pp"; cell.trend=n>0?'up':n<0?'down':null; return cell; }
  // 内网系数差源值为比例，展示时转换为 pp。
  function ratioPpCell(cell){ if(!cell||typeof cell.raw!=="number") return cell; var n=cell.raw; cell.text=(n>0?"+":n<0?"-":"")+Math.abs(n*100).toFixed(1); cell.unit="pp"; cell.trend=n>0?'up':n<0?'down':null; return cell; }
  function renderPriceIndexMtd(section){
    // 保留全部14列：保留天猫/抖音的“本月目标”，同时展示“完成差值”。
    // 一级表头天猫/抖音分区各4列（价指、对标值、本月目标、完成差值）。
    var keep=[0,1,2,3,4,5,6,7,8,9,10,11,12,13];
    var rows=(section.rows||[]).filter(function(row){return row.excelRow>=43&&row.excelRow<=52;}).map(function(row){
      var cells=keep.map(function(srcIdx){
        var c=Object.assign({},row.cells[srcIdx]);
        if(row.excelRow===44&&(srcIdx===6||srcIdx===10)&&c.merge){ c.merge=Object.assign({},c.merge,{colspan:4}); }
        return c;
      });
      if(row.excelRow===44&&(cells[1]||cells[6]||cells[10])){
        if(cells[1]) cells[1].sectionDivider=true;
        if(cells[6]) cells[6].sectionDivider=true;
        if(cells[10]) cells[10].sectionDivider=true;
      }
      if(cells[0]&&cells[0].text==='总'){cells[0].text='精品总';cells[0].raw='精品总';}
      // 价格指数、对标值、本月目标：百分比展示；完成差值：pp 展示。
      if(row.excelRow>=46){ [6,7,8,10,11,12].forEach(function(i){pricePctCell(cells[i]);}); [9,13].forEach(function(i){pricePpCell(cells[i]);}); }
      return {excelRow:row.excelRow,cells:cells};
    });
    rows=markHeaders(rows,2);
    return '<div class="section-title"><span></span>外网价指 · MTD</div>'+renderRows(rows,'price-index-mtd-grid');
  }
  function renderSixHighMtd(section){
    var source=getSection('six_high_price_index'); var d=source&&source.data; if(!d)return '';
    function numCell(v){if(v==='(NULL)'||v==null)return {text:'',raw:null,type:'blank'};var n=parseFloat(v);return isNaN(n)?{text:String(v),raw:v,type:'text'}:{text:n.toLocaleString('zh-CN',{maximumFractionDigits:1}),raw:n,type:'number'};}
    function pctCell(v,diff){if(v==='(NULL)'||v==null)return {text:'',raw:null,type:'blank'};var n=parseFloat(v);if(isNaN(n))return {text:String(v),raw:v,type:'text'};return {text:diff?((n>0?'+':n<0?'-':'')+Math.abs(n*100).toFixed(1)+'pp'):(n*100).toFixed(1)+'%',raw:n,type:'number',trend:diff?(n>0?'up':n<0?'down':null):null};}
    function dataRow(g,total){return {excelRow:total?1099:1000,total:!!total,cells:[{text:total?'精品总':g.group,raw:g.group,type:'text'},numCell(g['日均商品数']),numCell(g['可比商品数']),pctCell(g['可比率']),pctCell(g['天猫价格指数']),pctCell(g['天猫价格指数目标']),pctCell(g['天猫差值'],true),pctCell(g['抖音价格指数']),pctCell(g['抖音价格指数目标']),pctCell(g['抖音差值'],true),pctCell(g['调价率']),pctCell(g['调价率目标']),pctCell(g['调价率差值'],true),pctCell(g['断货率']),pctCell(g['断货率目标']),pctCell(g['断货率差值'],true),pctCell(g['7天缺货率'])]};}
    var headers=['小组','日均商品数','可比商品数','可比率','天猫价格指数','天猫目标','天猫差值','抖音价格指数','抖音目标','抖音差值','调价率','调价率目标','调价率差值','断货率','断货率目标','断货率差值','7天缺货率'];
    var rows=[headerRow(headers)]; (d.groups||[]).forEach(function(g){rows.push(dataRow(g,false));}); if(d.summary)rows.push(dataRow(d.summary,true));
    return '<div class="section-title"><span></span>六高 · MTD</div>'+renderRows(rows,'six-high-detail-grid');
  }
  function renderSixHighPriceIndex(section){
    var d=section.data; if(!d) return '';
    var groups=d.groups||[]; var summary=d.summary||{};
    function pctCell(v,diff){ if(v==='(NULL)'||v==null)return {text:'',raw:null,type:'blank'}; var n=parseFloat(v); if(isNaN(n))return {text:String(v),raw:v,type:'text'}; return {text:diff?((n>0?'+':n<0?'-':'')+Math.abs(n*100).toFixed(1)):(n*100).toFixed(1)+'%',raw:n,type:'number',unit:diff?'pp':'%',trend:diff?(n>0?'up':n<0?'down':null):null}; }
    var top={excelRow:0,cells:[{text:'小组',raw:'小组',type:'text',header:true,merge:{rowspan:2,colspan:1}},{text:'天猫',raw:'天猫',type:'text',header:true,sectionDivider:true,merge:{rowspan:1,colspan:3}},{merge:{covered:true}},{merge:{covered:true}},{text:'抖音',raw:'抖音',type:'text',header:true,sectionDivider:true,merge:{rowspan:1,colspan:3}},{merge:{covered:true}},{merge:{covered:true}}]};
    var sub={excelRow:0,cells:[{merge:{covered:true}},{text:'价格指数',raw:'价格指数',type:'text',header:true},{text:'目标',raw:'目标',type:'text',header:true},{text:'差值',raw:'差值',type:'text',header:true},{text:'价格指数',raw:'价格指数',type:'text',header:true},{text:'目标',raw:'目标',type:'text',header:true},{text:'差值',raw:'差值',type:'text',header:true}]};
    var rows=[top,sub];
    groups.forEach(function(g,idx){rows.push({excelRow:1000+idx,cells:[{text:g.group,raw:g.group,type:'text'},pctCell(g['天猫价格指数']),pctCell(g['天猫价格指数目标']),pctCell(g['天猫差值'],true),pctCell(g['抖音价格指数']),pctCell(g['抖音价格指数目标']),pctCell(g['抖音差值'],true)]});});
    if(summary&&summary.group){rows.push({excelRow:1099,total:true,cells:[{text:'精品总',raw:'精品总',type:'text'},pctCell(summary['天猫价格指数']),pctCell(summary['天猫价格指数目标']),pctCell(summary['天猫差值'],true),pctCell(summary['抖音价格指数']),pctCell(summary['抖音价格指数目标']),pctCell(summary['抖音差值'],true)]});}
    return '<div class="section-title"><span></span>六高价指 · MTD <small>（数据日期：'+escapeHtml(d.source_date||'')+'）</small></div>'+renderRows(rows,'six-high-price-grid');
  }
  // 唯一数据口径：经营会周报【优质款明细】C列=精品；不展示未同步的旧模板历史区。
  function renderQualityMtd(section){
    var headers=['小组','已引进','未引进','暂不引进','总计','引入率（目标50%）'];
    var rows=(section.rows||[]).filter(function(row){return row.excelRow>=93&&row.excelRow<=99;}).map(function(row){
      var cells=row.cells.slice(0,6).map(function(c){return Object.assign({},c);});
      var rate=cells[5];
      if(rate&&typeof rate.raw==='number'){
        rate.text=(rate.raw*100).toFixed(1)+'%'; rate.unit='%';
        if(Number(cells[4]&&cells[4].raw||0)>0&&rate.raw<0.5) rate.qualityFail=true;
      }
      return {excelRow:row.excelRow,cells:cells,total:normalizeGroupName(cells[0]&&cells[0].text)==='精品总'};
    });
    var date=(state.data&&state.data.dataDate)||'—';
    return '<div class="section-title"><span></span>优质款 · MTD <small>周报优质款明细 · 部门=精品 · 数据截至 '+escapeHtml(date)+'</small></div>'+renderRows([headerRow(headers)].concat(rows),'quality-weekly-grid');
  }
  function renderMtdWithoutTitle(section, title, firstRow, lastRow, cols){ var rows=(section.rows||[]).filter(function(row){return row.excelRow>=firstRow&&row.excelRow<=lastRow;}).map(function(row){return {excelRow:row.excelRow,cells:row.cells.slice(0,cols)};}); rows=markHeaders(rows,1); return '<div class="section-title"><span></span>'+escapeHtml(title)+'</div>'+renderRows(rows); }

  function renderPricePowerMtd(section){
    var headers=["小组","商品占比-曝光","商品占比-APP销售","APP占比-实际","APP占比-目标","APP占比-VS目标差距","APP占比-完成率","曝光","APP销售"];
    var rows=(section.rows||[]).filter(function(row){var label=row.cells&&row.cells[0]&&row.cells[0].text;return label==='四五星'||label==='大爆款';}).slice(0,2).map(function(row){
      var cells=row.cells.slice(0,9).map(function(c){return Object.assign({},c);});
      [3,4,5].forEach(function(i){var c=cells[i];if(c&&typeof c.raw==='number'){c.text=c.raw.toFixed(2);c.unit='';}});
      return {excelRow:row.excelRow,cells:cells};
    });
    return '<div class="section-title"><span></span>五星价格力 & 大爆款效率 · MTD</div>'+renderRows([headerRow(headers)].concat(rows));
  }
  function renderTrafficPanel(){
    var data=state.trafficFlowData;
    if(!data) return '<div class="loading">暂无流量数据</div>';
    var html='';
    var selectedGroup=state.periods.traffic_group||'精品总';
    
    function fmt(n){ if(n==null) return '—'; if(n>=1e8) return (n/1e8).toFixed(2)+'亿'; if(n>=1e4) return (n/1e4).toFixed(0)+'万'; return n.toLocaleString('zh-CN'); }
    function yoyPill(y){
      if(y==null) return '<span class="yoy-pill" style="color:#6b849e">—</span>';
      var cls=y>=0?'up':'down';
      var arrow=y>=0?'↑':'↓';
      var s=arrow+' '+Math.abs(y).toFixed(2)+'%';
      return '<span class="yoy-pill '+cls+'">'+s+'</span>';
    }
    
    // 模块1：当月MTD汇总
    var sourceDate=data.source_date||'—';
    html+='<div class="section-title"><span></span>曝光流量 & 浏览流量 · 当月MTD（截止 '+escapeHtml(sourceDate)+'）</div>';
    html+='<div class="excel-scroll"><table class="excel-table"><thead><tr>';
    html+='<th class="excel-cell is-header is-row-label">小组</th>';
    html+='<th class="excel-cell is-header">曝光流量</th>';
    html+='<th class="excel-cell is-header">曝光同比</th>';
    html+='<th class="excel-cell is-header">浏览流量</th>';
    html+='<th class="excel-cell is-header">浏览同比</th>';
    html+='<th class="excel-cell is-header">商详UV</th>';
    html+='</tr></thead><tbody>';
    
    var mtdData=data.mtd||[];
    var ordered=mtdData.slice().sort(function(a,b){return (a.group==='精品总')-(b.group==='精品总');});
    ordered.forEach(function(it){
      var isTotal=(it.group==='精品总');
      var rowCls=isTotal?'total-row':'';
      html+='<tr class="'+rowCls+'">';
      html+='<td class="excel-cell is-row-label">'+escapeHtml(it.group)+'</td>';
      html+='<td class="excel-cell">'+fmt(it.exposure)+'</td>';
      html+='<td class="excel-cell">'+yoyPill(it.exposure_yoy)+'</td>';
      html+='<td class="excel-cell">'+fmt(it.browse)+'</td>';
      html+='<td class="excel-cell">'+yoyPill(it.browse_yoy)+'</td>';
      html+='<td class="excel-cell">'+fmt(it.uv)+'</td>';
      html+='</tr>';
    });
    html+='</tbody></table></div>';
    
    // 模块2：当月明细（支持筛选小组）
    html+='<div class="section-title" style="margin-top:16px"><span></span>当月流量明细</div>';
    html+='<div class="filter-bar">';
    var groups=['精品总','饰品3组','珠宝1组','珠宝2组','珠宝3组','饰品1组','饰品2组'];
    groups.forEach(function(g){
      var active=selectedGroup===g?' active':'';
      html+='<button class="filter-btn'+active+'" data-section="traffic_group" data-period="'+g+'">'+g+'</button>';
    });
    html+='</div>';
    
    // 显示筛选小组的每日数据
    var dailyData=(data.daily&&data.daily[selectedGroup])||[];
    html+='<div class="excel-scroll"><table class="excel-table"><thead><tr>';
    html+='<th class="excel-cell is-header is-row-label">日期</th>';
    html+='<th class="excel-cell is-header">曝光流量</th>';
    html+='<th class="excel-cell is-header">曝光同比</th>';
    html+='<th class="excel-cell is-header">浏览流量</th>';
    html+='<th class="excel-cell is-header">浏览同比</th>';
    html+='<th class="excel-cell is-header">商详UV</th>';
    html+='</tr></thead><tbody>';
    
    dailyData.forEach(function(it){
      html+='<tr>';
      html+='<td class="excel-cell is-row-label">'+escapeHtml(it.date)+'</td>';
      html+='<td class="excel-cell">'+fmt(it.exposure)+'</td>';
      html+='<td class="excel-cell">'+yoyPill(it.exposure_yoy)+'</td>';
      html+='<td class="excel-cell">'+fmt(it.browse)+'</td>';
      html+='<td class="excel-cell">'+yoyPill(it.browse_yoy)+'</td>';
      html+='<td class="excel-cell">'+fmt(it.uv)+'</td>';
      html+='</tr>';
    });
    html+='</tbody></table></div>';
    
    html+='<div class="section-title" style="font-size:12px;color:#888;margin-top:8px"><span></span>数据来源：销售看板Excel · 总览表O-P列+BY-BZ列</div>';
    return html;
  }
  function renderBrandTierMtdPanel(){
    var d=state.brandTierData;
    if(!d) return '<div class="loading">暂无品牌分层数据</div>';
    var rows=d.summary||[];
    if(!rows.length) return '<div class="loading">暂无品牌分层汇总</div>';
    var sourceDate=d.source_date||'—';

    function fmt(n){if(n==null)return '—';if(Math.abs(n)>=1e8)return (n/1e8).toFixed(2)+'亿';if(Math.abs(n)>=1e4)return (n/1e4).toFixed(1)+'万';return String(Math.round(n));}
    function yoyPill(y){
      if(y==null) return '<span class="yoy-pill" style="color:#6b849e">—</span>';
      var cls=y>=0?'up':'down';
      var arrow=y>=0?'↑':'↓';
      return '<span class="yoy-pill '+cls+'">'+arrow+' '+Math.abs(y*100).toFixed(1)+'%</span>';
    }
    function pct(v){if(v==null)return '—';return (v*100).toFixed(1)+'%';}

    var html='<div class="section-title"><span></span>品牌分层 · MTD <small>（数据日期：'+escapeHtml(sourceDate)+'）</small></div>';
    html+='<div class="excel-scroll"><table class="excel-table brand-tier-grid"><thead><tr>';
    html+='<th class="excel-cell is-header is-row-label">标品/类穿戴</th><th class="excel-cell is-header" style="width:60px">品牌分层</th>';
    html+='<th class="excel-cell is-header">业绩</th><th class="excel-cell is-header" style="width:60px">业绩占比</th><th class="excel-cell is-header">同期业绩</th><th class="excel-cell is-header" style="width:70px">业绩同比</th>';
    html+='<th class="excel-cell is-header">曝光流量</th><th class="excel-cell is-header" style="width:60px">曝光占比</th><th class="excel-cell is-header">同期曝光</th><th class="excel-cell is-header" style="width:70px">曝光同比</th>';
    html+='</tr></thead><tbody>';

    var catLabels={};
    rows.forEach(function(r){
      var cat=r.category||'';
      var tier=r.tier||'';
      var isSub=r.is_subtotal;
      var isGrand=r.is_grand_total;
      var showCat=cat!==catLabels[cat];
      catLabels[cat]=cat;

      var rowCls='';
      if(isGrand) rowCls=' total-row';
      else if(isSub) rowCls=' subtotal-row';

      html+='<tr class="'+rowCls+'">';
      // 分类列
      if(isGrand){
        html+='<td class="excel-cell is-row-label" colspan="2"><b>精品总</b></td>';
      } else if(showCat){
        var catRowspan=0;
        for(var j=0;j<rows.length;j++){
          if(rows[j].category===cat&&!rows[j].is_grand_total) catRowspan++;
        }
        html+='<td class="excel-cell is-row-label" rowspan="'+catRowspan+'">'+escapeHtml(cat)+'</td>';
        html+='<td class="excel-cell is-row-label">'+escapeHtml(tier)+'</td>';
      } else {
        html+='<td class="excel-cell is-row-label">'+escapeHtml(tier)+'</td>';
      }

      html+='<td class="excel-cell">'+fmt(r.sales)+'</td>';
      html+='<td class="excel-cell">'+pct(r.sales_share)+'</td>';
      html+='<td class="excel-cell">'+fmt(r.sales_compare)+'</td>';
      html+='<td class="excel-cell">'+yoyPill(r.sales_yoy)+'</td>';
      html+='<td class="excel-cell">'+fmt(r.traffic)+'</td>';
      html+='<td class="excel-cell">'+pct(r.traffic_share)+'</td>';
      html+='<td class="excel-cell">'+fmt(r.traffic_compare)+'</td>';
      html+='<td class="excel-cell">'+yoyPill(r.traffic_yoy)+'</td>';
      html+='</tr>';
    });

    html+='</tbody></table></div>';
    return html;
  }
  function renderGrossProfit(section){
    var d=section&&section.data; if(!d)return '';
    var period=state.periods.gross_profit||'YTD';
    var headers=['时间','实收','净收入(毛利口径)','毛利值','毛利率','目标','落差','完成率','毛利盈余/缺口'];
    function pctCell(v){if(v==null)return{text:'',raw:null,type:'blank'};var n=parseFloat(v);if(isNaN(n))return{text:String(v),raw:v,type:'text'};return{text:(n*100).toFixed(2)+'%',raw:n,type:'number',unit:'%'};}
    function deltaCell(v){if(v==null)return{text:'',raw:null,type:'blank'};var n=parseFloat(v);if(isNaN(n))return{text:String(v),raw:v,type:'text'};var sign=n>0?'+':'';var trend=n>0?'up':n<0?'down':null;return{text:sign+(n*100).toFixed(2),raw:n,type:'number',unit:'pp',trend:trend};}
    function moneyCell(v){if(v==null)return{text:'',raw:null,type:'blank'};var n=parseFloat(v);if(isNaN(n))return{text:String(v),raw:v,type:'text'};var abs=Math.abs(n);var val=abs>=1e8?(n/1e8).toFixed(2):(n/1e4).toFixed(0);var unit=abs>=1e8?'亿':'万';return{text:val.toLocaleString('zh-CN'),raw:n,type:'number',unit:unit};}
    function dataRow(rowData,rowIdx,totalFlag){
      var cells=[{text:String(rowData.period),raw:rowData.period,type:'text'}];
      cells.push(moneyCell(rowData['实收']));
      cells.push(moneyCell(rowData['净收入']));
      cells.push(moneyCell(rowData['毛利值']));
      cells.push(pctCell(rowData['毛利率']));
      cells.push(pctCell(rowData['毛利率目标']));
      cells.push(deltaCell(rowData['落差']));
      cells.push(pctCell(rowData['完成率']));
      cells.push(moneyCell(rowData['毛利盈余']));
      return {excelRow:rowIdx,cells:cells,total:totalFlag};
    }
    // MTD 区
    var mtdHeader=markHeaders([{excelRow:0,cells:headers.map(function(h){return{text:h,raw:h,type:'text',header:true};})}],1);
    var mtdRows=[dataRow(d.mtd,1,true)];
    var mtdHtml='<div class="section-title" style="margin-top:8px;font-size:13px;font-weight:600"><span></span>MTD · '+escapeHtml(d.mtd.period)+'</div>'+renderRows(mtdHeader.concat(mtdRows));
    // YTD + 历史月份区
    var subPeriods=['YTD','历史月份'];
    var filterHtml='<div class="filter-bar">'+subPeriods.map(function(p){return'<button class="filter-btn'+(period===p?' active':'')+'" data-section="gross_profit" data-period="'+p+'">'+p+'</button>';}).join('')+'</div>';
    var subRows=[];
    if(period==='YTD'){
      subRows.push(dataRow(d.ytd,2,true));
      subRows.push(dataRow(d.year_target,3,true));
    }else{
      d.months.forEach(function(m,i){subRows.push(dataRow(m,i+4,false));});
    }
    var subHeader=markHeaders([{excelRow:0,cells:headers.map(function(h){return{text:h,raw:h,type:'text',header:true};})}],1);
    var subHtml='<div class="section-title" style="margin-top:8px;font-size:13px;font-weight:600"><span></span>'+escapeHtml(period==='YTD'?'YTD':'各月明细')+'</div>'+renderRows(subHeader.concat(subRows));
    // 标题+更新时间
    var sourceDate=d.source_date||'2026-08-03';
    var updatedAt=d.updated_at||'2026-08-03 14:28';
    var title='<div class="section-title gross-title"><span></span><div><b>毛利</b><em> · 数据截止 '+escapeHtml(sourceDate)+'</em></div><div style="font-size:11px;color:#999;margin-top:2px">更新于 '+escapeHtml(updatedAt)+'</div></div>';
    return title+mtdHtml+filterHtml+subHtml;
  }

  function renderTableSection(section){
    if(!section) return "";
    if(section.id==='gross_profit') return renderGrossProfit(section);
    if(section.id==='price_index_mtd') return renderPriceIndexMtd(section);
    if(section.id==='six_high') return renderSixHighMtd(section);
    if(section.id==='six_high_price_index') return renderSixHighPriceIndex(section);
    if(section.id==='quality_product_mtd') return renderQualityMtd(section);
    if(section.id==='machine_purchase_mtd'){
      var rows=(section.rows||[]).filter(function(row){return row.excelRow>=114&&row.excelRow<=118;}).map(function(row){
        var cells=row.cells.slice(0,8);
        normalizeTotalCell(cells[0]);
        ratePctCell(cells[3]); // 月完成率
        growthPctCell(cells[5]); // 业绩同比
        return {excelRow:row.excelRow,cells:cells};
      });
      rows=markHeaders(rows,1);
      return '<div class="section-title"><span></span>机采 · MTD</div>'+renderRows(rows);
    }
    if(section.id==='price_power_mtd') return renderPricePowerMtd(section);
    if(section.id==='traffic') return renderTrafficPanel();
    var hasFilter=!!PERIOD_CONFIG[section.id]; var baseTitle = section.title.replace(' · MTD / YTD / 历史月份',' · 历史月份').replace('YTD / 历史月份得分','历史月份得分').replace('YTD / 历史月份','历史月份'); var title=hasFilter ? baseTitle+' · '+activePeriod(section.id) : section.title; var rows=hasFilter ? periodRows(section, section.id) : (section.rows||[]); return '<div class="section-title"><span></span>'+escapeHtml(title)+'</div>'+renderPeriodFilter(section.id)+renderRows(rows); }

  function renderSalesPanel(){ var mtd=getSection('self_sales_mtd'), hist=getSection('self_sales_history'); return salesMtdTableSection(mtd)+renderTableSection(hist); }
  function discountRows(section, period){
    var cfg=PERIOD_CONFIG.internal_discount;
    var idx=cfg.periods[period] || cfg.periods.MTD;
    if(idx==='MTD_SNAPSHOT') idx=cfg.periods.MTD;
    var periodName=period==='MTD'?'MTD':period;
    // 不复用 Excel 的跨期合并表头；每个视图独立四列表头，避免表头混入数值。
    var headers=['小组',periodName+'去年内网价指',periodName+'今年内网价指','系数差'];
    var rows=(section.rows||[]).filter(function(row){return row.excelRow>=69&&row.excelRow<=75;}).map(function(row){
      var cells=idx.map(function(i){return Object.assign({},row.cells[i]||{text:'',type:'blank'});});
      // 去年/今年内网价指统一 xx.x%，系数差统一 ±x.xpp。
      pricePctCell(cells[1]); pricePctCell(cells[2]); ratioPpCell(cells[3]);
      if(cells[0]&&cells[0].text==='总'){cells[0].text='精品总';cells[0].raw='精品总';}
      return {excelRow:row.excelRow,total:row.excelRow===75,cells:cells};
    });
    return [headerRow(headers)].concat(rows);
  }

  function renderDiscountPanel(){
    var section=getSection('internal_discount');
    if(!section) return '<div class="loading">暂无数据</div>';
    var historyPeriod=state.periods.internal_discount || '1月';
    var html='';
    html+='<div class="section-title"><span></span>内网价指数 · MTD</div>'+renderRows(discountRows(section,'MTD'));
    html+='<div class="section-title"><span></span>内网价指数 · YTD</div>'+renderRows(discountRows(section,'YTD'));
    html+='<div class="section-title"><span></span>内网价指数 · 历史月份 · '+historyPeriod+'</div>';
    html+='<div class="filterbar">';
    ['1月','2月','3月','4月','5月','6月','7月'].forEach(function(p){ html+='<button class="filter-btn '+(historyPeriod===p?'active':'')+'" data-section="internal_discount" data-period="'+p+'">'+p+'</button>'; });
    html+='</div>'+renderRows(discountRows(section,historyPeriod));
    return html;
  }

  function renderAdjustmentPanel(){
    var d=state.adjustmentData;
    if(!d) return '<div class="loading">暂无调价率数据</div>';
    function num(v){return v==null?'—':Number(v).toLocaleString('zh-CN');}
    function statusRate(block){
      if(!block||block.rate==null) return '<span class="adjust-rate no-data">—</span>';
      var pct=(block.rate*100).toFixed(1)+'%';
      if(block.status==='fail') return '<span class="adjust-rate fail">'+pct+'<em>未达标</em></span>';
      return '<span class="adjust-rate pass">'+pct+'</span>';
    }
    function row(item,total){
      return '<tr class="'+(total?'total-row':'')+'"><td class="excel-cell is-row-label">'+escapeHtml(item.group)+'</td>'+
        '<td class="excel-cell is-number">'+num(item.overall.denominator)+'</td><td class="excel-cell is-number">'+num(item.overall.adjusted)+'</td><td class="excel-cell">'+statusRate(item.overall)+'</td>'+
        '<td class="excel-cell is-number section-divider">'+num(item.six_high.denominator)+'</td><td class="excel-cell is-number">'+num(item.six_high.adjusted)+'</td><td class="excel-cell">'+statusRate(item.six_high)+'</td></tr>';
    }
    var html='<div class="bi-card adjustment-head"><div class="bi-section-head"><h2>调价率</h2><span>整体目标 75% · 六高目标 80%</span></div><div class="adjustment-dates"><span>整体截至 '+escapeHtml(d.overall_source_date||'—')+'</span><span>六高截至 '+escapeHtml(d.six_high_source_date||'—')+'</span></div></div>';
    html+='<div class="excel-scroll"><table class="excel-table adjustment-grid"><thead><tr><th rowspan="2" class="excel-cell is-header is-row-label">小组</th><th colspan="3" class="excel-cell is-header">整体调价率</th><th colspan="3" class="excel-cell is-header section-divider">六高商品调价率</th></tr><tr><th class="excel-cell is-header">价高商品数</th><th class="excel-cell is-header">调价数</th><th class="excel-cell is-header">调价率</th><th class="excel-cell is-header section-divider">价高商品数</th><th class="excel-cell is-header">调价数</th><th class="excel-cell is-header">调价率</th></tr></thead><tbody>';
    (d.groups||[]).forEach(function(x){html+=row(x,false);});
    if(d.summary) html+=row(d.summary,true);
    html+='</tbody></table></div>';
    return html;
  }

  function normalizeSearchText(value){ return String(value||'').toLowerCase().replace(/\s+/g,''); }
  function brandMatches(query){
    var data=state.brandAdjustmentData; if(!data) return [];
    var q=normalizeSearchText(query);
    var rows=(data.brands||[]).map(function(b){
      var sn=String(b.sn||''), name=normalizeSearchText(b.brand), score=99;
      if(!q) score=b.denominator>0?20:30;
      else if(sn===q) score=0;
      else if(name===q) score=1;
      else if(sn.indexOf(q)===0) score=2;
      else if(name.indexOf(q)===0) score=3;
      else if(sn.indexOf(q)>=0) score=4;
      else if(name.indexOf(q)>=0) score=5;
      return {brand:b,score:score};
    }).filter(function(x){return x.score<99;});
    rows.sort(function(a,b){return a.score-b.score || b.brand.denominator-a.brand.denominator || String(a.brand.brand).localeCompare(String(b.brand.brand),'zh-CN');});
    return rows.slice(0,8).map(function(x){return x.brand;});
  }
  function brandTag(text,kind){ return text?'<span class="brand-tag '+(kind||'')+'">'+escapeHtml(text)+'</span>':''; }
  function renderBrandSearchResults(matches){
    if(!state.brandQuery) return '';
    if(!matches.length) return '<div class="brand-suggestions"><div class="brand-empty">未找到匹配品牌，请检查品牌名或SN</div></div>';
    var html='<div class="brand-suggestions">';
    matches.forEach(function(b){
      html+='<button type="button" class="brand-option" data-brand-sn="'+escapeHtml(b.sn)+'"><span><b>'+escapeHtml(b.brand||'未命名品牌')+'</b><small>SN '+escapeHtml(b.sn)+'</small></span><span class="brand-option-tags">'+brandTag(b.level,'level')+brandTag(b.shenyin,'shenyin')+'</span></button>';
    });
    return html+'</div>';
  }
  function brandMetricCard(label,value,kind){ return '<div class="brand-month-metric '+(kind||'')+'"><small>'+escapeHtml(label)+'</small><strong>'+escapeHtml(value)+'</strong></div>'; }
  function renderBrandCalendar(selected){
    var daily=selected.daily||[];
    if(!daily.length) return '<section class="brand-calendar-card"><div class="brand-calendar-empty">暂无日粒度数据</div></section>';
    var first=new Date(daily[0].date+'T00:00:00'); var offset=(first.getDay()+6)%7;
    var html='<section class="brand-calendar-card"><div class="brand-calendar-head"><div><b>'+escapeHtml(selected.month.split('截至')[0])+'调价日历</b><span>目标 75% · 数据截至 '+escapeHtml(selected.date)+'</span></div><div class="brand-calendar-legend"><i></i>未达标</div></div><div class="brand-weekdays"><span>一</span><span>二</span><span>三</span><span>四</span><span>五</span><span>六</span><span>日</span></div><div class="brand-calendar-grid">';
    for(var i=0;i<offset;i++) html+='<div class="brand-day spacer"></div>';
    daily.forEach(function(x){
      var rate=x.rate==null?'—':(x.rate*100).toFixed(1)+'%';
      var cls=x.is_future?'future':(!x.has_data?'empty':(x.rate<0.75?'fail':'pass'));
      var note=x.is_future?'待更新':(!x.has_data?'无明细':'');
      html+='<article class="brand-day '+cls+'"><div class="brand-day-date"><b>'+x.day+'</b>'+(note?'<em>'+note+'</em>':'')+'</div><dl><div><dt>价高</dt><dd>'+(x.has_data?Number(x.denominator).toLocaleString('zh-CN'):'—')+'</dd></div><div><dt>调价</dt><dd>'+(x.has_data?Number(x.adjusted).toLocaleString('zh-CN'):'—')+'</dd></div><div class="rate"><dt>调价率</dt><dd>'+rate+'</dd></div></dl></article>';
    });
    return html+'</div></section>';
  }
  function renderProblemBrandsSection(){
    var data=state.problemBrandsData;
    if(!data||!data.problems||data.problems.length===0) return '';
    
    var html='<section class="problem-brands-section">';
    html+='<div class="section-title"><span></span>问题品牌提醒 <small>（数据截至 '+escapeHtml(data.source_date||'—')+'）</small></div>';
    html+='<div class="problem-brands-summary">';
    html+='<span class="problem-count">共 <b>'+data.total_problem_count+'</b> 个问题品牌</span>';
    if(data.overall_problem_count>0) html+='<span class="problem-tag overall">整体调价率≤10%: '+data.overall_problem_count+'个</span>';
    if(data.six_high_problem_count>0) html+='<span class="problem-tag six-high">六高调价率<80%: '+data.six_high_problem_count+'个</span>';
    html+='</div>';
    
    html+='<div class="excel-scroll"><table class="excel-table problem-brands-table"><thead><tr>';
    html+='<th class="excel-cell is-header is-row-label">品牌</th>';
    html+='<th class="excel-cell is-header">小组</th>';
    html+='<th class="excel-cell is-header">价高商品数<br><small>(整体)</small></th>';
    html+='<th class="excel-cell is-header">调价数<br><small>(整体)</small></th>';
    html+='<th class="excel-cell is-header">调价率<br><small>(整体)</small></th>';
    html+='<th class="excel-cell is-header">六高价高数</th>';
    html+='<th class="excel-cell is-header">六高调价数</th>';
    html+='<th class="excel-cell is-header">六高调价率</th>';
    html+='<th class="excel-cell is-header">问题类型</th>';
    html+='</tr></thead><tbody>';
    
    data.problems.forEach(function(p){
      html+='<tr>';
      html+='<td class="excel-cell is-row-label"><b>'+escapeHtml(p.brand||'—')+'</b><br><small>SN: '+escapeHtml(p.sn||'')+'</small></td>';
      html+='<td class="excel-cell">'+escapeHtml(p.group||'—')+'</td>';
      
      // 整体调价率
      var overallRate=p.overall_rate;
      var overallRateText=overallRate!=null?(overallRate*100).toFixed(1)+'%':'—';
      var overallRateClass=overallRate!=null&&overallRate<=0.10?'problem-rate':'normal-rate';
      html+='<td class="excel-cell">'+(p.overall_denominator||0).toLocaleString('zh-CN')+'</td>';
      html+='<td class="excel-cell">'+(p.overall_adjusted||0).toLocaleString('zh-CN')+'</td>';
      html+='<td class="excel-cell '+overallRateClass+'">'+overallRateText+'</td>';
      
      // 六高调价率
      var sixHighRate=p.six_high_rate;
      var sixHighRateText=sixHighRate!=null?(sixHighRate*100).toFixed(1)+'%':'—';
      var sixHighRateClass=sixHighRate!=null&&sixHighRate<0.80?'problem-rate':'normal-rate';
      html+='<td class="excel-cell">'+(p.six_high_price_high||0).toLocaleString('zh-CN')+'</td>';
      html+='<td class="excel-cell">'+(p.six_high_adjusted||0).toLocaleString('zh-CN')+'</td>';
      html+='<td class="excel-cell '+sixHighRateClass+'">'+sixHighRateText+'</td>';
      
      // 问题类型
      var issueTypes=[];
      if(p.is_overall_problem) issueTypes.push('<span class="issue-tag overall">整体≤10%</span>');
      if(p.is_six_high_problem) issueTypes.push('<span class="issue-tag six-high">六高<80%</span>');
      html+='<td class="excel-cell">'+issueTypes.join(' ')+'</td>';
      html+='</tr>';
    });
    
    html+='</tbody></table></div>';
    html+='<div class="problem-brands-note"><small>筛选规则：① 整体调价率：价高商品数>10 且 调价率≤10%；② 六高调价率：调价率<80%</small></div>';
    html+='</section>';
    return html;
  }

  function renderBrandSalesTrafficPanel(){
    var d=state.brandTierData;
    if(!d) return '<div class="loading">暂无品牌销售流量数据</div>';
    var brands=(d.brands||[]).slice();
    if(!brands.length) return '<div class="loading">暂无品牌明细</div>';

    var sourceDate=d.source_date||'—';
    var selectedGroup=state.periods.brand_sales_traffic_group||'全部';
    var selectedTier=state.periods.brand_sales_traffic_tier||'全部';
    var brandQuery=state.brandSalesTrafficQuery||'';

    function fmt(n){if(n==null)return '—';if(Math.abs(n)>=1e8)return (n/1e8).toFixed(2)+'亿';if(Math.abs(n)>=1e4)return (n/1e4).toFixed(1)+'万';return String(Math.round(n));}
    function yoyPill(y){
      if(y==null) return '<span class="yoy-pill" style="color:#6b849e">—</span>';
      var cls=y>=0?'up':'down';
      var arrow=y>=0?'↑':'↓';
      return '<span class="yoy-pill '+cls+'">'+arrow+' '+Math.abs(y*100).toFixed(1)+'%</span>';
    }

    // 筛选
    if(selectedGroup!=='全部'){
      brands=brands.filter(function(b){return b.group===selectedGroup;});
    }
    if(selectedTier!=='全部'){
      brands=brands.filter(function(b){return b.tier===selectedTier;});
    }
    if(brandQuery){
      var q=brandQuery.toLowerCase();
      brands=brands.filter(function(b){
        return b.brand.toLowerCase().indexOf(q)!==-1 || b.sn.indexOf(q)!==-1;
      });
    }

    // 排序
    brands.sort(function(a,b){
      var sa=a.sales||0, sb=b.sales||0;
      if(sa!==sb) return sb-sa;
      return (b.traffic||0)-(a.traffic||0);
    });

    var html='<div class="section-title"><span></span>品牌销售流量 · MTD <small>（数据日期：'+escapeHtml(sourceDate)+'）</small></div>';

    // 筛选器
    html+='<div class="filter-bar" style="margin-bottom:8px">';
    var groups=['全部','饰品1组','饰品2组','饰品3组','珠宝1组','珠宝2组','珠宝3组'];
    groups.forEach(function(g){
      var active=selectedGroup===g?' active':'';
      html+='<button class="filter-btn'+active+'" data-section="brand_sales_traffic_group" data-period="'+g+'">'+g+'</button>';
    });
    html+='</div>';

    html+='<div class="filter-bar" style="margin-bottom:8px">';
    var tiers=['全部','S1','S2','S3','高价值','矩阵非高','双非'];
    tiers.forEach(function(t){
      var active=selectedTier===t?' active':'';
      html+='<button class="filter-btn'+active+'" data-section="brand_sales_traffic_tier" data-period="'+t+'">'+t+'</button>';
    });
    html+='</div>';

    // 搜索
    html+='<div class="brand-search-wrap" style="margin-bottom:8px"><span class="brand-search-icon">⌕</span><input id="brandSalesTrafficSearch" type="search" autocomplete="off" inputmode="search" placeholder="输入品牌名称或SN" value="'+escapeHtml(brandQuery)+'"><button type="button" class="brand-clear" aria-label="清空" style="right:4px">×</button></div>';

    // 汇总
    var totalSales=brands.reduce(function(s,b){return s+(b.sales||0);},0);
    var totalTraffic=brands.reduce(function(s,b){return s+(b.traffic||0);},0);
    html+='<div style="font-size:12px;color:#6b849e;margin-bottom:8px">品牌数：<b>'+brands.length+'</b> ｜ 业绩合计：<b>'+fmt(totalSales)+'</b> ｜ 曝光合计：<b>'+fmt(totalTraffic)+'</b></div>';

    // 表格
    html+='<div class="excel-scroll"><table class="excel-table brand-sales-traffic-table"><thead><tr>';
    html+='<th class="excel-cell is-header is-row-label">小组</th><th class="excel-cell is-header" style="width:60px">品牌分层</th><th class="excel-cell is-header" style="width:80px">品牌SN</th><th class="excel-cell is-header" style="min-width:100px">品牌名称</th>';
    html+='<th class="excel-cell is-header">业绩</th><th class="excel-cell is-header">同期业绩</th><th class="excel-cell is-header" style="width:70px">业绩同比</th>';
    html+='<th class="excel-cell is-header">曝光流量</th><th class="excel-cell is-header">同期曝光</th><th class="excel-cell is-header" style="width:70px">曝光同比</th>';
    html+='</tr></thead><tbody>';

    brands.forEach(function(b){
      html+='<tr>';
      html+='<td class="excel-cell is-row-label">'+escapeHtml(b.group||'')+'</td>';
      html+='<td class="excel-cell">'+escapeHtml(b.tier||'—')+'</td>';
      html+='<td class="excel-cell" style="font-size:11px;color:#6b849e">'+escapeHtml(b.sn)+'</td>';
      html+='<td class="excel-cell is-row-label" style="max-width:140px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="'+escapeHtml(b.brand||'')+'">'+escapeHtml(b.brand||'')+'</td>';
      html+='<td class="excel-cell">'+fmt(b.sales)+'</td>';
      html+='<td class="excel-cell">'+fmt(b.sales_compare)+'</td>';
      html+='<td class="excel-cell">'+yoyPill(b.sales_yoy)+'</td>';
      html+='<td class="excel-cell">'+fmt(b.traffic)+'</td>';
      html+='<td class="excel-cell">'+fmt(b.traffic_compare)+'</td>';
      html+='<td class="excel-cell">'+yoyPill(b.traffic_yoy)+'</td>';
      html+='</tr>';
    });
    html+='</tbody></table></div>';
    return html;
  }
  function renderBrandPriceIndexPanel(){
    var d=state.brandPriceIndexData;
    if(!d) return '<div class="loading">暂无品牌外网价格指数数据</div>';
    
    var selectedGroup=state.periods.brand_price_index_group||'饰品1组';
    var groups=d.groups||[];
    var currentGroup=groups.find(function(g){return g.group===selectedGroup;})||groups[0];
    
    function fmtRate(v){ if(v==null) return '—'; return (v*100).toFixed(1)+'%'; }
    function fmtDiff(v){ if(v==null) return '—'; var s=v>=0?'+':''; return s+v.toFixed(1)+'pp'; }
    function diffClass(v){ if(v==null) return ''; return v>=0?'normal-rate':'problem-rate'; }
    
    var html='<section class="brand-price-index-section">';
    html+='<div class="section-title"><span></span>品牌外网价格指数 <small>（数据截至 '+escapeHtml(d.source_date||'—')+'）</small></div>';
    
    // 小组筛选
    html+='<div class="filter-bar">';
    groups.forEach(function(g){
      var active=selectedGroup===g.group?' active':'';
      html+='<button class="filter-btn'+active+'" data-section="brand_price_index_group" data-period="'+g.group+'">'+g.group+'</button>';
    });
    html+='</div>';
    
    if(!currentGroup||!currentGroup.brands||currentGroup.brands.length===0){
      html+='<div class="brand-empty">该小组暂无品牌数据</div>';
      html+='</section>';
      return html;
    }
    
    // 合并表格：天猫+抖音
    html+='<div class="excel-scroll"><table class="excel-table brand-price-index-table"><thead><tr>';
    html+='<th class="excel-cell is-header is-row-label" rowspan="2">品牌</th>';
    html+='<th class="excel-cell is-header" rowspan="2">等级</th>';
        html+='<th class="excel-cell is-header" colspan="4">天猫</th>';
    html+='<th class="excel-cell is-header" colspan="4">抖音</th>';
    html+='</tr><tr>';
    html+='<th class="excel-cell is-header">价格指数</th>';
    html+='<th class="excel-cell is-header">对标值</th>';
    html+='<th class="excel-cell is-header">本月目标</th>';
    html+='<th class="excel-cell is-header">完成差值<br><small>(pp)</small></th>';
    html+='<th class="excel-cell is-header">价格指数</th>';
    html+='<th class="excel-cell is-header">对标值</th>';
    html+='<th class="excel-cell is-header">本月目标</th>';
    html+='<th class="excel-cell is-header">完成差值<br><small>(pp)</small></th>';
    html+='</tr></thead><tbody>';
    
    currentGroup.brands.forEach(function(b){
      html+='<tr>';
      html+='<td class="excel-cell is-row-label"><b>'+escapeHtml(b.brand||'—')+'</b><br><small>SN: '+escapeHtml(b.sn||'')+'</small></td>';
      html+='<td class="excel-cell">'+escapeHtml(b.level||'—')+'</td>';
      // 天猫
      html+='<td class="excel-cell">'+fmtRate(b.tmall_rate)+'</td>';
      html+='<td class="excel-cell">'+fmtRate(b.tmall_target)+'</td>';
      html+='<td class="excel-cell">'+fmtRate(b.tmall_goal)+'</td>';
      html+='<td class="excel-cell '+diffClass(b.tmall_diff)+'">'+fmtDiff(b.tmall_diff)+'</td>';
      // 抖音
      html+='<td class="excel-cell">'+fmtRate(b.douyin_rate)+'</td>';
      html+='<td class="excel-cell">'+fmtRate(b.douyin_target)+'</td>';
      html+='<td class="excel-cell">'+fmtRate(b.douyin_goal)+'</td>';
      html+='<td class="excel-cell '+diffClass(b.douyin_diff)+'">'+fmtDiff(b.douyin_diff)+'</td>';
      html+='</tr>';
    });
    
    html+='</tbody></table></div>';
    html+='</section>';
    return html;
  }

  function renderBrandAdjustmentPanel(){
    var d=state.brandAdjustmentData;
    if(!d) return '<div class="loading">暂无品牌调价率数据</div>';
    var matches=brandMatches(state.brandQuery);
    var selected=(d.brands||[]).find(function(b){return b.sn===state.selectedBrandSn;});
    var html='<section class="brand-search-card"><label for="brandSearchInput">搜索品牌</label><div class="brand-search-wrap"><span class="brand-search-icon">⌕</span><input id="brandSearchInput" type="search" autocomplete="off" inputmode="search" placeholder="输入品牌名称或品牌SN" value="'+escapeHtml(state.brandQuery)+'"><button type="button" class="brand-clear" aria-label="清空品牌搜索">×</button></div>'+renderBrandSearchResults(matches)+'</section>';
    
    // 问题品牌提醒区域
    html+=renderProblemBrandsSection();
    
    if(!selected){
      html+='<section class="brand-onboarding"><b>选择品牌查看本月调价率</b><span>支持中文、英文品牌名和品牌SN模糊搜索</span><small>数据截至 '+escapeHtml(d.source_date||'—')+'</small></section>';
      return html;
    }
    var hasMetric=Number(selected.row_count||0)>0;
    var rate=selected.rate==null?'—':(selected.rate*100).toFixed(1)+'%';
    var fail=selected.rate!=null&&selected.rate<0.75;
    html+='<section class="brand-context-card"><div class="brand-context-head"><div><h2>'+escapeHtml(selected.brand)+'</h2><p>品牌SN '+escapeHtml(selected.sn)+'</p></div><button type="button" class="brand-change">更换品牌</button></div><div class="brand-tags">'+brandTag(selected.level,'level')+brandTag(selected.shenyin,'shenyin')+(selected.groups||[]).map(function(g){return brandTag(g,'group');}).join('')+'</div></section>';
    html+='<section class="brand-month-summary">'+brandMetricCard('本月价高数',hasMetric?Number(selected.denominator).toLocaleString('zh-CN'):'—',hasMetric?'den':'empty')+brandMetricCard('本月调价数',hasMetric?Number(selected.adjusted).toLocaleString('zh-CN'):'—',hasMetric?'adj':'empty')+brandMetricCard('本月调价率',rate,selected.rate==null?'empty':fail?'fail':'pass')+'</section>';
    if(!hasMetric) html+='<div class="brand-no-metric">该品牌本月无调价指标明细；品牌存在于销售看板目录，但 data 与 data非神银指标表均无对应记录。</div>';
    html+=renderBrandCalendar(selected);
    html+='<p class="brand-scope-note">日历口径：销售看板 data 表按品牌SN+日期汇总价高商品数与调价商品数；月度汇总为每日原始计数之和。无明细显示“—”，截至日之后显示“待更新”。</p>';
    return html;
  }

  function renderViewModeSwitch(){ return '<div class="view-mode-switch" role="tablist" aria-label="数据视角"><button type="button" role="tab" class="view-mode-btn '+(state.viewMode==='group'?'active':'')+'" data-view-mode="group">小组视角</button><button type="button" role="tab" class="view-mode-btn '+(state.viewMode==='brand'?'active':'')+'" data-view-mode="brand">品牌视角</button></div>'; }
  function renderBrandTabs(){ var html='<div class="mobile-tabs brand-tabs">'; BRAND_TABS.forEach(function(tab){html+='<button class="tab-btn '+(state.activeBrandTab===tab.id?'active':'')+'" data-brand-tab="'+tab.id+'">'+escapeHtml(tab.label)+'</button>';});return html+'</div>'; }

  function renderGenericPanel(tab){ var html=""; tab.sectionIds.forEach(function(id){ var s=getSection(id); if(s) html+=renderTableSection(s); }); return html||'<div class="loading">暂无数据</div>'; }
  function renderTabs(){ var html='<div class="mobile-tabs">'; TABS.forEach(function(tab){html+='<button class="tab-btn '+(state.activeTab===tab.id?'active':'')+'" data-tab="'+tab.id+'">'+escapeHtml(tab.label)+'</button>';}); return html+'</div>'; }
  function updateBrandSuggestions(){
    var input=document.getElementById('brandSearchInput'); if(!input) return;
    var card=input.closest('.brand-search-card'); if(!card) return;
    var old=card.querySelector('.brand-suggestions');
    var html=renderBrandSearchResults(brandMatches(state.brandQuery));
    if(old) old.remove();
    if(html) card.insertAdjacentHTML('beforeend',html);
    bindBrandOptionClicks(card);
  }
  function bindBrandOptionClicks(scope){
    (scope||document).querySelectorAll('.brand-option').forEach(function(btn){
      btn.onclick=function(){state.selectedBrandSn=btn.getAttribute('data-brand-sn');state.brandQuery='';renderDashboard();};
      btn.ontouchstart=function(){state.selectedBrandSn=btn.getAttribute('data-brand-sn');state.brandQuery='';renderDashboard();};
    });
  }
  function bindBrandInteractions(){
    var input=document.getElementById('brandSearchInput');
    if(input&&!input.dataset.bound){
      input.dataset.bound='1';
      input.addEventListener('compositionstart',function(){state.brandComposing=true;});
      input.addEventListener('compositionend',function(){state.brandComposing=false;state.brandQuery=input.value;updateBrandSuggestions();});
      input.addEventListener('input',function(){
        if(state.brandComposing) return;
        state.brandQuery=input.value;
        updateBrandSuggestions();
      });
    }
    bindBrandOptionClicks(document);
    var clear=document.querySelector('.brand-clear');if(clear)clear.onclick=function(){state.brandQuery='';state.selectedBrandSn=null;renderDashboard();setTimeout(function(){var x=document.getElementById('brandSearchInput');if(x)x.focus();},0);};
    var change=document.querySelector('.brand-change');if(change)change.onclick=function(){state.brandQuery='';state.selectedBrandSn=null;renderDashboard();setTimeout(function(){var x=document.getElementById('brandSearchInput');if(x)x.focus();},0);};

    // 品牌销售流量搜索（防抖 300ms）
    var stInput=document.getElementById('brandSalesTrafficSearch');
    if(stInput&&!stInput.dataset.bound){
      stInput.dataset.bound='1';
      stInput.addEventListener('compositionstart',function(){state.brandSalesTrafficComposing=true;});
      stInput.addEventListener('compositionend',function(){state.brandSalesTrafficComposing=false;state.brandSalesTrafficQuery=stInput.value;if(state.brandSalesTrafficTimer)clearTimeout(state.brandSalesTrafficTimer);state.brandSalesTrafficTimer=setTimeout(function(){renderDashboard();},300);});
      stInput.addEventListener('input',function(){
        if(state.brandSalesTrafficComposing) return;
        state.brandSalesTrafficQuery=stInput.value;
        if(state.brandSalesTrafficTimer)clearTimeout(state.brandSalesTrafficTimer);
        state.brandSalesTrafficTimer=setTimeout(function(){renderDashboard();},300);
      });
    }
    var stClear=document.querySelector('#brandSalesTrafficSearch+.brand-clear');
    if(stClear)stClear.onclick=function(){state.brandSalesTrafficQuery='';renderDashboard();setTimeout(function(){var x=document.getElementById('brandSalesTrafficSearch');if(x)x.focus();},0);};
  }
  function renderDashboard(){ var data=state.data; if(!data) return; var meta=data.meta||{}; $navbarDate.textContent=meta.dataDate?'截止 '+meta.dataDate:'—'; if($periodToggle) $periodToggle.style.display='none'; var activeTab=TABS.find(function(t){return t.id===state.activeTab;})||TABS[0]; var body=state.viewMode==='brand'?(renderBrandTabs()+'<main class="mobile-panel">'+(state.activeBrandTab==='brand-price-index'?renderBrandPriceIndexPanel():(state.activeBrandTab==='brand-sales-traffic'?renderBrandSalesTrafficPanel():renderBrandAdjustmentPanel()))+'</main>'):(renderTabs()+'<main class="mobile-panel">'+(state.activeTab==='sales'?renderSalesPanel():(state.activeTab==='discount'?renderDiscountPanel():(state.activeTab==='traffic'?renderTrafficPanel():(state.activeTab==='adjustment'?renderAdjustmentPanel():(state.activeTab==='brand-tier'?renderBrandTierMtdPanel():renderGenericPanel(activeTab))))))+'</main>'); $modulesContainer.innerHTML=renderViewModeSwitch()+body; document.querySelectorAll('.view-mode-btn').forEach(function(btn){btn.onclick=function(){state.viewMode=btn.getAttribute('data-view-mode');renderDashboard();window.scrollTo(0,0);};}); document.querySelectorAll('[data-tab]').forEach(function(btn){btn.onclick=function(){state.activeTab=btn.getAttribute('data-tab');renderDashboard();window.scrollTo(0,0);};}); document.querySelectorAll('[data-brand-tab]').forEach(function(btn){btn.onclick=function(){state.activeBrandTab=btn.getAttribute('data-brand-tab');renderDashboard();window.scrollTo(0,0);};}); document.querySelectorAll('.filter-btn').forEach(function(btn){btn.onclick=function(){state.periods[btn.getAttribute('data-section')]=btn.getAttribute('data-period');renderDashboard();};}); bindBrandInteractions(); }

  function enterDashboard(){ $loginError.textContent=""; $loginPage.style.display="none"; $dashboard.classList.add("active"); loadData(); }
  function handleLogin(){ var pwd=$passwordInput.value.trim(); if(!pwd){$loginError.textContent='请输入密码';return;} $loginError.textContent='正在登录…'; loginByApi(pwd).then(function(){enterDashboard();}).catch(function(){ if(pwd===STATIC_PREVIEW_PASSWORD){state.token='static-preview';enterDashboard();return;} $loginError.textContent='密码错误'; $passwordInput.value=''; $passwordInput.focus(); }); }

  window.dashboardLogin=handleLogin; $loginBtn.onclick=handleLogin; $loginBtn.addEventListener('click',handleLogin); $passwordInput.addEventListener('keydown',function(e){if(e.key==='Enter')handleLogin();}); $logoutBtn.addEventListener('click',function(){state.token=null;state.data=null;$passwordInput.value='';$dashboard.classList.remove('active');$loginPage.style.display='flex';$loginError.textContent='';});
})();
