/* 精品货价监控数据看板 — Mobile BI / No Pie, No Ring */
(function () {
  "use strict";

  var API_BASE = window.DASHBOARD_API_BASE || "http://127.0.0.1:8900";
  var FALLBACK_URL = "data/excel_view.json";
  var STATIC_PREVIEW_PASSWORD = "vip2026";
  var PERIODS = ["YTD", "1月", "2月", "3月", "4月", "5月", "6月"];

  var state = { token: null, data: null, activeTab: "sales", periods: {} };

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
    { id: "gross", label: "毛利", sectionIds: ["gross_profit"] },
    { id: "price", label: "外网价指", sectionIds: ["price_index_mtd", "price_index_history"] },
    { id: "discount", label: "内网折扣", sectionIds: ["internal_discount"] },
    { id: "liugao", label: "六高", sectionIds: ["six_high"] },
    { id: "quality", label: "优质款", sectionIds: ["quality_product_mtd", "quality_product_history"] },
    { id: "machine", label: "机采", sectionIds: ["machine_purchase_mtd", "machine_purchase_history"] },
    { id: "power", label: "五星价格力", sectionIds: ["price_power_mtd", "price_power_history"] },
  ];

  var PERIOD_CONFIG = {
    self_sales_history: { rowStart: 17, rowEnd: 23, periods: { "YTD": [0,1,2,3,4,5], "1月": [0,6,7,8,9,10], "2月": [0,11,12,13,14,15], "3月": [0,16,17,18,19,20], "4月": [0,21,22,23,24,25], "5月": [0,26,27,28,29,30], "6月": [0,31,32,33,34,35] }, headers: function(p){ return p==="YTD" ? ["小组","YTD目标","YTD完成","YTD完成率","同期","业绩同比"] : ["小组",p+"目标",p+"完成",p+"完成率","同期","业绩同比"]; } },
    price_index_history: { rowStart: 56, rowEnd: 62, periods: { "YTD": [0,1,2,3,4,5,6,7,8,9,10,11], "1月": [0,12,13,14,15,16,17,18,19,20,21,22], "2月": [0,23,24,25,26,27,28,29,30,31,32,33], "3月": [0,34,35,36,37,38,39,40,41,42,43,44], "4月": [0,45,46,47,48,49,50,51,52,53,54,55], "5月": [0,56,57,58,59,60,61,62,63,64,65,66], "6月": [0,67,68,69,70,71,72,73,74,75,76,77] }, headers: function(p){ return ["小组",p+"综合得分",p+"天猫得分",p+"抖音得分",p+"天猫权重",p+"抖音权重",p+"天猫价指",p+"天猫目标",p+"天猫差值",p+"抖音价指",p+"抖音目标",p+"抖音差值"]; } },
    internal_discount: { rowStart: 68, rowEnd: 74, periods: { "MTD": [0,1,2,3], "YTD": [0,4,5,6], "1月": [0,7,8,9], "2月": [0,10,11,12], "3月": [0,13,14,15], "4月": [0,16,17,18], "5月": [0,19,20,21], "6月": [0,22,23,24] }, headers: function(p){ return ["小组",p+"去年内网价指",p+"今年内网价指",p+"系数差"]; }, periodList: ["MTD","YTD","1月","2月","3月","4月","5月","6月"] },
    quality_product_history: { rowStart: 103, rowEnd: 109, periods: { "YTD": [0,1,2,3,4,5], "1月": [0,6,7,8,9,10], "2月": [0,11,12,13,14,15], "3月": [0,16,17,18,19,20], "4月": [0,21,22,23,24,25], "5月": [0,26,27,28,29,30], "6月": [0,31,32,33,34,35] }, headers: function(p){ return ["小组",p+"已引进",p+"未引进",p+"暂不引进",p+"总计",p+"引入率"]; } },
    machine_purchase_history: { rowStart: 123, rowEnd: 126, periods: { "YTD": [0,1,2,3,4,5], "1月": [0,6,7,8,9,10], "2月": [0,11,12,13,14,15], "3月": [0,16,17,18,19,20], "4月": [0,21,22,23,24,25], "5月": [0,26,27,28,29,30], "6月": [0,31,32,33,34,35] }, headers: function(p){ return p==="YTD" ? ["小组","YTD目标","YTD完成","YTD完成率","同期","业绩同比"] : ["小组",p+"目标",p+"完成",p+"完成率","同期","业绩同比"]; } },
    price_power_history: { rowStart: 139, rowEnd: 152, periods: { "YTD": null, "6月": null }, headers: function(p){ return ["时间","维度","指标","曝光","曝光占比","APP销售","APP销售占比","实际","目标","VS目标差距","完成率"]; }, periodList: ["YTD","6月"] }
  };

  function apiFetch(path, options) {
    options = options || {}; options.headers = options.headers || {};
    if (state.token && state.token !== "static-preview") options.headers.Authorization = "Bearer " + state.token;
    return fetch(API_BASE + path, options).then(function (res) { if (!res.ok) return res.text().then(function (text) { throw new Error("HTTP " + res.status + " " + text); }); return res.json(); });
  }
  function loginByApi(password) { return apiFetch("/api/login", { method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify({password:password}) }).then(function(json){ state.token=json.token; return json; }); }
  function loadData() { return apiFetch("/api/excel_view").then(function(json){ state.data=json.data||json; renderDashboard(); }).catch(function(){ return fetch(FALLBACK_URL).then(function(res){ if(!res.ok) throw new Error("HTTP "+res.status); return res.json(); }).then(function(json){ state.data=json; renderDashboard(); }).catch(function(){ $modulesContainer.innerHTML='<div class="loading">数据加载失败，请稍后重试</div>'; }); }); }

  function escapeHtml(value){ return String(value==null?"":value).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"); }
  function parseNum(cell){ if(!cell) return null; var v=cell.raw; if(typeof v==="number") return v; var n=Number(v); return isNaN(n)?null:n; }
  function textWithUnit(cell){ if(!cell||!cell.text) return "—"; return cell.text + (cell.unit && cell.unit !== "%" ? cell.unit : ""); }
  function clamp(v,min,max){ return Math.max(min,Math.min(max,v)); }
  function getSection(id){ return (state.data.sections||[]).find(function(s){ return s.id===id; }); }
  function isTotalRow(row){ return row&&row.cells&&row.cells[0]&&row.cells[0].text==="总"; }
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

  function cellClass(cell,rowIndex,colIndex){ var classes=["excel-cell"]; var ctx=(cell&&cell.context||"")+" "+(cell&&cell.text||""); if(!cell||!cell.text) classes.push("is-blank"); if(cell&&cell.type==="number") classes.push("is-number"); if(rowIndex===0 || (cell&&cell.header)) classes.push("is-header"); if(colIndex===0) classes.push("is-row-label"); if(cell&&cell.merge&&cell.merge.colspan>1) classes.push("is-merged-head"); if(/同比/.test(ctx)) classes.push("is-yoy"); if(/完成率|达成率|阶段完成率|引入率|占比/.test(ctx)) classes.push("is-progress"); return classes.join(" "); }
  function renderEnhancedContent(cell){ if(!cell||!cell.text) return '<span class="blank-placeholder">—</span>'; var ctx=cell.context||""; var raw=parseNum(cell); var text=escapeHtml(cell.text); var unit=cell.unit&&cell.unit!=="%"?'<small>'+escapeHtml(cell.unit)+'</small>':""; if(/同比/.test(ctx)&&raw!=null){ var status=raw>=0?"up":"down"; return '<span class="yoy-pill '+status+'">'+(raw>=0?'↑':'↓')+' '+text+'</span>'; } if(/完成率|达成率|阶段完成率|引入率|占比/.test(ctx)&&raw!=null&&cell.unit==="%"){ var pct=raw>2?raw:raw*100; var cls=pct>=100?"good":pct<80?"bad":"normal"; return '<div class="progress-cell '+cls+'"><span>'+text+'</span><div class="mini-progress"><i style="width:'+clamp(pct,0,120)+'%"></i></div></div>'; } return text+unit; }
  function renderCell(cell,rowIndex,colIndex){ if(cell&&cell.merge&&cell.merge.covered) return ''; var title=cell&&cell.coord&&cell.text?' title="'+escapeHtml(cell.coord+' '+(cell.raw==null?'':cell.raw))+'"':""; var span=''; if(cell&&cell.merge&&cell.merge.rowspan>1) span+=' rowspan="'+cell.merge.rowspan+'"'; if(cell&&cell.merge&&cell.merge.colspan>1) span+=' colspan="'+cell.merge.colspan+'"'; return '<td class="'+cellClass(cell,rowIndex,colIndex)+'"'+title+span+'>'+renderEnhancedContent(cell)+'</td>'; }

  function markHeaders(rows, count){ return (rows||[]).map(function(row,idx){ if(idx<count){ row={excelRow:row.excelRow,cells:(row.cells||[]).map(function(cell){ var c=Object.assign({},cell); c.header=true; return c; })}; } return row; }); }
  function headerRow(names){ return {excelRow:0,cells:names.map(function(name){return {text:name,raw:name,type:"text",header:true};})}; }
  function renderRows(rows){ var html='<div class="excel-scroll"><table class="excel-table">'; rows.forEach(function(row,rowIndex){ var total=isTotalRow(row)?" total-row":""; html+='<tr class="'+total+'" data-excel-row="'+row.excelRow+'">'; (row.cells||[]).forEach(function(cell,colIndex){html+=renderCell(cell,rowIndex,colIndex);}); html+='</tr>'; }); html+='</table></div>'; return html; }

  function salesMtdTableSection(section){ var rows=(section.rows||[]).filter(function(row){return row.excelRow>=5&&row.excelRow<=12;}).map(function(row){return {excelRow:row.excelRow,cells:row.cells.slice(0,10)};}); rows=markHeaders(rows,1); return '<div class="section-title"><span></span>自营销售 · MTD</div>'+renderRows(rows); }

  function periodRows(section, sectionId){ var cfg=PERIOD_CONFIG[sectionId]; var period=activePeriod(sectionId); if(!cfg) return section.rows||[]; if(sectionId==='price_power_history'){ return pricePowerPeriodRows(section, period); } var idx=cfg.periods[period]||cfg.periods.YTD; var startHeader=Math.max(0,cfg.rowStart-2); var rows=(section.rows||[]).filter(function(row){return row.excelRow>=startHeader&&row.excelRow<=cfg.rowEnd;}).map(function(row){return {excelRow:row.excelRow,cells:idx.map(function(i){return row.cells[i]||{text:"",type:"blank"};})};}); return markHeaders(rows,2); }
  function pricePowerPeriodRows(section, period){ var rows=(section.rows||[]).filter(function(row){ if(row.excelRow===137||row.excelRow===138) return true; if(period==='YTD') return row.excelRow>=139&&row.excelRow<=140; return row.excelRow>=141&&row.excelRow<=152; }).map(function(row){return {excelRow:row.excelRow,cells:row.cells.slice(0,10)};}); return markHeaders(rows,2); }
  function renderPeriodFilter(sectionId){ var cfg=PERIOD_CONFIG[sectionId]; if(!cfg) return ""; var list=cfg.periodList||PERIODS; var current=activePeriod(sectionId); var html='<div class="filterbar">'; list.forEach(function(p){html+='<button class="filter-btn '+(current===p?'active':'')+'" data-section="'+sectionId+'" data-period="'+p+'">'+p+'</button>';}); html+='</div>'; return html; }

  function cloneCell(text){ return {text:text,raw:text,type:"text",unit:"",header:true}; }
  function renderPriceIndexMtd(section){ var rows=(section.rows||[]).filter(function(row){return row.excelRow>=43&&row.excelRow<=51;}).map(function(row){return {excelRow:row.excelRow,cells:row.cells.slice(0,14)};}); rows=markHeaders(rows,2); return '<div class="section-title"><span></span>外网价指 · MTD</div>'+renderRows(rows); }
  function renderSixHighMtd(section){
    var headers=["小组","款数目标","六高占比","天猫价指","天猫目标","vs天猫目标","抖音价指","抖音目标","vs抖音目标","折扣率","折扣目标","vs目标","动销率","动销目标","vs目标6.1%"];
    var rows=(section.rows||[]).filter(function(row){return row.excelRow>=80&&row.excelRow<=86;}).map(function(row){return {excelRow:row.excelRow,cells:row.cells.slice(0,15)};});
    return '<div class="section-title"><span></span>六高 · MTD</div>'+renderRows([headerRow(headers)].concat(rows));
  }
  function renderQualityMtd(section){
    var headers=["小组","已引进","未引进","暂不引进","总计","引入率（目标50%）"];
    var rows=(section.rows||[]).filter(function(row){return row.excelRow>=92&&row.excelRow<=98;}).map(function(row){return {excelRow:row.excelRow,cells:row.cells.slice(0,6)};});
    return '<div class="section-title"><span></span>优质款 · MTD</div>'+renderRows([headerRow(headers)].concat(rows));
  }
  function renderMtdWithoutTitle(section, title, firstRow, lastRow, cols){ var rows=(section.rows||[]).filter(function(row){return row.excelRow>=firstRow&&row.excelRow<=lastRow;}).map(function(row){return {excelRow:row.excelRow,cells:row.cells.slice(0,cols)};}); rows=markHeaders(rows,1); return '<div class="section-title"><span></span>'+escapeHtml(title)+'</div>'+renderRows(rows); }

  function renderPricePowerMtd(section){
    var rows=(section.rows||[]).filter(function(row){return row.excelRow>=131&&row.excelRow<=134;}).map(function(row){return {excelRow:row.excelRow,cells:row.cells.slice(0,9)};});
    rows=markHeaders(rows,2);
    return '<div class="section-title"><span></span>五星价格力 & 大爆款效率 · MTD</div>'+renderRows(rows);
  }
  function renderGrossProfit(section){
    var keepRows={27:true,34:true,35:true,36:true,37:true,38:true};
    var rows=(section.rows||[]).filter(function(row){return !!keepRows[row.excelRow];}).map(function(row){return {excelRow:row.excelRow,cells:row.cells.slice(0,9)};});
    rows=markHeaders(rows,1);
    var cutoff='毛利数据截止13号0点';
    var title='<div class="section-title gross-title"><span></span><div><b>毛利</b><em>'+escapeHtml(cutoff)+'</em></div></div>';
    return title+renderRows(rows);
  }

  function renderTableSection(section){
    if(!section) return "";
    if(section.id==='gross_profit') return renderGrossProfit(section);
    if(section.id==='price_index_mtd') return renderPriceIndexMtd(section);
    if(section.id==='six_high') return renderSixHighMtd(section);
    if(section.id==='quality_product_mtd') return renderQualityMtd(section);
    if(section.id==='machine_purchase_mtd') return renderMtdWithoutTitle(section,'机采 · MTD',114,118,8);
    if(section.id==='price_power_mtd') return renderPricePowerMtd(section);
    var hasFilter=!!PERIOD_CONFIG[section.id]; var baseTitle = section.title.replace(' · MTD / YTD / 历史月份',' · 历史月份').replace('YTD / 历史月份得分','历史月份得分').replace('YTD / 历史月份','历史月份'); var title=hasFilter ? baseTitle+' · '+activePeriod(section.id) : section.title; var rows=hasFilter ? periodRows(section, section.id) : (section.rows||[]); return '<div class="section-title"><span></span>'+escapeHtml(title)+'</div>'+renderPeriodFilter(section.id)+renderRows(rows); }

  function renderSalesPanel(){ var mtd=getSection('self_sales_mtd'), hist=getSection('self_sales_history'); return salesMtdTableSection(mtd)+renderTableSection(hist); }
  function discountRows(section, period){
    var cfg=PERIOD_CONFIG.internal_discount;
    var idx=cfg.periods[period] || cfg.periods.MTD;
    var startHeader=Math.max(0,cfg.rowStart-2);
    var rows=(section.rows||[]).filter(function(row){return row.excelRow>=startHeader&&row.excelRow<=cfg.rowEnd;}).map(function(row){return {excelRow:row.excelRow,cells:idx.map(function(i){return row.cells[i]||{text:'',type:'blank'};})};});
    return markHeaders(rows,2);
  }

  function renderDiscountPanel(){
    var section=getSection('internal_discount');
    if(!section) return '<div class="loading">暂无数据</div>';
    var historyPeriod=state.periods.internal_discount_history || '1月';
    var html='';
    html+='<div class="section-title"><span></span>内网价指数 · MTD</div>'+renderRows(discountRows(section,'MTD'));
    html+='<div class="section-title"><span></span>内网价指数 · YTD</div>'+renderRows(discountRows(section,'YTD'));
    html+='<div class="section-title"><span></span>内网价指数 · 历史月份 · '+historyPeriod+'</div>';
    html+='<div class="filterbar">';
    ['1月','2月','3月','4月','5月','6月'].forEach(function(p){ html+='<button class="filter-btn '+(historyPeriod===p?'active':'')+'" data-section="internal_discount_history" data-period="'+p+'">'+p+'</button>'; });
    html+='</div>'+renderRows(discountRows(section,historyPeriod));
    return html;
  }

  function renderGenericPanel(tab){ var html=""; tab.sectionIds.forEach(function(id){ var s=getSection(id); if(s) html+=renderTableSection(s); }); return html||'<div class="loading">暂无数据</div>'; }
  function renderTabs(){ var html='<div class="mobile-tabs">'; TABS.forEach(function(tab){html+='<button class="tab-btn '+(state.activeTab===tab.id?'active':'')+'" data-tab="'+tab.id+'">'+escapeHtml(tab.label)+'</button>';}); return html+'</div>'; }
  function renderDashboard(){ var data=state.data; if(!data) return; var meta=data.meta||{}; $navbarDate.textContent=meta.dataDate?'截止 '+meta.dataDate:'—'; if($periodToggle) $periodToggle.style.display='none'; var activeTab=TABS.find(function(t){return t.id===state.activeTab;})||TABS[0]; $modulesContainer.innerHTML=renderTabs()+'<main class="mobile-panel">'+(state.activeTab==='sales'?renderSalesPanel():(state.activeTab==='discount'?renderDiscountPanel():renderGenericPanel(activeTab)))+'</main>'; document.querySelectorAll('.tab-btn').forEach(function(btn){btn.onclick=function(){state.activeTab=btn.getAttribute('data-tab');renderDashboard();window.scrollTo(0,0);};}); document.querySelectorAll('.filter-btn').forEach(function(btn){btn.onclick=function(){state.periods[btn.getAttribute('data-section')]=btn.getAttribute('data-period');renderDashboard();};}); }

  function enterDashboard(){ $loginError.textContent=""; $loginPage.style.display="none"; $dashboard.classList.add("active"); loadData(); }
  function handleLogin(){ var pwd=$passwordInput.value.trim(); if(!pwd){$loginError.textContent='请输入密码';return;} $loginError.textContent='正在登录…'; loginByApi(pwd).then(function(){enterDashboard();}).catch(function(){ if(pwd===STATIC_PREVIEW_PASSWORD){state.token='static-preview';enterDashboard();return;} $loginError.textContent='密码错误'; $passwordInput.value=''; $passwordInput.focus(); }); }

  window.dashboardLogin=handleLogin; $loginBtn.onclick=handleLogin; $loginBtn.addEventListener('click',handleLogin); $passwordInput.addEventListener('keydown',function(e){if(e.key==='Enter')handleLogin();}); $logoutBtn.addEventListener('click',function(){state.token=null;state.data=null;$passwordInput.value='';$dashboard.classList.remove('active');$loginPage.style.display='flex';$loginError.textContent='';});
})();
