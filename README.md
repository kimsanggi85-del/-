[index.html](https://github.com/user-attachments/files/26478331/index.html)
<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>출하검사 주간보고서</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>

<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;600;700;900&family=JetBrains+Mono:wght@400;600&display=swap');
:root{
  --bg:#0b0e13;--bg2:#12161e;--bg3:#1a1f2b;--bg4:#222838;--bg5:#2a3148;
  --border:#252b3a;--border2:#343c50;
  --text:#e4e8f0;--text2:#8e96aa;--text3:#5c6478;
  --accent:#3b82f6;--accent2:#60a5fa;
  --green:#22c55e;--green-bg:rgba(34,197,94,.12);
  --red:#ef4444;--red-bg:rgba(239,68,68,.12);
  --yellow:#eab308;--yellow-bg:rgba(234,179,8,.12);
  --cyan:#06b6d4;--purple:#a78bfa;--orange:#f97316;--pink:#ec4899;
  --row-hover:#181d2a;--detail-bg:#0f1219;
}
*{margin:0;padding:0;box-sizing:border-box;}
body{font-family:'Noto Sans KR',sans-serif;background:var(--bg);color:var(--text);min-height:100vh;}
.app{max-width:1520px;margin:0 auto;padding:16px 20px;}

/* Header */
.header{display:flex;align-items:center;justify-content:space-between;padding:16px 0 20px;border-bottom:1px solid var(--border);margin-bottom:20px;}
.header h1{font-size:22px;font-weight:800;background:linear-gradient(135deg,var(--accent2),var(--cyan));-webkit-background-clip:text;-webkit-text-fill-color:transparent;}
.header .sub{font-size:12px;color:var(--text3);margin-top:2px;}
.hdr-right{text-align:right;font-size:11px;color:var(--text3);}

/* Upload */
.upload-bar{display:flex;gap:8px;margin-bottom:16px;flex-wrap:wrap;align-items:center;}
.upload-slot{background:var(--bg2);border:1px solid var(--border);border-radius:8px;padding:8px 14px;display:flex;align-items:center;gap:8px;cursor:pointer;transition:.2s;font-size:12px;white-space:nowrap;}
.upload-slot:hover{border-color:var(--accent);}
.upload-slot.loaded{border-color:var(--green);}
.upload-slot.loaded .sl-dot{background:var(--green);}
.sl-dot{width:8px;height:8px;border-radius:50%;background:var(--text3);flex-shrink:0;}
.sl-name{font-weight:600;}
.sl-info{color:var(--text3);font-size:11px;}
input[type="file"]{display:none;}

/* Filter */
.filter-bar{display:flex;align-items:center;gap:8px;background:var(--bg2);border:1px solid var(--border);border-radius:8px;padding:8px 12px;margin-bottom:16px;flex-wrap:wrap;font-size:12px;}
.filter-bar label{color:var(--text2);font-weight:500;}
.filter-bar input[type="date"]{background:var(--bg3);border:1px solid var(--border);border-radius:5px;color:var(--text);padding:4px 8px;font-size:12px;font-family:'Noto Sans KR';}
.filter-bar input[type="date"]::-webkit-calendar-picker-indicator{filter:invert(1);}
.btn{background:var(--accent);color:#fff;border:none;border-radius:5px;padding:5px 12px;font-size:12px;font-weight:600;cursor:pointer;font-family:'Noto Sans KR';transition:.2s;}
.btn:hover{background:#2563eb;}
.btn-sm{background:var(--bg3);border:1px solid var(--border);color:var(--text2);font-weight:500;}
.btn-sm:hover{border-color:var(--accent);color:var(--accent);background:var(--bg3);}
.btn-sm.active{background:var(--accent);color:#fff;border-color:var(--accent);}
.spacer{flex:1;}

/* Stats */
.stats-row{display:grid;grid-template-columns:repeat(5,1fr);gap:10px;margin-bottom:16px;}
.stat-card{background:var(--bg2);border:1px solid var(--border);border-radius:8px;padding:12px 14px;}
.stat-label{font-size:10px;color:var(--text3);font-weight:600;text-transform:uppercase;letter-spacing:.5px;}
.stat-value{font-size:24px;font-weight:700;font-family:'JetBrains Mono',monospace;margin-top:2px;}
.stat-sub{font-size:10px;color:var(--text3);margin-top:1px;}

/* Tabs */
.tabs{display:flex;gap:2px;background:var(--bg2);border-radius:8px;padding:3px;margin-bottom:16px;overflow-x:auto;}
.tab{padding:6px 16px;font-size:12px;font-weight:500;border-radius:6px;cursor:pointer;color:var(--text3);white-space:nowrap;transition:.2s;border:none;background:none;font-family:'Noto Sans KR';}
.tab:hover{color:var(--text2);}
.tab.active{background:var(--accent);color:#fff;}

/* Charts */
.charts-grid{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:16px;}
.chart-box{background:var(--bg2);border:1px solid var(--border);border-radius:10px;padding:16px;}
.chart-box.full{grid-column:1/-1;}
.chart-title{font-size:13px;font-weight:600;margin-bottom:12px;display:flex;align-items:center;gap:6px;}
.chart-title .dot{width:10px;height:10px;border-radius:3px;}
.chart-canvas{width:100%;height:260px;position:relative;}
.chart-canvas canvas{width:100%!important;height:100%!important;}

/* Table */
.table-wrap{background:var(--bg2);border:1px solid var(--border);border-radius:10px;overflow:hidden;margin-bottom:16px;}
.table-scroll{overflow-x:auto;max-height:700px;overflow-y:auto;}
table{width:100%;border-collapse:collapse;font-size:11px;}
thead th{background:var(--bg3);padding:8px 10px;text-align:center;font-weight:600;font-size:10px;color:var(--text2);border-bottom:1px solid var(--border);position:sticky;top:0;white-space:nowrap;text-transform:uppercase;letter-spacing:.3px;z-index:2;}
tbody td{padding:6px 10px;border-bottom:1px solid var(--border);text-align:right;font-family:'JetBrains Mono',monospace;font-size:11px;white-space:nowrap;}
tbody td.tl{text-align:left;font-family:'Noto Sans KR';}
tbody tr:hover{background:var(--row-hover);}

/* Expandable Rows */
.sum-row{cursor:pointer;transition:background .15s;}
.sum-row:hover{background:var(--bg4)!important;}
.sum-row td:first-child{position:relative;padding-left:24px;}
.sum-row .arrow{position:absolute;left:8px;top:50%;transform:translateY(-50%);font-size:10px;color:var(--text3);transition:transform .2s;}
.sum-row.open .arrow{transform:translateY(-50%) rotate(90deg);color:var(--accent);}
.sum-row.open{background:var(--bg3);}

.detail-row{display:none;}
.detail-row.show{display:table-row;}
.detail-row td{padding:0;border-bottom:none;}
.detail-panel{background:var(--detail-bg);border-left:3px solid var(--accent);padding:0;}
.detail-panel table{margin:0;font-size:10px;}
.detail-panel thead th{background:var(--bg4);font-size:9px;padding:5px 8px;position:static;border-bottom:1px solid var(--border2);}
.detail-panel tbody td{padding:4px 8px;border-bottom:1px solid rgba(37,43,58,.5);font-size:10px;}
.detail-panel tbody tr:hover{background:rgba(59,130,246,.06);}
.detail-panel .dp-header{display:flex;align-items:center;justify-content:space-between;padding:8px 12px;border-bottom:1px solid var(--border);}
.detail-panel .dp-title{font-size:11px;font-weight:600;font-family:'Noto Sans KR';color:var(--accent2);}
.detail-panel .dp-info{font-size:10px;color:var(--text3);font-family:'Noto Sans KR';}

/* Mini chart in detail panel */
.dp-mini-charts{display:flex;gap:12px;padding:8px 12px;border-bottom:1px solid var(--border);flex-wrap:wrap;}
.dp-mini{flex:1;min-width:120px;}
.dp-mini-label{font-size:9px;color:var(--text3);font-family:'Noto Sans KR';margin-bottom:4px;font-weight:500;}
.dp-mini-bar{display:flex;gap:1px;height:20px;border-radius:3px;overflow:hidden;background:var(--bg3);}
.dp-mini-seg{height:100%;min-width:2px;position:relative;}
.dp-mini-seg:hover{opacity:.8;}
.dp-mini-legend{display:flex;gap:8px;flex-wrap:wrap;margin-top:4px;}
.dp-mini-legend span{font-size:8px;color:var(--text3);display:flex;align-items:center;gap:2px;}
.dp-mini-legend .ldot{width:6px;height:6px;border-radius:2px;}

.sub-row{background:var(--bg4)!important;font-weight:700;cursor:pointer;}
.sub-row:hover{background:var(--bg5)!important;}
.sub-row td{border-top:2px solid var(--border2);}

.ppm-g{color:var(--green);background:var(--green-bg);border-radius:3px;padding:1px 5px;font-size:11px;}
.ppm-w{color:var(--yellow);background:var(--yellow-bg);border-radius:3px;padding:1px 5px;font-size:11px;}
.ppm-b{color:var(--red);background:var(--red-bg);border-radius:3px;padding:1px 5px;font-size:11px;}
.ctag{display:inline-block;padding:1px 6px;border-radius:3px;font-size:10px;font-weight:600;font-family:'Noto Sans KR';}
.ctag-SKF{background:#1e3a5f;color:#60a5fa;}
.ctag-SKF_SBC{background:#1e3a3f;color:#5eead4;}
.ctag-일진{background:#3a1e5f;color:#a78bfa;}
.ctag-진양오일씰{background:#3a2e1e;color:#fbbf24;}

.def-hi{color:var(--red);font-weight:600;}
.def-lo{color:var(--text3);}
.empty-state{text-align:center;padding:60px 20px;color:var(--text3);}

/* Customer Report Cards */
.rpt-week-sel{display:flex;align-items:center;gap:10px;margin-bottom:16px;flex-wrap:wrap;}
.rpt-week-sel select{background:var(--bg3);border:1px solid var(--border2);border-radius:6px;color:var(--text);padding:6px 12px;font-size:13px;font-family:'Noto Sans KR';cursor:pointer;min-width:260px;}
.rpt-week-sel select:focus{border-color:var(--accent);outline:none;}
.rpt-week-nav{display:flex;gap:4px;}
.rpt-week-nav button{background:var(--bg3);border:1px solid var(--border);border-radius:5px;color:var(--text2);width:32px;height:32px;cursor:pointer;font-size:14px;display:flex;align-items:center;justify-content:center;transition:.2s;}
.rpt-week-nav button:hover{border-color:var(--accent);color:var(--accent);}
.rpt-week-nav button:disabled{opacity:.3;cursor:default;}

.rpt-cards{display:flex;flex-direction:column;gap:16px;}
.rpt-card{background:var(--bg2);border:1px solid var(--border);border-radius:12px;overflow:hidden;}
.rpt-card-hdr{display:flex;align-items:center;justify-content:space-between;padding:14px 18px;border-bottom:1px solid var(--border);background:var(--bg3);}
.rpt-card-hdr .rpt-cname{font-size:15px;font-weight:700;display:flex;align-items:center;gap:10px;}
.rpt-card-hdr .rpt-cname .cdot{width:12px;height:12px;border-radius:4px;}
.rpt-card-hdr .rpt-period{font-size:11px;color:var(--text3);}

.rpt-stats{display:grid;grid-template-columns:repeat(5,1fr);gap:1px;background:var(--border);border-bottom:1px solid var(--border);}
.rpt-stat{background:var(--bg2);padding:12px 14px;text-align:center;}
.rpt-stat .rs-label{font-size:9px;color:var(--text3);text-transform:uppercase;letter-spacing:.5px;font-weight:600;}
.rpt-stat .rs-val{font-size:20px;font-weight:700;font-family:'JetBrains Mono',monospace;margin-top:2px;}
.rpt-stat .rs-delta{font-size:10px;margin-top:2px;font-family:'JetBrains Mono',monospace;}
.rs-up{color:var(--red);}
.rs-down{color:var(--green);}
.rs-flat{color:var(--text3);}

.rpt-body{padding:14px 18px;}
.rpt-section{margin-bottom:14px;}
.rpt-section:last-child{margin-bottom:0;}
.rpt-sec-title{font-size:12px;font-weight:600;color:var(--text2);margin-bottom:8px;display:flex;align-items:center;gap:6px;}
.rpt-sec-title .dot{width:8px;height:8px;border-radius:3px;}

.rpt-def-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(110px,1fr));gap:6px;}
.rpt-def-item{background:var(--bg3);border-radius:6px;padding:8px 10px;text-align:center;border:1px solid var(--border);}
.rpt-def-item .rdi-name{font-size:10px;color:var(--text2);font-weight:500;}
.rpt-def-item .rdi-val{font-size:16px;font-weight:700;font-family:'JetBrains Mono',monospace;margin-top:1px;}
.rpt-def-item .rdi-pct{font-size:9px;color:var(--text3);font-family:'JetBrains Mono',monospace;}
.rpt-def-item.zero{opacity:.4;}

.rpt-chart-row{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:14px;}
.rpt-chart-box{background:var(--bg3);border:1px solid var(--border);border-radius:8px;padding:12px;}
.rpt-chart-box .rch-title{font-size:11px;font-weight:600;color:var(--text2);margin-bottom:8px;}
.rpt-chart-box canvas{width:100%!important;height:180px!important;}

.rpt-no-data{text-align:center;padding:40px;color:var(--text3);font-size:13px;}
@media(max-width:900px){.rpt-stats{grid-template-columns:repeat(3,1fr);}.rpt-chart-row{grid-template-columns:1fr;}.rpt-def-grid{grid-template-columns:repeat(auto-fill,minmax(90px,1fr));}}

@media(max-width:900px){.stats-row{grid-template-columns:repeat(2,1fr);}.charts-grid{grid-template-columns:1fr;}}

/* ===== Print Report Page ===== */
.pr-page{background:#fff;color:#1a1a2e;border-radius:10px;padding:0;overflow:hidden;border:1px solid var(--border);}
.pr-topbar{display:flex;align-items:center;gap:10px;padding:10px 16px;background:#f0f4f8;border-bottom:1px solid #d0d5dd;}
.pr-topbar .pr-btn{background:#2563eb;color:#fff;border:none;border-radius:5px;padding:6px 14px;font-size:12px;font-weight:600;cursor:pointer;font-family:'Noto Sans KR';display:flex;align-items:center;gap:4px;}
.pr-topbar .pr-btn:hover{background:#1d4ed8;}
.pr-topbar .pr-btn-outline{background:#fff;border:1px solid #d0d5dd;color:#374151;font-weight:500;}
.pr-topbar .pr-btn-outline:hover{border-color:#2563eb;color:#2563eb;}
.pr-topbar .pr-info{flex:1;text-align:right;font-size:11px;color:#6b7280;}

.pr-inner{padding:14px 16px;}
.pr-title-row{display:flex;align-items:flex-end;justify-content:space-between;margin-bottom:10px;}
.pr-title{font-size:16px;font-weight:800;color:#1e293b;}
.pr-sub{font-size:10px;color:#94a3b8;}

/* Chart */
.pr-chart-wrap{background:#fff;border:1px solid #e2e8f0;border-radius:8px;padding:12px;margin-bottom:12px;}
.pr-chart-title{font-size:12px;font-weight:700;color:#1e293b;text-align:center;margin-bottom:8px;}
.pr-chart-canvas{width:100%;height:220px;position:relative;}
.pr-chart-canvas canvas{width:100%!important;height:100%!important;}

/* Weekly table */
.pr-wk-table{border:1px solid #e2e8f0;border-radius:6px;overflow:hidden;margin-bottom:12px;font-size:10px;}
.pr-wk-table table{width:100%;border-collapse:collapse;}
.pr-wk-table th{background:#f1f5f9;padding:4px 6px;text-align:center;font-weight:600;color:#475569;border:1px solid #e2e8f0;font-size:9px;white-space:nowrap;}
.pr-wk-table td{padding:3px 5px;text-align:right;border:1px solid #e2e8f0;font-family:'JetBrains Mono',monospace;font-size:9px;white-space:nowrap;color:#334155;}
.pr-wk-table td.rl{text-align:left;font-family:'Noto Sans KR';font-weight:600;background:#f8fafc;color:#1e293b;}

/* Bottom sections */
.pr-bottom{display:grid;grid-template-columns:1fr 1fr;gap:12px;}
.pr-section{border:1px solid #e2e8f0;border-radius:6px;overflow:hidden;}
.pr-sec-hdr{background:#f1f5f9;padding:6px 10px;font-size:11px;font-weight:700;color:#1e293b;border-bottom:1px solid #e2e8f0;}
.pr-sec-body{padding:8px 10px;}

/* Customer summary mini-tables */
.pr-cust-block{margin-bottom:8px;border:1px solid #e2e8f0;border-radius:4px;overflow:hidden;}
.pr-cust-block:last-child{margin-bottom:0;}
.pr-cust-hdr{display:flex;align-items:center;justify-content:space-between;padding:4px 8px;background:#f8fafc;border-bottom:1px solid #e2e8f0;}
.pr-cust-name{font-size:10px;font-weight:700;color:#1e293b;display:flex;align-items:center;gap:4px;}
.pr-cust-name .pc-dot{width:8px;height:8px;border-radius:2px;display:inline-block;}
.pr-cust-target{font-size:9px;color:#6b7280;font-weight:600;}
.pr-cust-target b{color:#2563eb;}
.pr-mini-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:1px;background:#e2e8f0;}
.pr-mini-cell{background:#fff;padding:5px 6px;text-align:center;}
.pr-mini-cell .pm-label{font-size:8px;color:#94a3b8;text-transform:uppercase;font-weight:600;}
.pr-mini-cell .pm-val{font-size:13px;font-weight:700;font-family:'JetBrains Mono',monospace;color:#1e293b;}
.pr-mini-cell .pm-val.pm-red{color:#dc2626;}
.pr-mini-cell .pm-val.pm-grn{color:#16a34a;}
.pr-mini-cell .pm-val.pm-warn{color:#d97706;}

/* Part defect table */
.pr-part-table{font-size:9px;}
.pr-part-table table{width:100%;border-collapse:collapse;}
.pr-part-table th{background:#f1f5f9;padding:3px 4px;text-align:center;font-weight:600;color:#475569;border:1px solid #e2e8f0;font-size:8px;white-space:nowrap;}
.pr-part-table td{padding:2px 4px;text-align:right;border:1px solid #e2e8f0;font-family:'JetBrains Mono',monospace;font-size:9px;color:#334155;}
.pr-part-table td.rl{text-align:left;font-family:'Noto Sans KR';font-weight:500;max-width:100px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.pr-part-table tr.pr-total{background:#f1f5f9;font-weight:700;}
.pr-part-table .pr-hi{color:#dc2626;font-weight:600;}

/* PPM badge for print */
.pr-ppm{display:inline-block;padding:1px 4px;border-radius:2px;font-size:9px;font-weight:700;font-family:'JetBrains Mono',monospace;}
.pr-ppm-g{background:#dcfce7;color:#16a34a;}
.pr-ppm-w{background:#fef9c3;color:#a16207;}
.pr-ppm-b{background:#fee2e2;color:#dc2626;}

/* Week selector for print tab */
.pr-week-sel{display:flex;align-items:center;gap:10px;margin-bottom:14px;flex-wrap:wrap;}
.pr-week-sel select{background:var(--bg3);border:1px solid var(--border2);border-radius:6px;color:var(--text);padding:6px 12px;font-size:13px;font-family:'Noto Sans KR';cursor:pointer;min-width:260px;}

@media print{
  body{background:#fff!important;color:#000!important;-webkit-print-color-adjust:exact;print-color-adjust:exact;}
  .app{padding:0!important;max-width:100%!important;}
  .header,.upload-bar,.filter-bar,.tabs,.stats-row{display:none!important;}
  .pr-page{border:none!important;border-radius:0!important;}
  .pr-topbar{display:none!important;}
  .pr-week-sel{display:none!important;}
  .pr-inner{padding:8px!important;}
  .pr-chart-canvas{height:200px!important;}
  @page{size:A4 landscape;margin:8mm;}
}
</style>
</head>
<body>
<div class="app" id="app"></div>
<script>
const CUSTS=[
  {key:'SKF',name:'SKF',hint:'SKF_.',color:'#3b82f6',bg:'rgba(59,130,246,'},
  {key:'SKF_SBC',name:'SKF SBC Raceway',hint:'SBC_Raceway',color:'#06b6d4',bg:'rgba(6,182,212,'},
  {key:'일진',name:'일진',hint:'일진',color:'#a78bfa',bg:'rgba(167,139,250,'},
  {key:'진양오일씰',name:'진양오일씰',hint:'진양오일씰',color:'#fbbf24',bg:'rgba(251,191,36,'},
];
const CMAP={
  'SKF':{d:0,p:1,pn:2,ins:3,q:5,td:12,df:{6:'찍힘',7:'휨',8:'기스',9:'Burr',10:'이물',11:'기타'}},
  'SKF_SBC':{d:0,p:1,pn:2,ins:3,q:5,td:null,df:{6:'높이',7:'Burr',8:'크랙',9:'찍힘',10:'눌림',11:'얼룩',12:'휨',13:'기타'}},
  '일진':{d:0,p:1,pn:2,ins:3,q:5,td:12,df:{6:'찍힘',7:'휨',8:'기스',9:'도포',10:'터짐',11:'기타'}},
  '진양오일씰':{d:0,p:1,pn:2,ins:3,q:5,td:12,df:{6:'찍힘',7:'휨',8:'기스',9:'Burr',10:'이물',11:'기타'}},
};
const DC={'찍힘':'#ef4444','휨':'#f97316','기스':'#eab308','Burr':'#22c55e','이물':'#06b6d4','기타':'#6b7280','높이':'#ec4899','크랙':'#a78bfa','눌림':'#f43f5e','얼룩':'#14b8a6','도포':'#8b5cf6','터짐':'#d946ef'};

function pD(v){const s=String(v).replace(/\D/g,'');if(s.length!==8)return null;const dt=new Date(+s.slice(0,4),+s.slice(4,6)-1,+s.slice(6,8));return isNaN(dt)?null:dt;}
function isoWk(d){const t=new Date(d);t.setHours(0,0,0,0);t.setDate(t.getDate()+3-(t.getDay()+6)%7);const w1=new Date(t.getFullYear(),0,4);return{y:t.getFullYear(),w:1+Math.round(((t-w1)/864e5-3+(w1.getDay()+6)%7)/7)};}
function wkMon(d){const t=new Date(d);t.setDate(t.getDate()-((t.getDay()+6)%7));return t;}
function F(n){return n.toLocaleString('ko-KR');}
function fp(n){return n.toFixed(1);}
function fd(d){return`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;}
function fsd(d){return`${d.getMonth()+1}/${String(d.getDate()).padStart(2,'0')}`;}
function ym(d){return`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;}
function ymL(s){const[y,m]=s.split('-');return`${y}년 ${+m}월`;}
function pC(v){return v<=1000?'ppm-g':v<=3000?'ppm-w':'ppm-b';}

// ===== Mixed Period: 2026은 상세, 나머지는 년도 =====
const DETAIL_YEAR=2026;
function mixedWkKey(r){return r.dt.getFullYear()===DETAIL_YEAR?r.wk:String(r.dt.getFullYear());}
function mixedMonKey(r){return r.dt.getFullYear()===DETAIL_YEAR?r.mon:String(r.dt.getFullYear());}
function mixedWkLabel(key,data){
  if(/^\d{4}-W\d{2}$/.test(key)){const r=data.find(x=>x.wk===key);return r?r.wkL:key;}
  return key+'년';
}
function mixedMonLabel(key){
  if(/^\d{4}-\d{2}$/.test(key)) return ymL(key);
  return key+'년';
}
function mixedWkShortLabel(key,data){
  if(/^\d{4}-W\d{2}$/.test(key)){const r=data.find(x=>x.wk===key);if(r){const l=r.wkL;return l.replace(/\d{4}년\s*/,'');}return key.split('-W')[1];}
  return key+'년';
}

function parseDB(wb,ck){
  const ws=wb.Sheets['DB'];if(!ws)return[];
  const raw=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});
  const cm=CMAP[ck],rows=[];
  for(let i=4;i<raw.length;i++){
    const r=raw[i];if(!r||!r[cm.d])continue;
    const dt=pD(r[cm.d]);if(!dt)continue;
    const q=parseInt(r[cm.q])||0;if(q<=0)continue;
    const dv={};let td=0;
    for(const[ci,dn]of Object.entries(cm.df)){const v=parseInt(r[+ci])||0;dv[dn]=v;td+=v;}
    if(cm.td!==null)td=parseInt(r[cm.td])||0;
    const iso=isoWk(dt),mon=wkMon(dt),sun=new Date(mon);sun.setDate(sun.getDate()+6);
    rows.push({c:ck,dt,ds:fd(dt),part:String(r[cm.p]||'').trim(),pn:String(r[cm.pn]||'').trim(),
      ins:String(r[cm.ins]||'').trim(),q,dv,td,ppm:q>0?td/q*1e6:0,
      wk:`${iso.y}-W${String(iso.w).padStart(2,'0')}`,
      wkL:`${iso.y}년 ${String(iso.w).padStart(2,'0')}주 (${fsd(mon)}~${fsd(sun)})`,
      mon:ym(dt)});
  }
  return rows;
}

let S={files:{},all:[],flt:[],sd:'',ed:'',tab:'dashboard',cust:'all',charts:{},expanded:new Set(),rptWeek:''};
function loadF(ck,file){const rd=new FileReader();rd.onload=async e=>{S.files[ck]=XLSX.read(e.target.result,{type:'array'});proc();render();};rd.readAsArrayBuffer(file);}
function proc(){S.all=[];for(const[k,wb]of Object.entries(S.files))S.all.push(...parseDB(wb,k));S.all.sort((a,b)=>a.dt-b.dt);filt();}
function filt(){let d=S.all;if(S.sd){const s=new Date(S.sd);d=d.filter(r=>r.dt>=s);}if(S.ed){const e=new Date(S.ed);e.setHours(23,59,59);d=d.filter(r=>r.dt<=e);}S.flt=d;}

function agg(data,keyFn){
  const m=new Map();
  for(const r of data){const k=keyFn(r)+'|'+r.c;if(!m.has(k))m.set(k,{key:keyFn(r),label:r.wkL||ymL(r.mon),c:r.c,n:0,q:0,td:0,dv:{},rows:[]});const g=m.get(k);g.n++;g.q+=r.q;g.td+=r.td;g.rows.push(r);for(const[dn,v]of Object.entries(r.dv))g.dv[dn]=(g.dv[dn]||0)+v;}
  const res=[...m.values()];res.forEach(g=>g.ppm=g.q>0?g.td/g.q*1e6:0);return res;
}
function aggTot(data,keyFn){
  const m=new Map();
  for(const r of data){const k=keyFn(r);if(!m.has(k))m.set(k,{key:k,label:r.wkL||ymL(r.mon),q:0,td:0,dv:{},rows:[]});const g=m.get(k);g.q+=r.q;g.td+=r.td;g.rows.push(r);for(const[dn,v]of Object.entries(r.dv))g.dv[dn]=(g.dv[dn]||0)+v;}
  const res=[...m.values()];res.forEach(g=>g.ppm=g.q>0?g.td/g.q*1e6:0);res.sort((a,b)=>a.key<b.key?-1:1);return res;
}
function topParts(data,n=20){const m=new Map();for(const r of data){const k=r.c+'|'+r.part;if(!m.has(k))m.set(k,{c:r.c,part:r.part,pn:r.pn,n:0,q:0,td:0});const g=m.get(k);g.n++;g.q+=r.q;g.td+=r.td;}const res=[...m.values()];res.forEach(g=>g.ppm=g.q>0?g.td/g.q*1e6:0);res.sort((a,b)=>b.ppm-a.ppm);return res.slice(0,n);}

// Mixed period aggregation (2026=detail, others=yearly)
function aggMixed(data,keyFn,labelFn){
  const m=new Map();
  for(const r of data){const k=keyFn(r)+'|'+r.c;if(!m.has(k))m.set(k,{key:keyFn(r),label:labelFn(keyFn(r),data),c:r.c,n:0,q:0,td:0,dv:{},rows:[]});const g=m.get(k);g.n++;g.q+=r.q;g.td+=r.td;g.rows.push(r);for(const[dn,v]of Object.entries(r.dv))g.dv[dn]=(g.dv[dn]||0)+v;}
  const res=[...m.values()];res.forEach(g=>g.ppm=g.q>0?g.td/g.q*1e6:0);return res;
}
function aggTotMixed(data,keyFn,labelFn){
  const m=new Map();
  for(const r of data){const k=keyFn(r);if(!m.has(k))m.set(k,{key:k,label:labelFn(k,data),q:0,td:0,dv:{},rows:[]});const g=m.get(k);g.q+=r.q;g.td+=r.td;g.rows.push(r);for(const[dn,v]of Object.entries(r.dv))g.dv[dn]=(g.dv[dn]||0)+v;}
  const res=[...m.values()];res.forEach(g=>g.ppm=g.q>0?g.td/g.q*1e6:0);res.sort((a,b)=>a.key<b.key?-1:1);return res;
}

function destroyCharts(){for(const c of Object.values(S.charts))c.destroy();S.charts={};}
Chart.defaults.color='#8e96aa';Chart.defaults.borderColor='#252b3a';Chart.defaults.font.family="'Noto Sans KR',sans-serif";Chart.defaults.font.size=11;
Chart.defaults.plugins.legend.labels.boxWidth=12;Chart.defaults.plugins.legend.labels.padding=12;
Chart.defaults.plugins.tooltip.backgroundColor='#1a1f2b';Chart.defaults.plugins.tooltip.borderColor='#343c50';Chart.defaults.plugins.tooltip.borderWidth=1;
Chart.defaults.plugins.tooltip.bodyFont={family:"'JetBrains Mono',monospace",size:12};Chart.defaults.plugins.tooltip.padding=10;Chart.defaults.plugins.tooltip.cornerRadius=6;

function mkChart(id,cfg){const el=document.getElementById(id);if(!el)return;if(S.charts[id])S.charts[id].destroy();S.charts[id]=new Chart(el,cfg);}

// ===== Detail Panel Builder =====
function buildDetailPanel(rows, custKey){
  const dns = Object.values(CMAP[custKey]?.df || {});
  // Mini defect bar
  const defTotals = {};
  let defSum = 0;
  for(const r of rows) for(const[dn,v] of Object.entries(r.dv)){ defTotals[dn]=(defTotals[dn]||0)+v; defSum+=v; }
  const defArr = Object.entries(defTotals).filter(([,v])=>v>0).sort((a,b)=>b[1]-a[1]);

  let h = `<div class="detail-panel"><div class="dp-header">
    <div class="dp-title">📋 ${custKey} 상세 내역 (${rows.length}건)</div>
    <div class="dp-info">검사수량 ${F(rows.reduce((s,r)=>s+r.q,0))} · 불량 ${F(rows.reduce((s,r)=>s+r.td,0))}</div></div>`;

  // Mini defect chart
  if(defArr.length > 0){
    h += `<div class="dp-mini-charts"><div class="dp-mini"><div class="dp-mini-label">불량유형 구성</div><div class="dp-mini-bar">`;
    for(const[dn,v] of defArr){
      const pct = defSum>0 ? v/defSum*100 : 0;
      h += `<div class="dp-mini-seg" style="width:${Math.max(pct,1)}%;background:${DC[dn]||'#6b7280'}" title="${dn}: ${F(v)} (${pct.toFixed(1)}%)"></div>`;
    }
    h += `</div><div class="dp-mini-legend">`;
    for(const[dn,v] of defArr.slice(0,6)){
      h += `<span><span class="ldot" style="background:${DC[dn]||'#6b7280'}"></span>${dn} ${F(v)}</span>`;
    }
    h += `</div></div></div>`;
  }

  // Detail table
  h += `<table><thead><tr><th>일자</th><th>품번</th><th>품명</th><th>검사자</th><th>검사수</th>`;
  for(const d of dns) h += `<th>${d}</th>`;
  h += `<th>불량계</th><th>PPM</th></tr></thead><tbody>`;
  for(const r of rows){
    h += `<tr><td>${r.ds}</td><td style="font-size:10px">${r.part}</td><td class="tl" style="font-size:10px">${r.pn}</td><td class="tl">${r.ins}</td><td>${F(r.q)}</td>`;
    for(const d of dns){ const v=r.dv[d]||0; h+=`<td class="${v>0?'def-hi':'def-lo'}">${v}</td>`; }
    h += `<td style="font-weight:600;${r.td>0?'color:var(--red)':''}">${r.td}</td><td><span class="${pC(r.ppm)}">${fp(r.ppm)}</span></td></tr>`;
  }
  h += `</tbody></table></div>`;
  return h;
}

// ===== Toggle expand =====
function toggleRow(id){
  if(S.expanded.has(id)) S.expanded.delete(id);
  else S.expanded.add(id);
  // DOM toggle for performance (no full re-render)
  const sr = document.querySelector(`[data-rid="${id}"]`);
  const dr = document.querySelector(`[data-did="${id}"]`);
  if(sr && dr){
    sr.classList.toggle('open');
    dr.classList.toggle('show');
  }
}

// ===== RENDER =====
function render(){
  destroyCharts();
  const app=document.getElementById('app');
  const data=S.flt;
  let h='';

  h+=`<div class="header"><div><h1>출하검사 주간보고서</h1><div class="sub">품질관리팀 · 출하검사 데이터 분석 시스템</div></div>
    <div class="hdr-right">${data.length>0?`${F(data.length)}건 · ${data[0].ds} ~ ${data[data.length-1].ds}`:'데이터 없음'}</div></div>`;

  h+=`<div class="upload-bar">`;
  for(const c of CUSTS){const ld=!!S.files[c.key];const fd2=ld?parseDB(S.files[c.key],c.key):[];
    h+=`<div class="upload-slot ${ld?'loaded':''}" onclick="document.getElementById('f-${c.key}').click()">
      <div class="sl-dot"></div><span class="sl-name">${c.name}</span><span class="sl-info">${ld?F(fd2.length)+'건':'파일선택'}</span>
      <input type="file" id="f-${c.key}" accept=".xls,.xlsx" onchange="handleF('${c.key}',this)"></div>`;}
  h+=`</div>`;

  if(!data.length){h+=`<div class="empty-state"><div style="font-size:40px;margin-bottom:12px;opacity:.3">📊</div>
    <div style="font-size:14px;margin-bottom:12px">출하검사일보 xls 파일을 업로드해주세요</div>
  </div>`;app.innerHTML=h;return;}

  h+=`<div class="filter-bar"><label>기간</label>
    <input type="date" value="${S.sd}" onchange="sf('sd',this.value)"><span style="color:var(--text3)">~</span>
    <input type="date" value="${S.ed}" onchange="sf('ed',this.value)">
    <button class="btn btn-sm" onclick="qr('1m')">1개월</button><button class="btn btn-sm" onclick="qr('3m')">3개월</button>
    <button class="btn btn-sm" onclick="qr('6m')">6개월</button><button class="btn btn-sm" onclick="qr('all')">전체</button>
    <div class="spacer"></div><button class="btn" onclick="expCSV()">📥 CSV</button><button class="btn btn-sm" onclick="window.print()">🖨️ 인쇄</button></div>`;

  const tQ=data.reduce((s,r)=>s+r.q,0),tD=data.reduce((s,r)=>s+r.td,0),tP=tQ>0?tD/tQ*1e6:0;
  const wSet=new Set(data.map(r=>r.wk)),cSet=new Set(data.map(r=>r.c));
  h+=`<div class="stats-row">
    <div class="stat-card"><div class="stat-label">검사건수</div><div class="stat-value">${F(data.length)}</div><div class="stat-sub">${wSet.size}주 · ${cSet.size}거래처</div></div>
    <div class="stat-card"><div class="stat-label">검사수량</div><div class="stat-value">${F(tQ)}</div></div>
    <div class="stat-card"><div class="stat-label">불량수량</div><div class="stat-value" style="color:${tD>0?'var(--red)':'var(--green)'}">${F(tD)}</div></div>
    <div class="stat-card"><div class="stat-label">종합 PPM</div><div class="stat-value"><span class="${pC(tP)}">${fp(tP)}</span></div></div>
    <div class="stat-card"><div class="stat-label">불량율</div><div class="stat-value">${tQ>0?(tD/tQ*100).toFixed(3):0}%</div></div></div>`;

  const tabs=[{id:'dashboard',l:'📊 대시보드'},{id:'weekly',l:'📅 주간 요약'},{id:'trend',l:'📈 PPM 추이'},{id:'monthly',l:'📆 월별 요약'},{id:'custrpt',l:'🏭 업체별 주간보고'},{id:'print',l:'🖨️ 인쇄보고서'},{id:'detail',l:'📝 상세 데이터'},{id:'top',l:'🔺 불량 TOP'}];
  h+=`<div class="tabs">`;for(const t of tabs)h+=`<button class="tab ${S.tab===t.id?'active':''}" onclick="st('${t.id}')">${t.l}</button>`;h+=`</div>`;

  if(S.tab==='dashboard') h+=rDash(data);
  else if(S.tab==='weekly') h+=rWeekly(data);
  else if(S.tab==='trend') h+=rTrend(data);
  else if(S.tab==='monthly') h+=rMonthly(data);
  else if(S.tab==='custrpt') h+=rCustReport(data);
  else if(S.tab==='print') h+=rPrintReport(data);
  else if(S.tab==='detail') h+=rDetail(data);
  else if(S.tab==='top') h+=rTop(data);

  app.innerHTML=h;
  if(S.tab==='dashboard') buildCharts(data);
  if(S.tab==='trend') buildTrendChart(data);
  if(S.tab==='custrpt') buildCustRptCharts(data);
  if(S.tab==='print') buildPrintCharts(data);
}

// ===== Weekly with expand (2026=주간, 나머지=연도) =====
function rWeekly(data){
  const grouped = aggMixed(data, mixedWkKey, mixedWkLabel);
  grouped.sort((a,b)=>a.key<b.key?-1:a.key>b.key?1:a.c<b.c?-1:1);
  const periods=[...new Set(grouped.map(g=>g.key))].sort();

  let h=`<div style="font-size:11px;color:var(--text3);margin-bottom:8px">💡 고객사 행을 클릭하면 상세 내역을 볼 수 있습니다 · ${DETAIL_YEAR}년 이외 데이터는 연도별 집계</div>`;
  h+=`<div class="table-wrap"><div class="table-scroll"><table><thead><tr>
    <th>기간</th><th>거래처</th><th>건수</th><th>검사수량</th><th>불량수량</th><th>PPM</th><th>주요불량</th></tr></thead><tbody>`;

  for(const pk of periods){
    const pkRows=grouped.filter(g=>g.key===pk);
    let sq=0,sd=0,sn=0;

    for(const r of pkRows){
      sq+=r.q; sd+=r.td; sn+=r.n;
      const td3=Object.entries(r.dv).filter(([,v])=>v>0).sort((a,b)=>b[1]-a[1]).slice(0,3).map(([n,v])=>`${n}(${F(v)})`).join(', ')||'-';
      const rid=`w_${pk}_${r.c}`;
      const isOpen=S.expanded.has(rid);

      h+=`<tr class="sum-row ${isOpen?'open':''}" data-rid="${rid}" onclick="toggleRow('${rid}')">
        <td class="tl" style="font-size:10px"><span class="arrow">▶</span>${r.label}</td>
        <td><span class="ctag ctag-${r.c}">${r.c}</span></td>
        <td>${F(r.n)}</td><td>${F(r.q)}</td><td>${F(r.td)}</td>
        <td><span class="${pC(r.ppm)}">${fp(r.ppm)}</span></td>
        <td class="tl" style="font-size:10px">${td3}</td></tr>`;

      // Detail row
      h+=`<tr class="detail-row ${isOpen?'show':''}" data-did="${rid}"><td colspan="7">${buildDetailPanel(r.rows, r.c)}</td></tr>`;
    }

    // Subtotal row
    const sp=sq>0?sd/sq*1e6:0;
    const subId=`wsub_${pk}`;
    const subOpen=S.expanded.has(subId);
    h+=`<tr class="sub-row sum-row ${subOpen?'open':''}" data-rid="${subId}" onclick="toggleRow('${subId}')">
      <td class="tl" style="font-weight:700"><span class="arrow">▶</span></td><td class="tl" style="font-weight:700">${/^\d{4}$/.test(pk)?pk+'년 소계':'주간 소계'}</td>
      <td>${F(sn)}</td><td>${F(sq)}</td><td>${F(sd)}</td><td><span class="${pC(sp)}">${fp(sp)}</span></td><td></td></tr>`;
    h+=`<tr class="detail-row ${subOpen?'show':''}" data-did="${subId}"><td colspan="7">`;
    for(const cr of pkRows){
      h+=buildDetailPanel(cr.rows, cr.c);
    }
    h+=`</td></tr>`;
  }
  h+=`</tbody></table></div></div>`;
  return h;
}

// ===== Monthly with expand (2026=월별, 나머지=연도) =====
function rMonthly(data){
  const grouped = aggMixed(data, mixedMonKey, (k,d)=>mixedMonLabel(k));
  grouped.sort((a,b)=>a.key<b.key?-1:a.key>b.key?1:a.c<b.c?-1:1);
  const periods=[...new Set(grouped.map(g=>g.key))].sort();

  let h=`<div style="font-size:11px;color:var(--text3);margin-bottom:8px">💡 고객사 행을 클릭하면 상세 내역을 볼 수 있습니다 · ${DETAIL_YEAR}년 이외 데이터는 연도별 집계</div>`;
  h+=`<div class="table-wrap"><div class="table-scroll"><table><thead><tr>
    <th>기간</th><th>거래처</th><th>건수</th><th>검사수량</th><th>불량수량</th><th>PPM</th><th>주요불량</th></tr></thead><tbody>`;

  for(const pk of periods){
    const pkRows=grouped.filter(g=>g.key===pk);
    let sq=0,sd=0,sn=0;

    for(const r of pkRows){
      sq+=r.q; sd+=r.td; sn+=r.n;
      const td3=Object.entries(r.dv).filter(([,v])=>v>0).sort((a,b)=>b[1]-a[1]).slice(0,3).map(([n,v])=>`${n}(${F(v)})`).join(', ')||'-';
      const rid=`m_${pk}_${r.c}`;
      const isOpen=S.expanded.has(rid);

      h+=`<tr class="sum-row ${isOpen?'open':''}" data-rid="${rid}" onclick="toggleRow('${rid}')">
        <td class="tl"><span class="arrow">▶</span>${r.label}</td>
        <td><span class="ctag ctag-${r.c}">${r.c}</span></td>
        <td>${F(r.n)}</td><td>${F(r.q)}</td><td>${F(r.td)}</td>
        <td><span class="${pC(r.ppm)}">${fp(r.ppm)}</span></td>
        <td class="tl" style="font-size:10px">${td3}</td></tr>`;

      h+=`<tr class="detail-row ${isOpen?'show':''}" data-did="${rid}"><td colspan="7">${buildDetailPanel(r.rows, r.c)}</td></tr>`;
    }

    const sp=sq>0?sd/sq*1e6:0;
    const subId=`msub_${pk}`;
    const subOpen=S.expanded.has(subId);
    h+=`<tr class="sub-row sum-row ${subOpen?'open':''}" data-rid="${subId}" onclick="toggleRow('${subId}')">
      <td class="tl" style="font-weight:700"><span class="arrow">▶</span></td><td class="tl" style="font-weight:700">${/^\d{4}$/.test(pk)?pk+'년 소계':'월간 소계'}</td>
      <td>${F(sn)}</td><td>${F(sq)}</td><td>${F(sd)}</td><td><span class="${pC(sp)}">${fp(sp)}</span></td><td></td></tr>`;
    h+=`<tr class="detail-row ${subOpen?'show':''}" data-did="${subId}"><td colspan="7">`;
    for(const cr of pkRows) h+=buildDetailPanel(cr.rows, cr.c);
    h+=`</td></tr>`;
  }
  h+=`</tbody></table></div></div>`;
  return h;
}

// ===== PPM 추이 (Cross-tab: Mixed Periods × Customers) =====
function rTrend(data){
  const grouped=aggMixed(data,mixedWkKey,mixedWkLabel);
  const wkTot=aggTotMixed(data,mixedWkKey,mixedWkLabel);
  const periods=[...new Set(grouped.map(g=>g.key))].sort();
  const custs=[...new Set(data.map(r=>r.c))].sort();

  let h=`<div style="font-size:12px;font-weight:600;margin-bottom:10px;color:var(--text2)">기간별 · 거래처별 PPM 추이 매트릭스 <span style="font-size:10px;font-weight:400;color:var(--text3)">(${DETAIL_YEAR}년 이외는 연도별 집계)</span></div>`;
  h+=`<div class="table-wrap" style="margin-bottom:20px"><div class="table-scroll"><table><thead><tr><th>기간</th>`;
  for(const c of custs) h+=`<th><span class="ctag ctag-${c}">${c}</span></th>`;
  h+=`<th style="background:var(--bg4);font-weight:700">전체</th></tr></thead><tbody>`;

  for(const pk of periods){
    const total=wkTot.find(g=>g.key===pk);
    const isYear=/^\d{4}$/.test(pk);
    const lbl=isYear?pk+'년':mixedWkShortLabel(pk,data);
    const rowStyle=isYear?' style="background:rgba(59,130,246,.06)"':'';
    h+=`<tr${rowStyle}><td class="tl" style="font-size:10px;${isYear?'font-weight:700':''}">` + lbl + `</td>`;
    for(const c of custs){
      const row=grouped.find(g=>g.key===pk&&g.c===c);
      if(row) h+=`<td><span class="${pC(row.ppm)}">${fp(row.ppm)}</span></td>`;
      else h+=`<td style="color:var(--text3)">-</td>`;
    }
    h+=`<td style="font-weight:700;background:rgba(34,40,56,.5)"><span class="${pC(total.ppm)}">${fp(total.ppm)}</span></td></tr>`;
  }

  // 전체 합계 행
  const custTotals={};
  for(const c of custs){
    const cd=data.filter(r=>r.c===c);
    const q=cd.reduce((s,r)=>s+r.q,0), d=cd.reduce((s,r)=>s+r.td,0);
    custTotals[c]=q>0?d/q*1e6:0;
  }
  const allQ=data.reduce((s,r)=>s+r.q,0), allD=data.reduce((s,r)=>s+r.td,0), allP=allQ>0?allD/allQ*1e6:0;
  h+=`<tr class="sub-row"><td class="tl" style="font-weight:700">전체 평균</td>`;
  for(const c of custs) h+=`<td style="font-weight:700"><span class="${pC(custTotals[c])}">${fp(custTotals[c])}</span></td>`;
  h+=`<td style="font-weight:700;background:rgba(34,40,56,.5)"><span class="${pC(allP)}">${fp(allP)}</span></td></tr>`;
  h+=`</tbody></table></div></div>`;

  // 기간별 검사수량/불량수량 요약 테이블
  h+=`<div style="font-size:12px;font-weight:600;margin-bottom:10px;color:var(--text2)">기간별 검사수량 · 불량수량 요약</div>`;
  h+=`<div class="table-wrap"><div class="table-scroll"><table><thead><tr><th>기간</th>`;
  for(const c of custs) h+=`<th colspan="2"><span class="ctag ctag-${c}">${c}</span></th>`;
  h+=`<th colspan="2" style="background:var(--bg4)">전체</th></tr><tr><th></th>`;
  for(const c of custs) h+=`<th style="font-size:9px">검사수</th><th style="font-size:9px">불량수</th>`;
  h+=`<th style="font-size:9px;background:var(--bg4)">검사수</th><th style="font-size:9px;background:var(--bg4)">불량수</th></tr></thead><tbody>`;

  for(const pk of periods){
    const total=wkTot.find(g=>g.key===pk);
    const isYear=/^\d{4}$/.test(pk);
    const lbl=isYear?pk+'년':mixedWkShortLabel(pk,data);
    const rowStyle=isYear?' style="background:rgba(59,130,246,.06)"':'';
    h+=`<tr${rowStyle}><td class="tl" style="font-size:10px;${isYear?'font-weight:700':''}">${lbl}</td>`;
    for(const c of custs){
      const row=grouped.find(g=>g.key===pk&&g.c===c);
      if(row){
        h+=`<td>${F(row.q)}</td><td style="${row.td>0?'color:var(--red);font-weight:600':''}">${F(row.td)}</td>`;
      } else {
        h+=`<td style="color:var(--text3)">-</td><td style="color:var(--text3)">-</td>`;
      }
    }
    h+=`<td style="font-weight:600;background:rgba(34,40,56,.5)">${F(total.q)}</td><td style="font-weight:600;background:rgba(34,40,56,.5);${total.td>0?'color:var(--red)':''}">${F(total.td)}</td></tr>`;
  }
  h+=`</tbody></table></div></div>`;

  // PPM 추이 차트
  h+=`<div class="charts-grid" style="margin-top:16px">
    <div class="chart-box full"><div class="chart-title"><div class="dot" style="background:var(--accent)"></div>PPM 추이 (거래처별) <span style="font-size:10px;font-weight:400;color:var(--text3)">${DETAIL_YEAR}년 이외 연도별</span></div><div class="chart-canvas"><canvas id="ch-trend-ppm"></canvas></div></div>
  </div>`;
  return h;
}

function buildTrendChart(data){
  setTimeout(()=>{
    const custs=[...new Set(data.map(r=>r.c))].sort();
    const grouped=aggMixed(data,mixedWkKey,mixedWkLabel);
    const wkTot=aggTotMixed(data,mixedWkKey,mixedWkLabel);
    const periods=[...new Set(grouped.map(g=>g.key))].sort();

    mkChart('ch-trend-ppm',{type:'line',data:{
      labels:periods.map(pk=>/^\d{4}$/.test(pk)?pk+'년':mixedWkShortLabel(pk,data)),
      datasets:[
        {label:'전체',data:periods.map(pk=>{const t=wkTot.find(x=>x.key===pk);return t?+fp(t.ppm):null;}),
          borderColor:'#fff',backgroundColor:'rgba(255,255,255,.1)',borderWidth:2.5,pointRadius:4,tension:.3,fill:false,order:0},
        ...custs.map(c=>{const ci=CUSTS.find(x=>x.key===c);return{
          label:c,data:periods.map(pk=>{const r=grouped.find(x=>x.key===pk&&x.c===c);return r?+fp(r.ppm):null;}),
          borderColor:ci?.color||'#666',backgroundColor:(ci?.bg||'rgba(100,100,100,')+'0.15)',
          borderWidth:1.8,pointRadius:3,tension:.3,fill:false,spanGaps:true,order:1};})
      ]},
      options:{responsive:true,maintainAspectRatio:false,interaction:{mode:'index',intersect:false},
        plugins:{legend:{position:'bottom'},tooltip:{callbacks:{label:ctx=>ctx.dataset.label+': '+ctx.parsed.y+' PPM'}}},
        scales:{y:{beginAtZero:true,title:{display:true,text:'PPM',font:{size:11}},grid:{color:'#1f2535'}},
          x:{grid:{display:false},ticks:{maxRotation:45,font:{size:9}}}}}});
  },50);
}


// ===== 업체별 주간보고 =====
const CUST_TARGETS={'SKF':600,'SKF_SBC':2000,'일진':1000,'진양오일씰':300};

function rCustReport(data){
  const weeks=[...new Set(data.map(r=>r.wk))].sort();
  if(!weeks.length) return `<div class="rpt-no-data">데이터가 없습니다</div>`;
  if(!S.rptWeek||!weeks.includes(S.rptWeek)) S.rptWeek=weeks[weeks.length-1];
  const wi=weeks.indexOf(S.rptWeek);
  const selWk=S.rptWeek;
  const wkData=data.filter(r=>r.wk===selWk);
  const wkLabel=wkData[0]?.wkL||selWk;
  const custs=[...new Set(data.map(r=>r.c))].sort();

  // --- 주차 선택 UI (인쇄시 숨김) ---
  let h=`<div class="pr-week-sel">
    <div style="font-size:12px;font-weight:600;color:var(--text2)">보고 주차</div>
    <select onchange="setRptWeek(this.value)">`;
  for(let i=weeks.length-1;i>=0;i--){
    const wk=weeks[i],lbl=data.find(r=>r.wk===wk)?.wkL||wk;
    h+=`<option value="${wk}" ${wk===selWk?'selected':''}>${lbl}</option>`;
  }
  h+=`</select>
    <div class="rpt-week-nav">
      <button onclick="navRptWeek(-1)" ${wi<=0?'disabled':''}>◀</button>
      <button onclick="navRptWeek(1)" ${wi>=weeks.length-1?'disabled':''}>▶</button>
    </div>
    <div style="flex:1"></div>
    <button class="btn" onclick="st('dashboard')" style="background:#475569;font-size:12px;padding:6px 14px">✕ 닫기</button>
  </div>`;

  // ===== 흰색 보고서 페이지 =====
  h+=`<div class="pr-page">`;

  // 상단 버튼바
  h+=`<div class="pr-topbar">
    <button class="pr-btn" onclick="window.print()">🖨️ 인쇄 / PDF</button>
    <button class="pr-btn pr-btn-outline" onclick="st('dashboard')">↻ 다시 업로드</button>
    <div class="pr-info">데이터 ${F(data.length)}건 · ${data[0]?.ds||''} ~ ${data[data.length-1]?.ds||''}</div>
  </div>`;

  h+=`<div class="pr-inner">`;

  // ========== 1. PPM 추이 차트 ==========
  h+=`<div class="pr-chart-wrap">
    <div class="pr-chart-title">주간 공정불량률 추이</div>
    <div class="pr-chart-canvas" style="height:230px"><canvas id="cr-main-chart"></canvas></div>
  </div>`;

  // ========== 2. 기간별 PPM 데이터 테이블 (2026=주간, 나머지=연도) ==========
  const mixedPeriods=[...new Set(data.map(r=>mixedWkKey(r)))].sort();
  const showN=Math.min(mixedPeriods.length,30);
  const showPeriods=mixedPeriods.slice(-showN);

  h+=`<div class="pr-wk-table"><div style="overflow-x:auto"><table><thead><tr><th style="min-width:80px;position:sticky;left:0;z-index:3;background:#f1f5f9">구분</th>`;
  for(const pk of showPeriods){
    const isYear=/^\d{4}$/.test(pk);
    let short;
    if(isYear){short=pk+'년';}
    else{const r0=data.find(r=>r.wk===pk);const dateP=r0?r0.wkL.match(/\(([^)]+)\)/):null;short=dateP?dateP[1].split('~')[0].trim():pk.split('-W')[1];}
    const isSel=pk===selWk;
    h+=`<th${isSel?' style="background:#dbeafe;color:#1d4ed8;font-weight:700"':''} ${isYear?'style="background:#e8f0fe;font-weight:700"':''} title="${isYear?pk+'년 집계':data.find(r=>r.wk===pk)?.wkL||pk}">${short}</th>`;
  }
  h+=`</tr></thead><tbody>`;

  const grouped=aggMixed(data,mixedWkKey,mixedWkLabel);
  for(const c of custs){
    const ci=CUSTS.find(x=>x.key===c);
    h+=`<tr><td class="rl" style="position:sticky;left:0;z-index:2;background:#f8fafc"><span style="display:inline-block;width:8px;height:8px;border-radius:2px;background:${ci?.color||'#666'};margin-right:4px;vertical-align:middle"></span>${ci?.name||c}</td>`;
    for(const pk of showPeriods){
      const row=grouped.find(g=>g.key===pk&&g.c===c);
      const isSel=pk===selWk;
      const isYear=/^\d{4}$/.test(pk);
      if(row){
        const v=Math.round(row.ppm);
        const bg=isSel?'background:#eff6ff;':isYear?'background:#f0f4ff;':'';
        h+=`<td style="${bg}">${F(v)}</td>`;
      } else {
        h+=`<td style="color:#d0d5dd;text-align:center">0</td>`;
      }
    }
    h+=`</tr>`;
  }
  h+=`</tbody></table></div></div>`;

  // ========== 3. 하단 2단 레이아웃 ==========
  h+=`<div class="pr-bottom">`;

  // ----- 좌측: 1. 주간 공정불량 -----
  h+=`<div class="pr-section"><div class="pr-sec-hdr">1. 주간 공정불량</div><div class="pr-sec-body">`;

  for(const c of custs){
    const ci=CUSTS.find(x=>x.key===c);
    const cd=wkData.filter(r=>r.c===c);
    const cQ=cd.reduce((s,r)=>s+r.q,0);
    const cD=cd.reduce((s,r)=>s+r.td,0);
    const cP=cQ>0?cD/cQ*1e6:0;
    const target=CUST_TARGETS[c]||1000;
    const dns=Object.values(CMAP[c]?.df||{});

    // 불량유형 집계
    const cDefs={};
    for(const r of cd) for(const[dn,v] of Object.entries(r.dv)) cDefs[dn]=(cDefs[dn]||0)+v;

    h+=`<div class="pr-cust-block">
      <div class="pr-cust-hdr">
        <div class="pr-cust-name"><span class="pc-dot" style="background:${ci?.color||'#666'}"></span>${ci?.name||c}</div>
        <div class="pr-cust-target">목표: <b>${F(target)}PPM</b></div>
      </div>`;

    // 미니 테이블: 검사수 / 불량수 / 불량률
    h+=`<table style="width:100%;border-collapse:collapse;font-size:9px">
      <thead><tr style="background:#f8fafc">
        <th style="padding:3px 6px;border:1px solid #e2e8f0;text-align:left;font-size:8px;color:#64748b;width:60px"></th>
        <th style="padding:3px 6px;border:1px solid #e2e8f0;text-align:right;font-size:8px;color:#64748b">검사수</th>
        <th style="padding:3px 6px;border:1px solid #e2e8f0;text-align:right;font-size:8px;color:#64748b">불량수</th>
        <th style="padding:3px 6px;border:1px solid #e2e8f0;text-align:right;font-size:8px;color:#64748b">불량률(PPM)</th>`;

    // 불량유형 상위 4개 열
    const topDefs=Object.entries(cDefs).filter(([,v])=>v>0).sort((a,b)=>b[1]-a[1]).slice(0,4);
    for(const[dn] of topDefs) h+=`<th style="padding:3px 4px;border:1px solid #e2e8f0;text-align:right;font-size:8px;color:#64748b">${dn}</th>`;

    h+=`</tr></thead><tbody>`;

    // 단일 요약 행
    h+=`<tr>
      <td style="padding:3px 6px;border:1px solid #e2e8f0;font-weight:600;font-family:'Noto Sans KR';background:#f8fafc">합계</td>
      <td style="padding:3px 6px;border:1px solid #e2e8f0;text-align:right;font-family:'JetBrains Mono',monospace">${F(cQ)}</td>
      <td style="padding:3px 6px;border:1px solid #e2e8f0;text-align:right;font-family:'JetBrains Mono',monospace;${cD>0?'color:#dc2626;font-weight:600':''}">${F(cD)}</td>
      <td style="padding:3px 6px;border:1px solid #e2e8f0;text-align:right;font-family:'JetBrains Mono',monospace;font-weight:700;${cP>target?'color:#dc2626':'color:#16a34a'}">${F(Math.round(cP))}</td>`;
    for(const[dn,v] of topDefs){
      h+=`<td style="padding:3px 4px;border:1px solid #e2e8f0;text-align:right;font-family:'JetBrains Mono',monospace;${v>0?'color:#dc2626;font-weight:600':''}">${F(v)}</td>`;
    }
    h+=`</tr></tbody></table></div>`;
  }

  // 전체 합계
  const tQ=wkData.reduce((s,r)=>s+r.q,0), tD=wkData.reduce((s,r)=>s+r.td,0), tP=tQ>0?tD/tQ*1e6:0;
  h+=`<div style="margin-top:6px;padding:5px 8px;background:#f1f5f9;border-radius:4px;border:1px solid #e2e8f0;display:flex;justify-content:space-between;font-size:10px">
    <b style="color:#1e293b">전체 합계</b>
    <span style="color:#475569">검사: <b>${F(tQ)}</b> · 불량: <b style="color:#dc2626">${F(tD)}</b> · PPM: <b><span class="pr-ppm ${prPC(tP)}">${F(Math.round(tP))}</span></b></span>
  </div>`;
  h+=`</div></div>`;

  // ----- 우측: 2. 품번별 불량 현황 (도넛차트) -----
  h+=`<div class="pr-section"><div class="pr-sec-hdr">2. 품번별 불량 현황</div><div class="pr-sec-body" style="display:flex;flex-direction:column;align-items:center;padding:12px">`;

  // 품번 집계
  const partMap=new Map();
  for(const r of wkData){
    const k=r.c+'|'+r.part;
    if(!partMap.has(k)) partMap.set(k,{c:r.c,part:r.part,pn:r.pn,q:0,td:0,dv:{}});
    const g=partMap.get(k);g.q+=r.q;g.td+=r.td;
    for(const[dn,v] of Object.entries(r.dv)) g.dv[dn]=(g.dv[dn]||0)+v;
  }
  let parts=[...partMap.values()].map(g=>({...g,ppm:g.q>0?g.td/g.q*1e6:0}));
  parts.sort((a,b)=>b.td-a.td);
  const defParts=parts.filter(p=>p.td>0);

  if(defParts.length>0){
    h+=`<div style="width:100%;height:280px;position:relative"><canvas id="cr-part-donut"></canvas></div>`;
    // 범례 테이블
    const topN=defParts.slice(0,10);
    h+=`<div style="width:100%;margin-top:8px;max-height:160px;overflow-y:auto"><table style="width:100%;border-collapse:collapse;font-size:9px">
      <thead><tr><th style="padding:3px 6px;border:1px solid #e2e8f0;text-align:left;background:#f8fafc;color:#64748b">품목코드</th>
      <th style="padding:3px 6px;border:1px solid #e2e8f0;text-align:right;background:#f8fafc;color:#64748b">불량수</th>
      <th style="padding:3px 6px;border:1px solid #e2e8f0;text-align:right;background:#f8fafc;color:#64748b">비율</th></tr></thead><tbody>`;
    for(const p of topN){
      const pct=tD>0?(p.td/tD*100).toFixed(1):'0';
      h+=`<tr><td style="padding:2px 6px;border:1px solid #e2e8f0;font-family:'Noto Sans KR'" title="${p.pn}">${p.part}</td>
        <td style="padding:2px 6px;border:1px solid #e2e8f0;text-align:right;font-family:'JetBrains Mono',monospace;color:#dc2626;font-weight:600">${F(p.td)}</td>
        <td style="padding:2px 6px;border:1px solid #e2e8f0;text-align:right;font-family:'JetBrains Mono',monospace">${pct}%</td></tr>`;
    }
    if(defParts.length>10){
      const otherTd=defParts.slice(10).reduce((s,p)=>s+p.td,0);
      const otherPct=tD>0?(otherTd/tD*100).toFixed(1):'0';
      h+=`<tr style="background:#f8fafc"><td style="padding:2px 6px;border:1px solid #e2e8f0;font-family:'Noto Sans KR';color:#94a3b8">기타 ${defParts.length-10}건</td>
        <td style="padding:2px 6px;border:1px solid #e2e8f0;text-align:right;font-family:'JetBrains Mono',monospace">${F(otherTd)}</td>
        <td style="padding:2px 6px;border:1px solid #e2e8f0;text-align:right;font-family:'JetBrains Mono',monospace">${otherPct}%</td></tr>`;
    }
    h+=`</tbody></table></div>`;
  } else {
    h+=`<div style="padding:40px;color:#94a3b8;font-size:12px;text-align:center">불량 데이터 없음</div>`;
  }
  h+=`</div></div>`;

  h+=`</div>`; // pr-bottom
  h+=`</div>`; // pr-inner
  h+=`</div>`; // pr-page
  return h;
}

// 인쇄용 PPM 색상
function prPC(v){return v<=1000?'pr-ppm-g':v<=3000?'pr-ppm-w':'pr-ppm-b';}
function prPmC(v,target){return v<=target?'pm-grn':v<=target*2?'pm-warn':'pm-red';}

// 업체별 주간보고 차트
function buildCustRptCharts(data){
  setTimeout(()=>{
    const custs=[...new Set(data.map(r=>r.c))].sort();
    const grouped=aggMixed(data,mixedWkKey,mixedWkLabel);
    const wkTot=aggTotMixed(data,mixedWkKey,mixedWkLabel);
    const periods=[...new Set(grouped.map(g=>g.key))].sort().slice(-30);

    // 바 차트 데이터셋 (거래처별)
    const datasets=custs.map(c=>{
      const ci=CUSTS.find(x=>x.key===c);
      return {
        type:'bar',
        label:ci?.name||c,
        data:periods.map(pk=>{const r=grouped.find(x=>x.key===pk&&x.c===c);return r?Math.round(r.ppm):0;}),
        backgroundColor:(ci?.bg||'rgba(100,100,100,')+'0.7)',
        borderColor:ci?.color||'#666',
        borderWidth:1,
        borderRadius:2,
        order:2,
        barPercentage:0.9,
        categoryPercentage:0.85
      };
    });

    // 전체 PPM 라인
    datasets.push({
      type:'line',
      label:'전체 PPM',
      data:periods.map(pk=>{const t=wkTot.find(x=>x.key===pk);return t?Math.round(t.ppm):0;}),
      borderColor:'#1e293b',backgroundColor:'transparent',
      borderWidth:2.5,pointRadius:2,tension:.3,fill:false,order:0,
      borderDash:[],
      yAxisID:'y'
    });

    // 휨불량 PPM 라인
    datasets.push({
      type:'line',
      label:'휨 PPM',
      data:periods.map(pk=>{
        const pRows=wkTot.find(x=>x.key===pk);
        if(!pRows) return 0;
        const hwim=pRows.dv['휨']||0;
        return pRows.q>0?Math.round(hwim/pRows.q*1e6):0;
      }),
      borderColor:'#ef4444',backgroundColor:'transparent',
      borderWidth:2,pointRadius:3,tension:.3,fill:false,order:0,
      borderDash:[6,3],
      yAxisID:'y'
    });

    // 라벨
    const labels=periods.map(pk=>/^\d{4}$/.test(pk)?pk+'년':(() => {
      const r=data.find(x=>x.wk===pk);
      if(!r) return pk.split('-W')[1];
      const m=r.wkL.match(/\(([^)]+)\)/);
      return m?m[1].split('~')[0].trim():pk.split('-W')[1];
    })());

    // 선택 주차 강조 배경
    const selIdx=periods.indexOf(S.rptWeek);

    mkChart('cr-main-chart',{
      type:'bar',
      data:{labels,datasets},
      plugins:[{
        id:'selWeekHighlight',
        beforeDraw(chart){
          if(selIdx<0) return;
          const {ctx}=chart;
          const xAxis=chart.scales.x;
          const yAxis=chart.scales.y;
          const barW=xAxis.width/labels.length;
          const x=xAxis.getPixelForValue(selIdx)-barW/2;
          ctx.save();
          ctx.fillStyle='rgba(37,99,235,0.06)';
          ctx.fillRect(x,yAxis.top,barW,yAxis.bottom-yAxis.top);
          ctx.restore();
        }
      }],
      options:{
        responsive:true,maintainAspectRatio:false,
        interaction:{mode:'index',intersect:false},
        plugins:{
          legend:{position:'right',labels:{font:{size:10},padding:6,boxWidth:12,color:'#475569'}},
          tooltip:{callbacks:{
            title:items=>{const i=items[0]?.dataIndex;if(i===undefined)return'';const pk=periods[i];if(/^\d{4}$/.test(pk))return pk+'년';const r=data.find(x=>x.wk===pk);return r?.wkL||pk;},
            label:ctx=>{
              return ctx.dataset.label+': '+F(ctx.parsed.y)+' PPM';
            }
          }}
        },
        scales:{
          y:{
            position:'left',beginAtZero:true,stacked:false,
            title:{display:true,text:'공정불량률(PPM)',font:{size:10},color:'#64748b'},
            grid:{color:'#e2e8f0'},ticks:{color:'#64748b',font:{size:9}}
          },
          x:{
            grid:{display:false},
            ticks:{color:'#64748b',font:{size:8},maxRotation:50,autoSkip:false}
          }
        }
      }
    });

    // ===== 품번별 불량 도넛차트 =====
    const selWk=S.rptWeek;
    const wkData=data.filter(r=>r.wk===selWk);
    const partMap=new Map();
    for(const r of wkData){
      const k=r.c+'|'+r.part;
      if(!partMap.has(k)) partMap.set(k,{c:r.c,part:r.part,pn:r.pn,td:0});
      partMap.get(k).td+=r.td;
    }
    const defParts=[...partMap.values()].filter(p=>p.td>0).sort((a,b)=>b.td-a.td);
    if(defParts.length>0){
      const donutColors=['#ef4444','#f97316','#eab308','#22c55e','#06b6d4','#3b82f6','#8b5cf6','#ec4899','#14b8a6','#f43f5e','#6366f1','#84cc16'];
      const topN=defParts.slice(0,10);
      const otherTd=defParts.slice(10).reduce((s,p)=>s+p.td,0);
      const dLabels=topN.map(p=>p.part);
      const dData=topN.map(p=>p.td);
      const dColors=topN.map((_,i)=>donutColors[i%donutColors.length]);
      if(otherTd>0){dLabels.push('기타');dData.push(otherTd);dColors.push('#9ca3af');}
      mkChart('cr-part-donut',{type:'doughnut',data:{labels:dLabels,datasets:[{data:dData,backgroundColor:dColors,borderColor:'#fff',borderWidth:2}]},
        plugins:[{
          id:'donutLabels',
          afterDraw(chart){
            const{ctx}=chart;
            const meta=chart.getDatasetMeta(0);
            const tot=dData.reduce((a,b)=>a+b,0);
            meta.data.forEach((arc,i)=>{
              if(dData[i]<=0) return;
              const pct=dData[i]/tot*100;
              if(pct<3) return;
              const{x,y}=arc.tooltipPosition();
              ctx.save();
              ctx.textAlign='center';ctx.textBaseline='middle';
              ctx.font='bold 11px JetBrains Mono,monospace';
              ctx.fillStyle='#fff';
              ctx.fillText(F(dData[i]),x,y-6);
              ctx.font='9px Noto Sans KR,sans-serif';
              ctx.fillStyle='rgba(255,255,255,.8)';
              ctx.fillText(pct.toFixed(1)+'%',x,y+8);
              ctx.restore();
            });
          }
        }],
        options:{responsive:true,maintainAspectRatio:false,cutout:'45%',
          plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>{const tot=dData.reduce((a,b)=>a+b,0);return ctx.label+': '+F(ctx.parsed)+'개 ('+(ctx.parsed/tot*100).toFixed(1)+'%)';}}}}
        }
      });
    }
  },80);
}
function navRptWeek(dir){
  const weeks=[...new Set(S.flt.map(r=>r.wk))].sort();
  const i=weeks.indexOf(S.rptWeek);
  const ni=i+dir;
  if(ni>=0&&ni<weeks.length){S.rptWeek=weeks[ni];S.expanded.clear();render();}
}

function rDetail(data){
  const custs=[...new Set(data.map(r=>r.c))].sort();
  let h=`<div style="display:flex;gap:6px;margin-bottom:12px;flex-wrap:wrap">
    <button class="btn btn-sm ${S.cust==='all'?'active':''}" onclick="sc('all')">전체</button>`;
  for(const c of custs)h+=`<button class="btn btn-sm ${S.cust===c?'active':''}" onclick="sc('${c}')">${c}</button>`;
  h+=`</div>`;
  const fd2=S.cust==='all'?data:data.filter(r=>r.c===S.cust);
  const dns=S.cust==='all'?[...new Set(data.flatMap(r=>Object.keys(r.dv)))]:Object.values(CMAP[S.cust]?.df||{});
  h+=`<div class="table-wrap"><div class="table-scroll"><table><thead><tr>
    <th>주차</th><th>일자</th><th>거래처</th><th>품번</th><th>품명</th><th>검사자</th><th>검사수</th>`;
  for(const d of dns)h+=`<th>${d}</th>`;
  h+=`<th>불량계</th><th>PPM</th></tr></thead><tbody>`;
  const show=fd2.slice(-500);
  for(const r of show){
    h+=`<tr><td class="tl" style="font-size:9px">${r.wkL}</td><td>${r.ds}</td><td><span class="ctag ctag-${r.c}">${r.c}</span></td>
      <td style="font-size:10px">${r.part}</td><td class="tl" style="font-size:10px">${r.pn}</td><td class="tl">${r.ins}</td><td>${F(r.q)}</td>`;
    for(const d of dns){const v=r.dv[d]||0;h+=`<td class="${v>0?'def-hi':'def-lo'}">${v}</td>`;}
    h+=`<td style="font-weight:600;${r.td>0?'color:var(--red)':''}">${r.td}</td><td><span class="${pC(r.ppm)}">${fp(r.ppm)}</span></td></tr>`;}
  if(fd2.length>500)h+=`<tr><td colspan="${7+dns.length+2}" style="text-align:center;color:var(--text3);font-family:'Noto Sans KR'">최근 500건 (전체 ${F(fd2.length)}건)</td></tr>`;
  h+=`</tbody></table></div></div>`;return h;
}

function rTop(data){
  const top=topParts(data,30);
  let h=`<div class="table-wrap"><div class="table-scroll"><table><thead><tr>
    <th>순위</th><th>거래처</th><th>품번</th><th>품명</th><th>건수</th><th>검사수량</th><th>불량수량</th><th>PPM</th></tr></thead><tbody>`;
  top.forEach((r,i)=>{
    const rs=i<3?'color:var(--red);font-weight:700':'';
    h+=`<tr><td style="text-align:center;${rs}">${i+1}</td><td><span class="ctag ctag-${r.c}">${r.c}</span></td>
      <td style="font-size:10px">${r.part}</td><td class="tl" style="font-size:10px">${r.pn}</td>
      <td>${F(r.n)}</td><td>${F(r.q)}</td><td>${F(r.td)}</td><td><span class="${pC(r.ppm)}">${fp(r.ppm)}</span></td></tr>`;});
  h+=`</tbody></table></div></div>`;return h;
}

// ===== Dashboard =====
function rDash(data){
  return `<div class="charts-grid">
    <div class="chart-box full"><div class="chart-title"><div class="dot" style="background:var(--accent)"></div>PPM 추이 (거래처별) <span style="font-size:10px;font-weight:400;color:var(--text3)">${DETAIL_YEAR}년 이외 연도별</span></div><div class="chart-canvas"><canvas id="ch-wk-ppm"></canvas></div></div>
    <div class="chart-box"><div class="chart-title"><div class="dot" style="background:var(--cyan)"></div>기간별 PPM & 검사수량</div><div class="chart-canvas"><canvas id="ch-mon-ppm"></canvas></div></div>
    <div class="chart-box"><div class="chart-title"><div class="dot" style="background:var(--purple)"></div>기간별 거래처별 PPM</div><div class="chart-canvas"><canvas id="ch-mon-cust"></canvas></div></div>
    <div class="chart-box"><div class="chart-title"><div class="dot" style="background:var(--orange)"></div>불량유형별 비율</div><div class="chart-canvas"><canvas id="ch-def-donut"></canvas></div></div>
    <div class="chart-box"><div class="chart-title"><div class="dot" style="background:var(--red)"></div>기간별 불량유형 추이</div><div class="chart-canvas"><canvas id="ch-def-trend"></canvas></div></div>
    <div class="chart-box full"><div class="chart-title"><div class="dot" style="background:var(--green)"></div>거래처별 불량유형 비교</div><div class="chart-canvas"><canvas id="ch-def-cust"></canvas></div></div>
    <div class="chart-box"><div class="chart-title"><div class="dot" style="background:#3b82f6"></div>기간별 검사수량 & 불량수량</div><div class="chart-canvas"><canvas id="ch-wk-qty"></canvas></div></div>
    <div class="chart-box"><div class="chart-title"><div class="dot" style="background:#ef4444"></div>PPM 분포</div><div class="chart-canvas"><canvas id="ch-ppm-dist"></canvas></div></div>
  </div>`;
}

function buildCharts(data){
  setTimeout(()=>{
    const custs=[...new Set(data.map(r=>r.c))].sort();
    const wkByCust=aggMixed(data,mixedWkKey,mixedWkLabel);wkByCust.sort((a,b)=>a.key<b.key?-1:a.key>b.key?1:0);
    const wkTot=aggTotMixed(data,mixedWkKey,mixedWkLabel);
    const monByCust=aggMixed(data,mixedMonKey,(k,d)=>mixedMonLabel(k));monByCust.sort((a,b)=>a.key<b.key?-1:a.key>b.key?1:0);
    const monTot=aggTotMixed(data,mixedMonKey,(k,d)=>mixedMonLabel(k));
    const periods=[...new Set(wkByCust.map(w=>w.key))].sort();
    const monPeriods=[...new Set(monByCust.map(m=>m.key))].sort();
    const allDN=[...new Set(data.flatMap(r=>Object.keys(r.dv)))].sort((a,b)=>{
      const ta=data.reduce((s,r)=>s+(r.dv[a]||0),0),tb=data.reduce((s,r)=>s+(r.dv[b]||0),0);return tb-ta;});
    const defTot={};for(const r of data)for(const[dn,v]of Object.entries(r.dv))defTot[dn]=(defTot[dn]||0)+v;
    const defArr=Object.entries(defTot).filter(([,v])=>v>0).sort((a,b)=>b[1]-a[1]);
    const defByCust=new Map();for(const r of data){if(!defByCust.has(r.c))defByCust.set(r.c,{});const g=defByCust.get(r.c);for(const[dn,v]of Object.entries(r.dv))g[dn]=(g[dn]||0)+v;}

    mkChart('ch-wk-ppm',{type:'line',data:{labels:periods.map(pk=>/^\d{4}$/.test(pk)?pk+'년':mixedWkShortLabel(pk,data)),
      datasets:[{label:'전체',data:periods.map(pk=>{const t=wkTot.find(x=>x.key===pk);return t?+fp(t.ppm):null;}),borderColor:'#fff',backgroundColor:'rgba(255,255,255,.1)',borderWidth:2.5,pointRadius:3,tension:.3,fill:false,order:0},
        ...custs.map(c=>{const ci=CUSTS.find(x=>x.key===c);return{label:c,data:periods.map(pk=>{const r=wkByCust.find(x=>x.key===pk&&x.c===c);return r?+fp(r.ppm):null;}),borderColor:ci?.color||'#666',backgroundColor:(ci?.bg||'rgba(100,100,100,')+'0.15)',borderWidth:1.8,pointRadius:2,tension:.3,fill:false,spanGaps:true,order:1};})]},
      options:{responsive:true,maintainAspectRatio:false,interaction:{mode:'index',intersect:false},plugins:{legend:{position:'bottom'}},
        scales:{y:{beginAtZero:true,title:{display:true,text:'PPM',font:{size:11}},grid:{color:'#1f2535'}},x:{grid:{display:false},ticks:{maxRotation:45,font:{size:9}}}}}});

    mkChart('ch-mon-ppm',{type:'bar',data:{labels:monPeriods.map(pk=>mixedMonLabel(pk)),
      datasets:[{type:'bar',label:'검사수량',data:monPeriods.map(pk=>{const t=monTot.find(x=>x.key===pk);return t?t.q:0;}),backgroundColor:'rgba(59,130,246,.25)',borderColor:'#3b82f6',borderWidth:1,yAxisID:'y1',order:2,borderRadius:4},
        {type:'line',label:'PPM',data:monPeriods.map(pk=>{const t=monTot.find(x=>x.key===pk);return t?+fp(t.ppm):0;}),borderColor:'#ef4444',backgroundColor:'rgba(239,68,68,.1)',borderWidth:2.5,pointRadius:4,tension:.3,yAxisID:'y',order:1,fill:false}]},
      options:{responsive:true,maintainAspectRatio:false,interaction:{mode:'index',intersect:false},plugins:{legend:{position:'bottom'}},
        scales:{y:{position:'left',beginAtZero:true,title:{display:true,text:'PPM'},grid:{color:'#1f2535'}},
          y1:{position:'right',beginAtZero:true,title:{display:true,text:'검사수량'},grid:{display:false},ticks:{callback:v=>v>=1e6?(v/1e6).toFixed(1)+'M':v>=1e3?(v/1e3).toFixed(0)+'K':v}},x:{grid:{display:false}}}}});

    mkChart('ch-mon-cust',{type:'bar',data:{labels:monPeriods.map(pk=>mixedMonLabel(pk)),
      datasets:custs.map(c=>{const ci=CUSTS.find(x=>x.key===c);return{label:c,data:monPeriods.map(pk=>{const r=monByCust.find(x=>x.key===pk&&x.c===c);return r?+fp(r.ppm):0;}),backgroundColor:(ci?.bg||'rgba(100,100,100,')+'0.6)',borderColor:ci?.color||'#666',borderWidth:1,borderRadius:3};})},
      options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'bottom'}},
        scales:{y:{beginAtZero:true,title:{display:true,text:'PPM'},grid:{color:'#1f2535'}},x:{grid:{display:false}}}}});

    if(defArr.length>0){mkChart('ch-def-donut',{type:'doughnut',data:{labels:defArr.map(([n])=>n),
      datasets:[{data:defArr.map(([,v])=>v),backgroundColor:defArr.map(([n])=>DC[n]||'#6b7280'),borderColor:'#12161e',borderWidth:2}]},
      options:{responsive:true,maintainAspectRatio:false,cutout:'55%',plugins:{legend:{position:'right',labels:{font:{size:11},padding:8}}}}});}

    mkChart('ch-def-trend',{type:'bar',data:{labels:monPeriods.map(pk=>mixedMonLabel(pk)),
      datasets:allDN.map(dn=>({label:dn,data:monPeriods.map(pk=>{const t=monTot.find(x=>x.key===pk);return t?(t.dv[dn]||0):0;}),backgroundColor:DC[dn]||'#6b7280',borderRadius:2}))},
      options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'bottom',labels:{font:{size:10}}}},
        scales:{y:{stacked:true,beginAtZero:true,title:{display:true,text:'불량수'},grid:{color:'#1f2535'}},x:{stacked:true,grid:{display:false}}}}});

    const topDef=allDN.slice(0,8);
    mkChart('ch-def-cust',{type:'bar',data:{labels:topDef,
      datasets:custs.map(c=>{const ci=CUSTS.find(x=>x.key===c);const cd=defByCust.get(c)||{};return{label:c,data:topDef.map(dn=>cd[dn]||0),backgroundColor:(ci?.bg||'rgba(100,100,100,')+'0.6)',borderColor:ci?.color||'#666',borderWidth:1,borderRadius:3};})},
      options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'bottom'}},
        scales:{x:{beginAtZero:true,grid:{color:'#1f2535'}},y:{grid:{display:false}}}}});

    mkChart('ch-wk-qty',{type:'bar',data:{labels:periods.map(pk=>/^\d{4}$/.test(pk)?pk+'년':mixedWkShortLabel(pk,data)),
      datasets:[{type:'bar',label:'검사수량',data:periods.map(pk=>{const t=wkTot.find(x=>x.key===pk);return t?t.q:0;}),backgroundColor:'rgba(59,130,246,.3)',borderColor:'#3b82f6',borderWidth:1,yAxisID:'y1',order:2,borderRadius:3},
        {type:'line',label:'불량수량',data:periods.map(pk=>{const t=wkTot.find(x=>x.key===pk);return t?t.td:0;}),borderColor:'#ef4444',backgroundColor:'rgba(239,68,68,.1)',borderWidth:2,pointRadius:3,tension:.3,yAxisID:'y',order:1}]},
      options:{responsive:true,maintainAspectRatio:false,interaction:{mode:'index',intersect:false},plugins:{legend:{position:'bottom'}},
        scales:{y:{position:'left',beginAtZero:true,title:{display:true,text:'불량수'},grid:{color:'#1f2535'}},
          y1:{position:'right',beginAtZero:true,title:{display:true,text:'검사수량'},grid:{display:false},ticks:{callback:v=>v>=1e6?(v/1e6).toFixed(1)+'M':v>=1e3?(v/1e3).toFixed(0)+'K':v}},x:{grid:{display:false},ticks:{maxRotation:45,font:{size:9}}}}}});

    const ppmVals=data.map(r=>r.ppm),bins=[0,500,1000,2000,3000,5000,10000,50000,Infinity],
      bL=['0~500','500~1K','1K~2K','2K~3K','3K~5K','5K~10K','10K~50K','50K+'],
      bC=['#22c55e','#4ade80','#a3e635','#eab308','#f97316','#ef4444','#dc2626','#991b1b'],
      bV=bL.map((_,i)=>ppmVals.filter(v=>v>=bins[i]&&v<bins[i+1]).length);
    mkChart('ch-ppm-dist',{type:'bar',data:{labels:bL,datasets:[{label:'건수',data:bV,backgroundColor:bC,borderRadius:4}]},
      options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},
        scales:{y:{beginAtZero:true,title:{display:true,text:'건수'},grid:{color:'#1f2535'}},x:{grid:{display:false},title:{display:true,text:'PPM 구간'}}}}});
  },50);
}

// ===== 인쇄용 보고서 =====

function rPrintReport(data){
  const weeks=[...new Set(data.map(r=>r.wk))].sort();
  if(!weeks.length) return `<div class="rpt-no-data">데이터가 없습니다</div>`;

  // 주차 선택 (필터에 따라 전체 or 특정 주차)
  if(!S.rptWeek||!weeks.includes(S.rptWeek)) S.rptWeek=weeks[weeks.length-1];

  // 주차 선택 UI (화면용, 인쇄시 숨김)
  let h=`<div class="pr-week-sel">
    <div style="font-size:12px;font-weight:600;color:var(--text2)">보고 주차</div>
    <select onchange="setRptWeek(this.value)">`;
  for(let i=weeks.length-1;i>=0;i--){
    const wk=weeks[i]; const lbl=data.find(r=>r.wk===wk)?.wkL||wk;
    h+=`<option value="${wk}" ${wk===S.rptWeek?'selected':''}>${lbl}</option>`;
  }
  h+=`</select>
    <div class="rpt-week-nav">
      <button onclick="navRptWeek(-1)" ${weeks.indexOf(S.rptWeek)<=0?'disabled':''}>◀</button>
      <button onclick="navRptWeek(1)" ${weeks.indexOf(S.rptWeek)>=weeks.length-1?'disabled':''}>▶</button>
    </div>
    <div style="flex:1"></div>
    <button class="btn" onclick="st('dashboard')" style="background:#475569;font-size:12px;padding:6px 14px">✕ 닫기</button>
  </div>`;

  const selWk=S.rptWeek;
  const wkData=data.filter(r=>r.wk===selWk);
  const wkLabel=wkData[0]?.wkL||selWk;
  const custs=[...new Set(data.map(r=>r.c))].sort();
  const allDNs=[...new Set(data.flatMap(r=>Object.keys(r.dv)))];
  // Sort by total count
  const dnTotals={};for(const r of data) for(const[dn,v] of Object.entries(r.dv)) dnTotals[dn]=(dnTotals[dn]||0)+v;
  allDNs.sort((a,b)=>(dnTotals[b]||0)-(dnTotals[a]||0));

  // ========== Print Page Start ==========
  h+=`<div class="pr-page">`;

  // Top bar
  h+=`<div class="pr-topbar">
    <button class="pr-btn" onclick="window.print()">🖨️ 인쇄 / PDF</button>
    <button class="pr-btn pr-btn-outline" onclick="st('dashboard')">↻ 다시 돌아가기</button>
    <div class="pr-info">데이터 ${F(data.length)}건 · ${data[0]?.ds||''} ~ ${data[data.length-1]?.ds||''}</div>
  </div>`;

  h+=`<div class="pr-inner">`;

  // Title
  h+=`<div class="pr-title-row">
    <div><div class="pr-title">주간 출하검사 불량률 추이</div></div>
    <div class="pr-sub">${wkLabel} 기준 · 품질관리팀</div>
  </div>`;

  // ===== 1. Main PPM Trend Chart =====
  h+=`<div class="pr-chart-wrap">
    <div class="pr-chart-canvas"><canvas id="pr-main-chart"></canvas></div>
  </div>`;

  // ===== 2. 기간별 데이터 테이블 (2026=주간, 나머지=연도) =====
  const mixedPeriods2=[...new Set(data.map(r=>mixedWkKey(r)))].sort();
  const recentN=Math.min(mixedPeriods2.length,25);
  const recentPks=mixedPeriods2.slice(-recentN);
  const grouped=aggMixed(data,mixedWkKey,mixedWkLabel);

  h+=`<div class="pr-wk-table"><table><thead><tr><th style="min-width:80px">구분</th>`;
  for(const pk of recentPks){
    const isYear=/^\d{4}$/.test(pk);
    let short;
    if(isYear){short=pk+'년';}
    else{const lbl=data.find(r=>r.wk===pk)?.wkL||pk;short=lbl.replace(/\d{4}년\s*/,'').replace(/주\s*\(/,'주(').replace(/\s*~\s*/,'~').split('(')[1]?.replace(')','').split('~')[0]||pk.split('-W')[1];}
    const isSel=pk===selWk;
    h+=`<th${isSel?' style="background:#dbeafe;color:#1d4ed8"':''}${isYear?' style="background:#e8f0fe;font-weight:700"':''}>${short}</th>`;
  }
  h+=`</tr></thead><tbody>`;

  for(const c of custs){
    const ci=CUSTS.find(x=>x.key===c);
    h+=`<tr><td class="rl"><span style="display:inline-block;width:8px;height:8px;border-radius:2px;background:${ci?.color||'#666'};margin-right:3px;vertical-align:middle"></span>${ci?.name||c}</td>`;
    for(const pk of recentPks){
      const row=grouped.find(g=>g.key===pk&&g.c===c);
      const isSel=pk===selWk;
      const isYear=/^\d{4}$/.test(pk);
      if(row){
        const bg=isSel?' style="background:#eff6ff"':isYear?' style="background:#f0f4ff"':'';
        h+=`<td${bg}><span class="pr-ppm ${prPC(row.ppm)}">${fp(row.ppm)}</span></td>`;
      } else {
        h+=`<td style="color:#ccc;text-align:center">-</td>`;
      }
    }
    h+=`</tr>`;
  }
  h+=`</tbody></table></div>`;

  // ===== 3. Bottom: Left=업체별요약, Right=품번별불량 =====
  h+=`<div class="pr-bottom">`;

  // ----- LEFT: 업체별 주간불량 요약 -----
  h+=`<div class="pr-section"><div class="pr-sec-hdr">1. 업체별 주간불량 요약 (${wkLabel})</div><div class="pr-sec-body">`;
  for(const c of custs){
    const ci=CUSTS.find(x=>x.key===c);
    const cd=wkData.filter(r=>r.c===c);
    const cQ=cd.reduce((s,r)=>s+r.q,0), cD=cd.reduce((s,r)=>s+r.td,0);
    const cP=cQ>0?cD/cQ*1e6:0;
    const target=CUST_TARGETS[c]||1000;

    // Per-customer defect breakdown
    const cDefs={};
    for(const r of cd) for(const[dn,v] of Object.entries(r.dv)) cDefs[dn]=(cDefs[dn]||0)+v;
    const topDefs=Object.entries(cDefs).filter(([,v])=>v>0).sort((a,b)=>b[1]-a[1]).slice(0,4);
    const topDefStr=topDefs.map(([n,v])=>`${n}(${F(v)})`).join(', ')||'-';

    h+=`<div class="pr-cust-block">
      <div class="pr-cust-hdr">
        <div class="pr-cust-name"><span class="pc-dot" style="background:${ci?.color||'#666'}"></span>${ci?.name||c}</div>
        <div class="pr-cust-target">목표: <b>${F(target)} PPM</b></div>
      </div>
      <div class="pr-mini-grid">
        <div class="pr-mini-cell"><div class="pm-label">검사수</div><div class="pm-val">${F(cQ)}</div></div>
        <div class="pr-mini-cell"><div class="pm-label">불량수</div><div class="pm-val ${cD>0?'pm-red':''}">${F(cD)}</div></div>
        <div class="pr-mini-cell"><div class="pm-label">불량률(PPM)</div><div class="pm-val ${prPmC(cP,target)}">${fp(cP)}</div></div>
      </div>`;
    if(topDefs.length>0){
      h+=`<div style="padding:3px 8px;font-size:9px;color:#6b7280;background:#f8fafc;border-top:1px solid #e2e8f0">주요불량: ${topDefStr}</div>`;
    }
    h+=`</div>`;
  }

  // 전체 합계
  const tQ=wkData.reduce((s,r)=>s+r.q,0), tD=wkData.reduce((s,r)=>s+r.td,0), tP=tQ>0?tD/tQ*1e6:0;
  h+=`<div style="margin-top:6px;padding:6px 8px;background:#f1f5f9;border-radius:4px;border:1px solid #e2e8f0;display:flex;justify-content:space-between;align-items:center">
    <span style="font-size:10px;font-weight:700;color:#1e293b">합계</span>
    <span style="font-size:10px;color:#475569">검사: ${F(tQ)} · 불량: ${F(tD)} · <b style="font-size:11px"><span class="pr-ppm ${prPC(tP)}">${fp(tP)} PPM</span></b></span>
  </div>`;

  h+=`</div></div>`;

  // ----- RIGHT: 품번별 불량 현황 (도넛차트) -----
  h+=`<div class="pr-section"><div class="pr-sec-hdr">2. 품번별 불량 현황 (${wkLabel})</div><div class="pr-sec-body" style="display:flex;flex-direction:column;align-items:center;padding:12px">`;

  // 품번 집계
  const partMap=new Map();
  for(const r of wkData){
    const k=r.c+'|'+r.part;
    if(!partMap.has(k)) partMap.set(k,{c:r.c,part:r.part,pn:r.pn,q:0,td:0,dv:{}});
    const g=partMap.get(k);g.q+=r.q;g.td+=r.td;
    for(const[dn,v] of Object.entries(r.dv)) g.dv[dn]=(g.dv[dn]||0)+v;
  }
  let parts=[...partMap.values()].map(g=>({...g,ppm:g.q>0?g.td/g.q*1e6:0}));
  parts.sort((a,b)=>b.td-a.td);
  const defParts2=parts.filter(p=>p.td>0);

  if(defParts2.length>0){
    h+=`<div style="width:100%;height:240px;position:relative"><canvas id="pr-part-donut"></canvas></div>`;
    const topN2=defParts2.slice(0,10);
    const prTotD=wkData.reduce((s,r)=>s+r.td,0);
    h+=`<div style="width:100%;margin-top:6px;max-height:140px;overflow-y:auto"><table style="width:100%;border-collapse:collapse;font-size:9px">
      <thead><tr><th style="padding:2px 4px;border:1px solid #e2e8f0;text-align:left;background:#f8fafc;color:#64748b;font-size:8px">품목코드</th>
      <th style="padding:2px 4px;border:1px solid #e2e8f0;text-align:right;background:#f8fafc;color:#64748b;font-size:8px">불량수</th>
      <th style="padding:2px 4px;border:1px solid #e2e8f0;text-align:right;background:#f8fafc;color:#64748b;font-size:8px">비율</th></tr></thead><tbody>`;
    for(const p of topN2){
      const pct=prTotD>0?(p.td/prTotD*100).toFixed(1):'0';
      h+=`<tr><td style="padding:2px 4px;border:1px solid #e2e8f0;font-family:'Noto Sans KR'" title="${p.pn}">${p.part}</td>
        <td style="padding:2px 4px;border:1px solid #e2e8f0;text-align:right;font-family:'JetBrains Mono',monospace;color:#dc2626;font-weight:600">${F(p.td)}</td>
        <td style="padding:2px 4px;border:1px solid #e2e8f0;text-align:right;font-family:'JetBrains Mono',monospace">${pct}%</td></tr>`;
    }
    if(defParts2.length>10){
      const otherTd2=defParts2.slice(10).reduce((s,p)=>s+p.td,0);
      const otherPct2=prTotD>0?(otherTd2/prTotD*100).toFixed(1):'0';
      h+=`<tr style="background:#f8fafc"><td style="padding:2px 4px;border:1px solid #e2e8f0;font-family:'Noto Sans KR';color:#94a3b8">기타 ${defParts2.length-10}건</td>
        <td style="padding:2px 4px;border:1px solid #e2e8f0;text-align:right;font-family:'JetBrains Mono',monospace">${F(otherTd2)}</td>
        <td style="padding:2px 4px;border:1px solid #e2e8f0;text-align:right;font-family:'JetBrains Mono',monospace">${otherPct2}%</td></tr>`;
    }
    h+=`</tbody></table></div>`;
  } else {
    h+=`<div style="padding:40px;color:#94a3b8;font-size:12px;text-align:center">불량 데이터 없음</div>`;
  }
  h+=`</div></div>`;

  h+=`</div>`; // pr-bottom
  h+=`</div>`; // pr-inner
  h+=`</div>`; // pr-page
  return h;
}

function buildPrintCharts(data){
  setTimeout(()=>{
    const custs=[...new Set(data.map(r=>r.c))].sort();
    const grouped=aggMixed(data,mixedWkKey,mixedWkLabel);
    const wkTot=aggTotMixed(data,mixedWkKey,mixedWkLabel);
    const periods=[...new Set(grouped.map(g=>g.key))].sort().slice(-25);

    const datasets=custs.map(c=>{
      const ci=CUSTS.find(x=>x.key===c);
      return {
        label:ci?.name||c,
        data:periods.map(pk=>{const r=grouped.find(x=>x.key===pk&&x.c===c);return r?+fp(r.ppm):null;}),
        borderColor:ci?.color||'#666',
        backgroundColor:(ci?.bg||'rgba(100,100,100,')+'0.12)',
        borderWidth:2,pointRadius:2.5,tension:.3,fill:false,spanGaps:true
      };
    });

    // Add total line
    datasets.unshift({
      label:'전체',
      data:periods.map(pk=>{const t=wkTot.find(x=>x.key===pk);return t?+fp(t.ppm):null;}),
      borderColor:'#1e293b',backgroundColor:'rgba(30,41,59,.08)',
      borderWidth:2.5,pointRadius:3,tension:.3,fill:false,borderDash:[5,3],order:0
    });

    // 휨불량 PPM 라인
    datasets.push({
      label:'휨 PPM',
      data:periods.map(pk=>{
        const pRows=wkTot.find(x=>x.key===pk);
        if(!pRows) return null;
        const hwim=pRows.dv['휨']||0;
        return pRows.q>0?+(hwim/pRows.q*1e6).toFixed(1):null;
      }),
      borderColor:'#ef4444',backgroundColor:'rgba(239,68,68,.08)',
      borderWidth:2,pointRadius:3,tension:.3,fill:false,borderDash:[6,3],spanGaps:true
    });

    mkChart('pr-main-chart',{type:'line',data:{
      labels:periods.map(pk=>{
        if(/^\d{4}$/.test(pk)) return pk+'년';
        const r=grouped.find(x=>x.key===pk);
        const lbl=r?r.label:'';
        return lbl.replace(/\d{4}년\s*/,'').replace(/주\s*\(/,'주(').replace(/\s*~\s*/,'~').split('(')[1]?.replace(')','').split('~')[0]||pk.split('-W')[1];
      }),
      datasets:datasets},
      options:{responsive:true,maintainAspectRatio:false,
        interaction:{mode:'index',intersect:false},
        plugins:{
          legend:{position:'right',labels:{font:{size:10},padding:6,boxWidth:12,color:'#475569'}},
          tooltip:{callbacks:{label:ctx=>ctx.dataset.label+': '+ctx.parsed.y+' PPM'}}
        },
        scales:{
          y:{beginAtZero:true,title:{display:true,text:'공정불량률(PPM)',font:{size:10},color:'#64748b'},
            grid:{color:'#e2e8f0'},ticks:{color:'#64748b',font:{size:9}}},
          x:{grid:{display:false},ticks:{color:'#64748b',font:{size:8},maxRotation:45}}
        }
      }
    });

    // ===== 품번별 불량 도넛차트 (인쇄용) =====
    const selWk2=S.rptWeek;
    const wkData2=data.filter(r=>r.wk===selWk2);
    const partMap2=new Map();
    for(const r of wkData2){
      const k=r.c+'|'+r.part;
      if(!partMap2.has(k)) partMap2.set(k,{c:r.c,part:r.part,pn:r.pn,td:0});
      partMap2.get(k).td+=r.td;
    }
    const defParts3=[...partMap2.values()].filter(p=>p.td>0).sort((a,b)=>b.td-a.td);
    if(defParts3.length>0){
      const donutColors=['#ef4444','#f97316','#eab308','#22c55e','#06b6d4','#3b82f6','#8b5cf6','#ec4899','#14b8a6','#f43f5e','#6366f1','#84cc16'];
      const topN3=defParts3.slice(0,10);
      const otherTd3=defParts3.slice(10).reduce((s,p)=>s+p.td,0);
      const dLabels=topN3.map(p=>p.part);
      const dData=topN3.map(p=>p.td);
      const dColors=topN3.map((_,i)=>donutColors[i%donutColors.length]);
      if(otherTd3>0){dLabels.push('기타');dData.push(otherTd3);dColors.push('#9ca3af');}
      mkChart('pr-part-donut',{type:'doughnut',data:{labels:dLabels,datasets:[{data:dData,backgroundColor:dColors,borderColor:'#fff',borderWidth:2}]},
        plugins:[{
          id:'donutLabels2',
          afterDraw(chart){
            const{ctx}=chart;
            const meta=chart.getDatasetMeta(0);
            const tot=dData.reduce((a,b)=>a+b,0);
            meta.data.forEach((arc,i)=>{
              if(dData[i]<=0) return;
              const pct=dData[i]/tot*100;
              if(pct<3) return;
              const{x,y}=arc.tooltipPosition();
              ctx.save();
              ctx.textAlign='center';ctx.textBaseline='middle';
              ctx.font='bold 10px JetBrains Mono,monospace';
              ctx.fillStyle='#fff';
              ctx.fillText(F(dData[i]),x,y-5);
              ctx.font='8px Noto Sans KR,sans-serif';
              ctx.fillStyle='rgba(255,255,255,.8)';
              ctx.fillText(pct.toFixed(1)+'%',x,y+7);
              ctx.restore();
            });
          }
        }],
        options:{responsive:true,maintainAspectRatio:false,cutout:'45%',
          plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>{const tot=dData.reduce((a,b)=>a+b,0);return ctx.label+': '+F(ctx.parsed)+'개 ('+(ctx.parsed/tot*100).toFixed(1)+'%)';}}}}
        }
      });
    }
  },80);
}

// ===== EVENTS =====
function handleF(k,el){if(el.files.length)loadF(k,el.files[0]);}
function sf(f,v){S[f]=v;filt();S.expanded.clear();render();}
function qr(r){if(r==='all'){S.sd='';S.ed='';}else{const e=new Date(),s=new Date();if(r==='1m')s.setMonth(s.getMonth()-1);if(r==='3m')s.setMonth(s.getMonth()-3);if(r==='6m')s.setMonth(s.getMonth()-6);S.sd=fd(s);S.ed=fd(e);}filt();S.expanded.clear();render();}
function st(t){S.tab=t;S.expanded.clear();render();}
function sc(c){S.cust=c;render();}
function expCSV(){const d=S.flt;if(!d.length)return;let csv='\uFEFF주차,일자,거래처,품번,품명,검사자,검사수,불량계,PPM\n';for(const r of d)csv+=`"${r.wkL}",${r.ds},${r.c},"${r.part}","${r.pn}",${r.ins},${r.q},${r.td},${fp(r.ppm)}\n`;const b=new Blob([csv],{type:'text/csv;charset=utf-8;'});const a=document.createElement('a');a.href=URL.createObjectURL(b);a.download=`출하검사_보고서_${fd(new Date())}.csv`;a.click();}

render();
</script>
</body>
</html>
