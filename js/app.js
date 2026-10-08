
"use strict";
let DATA = (typeof window !== "undefined" && window.DASH_DATA) ? window.DASH_DATA : [];
let CONFIG = (typeof window !== "undefined" && window.DASH_CONFIG) ? window.DASH_CONFIG : {plan_total:0,plan_al:0,plan_cu:0,month_ar:"",year:0,month_days:31,source:"",generated:""};
const ORIG_DATA = DATA;
const ORIG_CONFIG = Object.assign({}, CONFIG);
let PLAN = 0, PLAN_AL = 0, PLAN_CU = 0, MDAYS = 31, HAS_PLAN = false;
function syncPlanVars(){
  PLAN = Number(CONFIG.plan_total)||0;
  PLAN_AL = Number(CONFIG.plan_al)||0;
  PLAN_CU = Number(CONFIG.plan_cu)||0;
  MDAYS = Number(CONFIG.month_days)||31;
  HAS_PLAN = PLAN>0;
}
syncPlanVars();
const AR_MONTHS_JS = ["","يناير","فبراير","مارس","أبريل","مايو","يونيو","يوليو","أغسطس","سبتمبر","أكتوبر","نوفمبر","ديسمبر"];
const CAN_IMPORT = (typeof XLSX !== "undefined");
const FAM_ORDER = ["MV","LV","HV","OH","BC","CC","IC"];
const FAM_COLOR = {MV:"#1976d2",LV:"#0f9d58",HV:"#7c3aed",OH:"#dc2626",BC:"#ea580c",CC:"#ca8a04",IC:"#64748b"};
const FAM_NAME  = {MV:"جهد متوسط MV",LV:"جهد منخفض LV",HV:"جهد عالي HV",OH:"هوائي OH",BC:"مباني BC",CC:"كنترول CC",IC:"داخلي IC"};
const POWER_FAMS = new Set(["LV","MV","HV","OH"]);
const MAT_COLOR = {AL:"#8fa3bd",CU:"#c2570b",SCR:"#7c3aed",ARM:"#0d9488"};
const QC_LABEL = {OK:"OK — مقبول",Hold:"Hold — موقوف",P:"بدون حالة"};
const QC_COLOR = {OK:"#0b8a4b",Hold:"#c77700",P:"#64748b"};

const state = {days:new Set(),fams:new Set(),cust:"",mac:"",mkt:"all",con:"all",qc:"all"};
const tstate = {page:1,pageSize:50,sortKey:"d",sortDir:1};

const nf=(x,d=1)=>(x||0).toLocaleString("en-US",{minimumFractionDigits:d,maximumFractionDigits:d});
const n0=x=>Math.round(x||0).toLocaleString("en-US");
const pf=x=>((x||0)*100).toFixed(1)+"%";
const esc=s=>String(s).replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const $=id=>document.getElementById(id);

function computeKPIs(rows){
  const k={packs:rows.length,totalT:0,ct:0,st:0,at:0,al:0,cu:0,km:0,expT:0,locT:0,holdN:0,holdT:0,okN:0,pN:0,
    wos:new Set(),custs:new Set(),days:new Set(),maxDay:0,famT:{},
    dailyPow:new Array(MDAYS+2).fill(0),dailyTel:new Array(MDAYS+2).fill(0),dailyPacks:new Array(MDAYS+2).fill(0),cum:new Array(MDAYS+2).fill(0)};
  FAM_ORDER.forEach(f=>k.famT[f]=0);
  rows.forEach(r=>{
    k.totalT+=r.tt;k.ct+=r.ct;k.st+=r.st;k.at+=r.at;k.km+=r.q/1000;
    if(r.con==="A")k.al+=r.ct;else if(r.con==="C")k.cu+=r.ct;
    if(r.exp==="Y")k.expT+=r.tt;else k.locT+=r.tt;
    if(r.qc==="Hold"){k.holdN++;k.holdT+=r.tt;}else if(r.qc==="OK")k.okN++;else k.pN++;
    k.wos.add(r.wo);k.custs.add(r.cu);k.days.add(r.d);
    if(r.d>k.maxDay)k.maxDay=r.d;
    if(k.famT[r.fam]===undefined)k.famT[r.fam]=0;
    k.famT[r.fam]+=r.tt;
    if(POWER_FAMS.has(r.fam))k.dailyPow[r.d]+=r.tt;else k.dailyTel[r.d]+=r.tt;
    k.dailyPacks[r.d]++;
  });
  k.nDays=k.days.size;
  k.dailyAvg=k.nDays?k.totalT/k.nDays:0;
  k.paceReq=PLAN/MDAYS;
  k.ratio=HAS_PLAN?k.totalT/PLAN:0;
  k.expected=k.maxDay/MDAYS;
  k.forecast=k.dailyAvg*MDAYS;
  k.forecastRatio=HAS_PLAN?k.forecast/PLAN:0;
  k.behind=k.expected*PLAN-k.totalT;
  k.donePct=HAS_PLAN?k.totalT/PLAN:0;
  k.remaining=HAS_PLAN?Math.max(0,PLAN-k.totalT):0;
  k.daysLeft=Math.max(0,MDAYS-k.maxDay);
  k.paceToFinish=(HAS_PLAN&&k.daysLeft>0)?k.remaining/k.daysLeft:0;
  k.speedUpPct=(HAS_PLAN&&k.dailyAvg>0&&k.remaining>0)?(k.paceToFinish/k.dailyAvg-1):0;
  k.feasible=!HAS_PLAN?null:(k.remaining<=0.0001||k.paceToFinish<=k.dailyAvg*1.05);
  let c=0;for(let d=1;d<=MDAYS;d++){c+=k.dailyPow[d]+k.dailyTel[d];k.cum[d]=c;}
  return k;
}

function filteredRows(){
  return DATA.filter(r=>
    (state.days.size===0||state.days.has(r.d))&&
    (state.fams.size===0||state.fams.has(r.fam))&&
    (!state.cust||r.cu===state.cust)&&
    (!state.mac||r.mac===state.mac)&&
    (state.mkt==="all"||r.exp===(state.mkt==="exp"?"Y":"N"))&&
    (state.con==="all"||r.con===state.con)&&
    (state.qc==="all"||r.qc===state.qc));
}

function groupTons(rows,keyFn){const m=new Map();rows.forEach(r=>{const kk=keyFn(r);m.set(kk,(m.get(kk)||0)+r.tt);});return m;}
function groupCount(rows,keyFn){const m=new Map();rows.forEach(r=>{const kk=keyFn(r);m.set(kk,(m.get(kk)||0)+1);});return m;}
function topFrom(map,n){return [...map.entries()].sort((a,b)=>b[1]-a[1]).slice(0,n);}

/* ---------- tooltip ---------- */
function showTip(html,ev){const t=$("tip");t.innerHTML=html;t.style.display="block";moveTip(ev);}
function moveTip(ev){const t=$("tip");const w=t.offsetWidth,h=t.offsetHeight;let x=ev.clientX+14,y=ev.clientY+14;
  if(x+w>window.innerWidth-8)x=ev.clientX-w-14;if(y+h>window.innerHeight-8)y=ev.clientY-h-14;
  t.style.left=x+"px";t.style.top=y+"px";}
function hideTip(){$("tip").style.display="none";}
function bindTip(el,htmlFn){el.addEventListener("mouseenter",e=>showTip(htmlFn(),e));el.addEventListener("mousemove",moveTip);el.addEventListener("mouseleave",hideTip);}

/* ---------- svg helpers ---------- */
const NS="http://www.w3.org/2000/svg";
function svgEl(tag,attrs,parent){const e=document.createElementNS(NS,tag);for(const a in attrs)e.setAttribute(a,attrs[a]);if(parent)parent.appendChild(e);return e;}
function arcPath(cx,cy,R,r,a0,a1){
  const large=(a1-a0)>Math.PI?1:0;
  const p=(rr,a)=>[cx+rr*Math.cos(a),cy+rr*Math.sin(a)];
  const [x0,y0]=p(R,a0),[x1,y1]=p(R,a1),[x2,y2]=p(r,a1),[x3,y3]=p(r,a0);
  return `M${x0},${y0}A${R},${R} 0 ${large} 1 ${x1},${y1}L${x2},${y2}A${r},${r} 0 ${large} 0 ${x3},${y3}Z`;
}

function donut(elId,items,opts){
  opts=opts||{};
  const el=$(elId);el.innerHTML="";
  const unit=opts.unit||"طن", dec=opts.dec===undefined?1:opts.dec;
  const total=items.reduce((s,i)=>s+i.value,0);
  const wrap=document.createElement("div");wrap.className="donut-wrap";el.appendChild(wrap);
  const cx=90,cy=90,R=82,r=54;
  const svg=svgEl("svg",{viewBox:"0 0 180 180",width:"180",height:"180"});svg.style.flexShrink="0";wrap.appendChild(svg);
  if(total<=0){
    svgEl("circle",{cx:cx,cy:cy,r:(R+r)/2,fill:"none",stroke:"rgba(148,163,184,.15)","stroke-width":R-r},svg);
    const tx=svgEl("text",{x:cx,y:cy+4,"text-anchor":"middle",fill:"#5f7288","font-size":"12"},svg);tx.textContent="لا توجد بيانات";
  }else{
    let a0=-Math.PI/2;
    items.forEach(i=>{
      if(i.value<=0)return;
      const frac=i.value/total;let a1=a0+frac*2*Math.PI;if(frac>0.9999)a1=a0+2*Math.PI-0.0001;
      const p=svgEl("path",{d:arcPath(cx,cy,R,r,a0,a1),fill:i.color,stroke:"#ffffff","stroke-width":"1.5",cursor:opts.onSelect?"pointer":"default"},svg);
      p.addEventListener("mouseenter",()=>p.setAttribute("opacity","0.8"));
      p.addEventListener("mouseleave",()=>p.removeAttribute("opacity"));
      bindTip(p,()=>`<b style="color:${i.color}">${esc(i.label)}</b><br>${nf(i.value,dec)} ${unit} · ${pf(frac)}`);
      if(opts.onSelect)p.addEventListener("click",()=>opts.onSelect(i.key));
      a0=a1;
    });
    const t1=svgEl("text",{x:cx,y:cy-1,"text-anchor":"middle",fill:"#0d47a1","font-size":"19","font-weight":"700"},svg);t1.textContent=nf(total,dec);
    const t2=svgEl("text",{x:cx,y:cy+17,"text-anchor":"middle",fill:"#5f7288","font-size":"10.5"},svg);t2.textContent=unit;
  }
  const lg=document.createElement("div");lg.className="legend";wrap.appendChild(lg);
  items.forEach(i=>{
    const row=document.createElement("div");row.className="lg-row"+(opts.onSelect?" clickable":"");
    row.innerHTML=`<span class="dot" style="background:${i.color}"></span><span class="lg-l">${esc(i.label)}</span><span class="lg-v">${nf(i.value,dec)} ${unit}</span><span class="lg-p">${total>0?pf(i.value/total):"—"}</span>`;
    if(opts.onSelect)row.addEventListener("click",()=>opts.onSelect(i.key));
    lg.appendChild(row);
  });
}

function hbars(elId,items,opts){
  opts=opts||{};
  const el=$(elId);el.innerHTML="";
  if(!items.length){el.innerHTML='<div class="nodata">لا توجد بيانات مطابقة للفلاتر</div>';return;}
  const max=Math.max.apply(null,items.map(i=>i.value).concat([1e-9]));
  items.forEach(it=>{
    const row=document.createElement("div");row.className="hrow"+(opts.onClick?" clickable":"");
    const lab=document.createElement("div");lab.className="hlabel";lab.textContent=it.label;lab.title=it.label;
    const tr=document.createElement("div");tr.className="htrack";
    const bar=document.createElement("i");bar.style.width=(it.value/max*100)+"%";bar.style.background=it.color||"linear-gradient(90deg,#1976d2,#0d47a1)";
    tr.appendChild(bar);
    const val=document.createElement("div");val.className="hval";val.textContent=nf(it.value,1)+" طن"+(it.sub?" · "+it.sub:"");
    row.appendChild(lab);row.appendChild(tr);row.appendChild(val);
    if(it.tip)bindTip(row,()=>it.tip);
    if(opts.onClick)row.addEventListener("click",()=>opts.onClick(it));
    el.appendChild(row);
  });
}

function cumChart(k){
  const el=$("cumChart");el.innerHTML="";
  const W=760,H=250,padL=46,padR=12,padT=12,padB=24;
  const iw=W-padL-padR,ih=H-padT-padB;
  const yMax=Math.max(HAS_PLAN?PLAN:0,k.cum[MDAYS]||0,1)*1.08;
  const X=d=>padL+(d-1)/(MDAYS-1)*iw;
  const Y=v=>padT+ih*(1-v/yMax);
  const svg=svgEl("svg",{viewBox:`0 0 ${W} ${H}`});svg.style.width="100%";el.appendChild(svg);
  const defs=svgEl("defs",{},svg);
  defs.innerHTML='<linearGradient id="cumg" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#1976d2" stop-opacity=".38"/><stop offset="1" stop-color="#1976d2" stop-opacity="0"/></linearGradient>';
  for(let i=0;i<=5;i++){
    const v=yMax*i/5,y=Y(v);
    svgEl("line",{x1:padL,x2:W-padR,y1:y,y2:y,stroke:"rgba(13,71,161,.09)","stroke-width":1},svg);
    const t=svgEl("text",{x:padL-7,y:y+3.5,"text-anchor":"end",fill:"#5f7288","font-size":"10"},svg);t.textContent=n0(v);
  }
  [1,5,10,15,20,25,MDAYS].forEach(d=>{if(d<=MDAYS){const t=svgEl("text",{x:X(d),y:H-7,"text-anchor":"middle",fill:"#5f7288","font-size":"10"},svg);t.textContent=d;}});
  if(HAS_PLAN)svgEl("path",{d:`M${X(1)},${Y(PLAN/MDAYS)}L${X(MDAYS)},${Y(PLAN)}`,stroke:"#d97706","stroke-width":"1.8","stroke-dasharray":"6 5",fill:"none"},svg);
  const md=k.maxDay||0;
  if(md>0){
    const pts=[];for(let d=1;d<=md;d++)pts.push(`${X(d)},${Y(k.cum[d])}`);
    svgEl("path",{d:`M${X(1)},${Y(0)}L${pts.join("L")}L${X(md)},${Y(0)}Z`,fill:"url(#cumg)"},svg);
    svgEl("path",{d:`M${pts.join("L")}`,stroke:"#1976d2","stroke-width":"2.4",fill:"none","stroke-linejoin":"round","stroke-linecap":"round"},svg);
    for(let d=1;d<=md;d++)svgEl("circle",{cx:X(d),cy:Y(k.cum[d]),r:3.2,fill:"#1976d2",stroke:"#ffffff","stroke-width":1.5},svg);
  }
  for(let d=1;d<=MDAYS;d++){
    const rz=svgEl("rect",{x:X(d)-iw/MDAYS/2,y:padT,width:iw/MDAYS,height:ih,fill:"transparent"},svg);
    bindTip(rz,()=>{
      const ca=k.cum[d]||0;
      let s=`<b>يوم ${d} ${esc(CONFIG.month_ar)}</b><br>الفعلي التراكمي: ${nf(ca,1)} طن`;
      if(HAS_PLAN){
        const cp=PLAN*d/MDAYS;
        s+=`<br>الخطة التراكمية: ${nf(cp,1)} طن<br>الفرق: <span style="color:${ca>=cp?"#0b8a4b":"#d32f2f"}">${ca>=cp?"+":""}${nf(ca-cp,1)} طن</span>`;
      }
      return s;
    });
  }
}

function dailyChart(k){
  const el=$("dailyChart");el.innerHTML="";
  let max=1e-9;for(let d=1;d<=MDAYS;d++)max=Math.max(max,k.dailyPow[d]+k.dailyTel[d]);
  const cols=document.createElement("div");cols.className="cols";el.appendChild(cols);
  const labs=document.createElement("div");labs.className="dlabs";el.appendChild(labs);
  for(let d=1;d<=MDAYS;d++){
    const p=k.dailyPow[d],t=k.dailyTel[d],tot=p+t;
    const col=document.createElement("div");col.className="dcol";
    const st=document.createElement("div");st.className="dstack";
    if(t>0){const s=document.createElement("i");s.className="seg";s.style.height=(t/max*100)+"%";s.style.background="linear-gradient(180deg,#f59e0b,#d97706)";st.appendChild(s);}
    if(p>0){const s=document.createElement("i");s.className="seg";s.style.height=(p/max*100)+"%";s.style.background="linear-gradient(180deg,#42a5f5,#1976d2)";st.appendChild(s);}
    col.appendChild(st);cols.appendChild(col);
    const lb=document.createElement("div");lb.className="dlab"+(tot>0?"":" off");lb.textContent=d;labs.appendChild(lb);
    bindTip(col,()=>`<b>يوم ${d} ${esc(CONFIG.month_ar)}</b><br>Power: ${nf(p,1)} طن<br>Telecom: ${nf(t,1)} طن<br>الإجمالي: ${nf(tot,1)} طن · ${k.dailyPacks[d]} طرد`);
  }
}

/* ---------- KPIs ---------- */
function renderKPIs(k){
  $("kActual").textContent=nf(k.totalT,1);
  if(HAS_PLAN){
    $("kActualS").innerHTML=`من خطة <b>${n0(PLAN)}</b> طن (${pf(k.donePct)}) · العلامة البيضاء = المتوقع بعد ${k.maxDay} أيام (${pf(k.expected)})`;
    $("kActualBar").style.width=Math.min(100,k.ratio*100)+"%";
    $("kActualMark").style.display="";
    $("kActualMark").style.left=Math.min(100,k.expected*100)+"%";
  }else{
    $("kActualS").innerHTML=`${n0(k.packs)} طرد · سجل الخطة في <b>plan.json</b> لمتابعة التحقيق`;
    $("kActualBar").style.width="0%";$("kActualMark").style.display="none";
  }
  if(!HAS_PLAN){
    $("kRem").textContent="—";
    $("kRemS").innerHTML="أدخل خطة الشهر في <b>plan.json</b>";
    $("kRemBar").style.width="0%";
  }else if(k.remaining<=0.0001){
    $("kRem").innerHTML='<span class="good">0 🎉</span>';
    $("kRemS").innerHTML='<b class="good">الخطة اتحققت بالكامل!</b>';
    $("kRemBar").style.width="0%";
  }else{
    $("kRem").textContent=nf(k.remaining,1);
    $("kRemBar").style.width=Math.min(100,(k.remaining/PLAN)*100)+"%";
    if(k.daysLeft<=0){
      $("kRemS").innerHTML=`<span class="bad">الشهر انتهى — فاضل ${pf(k.remaining/PLAN)} من الخطة</span>`;
    }else{
      $("kRemS").innerHTML=`باقي <b>${k.daysLeft}</b> يوم · مطلوب <b>${nf(k.paceToFinish,1)}</b> طن/يوم — `+
        (k.feasible?`<span class="good">ممكن بالوتيرة الحالية ✓</span>`:`<span class="bad">محتاج تسريع ${pf(k.speedUpPct)} عن الحالي</span>`);
    }
  }
  if(HAS_PLAN){
    const behind=k.behind>0.5;
    $("kRatio").innerHTML=`<span class="${behind?"bad":"good"}">${pf(k.ratio)}</span>`;
    $("kRatioS").innerHTML=`المتوقع في اليوم ${k.maxDay}: ${pf(k.expected)} — <span class="${behind?"bad":"good"}">${behind?"متأخر":"متقدم"} ${nf(Math.abs(k.behind),1)} طن</span>`;
  }else{
    $("kRatio").textContent="—";$("kRatioS").textContent="بدون خطة مسجلة";
  }
  const gap=k.dailyAvg-k.paceReq;
  $("kPace").textContent=nf(k.dailyAvg,1);
  $("kPaceS").innerHTML=(HAS_PLAN?`المطلوب للخطة: ${nf(k.paceReq,1)} طن/يوم — <span class="${gap>=0?"good":"bad"}">${gap>=0?"+":""}${nf(gap,1)}</span>`:`معدل ${MDAYS} يوم`)+` · ${k.nDays} أيام إنتاج`+(HAS_PLAN&&k.remaining>0?` · لإنهاء الخطة: <b>${nf(k.paceToFinish,1)}</b> طن/يوم`:"");
  $("kFc").textContent=n0(k.forecast);
  $("kFcS").innerHTML=HAS_PLAN?`على المعدل الحالي = <b class="${k.forecastRatio>=1?"good":"warn"}">${pf(k.forecastRatio)}</b> من الخطة · باقي ${n0(k.remaining)} طن`:`إسقاط على ${MDAYS} يوم بنفس الوتيرة`;
  $("kKm").textContent=nf(k.km,1);
  $("kKmS").innerHTML=`${n0(k.packs)} طرد · ${k.wos.size} أمر تشغيل`;
  const expShare=k.totalT>0?k.expT/k.totalT:0;
  $("kExp").textContent=pf(expShare);
  $("kExpS").innerHTML=`تصدير ${nf(k.expT,1)} طن · محلي ${nf(k.locT,1)} طن`;
  $("kHold").innerHTML=`<span class="${k.holdN>0?"warn":"good"}">${k.holdN}</span>`;
  $("kHoldS").innerHTML=`${nf(k.holdT,1)} طن موقوفة · ${k.pN} طرد بدون حالة QC`;
  $("kCust").textContent=k.custs.size;
  $("kCustS").innerHTML=`عميل نشط · ${k.days.size} أيام تشغيل من ${MDAYS}`;
  $("subDate").innerHTML=`${esc(CONFIG.month_ar)} ${CONFIG.year} · البيانات حتى يوم ${k.maxDay||0} · ${n0(DATA.length)} طرد · <span style="color:#1976d2">EPC — Energya Power Cables</span>`;
}

/* ---------- slicers ---------- */
function chipEl(label,active,dim,onclick){
  const b=document.createElement("button");
  b.className="chip"+(active?" on":"")+(dim?" dim":"");
  b.innerHTML=label;b.onclick=onclick;return b;
}
function toggleIn(set,v){set.has(v)?set.delete(v):set.add(v);renderAll();}
function resetFilters(){state.days.clear();state.fams.clear();state.cust="";state.mac="";state.mkt="all";state.con="all";state.qc="all";tstate.page=1;renderAll();}

function renderSlicers(){
  const dayCounts=groupCount(DATA,r=>r.d);
  const sd=$("slDays");sd.innerHTML="";
  sd.appendChild(chipEl("الكل",state.days.size===0,false,()=>{state.days.clear();renderAll();}));
  for(let d=1;d<=MDAYS;d++){
    sd.appendChild(chipEl(String(d),state.days.has(d),!dayCounts.has(d),()=>toggleIn(state.days,d)));
  }
  const sf=$("slFams");sf.innerHTML="";
  sf.appendChild(chipEl("الكل",state.fams.size===0,false,()=>{state.fams.clear();renderAll();}));
  FAM_ORDER.forEach(f=>{
    sf.appendChild(chipEl(`<span class="dot" style="background:${FAM_COLOR[f]}"></span>${f}`,state.fams.has(f),false,()=>toggleIn(state.fams,f)));
  });
  const mk=[["all","الكل"],["exp","🌍 Export"],["loc","🏠 Local"]];
  const sm=$("slMkt");sm.innerHTML="";
  mk.forEach(([v,l])=>sm.appendChild(chipEl(l,state.mkt===v,false,()=>{state.mkt=v;renderAll();})));
  const cn=[["all","الكل"],["A","AL ألومنيوم"],["C","CU نحاس"]];
  const sc=$("slCon");sc.innerHTML="";
  cn.forEach(([v,l])=>sc.appendChild(chipEl(l,state.con===v,false,()=>{state.con=v;renderAll();})));
  const qc=[["all","الكل"],["OK","OK"],["Hold","Hold"],["P","بدون حالة"]];
  const sq=$("slQc");sq.innerHTML="";
  qc.forEach(([v,l])=>sq.appendChild(chipEl(l,state.qc===v,false,()=>{state.qc=v;renderAll();})));
  $("selCust").value=state.cust;
  $("selMac").value=state.mac;
}

function renderFilterBar(rows,k){
  const fb=$("filterBar");fb.innerHTML="";
  const add=(label,clear)=>{
    const c=document.createElement("span");c.className="fchip";
    c.innerHTML=`${esc(label)} <b>✕</b>`;
    c.querySelector("b").onclick=()=>{clear();renderAll();};
    fb.appendChild(c);
  };
  if(state.days.size)add("أيام: "+[...state.days].sort((a,b)=>a-b).join("، "),()=>state.days.clear());
  if(state.fams.size)add("عائلات: "+[...state.fams].join("، "),()=>state.fams.clear());
  if(state.mkt!=="all")add(state.mkt==="exp"?"تصدير":"محلي",()=>state.mkt="all");
  if(state.con!=="all")add(state.con==="A"?"ألومنيوم":"نحاس",()=>state.con="all");
  if(state.qc!=="all")add("QC: "+QC_LABEL[state.qc],()=>state.qc="all");
  if(state.cust)add("عميل: "+state.cust,()=>{state.cust="";$("selCust").value="";});
  if(state.mac)add("ماكينة: "+state.mac,()=>{state.mac="";$("selMac").value="";});
  const info=document.createElement("span");
  info.style.cssText="font-size:11.5px;color:var(--mut)";
  info.textContent=`النتيجة: ${n0(k.packs)} من ${DATA.length} طرد · ${nf(k.totalT,1)} طن`;
  fb.appendChild(info);
}

/* ---------- table ---------- */
const COLS=[
  {k:"d",l:"اليوم"},{k:"pk",l:"الطرد"},{k:"wo",l:"أمر التشغيل"},{k:"cu",l:"العميل"},
  {k:"fam",l:"العائلة"},{k:"vs",l:"المواصفة"},{k:"mac",l:"الماكينة"},{k:"q",l:"أمتار"},
  {k:"tt",l:"طن"},{k:"exp",l:"السوق"},{k:"qc",l:"QC"}
];
function renderTable(rows,k){
  const sorted=rows.slice().sort((a,b)=>{
    const key=tstate.sortKey,va=a[key],vb=b[key];
    if(typeof va==="number"&&typeof vb==="number")return (va-vb)*tstate.sortDir;
    return String(va).localeCompare(String(vb),"ar")*tstate.sortDir;
  });
  const pages=Math.max(1,Math.ceil(sorted.length/tstate.pageSize));
  if(tstate.page>pages)tstate.page=pages;
  const start=(tstate.page-1)*tstate.pageSize;
  const pageRows=sorted.slice(start,start+tstate.pageSize);
  let h='<table><thead><tr>';
  COLS.forEach(c=>{
    const arrow=tstate.sortKey===c.k?(tstate.sortDir>0?" ▲":" ▼"):"";
    h+=`<th data-k="${c.k}">${c.l}${arrow}</th>`;
  });
  h+="</tr></thead><tbody>";
  pageRows.forEach(r=>{
    h+=`<tr><td>${r.d}</td><td>${esc(r.pk)}</td><td>${esc(r.wo)}</td><td>${esc(r.cu)}</td>`+
       `<td><span class="bdg" style="background:${FAM_COLOR[r.fam]||"#64748b"}22;color:${FAM_COLOR[r.fam]||"#94a3b8"}">${esc(r.fam)}</span></td>`+
       `<td>${esc(r.vs)}</td><td>${esc(r.mac)}</td><td class="num">${n0(r.q)}</td><td class="num">${nf(r.tt,3)}</td>`+
       `<td><span class="bdg" style="background:${r.exp==="Y"?"#0f9d5822;color:#0b8a4b":"#1976d222;color:#1266b0"}">${r.exp==="Y"?"Export":"Local"}</span></td>`+
       `<td><span class="bdg" style="background:${QC_COLOR[r.qc]}22;color:${QC_COLOR[r.qc]}">${r.qc==="OK"?"OK":r.qc==="Hold"?"Hold":"—"}</span></td></tr>`;
  });
  h+="</tbody></table>";
  $("tableWrap").innerHTML=h;
  $("tableWrap").querySelectorAll("th").forEach(th=>{
    th.onclick=()=>{
      const kk=th.dataset.k;
      if(tstate.sortKey===kk)tstate.sortDir*=-1;else{tstate.sortKey=kk;tstate.sortDir=1;}
      renderTable(rows,k);
    };
  });
  $("tblHint").textContent=`عرض ${sorted.length?start+1:0}–${Math.min(start+tstate.pageSize,sorted.length)} من ${sorted.length} طرد (${nf(k.totalT,1)} طن · ${nf(k.km,1)} كم) — اضغط على عنوان أي عمود للترتيب`;
  const pg=$("pager");pg.innerHTML="";
  const prev=document.createElement("button");prev.textContent="→ السابق";prev.disabled=tstate.page<=1;
  prev.onclick=()=>{tstate.page--;renderTable(rows,k);};
  const next=document.createElement("button");next.textContent="التالي ←";next.disabled=tstate.page>=pages;
  next.onclick=()=>{tstate.page++;renderTable(rows,k);};
  const info=document.createElement("span");info.textContent=`صفحة ${tstate.page} من ${pages}`;
  pg.appendChild(prev);pg.appendChild(info);pg.appendChild(next);
}

/* ---------- charts render ---------- */
function renderCharts(rows,k){
  cumChart(k);
  dailyChart(k);
  donut("famDonut",FAM_ORDER.filter(f=>(k.famT[f]||0)>0).map(f=>({key:f,label:FAM_NAME[f],value:k.famT[f],color:FAM_COLOR[f]})),
    {onSelect:f=>toggleIn(state.fams,f)});
  donut("mktDonut",[
    {key:"Y",label:"Export — تصدير",value:k.expT,color:"#0f9d58"},
    {key:"N",label:"Local — محلي",value:k.locT,color:"#1976d2"}
  ],{onSelect:v=>{state.mkt=v==="Y"?(state.mkt==="exp"?"all":"exp"):(state.mkt==="loc"?"all":"loc");renderAll();}});
  donut("qcDonut",[
    {key:"OK",label:QC_LABEL.OK,value:k.okN,color:QC_COLOR.OK},
    {key:"Hold",label:QC_LABEL.Hold,value:k.holdN,color:QC_COLOR.Hold},
    {key:"P",label:QC_LABEL.P,value:k.pN,color:QC_COLOR.P}
  ],{unit:"طرد",dec:0,onSelect:v=>{state.qc=state.qc===v?"all":v;renderAll();}});
  const custT=groupTons(rows,r=>r.cu),custC=groupCount(rows,r=>r.cu);
  hbars("custBars",topFrom(custT,12).map(([cu,v])=>({
    key:cu,label:cu,value:v,sub:custC.get(cu)+" طرد",
    color:state.cust===cu?"linear-gradient(90deg,#0f9d58,#0b8a4b)":null,
    tip:`<b>${esc(cu)}</b><br>${nf(v,2)} طن · ${pf(k.totalT?v/k.totalT:0)} من المعروض<br>${custC.get(cu)} طرد — اضغط للفلترة`
  })),{onClick:it=>{state.cust=state.cust===it.key?"":it.key;tstate.page=1;renderAll();}});
  const macT=groupTons(rows,r=>r.mac),macC=groupCount(rows,r=>r.mac);
  hbars("macBars",topFrom(macT,12).map(([m,v])=>({
    key:m,label:m,value:v,sub:macC.get(m)+" طرد",
    color:state.mac===m?"linear-gradient(90deg,#0f9d58,#0b8a4b)":null,
    tip:`<b>ماكينة ${esc(m)}</b><br>${nf(v,2)} طن · ${pf(k.totalT?v/k.totalT:0)}<br>${macC.get(m)} طرد — اضغط للفلترة`
  })),{onClick:it=>{state.mac=state.mac===it.key?"":it.key;tstate.page=1;renderAll();}});
  const tot=k.totalT||1;
  hbars("matBars",[
    {label:"AL — ألومنيوم",value:k.al,color:MAT_COLOR.AL,sub:pf(k.al/tot)+(PLAN_AL?` · ${pf(PLAN_AL?k.al/PLAN_AL:0)} من الخطة`:""),tip:`<b>AL — ألومنيوم</b><br>${nf(k.al,2)} طن (${pf(k.al/tot)} من الإجمالي)`+(PLAN_AL?`<br>خطة AL: ${n0(PLAN_AL)} طن — تحقق ${pf(k.al/PLAN_AL)} · فاضل ${nf(Math.max(0,PLAN_AL-k.al),1)} طن`:"")},
    {label:"CU — نحاس",value:k.cu,color:MAT_COLOR.CU,sub:pf(k.cu/tot)+(PLAN_CU?` · ${pf(PLAN_CU?k.cu/PLAN_CU:0)} من الخطة`:""),tip:`<b>CU — نحاس</b><br>${nf(k.cu,2)} طن (${pf(k.cu/tot)} من الإجمالي)`+(PLAN_CU?`<br>خطة CU: ${n0(PLAN_CU)} طن — تحقق ${pf(k.cu/PLAN_CU)} · فاضل ${nf(Math.max(0,PLAN_CU-k.cu),1)} طن`:"")},
    {label:"Screen",value:k.st,color:MAT_COLOR.SCR,sub:pf(k.st/tot),tip:`<b>Screen</b><br>${nf(k.st,2)} طن (${pf(k.st/tot)})`},
    {label:"Armour",value:k.at,color:MAT_COLOR.ARM,sub:pf(k.at/tot),tip:`<b>Armour</b><br>${nf(k.at,2)} طن (${pf(k.at/tot)})`}
  ]);
}

/* ---------- banner messages ---------- */
let DEFAULT_BANNER="";
function showBanner(html,type){
  const b=$("banner");b.style.display="flex";
  b.style.borderColor=type==="ok"?"rgba(15,157,88,.55)":type==="err"?"rgba(220,38,38,.55)":"rgba(25,118,210,.45)";
  b.style.background=type==="ok"?"#e9f8f0":type==="err"?"#fdecec":"#e8f2fd";
  $("bannerTxt").innerHTML=html;
}
function nowStr(){const d=new Date();const p=n=>String(n).padStart(2,"0");return `${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`;}

/* ---------- localStorage (guarded for sandboxed previews) ---------- */
const LS_PLAN="epcDashPlan_v1";
function savePlanLS(){try{localStorage.setItem(LS_PLAN,JSON.stringify({t:CONFIG.plan_total,a:CONFIG.plan_al,c:CONFIG.plan_cu}));}catch(e){}}
function loadPlanLS(){try{const s=localStorage.getItem(LS_PLAN);if(!s)return false;const o=JSON.parse(s);CONFIG.plan_total=Number(o.t)||0;CONFIG.plan_al=Number(o.a)||0;CONFIG.plan_cu=Number(o.c)||0;return true;}catch(e){return false;}}
function clearPlanLS(){try{localStorage.removeItem(LS_PLAN);}catch(e){}}

/* ---------- ERP file parsing (mirrors update_dashboard.py) ---------- */
const REQUIRED=["wo","export","cond_kg","screen_kgm","armour_kgm","prodfamily","customer","packno","packqty","prdate","maccode","conductor","qcstatus"];
function parseErpRows(rows,fname){
  let hi=-1,idx=null;
  for(let i=0;i<Math.min(rows.length,10);i++){
    const cells=new Set((rows[i]||[]).map(v=>v==null?"":String(v).trim().toLowerCase()));
    if(REQUIRED.every(c=>cells.has(c))){
      hi=i;idx={};
      (rows[i]||[]).forEach((v,j)=>{if(v!=null)idx[String(v).trim().toLowerCase()]=j;});
      break;
    }
  }
  if(hi<0)throw new Error("لم أجد صف العناوين المطلوب (wo / packno / prdate / cond_kg …) — تأكد إن ده ملف تصدير الـ ERP الصحيح");
  const g=(r,n)=>{const j=idx[n];return (j===undefined||r[j]===undefined)?null:r[j];};
  const sval=v=>v==null?"":(typeof v==="number"&&Number.isInteger(v)?String(v):String(v).trim());
  const fnum=v=>{const n=Number(v);return isFinite(n)?n:0;};
  const pdate=v=>{
    if(v==null||v==="")return null;
    if(v instanceof Date&&!isNaN(v.getTime()))return{y:v.getFullYear(),m:v.getMonth()+1,d:v.getDate()};
    if(typeof v==="number"&&v>20000){const t=new Date(Date.UTC(1899,11,30)+v*86400000);return{y:t.getUTCFullYear(),m:t.getUTCMonth()+1,d:t.getUTCDate()};}
    if(typeof v==="string"){
      let mm=v.trim().match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);if(mm)return{y:+mm[1],m:+mm[2],d:+mm[3]};
      mm=v.trim().match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})/);if(mm)return{y:+mm[3],m:+mm[2],d:+mm[1]};
    }
    return null;
  };
  const data=[];let skipped=0;
  for(let i=hi+1;i<rows.length;i++){
    const r=rows[i]||[];
    const dt=pdate(g(r,"prdate")),pk=sval(g(r,"packno"));
    if(!dt||pk===""){skipped++;continue;}
    const q=fnum(g(r,"packqty"));
    const ct=q*fnum(g(r,"cond_kg"))/1000;
    const st=q*fnum(g(r,"screen_kgm"))/1000;
    const at=q*fnum(g(r,"armour_kgm"))/1000;
    data.push({d:dt.d,pk:pk,wo:sval(g(r,"wo")),cu:sval(g(r,"customer"))||"?",fam:sval(g(r,"prodfamily"))||"?",
      con:sval(g(r,"conductor"))||"?",exp:sval(g(r,"export")).toUpperCase()==="Y"?"Y":"N",
      q:Math.round(q),ct:+ct.toFixed(4),st:+st.toFixed(4),at:+at.toFixed(4),tt:+(ct+st+at).toFixed(4),
      mac:sval(g(r,"maccode"))||"?",pt:sval(g(r,"packtype"))||"?",qc:sval(g(r,"qcstatus"))||"P",
      vs:(sval(g(r,"volt"))+" "+sval(g(r,"size"))).trim(),_ym:dt.y*100+dt.m});
  }
  if(!data.length)throw new Error("مفيش صفوف إنتاج صالحة في الملف (لازم يكون فيه prdate و packno)");
  const cnt=new Map();data.forEach(r=>cnt.set(r._ym,(cnt.get(r._ym)||0)+1));
  let ym=0,best=-1;cnt.forEach((v,k)=>{if(v>best){best=v;ym=k;}});
  const year=Math.floor(ym/100),month=ym%100;
  data.forEach(r=>delete r._ym);
  const monthDays=new Date(year,month,0).getDate();
  return {data:data,year:year,month:month,monthDays:monthDays,skipped:skipped,fname:fname};
}

function applyImport(res){
  const newKey=res.year+"-"+String(res.month).padStart(2,"0");
  let archNote="";
  if(DATA.length){
    const oldKey=(ACTIVE&&ACTIVE!=="__embed__")?ACTIVE:cfgMonthKey(CONFIG);
    if(oldKey&&oldKey!==newKey){
      storeMonth(oldKey,DATA,CONFIG);
      archNote=` · 🗄️ اتأرشف <b>${esc(monthLabel(oldKey))}</b> تلقائياً`;
    }
  }
  DATA=res.data;
  ACTIVE=newKey;
  CONFIG=Object.assign({},CONFIG,{month_ar:AR_MONTHS_JS[res.month]||String(res.month),year:res.year,month_num:res.month,month_days:res.monthDays,source:res.fname,generated:nowStr()});
  if(MONTHS[newKey]&&MONTHS[newKey].plan&&MONTHS[newKey].plan.t){
    CONFIG.plan_total=MONTHS[newKey].plan.t;CONFIG.plan_al=MONTHS[newKey].plan.a;CONFIG.plan_cu=MONTHS[newKey].plan.c;
  }
  syncPlanVars();
  resetFiltersSilent();
  buildSelects();
  renderAll();
  storeMonth(newKey,DATA,CONFIG);
  rebuildMonthSelect();
  const k=computeKPIs(DATA);
  showBanner(`✅ تم استيراد <b>${esc(res.fname)}</b> — ${n0(DATA.length)} طرد · ${nf(k.totalT,1)} طن · ${esc(CONFIG.month_ar)} ${CONFIG.year}`+
    (res.skipped?` · (${n0(res.skipped)} صف اتجاوز) `:" ")+
    (HAS_PLAN?` · خلص <b>${pf(k.donePct)}</b> من الخطة، فاضل <b>${nf(k.remaining,1)}</b> طن`:"")+
    archNote+
    ` — <span id="restoreLink" style="cursor:pointer;text-decoration:underline">عرض بيانات الملف المدمجة</span>`,"ok");
  const rl=$("restoreLink");if(rl)rl.onclick=()=>loadEmbedded(false);
}

function importFile(file){
  if(!CAN_IMPORT){showBanner("⚠️ مكتبة قراءة الإكسيل مش مدمجة في النسخة دي — استخدم السكريبت update_dashboard.py","err");return;}
  showBanner(`⏳ جارِ قراءة <b>${esc(file.name)}</b> …`,"info");
  const reader=new FileReader();
  reader.onerror=()=>showBanner("❌ فشل قراءة الملف — جرب تاني","err");
  reader.onload=e=>{
    try{
      const wb=XLSX.read(new Uint8Array(e.target.result),{type:"array",cellDates:true});
      const sn=wb.SheetNames.includes("0040_qa")?"0040_qa":wb.SheetNames[0];
      const rows=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,raw:true,defval:null});
      applyImport(parseErpRows(rows,file.name));
    }catch(err){
      showBanner("❌ "+esc(err&&err.message?err.message:String(err)),"err");
    }
  };
  reader.readAsArrayBuffer(file);
}

/* ---------- month archive ---------- */
const LS_MONTHS="epcMonths_v1", LS_ACTIVE="epcActive_v1";
let MONTHS={}, ACTIVE="__embed__";
let LS_OK=null;
function lsAvailable(){
  if(LS_OK!==null)return LS_OK;
  try{localStorage.setItem("__epc_t","1");localStorage.removeItem("__epc_t");LS_OK=true;}
  catch(e){LS_OK=false;}
  return LS_OK;
}
function cfgMonthNum(c){const n=Number(c.month_num);if(n)return n;return Math.max(0,AR_MONTHS_JS.indexOf(c.month_ar));}
function cfgMonthKey(c){const m=cfgMonthNum(c);return m?c.year+"-"+String(m).padStart(2,"0"):"";}
function monthLabel(key){const parts=String(key).split("-");const y=+parts[0],m=+parts[1];return (AR_MONTHS_JS[m]||m)+" "+y;}
function persistMonths(){
  if(!lsAvailable())return;
  try{localStorage.setItem(LS_MONTHS,JSON.stringify(MONTHS));localStorage.setItem(LS_ACTIVE,ACTIVE||"");}
  catch(e){showBanner("⚠️ مساحة حفظ المتصفح امتلت — نزّل نسخة احتياطية بزرار ⬇️ نسخة وامسح شهور قديمة من القائمة.","err");}
}
function loadMonths(){
  try{
    const s=localStorage.getItem(LS_MONTHS);
    if(s){const o=JSON.parse(s);if(o&&typeof o==="object")MONTHS=o;}
    const a=localStorage.getItem(LS_ACTIVE);
    if(a&&MONTHS[a])ACTIVE=a;else ACTIVE="__embed__";
  }catch(e){ACTIVE="__embed__";}
}
function storeMonth(key,data,cfg){
  if(!key)return;
  MONTHS[key]={data:data,cfg:{month_ar:cfg.month_ar,year:cfg.year,month_num:cfgMonthNum(cfg),month_days:cfg.month_days,source:cfg.source,generated:cfg.generated},plan:{t:cfg.plan_total,a:cfg.plan_al,c:cfg.plan_cu}};
  persistMonths();
}
function rebuildMonthSelect(){
  const sel=$("selMonth");if(!sel)return;
  const keys=Object.keys(MONTHS).sort().reverse();
  let h="";
  if(ACTIVE==="__embed__")h+=`<option value="__embed__">${esc(monthLabel(cfgMonthKey(ORIG_CONFIG)))} (الملف الأصلي)</option>`;
  keys.forEach(k=>{h+=`<option value="${esc(k)}">🗄️ ${esc(monthLabel(k))} (${n0((MONTHS[k].data||[]).length)} طرد)</option>`;});
  h+=`<option value="__new__">🆕 شهر جديد (فارغ)</option>`;
  sel.innerHTML=h;
  if(ACTIVE===null)sel.value="__new__";
  else if(ACTIVE==="__embed__")sel.value="__embed__";
  else sel.value=MONTHS[ACTIVE]?ACTIVE:(keys[0]||"__new__");
}
function resetFiltersSilent(){
  state.days.clear();state.fams.clear();state.cust="";state.mac="";state.mkt="all";state.con="all";state.qc="all";tstate.page=1;
}
function loadMonthEntry(key,silent){
  const e=MONTHS[key];if(!e)return false;
  DATA=e.data||[];
  CONFIG=Object.assign({},ORIG_CONFIG,{month_ar:e.cfg.month_ar,year:e.cfg.year,month_num:e.cfg.month_num,month_days:e.cfg.month_days,source:e.cfg.source,generated:e.cfg.generated,plan_total:e.plan.t,plan_al:e.plan.a,plan_cu:e.plan.c});
  ACTIVE=key;
  syncPlanVars();
  resetFiltersSilent();
  buildSelects();applyPlanUI();rebuildMonthSelect();renderAll();
  if(!silent)showBanner(`🗄️ تم فتح أرشيف <b>${esc(monthLabel(key))}</b> — ${n0(DATA.length)} طرد · ${nf(computeKPIs(DATA).totalT,1)} طن`,"info");
  return true;
}
function loadEmbedded(silent){
  DATA=ORIG_DATA;
  CONFIG=Object.assign({},ORIG_CONFIG);
  loadPlanLS();
  ACTIVE="__embed__";
  syncPlanVars();
  resetFiltersSilent();
  buildSelects();applyPlanUI();rebuildMonthSelect();renderAll();
  if(!silent)$("banner").style.display="none";
}
function restoreEmbedded(){loadEmbedded(false);}
let nmConfirm=false;
function startNewMonth(){
  const b=$("btnNewMonth");
  if(!nmConfirm){
    nmConfirm=true;b.classList.add("arm");b.textContent="⚠️ تأكيد: اضغط تاني";
    setTimeout(()=>{nmConfirm=false;b.classList.remove("arm");b.textContent="🆕 شهر جديد";},4000);
    return;
  }
  nmConfirm=false;b.classList.remove("arm");b.textContent="🆕 شهر جديد";
  let note="";
  if(DATA.length){
    const k=ACTIVE&&ACTIVE!=="__embed__"?ACTIVE:cfgMonthKey(CONFIG);
    if(k){storeMonth(k,DATA,CONFIG);note=` — تم حفظ <b>${esc(monthLabel(k))}</b> في الأرشيف 🗄️ وتقدر ترجعله من القائمة فوق`;}
  }
  DATA=[];ACTIVE=null;
  CONFIG=Object.assign({},CONFIG,{source:"—",generated:nowStr()});
  syncPlanVars();
  resetFiltersSilent();
  buildSelects();applyPlanUI();rebuildMonthSelect();renderAll();
  showBanner(`🆕 الداشبورد فاضي وجاهز للشهر الجديد${note} — اسحب ملف الـ ERP الجديد هنا أو دوس 📂 استيراد.`,"ok");
}
function exportBackup(){
  const payload={app:"EPC-Dashboard",version:1,exported:nowStr(),months:MONTHS,active:ACTIVE};
  try{
    const blob=new Blob([JSON.stringify(payload)],{type:"application/json"});
    const a=document.createElement("a");
    a.href=URL.createObjectURL(blob);
    a.download="EPC-Dashboard-Backup-"+nowStr().slice(0,10)+".json";
    document.body.appendChild(a);a.click();document.body.removeChild(a);
    setTimeout(()=>URL.revokeObjectURL(a.href),60000);
    showBanner("⬇️ اتنزّلت نسخة احتياطية بكل الشهور والخطط — احتفظ بيها في مكان أمين.","ok");
  }catch(e){showBanner("❌ فشل تنزيل النسخة الاحتياطية: "+esc(String(e&&e.message||e)),"err");}
}
function importBackup(file){
  const reader=new FileReader();
  reader.onload=e=>{
    try{
      const o=JSON.parse(String(e.target.result));
      if(!o||typeof o!=="object"||!o.months)throw new Error("الملف مش نسخة احتياطية صالحة");
      let added=0;
      Object.keys(o.months).forEach(k=>{const m=o.months[k];if(m&&Array.isArray(m.data)){MONTHS[k]=m;added++;}});
      persistMonths();
      const act=(o.active&&MONTHS[o.active])?o.active:(Object.keys(MONTHS).sort().reverse()[0]||null);
      if(act)loadMonthEntry(act,true);else rebuildMonthSelect();
      showBanner(`⬆️ تم استرجاع النسخة الاحتياطية — ${added} شهر في الأرشيف`,"ok");
    }catch(err){showBanner("❌ فشل قراءة النسخة الاحتياطية: "+esc(String(err&&err.message||err)),"err");}
  };
  reader.readAsText(file);
}

function applyPlanUI(){
  const lp=$("lgPlan");if(lp)lp.style.display=HAS_PLAN?"":"none";
  const ch=$("cumHint");
  if(ch)ch.textContent=HAS_PLAN
    ?"الخط الأزرق = الفعلي التراكمي (طن كامل: موصلات + Screen + Armour) · الأصفر المتقطع = الخطة موزعة على أيام الشهر · hover للتفاصيل"
    :"الخط الأزرق = الفعلي التراكمي (طن كامل) — سجّل الخطة بزرار 🎯 الخطة فوق لإضافة خط الخطة والمقارنة";
  if(!HAS_PLAN){
    const b=$("banner");b.style.display="flex";b.style.borderColor="rgba(217,119,6,.6)";b.style.background="#fef6e0";
    $("bannerTxt").innerHTML="⚠️ <b>الخطة مش مسجلة</b> — دوس على زرار <b>🎯 الخطة</b> فوق وسجّل أرقام خطة الشهر بالطن (أو عدّل plan.json لو بتستخدم السكريبت). الداشبورد شغال حالياً بعرض الإنتاج فقط بدون مقارنة بالخطة.";
  }
}

/* ---------- main ---------- */
function updateMeta(){
  const bg=$("bGen");if(bg)bg.textContent=CONFIG.generated||"—";
  $("fUpd").textContent=`آخر تحديث: ${CONFIG.generated||"—"} · مصدر البيانات: ${CONFIG.source||"—"} · الخطة: ${HAS_PLAN?n0(PLAN)+" طن":"غير مسجلة"} · بصمة التحقق: ${CONFIG.fingerprint||"—"}`;
}
function renderAll(){
  updateMeta();
  const rows=filteredRows();
  const k=computeKPIs(rows);
  renderKPIs(k);
  renderSlicers();
  renderFilterBar(rows,k);
  renderCharts(rows,k);
  renderTable(rows,k);
}

function buildSelects(){
  const custT=groupTons(DATA,r=>r.cu);
  const selC=$("selCust");
  selC.innerHTML='<option value="">كل العملاء</option>'+
    topFrom(custT,custT.size).map(([c,v])=>`<option value="${esc(c)}">${esc(c)} (${nf(v,1)} طن)</option>`).join("");
  selC.onchange=()=>{state.cust=selC.value;tstate.page=1;renderAll();};
  const macT=groupTons(DATA,r=>r.mac);
  const selM=$("selMac");
  selM.innerHTML='<option value="">كل الماكينات</option>'+
    topFrom(macT,macT.size).map(([m,v])=>`<option value="${esc(m)}">${esc(m)} (${nf(v,1)} طن)</option>`).join("");
  selM.onchange=()=>{state.mac=selM.value;tstate.page=1;renderAll();};
}

function init(){
  DEFAULT_BANNER=$("bannerTxt").innerHTML;
  loadPlanLS();
  loadMonths();
  if(ACTIVE&&ACTIVE!=="__embed__"&&MONTHS[ACTIVE])loadMonthEntry(ACTIVE,true);
  else ACTIVE="__embed__";
  buildSelects();
  document.title=`EPC Production — ${CONFIG.month_ar} ${CONFIG.year}`;
  $("btnReset").onclick=resetFilters;
  $("btnPrint").onclick=()=>window.print();
  $("bannerX").onclick=()=>$("banner").style.display="none";
  /* month controls */
  $("selMonth").onchange=e=>{
    const v=e.target.value;
    if(v==="__new__")startNewMonth();
    else if(v==="__embed__")loadEmbedded(false);
    else loadMonthEntry(v,false);
  };
  $("btnNewMonth").onclick=()=>startNewMonth();
  $("btnBackupDl").onclick=()=>exportBackup();
  $("btnBackupUp").onclick=()=>$("fileBackup").click();
  $("fileBackup").onchange=e=>{if(e.target.files&&e.target.files[0])importBackup(e.target.files[0]);e.target.value="";};
  rebuildMonthSelect();
  /* import wiring */
  const imp=$("btnImport");
  if(!CAN_IMPORT){imp.disabled=true;imp.style.opacity=.45;imp.title="مكتبة قراءة الإكسيل مش مدمجة — استخدم السكريبت";}
  imp.onclick=()=>$("fileInp").click();
  $("fileInp").onchange=e=>{if(e.target.files&&e.target.files[0])importFile(e.target.files[0]);e.target.value="";};
  const ov=$("dropOv");
  ["dragenter","dragover"].forEach(ev=>document.addEventListener(ev,e=>{e.preventDefault();ov.style.display="flex";}));
  document.addEventListener("dragleave",e=>{if(!e.relatedTarget)ov.style.display="none";});
  document.addEventListener("drop",e=>{
    e.preventDefault();ov.style.display="none";
    if(e.dataTransfer&&e.dataTransfer.files&&e.dataTransfer.files[0])importFile(e.dataTransfer.files[0]);
  });
  /* plan modal wiring */
  $("btnPlan").onclick=()=>{
    $("plTotal").value=CONFIG.plan_total||"";
    $("plAL").value=CONFIG.plan_al||"";
    $("plCU").value=CONFIG.plan_cu||"";
    $("planModal").style.display="flex";
  };
  $("plClose").onclick=()=>$("planModal").style.display="none";
  $("planModal").onclick=e=>{if(e.target===$("planModal"))$("planModal").style.display="none";};
  $("plSave").onclick=()=>{
    CONFIG.plan_total=Number($("plTotal").value)||0;
    CONFIG.plan_al=Number($("plAL").value)||0;
    CONFIG.plan_cu=Number($("plCU").value)||0;
    syncPlanVars();savePlanLS();applyPlanUI();
    if(ACTIVE&&MONTHS[ACTIVE]){MONTHS[ACTIVE].plan={t:CONFIG.plan_total,a:CONFIG.plan_al,c:CONFIG.plan_cu};persistMonths();}
    $("planModal").style.display="none";
    renderAll();
    showBanner(HAS_PLAN?`✅ الخطة اتحفظت: <b>${n0(PLAN)}</b> طن (AL ${n0(PLAN_AL)} · CU ${n0(PLAN_CU)})${ACTIVE&&MONTHS[ACTIVE]?" لشهر "+esc(monthLabel(ACTIVE)):""}`:"✅ الخطة اتمسحت — الداشبورد بيعرض الإنتاج فقط","ok");
  };
  $("plFile").onclick=()=>{
    CONFIG.plan_total=Number(ORIG_CONFIG.plan_total)||0;
    CONFIG.plan_al=Number(ORIG_CONFIG.plan_al)||0;
    CONFIG.plan_cu=Number(ORIG_CONFIG.plan_cu)||0;
    syncPlanVars();clearPlanLS();applyPlanUI();
    $("planModal").style.display="none";
    renderAll();
    showBanner("✅ رجعت لخطة الملف المدمجة","ok");
  };
  applyPlanUI();
  renderAll();
  if(ACTIVE&&ACTIVE!=="__embed__"&&MONTHS[ACTIVE]){
    showBanner(`🗄️ تم تحميل آخر شهر كنت شغال عليه: <b>${esc(monthLabel(ACTIVE))}</b> (${n0(DATA.length)} طرد) — محفوظ في المتصفح على الجهاز ده`,"info");
  }
}
if(typeof document!=="undefined"){document.addEventListener("DOMContentLoaded",init);}
globalThis.__APP__={DATA,ORIG_DATA,CONFIG,ORIG_CONFIG,PLAN,PLAN_AL,PLAN_CU,MDAYS,HAS_PLAN,state,computeKPIs,filteredRows,groupTons,groupCount,topFrom,parseErpRows,syncPlanVars,cfgMonthKey,cfgMonthNum,monthLabel,storeMonth,loadMonthEntry,getMONTHS:()=>MONTHS,getACTIVE:()=>ACTIVE};
