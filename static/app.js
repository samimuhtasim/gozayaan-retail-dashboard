
const state={page:"dashboard",start:null,end:null,weekEnd:null};

const labels={
 dashboard:["Management View","H-2 Achievement Dashboard"],
 banani:["Branch","Banani Sales Summary"],
 chattogram:["Branch","Chattogram Sales Summary"],
 hq:["Control Tower","HQ Sales Summary"],
 motijheel:["Branch","Motijheel Sales Summary"],
 wbr:["Weekly Review","Weekly Business Review"]
};

function esc(v){return String(v??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#039;"}[c]))}
function usd(v){if(v===null||v===undefined||v==="")return"N/A";const n=Number(v);return Number.isFinite(n)?"$"+n.toLocaleString("en-US",{maximumFractionDigits:2}):"N/A"}
function n(v){const x=Number(v);return Number.isFinite(x)?x:0}
function pct(v){if(v===null||v===undefined||!Number.isFinite(Number(v)))return"N/A";return(Number(v)*100).toFixed(1)+"%"}
function today(){return new Date().toISOString().slice(0,10)}

function readUrl(){
 const q=new URLSearchParams(location.search);
 state.page=q.get("page")||"dashboard";
 state.start=q.get("start");
 state.end=q.get("end");
 state.weekEnd=q.get("week_end")||q.get("end");
 if(state.page!=="dashboard" && state.page!=="wbr"){
   state.start=state.start||today();
   state.end=state.end||today();
 }
 if(state.page==="wbr") state.weekEnd=state.weekEnd||today();
}

function writeUrl(push=true){
 const q=new URLSearchParams();
 q.set("page",state.page);
 if(state.page!=="dashboard"){
   if(state.page==="wbr")q.set("week_end",state.weekEnd||today());
   else{q.set("start",state.start||today());q.set("end",state.end||today());}
 }
 const url=location.pathname+"?"+q.toString();
 if(push)history.pushState({},'',url);else history.replaceState({},'',url);
}

function setStatus(text,mode="ok"){
 document.getElementById("statusText").textContent=text;
 const dot=document.querySelector(".status-dot");
 dot.className="status-dot"+(mode==="warn"?" warn":mode==="error"?" err":"");
}

function setControls(){
 const c=document.getElementById("controls");
 if(state.page==="dashboard"){
   c.innerHTML=`<div class="locked">Date locked • live Dashboard</div>`;
   return;
 }
 if(state.page==="wbr"){
   c.innerHTML=`<div class="control"><label>Week ending</label><input id="weekEnd" type="date" value="${state.weekEnd||today()}"></div>`;
   return;
 }
 c.innerHTML=`
   <div class="control"><label>Start date</label><input id="start" type="date" value="${state.start||today()}"></div>
   <div class="control"><label>End date</label><input id="end" type="date" value="${state.end||today()}"></div>`;
}

function setExports(){
 if(state.page==="dashboard"){
   document.getElementById("exports").innerHTML=`
     <button class="export-btn" onclick="exportFile('xlsx')">↓ XLSX</button>
     <button class="export-btn" onclick="exportFile('pdf')">↓ PDF</button>`;
   return;
 }
 document.getElementById("exports").innerHTML=`
   <button class="export-btn" onclick="exportFile('xlsx')">↓ XLSX</button>
   <button class="export-btn" onclick="exportFile('pdf')">↓ PDF</button>`;
}

function setPage(p,push=true){
 state.page=p;
 document.querySelectorAll("#nav button").forEach(b=>b.classList.toggle("active",b.dataset.page===p));
 document.getElementById("eyebrow").textContent=labels[p][0];
 document.getElementById("title").textContent=labels[p][1];
 setControls();setExports();
 writeUrl(push);
 load();
}

async function api(url,options={}){
 const r=await fetch(url,options);
 const data=await r.json();
 if(!r.ok)throw new Error(data.detail||"Request failed");
 return data;
}

function currentExportParams(){
 const q=new URLSearchParams();
 q.set("page",state.page);
 if(state.page==="wbr")q.set("end",state.weekEnd||today());
 else if(state.page!=="dashboard"){q.set("start",state.start||today());q.set("end",state.end||today());}
 return q;
}

function exportFile(fmt){
 const q=currentExportParams();
 location.href=`/api/export/${fmt}?${q.toString()}`;
}

async function load(){
 const root=document.getElementById("view");
 root.innerHTML=`<div class="card" style="padding:60px;text-align:center;color:#718096">Loading data…</div>`;
 try{
   let d;
   if(state.page==="dashboard")d=await api("/api/dashboard");
   else if(state.page==="wbr")d=await api(`/api/wbr?week_end=${encodeURIComponent(state.weekEnd||today())}`);
   else d=await api(`/api/branch/${state.page}?start=${encodeURIComponent(state.start||today())}&end=${encodeURIComponent(state.end||today())}`);
   state.lastRefresh=d.last_refresh||null;
   render(d);
   setStatus((d.source||"Source")+" • last data refresh "+(state.lastRefresh?new Date(state.lastRefresh).toLocaleString("en-GB",{day:"2-digit",month:"short",hour:"2-digit",minute:"2-digit"}):"unknown"));
 }catch(err){
   root.innerHTML=`<div class="error">${esc(err.message||err)}</div>`;
   setStatus("Data request failed","error");
 }
}

async function manualRefresh(){
 const buttons=document.querySelectorAll(".refresh-main");
 buttons.forEach(b=>{b.disabled=true;b.textContent="↻ REFRESHING…"});
 setStatus("Refreshing source data…","warn");
 try{
   const d=await api("/api/refresh",{method:"POST"});
   state.lastRefresh=d.refreshed_at||state.lastRefresh;
   await load();
 }catch(err){
   setStatus("Refresh failed • showing last available data","error");
 }finally{
   buttons.forEach(b=>{b.disabled=false;b.textContent="↻ REFRESH DATA"});
 }
}

/* Add refresh button at the top right of every page. */
function addRefreshButton(){
 const controls=document.getElementById("controls");
 const wrap=document.createElement("div");
 wrap.className="refresh-wrap";
 wrap.innerHTML=`<button class="button primary refresh refresh-main" onclick="manualRefresh()">↻ REFRESH DATA</button><div class="refresh-caption">Automatic refresh every 30 minutes</div>`;
 controls.appendChild(wrap);
}

/* Renderers */
function render(d){
 if(d.page==="dashboard"){renderDashboard(d);return}
 if(d.page==="wbr"){renderWbr(d);return}
 renderBranch(d);
 addRefreshButton();
}

function renderDashboard(d){
 const rows=d.rows.slice(1).filter(r=>r.some(x=>String(x).trim()!==""));
 document.getElementById("view").innerHTML=`
 <div class="cards">
 ${rows.map((r,i)=>`<div class="card ${i===rows.length-1?"blue":""}">
   <div class="label">${esc(r[0])}</div>
   <div class="big">${usd(r[8])}</div>
   <div class="small" style="margin-top:6px">H2 achievement ${pct(Number(r[9]))}</div>
   <div class="currency">USD</div>
   ${i===rows.length-1?'<div class="rule"></div>':""}
 </div>`).join("")}
 </div>
 <div class="section"><div class="card">
   <div class="section-title">H2 Performance</div>
   <div class="small" style="margin-bottom:8px">Product is a row, not an input • USD</div>
   <table class="table"><thead><tr>
   <th>Product</th><th>Q3 Target</th><th>Q3 Actual</th><th>Q3 Ach.</th>
   <th>Q4 Target</th><th>Q4 Actual</th><th>Q4 Ach.</th><th>H2 Target</th><th>H2 Actual</th><th>H2 Ach.</th>
   </tr></thead><tbody>
   ${rows.map(r=>`<tr><td>${esc(r[0])}</td><td>${usd(r[1])}</td><td>${usd(r[2])}</td><td>${pct(Number(r[3]))}</td><td>${usd(r[4])}</td><td>${usd(r[5])}</td><td>${pct(Number(r[6]))}</td><td>${usd(r[7])}</td><td>${usd(r[8])}</td><td>${pct(Number(r[9]))}</td></tr>`).join("")}
   </tbody></table>
 </div></div>`;
 addRefreshButton();
}

function renderBranch(d){
 const products=Object.entries(d.products).filter(([name])=>!["Vouchers","Others"].includes(name));
 document.getElementById("view").innerHTML=`
 <div class="cards">
 ${[
  ["Selected Period GMV",usd(d.total_gmv),`${d.start} → ${d.end}`],
  ["MTD GMV",usd(d.progress.mtd_gmv),"Through selected end date"],
  ["Projected GMV",usd(d.progress.projected),"Month-end projection"],
  ["Current Run Rate",usd(d.progress.current_run_rate),"Per day"]
 ].map(x=>`<div class="card"><div class="label">${x[0]}</div><div class="big">${x[1]}</div><div class="small" style="margin-top:5px">${x[2]}</div><div class="currency">USD</div></div>`).join("")}
 </div>
 <div class="grid2 section">
 <div class="card"><div class="section-title">Product Performance</div><div class="small" style="margin-bottom:8px">${esc(d.start)} → ${esc(d.end)} • USD</div>
 <table class="table"><thead><tr><th>Product</th><th>CSS GMV</th><th>Retail GMV</th><th>Total GMV</th><th>Receivable</th><th>Bookings</th></tr></thead><tbody>
 ${products.map(([name,x])=>`<tr><td>${name}</td><td>${usd(x.css)}</td><td>${usd(x.retail)}</td><td><b>${usd(x.gmv)}</b></td><td>${usd(x.receivable)}</td><td>${n(x.bookings)}</td></tr>`).join("")}
 </tbody></table></div>
 <div class="card"><div class="section-title">Target Achievement</div>
 ${products.map(([name,x])=>{const t=d.progress.targets[name]||0;const a=t?x.gmv/t:null;return `<div class="mini"><span>${name}</span><b>${pct(a)}</b></div><div class="bar"><i style="width:${Math.min(100,(a||0)*100)}%"></i></div><div class="small" style="margin:4px 0 10px">${usd(x.gmv)} / ${usd(t)}</div>`}).join("")}
 </div></div>
 <div class="grid3 section">
 <div class="card"><div class="section-title">Run Rate</div><div class="mini"><span>Required Run Rate</span><b>${usd(d.progress.required_run_rate)}</b></div><div class="mini"><span>Gap to Required</span><b class="${d.progress.gap>=0?"good":"bad"}">${usd(d.progress.gap)}</b></div></div>
 <div class="card"><div class="section-title">Operations</div><div class="mini"><span>Active CSS</span><b>${n(d.progress.active_css)}</b></div><div class="mini"><span>Pipeline Number</span><b>${n(d.progress.pipeline_number)}</b></div><div class="mini"><span>Pipeline Worth</span><b>${usd(d.progress.pipeline_worth)}</b></div></div>
 <div class="card"><div class="section-title">Footfall</div><div class="big">${n(d.progress.footfall).toLocaleString("en-US")}</div><div class="small">Selected period</div></div>
 </div>`;
}

function renderWbr(d){
 const g=d.summary.gmv,b=d.summary.bookings,r=d.summary.receivable,nr=d.summary.net_revenue;
 document.getElementById("view").innerHTML=`
 <div class="notice"><b>Week:</b> ${esc(d.start)} → ${esc(d.end)} <span style="margin-left:12px">• Currency: USD</span></div>
 <div class="cards">
 ${[
 ["Weekly GMV",usd(g.Total),pct(g.WoW)],
 ["Bookings",n(b.Total).toLocaleString("en-US"),pct(b.WoW)],
 ["Receivable",usd(r.Total),pct(r.Ratio)],
 ["Net Revenue",usd(nr.Total),pct(nr.Ratio)]
 ].map((x,i)=>`<div class="card ${i===0?"blue":""}"><div class="label">${x[0]}</div><div class="big">${x[1]}</div><div class="small" style="margin-top:6px">${i<2?x[2]+" WoW":x[2]+" of GMV"}</div><div class="currency">${i===1?"Transactions":"USD"}</div></div>`).join("")}
 </div>
 <div class="section"><div class="card"><div class="section-title">Weekly Performance Summary</div>
 <table class="table"><thead><tr><th>Metric</th><th>Flight</th><th>Tour</th><th>Visa</th><th>Hotel</th><th>Total</th><th>Change / Ratio</th></tr></thead><tbody>
 <tr><td>GMV</td><td>${usd(g.Flight)}</td><td>${usd(g.Tour)}</td><td>${usd(g.Visa)}</td><td>${usd(g.Hotel)}</td><td><b>${usd(g.Total)}</b></td><td>${pct(g.WoW)}</td></tr>
 <tr><td>Bookings</td><td>${n(b.Flight)}</td><td>${n(b.Tour)}</td><td>${n(b.Visa)}</td><td>${n(b.Hotel)}</td><td><b>${n(b.Total)}</b></td><td>${pct(b.WoW)}</td></tr>
 <tr><td>Receivable</td><td>${usd(r.Flight)}</td><td>${usd(r.Tour)}</td><td>${usd(r.Visa)}</td><td>${usd(r.Hotel)}</td><td><b>${usd(r.Total)}</b></td><td>${pct(r.Ratio)}</td></tr>
 <tr><td>Net Revenue</td><td>${usd(nr.Flight)}</td><td>${usd(nr.Tour)}</td><td>${usd(nr.Visa)}</td><td>${usd(nr.Hotel)}</td><td><b>${usd(nr.Total)}</b></td><td>${pct(nr.Ratio)}</td></tr>
 </tbody></table></div></div>
 <div class="grid2 section"><div class="card"><div class="section-title">Weekly Target vs Achievement</div>
 <table class="table"><thead><tr><th>Product</th><th>Weekly Target</th><th>Actual</th><th>Achievement</th><th>Status</th></tr></thead><tbody>
 ${Object.entries(d.weekly_targets).map(([name,target])=>{const actual=g[name]||0;const a=target?actual/target:null;const status=a==null?"N/A":a>=1?"On Track":a>=.8?"At Risk":"Behind";const cls=status==="On Track"?"good":status==="At Risk"?"warn":"bad";return `<tr><td>${name}</td><td>${usd(target)}</td><td>${usd(actual)}</td><td>${pct(a)}</td><td><span class="badge ${cls}">${status}</span></td></tr>`}).join("")}
 </tbody></table></div>
 <div class="card"><div class="section-title">Operations</div><div class="mini"><span>Weekly Footfall</span><b>${n(d.footfall).toLocaleString("en-US")}</b></div><div class="mini"><span>Avg Daily Footfall</span><b>${(n(d.footfall)/7).toFixed(1)}</b></div></div></div>
 ${narrative("Detailed Analysis",d.narrative.analysis)}
 ${narrative("Challenges",d.narrative.challenges)}
 ${narrative("Way Forward / Action Plan",d.narrative.actions)}
 `;
 addRefreshButton();
}

function narrative(title,rows){
 const text=rows.flat().filter(x=>String(x||"").trim()).join("\n");
 return `<div class="section"><div class="card"><div class="section-title">${title}</div><div class="narrative">${esc(text)}</div></div></div>`;
}

document.querySelectorAll("#nav button").forEach(b=>b.addEventListener("click",()=>setPage(b.dataset.page)));

document.addEventListener("change",e=>{
 if(e.target.id==="start"){state.start=e.target.value;writeUrl();load();setExports();addRefreshButton()}
 if(e.target.id==="end"){state.end=e.target.value;writeUrl();load();setExports();addRefreshButton()}
 if(e.target.id==="weekEnd"){state.weekEnd=e.target.value;writeUrl();load();setExports();addRefreshButton()}
});

window.addEventListener("popstate",()=>{readUrl();setPage(state.page,false)});

// Page visibility resume: a 30-minute browser refresh, while backend auto refreshes every 30 minutes.
setInterval(()=>load(),30*60*1000);

readUrl();
setPage(state.page,false);
