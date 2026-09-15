
const state={page:"dashboard",start:null,end:null,weekEnd:null};

const labels={
 dashboard:["Management View","H-2 Achievement Dashboard"],
 banani:["Branch","Banani Sales Summary"],
 chattogram:["Branch","Chattogram Sales Summary"],
 hq:["Control Tower","HQ Sales Summary"],
 motijheel:["Branch","Motijheel Sales Summary"],
 "daily-gmv":["Company-wide","Daily GMV"],
 wbr:["Weekly Review","Weekly Business Review"]
};

function esc(v){return String(v??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#039;"}[c]))}
function usd(v){if(v===null||v===undefined||v==="")return"N/A";const n=Number(v);return Number.isFinite(n)?"$"+n.toLocaleString("en-US",{maximumFractionDigits:2}):"N/A"}
function n(v){const x=Number(v);return Number.isFinite(x)?x:0}
function pct(v){if(v===null||v===undefined||!Number.isFinite(Number(v)))return"N/A";return(Number(v)*100).toFixed(1)+"%"}
function today(){return new Intl.DateTimeFormat("en-CA",{timeZone:"Asia/Dhaka"}).format(new Date())}

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
 if(state.page==="daily-gmv"){
   // No export endpoint for this page yet — build_export_payload() in
   // app.py only knows dashboard/wbr/branch pages. Leaving this empty
   // rather than wiring a button to a 400.
   document.getElementById("exports").innerHTML="";
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

/* --- Data loading, with a short client-side cache ---------------------
   The backend now caches its own heavy computation per refresh cycle
   (30 min), but every page NAVIGATION still round-trips the network
   even when you're just flipping back to a page you looked at 10
   seconds ago. This adds a short (60s) in-memory cache so revisiting a
   page within that window renders instantly from what's already in
   hand, while a background fetch quietly confirms/updates it — data
   can't meaningfully change in 60s given the 30-minute refresh cycle,
   so this is purely about perceived speed, not staleness risk.
   Session-only by design (a plain JS Map, not localStorage): there's no
   reason a cached figure should outlive the tab, and this keeps it dead
   simple. manualRefresh() clears it explicitly — see below — so the
   REFRESH DATA button always hits the network, never a stale cache. */
const CACHE_TTL_MS=60*1000;
const dataCache=new Map();

function cacheKeyFor(page){
 if(page==="dashboard")return "dashboard";
 if(page==="wbr")return "wbr|"+(state.weekEnd||today());
 return page+"|"+(state.start||today())+".."+(state.end||today());
}

function urlFor(page){
 if(page==="dashboard")return "/api/dashboard";
 if(page==="wbr")return `/api/wbr?week_end=${encodeURIComponent(state.weekEnd||today())}`;
 if(page==="daily-gmv")return `/api/daily-gmv?start=${encodeURIComponent(state.start||today())}&end=${encodeURIComponent(state.end||today())}`;
 return `/api/branch/${page}?start=${encodeURIComponent(state.start||today())}&end=${encodeURIComponent(state.end||today())}`;
}

function setStatusFrom(d){
 state.lastRefresh=d.last_refresh||null;
 setStatus((d.source||"Source")+" • last data refresh "+(state.lastRefresh?new Date(state.lastRefresh).toLocaleString("en-GB",{timeZone:"Asia/Dhaka",day:"2-digit",month:"short",year:"numeric",hour:"2-digit",minute:"2-digit",hour12:false})+" BDT":"unknown"));
}

async function load(){
 const root=document.getElementById("view");
 const page=state.page;
 const key=cacheKeyFor(page);
 const url=urlFor(page);
 const cached=dataCache.get(key);
 const isFresh=cached && (Date.now()-cached.ts<CACHE_TTL_MS);

 if(cached){
   render(cached.data);
   setStatusFrom(cached.data);
 }else{
   root.innerHTML=`<div class="card" style="padding:60px;text-align:center;color:#718096">Loading data…</div>`;
 }

 if(isFresh)return;

 try{
   const d=await api(url);
   dataCache.set(key,{data:d,ts:Date.now()});
   if(state.page!==page)return; // navigated elsewhere while this was in flight
   render(d);
   setStatusFrom(d);
 }catch(err){
   if(state.page!==page)return;
   if(cached){
     // Already showing the last good data for this page — a failed
     // background refresh shouldn't blank that out from under the user.
     setStatus("Refresh failed • showing cached data","warn");
     return;
   }
   root.innerHTML=`<div class="error"><b>Live data unavailable.</b><br>${esc(err.message||err)}<br><br>The dashboard is not showing a snapshot because live mode is enabled.</div>`;
   setStatus("LIVE SOURCE UNAVAILABLE","error");
 }
}

async function manualRefresh(){
 const buttons=document.querySelectorAll(".refresh-main");
 buttons.forEach(b=>{b.disabled=true;b.textContent="↻ REFRESHING…"});
 setStatus("Refreshing source data…","warn");
 try{
   const d=await api("/api/refresh",{method:"POST"});
   state.lastRefresh=d.refreshed_at||state.lastRefresh;
   dataCache.clear(); // force load() past the 60s cache onto the network
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
 if(controls.querySelector(".refresh-main")) return;
 const wrap=document.createElement("div");
 wrap.className="refresh-wrap";
 wrap.innerHTML=`<button class="button primary refresh refresh-main" onclick="manualRefresh()">↻ REFRESH DATA</button><div class="refresh-caption">Automatic refresh every 30 minutes</div>`;
 controls.appendChild(wrap);
}

/* Renderers */
function render(d){
 if(d.page==="dashboard"){renderDashboard(d);return}
 if(d.page==="wbr"){renderWbr(d);return}
 if(d.page==="daily-gmv"){renderDailyGmv(d);addRefreshButton();return}
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
 <div class="card"><div class="section-title">Target Achievement <span class="small">• ${esc(d.locked_month_progress?.month||"")}</span></div>
 ${Object.entries(d.locked_month_progress?.products||{}).filter(([name])=>["Flight","Hotel","Tour","Visa"].includes(name)).map(([name,x])=>{const t=(d.locked_month_progress.targets||{})[name]||0;const a=t?x.gmv/t:null;return `<div class="mini"><span>${name}</span><b>${pct(a)}</b></div><div class="bar"><i style="width:${Math.min(100,(a||0)*100)}%"></i></div><div class="small" style="margin:4px 0 10px">${usd(x.gmv)} / ${usd(t)}</div>`}).join("")}
 </div></div>
 <div class="grid3 section">
 <div class="card"><div class="section-title">Run Rate</div><div class="mini"><span>Required Run Rate</span><b>${usd(d.progress.required_run_rate)}</b></div><div class="mini"><span>Gap to Required</span><b class="${d.progress.gap>=0?"good":"bad"}">${usd(d.progress.gap)}</b></div></div>
 <div class="card"><div class="section-title">Operations</div><div class="mini"><span>Active CSS</span><b>${n(d.progress.active_css)}</b></div><div class="mini"><span>Pipeline Number</span><b>${n(d.progress.pipeline_number)}</b></div><div class="mini"><span>Pipeline Worth</span><b>${usd(d.progress.pipeline_worth)}</b></div></div>
 <div class="card"><div class="section-title">Footfall</div><div class="big">${n(d.progress.footfall).toLocaleString("en-US")}</div><div class="small">Selected period • avg ${n(d.progress.footfall_avg_daily).toFixed(1)}/day</div></div>
 </div>`;
}

function renderDailyGmv(d){
 const rows=d.rows||[];
 const maxTotal=Math.max(1,...rows.map(r=>n(r.total)));
 const totals={flight:0,hotel:0,tour:0,visa:0};
 rows.forEach(r=>{totals.flight+=n(r.flight);totals.hotel+=n(r.hotel);totals.tour+=n(r.tour);totals.visa+=n(r.visa)});
 let peak=null;
 for(const r of rows){if(!peak||n(r.total)>n(peak.total))peak=r}

 document.getElementById("view").innerHTML=`
 <div class="cards">
 ${[
  ["Total GMV",usd(d.total_gmv),`${d.start} → ${d.end} • ${d.days} day${d.days===1?"":"s"}`,true],
  ["Average Daily GMV",usd(d.average_daily_gmv),"Across selected range",false],
  ["Best Day",peak?usd(peak.total):"N/A",peak?peak.date:"No data in range",false],
  ["Days Selected",String(d.days),"In current range",false]
 ].map(x=>`<div class="card ${x[3]?"blue":""}"><div class="label">${x[0]}</div><div class="big">${x[1]}</div><div class="small" style="margin-top:6px">${x[2]}</div><div class="currency">USD</div></div>`).join("")}
 </div>
 <div class="grid2 section">
 <div class="card">
   <div class="section-title">Daily GMV</div>
   <div class="small" style="margin-bottom:8px">${esc(d.start)} → ${esc(d.end)} • bar shows each day relative to the range's highest day</div>
   ${rows.length?rows.map(r=>`<div class="daily-row"><span class="daily-date">${esc(r.date)}</span><div class="bar"><i style="width:${Math.max(2,n(r.total)/maxTotal*100)}%"></i></div><span class="daily-value">${usd(r.total)}</span></div>`).join(""):`<div class="small">No data in the selected range.</div>`}
 </div>
 <div class="card">
   <div class="section-title">Product Split</div>
   <div class="small" style="margin-bottom:8px">Sum across selected range • USD</div>
   <table class="table"><thead><tr><th>Product</th><th>GMV</th></tr></thead><tbody>
   ${[["Flight",totals.flight],["Hotel",totals.hotel],["Tour",totals.tour],["Visa",totals.visa]].map(([name,v])=>`<tr><td>${name}</td><td>${usd(v)}</td></tr>`).join("")}
   </tbody></table>
 </div>
 </div>`;
}

function renderWbr(d){
 const g=d.summary.gmv,b=d.summary.bookings,r=d.summary.receivable,nr=d.summary.net_revenue;
 document.getElementById("view").innerHTML=`
 <div class="notice"><b>Week:</b> ${esc(d.start)} → ${esc(d.end)} <span style="margin-left:12px">• Currency: USD</span>${d.manual_override?` <span style="margin-left:12px">• Reviewing a past week — <a href="#" onclick="document.getElementById('weekEnd').value='';document.getElementById('weekEnd').dispatchEvent(new Event('change'));return false;" style="color:inherit">back to current</a></span>`:""}</div>
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
 <div class="card"><div class="section-title">Operations</div><div class="mini"><span>Weekly Footfall</span><b>${n(d.footfall).toLocaleString("en-US")}</b></div><div class="mini"><span>Avg Daily Footfall</span><b>${n(d.footfall_avg_daily??(n(d.footfall)/7)).toFixed(1)}</b></div></div></div>
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

/* --- Sidebar toggle ---------------------------------------------------
   The button's own `left` position transitions on click (see
   styles.css). If the cursor doesn't move afterward, the browser can
   keep rendering :hover from the button's old position — the .no-hover
   class forces the resting style for the transition's duration so that
   can't happen, then gets out of the way so real hover behaves normally
   again. blur() additionally drops the click's own focus ring. */
const sidebarToggle=document.getElementById("sidebarToggle");
sidebarToggle.addEventListener("click",()=>{
 document.body.classList.toggle("sidebar-hidden");
 const hidden=document.body.classList.contains("sidebar-hidden");
 sidebarToggle.textContent=hidden?"☰":"×";
 sidebarToggle.setAttribute("aria-label",hidden?"Show navigation":"Hide navigation");
 sidebarToggle.title=hidden?"Show navigation":"Hide navigation";
 sidebarToggle.blur();
 sidebarToggle.classList.add("no-hover");
 clearTimeout(sidebarToggle._hoverResetTimer);
 sidebarToggle._hoverResetTimer=setTimeout(()=>sidebarToggle.classList.remove("no-hover"),220);
});

/* --- Theme (light / dark) ----------------------------------------------
   Defaults to the OS preference (styles.css handles that with no JS at
   all, via prefers-color-scheme) until the person picks explicitly, at
   which point that choice is remembered and wins regardless of OS
   setting. Stored in localStorage since this is a real per-browser
   preference on a real deployed page, not a Claude-artifact sandbox. */
const THEME_KEY="gz-theme";
const themeToggle=document.getElementById("themeToggle");

function systemPrefersDark(){
 return window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches;
}
function isDarkActive(){
 const explicit=document.documentElement.getAttribute("data-theme");
 return explicit?explicit==="dark":systemPrefersDark();
}
function applyTheme(theme){
 if(theme==="light"||theme==="dark")document.documentElement.setAttribute("data-theme",theme);
 else document.documentElement.removeAttribute("data-theme");
 themeToggle.textContent=isDarkActive()?"☀️":"🌙";
}
function initTheme(){
 let saved=null;
 try{saved=localStorage.getItem(THEME_KEY)}catch(e){/* storage blocked — fall back to OS preference */}
 applyTheme(saved);
}
themeToggle.addEventListener("click",()=>{
 const next=isDarkActive()?"light":"dark";
 try{localStorage.setItem(THEME_KEY,next)}catch(e){/* storage blocked — theme still applies for this load */}
 applyTheme(next);
});

document.querySelectorAll("#nav button").forEach(b=>b.addEventListener("click",()=>setPage(b.dataset.page)));

document.addEventListener("change",e=>{
 if(e.target.id==="start"){state.start=e.target.value;writeUrl();load();setExports();addRefreshButton()}
 if(e.target.id==="end"){state.end=e.target.value;writeUrl();load();setExports();addRefreshButton()}
 if(e.target.id==="weekEnd"){state.weekEnd=e.target.value||today();writeUrl();load();setExports();addRefreshButton()}
});

window.addEventListener("popstate",()=>{readUrl();setPage(state.page,false)});

// Page visibility resume: a 30-minute browser refresh, while backend auto refreshes every 30 minutes.
setInterval(()=>load(),30*60*1000);

initTheme();
readUrl();
setPage(state.page,false);
