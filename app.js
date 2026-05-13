// ═══════════════════════════════════════════════════════════════════════════════
// MAIN INIT
// ═══════════════════════════════════════════════════════════════════════════════
function initDashboard() {
  const today = new Date().toISOString().split('T')[0];
  document.getElementById('inv-date').value   = today;
  document.getElementById('rma-date').value   = today;

  renderChecklist();
  initLogisticsState();
  renderLogistics();
  renderRMA();

  // Listen to Firebase collections
  listenShowroom();
  listenLogistics();
  listenInventory();
  listenRMA();
}

// Wait for Firebase module to expose globals
function waitForFirebase(cb) {
  if (window._db) cb();
  else window.addEventListener('firebase-ready', cb);
}

// ═══════════════════════════════════════════════════════════════════════════════
// PO DATA (local only — from D-Tools export)
// ═══════════════════════════════════════════════════════════════════════════════
let allData = [], currentFilter = 'All', showReceived = false, sortCol = 'Date', sortAsc = false;


function handleFileInput(input) { if (input.files[0]) processFile(input.files[0]); }

function processFile(file) {
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb = XLSX.read(new Uint8Array(e.target.result), { type:'array' });
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval:'' });
      const sixMonthsAgo = new Date(); sixMonthsAgo.setMonth(sixMonthsAgo.getMonth()-6);
      allData = rows.map(row => ({
        Number:    String(row['Number']||row['PO #']||row['PO#']||'').trim(),
        Date:      formatDate(row['Date']),
        Vendor:    String(row['Vendor']||'').trim(),
        Status:    String(row['Status']||'').trim(),
        Projects:  cleanProject(String(row['Projects']||'')) || cleanProject(String(row['Service Orders']||'')),
        VendorRef: String(row['Vendor Reference Number']||row['Vendor Ref']||'').trim(),
        Tracking:  String(row['Tracking Number']||row['Tracking']||'').trim(),
        Notes:     String(row['Notes']||'').trim(),
      })).filter(r => {
        if (r.Status !== 'Received') return true;
        const d = new Date(r.Date); return !isNaN(d) && d >= sixMonthsAgo;
      });
      const ts = 'POs updated ' + new Date().toLocaleString('en-US',{month:'short',day:'numeric',hour:'numeric',minute:'2-digit'});
      document.getElementById('last-updated').textContent = ts;
      updateKPIs(); renderTable();
      window._saveDoc('dashboard/podata', { rows: allData, timestamp: ts, filename: file.name });
    } catch(err) { alert("Could not read file."); console.error(err); }
  };
  reader.readAsArrayBuffer(file);
}

function formatDate(val) {
  if (!val) return '';
  if (typeof val === 'number') { const d = new Date(Math.round((val-25569)*86400*1000)); return d.toLocaleDateString('en-US',{month:'numeric',day:'numeric',year:'numeric'}); }
  const str = String(val).trim();
  const match = str.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (match) { const d = new Date(parseInt(match[1]),parseInt(match[2])-1,parseInt(match[3])); return d.toLocaleDateString('en-US',{month:'numeric',day:'numeric',year:'numeric'}); }
  return str;
}
function cleanProject(raw) {
  if (!raw || !String(raw).trim()) return '';
  let p = String(raw).trim();
  let client = '';
  if (p.includes(' : ')) { const parts = p.split(' : '); client = parts[0].trim(); p = parts[1].trim(); }
  p = p.replace(/\(PORT\s*-?\s*(\d+)\)/i, '($1)');
  if (client) p = client + ' - ' + p;
  return p.trim();
}
function trackingLink(raw) {
  if (!raw) return '<span style="opacity:0.3">—</span>';
  const str = raw.trim(); let carrier='', number=str;
  if (str.includes(':')) { const p=str.split(':'); carrier=p[0].trim().toLowerCase(); number=p.slice(1).join(':').trim(); }
  else if (/^1Z/i.test(str)) carrier='ups';
  else if (/^\d{15,22}$/.test(str)) carrier='fedex';
  else if (/^9[2345]\d{18,}$/.test(str)) carrier='usps';
  let url='';
  if (carrier.includes('fedex')) url=`https://www.fedex.com/fedextrack/?trknbr=${encodeURIComponent(number)}`;
  else if (carrier.includes('ups')) url=`https://www.ups.com/track?tracknum=${encodeURIComponent(number)}`;
  else if (carrier.includes('usps')) url=`https://tools.usps.com/go/TrackConfirmAction?tLabels=${encodeURIComponent(number)}`;
  if (url) return `<a href="${url}" target="_blank" class="tracking-link">${number} ↗</a>`;
  return `<span class="tracking-plain">${str}</span>`;
}
function updateKPIs() {
  const c={Issued:0,'Partially Received':0,Received:0,Draft:0};
  allData.forEach(r=>{if(c[r.Status]!==undefined)c[r.Status]++;});
  document.getElementById('kpi-total').textContent=allData.length;
  document.getElementById('kpi-issued').textContent=c['Issued'];
  document.getElementById('kpi-partial').textContent=c['Partially Received'];
  document.getElementById('kpi-received').textContent=c['Received'];
  document.getElementById('kpi-draft').textContent=c['Draft'];
}
function setFilter(val,btn) { currentFilter=val; document.querySelectorAll('.filter-btn:not(.toggle-received)').forEach(b=>b.classList.remove('active')); btn.classList.add('active'); renderTable(); }
function toggleReceived() { showReceived=!showReceived; const b=document.getElementById('toggle-received-btn'); b.textContent=showReceived?'Hide Received':'Show Received'; b.classList.toggle('on',showReceived); renderTable(); }
function sortBy(col) { if(sortCol===col)sortAsc=!sortAsc; else{sortCol=col;sortAsc=true;} document.querySelectorAll('thead th').forEach(t=>t.classList.remove('sorted')); const i=['Number','Date','Vendor','Status','Projects'].indexOf(col); if(i>=0)document.querySelectorAll('thead th')[i].classList.add('sorted'); renderTable(); }
function statusBadge(s) { const m={'Issued':'badge-issued','Partially Received':'badge-partial','Received':'badge-received','Draft':'badge-draft','Cancelled':'badge-cancelled'}; return `<span class="badge ${m[s]||'badge-draft'}">${s}</span>`; }
function renderTable() {
  const search=document.getElementById('search-input').value.toLowerCase(), has=search.length>0;
  let data=allData.filter(r=>{
    const isR=r.Status==='Received';
    const mS=!has||(r.Number.toLowerCase().includes(search)||r.Vendor.toLowerCase().includes(search)||r.Projects.toLowerCase().includes(search)||r.Notes.toLowerCase().includes(search)||r.VendorRef.toLowerCase().includes(search));
    if(!mS)return false; if(has)return true;
    if(isR&&!showReceived)return false;
    if(currentFilter!=='All'&&r.Status!==currentFilter)return false;
    return true;
  });
  data.sort((a,b)=>{let av=a[sortCol]||'',bv=b[sortCol]||''; if(sortCol==='Date'){av=new Date(av);bv=new Date(bv);return sortAsc?av-bv:bv-av;} av=String(av).toLowerCase();bv=String(bv).toLowerCase();return sortAsc?av.localeCompare(bv):bv.localeCompare(av);});
  const tbody=document.getElementById('po-tbody');
  if(!data.length){tbody.innerHTML=`<tr><td colspan="8"><div id="empty-state"><div class="empty-icon">${allData.length?'🔍':'📋'}</div><div class="empty-title">${allData.length?'No results found':'No data loaded yet'}</div><div class="empty-sub">${allData.length?'Try adjusting your filter or search':'Drop your D-Tools PO export above to get started'}</div></div></td></tr>`;return;}
  tbody.innerHTML=data.map((r,i)=>`<tr class="fade-in ${r.Status==='Received'?'row-received':''}" style="animation-delay:${Math.min(i*0.02,0.3)}s"><td><span class="po-num">${r.Number}</span></td><td style="white-space:nowrap;color:var(--muted);font-size:12px">${r.Date}</td><td><span class="vendor-name">${r.Vendor}</span></td><td>${statusBadge(r.Status)}</td><td><span class="project-name">${r.Projects||'<span style="opacity:0.3">—</span>'}</span></td><td><span class="po-num">${r.VendorRef||'<span style="opacity:0.3">—</span>'}</span></td><td>${trackingLink(r.Tracking)}</td><td><span class="notes-text">${r.Notes||''}</span></td></tr>`).join('');
}

// ═══════════════════════════════════════════════════════════════════════════════
// SHOWROOM CHECKLIST — Firebase
// ═══════════════════════════════════════════════════════════════════════════════
const SHOWROOM_ITEMS = [
  {id:'foyer',    area:'Foyer',       detail:'TV / Display / Sound Bar'},
  {id:'showroom', area:'Showroom',    detail:'Can Lights / Floor Lights / Wall Speakers / Touch Panel'},
  {id:'fireplace',area:'Fireplace',   detail:'TV / Sound / Control Rack / Shelf Lights'},
  {id:'island',   area:'Island',      detail:'Lights / Pendant Lights / Control Device'},
  {id:'couch',    area:'Couch Area',  detail:'Big TV / Thermostat / Ring Camera'},
  {id:'kitchen',  area:'Kitchenette', detail:'Wall Lights / TV / Fridge / Lights'},
  {id:'theater',  area:'Theater Room',detail:'Rack / Control Device / Projector / Sound / Lights / Thermostat'},
];
let showroomItems = SHOWROOM_ITEMS.map(i=>({...i,status:'unchecked'}));

function getWeekLabel() {
  const now=new Date(), mon=new Date(now); mon.setDate(now.getDate()-((now.getDay()+6)%7));
  return 'Week of '+mon.toLocaleDateString('en-US',{month:'short',day:'numeric',year:'numeric'});
}

function listenPO() {
  waitForFirebase(()=>{
    window._onSnapshot(window._doc(window._db,'dashboard','podata'), snap=>{
      if(snap.exists()&&snap.data().rows) {
        allData = snap.data().rows;
        const ts = snap.data().timestamp || '';
        if(ts) document.getElementById('last-updated').textContent = ts;
        updateKPIs();
        renderTable();
      }
    });
  });
}

function listenShowroom() {
  waitForFirebase(()=>{
    window._onSnapshot(window._doc(window._db,'dashboard','showroom'), snap=>{
      if(snap.exists()) { const d=snap.data(); if(d.items)showroomItems=d.items; if(d.notes)document.getElementById('showroom-notes').value=d.notes; renderChecklist(); }
    });
  });
}

async function toggleChecklistItem(id) {
  const item=showroomItems.find(i=>i.id===id); if(!item)return;
  item.status=item.status==='unchecked'?'pass':item.status==='pass'?'fail':'unchecked';
  renderChecklist();
  await window._saveDoc('dashboard/showroom',{items:showroomItems});
}

async function resetChecklist() {
  showroomItems=SHOWROOM_ITEMS.map(i=>({...i,status:'unchecked'}));
  document.getElementById('showroom-notes').value='';
  renderChecklist();
  await window._saveDoc('dashboard/showroom',{items:showroomItems,notes:''});
}

let showroomNoteTimer;
function saveShowroomNotes() {
  clearTimeout(showroomNoteTimer);
  showroomNoteTimer = setTimeout(async()=>{
    await window._saveDoc('dashboard/showroom',{notes:document.getElementById('showroom-notes').value});
  },800);
}

function renderChecklist() {
  document.getElementById('showroom-week').textContent='· '+getWeekLabel();
  document.getElementById('showroom-grid').innerHTML=showroomItems.map(item=>{
    const s=item.status,cls=s==='pass'?'pass':s==='fail'?'fail':'unknown',icon=s==='pass'?'✓':s==='fail'?'✕':'—',label=s==='pass'?'Pass':s==='fail'?'Fail':'Tap to check',lc=s==='pass'?'#86efac':s==='fail'?'#fca5a5':'#475569';
    return `<div class="checklist-item ${cls}" onclick="toggleChecklistItem('${item.id}')"><div class="checklist-dot ${cls}">${icon}</div><div style="flex:1"><div class="checklist-name">${item.area}</div><div class="checklist-sub">${item.detail}</div></div><div style="font-size:11px;font-weight:600;color:${lc};white-space:nowrap">${label}</div></div>`;
  }).join('');
}

// ═══════════════════════════════════════════════════════════════════════════════
// LOGISTICS CHECKLIST — Firebase
// ═══════════════════════════════════════════════════════════════════════════════
const LOGISTICS_SECTIONS=[
  {id:'fleet',title:'Fleet',tasks:[
    {id:'fl1',text:'All tech vans inspected for tire pressure, glass cracks, and body damage. Exterior branding is clean and no visible maintenance codes (Oil, Brakes, Check Engine, etc.)',monthly:false},
    {id:'fl2',text:'Service appointments made for any vehicle issues.',monthly:false},
    {id:'fl3',text:'Inspect vinyl condition and overall exterior presentation of vehicles.',monthly:false},
    {id:'fl4',text:'Van products, parts, and tools inventoried and adjusted as needed.',monthly:true},
  ]},
  {id:'inventory',title:'Inventory & Staging',tasks:[
    {id:'in1',text:"All parts for the next day's projects are labeled and staged, ready to be loaded into the vans.",monthly:false},
    {id:'in2',text:'Packages delivered checked against packing slips and received. Items for orders placed into appropriate staging area, items for stock placed onto appropriate shelves.',monthly:false},
    {id:'in3',text:'Process any defective items for RMA return and reorder any necessary replacement parts.',monthly:false},
    {id:'in4',text:'Organize and clean trash and debris from warehouse.',monthly:false},
  ]},
  {id:'facilities',title:'Facilities',tasks:[
    {id:'fa1',text:'Be present at Morning Stand Up.',monthly:false},
    {id:'fa2',text:'Exterior/Interior signs, lighting, HVAC, plumbing all working properly and shop floor clear of trip hazards. Entrances free of trash and debris. Trash checked and emptied.',monthly:false},
    {id:'fa3',text:'All van keys are accounted for at the start/end of shift.',monthly:false},
    {id:'fa4',text:'Racks are cool, cable management is tidy, UPS backups show "Green", OVRC cleared of unnecessary devices and labeled properly.',monthly:false},
    {id:'fa5',text:'Fire extinguishers and first aid kits checked and restocked when necessary.',monthly:true},
  ]},
];
const DAYS=['M','T','W','T','F'];
let logisticsState={};

function initLogisticsState() {
  logisticsState={};
  LOGISTICS_SECTIONS.forEach(sec=>{
    logisticsState[sec.id]={initials:''};
    sec.tasks.forEach(t=>{ logisticsState[sec.id][t.id]=['unchecked','unchecked','unchecked','unchecked','unchecked']; });
  });
}

function listenLogistics() {
  waitForFirebase(()=>{
    window._onSnapshot(window._doc(window._db,'dashboard','logistics'), snap=>{
      if(snap.exists()) {
        const d=snap.data();
        if(d.state) logisticsState=d.state;
        if(d.notes) document.getElementById('logistics-notes').value=d.notes;
        renderLogistics();
      }
    });
  });
}

async function toggleDayCell(secId,taskId,dayIdx) {
  const cur=logisticsState[secId][taskId][dayIdx];
  logisticsState[secId][taskId][dayIdx]=cur==='unchecked'?'pass':cur==='pass'?'fail':'unchecked';
  renderLogistics();
  await window._saveDoc('dashboard/logistics',{state:logisticsState});
}

async function resetLogistics() {
  initLogisticsState();
  document.getElementById('logistics-notes').value='';
  renderLogistics();
  await window._saveDoc('dashboard/logistics',{state:logisticsState,notes:''});
}

let logisticsNoteTimer;
function saveLogisticsNotes() {
  clearTimeout(logisticsNoteTimer);
  logisticsNoteTimer=setTimeout(async()=>{
    await window._saveDoc('dashboard/logistics',{notes:document.getElementById('logistics-notes').value});
  },800);
}

function renderLogistics() {
  document.getElementById('logistics-week').textContent='· '+getWeekLabel();
  document.getElementById('logistics-wrap').innerHTML=LOGISTICS_SECTIONS.map(sec=>{
    const state=logisticsState[sec.id]||{};
    const tasksHtml=sec.tasks.map(task=>{
      const days=state[task.id]||['unchecked','unchecked','unchecked','unchecked','unchecked'];
      let cellsHtml;
      if(task.monthly){
        const s=days[0],cls=s==='pass'?'pass':s==='fail'?'fail':'unknown',label=s==='pass'?'✓ Done':s==='fail'?'✕ Issue':'Mark Done',color=s==='pass'?'#86efac':s==='fail'?'#fca5a5':'#64748b';
        cellsHtml=`<div class="day-cell ${cls}" onclick="toggleDayCell('${sec.id}','${task.id}',0)" style="width:90px;border-radius:8px;padding:0 10px"><span class="day-val" style="font-size:11px;font-weight:600;color:${color}">${label}</span></div>`;
      } else {
        cellsHtml=DAYS.map((d,i)=>{const s=days[i],cls=s==='pass'?'pass':s==='fail'?'fail':'unknown',val=s==='pass'?'✓':s==='fail'?'✕':'·';return `<div class="day-cell ${cls}" onclick="toggleDayCell('${sec.id}','${task.id}',${i})" title="${['Mon','Tue','Wed','Thu','Fri'][i]}"><span class="day-label">${d}</span><span class="day-val">${val}</span></div>`;}).join('');
      }
      return `<div class="logistics-row${task.monthly?' monthly':''}"><div class="logistics-task">${task.text}${task.monthly?'<span class="monthly-tag">Monthly</span>':''}</div><div class="day-cells">${cellsHtml}</div></div>`;
    }).join('');
    return `<div class="logistics-section"><div class="logistics-section-header"><span class="logistics-section-title">${sec.title}</span><div class="logistics-initials">Initials: <input type="text" maxlength="5" value="${state.initials||''}" placeholder="JMS" oninput="updateInitials('${sec.id}',this.value)"></div></div>${tasksHtml}</div>`;
  }).join('');
}

let initialsTimer;
function updateInitials(secId, val) {
  logisticsState[secId].initials=val;
  clearTimeout(initialsTimer);
  initialsTimer=setTimeout(async()=>{ await window._saveDoc('dashboard/logistics',{state:logisticsState}); },800);
}

// ═══════════════════════════════════════════════════════════════════════════════
// INVENTORY CHART — Firebase
// ═══════════════════════════════════════════════════════════════════════════════
let inventoryEntries=[], invChart=null;

function listenInventory() {
  waitForFirebase(()=>{
    window._onSnapshot(window._doc(window._db,'dashboard','inventory'), snap=>{
      if(snap.exists()&&snap.data().entries) { inventoryEntries=snap.data().entries; renderInventoryChart(); }
    });
  });
}

async function addInventoryEntry() {
  const date=document.getElementById('inv-date').value, inv=parseFloat(document.getElementById('inv-inventory').value), stg=parseFloat(document.getElementById('inv-staging').value);
  if(!date||isNaN(inv)||isNaN(stg)){alert('Please fill in all fields.');return;}
  inventoryEntries.push({date,inventory:inv,staging:stg});
  inventoryEntries.sort((a,b)=>a.date.localeCompare(b.date));
  document.getElementById('inv-inventory').value=''; document.getElementById('inv-staging').value='';
  const next=new Date(date); next.setDate(next.getDate()+7); document.getElementById('inv-date').value=next.toISOString().split('T')[0];
  renderInventoryChart();
  await window._saveDoc('dashboard/inventory',{entries:inventoryEntries});
}

async function clearLastEntry() {
  if(!inventoryEntries.length)return;
  inventoryEntries.pop(); renderInventoryChart();
  await window._saveDoc('dashboard/inventory',{entries:inventoryEntries});
}

function fmtMoney(n){return '$'+n.toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2});}

function renderInventoryChart() {
  const area=document.getElementById('chart-area'), totals=document.getElementById('chart-totals');
  if(!inventoryEntries.length){area.innerHTML='<div class="no-chart-data">No entries yet — add your first inventory total above</div>';totals.innerHTML='';if(invChart){invChart.destroy();invChart=null;}return;}
  const latest=inventoryEntries[inventoryEntries.length-1], combined=latest.inventory+latest.staging;
  totals.innerHTML=`<div class="chart-total">Current Inventory <span class="inv">${fmtMoney(latest.inventory)}</span></div><div class="chart-total">Current Staging <span class="stg">${fmtMoney(latest.staging)}</span></div><div class="chart-total">Combined Total <span class="comb">${fmtMoney(combined)}</span></div>`;
  const labels=inventoryEntries.map(e=>{const d=new Date(e.date+'T00:00:00');return d.toLocaleDateString('en-US',{month:'short',day:'numeric'});});
  if(invChart)invChart.destroy();
  area.innerHTML='<canvas id="inv-chart" style="width:100%!important"></canvas>';
  invChart=new Chart(document.getElementById('inv-chart').getContext('2d'),{
    type:'line',
    data:{labels,datasets:[
      {label:'Inventory',data:inventoryEntries.map(e=>e.inventory),borderColor:'#93c5fd',backgroundColor:'rgba(147,197,253,0.08)',pointBackgroundColor:'#93c5fd',tension:0.3,pointRadius:4},
      {label:'Staging',data:inventoryEntries.map(e=>e.staging),borderColor:'#86efac',backgroundColor:'rgba(134,239,172,0.08)',pointBackgroundColor:'#86efac',tension:0.3,pointRadius:4},
      {label:'Combined',data:inventoryEntries.map(e=>e.inventory+e.staging),borderColor:'#fbbf24',backgroundColor:'rgba(251,191,36,0.06)',pointBackgroundColor:'#fbbf24',tension:0.3,pointRadius:4,borderDash:[4,3]},
    ]},
    options:{responsive:true,maintainAspectRatio:true,onResize:(chart,size)=>{ chart.resize(); },plugins:{legend:{labels:{color:'#94a3b8',font:{family:'DM Sans',size:12},boxWidth:12,padding:20}},tooltip:{backgroundColor:'#1c2333',borderColor:'rgba(255,255,255,0.1)',borderWidth:1,titleColor:'#e2e8f0',bodyColor:'#94a3b8',callbacks:{label:c=>` ${c.dataset.label}: ${fmtMoney(c.raw)}`}}},scales:{x:{grid:{color:'rgba(255,255,255,0.04)'},ticks:{color:'#64748b',font:{family:'DM Mono',size:11}}},y:{grid:{color:'rgba(255,255,255,0.04)'},ticks:{color:'#64748b',font:{family:'DM Mono',size:11},callback:v=>'$'+Math.round(v/1000)+'K'}}}}
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// RMA TRACKER — Firebase
// ═══════════════════════════════════════════════════════════════════════════════
let rmaData=[];

function listenRMA() {
  waitForFirebase(()=>{
    window._onSnapshot(window._doc(window._db,'dashboard','rma'), snap=>{
      if(snap.exists()&&snap.data().items){rmaData=snap.data().items;renderRMA();}
    });
  });
}

async function addRMA() {
  const client=document.getElementById('rma-client').value.trim(), vendor=document.getElementById('rma-vendor').value.trim(), item=document.getElementById('rma-item').value.trim(), number=document.getElementById('rma-number').value.trim(), date=document.getElementById('rma-date').value, status=document.getElementById('rma-status').value;
  if(!vendor||!item){alert('Please enter at least a vendor and item.');return;}
  rmaData.push({id:Date.now(),client,vendor,item,number,date,status,notes:''});
  document.getElementById('rma-client').value=''; document.getElementById('rma-vendor').value=''; document.getElementById('rma-item').value=''; document.getElementById('rma-number').value=''; document.getElementById('rma-status').value='Submitted';
  renderRMA();
  await window._saveDoc('dashboard/rma',{items:rmaData});
}

let rmaTimer;
async function updateRMAStatus(id,val) { const r=rmaData.find(r=>r.id===id); if(r){r.status=val;renderRMA();await window._saveDoc('dashboard/rma',{items:rmaData});} }
function updateRMANotes(id,val) { const r=rmaData.find(r=>r.id===id); if(r){r.notes=val;clearTimeout(rmaTimer);rmaTimer=setTimeout(async()=>{await window._saveDoc('dashboard/rma',{items:rmaData});},800);} }
async function removeRMA(id) { rmaData=rmaData.filter(r=>r.id!==id); renderRMA(); await window._saveDoc('dashboard/rma',{items:rmaData}); }

function renderRMA() {
  const open=rmaData.filter(r=>r.status!=='Resolved').length;
  document.getElementById('rma-count').textContent=open>0?`· ${open} open`:'';
  const wrap=document.getElementById('rma-wrap');
  if(!rmaData.length){wrap.innerHTML='<div class="rma-empty">No open RMAs — add one above</div>';return;}
  wrap.innerHTML=`<table><thead><tr><th>Client</th><th>Vendor</th><th>Item / Description</th><th>RMA #</th><th>Date Submitted</th><th>Status</th><th>Notes</th><th></th></tr></thead><tbody>${rmaData.map(r=>`<tr class="fade-in ${r.status==='Resolved'?'row-received':''}"><td><span class="project-name">${r.client||'<span style="opacity:0.3">—</span>'}</span></td><td><span class="vendor-name">${r.vendor}</span></td><td style="max-width:260px">${r.item}</td><td><span class="po-num">${r.number||'<span style="opacity:0.3">—</span>'}</span></td><td style="white-space:nowrap;color:var(--muted);font-size:12px">${r.date?new Date(r.date+'T00:00:00').toLocaleDateString('en-US',{month:'numeric',day:'numeric',year:'numeric'}):'—'}</td><td><select onchange="updateRMAStatus(${r.id},this.value)" style="background:var(--bg3);border:1px solid var(--border);border-radius:6px;padding:3px 8px;font-size:11px;color:var(--text);font-family:'DM Sans',sans-serif;outline:none;cursor:pointer">${['Submitted','Pending Vendor','Replacement Ordered','Resolved'].map(s=>`<option value="${s}" ${s===r.status?'selected':''}>${s}</option>`).join('')}</select></td><td><input type="text" value="${r.notes}" oninput="updateRMANotes(${r.id},this.value)" placeholder="Add notes…" style="background:transparent;border:none;border-bottom:1px solid var(--border);color:var(--text);font-size:12px;font-family:'DM Sans',sans-serif;outline:none;width:100%;padding:2px 0;min-width:140px"></td><td><button onclick="removeRMA(${r.id})" style="background:transparent;border:none;color:var(--muted);cursor:pointer;font-size:14px;padding:4px 8px">✕</button></td></tr>`).join('')}</tbody></table>`;
}

// Fix Chart.js resize
window.addEventListener('resize', () => {
  if (invChart) {
    invChart.resize();
  }
});

// Init on load
window.addEventListener('DOMContentLoaded', ()=>{
  document.getElementById('inv-date').value = new Date().toISOString().split('T')[0];
  document.getElementById('rma-date').value = new Date().toISOString().split('T')[0];
  renderChecklist();
  initLogisticsState();
  renderLogistics();
  renderRMA();
  waitForFirebase(()=>{ listenPO(); listenShowroom(); listenLogistics(); listenInventory(); listenRMA(); listenFleet(); });
});
// ═══════════════════════════════════════════════════════════════════════════════
// FLEET HUB — Firebase
// ═══════════════════════════════════════════════════════════════════════════════
let fleetData = [], serviceData = [], contactsData = [];
let fleetModalId = null; // vehicle being edited in modal

// ── Registration badge helper ─────────────────────────────────────────────────
function regBadge(expDate) {
  if (!expDate) return '<span class="reg-badge reg-none">No Date</span>';
  const today = new Date(); today.setHours(0,0,0,0);
  const exp   = new Date(expDate + 'T00:00:00');
  const days  = Math.round((exp - today) / 86400000);
  const label = exp.toLocaleDateString('en-US', {month:'short', day:'numeric', year:'numeric'});
  if (days < 0)  return `<span class="reg-badge reg-expired">EXPIRED · ${label}</span>`;
  if (days <= 30) return `<span class="reg-badge reg-red">Exp. ${label}</span>`;
  if (days <= 60) return `<span class="reg-badge reg-yellow">Exp. ${label}</span>`;
  return `<span class="reg-badge reg-green">${label}</span>`;
}

function regDaysUntil(expDate) {
  if (!expDate) return Infinity;
  const today = new Date(); today.setHours(0,0,0,0);
  const exp   = new Date(expDate + 'T00:00:00');
  return Math.round((exp - today) / 86400000);
}

// ── Registration alert banner ─────────────────────────────────────────────────
function renderRegBanner() {
  const banner = document.getElementById('fleet-reg-banner');
  const urgent = fleetData.filter(v => regDaysUntil(v.registration) <= 60);
  if (!urgent.length) { banner.style.display = 'none'; return; }
  banner.style.display = 'block';
  banner.innerHTML = `<span class="reg-banner-icon">⚠</span> <strong>Registration Alert:</strong> ` +
    urgent.map(v => {
      const d = regDaysUntil(v.registration);
      return `<span class="reg-banner-item ${d < 0 ? 'expired' : d <= 30 ? 'soon' : 'upcoming'}">${v.name} ${d < 0 ? '(EXPIRED)' : `(${d}d)`}</span>`;
    }).join('  ·  ');
}

// ── Fleet listeners ───────────────────────────────────────────────────────────
function listenFleet() {
  waitForFirebase(() => {
    window._onSnapshot(window._doc(window._db, 'dashboard', 'fleet'), snap => {
      if (snap.exists() && snap.data().vehicles) fleetData = snap.data().vehicles;
      renderFleetCards();
      renderRegBanner();
      renderServiceRequests();
    });
    window._onSnapshot(window._doc(window._db, 'dashboard', 'service'), snap => {
      if (snap.exists() && snap.data().requests) serviceData = snap.data().requests;
      renderFleetCards();
      renderServiceRequests();
    });
    window._onSnapshot(window._doc(window._db, 'dashboard', 'contacts'), snap => {
      if (snap.exists() && snap.data().contacts) contactsData = snap.data().contacts;
      renderContacts();
      renderServiceRequests(); // re-render so vendor dropdowns update
    });
  });
}

// ── Fleet cards ───────────────────────────────────────────────────────────────
function renderFleetCards() {
  const grid = document.getElementById('fleet-grid');
  if (!fleetData.length) {
    grid.innerHTML = '<div class="fleet-empty">No vehicles added yet — click <strong>+ Add Vehicle</strong> to get started</div>';
    return;
  }
  // Sort by vehicle number if present
  const sorted = [...fleetData].sort((a,b) => {
    const an = parseInt(a.number) || 999, bn = parseInt(b.number) || 999;
    return an - bn;
  });
  grid.innerHTML = sorted.map(v => {
    const open = serviceData.filter(s => s.vehicleId === v.id && s.status !== 'Resolved').length;
    const vehicleLabel = [v.number ? `#${v.number}` : '', v.year, v.make, v.model].filter(Boolean).join(' ');
    return `
    <div class="fleet-card fade-in">
      ${v.photo ? `<div class="fleet-card-photo"><img src="${v.photo}" alt="${v.name}"></div>` : '<div class="fleet-card-photo no-photo"><span>🚐</span></div>'}
      <div class="fleet-card-inner">
        <div class="fleet-card-header">
          <div>
            <div class="fleet-card-name">${v.name || '—'}${v.number ? `<span class="fleet-num-tag">#${v.number}</span>` : ''}</div>
            <div class="fleet-card-ymm">${[v.year, v.make, v.model].filter(Boolean).join(' ') || '—'}</div>
          </div>
          ${open > 0 ? `<span class="fleet-service-badge">${open} open</span>` : ''}
        </div>
        <div class="fleet-card-details">
          <div class="fleet-card-row"><span class="fleet-card-label">Plate</span><span class="fleet-card-val mono">${v.plate || '—'}</span></div>
          <div class="fleet-card-row"><span class="fleet-card-label">VIN</span><span class="fleet-card-val vin">${v.vin ? v.vin.toUpperCase() : '—'}</span></div>
          ${v.tiresize ? `<div class="fleet-card-row"><span class="fleet-card-label">Tires</span><span class="fleet-card-val mono">${v.tiresize}</span></div>` : ''}
          <div class="fleet-card-row reg-row"><span class="fleet-card-label">Reg.</span>${regBadge(v.registration)}</div>
        </div>
        ${v.notes ? `<div class="fleet-card-notes">${v.notes}</div>` : ''}
        <div class="fleet-card-actions">
          <button class="fleet-action-btn" onclick="openServiceModal(${v.id})">+ Service Request</button>
          <button class="fleet-action-btn ghost" onclick="openVehicleModal(${v.id})">Edit</button>
          <button class="fleet-action-btn danger" onclick="removeVehicle(${v.id})">✕</button>
        </div>
      </div>
    </div>`;
  }).join('');
}

// ── Vehicle modal (add / edit) ─────────────────────────────────────────────────
function openVehicleModal(id) {
  fleetModalId = id || null;
  const v = id ? fleetData.find(v => v.id === id) : {};
  document.getElementById('fleet-modal-title').textContent = id ? 'Edit Vehicle' : 'Add Vehicle';
  document.getElementById('fm-number').value       = v?.number       || '';
  document.getElementById('fm-name').value         = v?.name         || '';
  document.getElementById('fm-year').value         = v?.year         || '';
  document.getElementById('fm-make').value         = v?.make         || '';
  document.getElementById('fm-model').value        = v?.model        || '';
  document.getElementById('fm-color').value        = v?.color        || '';
  document.getElementById('fm-plate').value        = v?.plate        || '';
  document.getElementById('fm-vin').value          = v?.vin          || '';
  document.getElementById('fm-registration').value = v?.registration || '';
  document.getElementById('fm-tiresize').value     = v?.tiresize     || '';
  document.getElementById('fm-notes').value        = v?.notes        || '';
  // Photo
  const preview = document.getElementById('fm-photo-preview');
  const placeholder = document.getElementById('photo-upload-placeholder');
  const clearBtn = document.getElementById('fm-photo-clear');
  if (v?.photo) {
    preview.src = v.photo; preview.style.display = 'block';
    placeholder.style.display = 'none'; clearBtn.style.display = 'inline-flex';
  } else {
    preview.src = ''; preview.style.display = 'none';
    placeholder.style.display = 'flex'; clearBtn.style.display = 'none';
  }
  document.getElementById('fm-photo-input').value = '';
  document.getElementById('fleet-modal').style.display = 'flex';
}

function handlePhotoUpload(input) {
  if (!input.files[0]) return;
  const file = input.files[0];
  if (file.size > 2 * 1024 * 1024) { alert('Photo must be under 2MB. Try a smaller image.'); return; }
  const reader = new FileReader();
  reader.onload = e => {
    const preview = document.getElementById('fm-photo-preview');
    const placeholder = document.getElementById('photo-upload-placeholder');
    preview.src = e.target.result;
    preview.style.display = 'block';
    placeholder.style.display = 'none';
    document.getElementById('fm-photo-clear').style.display = 'inline-flex';
  };
  reader.readAsDataURL(file);
}

function clearPhoto() {
  document.getElementById('fm-photo-preview').src = '';
  document.getElementById('fm-photo-preview').style.display = 'none';
  document.getElementById('photo-upload-placeholder').style.display = 'flex';
  document.getElementById('fm-photo-clear').style.display = 'none';
  document.getElementById('fm-photo-input').value = '';
}

function closeVehicleModal() {
  document.getElementById('fleet-modal').style.display = 'none';
  fleetModalId = null;
}

async function saveVehicle() {
  const name  = document.getElementById('fm-name').value.trim();
  if (!name) { alert('Assigned technician is required.'); return; }
  const photoPreview = document.getElementById('fm-photo-preview');
  const existingVehicle = fleetModalId ? fleetData.find(v => v.id === fleetModalId) : null;
  const photo = photoPreview.style.display !== 'none' && photoPreview.src
    ? photoPreview.src
    : (existingVehicle?.photo && photoPreview.style.display !== 'none' ? existingVehicle.photo : '');
  const vehicle = {
    id:           fleetModalId || Date.now(),
    number:       document.getElementById('fm-number').value.trim(),
    name,
    year:         document.getElementById('fm-year').value.trim(),
    make:         document.getElementById('fm-make').value.trim(),
    model:        document.getElementById('fm-model').value.trim(),
    color:        document.getElementById('fm-color').value.trim(),
    plate:        document.getElementById('fm-plate').value.trim().toUpperCase(),
    vin:          document.getElementById('fm-vin').value.trim().toUpperCase(),
    registration: document.getElementById('fm-registration').value,
    tiresize:     document.getElementById('fm-tiresize').value.trim(),
    notes:        document.getElementById('fm-notes').value.trim(),
    photo:        photoPreview.src && photoPreview.style.display !== 'none' ? photoPreview.src : '',
  };
  if (fleetModalId) {
    const idx = fleetData.findIndex(v => v.id === fleetModalId);
    if (idx >= 0) fleetData[idx] = vehicle; else fleetData.push(vehicle);
  } else {
    fleetData.push(vehicle);
  }
  closeVehicleModal();
  renderFleetCards(); renderRegBanner();
  await window._saveDoc('dashboard/fleet', { vehicles: fleetData });
}

async function removeVehicle(id) {
  if (!confirm('Remove this vehicle? This cannot be undone.')) return;
  fleetData = fleetData.filter(v => v.id !== id);
  serviceData = serviceData.filter(s => s.vehicleId !== id);
  renderFleetCards(); renderRegBanner(); renderServiceRequests();
  await window._saveDoc('dashboard/fleet', { vehicles: fleetData });
  await window._saveDoc('dashboard/service', { requests: serviceData });
}

// ── Service request modal ──────────────────────────────────────────────────────
function openServiceModal(vehicleId) {
  document.getElementById('sr-vehicle').value = vehicleId || '';
  // populate vehicle dropdown
  const sel = document.getElementById('sr-vehicle');
  sel.innerHTML = fleetData.map(v => `<option value="${v.id}" ${v.id === vehicleId ? 'selected' : ''}>${v.name}</option>`).join('');
  // populate vendor dropdown
  const vsel = document.getElementById('sr-vendor');
  vsel.innerHTML = '<option value="">— None —</option>' + contactsData.map(c => `<option value="${c.id}">${c.name}</option>`).join('');
  document.getElementById('sr-issue').value = '';
  document.getElementById('sr-date-scheduled').value = '';
  document.getElementById('sr-notes').value = '';
  document.getElementById('sr-status').value = 'Pending';
  document.getElementById('service-modal').style.display = 'flex';
}

function closeServiceModal() {
  document.getElementById('service-modal').style.display = 'none';
}

async function saveServiceRequest() {
  const vehicleId = parseInt(document.getElementById('sr-vehicle').value);
  const issue     = document.getElementById('sr-issue').value.trim();
  if (!vehicleId || !issue) { alert('Please select a vehicle and describe the issue.'); return; }
  const vehicle   = fleetData.find(v => v.id === vehicleId);
  const vendorId  = parseInt(document.getElementById('sr-vendor').value) || null;
  const vendor    = vendorId ? contactsData.find(c => c.id === vendorId) : null;
  const today     = new Date().toISOString().split('T')[0];
  serviceData.push({
    id:            Date.now(),
    vehicleId,
    vehicleName:   vehicle?.name || '',
    issue,
    status:        document.getElementById('sr-status').value,
    vendorId,
    vendorName:    vendor?.name || '',
    dateSubmitted: today,
    dateScheduled: document.getElementById('sr-date-scheduled').value,
    notes:         document.getElementById('sr-notes').value.trim(),
  });
  closeServiceModal();
  renderFleetCards(); renderServiceRequests();
  await window._saveDoc('dashboard/service', { requests: serviceData });
}

// ── Service requests list ──────────────────────────────────────────────────────
function renderServiceRequests() {
  const open = serviceData.filter(s => s.status !== 'Resolved').length;
  document.getElementById('service-count').textContent = open > 0 ? `· ${open} open` : '';
  const wrap = document.getElementById('service-wrap');
  if (!serviceData.length) { wrap.innerHTML = '<div class="rma-empty">No service requests yet</div>'; return; }
  wrap.innerHTML = `<table><thead><tr>
    <th>Vehicle</th><th>Issue</th><th>Status</th><th>Vendor</th>
    <th>Submitted</th><th>Scheduled</th><th>Notes</th><th></th>
  </tr></thead><tbody>${serviceData.map(r => {
    const statusColors = { 'Pending':'badge-draft','Scheduled':'badge-issued','In Progress':'badge-partial','Resolved':'badge-received' };
    const vendorOpts = ['','...contactsData...'].join(''); // built inline below
    const vOpts = '<option value="">— None —</option>' + contactsData.map(c => `<option value="${c.id}" ${c.id === r.vendorId ? 'selected' : ''}>${c.name}</option>`).join('');
    return `<tr class="fade-in ${r.status==='Resolved'?'row-received':''}">
      <td><span class="vendor-name">${r.vehicleName}</span></td>
      <td style="max-width:220px;font-size:12px">${r.issue}</td>
      <td>
        <select onchange="updateServiceStatus(${r.id},this.value)" style="background:var(--bg3);border:1px solid var(--border);border-radius:2px;padding:3px 8px;font-size:11px;color:var(--text);font-family:'Karla',sans-serif;outline:none;cursor:pointer">
          ${['Pending','Scheduled','In Progress','Resolved'].map(s=>`<option value="${s}" ${s===r.status?'selected':''}>${s}</option>`).join('')}
        </select>
      </td>
      <td>
        <select onchange="updateServiceVendor(${r.id},this.value)" style="background:var(--bg3);border:1px solid var(--border);border-radius:2px;padding:3px 8px;font-size:11px;color:var(--text);font-family:'Karla',sans-serif;outline:none;cursor:pointer;max-width:140px">
          ${vOpts}
        </select>
      </td>
      <td style="white-space:nowrap;color:var(--muted);font-size:12px">${r.dateSubmitted ? new Date(r.dateSubmitted+'T00:00:00').toLocaleDateString('en-US',{month:'numeric',day:'numeric',year:'numeric'}) : '—'}</td>
      <td style="white-space:nowrap;color:var(--muted);font-size:12px">${r.dateScheduled ? new Date(r.dateScheduled+'T00:00:00').toLocaleDateString('en-US',{month:'numeric',day:'numeric',year:'numeric'}) : '—'}</td>
      <td><input type="text" value="${r.notes||''}" oninput="updateServiceNotes(${r.id},this.value)" placeholder="Notes…" style="background:transparent;border:none;border-bottom:1px solid var(--border);color:var(--text);font-size:12px;font-family:'Karla',sans-serif;outline:none;width:100%;padding:2px 0;min-width:130px"></td>
      <td><button onclick="removeService(${r.id})" style="background:transparent;border:none;color:var(--muted);cursor:pointer;font-size:14px;padding:4px 8px;transition:color 0.15s" onmouseover="this.style.color='#dc2626'" onmouseout="this.style.color='var(--muted)'">✕</button></td>
    </tr>`;
  }).join('')}</tbody></table>`;
}

let serviceTimer;
async function updateServiceStatus(id, val) {
  const r = serviceData.find(r => r.id === id); if (!r) return;
  r.status = val; renderFleetCards(); renderServiceRequests();
  await window._saveDoc('dashboard/service', { requests: serviceData });
}
async function updateServiceVendor(id, val) {
  const r = serviceData.find(r => r.id === id); if (!r) return;
  const vendorId = parseInt(val) || null;
  const vendor = vendorId ? contactsData.find(c => c.id === vendorId) : null;
  r.vendorId = vendorId; r.vendorName = vendor?.name || '';
  await window._saveDoc('dashboard/service', { requests: serviceData });
}
function updateServiceNotes(id, val) {
  const r = serviceData.find(r => r.id === id); if (!r) return;
  r.notes = val; clearTimeout(serviceTimer);
  serviceTimer = setTimeout(async () => { await window._saveDoc('dashboard/service', { requests: serviceData }); }, 800);
}
async function removeService(id) {
  serviceData = serviceData.filter(r => r.id !== id);
  renderFleetCards(); renderServiceRequests();
  await window._saveDoc('dashboard/service', { requests: serviceData });
}

// ── Contacts ──────────────────────────────────────────────────────────────────
let contactModalId = null;

function openContactModal(id) {
  contactModalId = id || null;
  const c = id ? contactsData.find(c => c.id === id) : {};
  document.getElementById('contact-modal-title').textContent = id ? 'Edit Contact' : 'Add Contact';
  document.getElementById('cm-name').value    = c?.name    || '';
  document.getElementById('cm-type').value    = c?.type    || 'Repair Shop';
  document.getElementById('cm-contact').value = c?.contact || '';
  document.getElementById('cm-phone').value   = c?.phone   || '';
  document.getElementById('cm-address').value = c?.address || '';
  document.getElementById('cm-notes').value   = c?.notes   || '';
  document.getElementById('contact-modal').style.display = 'flex';
}

function closeContactModal() {
  document.getElementById('contact-modal').style.display = 'none';
  contactModalId = null;
}

async function saveContact() {
  const name = document.getElementById('cm-name').value.trim();
  if (!name) { alert('Business name is required.'); return; }
  const contact = {
    id:      contactModalId || Date.now(),
    name,
    type:    document.getElementById('cm-type').value,
    contact: document.getElementById('cm-contact').value.trim(),
    phone:   document.getElementById('cm-phone').value.trim(),
    address: document.getElementById('cm-address').value.trim(),
    notes:   document.getElementById('cm-notes').value.trim(),
  };
  if (contactModalId) {
    const idx = contactsData.findIndex(c => c.id === contactModalId);
    if (idx >= 0) contactsData[idx] = contact; else contactsData.push(contact);
  } else {
    contactsData.push(contact);
  }
  closeContactModal();
  renderContacts(); renderServiceRequests();
  await window._saveDoc('dashboard/contacts', { contacts: contactsData });
}

async function removeContact(id) {
  if (!confirm('Remove this contact?')) return;
  contactsData = contactsData.filter(c => c.id !== id);
  renderContacts(); renderServiceRequests();
  await window._saveDoc('dashboard/contacts', { contacts: contactsData });
}

const CONTACT_TYPE_COLORS = {
  'Repair Shop': { bg:'#eff6ff', fg:'#1d4ed8', border:'#0062a4' },
  'Vinyl / Graphics': { bg:'#fdf4ff', fg:'#7e22ce', border:'#a855f7' },
  'Tires': { bg:'#fff7ed', fg:'#c2410c', border:'#f97316' },
  'Glass': { bg:'#f0fdf4', fg:'#15803d', border:'#22c55e' },
  'Other': { bg:'#f5f6f8', fg:'#888896', border:'#d0d5dd' },
};

function renderContacts() {
  const wrap = document.getElementById('contacts-wrap');
  if (!contactsData.length) { wrap.innerHTML = '<div class="rma-empty">No contacts yet — add one above</div>'; return; }
  const byType = {};
  contactsData.forEach(c => { if (!byType[c.type]) byType[c.type] = []; byType[c.type].push(c); });
  wrap.innerHTML = Object.entries(byType).map(([type, contacts]) => {
    const col = CONTACT_TYPE_COLORS[type] || CONTACT_TYPE_COLORS['Other'];
    return `<div class="contacts-group">
      <div class="contacts-group-label" style="color:${col.fg}">${type}</div>
      <div class="contacts-grid">${contacts.map(c => `
        <div class="contact-card fade-in">
          <div class="contact-card-header">
            <span class="contact-name">${c.name}</span>
            <span class="contact-type-badge" style="background:${col.bg};color:${col.fg};border-color:${col.border}">${c.type}</span>
          </div>
          ${c.contact ? `<div class="contact-row"><span class="contact-label">Contact</span><span>${c.contact}</span></div>` : ''}
          ${c.phone   ? `<div class="contact-row"><span class="contact-label">Phone</span><a href="tel:${c.phone}" class="contact-phone">${c.phone}</a></div>` : ''}
          ${c.address ? `<div class="contact-row"><span class="contact-label">Address</span><span style="font-size:12px;color:var(--muted)">${c.address}</span></div>` : ''}
          ${c.notes   ? `<div class="contact-notes">${c.notes}</div>` : ''}
          <div class="contact-actions">
            <button class="fleet-action-btn ghost" onclick="openContactModal(${c.id})">Edit</button>
            <button class="fleet-action-btn danger" onclick="removeContact(${c.id})">✕</button>
          </div>
        </div>`).join('')}
      </div>
    </div>`;
  }).join('');
}

// Close modals on backdrop click
document.addEventListener('click', e => {
  if (e.target.id === 'fleet-modal')   closeVehicleModal();
  if (e.target.id === 'service-modal') closeServiceModal();
  if (e.target.id === 'contact-modal') closeContactModal();
});
