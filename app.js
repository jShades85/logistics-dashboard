// ═══════════════════════════════════════════════════════════════════════════════
// SLACK CONFIG — Order Request Integration
// ═══════════════════════════════════════════════════════════════════════════════
const SLACK_TOKEN   = 'PASTE_YOUR_TOKEN_HERE';   // xoxb-...
const SLACK_CHANNEL = 'C068G11H94P';             // #team-purchase-orders
const SLACK_RESOLVE_EMOJI = 'white_check_mark';  // ✅
const SLACK_POLL_INTERVAL = 60000;               // 60 seconds

let orderRequests = [];
let slackPollTimer = null;

async function fetchOrderRequests() {
  if (!SLACK_TOKEN || SLACK_TOKEN === 'PASTE_YOUR_TOKEN_HERE') return;
  try {
    const res = await fetch(
      `https://slack.com/api/conversations.history?channel=${SLACK_CHANNEL}&limit=100`,
      { headers: { Authorization: `Bearer ${SLACK_TOKEN}` } }
    );
    const data = await res.json();
    if (!data.ok) { console.warn('Slack API error:', data.error); return; }

    const msgs = (data.messages || []).filter(m => {
      // Only workflow bot messages that contain our expected fields
      if (!m.text) return false;
      const t = m.text;
      return t.includes('Port Number:') && t.includes('Phase:') && t.includes('Submitted by:');
    });

    // Filter out messages already resolved with ✅
    orderRequests = msgs
      .filter(m => {
        const reactions = m.reactions || [];
        return !reactions.some(r => r.name === SLACK_RESOLVE_EMOJI);
      })
      .map(m => {
        const text = m.text;
        const portMatch = text.match(/Port Number:\s*(\S+)/i);
        const phaseMatch = text.match(/Phase:\s*([^\n]+)/i);
        const submittedMatch = text.match(/Submitted by:\s*<@([^>]+)>|Submitted by:\s*([^\n]+)/i);
        return {
          ts: m.ts,
          port: portMatch ? portMatch[1].trim() : '—',
          phase: phaseMatch ? phaseMatch[1].trim() : '—',
          submitted: submittedMatch ? (submittedMatch[2] || submittedMatch[1] || '—').trim() : '—',
          time: new Date(parseFloat(m.ts) * 1000).toLocaleString('en-US', {
            month: 'short', day: 'numeric', hour: 'numeric', minute: '2-digit'
          }),
        };
      });

    renderOrderRequests();
  } catch (err) {
    console.error('Slack fetch error:', err);
  }
}

function renderOrderRequests() {
  const countEl = document.getElementById('kpi-orders');
  const listEl  = document.getElementById('order-request-list');
  if (countEl) countEl.textContent = orderRequests.length || '0';

  if (!listEl) return;
  if (!orderRequests.length) {
    listEl.innerHTML = '<div class="order-empty">No open order requests — react with ✅ in Slack to clear them</div>';
    return;
  }
  listEl.innerHTML = `<table>
    <thead><tr>
      <th>Port #</th><th>Phase</th><th>Submitted By</th><th>Time</th>
    </tr></thead>
    <tbody>
      ${orderRequests.map(r => `
        <tr class="fade-in">
          <td><span class="po-num">${r.port}</span></td>
          <td><span class="badge badge-issued">${r.phase}</span></td>
          <td><span class="vendor-name">${r.submitted}</span></td>
          <td style="white-space:nowrap;color:var(--muted);font-size:12px">${r.time}</td>
        </tr>`).join('')}
    </tbody>
  </table>`;
}

function startSlackPolling() {
  if (!SLACK_TOKEN || SLACK_TOKEN === 'PASTE_YOUR_TOKEN_HERE') {
    const listEl = document.getElementById('order-request-list');
    if (listEl) listEl.innerHTML = '<div class="order-empty">⚠ Add your Slack token to app.js to enable order request sync</div>';
    const countEl = document.getElementById('kpi-orders');
    if (countEl) countEl.textContent = '—';
    return;
  }
  fetchOrderRequests();
  slackPollTimer = setInterval(fetchOrderRequests, SLACK_POLL_INTERVAL);
}

let orderRequestsVisible = true;
function toggleOrderRequests() {
  orderRequestsVisible = !orderRequestsVisible;
  const wrap  = document.getElementById('order-request-wrap');
  const arrow = document.getElementById('order-toggle-arrow');
  if (wrap)  wrap.style.display  = orderRequestsVisible ? '' : 'none';
  if (arrow) arrow.textContent   = orderRequestsVisible ? '▲ Hide' : '▼ Show';
}

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
  startSlackPolling();

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
  waitForFirebase(()=>{ listenPO(); listenShowroom(); listenLogistics(); listenInventory(); listenRMA(); });
});