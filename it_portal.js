// ── CONSTANTS ────────────────────────────────────────────────
const SP_SITE   = 'https://furuyath.sharepoint.com/sites/ITHelpdesk';
const LIST_NAME = 'ตัวติดตามปัญหา';

// ── MSAL CONFIG ──────────────────────────────────────────────
const MSAL_CONFIG = {
  auth: {
    clientId: '9c3dd0c6-96e6-43a0-8b79-9cb66040636c',
    authority: 'https://login.microsoftonline.com/258c51b5-9907-453b-ae52-2f4d2acb7f00',
    redirectUri: window.location.origin + window.location.pathname
  },
  cache: { cacheLocation: 'sessionStorage' }
};
const SP_SCOPES   = ['https://furuyath.sharepoint.com/.default'];
const GRAPH_SCOPES = ['https://graph.microsoft.com/.default'];
const msalInstance = new msal.PublicClientApplication(MSAL_CONFIG);

// Role map — เพิ่ม email ได้ที่นี่
const ROLE_MAP = {
  'nutthawut@furuya.co.th': 'it_admin',
  'it@furuya.co.th':        'it_staff',
  'anuchit.po@furuya.co.th':'gm'
};

function autoSetRole(account) {
  const email = (account?.username || '').toLowerCase();
  const role = ROLE_MAP[email] || 'employee';
  setRole(role);
  // ซ่อน dropdown ถ้า role ถูก detect อัตโนมัติ
  const sel = document.getElementById('role-select');
  if (sel) { sel.value = role; sel.closest('.role-wrap').style.display = 'none'; }
  console.log(`MSAL: logged in as ${email} → role: ${role}`);
}

// Handle redirect response first — then auto-login if no account
// flag บอกว่า redirect ครั้งนี้เป็น Graph consent หรือเปล่า
const _isGraphConsent = sessionStorage.getItem('graphConsentPending') === '1';

const _authReady = msalInstance.handleRedirectPromise()
  .then(response => {
    if (response) {
      console.log('MSAL: redirect OK, scopes:', response.scopes);
      sessionStorage.removeItem('graphConsentPending');
    }
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length === 0) {
      msalInstance.loginRedirect({ scopes: SP_SCOPES });
    } else {
      autoSetRole(accounts[0]);
      const returnTab = sessionStorage.getItem('returnTab');
      if (returnTab) {
        sessionStorage.removeItem('returnTab');
        setTimeout(() => document.querySelector(`.nav-tab[data-tab="${returnTab}"]`)?.click(), 300);
      }
    }
  })
  .catch(err => console.error('MSAL redirect error:', err));

// เรียกเมื่อกดปุ่ม "เชื่อมต่อ Email" — redirect ขอ Graph token แยกต่างหาก
function ops_connectEmail() {
  const account = msalInstance.getAllAccounts()[0];
  if (!account) return;
  sessionStorage.setItem('graphConsentPending', '1');
  sessionStorage.setItem('returnTab', 'ops');
  msalInstance.acquireTokenRedirect({ scopes: GRAPH_SCOPES, account });
}

async function getToken(scopes) {
  await _authReady;
  const account = msalInstance.getAllAccounts()[0];
  if (!account) {
    msalInstance.loginRedirect({ scopes: SP_SCOPES });
    return new Promise(() => {});
  }
  try {
    const r = await msalInstance.acquireTokenSilent({ scopes, account });
    return r.accessToken;
  } catch {
    // SP scope: redirect login ใหม่ / Graph scope: throw ให้ caller แสดง demo
    if (scopes.some(s => s.includes('sharepoint.com'))) {
      msalInstance.acquireTokenRedirect({ scopes: SP_SCOPES, account });
      return new Promise(() => {});
    }
    throw new Error('cannot acquire Graph token silently');
  }
}

async function sp_headers(extra = {}) {
  const token = await getToken(SP_SCOPES);
  return { 'Accept': 'application/json;odata=verbose', 'Authorization': 'Bearer ' + token, ...extra };
}

// ── SHARED ──────────────────────────────────────────────────
let tabLoaded = {};

function setRole(role) {
  document.querySelectorAll('.nav-tab[data-roles]').forEach(btn => {
    const ok = btn.getAttribute('data-roles').split(',').includes(role);
    btn.hidden = !ok;
    if (!ok && btn.classList.contains('active')) {
      document.querySelector('.nav-tab[data-tab="form"]').click();
    }
  });
}

function showTab(name, el) {
  document.querySelectorAll('.tab-page').forEach(p => p.classList.remove('active'));
  document.getElementById('tab-' + name).classList.add('active');
  document.querySelectorAll('.nav-tab').forEach(t => t.classList.remove('active'));
  if (el) el.classList.add('active');
  if (!tabLoaded[name]) {
    tabLoaded[name] = true;
    if (name === 'mytickets') frm_loadMyTickets();
    if (name === 'staff')     std_load();
    if (name === 'manager')   mgr_load();
    if (name === 'ops')       ops_loadAll();
  }
}

function showToast(msg, isErr = false) {
  const t = document.getElementById('toast');
  t.textContent = msg; t.className = 'toast' + (isErr ? ' error' : '');
  t.style.display = 'block';
  setTimeout(() => t.style.display = 'none', 4500);
}

async function sp_digest() {
  const r = await fetch(`${SP_SITE}/_api/contextinfo`, {
    method: 'POST',
    headers: await sp_headers({ 'Content-Type': 'application/json;odata=verbose' })
  });
  return (await r.json()).d.GetContextWebInformation.FormDigestValue;
}

async function sp_entityType() {
  const r = await fetch(
    `${SP_SITE}/_api/web/lists/getbytitle('${encodeURIComponent(LIST_NAME)}')?$select=ListItemEntityTypeFullName`,
    { headers: await sp_headers() }
  );
  return (await r.json()).d.ListItemEntityTypeFullName;
}

// ── FORM TAB ─────────────────────────────────────────────────
const SLA_MAP = { Critical:{h:2,l:'2 ชั่วโมง'}, High:{h:4,l:'4 ชั่วโมง'}, Medium:{h:8,l:'8 ชั่วโมง'}, Low:{h:24,l:'24 ชั่วโมง'} };
let frm_prio = 'Medium';

function frm_setPrio(p) {
  frm_prio = p;
  document.querySelectorAll('.prio-btn').forEach(b => b.classList.remove('selected'));
  document.getElementById('p-' + p).classList.add('selected');
  document.getElementById('sla-hint').textContent = `⏱ SLA: ทีม IT จะแก้ไขภายใน ${SLA_MAP[p].l} (${p})`;
}

function frm_clear() {
  ['f-name','f-email','f-dept','f-category','f-title','f-detail'].forEach(id => document.getElementById(id).value = '');
  frm_setPrio('Medium');
}

function frm_genId() {
  const d = new Date();
  return 'TK' + String(d.getFullYear()).slice(-2) + String(d.getMonth()+1).padStart(2,'0') + String(d.getDate()).padStart(2,'0') + '-' + (Math.floor(Math.random()*9000)+1000);
}

async function frm_submit() {
  const v = id => document.getElementById(id).value.trim();
  const name = v('f-name'), email = v('f-email'), dept = document.getElementById('f-dept').value,
        cat  = document.getElementById('f-category').value, title = v('f-title'), detail = v('f-detail');
  if (!name||!email||!dept||!cat||!title||!detail) { showToast('กรุณากรอกข้อมูลให้ครบทุกช่อง', true); return; }
  const btn = document.getElementById('submit-btn');
  btn.disabled = true; btn.textContent = 'กำลังส่ง...';
  const ticketId = frm_genId();
  try {
    const [digest, entityType] = await Promise.all([sp_digest(), sp_entityType()]);
    const res = await fetch(`${SP_SITE}/_api/web/lists/getbytitle('${encodeURIComponent(LIST_NAME)}')/items`, {
      method: 'POST',
      headers: await sp_headers({
        'Content-Type': 'application/json;odata=verbose',
        'X-RequestDigest': digest
      }),
      body: JSON.stringify({
        __metadata: { type: entityType },
        Title: title, RequesterName: name, RequesterEmail: email,
        Department: dept, Category: cat, Description: detail,
        Priority: frm_prio, Status: 'ใหม่', TicketID: ticketId,
        SLA_Deadline: new Date(Date.now() + SLA_MAP[frm_prio].h * 3600000).toISOString(),
        SLA_Status: 'On Track'
      })
    });
    if (!res.ok) { const e = await res.json(); throw new Error(e.error?.message?.value || res.statusText); }
    showToast('✅ ส่งสำเร็จ! Ticket ID: ' + ticketId);
    try {
      const saved = JSON.parse(localStorage.getItem('myTickets')||'[]');
      saved.unshift({ ticketId, title, dept, cat, priority: frm_prio, status: 'ใหม่', created: new Date().toISOString() });
      localStorage.setItem('myTickets', JSON.stringify(saved.slice(0,50)));
    } catch(e) {}
    frm_clear();
  } catch(e) {
    showToast('เกิดข้อผิดพลาด: ' + e.message, true);
  } finally {
    btn.disabled = false; btn.textContent = 'ส่งแจ้งซ่อม →';
  }
}

function frm_loadMyTickets() {
  let saved = [];
  try { saved = JSON.parse(localStorage.getItem('myTickets')||'[]'); } catch(e) {}
  const el = document.getElementById('my-tickets-list');
  if (!saved.length) { el.innerHTML = '<div class="empty-state"><div class="icon">📭</div><p>ยังไม่มี Ticket ที่ส่งจากอุปกรณ์นี้</p></div>'; return; }
  const sb = s => { const m={'ใหม่':'b-new','กำลังดำเนินการ':'b-prog','รอวัสดุ':'b-wait','แก้ไขแล้ว':'b-done'}; return `<span class="badge ${m[s]||'b-new'}">${s||'ใหม่'}</span>`; };
  const pb = p => { const m={Critical:'b-crit',High:'b-high',Medium:'b-med',Low:'b-low'}; return `<span class="badge ${m[p]||'b-med'}">${p}</span>`; };
  el.innerHTML = saved.map(t => `
    <div class="ticket-card"><div class="ticket-top"><div>
      <div class="ticket-id">${t.ticketId}</div>
      <div class="ticket-title">${t.title}</div>
      <div class="ticket-meta">${t.dept} · ${t.cat} · ${new Date(t.created).toLocaleDateString('th-TH')}</div>
    </div><div style="display:flex;flex-direction:column;gap:6px;align-items:flex-end;">${sb(t.status)}${pb(t.priority)}</div></div></div>`).join('');
}

// ── STAFF TAB ────────────────────────────────────────────────
let std_tickets = [], std_curId = null;

async function std_load() {
  document.getElementById('staff-tbody').innerHTML = '<tr><td colspan="7" class="loading-td"><div class="spinner"></div> กำลังโหลด...</td></tr>';
  try {
    const r = await fetch(
      `${SP_SITE}/_api/web/lists/getbytitle('${encodeURIComponent(LIST_NAME)}')/items?$orderby=Created desc&$top=200&$select=ID,Title,TicketID,RequesterName,RequesterEmail,Department,Category,Priority,Status,SLA_Deadline,Description,ResolutionNote,ConfirmDate`,
      { headers: await sp_headers() }
    );
    if (!r.ok) throw new Error();
    std_tickets = (await r.json()).d.results;
    std_kpi(); std_renderTable(std_tickets);
  } catch(e) {
    document.getElementById('staff-tbody').innerHTML = '<tr><td colspan="7" class="loading-td">❌ โหลดไม่ได้ — กรุณา Login ใหม่</td></tr>';
  }
}

function std_kpi() {
  const c = s => std_tickets.filter(t => t.Status === s).length;
  const breach = std_tickets.filter(t => t.SLA_Deadline && new Date(t.SLA_Deadline) < new Date() && t.Status !== 'แก้ไขแล้ว').length;
  document.getElementById('sk-new').textContent    = c('ใหม่');
  document.getElementById('sk-prog').textContent   = c('กำลังดำเนินการ');
  document.getElementById('sk-wait').textContent   = c('รอวัสดุ');
  document.getElementById('sk-done').textContent   = c('แก้ไขแล้ว');
  document.getElementById('sk-breach').textContent = breach;
  document.getElementById('fc-all').textContent    = std_tickets.length;
  document.getElementById('fc-new').textContent    = c('ใหม่');
  document.getElementById('fc-prog').textContent   = c('กำลังดำเนินการ');
}

function std_filter(f, el) {
  document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
  if (el) el.classList.add('active');
  std_renderTable(f === 'all' ? std_tickets : std_tickets.filter(t => t.Status === f));
}

function std_sla(deadline, status) {
  if (!deadline || status === 'แก้ไขแล้ว') return ['#4ade80', 100];
  const now = Date.now(), dl = new Date(deadline).getTime();
  if (dl < now) return ['#f87171', 100];
  const pct = Math.max(0, Math.min(100, ((dl - now) / (8*3600000)) * 100));
  return [pct < 25 ? '#fb923c' : '#4ade80', pct];
}

function std_pb(p) { const m={Critical:'b-crit',High:'b-high',Medium:'b-med',Low:'b-low'}; return `<span class="badge ${m[p]||'b-med'}">${p||'—'}</span>`; }
function std_sb(s) { const m={'ใหม่':'b-new','กำลังดำเนินการ':'b-prog','รอวัสดุ':'b-wait','แก้ไขแล้ว':'b-done'}; return `<span class="badge ${m[s]||'b-new'}">${s||'ใหม่'}</span>`; }

function std_renderTable(items) {
  if (!items.length) { document.getElementById('staff-tbody').innerHTML = '<tr><td colspan="7" class="loading-td">ไม่มี Ticket ในหมวดนี้</td></tr>'; return; }
  document.getElementById('staff-tbody').innerHTML = items.map(t => {
    const [color, pct] = std_sla(t.SLA_Deadline, t.Status);
    return `<tr>
      <td><span style="color:#0070C0;font-size:12px;font-weight:600">${t.TicketID||'#'+t.ID}</span></td>
      <td style="max-width:200px"><div style="font-weight:600">${t.Title||''}</div><div style="font-size:11px;color:#666">${t.Category||''}</div></td>
      <td><div>${t.RequesterName||''}</div><div style="font-size:11px;color:#666">${t.Department||''}</div></td>
      <td>${std_pb(t.Priority)}</td><td>${std_sb(t.Status)}</td>
      <td style="min-width:80px"><div style="font-size:11px;color:#666">${t.SLA_Deadline?new Date(t.SLA_Deadline).toLocaleDateString('th-TH'):'-'}</div><div class="sla-bar"><div class="sla-fill" style="width:${pct}%;background:${color}"></div></div></td>
      <td><button class="btn-sm btn-manage" onclick="std_openPanel(${t.ID})">จัดการ</button></td>
    </tr>`;
  }).join('');
}

function std_openPanel(id) {
  const t = std_tickets.find(x => x.ID === id);
  if (!t) return;
  std_curId = id;
  document.getElementById('sp-title').textContent = t.TicketID || '#'+id;
  document.getElementById('sp-info').innerHTML = `<div style="background:#f0f5f0;border-radius:8px;padding:14px;margin-bottom:16px;font-size:13px">
    <div style="font-weight:600;font-size:15px;margin-bottom:8px">${t.Title}</div>
    <div style="color:#888;margin-bottom:6px;white-space:pre-wrap">${t.Description||''}</div>
    <div style="font-size:12px;color:#666">ผู้แจ้ง: ${t.RequesterName||''} (${t.RequesterEmail||''}) · ${t.Department||''}</div>
  </div>`;
  document.getElementById('sp-status').value      = t.Status || 'ใหม่';
  document.getElementById('sp-resolution').value  = t.ResolutionNote || '';
  document.getElementById('sp-confirm-date').value = t.ConfirmDate ? new Date(t.ConfirmDate).toISOString().slice(0,16) : '';
  document.getElementById('staff-panel').classList.add('open');
}

function std_closePanel() {
  document.getElementById('staff-panel').classList.remove('open');
  std_curId = null;
}

async function std_save() {
  if (!std_curId) return;
  const btn = document.getElementById('sp-save-btn');
  btn.textContent = 'กำลังบันทึก...'; btn.disabled = true;
  try {
    const [digest, entityType] = await Promise.all([sp_digest(), sp_entityType()]);
    const status = document.getElementById('sp-status').value;
    const payload = {
      __metadata: { type: entityType },
      Status: status,
      ResolutionNote: document.getElementById('sp-resolution').value
    };
    if (status === 'แก้ไขแล้ว') {
      const cd = document.getElementById('sp-confirm-date').value;
      payload.ConfirmDate = cd ? new Date(cd).toISOString() : new Date().toISOString();
    }
    const res = await fetch(
      `${SP_SITE}/_api/web/lists/getbytitle('${encodeURIComponent(LIST_NAME)}')/items(${std_curId})`,
      {
        method: 'POST',
        headers: await sp_headers({
          'Content-Type': 'application/json;odata=verbose',
          'X-RequestDigest': digest,
          'X-HTTP-Method': 'MERGE',
          'IF-MATCH': '*'
        }),
        body: JSON.stringify(payload)
      }
    );
    if (!res.ok && res.status !== 204) throw new Error('Update failed');
    showToast('✅ บันทึกสำเร็จ!');
    std_closePanel(); std_load();
  } catch(e) {
    showToast('❌ เกิดข้อผิดพลาด: ' + e.message, true);
  } finally {
    btn.textContent = '💾 บันทึก'; btn.disabled = false;
  }
}

// ── MANAGER TAB ──────────────────────────────────────────────
let mgr_data = [], mgr_tc, mgr_sc, mgr_pc;

async function mgr_load() {
  try {
    const r = await fetch(
      `${SP_SITE}/_api/web/lists/getbytitle('${encodeURIComponent(LIST_NAME)}')/items?$top=500&$orderby=Created desc&$select=ID,Title,TicketID,Department,Priority,Status,SLA_Deadline,Created`,
      { headers: await sp_headers() }
    );
    if (!r.ok) throw new Error();
    mgr_data = (await r.json()).d.results;
  } catch(e) {
    mgr_data = mgr_demo();
  }
  mgr_filter();
}

function mgr_demo() {
  const depts=['IT Dept','Admin Dept','Assembly Dept','Account Dept','QA Dept','Marketing Dept'];
  const prios=['Critical','High','Medium','Low'];
  const sts=['ใหม่','กำลังดำเนินการ','แก้ไขแล้ว','แก้ไขแล้ว','แก้ไขแล้ว'];
  return Array.from({length:42},(_,i)=>{
    const c=new Date(Date.now()-Math.random()*30*86400000).toISOString(),p=prios[i%4];
    return {ID:i+1,TicketID:`TK2605${String(i+1).padStart(4,'0')}`,Title:`Ticket #${i+1}`,Department:depts[i%depts.length],Priority:p,Status:sts[i%sts.length],SLA_Deadline:new Date(new Date(c).getTime()+{Critical:2,High:4,Medium:8,Low:24}[p]*3600000).toISOString(),Created:c};
  });
}

function mgr_filter() {
  const months = parseInt(document.getElementById('mgr-month').value);
  const cutoff = new Date();
  if (months===0) cutoff.setDate(1); else cutoff.setMonth(cutoff.getMonth()-months);
  const d = mgr_data.filter(t => new Date(t.Created) >= cutoff);
  mgr_kpi(d); mgr_trend(d); mgr_sla(d); mgr_prio(d); mgr_dept(d); mgr_breach(d);
}

function mgr_kpi(d) {
  const now=new Date(),t=d.length,res=d.filter(x=>x.Status==='แก้ไขแล้ว').length;
  const br=d.filter(x=>x.SLA_Deadline&&new Date(x.SLA_Deadline)<now&&x.Status!=='แก้ไขแล้ว').length;
  document.getElementById('mk-total').textContent = t;
  document.getElementById('mk-res').textContent   = res;
  document.getElementById('mk-breach').textContent= br;
  document.getElementById('mk-pend').textContent  = d.filter(x=>x.Status!=='แก้ไขแล้ว').length;
  document.getElementById('mk-sla').textContent   = (t?Math.round(((t-br)/t)*100):100)+'%';
}

function mgr_trend(d) {
  const w={};d.forEach(t=>{const dt=new Date(t.Created),k=`สัปดาห์ ${Math.ceil(dt.getDate()/7)}`;w[k]=w[k]||{t:0,r:0};w[k].t++;if(t.Status==='แก้ไขแล้ว')w[k].r++;});
  const lb=Object.keys(w);if(mgr_tc)mgr_tc.destroy();
  mgr_tc=new Chart(document.getElementById('mgr-trend'),{type:'bar',data:{labels:lb,datasets:[{label:'ทั้งหมด',data:lb.map(k=>w[k].t),backgroundColor:'rgba(0,112,192,0.25)',borderColor:'#0070C0',borderWidth:2},{label:'แก้แล้ว',data:lb.map(k=>w[k].r),backgroundColor:'#22c55e40',borderColor:'#22c55e',borderWidth:2}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:'#888',boxWidth:12}}},scales:{x:{ticks:{color:'#666'},grid:{color:'#e2e4ed'}},y:{ticks:{color:'#666'},grid:{color:'#e2e4ed'}}}}});
}

function mgr_sla(d) {
  const now=new Date(),br=d.filter(t=>t.SLA_Deadline&&new Date(t.SLA_Deadline)<now&&t.Status!=='แก้ไขแล้ว').length;
  if(mgr_sc)mgr_sc.destroy();
  mgr_sc=new Chart(document.getElementById('mgr-sla-chart'),{type:'doughnut',data:{labels:['On Track','Breached'],datasets:[{data:[d.length-br,br],backgroundColor:['#22c55e80','#f8717180'],borderColor:['#22c55e','#f87171'],borderWidth:2}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:'#888'}}}}});
}

function mgr_prio(d) {
  const p={Critical:0,High:0,Medium:0,Low:0};d.forEach(t=>{if(p[t.Priority]!==undefined)p[t.Priority]++;});
  if(mgr_pc)mgr_pc.destroy();
  mgr_pc=new Chart(document.getElementById('mgr-prio-chart'),{type:'doughnut',data:{labels:Object.keys(p),datasets:[{data:Object.values(p),backgroundColor:['#f8717180','#fb923c80','#60a5fa80','#4ade8080'],borderColor:['#f87171','#fb923c','#60a5fa','#4ade80'],borderWidth:2}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:'#888'}}}}});
}

function mgr_dept(d) {
  const depts={};d.forEach(t=>{depts[t.Department]=(depts[t.Department]||0)+1;});
  const s=Object.entries(depts).sort((a,b)=>b[1]-a[1]).slice(0,8),max=s[0]?.[1]||1;
  document.getElementById('mgr-dept').innerHTML=s.map(([n,c])=>`<div class="dept-row"><div class="dept-name">${n||'Unknown'}</div><div class="dept-bar-wrap"><div class="dept-bar" style="width:${Math.round(c/max*100)}%"></div></div><div class="dept-count">${c}</div></div>`).join('')||'<div style="text-align:center;padding:30px;color:#555">ไม่มีข้อมูล</div>';
}

function mgr_breach(d) {
  const now=new Date(),br=d.filter(t=>t.SLA_Deadline&&new Date(t.SLA_Deadline)<now&&t.Status!=='แก้ไขแล้ว');
  const pb=p=>{const m={Critical:'b-crit',High:'b-high',Medium:'b-med',Low:'b-low'};return `<span class="badge ${m[p]||'b-med'}">${p||'—'}</span>`;};
  document.getElementById('mgr-breach-tbl').innerHTML=!br.length
    ?'<tr><td colspan="6" style="text-align:center;padding:30px;color:#22c55e">✅ ไม่มี Ticket เกิน SLA!</td></tr>'
    :br.map(t=>{const dl=new Date(t.SLA_Deadline);return `<tr><td style="color:#0070C0;font-size:12px;font-weight:600">${t.TicketID||'#'+t.ID}</td><td>${t.Title}</td><td>${t.Department||'—'}</td><td>${pb(t.Priority)}</td><td>${dl.toLocaleString('th-TH')}</td><td class="overdue">+${Math.round((now-dl)/3600000)} ชม.</td></tr>`;}).join('');
}

function mgr_export() {
  const months=parseInt(document.getElementById('mgr-month').value),cutoff=new Date();
  if(months===0)cutoff.setDate(1);else cutoff.setMonth(cutoff.getMonth()-months);
  const d=mgr_data.filter(t=>new Date(t.Created)>=cutoff);
  const h=['TicketID','Title','Department','Priority','Status','SLA_Deadline','Created'];
  const a=document.createElement('a');a.href=URL.createObjectURL(new Blob(['﻿'+[h.join(','),...d.map(t=>h.map(k=>`"${(t[k]||'').toString().replace(/"/g,'""')}"`).join(','))].join('\n')],{type:'text/csv;charset=utf-8;'}));a.download='it_report.csv';a.click();
}

// ── OPS TAB ──────────────────────────────────────────────────
function ops_ago(d){if(!d)return '';const m=Math.floor((Date.now()-new Date(d))/60000);if(m<1)return 'เมื่อกี้';if(m<60)return m+' นาทีที่แล้ว';const h=Math.floor(m/60);return h<24?h+' ชม.ที่แล้ว':Math.floor(h/24)+' วันที่แล้ว';}
function ops_sev(t){const s=(t||'').toLowerCase();return s.includes('fail')||s.includes('error')?'err':s.includes('warn')?'warn':s.includes('success')?'ok':'info';}
function ops_dot(s){const c={ok:'#22c55e',err:'#ef4444',warn:'#f59e0b',info:'#3b82f6'};return `<div class="ops-sev" style="background:${c[s]||c.info}"></div>`;}
function ops_tag(s){const l={ok:'Success',err:'Failed',warn:'Warning',info:'Info'};return `<span class="ops-tag ops-tag-${s}">${l[s]||s}</span>`;}

async function ops_loadVeem() {
  const filters=document.getElementById('veem-filter').value.split(',').map(s=>s.trim()).filter(Boolean);
  const list=document.getElementById('veem-list');
  try {
    const token = await getToken(GRAPH_SCOPES);
    // ค้นหาแบบ OR ทุก keyword ใน subject — ค้นข้าม folder อัตโนมัติ
    const searchQ = filters.join(' OR ');
    const r=await fetch(
      `https://graph.microsoft.com/v1.0/me/messages?$search=${encodeURIComponent('"'+searchQ+'"')}&$top=20&$select=subject,receivedDateTime,from,bodyPreview&$orderby=receivedDateTime+desc`,
      { headers: { 'Authorization': 'Bearer ' + token } }
    );
    if(!r.ok) { const e=await r.json(); throw new Error(e.error?.message||'HTTP '+r.status); }
    ops_renderVeem((await r.json()).value||[]);
  } catch(e){
    ops_renderVeem([
      {subject:'[Success] BK_FIT_PAYROLL_to_FitWasabi_Repo',receivedDateTime:new Date(Date.now()-3*3600000).toISOString(),from:{emailAddress:{address:'veeam@furuya.co.th'}},bodyPreview:'Duration: 00:05:23'},
      {subject:'[Failed] Backup Configuration Job',receivedDateTime:new Date(Date.now()-6*3600000).toISOString(),from:{emailAddress:{address:'veeam@furuya.co.th'}},bodyPreview:'Error: Connection timeout'}
    ]);
    document.getElementById('ops-status').innerHTML='⚠️ Demo data — <button onclick="ops_connectEmail()" style="background:none;border:none;color:#0070C0;font-weight:600;cursor:pointer;padding:0;font-size:inherit;text-decoration:underline;">คลิกเพื่อเชื่อมต่อ Email</button> (ครั้งแรกเท่านั้น)';
  }
}

function ops_renderVeem(emails){
  const l=document.getElementById('veem-list'),c=document.getElementById('veem-count');
  if(!emails.length){l.innerHTML='<div style="text-align:center;padding:40px;color:#555;">ไม่พบอีเมล Veeam</div>';c.textContent='0';ops_kpi('veem',[]);return;}
  c.textContent=emails.length;
  l.innerHTML=emails.map(e=>{const s=ops_sev(e.subject);return `<div class="ops-item">${ops_dot(s)}<div class="ops-body"><div class="ops-subj" title="${e.subject}">${e.subject}</div><div class="ops-meta">${e.from?.emailAddress?.address||''} · ${ops_ago(e.receivedDateTime)}</div></div>${ops_tag(s)}</div>`;}).join('');
  ops_kpi('veem',emails);
}

let _opmActiveSection = null;
async function ops_section(name, el) {
  const body = document.getElementById('opm-section-body');
  // toggle off ถ้ากดซ้ำ
  if (_opmActiveSection === name) {
    _opmActiveSection = null;
    body.style.display = 'none';
    document.querySelectorAll('.opm-nav-btn').forEach(b => b.classList.remove('active'));
    return;
  }
  _opmActiveSection = name;
  document.querySelectorAll('.opm-nav-btn').forEach(b => b.classList.remove('active'));
  if (el) el.classList.add('active');
  body.style.display = '';
  body.innerHTML = '<div style="text-align:center;padding:30px"><div class="spinner"></div></div>';

  const pathMap = {
    dashboard: ['/apiclient/api/json/alarm','/api/json/alarms','/api/json/alarm'],
    servers:   ['/apiclient/api/json/device','/api/json/device','/api/json/resources'],
    alarms:    ['/apiclient/api/v2/alarm','/api/json/alarms','/api/json/alarm'],
    network:   ['/apiclient/api/json/device','/api/json/device','/api/json/resources'],
  };
  const raw = await ops_fetchOpm(pathMap[name] || pathMap.alarms);
  if (!raw) {
    body.innerHTML = `<div style="text-align:center;padding:30px;font-size:13px;color:#555;">
      เชื่อมต่อ OpManager ไม่ได้<br>
      <small style="color:#888;margin-top:8px;display:block;">ต้องเปิด CORS ใน OpManager ก่อน<br>
      Settings → General → API → CORS Origins → เพิ่ม <b>https://nutthawut-a11y.github.io</b></small></div>`;
    return;
  }
  let items = Array.isArray(raw) ? raw : (raw.response?.result || raw.data || raw.devices || raw.alarms || raw.alarm || []);

  if (name === 'dashboard') {
    const crit = items.filter(a => (a.severity||'').toLowerCase().includes('crit')).length;
    const warn = items.filter(a => (a.severity||'').toLowerCase().includes('warn')).length;
    const ok   = items.filter(a => (a.severity||'').toLowerCase().includes('clear')).length;
    body.innerHTML = `<div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(130px,1fr));gap:12px;padding:6px;">
      <div style="background:#f0f5f0;border-radius:8px;padding:16px;text-align:center;"><div style="font-size:28px;font-weight:700;color:#0070C0">${items.length}</div><div style="font-size:12px;color:#666;margin-top:4px;">Total Alarms</div></div>
      <div style="background:#f0f5f0;border-radius:8px;padding:16px;text-align:center;"><div style="font-size:28px;font-weight:700;color:#ef4444">${crit}</div><div style="font-size:12px;color:#666;margin-top:4px;">Critical</div></div>
      <div style="background:#f0f5f0;border-radius:8px;padding:16px;text-align:center;"><div style="font-size:28px;font-weight:700;color:#f59e0b">${warn}</div><div style="font-size:12px;color:#666;margin-top:4px;">Warning</div></div>
      <div style="background:#f0f5f0;border-radius:8px;padding:16px;text-align:center;"><div style="font-size:28px;font-weight:700;color:#22c55e">${ok}</div><div style="font-size:12px;color:#666;margin-top:4px;">Cleared</div></div>
    </div>`;
  } else if (name === 'alarms') {
    if (!items.length) { body.innerHTML = '<div style="text-align:center;padding:30px;color:#22c55e;">✅ ไม่มี Alarm</div>'; return; }
    body.innerHTML = items.slice(0,50).map(a => {
      const n = a.device||a.displayName||a.entityName||'Unknown';
      const msg = a.message||a.alarmMessage||a.description||'';
      const sv = (a.severity||'').toLowerCase();
      const s = sv.includes('crit')||sv.includes('error')?'err':sv.includes('warn')?'warn':sv.includes('clear')?'ok':'info';
      return `<div class="ops-item">${ops_dot(s)}<div class="ops-body"><div class="ops-subj">${n}</div><div class="ops-meta">${msg}${msg&&ops_ago(a.lastUpdatedTime||a.createdTime||'')?' · ':''}${ops_ago(a.lastUpdatedTime||a.createdTime||'')}</div></div>${ops_tag(s)}</div>`;
    }).join('');
  } else {
    if (!items.length) { body.innerHTML = '<div style="text-align:center;padding:30px;color:#22c55e;">✅ ไม่มี Device</div>'; return; }
    body.innerHTML = items.slice(0,50).map(d => {
      const nm = d.deviceName||d.displayName||d.entityName||'Unknown';
      const ip = d.ipAddress||d.ip||'';
      const tp = d.type||d.deviceType||'';
      const st = (d.status||d.deviceStatus||'').toString().toUpperCase();
      const s  = (st==='DOWN'||st==='1'||st==='CRITICAL')?'err':(st==='UP'||st==='0'||st==='ACTIVE')?'ok':'warn';
      return `<div class="ops-item">${ops_dot(s)}<div class="ops-body"><div class="ops-subj">${nm}${ip?' <span style="color:#888;font-size:11px;">('+ip+')</span>':''}</div><div class="ops-meta">${tp}${tp?' · ':''}${ops_ago(d.lastPollTime||d.lastUpdatedTime||'')}</div></div>${ops_tag(s)}</div>`;
    }).join('');
  }
}

async function ops_fetchOpm(paths) {
  const base = document.getElementById('opm-url').value.trim().replace(/\/$/,'');
  const key  = document.getElementById('opm-key').value.trim();
  // OpManager REST API: query param คือ AUTHTOKEN= (ไม่ใช่ apiKey=)
  // simple request (ไม่มี custom header) → ไม่มี preflight → CORS ผ่านได้แม้ web.xml ยังไม่สมบูรณ์
  for (const path of paths) {
    for (const [url, hdrs] of [
      [`${base}${path}?AUTHTOKEN=${key}`, {}],
      [`${base}${path}`,                  { 'Accept':'application/json', ...(key?{'AUTHTOKEN':key}:{}) }]
    ]) {
      try {
        const r = await fetch(url, { mode:'cors', headers: hdrs });
        if (r.ok) return await r.json();
      } catch(e) {}
    }
  }
  return null;
}

async function ops_loadOpm(){
  const l=document.getElementById('opm-list'),c=document.getElementById('opm-count');
  const raw = await ops_fetchOpm(['/apiclient/api/json/alarm','/api/json/alarms','/api/json/alarm']);
  let items = raw ? (Array.isArray(raw) ? raw : (raw.data||raw.alarms||raw.alarm||[])) : null;
  if(!items){l.innerHTML='<div style="text-align:center;padding:30px;color:#555;font-size:13px;">เชื่อมต่อ OpManager ไม่ได้<br><small style="color:#888;margin-top:8px;display:block;">ตรวจสอบ URL และ API Key<br>OpManager → Settings → General → API</small></div>';c.textContent='–';ops_kpi('opm',[]);return;}
  c.textContent=items.length;
  l.innerHTML=!items.length?'<div style="text-align:center;padding:30px;color:#22c55e;">✅ ไม่มี Alert</div>':items.slice(0,30).map(a=>{const n=a.device||a.displayName||a.entityName||'Unknown',msg=a.message||a.alarmMessage||a.description||'',sv=(a.severity||'').toLowerCase(),s=sv.includes('crit')||sv.includes('error')?'err':sv.includes('warn')?'warn':sv.includes('clear')?'ok':'info';return `<div class="ops-item">${ops_dot(s)}<div class="ops-body"><div class="ops-subj">${n}</div><div class="ops-meta">${msg} · ${ops_ago(a.lastUpdatedTime||a.createdTime||'')}</div></div>${ops_tag(s)}</div>`;}).join('');
  ops_kpi('opm',items);
}


function ops_kpi(src,items){
  if(src==='veem'){document.getElementById('ops-total').textContent=items.length;document.getElementById('ops-ok').textContent=items.filter(e=>ops_sev(e.subject)==='ok').length;document.getElementById('ops-fail').textContent=items.filter(e=>ops_sev(e.subject)==='err').length;}
  if(src==='opm')document.getElementById('ops-alerts').textContent=items.length;
}

async function ops_loadAll(){
  document.getElementById('ops-status').textContent='⏳ กำลังโหลด...';
  await Promise.all([ops_loadVeem(),ops_loadOpm()]);
  document.getElementById('ops-updated').textContent='อัปเดต: '+new Date().toLocaleTimeString('th-TH');
}

setInterval(()=>{if(document.getElementById('tab-ops').classList.contains('active'))ops_loadAll();},5*60*1000);

// Init
setRole('employee');
