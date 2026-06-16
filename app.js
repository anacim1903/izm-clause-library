'use strict';

// ── Configuration ─────────────────────────────────────────────────────────────

const CLIENT_ID = '44d73efb-0980-447f-8629-04c970ac0e74';
const TENANT_ID = 'd322a5f5-4eaf-439f-a75a-d4c3822080db';
const SCOPES    = ['Files.Read', 'Sites.Read.All'];

const WORKBOOK_BASE =
  'https://graph.microsoft.com/v1.0' +
  '/sites/izumedeirosadvogados.sharepoint.com:/sites/IzuMedeiros:' +
  '/drive/root:/AI/260616%20-%20IZM%20Clauses.xlsx:';

let msalInstance;
let allClauses = [];

Office.onReady(() => {
  msalInstance = new msal.PublicClientApplication({
    auth: {
      clientId:    CLIENT_ID,
      authority:   `https://login.microsoftonline.com/${TENANT_ID}`,
      redirectUri: window.location.href.split('?')[0].split('#')[0],
    },
    cache: { cacheLocation: 'sessionStorage', storeAuthStateInCookie: false },
  });

  msalInstance.handleRedirectPromise().then(init).catch(showError);
});

async function init() {
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length > 0) {
    await loadClauses();
  } else {
    showStatus('Signing in…');
    try {
      await msalInstance.loginPopup({ scopes: SCOPES });
      await loadClauses();
    } catch (e) {
      showError(e);
    }
  }
}

async function getToken() {
  const account = msalInstance.getAllAccounts()[0];
  try {
    const r = await msalInstance.acquireTokenSilent({ scopes: SCOPES, account });
    return r.accessToken;
  } catch {
    const r = await msalInstance.acquireTokenPopup({ scopes: SCOPES, account });
    return r.accessToken;
  }
}

async function loadClauses() {
  showStatus('Loading clauses…');
  try {
    const token   = await getToken();
    const headers = { Authorization: `Bearer ${token}` };

    const wsRes = await fetch(`${WORKBOOK_BASE}/workbook/worksheets`, { headers });
    if (!wsRes.ok) throw new Error(`Could not open workbook (HTTP ${wsRes.status}). Check the file path and permissions.`);
    const { value: sheets } = await wsRes.json();
    if (!sheets || !sheets.length) throw new Error('No worksheets found in the Excel file.');
    const sheetName = sheets[0].name;

    const dataRes = await fetch(
      `${WORKBOOK_BASE}/workbook/worksheets('${encodeURIComponent(sheetName)}')/usedRange`,
      { headers }
    );
    if (!dataRes.ok) throw new Error(`Could not read sheet data (HTTP ${dataRes.status}).`);
    const { values } = await dataRes.json();

    if (!values || values.length < 2) {
      showStatus('No clauses found. Add rows to the Excel file and reload.');
      return;
    }

    allClauses = values
      .slice(1)
      .filter(r => r[0])
      .map(r => ({
        title: String(r[0] ?? '').trim(),
        type:  String(r[1] ?? '').trim(),
        text:  String(r[2] ?? '').trim(),
      }));

    populateFilter();
    renderClauses(allClauses);
    document.getElementById('searchArea').style.display = 'flex';
    document.getElementById('status').style.display     = 'none';

  } catch (e) {
    showError(e);
  }
}

function populateFilter() {
  const types  = [...new Set(allClauses.map(c => c.type).filter(Boolean))].sort();
  const select = document.getElementById('typeFilter');
  while (select.options.length > 1) select.remove(1);
  types.forEach(t => {
    const o = document.createElement('option');
    o.value = t;
    o.textContent = t;
    select.appendChild(o);
  });
}

function filterClauses() {
  const q    = document.getElementById('searchInput').value.toLowerCase();
  const type = document.getElementById('typeFilter').value;
  renderClauses(allClauses.filter(c => {
    const matchQ = !q || c.title.toLowerCase().includes(q) || c.text.toLowerCase().includes(q);
    const matchT = !type || c.type === type;
    return matchQ && matchT;
  }));
}

function renderClauses(clauses) {
  const list  = document.getElementById('clauseList');
  const count = document.getElementById('count');

  count.textContent = `${clauses.length} clause${clauses.length !== 1 ? 's' : ''}`;
  list.innerHTML = '';

  if (!clauses.length) {
    list.innerHTML = '<p class="empty">No clauses match your search.</p>';
    return;
  }

  clauses.forEach(c => {
    const card = document.createElement('div');
    card.className = 'clause-card';
    card.innerHTML = `
      <div class="clause-title">${esc(c.title)}</div>
      ${c.type ? `<div class="clause-type">${esc(c.type)}</div>` : ''}
      <div class="clause-preview">${esc(c.text.slice(0, 140))}${c.text.length > 140 ? '…' : ''}</div>
    `;
    card.addEventListener('click', () => insertClause(c));
    list.appendChild(card);
  });
}

async function insertClause(c) {
  try {
    await Word.run(async ctx => {
      ctx.document.getSelection().insertText(c.text, Word.InsertLocation.replace);
      await ctx.sync();
    });
  } catch (e) {
    alert('Could not insert clause:\n' + e.message);
  }
}

function showStatus(msg) {
  const s = document.getElementById('status');
  s.textContent       = msg;
  s.style.display     = 'block';
  s.style.color       = '#666';
  document.getElementById('searchArea').style.display = 'none';
  document.getElementById('clauseList').innerHTML      = '';
}

function showError(err) {
  const s = document.getElementById('status');
  s.textContent   = '⚠ ' + (err?.message || String(err));
  s.style.display = 'block';
  s.style.color   = '#c0392b';
}

function esc(str) {
  const d = document.createElement('div');
  d.appendChild(document.createTextNode(str));
  return d.innerHTML;
}
