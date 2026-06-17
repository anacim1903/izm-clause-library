'use strict';

const SHEET_URL =
  'https://docs.google.com/spreadsheets/d/' +
  '1ebPCHEkZz5sDRTrobt_NPmXGohFhKVjzQUpOMZCkegI' +
  '/gviz/tq?tqx=out:csv';

let allClauses = [];

Office.onReady(async () => {
  await loadClauses();
});

async function loadClauses() {
  showStatus('Loading clauses…');
  try {
    const res = await fetch(SHEET_URL);
    if (!res.ok) throw new Error(`Could not load sheet (HTTP ${res.status})`);
    const text = await res.text();

    const { data } = Papa.parse(text, { skipEmptyLines: true });

    if (!data || data.length < 2) {
      showStatus('No clauses found. Add rows to the Google Sheet and reload.');
      return;
    }

    allClauses = data.slice(1)
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
