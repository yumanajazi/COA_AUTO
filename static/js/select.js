// Page 2: search / filter the uploaded rows, tick the ones to convert.
const rows = Array.from(document.querySelectorAll('#rows tbody tr'));
const q = document.getElementById('q');
const onlyReady = document.getElementById('onlyready');
const dfrom = document.getElementById('dfrom');
const dto = document.getElementById('dto');
const count = document.getElementById('count');
const go = document.getElementById('go');

const hits = document.getElementById('hits');
const esc = t => t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
// Wrap every match of the search text in <mark> inside the product / batch / A.R. cells.
function highlight(tr, term) {
  tr.querySelectorAll('.hl-cell').forEach(td => {
    if (td.dataset.orig === undefined) td.dataset.orig = td.textContent;
    const text = td.dataset.orig;
    if (!term) { td.textContent = text; return; }
    const low = text.toLowerCase();
    let out = '', i = 0, k;
    while ((k = low.indexOf(term, i)) !== -1) {
      out += esc(text.slice(i, k)) + '<mark>' + esc(text.slice(k, k + term.length)) + '</mark>';
      i = k + term.length;
    }
    td.innerHTML = out + esc(text.slice(i));
  });
}
function visible(tr) {
  if (onlyReady.checked && tr.dataset.ready !== '1') return false;
  if (q.value && !tr.dataset.text.includes(q.value.trim().toLowerCase())) return false;
  const d = tr.dataset.date;           // ISO date, so plain string comparison works
  if (dfrom.value && (!d || d < dfrom.value)) return false;
  if (dto.value && (!d || d > dto.value)) return false;
  return true;
}
function applyFilter() {
  const term = q.value.trim().toLowerCase();
  let n = 0;
  rows.forEach(tr => {
    tr.hidden = !visible(tr);
    highlight(tr, tr.hidden ? '' : term);
    tr.classList.toggle('hit', !!term && !tr.hidden);
    if (term && !tr.hidden) n++;
  });
  hits.textContent = term ? n + ' matching record' + (n === 1 ? '' : 's') : '';
  updateCount();
}
function updateCount() {
  const n = rows.filter(tr => tr.querySelector('input').checked).length;
  count.textContent = n;
  go.disabled = n === 0;
}
[q, onlyReady, dfrom, dto].forEach(el => el.addEventListener('input', applyFilter));
document.getElementById('selall').addEventListener('click', () => {
  rows.forEach(tr => { if (!tr.hidden) tr.querySelector('input').checked = true; });
  updateCount();
});
document.getElementById('clearall').addEventListener('click', () => {
  rows.forEach(tr => { tr.querySelector('input').checked = false; });
  updateCount();
});
document.getElementById('rows').addEventListener('change', updateCount);
// Clicking anywhere on a row toggles it.
document.querySelector('#rows tbody').addEventListener('click', e => {
  if (e.target.tagName === 'INPUT') return;
  const cb = e.target.closest('tr').querySelector('input');
  cb.checked = !cb.checked; updateCount();
});
applyFilter();
