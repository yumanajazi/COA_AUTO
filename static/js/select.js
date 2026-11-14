// Page 2: search / filter the uploaded rows, tick the ones to convert.
const rows = Array.from(document.querySelectorAll('#rows tbody tr'));
const q = document.getElementById('q');
const onlyReady = document.getElementById('onlyready');
const dfrom = document.getElementById('dfrom');
const dto = document.getElementById('dto');
const count = document.getElementById('count');
const go = document.getElementById('go');

function visible(tr) {
  if (onlyReady.checked && tr.dataset.ready !== '1') return false;
  if (q.value && !tr.dataset.text.includes(q.value.trim().toLowerCase())) return false;
  const d = tr.dataset.date;           // ISO date, so plain string comparison works
  if (dfrom.value && (!d || d < dfrom.value)) return false;
  if (dto.value && (!d || d > dto.value)) return false;
  return true;
}
function applyFilter() {
  rows.forEach(tr => { tr.hidden = !visible(tr); });
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
