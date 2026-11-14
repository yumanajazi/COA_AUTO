// Page 3: read the form -> send to server -> show the returned preview. Repeat on every change.
const D = window.COA;
let pos = 0;               // which selected sample is on screen
let timer = null;
let seq = 0;               // ignore out-of-order responses

const $ = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => Array.from(r.querySelectorAll(s));

// ---- "Others": rows typed by hand -------------------------------------------------
function addOther(o = {label: '', spec: 'Should be absent', obs: 'Absent'}) {
  const div = document.createElement('div');
  div.className = 'orow';
  div.innerHTML = `<input class="o-label" placeholder="Test name (e.g. Klebsiella)">
                   <input class="o-spec" placeholder="Specification">
                   <input class="o-obs" placeholder="Result">
                   <button type="button" class="x" aria-label="Remove">&times;</button>`;
  $('.o-label', div).value = o.label; $('.o-spec', div).value = o.spec; $('.o-obs', div).value = o.obs;
  $('.x', div).onclick = () => { div.remove(); schedule(); };
  $('#others').appendChild(div);
}
(D.settings.others || []).forEach(addOther);
$('#addother').onclick = () => { addOther(); schedule(); };

// ---- Read the whole form into one settings object ---------------------------------
function collect() {
  const s = {fields: {}, params: {}, others: [], signatories: {}};
  $$('.frow[data-field]').forEach(row => {
    const useBox = $('.f-use', row);
    s.fields[row.dataset.field] = {use_file: useBox ? useBox.checked : false, manual: $('.f-manual', row).value};
  });
  $$('.prow').forEach(row => {
    s.params[row.dataset.param] = {
      include: $('.p-include', row).checked, spec: $('.p-spec', row).value,
      use_file: $('.p-use', row).checked, manual: $('.p-manual', row).value};
  });
  $$('.orow').forEach(r => s.others.push({label: $('.o-label', r).value, spec: $('.o-spec', r).value, obs: $('.o-obs', r).value}));
  s.remarks = $('#remarks').value;
  s.ar_prefix = $('#arprefix').value;
  $$('.sig').forEach(i => { s.signatories[i.dataset.sig] = i.value; });
  return s;
}

// ---- Show/hide the manual boxes depending on the ticks ----------------------------
function syncRows() {
  $$('.frow[data-field][data-src="1"]').forEach(row => {
    const on = $('.f-use', row).checked;
    $('.f-manual', row).hidden = on;
    $('.srcval', row).hidden = !on;
  });
  $$('.prow').forEach(row => {
    const inc = $('.p-include', row).checked, use = $('.p-use', row).checked;
    row.classList.toggle('off', !inc);
    $('.p-manual', row).hidden = use;
  });
}

// ---- Preview ----------------------------------------------------------------------
function fitPreview() {
  const box = $('#previewbox');
  const z = Math.min(1, (box.clientWidth - 2) / 794);
  $('#preview').style.zoom = z;
}
async function refresh() {
  const mine = ++seq;
  const item = D.items[pos];
  const res = await fetch(D.previewUrl, {method: 'POST', headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({id: item.id, settings: collect()})});
  if (mine !== seq) return;
  if (!res.ok) { $('#preview').textContent = 'Could not build the preview.'; return; }
  const data = await res.json();
  $('#preview').innerHTML = data.html;
  fitPreview();
  const w = $('#warn');
  w.hidden = !data.warnings.length;
  w.textContent = data.warnings.length ? 'Check this sample: ' + data.warnings.join(' · ') : '';
  $$('.frow[data-field]').forEach(row => {
    const el = $('.srcval', row);
    if (el && data.source[row.dataset.field] !== undefined) el.textContent = '→ ' + data.source[row.dataset.field];
  });
}
function schedule() { syncRows(); clearTimeout(timer); timer = setTimeout(refresh, 120); }

$('#panel').addEventListener('input', schedule);
$('#panel').addEventListener('change', schedule);

// ---- Pager ------------------------------------------------------------------------
function go(n) {
  pos = (n + D.items.length) % D.items.length;
  $('#jump').value = pos;
  refresh();
}
$('#prev').onclick = () => go(pos - 1);
$('#next').onclick = () => go(pos + 1);
$('#jump').onchange = e => go(+e.target.value);
window.addEventListener('resize', fitPreview);
document.addEventListener('keydown', e => {
  if (['INPUT', 'SELECT', 'TEXTAREA'].includes(e.target.tagName)) return;
  if (e.key === 'ArrowLeft') go(pos - 1);
  if (e.key === 'ArrowRight') go(pos + 1);
});

// ---- Process final COA ------------------------------------------------------------
$('#process').onclick = async () => {
  $('#overlay').hidden = false;
  try {
    const res = await fetch(D.generateUrl, {method: 'POST', headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({settings: collect()})});
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Something went wrong');
    location.href = data.redirect;
  } catch (err) {
    $('#overlay').hidden = true;
    alert(err.message);
  }
};

syncRows();
go(0);
