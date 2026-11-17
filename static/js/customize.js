// Page 3: read the form -> send to server -> show the returned preview. Repeat on every change.
const D = JSON.parse(document.getElementById('coa-data').textContent);
let G = D.settings;          // shared settings (all COAs)
let O = D.overrides || {};   // per-COA settings, only for samples ticked 'RM & WS'
let C = D.categories || {};  // category of each sample
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
  s.unit = $('#unit').value;
  s.ar_prefix = $('#arprefix').value;
  $$('.sig').forEach(i => { s.signatories[i.dataset.sig] = i.value; });
  return s;
}

// ---- Put a settings object back into the form (used when you move to another sample) ----
function applyToForm(s) {
  $$('.frow[data-field]').forEach(row => {
    const f = (s.fields || {})[row.dataset.field];
    if (!f) return;
    const useBox = $('.f-use', row);
    if (useBox) useBox.checked = !!f.use_file;
    $('.f-manual', row).value = f.manual || '';
  });
  $$('.prow').forEach(row => {
    const p = (s.params || {})[row.dataset.param];
    if (!p) return;
    $('.p-include', row).checked = !!p.include; $('.p-spec', row).value = p.spec;
    $('.p-use', row).checked = !!p.use_file; $('.p-manual', row).value = p.manual || '';
  });
  $('#others').innerHTML = '';
  (s.others || []).forEach(addOther);
  $('#remarks').value = s.remarks; $('#unit').value = s.unit || 'auto'; $('#arprefix').value = s.ar_prefix;
  $$('.sig').forEach(i => { i.value = (s.signatories || {})[i.dataset.sig] || ''; });
}
const cur = () => D.items[pos];
function payload() {
  return {id: cur().id, settings: G, overrides: O, categories: C};
}
// Where do the form's values go? Into this sample's own copy when 'RM & WS' is ticked, else into the shared settings.
function store() {
  const s = collect();
  if (O[cur().id]) O[cur().id] = s; else G = s;
}
function showSample() {
  const it = cur(), own = !!O[it.id];
  $('#sname').textContent = `${it.sr} · ${it.product} · ${it.batch}` + (it.sku ? ` · ${it.sku}` : '');
  $('#cat').value = C[it.id] || 'FG';
  $('#own').checked = own;
  $('#scope').innerHTML = own ? 'Changes below apply to <b>this COA only</b>.' : 'Changes below apply to <b>all COAs</b>.';
  $('#samplecard').classList.toggle('own', own);
  applyToForm(own ? O[it.id] : G);
  syncRows();
}
$('#own').onchange = e => {
  const id = cur().id;
  if (e.target.checked) O[id] = JSON.parse(JSON.stringify(G)); else delete O[id];
  showSample(); schedule();
};
$('#cat').onchange = e => { C[cur().id] = e.target.value; schedule(); };

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
    body: JSON.stringify(payload())});
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
function schedule() { store(); syncRows(); clearTimeout(timer); timer = setTimeout(refresh, 120); }

$('#panel').addEventListener('input', schedule);
$('#panel').addEventListener('change', schedule);

// ---- Pager ------------------------------------------------------------------------
function go(n) {
  store();
  pos = (n + D.items.length) % D.items.length;
  $('#jump').value = pos;
  showSample();
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
  store();
  $('#overlay').hidden = false;
  try {
    const res = await fetch(D.generateUrl, {method: 'POST', headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({settings: G, overrides: O, categories: C})});
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Something went wrong');
    location.href = data.redirect;
  } catch (err) {
    $('#overlay').hidden = true;
    alert(err.message);
  }
};

showSample();
refresh();
