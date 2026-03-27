function applyTheme(mode) {
  const isDark = mode === 'dark';
  document.body.classList.toggle('dark-mode', isDark);
  const btn = document.getElementById('themeToggle');
  if (btn) btn.textContent = isDark ? '☀️ Modo claro' : '🌙 Modo oscuro';
}

function initThemeToggle() {
  const savedTheme = localStorage.getItem('themeMode') || 'light';
  applyTheme(savedTheme);
  const btn = document.getElementById('themeToggle');
  if (!btn) return;
  btn.addEventListener('click', () => {
    const next = document.body.classList.contains('dark-mode') ? 'light' : 'dark';
    localStorage.setItem('themeMode', next);
    applyTheme(next);
  });
}

const API_BASE = 'http://127.0.0.1:8000';
let allItems = [];
let currentPage = 1;
const PER_PAGE = 20;
let globalFilterOptions = null;
let activeSearchPrograma = '';
let uniqueProgramDenoms = [];

function setupMultiSelect(selectId){
  const select = document.getElementById(selectId);
  if(!select) return;

  select.classList.add('multi-hidden-select');

  let labelText = selectId;
  const prev = select.previousElementSibling;
  if(prev && prev.tagName === 'LABEL'){
    labelText = (prev.textContent || '').trim();
  }

  let wrapper = select.nextElementSibling;
  if(!wrapper || !wrapper.classList || !wrapper.classList.contains('multi-select')){
    wrapper = document.createElement('div');
    wrapper.className = 'multi-select mt-1';
    wrapper.innerHTML = `
      <button type="button" class="btn btn-outline-secondary btn-sm multi-select-toggle" data-target="${selectId}">
        <span class="me-2">${labelText}</span>
        <span class="multi-select-summary">(todos)</span>
      </button>
      <div class="multi-select-menu" data-target="${selectId}" style="display:none;"></div>
    `;
    select.parentNode.insertBefore(wrapper, select.nextSibling);

    const toggle = wrapper.querySelector('.multi-select-toggle');
    const menu = wrapper.querySelector('.multi-select-menu');
    if(toggle && menu){
      toggle.addEventListener('click', (ev)=>{
        ev.stopPropagation();
        const isOpen = menu.style.display === 'block';
        document.querySelectorAll('.multi-select-menu').forEach(m => { m.style.display = 'none'; });
        menu.style.display = isOpen ? 'none' : 'block';
      });
    }
  }

  const menu = wrapper.querySelector('.multi-select-menu');
  if(!menu) return;

  menu.innerHTML = '';
  const options = Array.from(select.options || []);
  options.forEach((opt, idx) => {
    const value = (opt.value || '').toString();
    if(value === '') return;
    const id = `${selectId}_opt_${idx}`;
    const row = document.createElement('div');
    row.className = 'multi-select-option';
    row.innerHTML = `
      <input type="checkbox" id="${id}">
      <label for="${id}" class="mb-0">${opt.text}</label>
    `;
    const cb = row.querySelector('input[type="checkbox"]');
    if(cb){
      cb.checked = opt.selected;
      cb.dataset.value = value;
      cb.addEventListener('change', ()=>{
        opt.selected = cb.checked;
        updateMultiSelectSummary(selectId);
        currentPage = 1;
        loadProgramas();
      });
    }
    menu.appendChild(row);
  });

  updateMultiSelectSummary(selectId);
}

function updateMultiSelectSummary(selectId){
  const select = document.getElementById(selectId);
  if(!select) return;
  const wrapper = select.nextElementSibling;
  if(!wrapper || !wrapper.classList || !wrapper.classList.contains('multi-select')) return;
  const summaryEl = wrapper.querySelector('.multi-select-summary');
  if(!summaryEl) return;

  const options = Array.from(select.options || []).filter(o => o.value !== '');
  const selected = options.filter(o => o.selected);
  if(!selected.length){
    summaryEl.textContent = '(todos)';
  }else if(selected.length === 1){
    summaryEl.textContent = selected[0].text;
  }else{
    summaryEl.textContent = `${selected.length} seleccionados`;
  }
}

document.addEventListener('DOMContentLoaded', async () => {
  initThemeToggle();
  initFileCounters();
  document.getElementById('uploadProgramasBtn').addEventListener('click', uploadProgramasExcel);
  document.getElementById('uploadCertificadosBtn').addEventListener('click', uploadCertificadosExcel);
  const historicosBtn = document.getElementById('uploadProgramasHistoricosBtn');
  if (historicosBtn) historicosBtn.addEventListener('click', uploadProgramasHistoricos);
  const comboBtn = document.getElementById('uploadProgramasCertificadosBtn');
  if (comboBtn) comboBtn.addEventListener('click', uploadProgramasYCertificados);
  const prevBtn = document.getElementById('prevPageBtn');
  const nextBtn = document.getElementById('nextPageBtn');
  if (prevBtn) prevBtn.addEventListener('click', () => changePage(-1));
  if (nextBtn) nextBtn.addEventListener('click', () => changePage(1));
  document.getElementById('applyFiltersBtn').addEventListener('click', loadProgramas);
  document.getElementById('clearFiltersBtn').addEventListener('click', () => {
    ['filterYear','filterMunicipio','filterEstrategia','filterConvenio','filterVigencia'].forEach(id => {
      const sel = document.getElementById(id);
      if(sel){
        Array.from(sel.options || []).forEach(o => { o.selected = false; });
        setupMultiSelect(id);
      }
    });
    const nf = document.getElementById('filterNumeroFicha');
    const sc = document.getElementById('filterSoloCertificados');
    if (nf) nf.value = '';
    if (sc) sc.checked = false;
    const sp = document.getElementById('searchPrograma');
    if (sp) sp.value = '';
    activeSearchPrograma = '';
    loadProgramas();
  });
  const exportBtn = document.getElementById('exportProgramasBtn');
  if (exportBtn) exportBtn.addEventListener('click', exportProgramasExcel);
  const searchInput = document.getElementById('searchPrograma');
  if (searchInput) {
    searchInput.addEventListener('input', () => {
      activeSearchPrograma = (searchInput.value || '').trim();
      updateProgramSearchSuggestions(activeSearchPrograma);
      currentPage = 1;
      loadProgramas();
    });
  }
  // Cerrar menús de multi-select al hacer clic fuera
  document.addEventListener('click', (ev)=>{
    const target = ev.target;
    if(!target.closest || !target.closest('.multi-select')){
      document.querySelectorAll('.multi-select-menu').forEach(m => { m.style.display = 'none'; });
    }
  });
  await loadGlobalFilterOptions();
  loadProgramas();
});

function buildUniqueProgramDenoms() {
  const s = new Set();
  allItems.forEach(r => {
    const d = (r.denominacion_programa || '').toString().trim();
    if (d) s.add(d);
  });
  uniqueProgramDenoms = Array.from(s).sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' }));
}

function updateProgramSearchSuggestions(prefix) {
  const list = document.getElementById('searchProgramaSuggestions');
  if (!list) return;
  const p = (prefix || '').toString().toLowerCase();
  const matches = p
    ? uniqueProgramDenoms.filter(d => d.toLowerCase().includes(p)).slice(0, 25)
    : uniqueProgramDenoms.slice(0, 25);
  list.innerHTML = '';
  matches.forEach(m => {
    const opt = document.createElement('option');
    opt.value = m;
    list.appendChild(opt);
  });
}

async function loadGlobalFilterOptions() {
  try {
    const resp = await fetch(`${API_BASE}/programas/filters`);
    if (!resp.ok) throw new Error(resp.status + ' ' + resp.statusText);
    globalFilterOptions = await resp.json();
  } catch (e) {
    console.error('Error cargando filtros globales de programas:', e);
    globalFilterOptions = null;
  }
}

function setStatus(msg) {
  const el = document.getElementById('status');
  if (el) el.textContent = msg || '';
}

function setUploadProgress(percent) {
  const container = document.getElementById('uploadProgressContainer');
  const bar = document.getElementById('uploadProgressBar');
  if (!container || !bar) return;
  if (!percent || percent <= 0) {
    container.style.display = 'none';
    bar.style.width = '0%';
    bar.setAttribute('aria-valuenow', '0');
    return;
  }
  const p = Math.max(0, Math.min(100, percent));
  container.style.display = 'block';
  bar.style.width = p + '%';
  bar.setAttribute('aria-valuenow', String(p));
}

function initFileCounters() {
  const inputs = document.querySelectorAll('input[type="file"][data-counter-target]');
  inputs.forEach((input) => {
    const spanId = input.getAttribute('data-counter-target');
    const span = spanId ? document.getElementById(spanId) : null;
    if (!span) return;
    const update = () => {
      const count = input.files ? input.files.length : 0;
      span.textContent = count ? `${count} archivo${count > 1 ? 's' : ''}` : '';
    };
    input.addEventListener('change', update);
    update();
  });
}

function escapeHtml(v) {
  if (v === null || v === undefined) return '';
  return String(v).replace(/[&<>\"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}

function toLabel(key) {
  if (key === 'fecha_corte') return 'Fecha de corte PE-04';
  if (key === 'cupos') return 'Aprendices matriculados';
  return String(key || '')
    .replace(/_/g, ' ')
    .split(' ')
    .map((w) => w ? w.charAt(0).toUpperCase() + w.slice(1) : '')
    .join(' ');
}

// Helpers numéricos similares al módulo de fichas
function toNumberProgramas(v) {
  if (v === null || v === undefined) return 0;
  if (typeof v === 'number') return Number.isFinite(v) ? v : 0;
  const cleaned = String(v).trim().replace(/\./g, '').replace(',', '.');
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : 0;
}

function formatNumberProgramas(n) {
  return new Intl.NumberFormat('es-CO').format(n || 0);
}

function getFilters() {
  const getSelectedValues = (sel) => {
    if(!sel) return [];
    return Array.from(sel.selectedOptions || [])
      .map(o => (o.value || '').toString().trim())
      .filter(v => v !== '');
  };

  const years = getSelectedValues(document.getElementById('filterYear'));
  const municipios = getSelectedValues(document.getElementById('filterMunicipio'));
  const estrategias = getSelectedValues(document.getElementById('filterEstrategia'));
  const convenios = getSelectedValues(document.getElementById('filterConvenio'));
  const vigencias = getSelectedValues(document.getElementById('filterVigencia'));
  const soloCertificados = !!document.getElementById('filterSoloCertificados')?.checked;
  const numeroFicha = (document.getElementById('filterNumeroFicha')?.value || '').trim();
  return { years, municipios, estrategias, convenios, vigencias, soloCertificados, numeroFicha };
}

function buildUrl() {
  const { years, municipios, estrategias, convenios, vigencias, soloCertificados, numeroFicha } = getFilters();
  const params = new URLSearchParams();
  if (years.length) params.set('year', years.join(','));
  if (municipios.length) params.set('municipio', municipios.join(','));
  if (estrategias.length) params.set('estrategia', estrategias.join(','));
  if (convenios.length) params.set('convenio', convenios.join(','));
  if (vigencias.length) params.set('vigencia', vigencias.join(','));
   if (numeroFicha) params.set('numero_ficha', numeroFicha);
  if (soloCertificados) params.set('solo_certificados', '1');
  if (activeSearchPrograma) params.set('search', activeSearchPrograma);
  params.set('page', String(currentPage || 1));
  params.set('per_page', String(PER_PAGE));
  const q = params.toString();
  return `${API_BASE}/programas${q ? `?${q}` : ''}`;
}

function getFilenameFromDisposition(contentDisposition) {
  if (!contentDisposition) return '';
  const match = contentDisposition.match(/filename="?([^";]+)"?/i);
  return match && match[1] ? match[1] : '';
}

function buildProgramasExportUrl() {
  const { years, municipios, estrategias, convenios, vigencias, soloCertificados, numeroFicha } = getFilters();
  const params = new URLSearchParams();
  if (years.length) params.set('year', years.join(','));
  if (municipios.length) params.set('municipio', municipios.join(','));
  if (estrategias.length) params.set('estrategia', estrategias.join(','));
  if (convenios.length) params.set('convenio', convenios.join(','));
  if (vigencias.length) params.set('vigencia', vigencias.join(','));
  if (soloCertificados) params.set('solo_certificados', '1');
  if (numeroFicha) params.set('numero_ficha', numeroFicha);
  if (activeSearchPrograma) params.set('search', activeSearchPrograma);
  const q = params.toString();
  return `${API_BASE}/programas/export${q ? `?${q}` : ''}`;
}

// URL para obtener TODOS los programas filtrados (sin paginación) para totales
function buildProgramasAllUrl() {
  const { years, municipios, estrategias, convenios, vigencias, soloCertificados, numeroFicha } = getFilters();
  const params = new URLSearchParams();
  if (years.length) params.set('year', years.join(','));
  if (municipios.length) params.set('municipio', municipios.join(','));
  if (estrategias.length) params.set('estrategia', estrategias.join(','));
  if (convenios.length) params.set('convenio', convenios.join(','));
  if (vigencias.length) params.set('vigencia', vigencias.join(','));
  if (soloCertificados) params.set('solo_certificados', '1');
  if (numeroFicha) params.set('numero_ficha', numeroFicha);
  if (activeSearchPrograma) params.set('search', activeSearchPrograma);
  const q = params.toString();
  return `${API_BASE}/programas/all${q ? `?${q}` : ''}`;
}

async function exportProgramasExcel() {
  setStatus('Generando Excel de programas...');
  try {
    const url = buildProgramasExportUrl();
    const resp = await fetch(url);
    if (!resp.ok) {
      const data = await resp.json().catch(() => null);
      const msg = data && data.detail ? data.detail : (resp.status + ' ' + resp.statusText);
      throw new Error(msg);
    }
    const blob = await resp.blob();
    const disposition = resp.headers.get('Content-Disposition') || '';
    const filename = getFilenameFromDisposition(disposition) || 'programas_export.xlsx';

    const objectUrl = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = objectUrl;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(objectUrl);

    setStatus('Excel de programas exportado correctamente.');
  } catch (e) {
    setStatus('Error exportando programas: ' + e.message);
    console.error(e);
  }
}

async function loadProgramas() {
  setStatus('Cargando programas...');
  try {
    const resp = await fetch(buildUrl());
    if (!resp.ok) throw new Error(resp.status + ' ' + resp.statusText);
    const data = await resp.json();
    allItems = Array.isArray(data.items) ? data.items : [];

    buildUniqueProgramDenoms();
    updateProgramSearchSuggestions(activeSearchPrograma);
    renderTable(allItems);
    updateHeaderInfo(data.fecha_corte, data.total || allItems.length);
    // Totales globales: sumar todos los registros que cumplen los filtros (no solo la página actual)
    fetchAllProgramasAndUpdateTotales();
    populateFilterOptions(allItems, data);
    updatePagination(data.total || 0, data.page || currentPage, data.per_page || PER_PAGE);

    setStatus(`Mostrando ${allItems.length} registros (total ${data.total || allItems.length}).`);
  } catch (e) {
    setStatus('Error: ' + e.message);
    console.error(e);
  }
}

// Obtener todos los registros filtrados desde el backend y actualizar los totales (lógica tipo fichas)
async function fetchAllProgramasAndUpdateTotales() {
  try {
    const url = buildProgramasAllUrl();
    const resp = await fetch(url);
    if (!resp.ok) throw new Error(resp.status + ' ' + resp.statusText);
    const items = await resp.json();
    const rows = Array.isArray(items) ? items : [];
    updateProgramasTotalsSummary(rows);
  } catch (e) {
    // Si falla, dejar los totales en cero
    updateProgramasTotalsSummary([]);
    console.error('Error actualizando totales de programas:', e);
  }
}

// Sumar matriculados (cupos), activos y certificados sobre el conjunto filtrado completo
function updateProgramasTotalsSummary(rows) {
  const totals = rows.reduce((acc, r) => {
    acc.matriculados += toNumberProgramas(r.cupos ?? r.cupo);
    acc.activos += toNumberProgramas(r.aprendices_activos);
    acc.certificados += toNumberProgramas(r.certificado);
    return acc;
  }, { matriculados: 0, activos: 0, certificados: 0 });

  const elRows = document.getElementById('sumProgRegistros');
  const elMat = document.getElementById('sumProgMatriculados');
  const elAct = document.getElementById('sumProgActivos');
  const elCert = document.getElementById('sumProgCertificados');

  if (elRows) elRows.textContent = formatNumberProgramas(rows.length);
  if (elMat) elMat.textContent = formatNumberProgramas(totals.matriculados);
  if (elAct) elAct.textContent = formatNumberProgramas(totals.activos);
  if (elCert) elCert.textContent = formatNumberProgramas(totals.certificados);
}




function updateHeaderInfo(fechaCorte, total) {
  const fc = document.getElementById('fechaCorteValue');
  const tr = document.getElementById('totalRowsValue');
  if (fc) fc.textContent = fechaCorte || '-';
  if (tr) tr.textContent = String(total || 0);
}

function populateFilterOptions(items, meta) {
  const yearEl = document.getElementById('filterYear');
  const muniEl = document.getElementById('filterMunicipio');
  const estEl = document.getElementById('filterEstrategia');
  const convEl = document.getElementById('filterConvenio');
  const vigEl = document.getElementById('filterVigencia');
  if (!yearEl || !muniEl || !estEl || !convEl || !vigEl) return;

  const selectedYears = Array.from(yearEl.selectedOptions || []).map(o => o.value);
  const selectedMunis = Array.from(muniEl.selectedOptions || []).map(o => o.value);

  // Usar opciones globales calculadas en el backend para que los filtros
  // sean consistentes para toda la tabla, no solo para la pagina actual.
  const years = new Set(globalFilterOptions?.years || []);
  const municipios = new Set(globalFilterOptions?.municipios || []);
  const estrategias = new Set(globalFilterOptions?.estrategias || []);
  const convenios = new Set(globalFilterOptions?.convenios || []);
  const vigencias = new Set(globalFilterOptions?.vigencias || []);

  yearEl.options.length = 1;
  Array.from(years).sort((a, b) => Number(b) - Number(a)).forEach((v) => {
    const o = document.createElement('option');
    o.value = v;
    o.text = v;
    yearEl.appendChild(o);
  });

  muniEl.options.length = 1;
  Array.from(municipios).sort((a, b) => a.localeCompare(b)).forEach((v) => {
    const o = document.createElement('option');
    o.value = v;
    o.text = v;
    muniEl.appendChild(o);
  });

  Array.from(yearEl.options || []).forEach(o => { o.selected = selectedYears.includes(o.value); });
  Array.from(muniEl.options || []).forEach(o => { o.selected = selectedMunis.includes(o.value); });

  const selectedEsts = Array.from(estEl.selectedOptions || []).map(o => o.value);
  const selectedConvs = Array.from(convEl.selectedOptions || []).map(o => o.value);
  const selectedVigs = Array.from(vigEl.selectedOptions || []).map(o => o.value);

  estEl.options.length = 1;
  Array.from(estrategias).sort((a, b) => a.localeCompare(b)).forEach((v) => {
    const o = document.createElement('option');
    o.value = v;
    o.text = v;
    estEl.appendChild(o);
  });

  convEl.options.length = 1;
  Array.from(convenios).sort((a, b) => a.localeCompare(b)).forEach((v) => {
    const o = document.createElement('option');
    o.value = v;
    o.text = v;
    convEl.appendChild(o);
  });

  Array.from(estEl.options || []).forEach(o => { o.selected = selectedEsts.includes(o.value); });
  Array.from(convEl.options || []).forEach(o => { o.selected = selectedConvs.includes(o.value); });

  vigEl.options.length = 1;
  Array.from(vigencias).sort((a, b) => Number(b) - Number(a)).forEach((v) => {
    const o = document.createElement('option');
    o.value = v;
    o.text = v;
    vigEl.appendChild(o);
  });
  Array.from(vigEl.options || []).forEach(o => { o.selected = selectedVigs.includes(o.value); });

  ['filterYear','filterMunicipio','filterEstrategia','filterConvenio','filterVigencia'].forEach(id => {
    setupMultiSelect(id);
  });
}

function updatePagination(total, page, perPage) {
  const infoEl = document.getElementById('pageInfo');
  const prevBtn = document.getElementById('prevPageBtn');
  const nextBtn = document.getElementById('nextPageBtn');
  if (!infoEl || !prevBtn || !nextBtn) return;

  const totalPages = Math.max(1, Math.ceil((total || 0) / (perPage || PER_PAGE)));
  currentPage = Math.min(Math.max(1, page || 1), totalPages);

  infoEl.textContent = `Página ${currentPage} de ${totalPages}`;
  prevBtn.disabled = currentPage <= 1;
  nextBtn.disabled = currentPage >= totalPages;
}

function changePage(delta) {
  currentPage = Math.max(1, (currentPage || 1) + delta);
  loadProgramas();
}

function renderTable(rows) {
  const thead = document.querySelector('#programasTable thead');
  const tbody = document.querySelector('#programasTable tbody');
  if (!thead || !tbody) return;
  thead.innerHTML = '';
  tbody.innerHTML = '';

  if (!rows.length) {
    thead.innerHTML = '<tr><th>No hay datos</th></tr>';
    return;
  }

  const keys = Object.keys(rows[0]).filter((k) => k !== 'id');
  const trh = document.createElement('tr');
  trh.innerHTML = keys.map((k) => `<th>${toLabel(k)}</th>`).join('');
  thead.appendChild(trh);

  rows.forEach((row) => {
    const tr = document.createElement('tr');
    tr.innerHTML = keys.map((k) => `<td>${escapeHtml(row[k])}</td>`).join('');
    tbody.appendChild(tr);
  });
}

async function uploadProgramasExcel() {
  const input = document.getElementById('programasFile');
  const files = input && input.files ? Array.from(input.files) : [];
  if (!files.length) {
    alert('Selecciona al menos un archivo Excel primero.');
    return;
  }

  let totalInserted = 0;
  setUploadProgress(0);
  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    const fd = new FormData();
    fd.append('file', file);

    const percent = Math.round((i / files.length) * 100);
    setUploadProgress(percent);
    setStatus(`Subiendo Excel de programas (${i + 1}/${files.length})...`);
    try {
      const resp = await fetch(`${API_BASE}/programas/upload-excel`, {
        method: 'POST',
        body: fd,
      });
      const data = await resp.json().catch(() => null);
      if (!resp.ok) {
        const msg = data && data.detail ? data.detail : (resp.status + ' ' + resp.statusText);
        throw new Error(msg);
      }
      totalInserted += data && typeof data.inserted === 'number' ? data.inserted : 0;
    } catch (e) {
      setStatus('Error al subir programas: ' + e.message + ` (archivo ${file.name})`);
      console.error(e);
      return;
    }
  }

  setUploadProgress(100);
  setStatus(`Subida completada. Filas insertadas en total: ${totalInserted}.`);
  loadProgramas();
}

async function uploadProgramasHistoricos() {
  const input = document.getElementById('programasHistoricosFile');
  const yearInput = document.getElementById('historicoYear');
  const yearVal = yearInput && yearInput.value ? yearInput.value.trim() : '';

  if (!yearVal) {
    alert('Ingresa el año histórico.');
    return;
  }

  const yearNum = parseInt(yearVal, 10);
  if (!Number.isFinite(yearNum) || yearNum < 1900 || yearNum > 2100) {
    alert('El año histórico debe estar entre 1900 y 2100.');
    return;
  }

  const files = input && input.files ? Array.from(input.files) : [];
  if (!files.length) {
    alert('Selecciona al menos un archivo Excel histórico de programas.');
    return;
  }

  let totalInserted = 0;
  let totalUpdated = 0;
  let totalDuplicateRows = 0;

  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    const fd = new FormData();
    fd.append('file', file);
    fd.append('year', String(yearNum));

    setStatus(`Subiendo Excel histórico de programas (${i + 1}/${files.length})...`);
    try {
      const resp = await fetch(`${API_BASE}/programas/upload-excel-historico`, {
        method: 'POST',
        body: fd,
      });
      const data = await resp.json().catch(() => null);
      if (!resp.ok) {
        const msg = data && data.detail ? data.detail : (resp.status + ' ' + resp.statusText);
        throw new Error(msg);
      }

      totalInserted += data && typeof data.inserted === 'number' ? data.inserted : 0;
      totalUpdated += data && typeof data.updated_fichas === 'number' ? data.updated_fichas : 0;
      if (data && typeof data.duplicate_rows === 'number') {
        totalDuplicateRows += data.duplicate_rows;
      }
    } catch (e) {
      setStatus('Error al subir históricos: ' + e.message + ` (archivo ${file.name})`);
      console.error(e);
      return;
    }
  }

  setStatus(
    `Subida histórica completada. Filas insertadas: ${totalInserted}. ` +
    `Fichas actualizadas (estrategia/estado) a partir de históricos: ${totalUpdated}. ` +
    `Registros del histórico que ya existían en la tabla (por numero_ficha): ${totalDuplicateRows}.`
  );
  loadProgramas();
}

async function uploadCertificadosExcel() {
  const input = document.getElementById('certificadosFile');
  const files = input && input.files ? Array.from(input.files) : [];
  if (!files.length) {
    alert('Selecciona al menos un archivo complementario de certificados.');
    return;
  }

  let totalUpdatedRows = 0;
  let totalUpdatedFichas = 0;
  let totalUnmatched = 0;

  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    const fd = new FormData();
    fd.append('file', file);

    setStatus(`Actualizando certificados (${i + 1}/${files.length})...`);
    try {
      const resp = await fetch(`${API_BASE}/programas/upload-certificados`, {
        method: 'POST',
        body: fd,
      });
      const data = await resp.json().catch(() => null);
      if (!resp.ok) {
        const msg = data && data.detail ? data.detail : (resp.status + ' ' + resp.statusText);
        throw new Error(msg);
      }

      totalUpdatedRows += data && data.updated_rows ? data.updated_rows : 0;
      totalUpdatedFichas += data && data.updated_fichas ? data.updated_fichas : 0;
      totalUnmatched += data && data.unmatched_fichas ? data.unmatched_fichas : 0;
    } catch (e) {
      setStatus('Error certificados: ' + e.message + ` (archivo ${file.name})`);
      console.error(e);
      return;
    }
  }

  setStatus(
    `Certificados actualizados. Filas impactadas: ${totalUpdatedRows}. ` +
    `Fichas actualizadas: ${totalUpdatedFichas}. Sin coincidencia (suma de todos los archivos): ${totalUnmatched}.`
  );
  loadProgramas();
}

async function uploadProgramasYCertificados() {
  const progInput = document.getElementById('programasFileCombo');
  const certInput = document.getElementById('certificadosFileCombo');
  const progFiles = progInput && progInput.files ? Array.from(progInput.files) : [];
  const certFiles = certInput && certInput.files ? Array.from(certInput.files) : [];

  if (!progFiles.length || !certFiles.length) {
    alert('Selecciona al menos un archivo de programas y uno de certificados.');
    return;
  }

  let totalInsertedProg = 0;
  let totalUpdatedRows = 0;
  let totalUpdatedFichas = 0;
  let totalUnmatched = 0;
  setUploadProgress(0);

  // 1) Subir todos los programas
  for (let i = 0; i < progFiles.length; i++) {
    const file = progFiles[i];
    const fdProg = new FormData();
    fdProg.append('file', file);

    const percentProg = Math.round((i / (progFiles.length + certFiles.length)) * 100);
    setUploadProgress(percentProg);
    setStatus(`Subiendo Excel de programas (${i + 1}/${progFiles.length})...`);
    try {
      const respProg = await fetch(`${API_BASE}/programas/upload-excel`, {
        method: 'POST',
        body: fdProg,
      });
      const dataProg = await respProg.json().catch(() => null);
      if (!respProg.ok) {
        const msgProg = dataProg && dataProg.detail ? dataProg.detail : (respProg.status + ' ' + respProg.statusText);
        throw new Error('Programas: ' + msgProg + ` (archivo ${file.name})`);
      }
      totalInsertedProg += dataProg && dataProg.inserted ? dataProg.inserted : 0;
    } catch (e) {
      setStatus('Error combinando carga (programas): ' + e.message);
      console.error(e);
      return;
    }
  }

  // 2) Subir todos los certificados
  for (let j = 0; j < certFiles.length; j++) {
    const file = certFiles[j];
    const fdCert = new FormData();
    fdCert.append('file', file);

    const percentCert = Math.round(((progFiles.length + j) / (progFiles.length + certFiles.length)) * 100);
    setUploadProgress(percentCert);
    setStatus(`Programas cargados. Actualizando certificados (${j + 1}/${certFiles.length})...`);
    try {
      const respCert = await fetch(`${API_BASE}/programas/upload-certificados`, {
        method: 'POST',
        body: fdCert,
      });
      const dataCert = await respCert.json().catch(() => null);
      if (!respCert.ok) {
        const msgCert = dataCert && dataCert.detail ? dataCert.detail : (respCert.status + ' ' + respCert.statusText);
        throw new Error('Certificados: ' + msgCert + ` (archivo ${file.name})`);
      }

      totalUpdatedRows += dataCert && dataCert.updated_rows ? dataCert.updated_rows : 0;
      totalUpdatedFichas += dataCert && dataCert.updated_fichas ? dataCert.updated_fichas : 0;
      totalUnmatched += dataCert && dataCert.unmatched_fichas ? dataCert.unmatched_fichas : 0;
    } catch (e) {
      setStatus('Error combinando carga (certificados): ' + e.message);
      console.error(e);
      return;
    }
  }

  setUploadProgress(100);
  setStatus(
    `Programas y certificados procesados. Programas insertados (suma de todos los archivos): ${totalInsertedProg}. ` +
    `Filas impactadas en certificados: ${totalUpdatedRows}. Fichas actualizadas: ${totalUpdatedFichas}. ` +
    `Sin coincidencia (suma de todos los archivos): ${totalUnmatched}.`
  );
  loadProgramas();
}
