const API_BASE = 'http://127.0.0.1:8000';
let allItems = [];
let currentPage = 1;
const PER_PAGE = 30;

document.addEventListener('DOMContentLoaded', () => {
  document.getElementById('uploadProgramasBtn').addEventListener('click', uploadProgramasExcel);
  document.getElementById('uploadCertificadosBtn').addEventListener('click', uploadCertificadosExcel);
  const comboBtn = document.getElementById('uploadProgramasCertificadosBtn');
  if (comboBtn) comboBtn.addEventListener('click', uploadProgramasYCertificados);
  const prevBtn = document.getElementById('prevPageBtn');
  const nextBtn = document.getElementById('nextPageBtn');
  if (prevBtn) prevBtn.addEventListener('click', () => changePage(-1));
  if (nextBtn) nextBtn.addEventListener('click', () => changePage(1));
  document.getElementById('applyFiltersBtn').addEventListener('click', loadProgramas);
  document.getElementById('clearFiltersBtn').addEventListener('click', () => {
    const y = document.getElementById('filterYear');
    const m = document.getElementById('filterMunicipio');
    if (y) y.value = '';
    if (m) m.value = '';
    loadProgramas();
  });
  loadProgramas();
});

function setStatus(msg) {
  const el = document.getElementById('status');
  if (el) el.textContent = msg || '';
}

function escapeHtml(v) {
  if (v === null || v === undefined) return '';
  return String(v).replace(/[&<>\"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}

function toLabel(key) {
  return String(key || '')
    .replace(/_/g, ' ')
    .split(' ')
    .map((w) => w ? w.charAt(0).toUpperCase() + w.slice(1) : '')
    .join(' ');
}

function getFilters() {
  const year = (document.getElementById('filterYear')?.value || '').trim();
  const municipio = (document.getElementById('filterMunicipio')?.value || '').trim();
  const estrategia = (document.getElementById('filterEstrategia')?.value || '').trim();
  const convenio = (document.getElementById('filterConvenio')?.value || '').trim();
  return { year, municipio, estrategia, convenio };
}

function buildUrl() {
  const { year, municipio, estrategia, convenio } = getFilters();
  const params = new URLSearchParams();
  if (year) params.set('year', year);
  if (municipio) params.set('municipio', municipio);
  if (estrategia) params.set('estrategia', estrategia);
  if (convenio) params.set('convenio', convenio);
  params.set('page', String(currentPage || 1));
  params.set('per_page', String(PER_PAGE));
  const q = params.toString();
  return `${API_BASE}/programas${q ? `?${q}` : ''}`;
}

async function loadProgramas() {
  setStatus('Cargando programas...');
  try {
    const resp = await fetch(buildUrl());
    if (!resp.ok) throw new Error(resp.status + ' ' + resp.statusText);
    const data = await resp.json();
    allItems = Array.isArray(data.items) ? data.items : [];

    renderTable(allItems);
    updateHeaderInfo(data.fecha_corte, data.total || allItems.length);
    populateFilterOptions(allItems, data);
    updatePagination(data.total || 0, data.page || currentPage, data.per_page || PER_PAGE);

    setStatus(`Mostrando ${allItems.length} registros (total ${data.total || allItems.length}).`);
  } catch (e) {
    setStatus('Error: ' + e.message);
    console.error(e);
  }
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
  if (!yearEl || !muniEl || !estEl || !convEl) return;

  const selectedYear = yearEl.value;
  const selectedMuni = muniEl.value;

  const years = new Set();
  const municipios = new Set();
  const estrategias = new Set();
  const convenios = new Set();

  items.forEach((r) => {
    const fc = (r.fecha_corte || '').toString().trim();
    if (fc.length >= 4) years.add(fc.slice(0, 4));
    const m = (r.ciudad_municipio || '').toString().trim();
    if (m) municipios.add(m);
    const est = (r.estrategia_programa || '').toString().trim();
    if (est) estrategias.add(est);
    const cv = (r.convenio || '').toString().trim();
    if (cv) convenios.add(cv);
  });

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

  if (selectedYear) yearEl.value = selectedYear;
  if (selectedMuni) muniEl.value = selectedMuni;

  const selectedEst = estEl.value;
  const selectedConv = convEl.value;

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

  if (selectedEst) estEl.value = selectedEst;
  if (selectedConv) convEl.value = selectedConv;
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
  const file = input && input.files ? input.files[0] : null;
  if (!file) {
    alert('Selecciona un archivo Excel primero.');
    return;
  }

  const fd = new FormData();
  fd.append('file', file);

  setStatus('Subiendo Excel...');
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

    setStatus(`Subido correctamente. Filas insertadas: ${data.inserted || 0}.`);
    loadProgramas();
  } catch (e) {
    setStatus('Error upload: ' + e.message);
    console.error(e);
  }
}

async function uploadCertificadosExcel() {
  const input = document.getElementById('certificadosFile');
  const file = input && input.files ? input.files[0] : null;
  if (!file) {
    alert('Selecciona el archivo complementario de certificados.');
    return;
  }

  const fd = new FormData();
  fd.append('file', file);

  setStatus('Actualizando certificados...');
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

    const updatedRows = data && data.updated_rows ? data.updated_rows : 0;
    const updatedFichas = data && data.updated_fichas ? data.updated_fichas : 0;
    const unmatched = data && data.unmatched_fichas ? data.unmatched_fichas : 0;
    setStatus(`Certificados actualizados. Filas impactadas: ${updatedRows}. Fichas actualizadas: ${updatedFichas}. Sin coincidencia: ${unmatched}.`);
    loadProgramas();
  } catch (e) {
    setStatus('Error certificados: ' + e.message);
    console.error(e);
  }
}

async function uploadProgramasYCertificados() {
  const progInput = document.getElementById('programasFileCombo');
  const certInput = document.getElementById('certificadosFileCombo');
  const progFile = progInput && progInput.files ? progInput.files[0] : null;
  const certFile = certInput && certInput.files ? certInput.files[0] : null;

  if (!progFile || !certFile) {
    alert('Selecciona ambos archivos: programas y certificados.');
    return;
  }

  // 1) Subir programas
  const fdProg = new FormData();
  fdProg.append('file', progFile);

  setStatus('Subiendo Excel de programas...');
  try {
    const respProg = await fetch(`${API_BASE}/programas/upload-excel`, {
      method: 'POST',
      body: fdProg,
    });
    const dataProg = await respProg.json().catch(() => null);
    if (!respProg.ok) {
      const msgProg = dataProg && dataProg.detail ? dataProg.detail : (respProg.status + ' ' + respProg.statusText);
      throw new Error('Programas: ' + msgProg);
    }

    // 2) Subir certificados usando el segundo archivo
    const fdCert = new FormData();
    fdCert.append('file', certFile);

    setStatus('Programas cargados. Actualizando certificados...');
    const respCert = await fetch(`${API_BASE}/programas/upload-certificados`, {
      method: 'POST',
      body: fdCert,
    });
    const dataCert = await respCert.json().catch(() => null);
    if (!respCert.ok) {
      const msgCert = dataCert && dataCert.detail ? dataCert.detail : (respCert.status + ' ' + respCert.statusText);
      throw new Error('Certificados: ' + msgCert);
    }

    const inserted = dataProg && dataProg.inserted ? dataProg.inserted : 0;
    const updatedRows = dataCert && dataCert.updated_rows ? dataCert.updated_rows : 0;
    const updatedFichas = dataCert && dataCert.updated_fichas ? dataCert.updated_fichas : 0;
    const unmatched = dataCert && dataCert.unmatched_fichas ? dataCert.unmatched_fichas : 0;

    setStatus(
      `Programas y certificados procesados. Programas insertados: ${inserted}. ` +
      `Filas impactadas en certificados: ${updatedRows}. Fichas actualizadas: ${updatedFichas}. Sin coincidencia: ${unmatched}.`
    );
    loadProgramas();
  } catch (e) {
    setStatus('Error combinando carga: ' + e.message);
    console.error(e);
  }
}
