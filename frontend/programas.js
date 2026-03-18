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
const PER_PAGE = 30;
let globalFilterOptions = null;

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
    const y = document.getElementById('filterYear');
    const m = document.getElementById('filterMunicipio');
    const nf = document.getElementById('filterNumeroFicha');
    if (y) y.value = '';
    if (m) m.value = '';
    const e = document.getElementById('filterEstrategia');
    const c = document.getElementById('filterConvenio');
    const v = document.getElementById('filterVigencia');
    const sc = document.getElementById('filterSoloCertificados');
    if (e) e.value = '';
    if (c) c.value = '';
    if (v) v.value = '';
    if (sc) sc.checked = false;
    if (nf) nf.value = '';
    loadProgramas();
  });
  const exportBtn = document.getElementById('exportProgramasBtn');
  if (exportBtn) exportBtn.addEventListener('click', exportProgramasExcel);
  await loadGlobalFilterOptions();
  loadProgramas();
});

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
  const vigencia = (document.getElementById('filterVigencia')?.value || '').trim();
  const soloCertificados = !!document.getElementById('filterSoloCertificados')?.checked;
  const numeroFicha = (document.getElementById('filterNumeroFicha')?.value || '').trim();
  return { year, municipio, estrategia, convenio, vigencia, soloCertificados, numeroFicha };
}

function buildUrl() {
  const { year, municipio, estrategia, convenio, vigencia, soloCertificados, numeroFicha } = getFilters();
  const params = new URLSearchParams();
  if (year) params.set('year', year);
  if (municipio) params.set('municipio', municipio);
  if (estrategia) params.set('estrategia', estrategia);
  if (convenio) params.set('convenio', convenio);
  if (vigencia) params.set('vigencia', vigencia);
   if (numeroFicha) params.set('numero_ficha', numeroFicha);
  if (soloCertificados) params.set('solo_certificados', '1');
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
  const { year, municipio, estrategia, convenio, vigencia, soloCertificados, numeroFicha } = getFilters();
  const params = new URLSearchParams();
  if (year) params.set('year', year);
  if (municipio) params.set('municipio', municipio);
  if (estrategia) params.set('estrategia', estrategia);
  if (convenio) params.set('convenio', convenio);
  if (vigencia) params.set('vigencia', vigencia);
  if (soloCertificados) params.set('solo_certificados', '1');
  if (numeroFicha) params.set('numero_ficha', numeroFicha);
  const q = params.toString();
  return `${API_BASE}/programas/export${q ? `?${q}` : ''}`;
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
  const vigEl = document.getElementById('filterVigencia');
  if (!yearEl || !muniEl || !estEl || !convEl || !vigEl) return;

  const selectedYear = yearEl.value;
  const selectedMuni = muniEl.value;

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

  if (selectedYear) yearEl.value = selectedYear;
  if (selectedMuni) muniEl.value = selectedMuni;

  const selectedEst = estEl.value;
  const selectedConv = convEl.value;
  const selectedVig = vigEl.value;

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

  vigEl.options.length = 1;
  Array.from(vigencias).sort((a, b) => Number(b) - Number(a)).forEach((v) => {
    const o = document.createElement('option');
    o.value = v;
    o.text = v;
    vigEl.appendChild(o);
  });
  if (selectedVig) vigEl.value = selectedVig;
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
