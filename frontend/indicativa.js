function applyTheme(mode){
  const isDark = mode === 'dark';
  document.body.classList.toggle('dark-mode', isDark);
  const btn = document.getElementById('themeToggle');
  if(btn) btn.textContent = isDark ? '☀️ Modo claro' : '🌙 Modo oscuro';
}

function initThemeToggle(){
  const savedTheme = localStorage.getItem('themeMode') || 'light';
  applyTheme(savedTheme);
  const btn = document.getElementById('themeToggle');
  if(!btn) return;
  btn.addEventListener('click', ()=>{
    const next = document.body.classList.contains('dark-mode') ? 'light' : 'dark';
    localStorage.setItem('themeMode', next);
    applyTheme(next);
  });
}

const API_BASE = 'http://127.0.0.1:8000';
let allIndicativa = [];
let currentPage = 1;
const PER_PAGE = 50;
let indicativaFiltersMeta = null;

function setStatus(msg){
  const el = document.getElementById('status');
  if(el) el.textContent = msg || '';
}

function setUploadProgress(percent){
  const container = document.getElementById('uploadProgressContainer');
  const bar = document.getElementById('uploadProgressBar');
  if(!container || !bar) return;
  if(!percent || percent <= 0){
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

function escapeHtml(v){
  if(v === null || v === undefined) return '';
  return String(v).replace(/[&<>\"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
}

async function loadIndicativa(){
  setStatus('Cargando indicativa...');
  try{
    const resp = await fetch(buildIndicativaUrl());
    if(!resp.ok) throw new Error(resp.status + ' ' + resp.statusText);
    const data = await resp.json();
    allIndicativa = Array.isArray(data.items) ? data.items : [];
    renderTable(allIndicativa);
    updatePagination(data.total || allIndicativa.length, data.page || currentPage, data.per_page || PER_PAGE);
    setStatus(`Mostrando ${allIndicativa.length} registros (total ${data.total || allIndicativa.length}).`);
  }catch(e){
    setStatus('Error: ' + e.message);
    console.error(e);
  }
}

function getIndicativaFilters(){
  const centro = (document.getElementById('filterCentro')?.value || '').trim();
  const nivel = (document.getElementById('filterNivel')?.value || '').trim();
  const periodo = (document.getElementById('filterPeriodoOferta')?.value || '').trim();
  return { centro, nivel, periodo };
}

function buildIndicativaUrl(){
  const { centro, nivel, periodo } = getIndicativaFilters();
  const params = new URLSearchParams();
  params.set('page', String(currentPage || 1));
  params.set('per_page', String(PER_PAGE));
  if(centro) params.set('centro', centro);
  if(nivel) params.set('nivel', nivel);
  if(periodo) params.set('periodo_oferta', periodo);
  const q = params.toString();
  return `${API_BASE}/indicativa${q ? `?${q}` : ''}`;
}

function updatePagination(total, page, perPage){
  const infoEl = document.getElementById('pageInfo');
  const prevBtn = document.getElementById('prevPageBtn');
  const nextBtn = document.getElementById('nextPageBtn');
  if(!infoEl || !prevBtn || !nextBtn) return;

  const totalPages = Math.max(1, Math.ceil((total || 0) / (perPage || PER_PAGE)));
  currentPage = Math.min(Math.max(1, page || 1), totalPages);

  infoEl.textContent = `Página ${currentPage} de ${totalPages}`;
  prevBtn.disabled = currentPage <= 1;
  nextBtn.disabled = currentPage >= totalPages;
}

function changePage(delta){
  currentPage = Math.max(1, (currentPage || 1) + delta);
  loadIndicativa();
}

function renderTable(rows){
  const tbody = document.querySelector('#indicativaTable tbody');
  if(!tbody) return;
  tbody.innerHTML = '';
  if(!rows.length){
    const tr = document.createElement('tr');
    tr.innerHTML = '<td colspan="6" class="text-center">No hay datos</td>';
    tbody.appendChild(tr);
    return;
  }

  rows.forEach((row, idx) => {
    const tr = document.createElement('tr');
    const displayId = (currentPage - 1) * PER_PAGE + idx + 1;
    tr.innerHTML = `
      <td>${escapeHtml(displayId)}</td>
      <td>${escapeHtml(row.centro_formacion)}</td>
      <td>${escapeHtml(row.nivel_formacion)}</td>
      <td>${escapeHtml(row.denominacion_programa)}</td>
      <td>${escapeHtml(row.periodo_oferta)}</td>
      <td>${escapeHtml(row.tipo_oferta)}</td>
    `;
    tbody.appendChild(tr);
  });
}

async function uploadIndicativaExcel(){
  const input = document.getElementById('indicativaFile');
  const files = input && input.files ? Array.from(input.files) : [];
  if(!files.length){
    alert('Selecciona al menos un archivo Excel primero.');
    return;
  }

  setStatus(`Subiendo ${files.length} archivo(s) de indicativa...`);
  setUploadProgress(5);

  let totalInserted = 0;
  let processed = 0;

  try{
    for(const file of files){
      const fd = new FormData();
      fd.append('file', file);

      const resp = await fetch(`${API_BASE}/indicativa/upload-excel`, {
        method: 'POST',
        body: fd,
      });
      const data = await resp.json().catch(() => null);
      if(!resp.ok){
        const msg = data && data.detail ? data.detail : (resp.status + ' ' + resp.statusText);
        throw new Error(`Error con el archivo "${file.name}": ${msg}`);
      }
      const inserted = data && data.inserted ? data.inserted : 0;
      totalInserted += inserted;
      processed++;
      const percent = Math.round((processed / files.length) * 100);
      setUploadProgress(percent);
      setStatus(`Procesado ${processed} de ${files.length} archivo(s). Filas insertadas acumuladas: ${totalInserted}.`);
    }

    setUploadProgress(100);
    setStatus(`Subida finalizada. Archivos procesados: ${files.length}. Filas insertadas totales: ${totalInserted}.`);
    await loadIndicativa();
  }catch(e){
    setStatus('Error upload: ' + e.message);
    console.error(e);
  }
}

function getFilenameFromDisposition(contentDisposition){
  if(!contentDisposition) return '';
  const match = contentDisposition.match(/filename="?([^";]+)"?/i);
  return match && match[1] ? match[1] : '';
}

function buildIndicativaExportUrl(){
  const { centro, nivel, periodo } = getIndicativaFilters();
  const params = new URLSearchParams();
  if(centro) params.set('centro', centro);
  if(nivel) params.set('nivel', nivel);
  if(periodo) params.set('periodo_oferta', periodo);
  const q = params.toString();
  return `${API_BASE}/indicativa/export${q ? `?${q}` : ''}`;
}

async function exportIndicativaExcel(){
  const url = buildIndicativaExportUrl();
  setStatus('Generando Excel de indicativa...');
  try{
    const resp = await fetch(url);
    if(!resp.ok){
      const data = await resp.json().catch(() => null);
      const msg = data && data.detail ? data.detail : (resp.status + ' ' + resp.statusText);
      throw new Error(msg);
    }
    const blob = await resp.blob();
    const disposition = resp.headers.get('Content-Disposition') || '';
    const filename = getFilenameFromDisposition(disposition) || 'indicativa_export.xlsx';

    const objectUrl = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = objectUrl;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(objectUrl);

    setStatus('Excel de indicativa exportado correctamente.');
  }catch(e){
    setStatus('Error exportando indicativa: ' + e.message);
    console.error(e);
  }
}

async function loadIndicativaFilterOptions(){
  try{
    const resp = await fetch(`${API_BASE}/indicativa/filters`);
    if(!resp.ok) throw new Error(resp.status + ' ' + resp.statusText);
    indicativaFiltersMeta = await resp.json();
    populateIndicativaFilterOptions();
  }catch(e){
    console.error('Error cargando filtros de indicativa:', e);
    indicativaFiltersMeta = null;
  }
}

function populateIndicativaFilterOptions(){
  if(!indicativaFiltersMeta) return;
  const centroEl = document.getElementById('filterCentro');
  const nivelEl = document.getElementById('filterNivel');
  const periodoEl = document.getElementById('filterPeriodoOferta');
  if(!centroEl || !nivelEl || !periodoEl) return;

  const selectedCentro = centroEl.value;
  const selectedNivel = nivelEl.value;
  const selectedPeriodo = periodoEl.value;

  centroEl.options.length = 1;
  (indicativaFiltersMeta.centros || []).forEach(v => {
    const o = document.createElement('option');
    o.value = v;
    o.textContent = v;
    centroEl.appendChild(o);
  });

  nivelEl.options.length = 1;
  (indicativaFiltersMeta.niveles || []).forEach(v => {
    const o = document.createElement('option');
    o.value = v;
    o.textContent = v;
    nivelEl.appendChild(o);
  });

  periodoEl.options.length = 1;
  (indicativaFiltersMeta.periodos_oferta || []).forEach(v => {
    const o = document.createElement('option');
    o.value = v;
    o.textContent = v;
    periodoEl.appendChild(o);
  });

  if(selectedCentro) centroEl.value = selectedCentro;
  if(selectedNivel) nivelEl.value = selectedNivel;
  if(selectedPeriodo) periodoEl.value = selectedPeriodo;
}

document.addEventListener('DOMContentLoaded', () => {
  initThemeToggle();
  const prevBtn = document.getElementById('prevPageBtn');
  const nextBtn = document.getElementById('nextPageBtn');
  if(prevBtn) prevBtn.addEventListener('click', () => changePage(-1));
  if(nextBtn) nextBtn.addEventListener('click', () => changePage(1));
  const uploadBtn = document.getElementById('uploadIndicativaBtn');
  if(uploadBtn) uploadBtn.addEventListener('click', uploadIndicativaExcel);
  const applyBtn = document.getElementById('applyFiltersBtn');
  const clearBtn = document.getElementById('clearFiltersBtn');
  const exportBtn = document.getElementById('exportIndicativaBtn');
  if(applyBtn) applyBtn.addEventListener('click', () => { currentPage = 1; loadIndicativa(); });
  if(clearBtn) clearBtn.addEventListener('click', () => {
    const centroEl = document.getElementById('filterCentro');
    const nivelEl = document.getElementById('filterNivel');
    const periodoEl = document.getElementById('filterPeriodoOferta');
    if(centroEl) centroEl.value = '';
    if(nivelEl) nivelEl.value = '';
    if(periodoEl) periodoEl.value = '';
    currentPage = 1;
    loadIndicativa();
  });
  if(exportBtn) exportBtn.addEventListener('click', exportIndicativaExcel);
  loadIndicativaFilterOptions().finally(() => {
    loadIndicativa();
  });
});
