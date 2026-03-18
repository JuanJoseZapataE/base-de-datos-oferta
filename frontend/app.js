const perPage = 30;
let items = [];
let page = 1;
let activeFilters = { centro: '', oferta: '', estado: '', tipo: '', nivel: '', periodo: '' };
let activeSearch = '';
let lastSearchSource = '';
let lastTotalCount = 0;
let pollIntervalId = null;
const API_BASE = 'http://127.0.0.1:8000';

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

// Autocompletado: lista única de denominaciones
let uniqueDenoms = [];
function buildUniqueDenoms(){
  const s = new Set();
  items.forEach(r => { const d = (r.denominacion_programa||'').toString().trim(); if(d) s.add(d); });
  uniqueDenoms = Array.from(s).sort((a,b)=>a.localeCompare(b, undefined, {sensitivity:'base'}));
}
function updateSearchSuggestions(prefix){
  const list = document.getElementById('searchSuggestions');
  if(!list) return;
  const p = (prefix||'').toString().toLowerCase();
  const matches = p ? uniqueDenoms.filter(d=>d.toLowerCase().includes(p)).slice(0,25) : uniqueDenoms.slice(0,25);
  list.innerHTML = '';
  matches.forEach(m => { const opt = document.createElement('option'); opt.value = m; list.appendChild(opt); });
}

document.addEventListener('DOMContentLoaded', ()=>{
  initThemeToggle();
  initUploadBlocks();
  // Cargar datos automáticamente al iniciar
  loadAll();
  const filterBtn = document.getElementById('filterBtn');
  if(filterBtn){
    filterBtn.addEventListener('click', ()=>{ readFilters(); page = 1; render(); });
  }
  const clearBtn = document.getElementById('clearFiltersBtn');
  if(clearBtn){
    clearBtn.addEventListener('click', ()=>{
    const fc = document.getElementById('filterCentro'); if(fc) fc.value='';
    const fo = document.getElementById('filterOferta'); if(fo) fo.value='';
    const ft = document.getElementById('filterTipo'); if(ft) ft.value='';
    const fn = document.getElementById('filterNivel'); if(fn) fn.value='';
    const fe = document.getElementById('filterEstado'); if(fe) fe.value='';
    const fp = document.getElementById('filterPeriodo'); if(fp) fp.value='';
    const s = document.getElementById('searchInput'); if(s) s.value='';
    readFilters(); page=1; render();
    });
  }
  const searchEl = document.getElementById('searchInput');
  if(searchEl){
    searchEl.addEventListener('input', ev => {
      const v = ev.target.value || '';
      activeSearch = v;
      lastSearchSource = 'filter';
      // Sincronizar con el buscador del encabezado si existe
      const headerSearch = document.getElementById('searchInputHeader');
      if(headerSearch && headerSearch.value !== v) headerSearch.value = v;
      updateSearchSuggestions(v);
      page = 1;
      render();
    });
  }
  bindUploadModuleEvents();
  const exportBtn = document.getElementById('exportExcelBtn');
  if(exportBtn){
    exportBtn.addEventListener('click', exportFilteredExcel);
  }
  const firstBtn = document.getElementById('first');
  const prevBtn = document.getElementById('prev');
  const nextBtn = document.getElementById('next');
  const lastBtn = document.getElementById('last');
  if(firstBtn){ firstBtn.addEventListener('click', ()=>{ page = 1; render(); }); }
  if(prevBtn){ prevBtn.addEventListener('click', ()=>{ if(page>1){ page--; render(); } }); }
  if(nextBtn){
    nextBtn.addEventListener('click', ()=>{
      const totalPages = Math.max(1, Math.ceil(items.length / perPage));
      if(page<totalPages){ page++; render(); }
    });
  }
  if(lastBtn){
    lastBtn.addEventListener('click', ()=>{
      const totalPages = Math.max(1, Math.ceil(items.length / perPage));
      page = totalPages;
      render();
    });
  }
  // Iniciar polling para detectar nuevos registros en segundo plano
  startPolling();
});

async function loadAll(){
  const base = API_BASE;
  const url = base + '/fichas/all';
  setStatus('Cargando...');
  try{
    const resp = await fetch(url);
    if(!resp.ok) throw new Error(resp.status + ' ' + resp.statusText);
    items = await resp.json();
    page = 1;
    buildUniqueDenoms();
    updateSearchSuggestions('');

    populateFilterOptions();
    render();
    // actualizar contador conocido
    lastTotalCount = items.length || 0;
    // render() actualizará el estado mostrando los registros totales o filtrados
  }catch(e){
    setStatus('Error: '+e.message);
    console.error(e);
  }
}

// Polling: consultar /fichas/count periódicamente y recargar si cambió el total
function startPolling(intervalMs = 8000){
  if(pollIntervalId) return;
  pollIntervalId = setInterval(async ()=>{
    try{
      const resp = await fetch(API_BASE + '/fichas/count');
      if(!resp.ok) return;
      const j = await resp.json();
      const total = typeof j.total === 'number' ? j.total : (Array.isArray(j.sample) ? j.sample.length : 0);
      if(total !== lastTotalCount){
        // hubo cambios, recargar
        await loadAll();
      }
    }catch(e){ /* no interrumpir por errores de red */ }
  }, intervalMs);
}

function stopPolling(){ if(pollIntervalId){ clearInterval(pollIntervalId); pollIntervalId = null; } }

function initUploadBlocks(){
  const wrap = document.getElementById('uploadBlocks');
  if(!wrap) return;
  if(wrap.children.length === 0) addUploadBlock();
}

function bindUploadModuleEvents(){
  const addBtn = document.getElementById('addUploadBlockBtn');
  const uploadAllBtn = document.getElementById('uploadAllBtn');
  if(addBtn){
    addBtn.addEventListener('click', ()=> addUploadBlock());
  }
  if(uploadAllBtn){
    uploadAllBtn.addEventListener('click', uploadAllExcels);
  }
}

function addUploadBlock(){
  const wrap = document.getElementById('uploadBlocks');
  if(!wrap) return;
  const idx = wrap.children.length + 1;
  const block = document.createElement('div');
  block.className = 'border rounded p-3 upload-block';
  block.innerHTML = `
    <div class="d-flex justify-content-between align-items-center mb-2">
      <strong>Archivo ${idx}</strong>
      <div class="d-flex gap-2">
        <button type="button" class="btn btn-sm btn-primary upload-single-block">Subir este archivo</button>
        <button type="button" class="btn btn-sm btn-outline-danger remove-upload-block">Quitar</button>
      </div>
    </div>
    <div class="row g-2 align-items-center">
      <div class="col-md-2">
        <label class="form-label">Periodo (Ano) *</label>
        <input class="form-control upload-periodo" type="number" placeholder="2025">
      </div>
      <div class="col-md-3">
        <label class="form-label">Oferta *</label>
        <select class="form-select upload-oferta">
          <option value="">(selecciona)</option>
          <option value="1">1</option>
          <option value="2">2</option>
          <option value="3">3</option>
          <option value="4">4</option>
        </select>
      </div>
      <div class="col-md-3">
        <label class="form-label">Tipo *</label>
        <select class="form-select upload-tipo">
          <option value="">(selecciona)</option>
          <option value="PRESENCIAL Y A DISTANCIA">PRESENCIAL Y A DISTANCIA</option>
          <option value="VIRTUAL">VIRTUAL</option>
        </select>
      </div>
      <div class="col-md-4">
        <label class="form-label">Archivo Excel *</label>
        <div class="excel-picker">
          <input type="file" accept=".xls,.xlsx,.xml" class="form-control upload-file" multiple>
        </div>
      </div>
    </div>
  `;
  wrap.appendChild(block);

  const removeBtn = block.querySelector('.remove-upload-block');
  if(removeBtn){
    removeBtn.addEventListener('click', ()=>{
      block.remove();
      renumberUploadBlocks();
      if(wrap.children.length === 0) addUploadBlock();
    });
  }
  const uploadSingleBtn = block.querySelector('.upload-single-block');
  if(uploadSingleBtn){
    uploadSingleBtn.addEventListener('click', async ()=>{
      const blocks = Array.from(document.querySelectorAll('#uploadBlocks .upload-block'));
      const position = blocks.indexOf(block) + 1;
      const entry = readUploadBlock(block);
      const err = validateUploadEntry(entry, position || 1);
      if(err){
        alert(err);
        return;
      }
      const totalFiles = entry.files ? entry.files.length : 0;
      showLoading();
      disableActions(true);
      try{
        let totalInserted = 0;
        for(let i = 0; i < entry.files.length; i++){
          const file = entry.files[i];
          const percent = totalFiles ? Math.round((i / totalFiles) * 100) : 0;
          setUploadProgress(percent);
          setStatus(`Subiendo archivo ${i + 1} de ${totalFiles} (bloque ${position || 1})...`);
          const result = await uploadOneExcel(entry, file);
          const inserted = result && result.inserted ? Number(result.inserted) : 0;
          totalInserted += inserted;
        }
        setUploadProgress(100);
        setStatus(`Subida completada. Archivos en este bloque: ${totalFiles}. Filas insertadas: ${totalInserted}. Recargando datos...`);
        showToast(`Subidos ${totalFiles} archivos en este bloque. Filas insertadas: ${totalInserted}.`, 'success');
        await loadAll();
      }catch(e){
        setStatus('Error upload: '+e.message);
        showToast('Error upload: '+e.message, 'danger');
        console.error(e);
      }finally{
        hideLoading();
        disableActions(false);
      }
    });
  }
  renumberUploadBlocks();
}

function renumberUploadBlocks(){
  const blocks = Array.from(document.querySelectorAll('#uploadBlocks .upload-block'));
  blocks.forEach((block, i)=>{
    const title = block.querySelector('strong');
    if(title) title.textContent = `Archivo ${i + 1}`;
  });
}

function readUploadBlock(block){
  const periodo = (block.querySelector('.upload-periodo')?.value || '').trim();
  const oferta = (block.querySelector('.upload-oferta')?.value || '').trim();
  const tipo = (block.querySelector('.upload-tipo')?.value || '').trim();
  const fileInput = block.querySelector('.upload-file');
  const files = fileInput && fileInput.files ? Array.from(fileInput.files) : [];
  return { periodo, oferta, tipo, files };
}

function validateUploadEntry(entry, position){
  if(!entry.periodo) return `Archivo ${position}: Debes indicar el Periodo (ano).`;
  if(!entry.oferta) return `Archivo ${position}: Debes indicar la Oferta.`;
  if(!entry.tipo) return `Archivo ${position}: Debes indicar el Tipo.`;
  if(!entry.files || entry.files.length === 0) return `Archivo ${position}: Debes seleccionar al menos un archivo Excel.`;
  return '';
}

async function uploadOneExcel(entry, file){
  const fd = new FormData();
  fd.append('file', file);
  fd.append('periodo', entry.periodo);
  fd.append('oferta', entry.oferta);
  fd.append('tipo', entry.tipo);
  const resp = await fetch(API_BASE + '/upload-excel', { method: 'POST', body: fd });
  const data = await resp.json().catch(()=>null);
  if(!resp.ok){
    const msg = data && data.detail ? data.detail : (resp.status + ' ' + resp.statusText);
    throw new Error(msg);
  }
  return data;
}

async function uploadAllExcels(){
  const blocks = Array.from(document.querySelectorAll('#uploadBlocks .upload-block'));
  if(blocks.length === 0){
    alert('Agrega al menos un archivo para subir.');
    return;
  }

  const entries = blocks.map(readUploadBlock);
  for(let i = 0; i < entries.length; i++){
    const err = validateUploadEntry(entries[i], i + 1);
    if(err){
      alert(err);
      return;
    }
  }

  showLoading();
  disableActions(true);
  let totalInserted = 0;
  // total de archivos (sumando todos los bloques)
  const totalFiles = entries.reduce((acc, e) => acc + (e.files ? e.files.length : 0), 0) || 0;
  let processedFiles = 0;
  try{
    for(let i = 0; i < entries.length; i++){
      const entry = entries[i];
      for(let j = 0; j < entry.files.length; j++){
        processedFiles += 1;
        const percent = totalFiles ? Math.round(((processedFiles - 1) / totalFiles) * 100) : 0;
        setUploadProgress(percent);
        setStatus(`Subiendo archivo ${processedFiles} de ${totalFiles} (bloque ${i + 1})...`);
        const result = await uploadOneExcel(entry, entry.files[j]);
        const inserted = result && result.inserted ? Number(result.inserted) : 0;
        totalInserted += inserted;
      }
    }
    setUploadProgress(100);
    setStatus(`Subidos ${totalFiles} archivos. Filas insertadas: ${totalInserted}. Recargando datos...`);
    showToast(`Subidos ${totalFiles} archivos. Filas insertadas: ${totalInserted}.`, 'success');
    await loadAll();
  }catch(e){
    setStatus('Error upload: '+e.message);
    showToast('Error upload: '+e.message, 'danger');
    console.error(e);
  }finally{
    hideLoading();
    disableActions(false);
  }
}

function buildExportUrl(){
  const params = new URLSearchParams();
  if(activeFilters.centro) params.set('centro', activeFilters.centro);
  if(activeFilters.oferta) params.set('oferta', activeFilters.oferta);
  if(activeFilters.estado) params.set('estado', activeFilters.estado);
  if(activeFilters.tipo) params.set('tipo', activeFilters.tipo);
  if(activeFilters.nivel) params.set('nivel', activeFilters.nivel);
  if(activeFilters.periodo) params.set('periodo', activeFilters.periodo);
  if(activeSearch) params.set('search', activeSearch);
  const q = params.toString();
  return `${API_BASE}/fichas/export${q ? `?${q}` : ''}`;
}

function getFilenameFromDisposition(contentDisposition){
  if(!contentDisposition) return '';
  const match = contentDisposition.match(/filename="?([^";]+)"?/i);
  return match && match[1] ? match[1] : '';
}

async function exportFilteredExcel(){
  const filteredNow = getFilteredItems();
  if(filteredNow.length === 0){
    alert('No hay registros para exportar con los filtros actuales.');
    return;
  }

  const url = buildExportUrl();
  setStatus('Generando Excel de registros filtrados...');
  showLoading();
  disableActions(true);
  try{
    const resp = await fetch(url);
    if(!resp.ok){
      const data = await resp.json().catch(()=>null);
      const msg = data && data.detail ? data.detail : (resp.status + ' ' + resp.statusText);
      throw new Error(msg);
    }
    const blob = await resp.blob();
    const disposition = resp.headers.get('Content-Disposition') || '';
    const filename = getFilenameFromDisposition(disposition) || 'fichas_export.xlsx';

    const objectUrl = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = objectUrl;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(objectUrl);

    setStatus(`Excel exportado (${filteredNow.length} registros).`);
    showToast('Excel exportado correctamente.', 'success');
  }catch(e){
    setStatus('Error exportando: ' + e.message);
    showToast('Error exportando: ' + e.message, 'danger');
    console.error(e);
  }finally{
    hideLoading();
    disableActions(false);
  }
}

function getFilteredItems(){
  return items.filter(row => {
    if(activeFilters.centro){
      const v = (row.centro_formacion || '').toString().trim();
      if(v.toLowerCase() !== activeFilters.centro.toLowerCase()) return false;
    }
    if(activeFilters.oferta){
      const v = (row.oferta || '').toString();
      if(v !== activeFilters.oferta) return false;
    }
    if(activeFilters.estado){
      const v = (row.estado_ficha || '').toString().trim();
      if(v.toLowerCase() !== activeFilters.estado.toLowerCase()) return false;
    }
    if(activeFilters.tipo){
      const v = (row.tipo || '').toString().trim();
      if(v.toLowerCase() !== activeFilters.tipo.toLowerCase()) return false;
    }
    if(activeFilters.nivel){
      const v = (row.nivel_formacion || '').toString().trim();
      if(v.toLowerCase() !== activeFilters.nivel.toLowerCase()) return false;
    }
    if(activeFilters.periodo){
      const v = (row.periodo || '').toString().trim();
      if(v !== activeFilters.periodo) return false;
    }
    if(activeSearch){
      const denom = (row.denominacion_programa || '').toString().toLowerCase();
      if(!denom.includes(activeSearch.toLowerCase())) return false;
    }
    return true;
  });
}

function render(){
  // Aplicar filtros cliente antes de paginar
  const filtered = getFilteredItems();
  updateTotalsSummary(filtered);
  const totalPages = Math.max(1, Math.ceil(filtered.length / perPage));
  if(page < 1) page = 1;
  if(page > totalPages) page = totalPages;
  const start = (page-1)*perPage;
  const slice = filtered.slice(start, start+perPage);
  const thead = document.querySelector('#table thead');
  const tbody = document.querySelector('#table tbody');
  thead.innerHTML = '';
  tbody.innerHTML = '';
  if(slice.length === 0){ thead.innerHTML = '<tr><th>No hay datos</th></tr>'; return; }
  // Excluir `cod_municipio` y `cod_regional` de la vista para dejar más espacio (no se eliminan de la base de datos)
  const keys = Object.keys(slice[0]).filter(k => k !== 'cod_municipio' && k !== 'cod_regional' && k !== 'perfil_ingreso' && k !== 'cod_centro');
  const trh = document.createElement('tr');
  // Construir encabezados; en la columna denominacion_programa insertamos
  // un buscador en tiempo real dentro de la celda del encabezado, que
  // comparte estado con el buscador de la sección de filtros.
  trh.innerHTML = keys.map(k => {
    if(k === 'denominacion_programa'){
      return `
        <th>
          ${friendlyLabel(k)}
          <div class="mt-1">
            <input id="searchInputHeader" class="form-control form-control-sm" placeholder="Buscar programa..." list="searchSuggestions" autocomplete="off" />
          </div>
        </th>`;
    }
    return `<th>${friendlyLabel(k)}</th>`;
  }).join('');
  trh.innerHTML += '<th>Acciones</th>';
  thead.appendChild(trh);
  // Re-vincular el buscador del encabezado ahora que existe
  const headerSearch = document.getElementById('searchInputHeader');
  if(headerSearch){
    headerSearch.value = activeSearch || '';
    headerSearch.addEventListener('input', ev => {
      const v = ev.target.value || '';
      activeSearch = v;
      lastSearchSource = 'header';
      // Sincronizar con el buscador de filtros
      const filterSearch = document.getElementById('searchInput');
      if(filterSearch && filterSearch.value !== v) filterSearch.value = v;
      updateSearchSuggestions(v);
      page = 1;
      render();
    });
    // Actualizar sugerencias con el valor actual
    updateSearchSuggestions(headerSearch.value || '');
    // Si la última interacción de búsqueda fue en el encabezado,
    // devolver el foco al input para permitir escribir seguido.
    if(lastSearchSource === 'header'){
      headerSearch.focus();
      const len = headerSearch.value.length;
      try{ headerSearch.setSelectionRange(len, len); }catch(e){}
    }
  }
  slice.forEach((r,i)=>{
    const tr = document.createElement('tr');
    const trafficClass = getTrafficClass(r);
    if(trafficClass) tr.classList.add(trafficClass);
    const cells = keys.map(k=>`<td>${escapeHtml(r[k])}</td>`).join('');
    const editLink = `edit.html?cod_ficha=${encodeURIComponent(r.cod_ficha)}`;
    const actionCell = `<td>
      <a class="btn btn-sm btn-primary me-1" href="${editLink}">Editar</a>
      <button class="btn btn-sm btn-danger delete-btn" data-cod="${escapeHtml(r.cod_ficha)}">Eliminar</button>
    </td>`;
    tr.innerHTML = cells + actionCell;
    tbody.appendChild(tr);
  });
  document.getElementById('pageInfo').innerText = `Página ${page} de ${totalPages}`;
  const prevPageNum = document.getElementById('prevPageNum');
  const currentPageNum = document.getElementById('currentPageNum');
  const nextPageNum = document.getElementById('nextPageNum');
  if(prevPageNum) prevPageNum.innerText = page > 1 ? String(page - 1) : '-';
  if(currentPageNum) currentPageNum.innerText = String(page);
  if(nextPageNum) nextPageNum.innerText = page < totalPages ? String(page + 1) : '-';

  const prevBtn = document.getElementById('prev');
  const nextBtn = document.getElementById('next');
  const firstBtn = document.getElementById('first');
  const lastBtn = document.getElementById('last');
  if(prevBtn) prevBtn.disabled = page <= 1;
  if(firstBtn) firstBtn.disabled = page <= 1;
  if(nextBtn) nextBtn.disabled = page >= totalPages;
  if(lastBtn) lastBtn.disabled = page >= totalPages;
  // Mostrar conteo dinámico: cuántos se muestran y total
  const total = items.length || 0;
  const anyFilter = Boolean(activeFilters.centro || activeFilters.oferta || activeFilters.estado || activeFilters.tipo || activeFilters.nivel || activeFilters.periodo);
  if(anyFilter){
    setStatus(`Mostrando ${filtered.length} de ${total} registros (filtrados)`);
  } else {
    setStatus(`Mostrando ${filtered.length} registros`);
  }
}

function toNumber(v){
  if(v === null || v === undefined) return 0;
  if(typeof v === 'number') return Number.isFinite(v) ? v : 0;
  const cleaned = String(v).trim().replace(/\./g, '').replace(',', '.');
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : 0;
}

function getTrafficClass(row){
  const cupo = toNumber(row.cupo);
  const inscritos = toNumber(row.inscritos_primera_opcion) + toNumber(row.inscritos_segunda_opcion);
  const diff = inscritos - cupo;
  if(diff > 10) return 'row-traffic-green';
  if(diff >= 1 && diff <= 10) return 'row-traffic-yellow';
  return 'row-traffic-red';
}

function formatNumber(n){
  return new Intl.NumberFormat('es-CO').format(n || 0);
}

function updateTotalsSummary(rows){
  const totals = rows.reduce((acc, r) => {
    acc.cupos += toNumber(r.cupo);
    acc.primera += toNumber(r.inscritos_primera_opcion);
    acc.segunda += toNumber(r.inscritos_segunda_opcion);
    return acc;
  }, { cupos: 0, primera: 0, segunda: 0 });

  const elRows = document.getElementById('sumRegistros');
  const elCupos = document.getElementById('sumCupos');
  const elPrimera = document.getElementById('sumPrimera');
  const elSegunda = document.getElementById('sumSegunda');

  if(elRows) elRows.textContent = formatNumber(rows.length);
  if(elCupos) elCupos.textContent = formatNumber(totals.cupos);
  if(elPrimera) elPrimera.textContent = formatNumber(totals.primera);
  if(elSegunda) elSegunda.textContent = formatNumber(totals.segunda);
}

// Delete modal handling
let deleteModalInstance = null;
function openDeleteModal(cod){
  const modalEl = document.getElementById('confirmDeleteModal');
  if(!modalEl) return;
  // fetch ficha details
  fetch(`${API_BASE}/fichas/${encodeURIComponent(cod)}`)
    .then(r => { if(!r.ok) throw new Error('No se pudo obtener los datos'); return r.json(); })
    .then(data => {
      const details = [];
      const fields = ['cod_ficha','denominacion_programa','centro_formacion','periodo','oferta','tipo','cupo','inscritos_primera_opcion','inscritos_segunda_opcion'];
      fields.forEach(f => { if(f in data){ details.push(`${friendlyLabel(f)}: ${data[f]}`); } });
      const dd = document.getElementById('deleteDetails');
      if(dd) dd.innerText = details.join('\n');
      const btn = document.getElementById('confirmDeleteBtn');
      if(btn) btn.dataset.cod = cod;
      deleteModalInstance = new bootstrap.Modal(modalEl);
      deleteModalInstance.show();
    })
    .catch(e => { showToast('Error obteniendo ficha: '+e.message, 'danger'); console.error(e); });
}

document.addEventListener('click', ev => {
  const t = ev.target;
  if(t && t.classList && t.classList.contains('delete-btn')){
    const cod = t.getAttribute('data-cod');
    if(cod) openDeleteModal(cod);
  }
});

document.getElementById && (()=>{
  const confirmBtn = document.getElementById('confirmDeleteBtn');
  if(confirmBtn){
    confirmBtn.addEventListener('click', async ()=>{
      const cod = confirmBtn.dataset.cod;
      if(!cod) return;
      try{
        const resp = await fetch(`${API_BASE}/fichas/${encodeURIComponent(cod)}`, { method: 'DELETE' });
        if(resp.status === 204 || resp.ok){
          // Actualizar UI inmediatamente: eliminar del array local y re-renderizar
          const codNum = isNaN(Number(cod)) ? cod : Number(cod);
          items = items.filter(it => String(it.cod_ficha) !== String(codNum));
          lastTotalCount = items.length;
          showToast('Registro eliminado', 'success');
          if(deleteModalInstance) deleteModalInstance.hide();
          render();
          // Refrescar en segundo plano para mantener sincronía con el servidor
          loadAll().catch(()=>{});
        } else {
          const j = await resp.json().catch(()=>null);
          const msg = j && j.detail ? j.detail : resp.statusText;
          showToast('Error al eliminar: '+msg, 'danger');
        }
      }catch(e){ showToast('Error al eliminar: '+e.message, 'danger'); console.error(e); }
    });
  }
})();

function setStatus(s){
  const el = document.getElementById('status');
  if(el) el.innerText = s || '';
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

function escapeHtml(v){ if(v === null || v === undefined) return ''; return String(v).replace(/[&<>\"]/g, c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c])); }

function friendlyLabel(key){
  if(!key) return '';
  let s = String(key).replace(/_/g,' ');
  s = s.replace(/\bcod\b/gi, 'código');
  s = s.replace(/\bdenominacion\b/gi, 'denominación');
  s = s.replace(/\bformacion\b/gi, 'formación');
  s = s.replace(/\bcentro\b/gi, 'centro');
  s = s.replace(/\bmunicipio\b/gi, 'municipio');
  s = s.replace(/\bregional\b/gi, 'regional');
  s = s.replace(/\bficha\b/gi, 'ficha');
  s = s.replace(/\bprograma\b/gi, 'programa');
  s = s.replace(/\binscritos\b/gi, 'inscritos');
  s = s.replace(/\bprimera\b/gi, 'primera');
  s = s.replace(/\bsegunda\b/gi, 'segunda');
  s = s.replace(/\boferta\b/gi, 'oferta');
  s = s.replace(/\btipo\b/gi, 'tipo');
  s = s.replace(/\bperiodo\b/gi, 'periodo');
  s = s.replace(/\bcupo\b/gi, 'cupo');
  s = s.split(' ').map(w => w.charAt(0).toUpperCase() + w.slice(1)).join(' ');
  return s;
}

function readFilters(){
  const fc = document.getElementById('filterCentro');
  const fo = document.getElementById('filterOferta');
  const ft = document.getElementById('filterTipo');
  const fn = document.getElementById('filterNivel');
  const fe = document.getElementById('filterEstado');
  const fp = document.getElementById('filterPeriodo');
  activeFilters.centro = fc ? fc.value.trim() : '';
  activeFilters.oferta = fo ? fo.value : '';
  activeFilters.tipo = ft ? ft.value : '';
  activeFilters.nivel = fn ? fn.value : '';
  activeFilters.estado = fe ? fe.value.trim() : '';
  activeFilters.periodo = fp ? fp.value : '';
  // Leer también la búsqueda global para mantener consistencia al aplicar/limpiar filtros
  const s = document.getElementById('searchInput');
  activeSearch = s ? (s.value || '').toString().trim() : '';
}

function populateFilterOptions(){
  // rellenar selects con valores únicos ordenados
  const centros = new Set();
  const estados = new Set();
  const tipos = new Set();
  const niveles = new Set();
  const periodos = new Set();
  items.forEach(r => { if(r.centro_formacion) centros.add(String(r.centro_formacion).trim()); if(r.estado_ficha) estados.add(String(r.estado_ficha).trim()); if(r.nivel_formacion) niveles.add(String(r.nivel_formacion).trim()); });
  items.forEach(r => { if(r.tipo) tipos.add(String(r.tipo).trim()); });
  items.forEach(r => {
    if(r.periodo !== null && r.periodo !== undefined && String(r.periodo).trim() !== '') {
      periodos.add(String(r.periodo).trim());
    }
  });
  const centroArr = Array.from(centros).sort((a,b)=>a.localeCompare(b));
  const estadoArr = Array.from(estados).sort((a,b)=>a.localeCompare(b));
  const tipoArr = Array.from(tipos).sort((a,b)=>a.localeCompare(b));
  const nivelArr = Array.from(niveles).sort((a,b)=>a.localeCompare(b));
  const periodoArr = Array.from(periodos).sort((a,b)=>Number(b)-Number(a));

  const sc = document.getElementById('filterCentro');
  const se = document.getElementById('filterEstado');
  const st = document.getElementById('filterTipo');
  const sn = document.getElementById('filterNivel');
  const sp = document.getElementById('filterPeriodo');
  // limpiar manteniendo la primera opción (todos)
  if(sc){ sc.options.length = 1; centroArr.forEach(v=>{ const o = document.createElement('option'); o.value = v; o.text = v; sc.appendChild(o); }); }
  if(se){ se.options.length = 1; estadoArr.forEach(v=>{ const o = document.createElement('option'); o.value = v; o.text = v; se.appendChild(o); }); }
  if(st){ st.options.length = 1; tipoArr.forEach(v=>{ const o = document.createElement('option'); o.value = v; o.text = v; st.appendChild(o); }); }
  if(sn){ sn.options.length = 1; nivelArr.forEach(v=>{ const o = document.createElement('option'); o.value = v; o.text = v; sn.appendChild(o); }); }
  if(sp){ sp.options.length = 1; periodoArr.forEach(v=>{ const o = document.createElement('option'); o.value = v; o.text = v; sp.appendChild(o); }); }
}

// bulk-update UI removed — function applyUpdateToVisible deleted

// UX helpers: spinner, toasts, disable buttons
function showLoading(){
  const s = document.getElementById('loadingSpinner'); if(s) s.style.display = 'block';
}
function hideLoading(){
  const s = document.getElementById('loadingSpinner'); if(s) s.style.display = 'none';
}
function disableActions(dis){
  const ids = ['addUploadBlockBtn','uploadAllBtn','filterBtn','clearFiltersBtn','exportExcelBtn','first','prev','next','last'];
  ids.forEach(id=>{ const el = document.getElementById(id); if(el) el.disabled = dis; });
  document.querySelectorAll('.remove-upload-block').forEach(btn => { btn.disabled = dis; });
}
function showToast(message, type='info'){ // type: 'success'|'danger'|'info'
  const container = document.getElementById('toastContainer');
  if(!container) return;
  const div = document.createElement('div');
  div.className = `alert alert-${type} shadow-sm`;
  div.style.minWidth = '220px';
  div.style.marginTop = '6px';
  div.style.opacity = '0.95';
  div.innerText = message;
  container.appendChild(div);
  setTimeout(()=>{ div.style.transition = 'opacity 400ms'; div.style.opacity='0'; setTimeout(()=>div.remove(),450); }, 3500);
}
