const API_BASE = 'http://127.0.0.1:8000';

function escapeHtml(v) {
  if (v === null || v === undefined) return '';
  return String(v).replace(/[&<>\"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}

function showStatus(message, type = 'secondary') {
  const el = document.getElementById('status');
  if (!el) return;
  el.innerHTML = `<div class="alert alert-${type} py-2 mb-0">${escapeHtml(message)}</div>`;
}

function setProgress(percent) {
  const container = document.getElementById('uploadProgressContainer');
  const bar = document.getElementById('uploadProgressBar');
  if (!container || !bar) return;
  container.style.display = 'block';
  const value = Math.max(0, Math.min(100, Number(percent) || 0));
  bar.style.width = `${value}%`;
  bar.setAttribute('aria-valuenow', String(value));
}

function hideProgress() {
  const container = document.getElementById('uploadProgressContainer');
  if (container) container.style.display = 'none';
  setProgress(0);
}

function renderCatalogo(rows) {
  const tbody = document.getElementById('catalogoTableBody');
  if (!tbody) return;


  if (!rows.length) {
    tbody.innerHTML = '<tr><td colspan="6" class="text-center text-muted py-4">Sin registros</td></tr>';
    return;
  }

  tbody.innerHTML = rows.map((row) => `
    <tr>
      <td><strong>${escapeHtml(row.cod_ver)}</strong></td>
      <td>${escapeHtml(row.prf_denominacion)}</td>
      <td>${escapeHtml(row.nivel_de_formacion)}</td>
      <td>${escapeHtml(row.prf_duracion_maxima || '—')}</td>
      <td>${escapeHtml(row.prf_dur_etapa_lectiva || '—')}</td>
      <td>${escapeHtml(row.prf_dur_etapa_prod || '—')}</td>
    </tr>
  `).join('');
}

let catalogoCurrentPage = 1;
let catalogoTotalPages = 1;
let catalogoPerPage = 25;

async function loadCatalogo(page = 1) {
  const search = (document.getElementById('catalogoSearch')?.value || '').trim();
  // Sanear y validar 'page' para evitar NaN que genera 422 en el backend
  page = Number(page);
  if (!Number.isFinite(page) || page <= 0) {
    page = 1;
  } else {
    page = Math.floor(page);
  }
  catalogoCurrentPage = page;
  
  const params = new URLSearchParams();
  const selectedNivel = (document.getElementById('filterNivelFormacion')?.value || '').trim();
  params.set('page', String(catalogoCurrentPage));
  params.set('per_page', String(catalogoPerPage));
  if (search) params.set('search', search);
  if (selectedNivel) params.set('nivel', selectedNivel);

  try {
    showStatus('Cargando catalogo...', 'info');
    const resp = await fetch(`${API_BASE}/catalogo?${params.toString()}`);
    const data = await resp.json().catch(() => null);
    if (!resp.ok) {
      const msg = data && data.detail ? data.detail : `${resp.status} ${resp.statusText}`;
      throw new Error(msg);
    }
    
    let items = Array.isArray(data?.items) ? data.items : [];

    // poblar select de niveles usando distinct_niveles devuelto por el backend
    try {
      const nivelSelect = document.getElementById('filterNivelFormacion');
      if (nivelSelect && Array.isArray(data?.distinct_niveles)) {
        // limpiar opciones excepto la primera (Todos)
        const keepFirst = nivelSelect.options[0];
        nivelSelect.innerHTML = '';
        nivelSelect.appendChild(keepFirst);
        data.distinct_niveles.forEach(v => {
          if (v) {
            const opt = document.createElement('option'); opt.value = v; opt.textContent = v; nivelSelect.appendChild(opt);
          }
        });
        // si el usuario ya tenía un valor seleccionado, mantenerlo si existe
        const prev = (selectedNivel || '');
        if (prev) { nivelSelect.value = prev; }
      }
    } catch (e) { console.warn(e); }

    // Renderizar items tal cual vienen del servidor (server-side paging)
    renderCatalogo(items);

    // mostrar total real de la base de datos (no el tamaño de la página)
    const total = Number.isFinite(Number(data?.total)) ? Number(data.total) : 0;
    const totalEl = document.getElementById('catalogoTotal');
    if (totalEl) totalEl.textContent = String(total);
    catalogoTotalPages = Math.ceil(total / catalogoPerPage) || 1;
    
    const pageEl = document.getElementById('catalogoCurrentPage');
    const pagesEl = document.getElementById('catalogoTotalPages');
    if (pageEl) pageEl.textContent = String(catalogoCurrentPage);
    if (pagesEl) pagesEl.textContent = String(catalogoTotalPages);
    
    updatePaginationButtons();
    showStatus(`Catalogo listo. Fecha de corte actual: ${data?.fecha_corte || 'sin definir'}.`, 'success');
  } catch (err) {
    console.error(err);
    renderCatalogo([]);
    showStatus(`Error al cargar catalogo: ${err.message}`, 'danger');
  }
}

function updatePaginationButtons() {
  const prevBtn = document.getElementById('catalogoPrevBtn');
  const nextBtn = document.getElementById('catalogoNextBtn');
  
  if (prevBtn) {
    if (catalogoCurrentPage <= 1) {
      prevBtn.parentElement.classList.add('disabled');
      prevBtn.disabled = true;
    } else {
      prevBtn.parentElement.classList.remove('disabled');
      prevBtn.disabled = false;
    }
  }
  
  if (nextBtn) {
    if (catalogoCurrentPage >= catalogoTotalPages) {
      nextBtn.parentElement.classList.add('disabled');
      nextBtn.disabled = true;
    } else {
      nextBtn.parentElement.classList.remove('disabled');
      nextBtn.disabled = false;
    }
  }
}

async function uploadCatalogoExcel() {
  const input = document.getElementById('catalogoFile');
  const files = input && input.files ? Array.from(input.files) : [];
  if (!files.length) {
    alert('Selecciona un archivo Excel primero.');
    return;
  }

  const fechaManual = (document.getElementById('catalogoFechaCorte')?.value || '').trim();
  const fd = new FormData();
  fd.append('file', files[0]);
  if (fechaManual) {
    fd.append('fecha_corte_manual', fechaManual);
  }

  try {
    setProgress(10);
    showStatus('Subiendo Excel de catalogo...', 'info');
    const resp = await fetch(`${API_BASE}/catalogo/upload-excel`, {
      method: 'POST',
      body: fd,
    });
    setProgress(75);
    const data = await resp.json().catch(() => null);
    if (!resp.ok) {
      const msg = data && data.detail ? data.detail : `${resp.status} ${resp.statusText}`;
      throw new Error(msg);
    }
    setProgress(100);
    showStatus(`Subida completada. Fecha de corte: ${data.fecha_corte}. Filas procesadas: ${data.inserted}.`, 'success');
    await loadCatalogo();
  } catch (err) {
    console.error(err);
    showStatus(`Error al subir catalogo: ${err.message}`, 'danger');
  } finally {
    setTimeout(hideProgress, 600);
  }
}

document.addEventListener('DOMContentLoaded', () => {
  const uploadBtn = document.getElementById('uploadCatalogoBtn');
  if (uploadBtn) uploadBtn.addEventListener('click', uploadCatalogoExcel);

  const reloadBtn = document.getElementById('reloadCatalogoBtn');
  if (reloadBtn) reloadBtn.addEventListener('click', loadCatalogo);

  const searchBtn = document.getElementById('searchCatalogoBtn');
  if (searchBtn) searchBtn.addEventListener('click', loadCatalogo);

  const searchInput = document.getElementById('catalogoSearch');
  if (searchInput) {
    searchInput.addEventListener('keydown', (event) => {
      if (event.key === 'Enter') {
        event.preventDefault();
        loadCatalogo();
      }
    });
  }

  const toggleBtn = document.getElementById('toggleTableBtn');
  if (toggleBtn) {
    toggleBtn.addEventListener('click', () => {
      const container = document.getElementById('catalogoTableContainer');
      if (container) {
        container.classList.toggle('show');
        toggleBtn.textContent = container.classList.contains('show') ? 'Ocultar tabla' : 'Mostrar tabla';
        // sincronizar visibilidad de la paginación
        const pagNav = document.getElementById('catalogoPagination');
        if (pagNav) {
          pagNav.style.display = container.classList.contains('show') ? 'block' : 'none';
        }
      }
    });
  }

  const prevBtn = document.getElementById('catalogoPrevBtn');
  if (prevBtn) {
    prevBtn.addEventListener('click', () => {
      if (catalogoCurrentPage > 1) {
        loadCatalogo(catalogoCurrentPage - 1);
      }
    });
  }

  const nextBtn = document.getElementById('catalogoNextBtn');
  if (nextBtn) {
    nextBtn.addEventListener('click', () => {
      if (catalogoCurrentPage < catalogoTotalPages) {
        loadCatalogo(catalogoCurrentPage + 1);
      }
    });
  }

  const nivelSelect = document.getElementById('filterNivelFormacion');
  if (nivelSelect) {
    nivelSelect.addEventListener('change', () => {
      catalogoCurrentPage = 1;
      loadCatalogo(1);
    });
  }

  loadCatalogo();
});
