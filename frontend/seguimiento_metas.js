const API_BASE = 'http://127.0.0.1:8000';

function escapeHtml(v) {
  if (v === null || v === undefined) return '';
  return String(v).replace(/[&<>\"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}

function updateThemeButton(theme) {
  const btn = document.getElementById('themeToggle');
  if (btn) {
    btn.textContent = theme === 'dark' ? '☀️ Modo claro' : '🌙 Modo oscuro';
  }
}

function showStatus(message, type = 'secondary') {
  const el = document.getElementById('status');
  if (!el) return;
  el.innerHTML = `<div class="alert alert-${type} py-2 mb-0">${escapeHtml(message)}</div>`;
}

function showOfertaStatus(message, type = 'secondary') {
  const el = document.getElementById('ofertaStatus');
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
      <td><small><span class="badge ${row.tipo_programa ? 'bg-info' : 'bg-secondary'}">${escapeHtml(row.tipo_programa || 'Sin asignar')}</span></small></td>
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

// ===== MINI MÓDULO: AGREGAR PROGRAMAS EN CATÁLOGO =====

function showAgregarProgramasStatus(message, type = 'secondary') {
  const el = document.getElementById('agregarProgramasStatus');
  if (!el) return;
  el.innerHTML = `<div class="alert alert-${type} py-2 mb-0">${escapeHtml(message)}</div>`;
}

function setAgregarProgramasProgress(percent) {
  const container = document.getElementById('agregarProgramasProgressContainer');
  const bar = document.getElementById('agregarProgramasProgressBar');
  if (!container || !bar) return;
  container.style.display = 'block';
  const value = Math.max(0, Math.min(100, Number(percent) || 0));
  bar.style.width = `${value}%`;
  bar.setAttribute('aria-valuenow', String(value));
}

function hideAgregarProgramasProgress() {
  const container = document.getElementById('agregarProgramasProgressContainer');
  if (container) container.style.display = 'none';
  setAgregarProgramasProgress(0);
}

async function uploadAgregarProgramas() {
  const tipoProgramaSelect = document.getElementById('tipoProgramaSelect');
  const input = document.getElementById('agregarProgramasFile');
  
  const tipoPrograma = tipoProgramaSelect?.value?.trim() || '';
  const files = input && input.files ? Array.from(input.files) : [];
  
  if (!tipoPrograma) {
    alert('Selecciona un tipo de programa primero.');
    return;
  }
  
  if (!files.length) {
    alert('Selecciona un archivo Excel primero.');
    return;
  }

  const fd = new FormData();
  fd.append('file', files[0]);
  fd.append('tipo_programa', tipoPrograma);

  try {
    setAgregarProgramasProgress(10);
    showAgregarProgramasStatus('Procesando códigos...', 'info');
    
    const resp = await fetch(`${API_BASE}/catalogo/agregar-programas`, {
      method: 'POST',
      body: fd,
    });
    
    setAgregarProgramasProgress(75);
    const data = await resp.json().catch(() => null);
    
    if (!resp.ok) {
      const msg = data && data.detail ? data.detail : `${resp.status} ${resp.statusText}`;
      throw new Error(msg);
    }
    
    setAgregarProgramasProgress(100);
    const message = `✓ Completado: ${data.updated} programas actualizados, ${data.not_found} no encontrados. Tipo: ${escapeHtml(data.tipo_programa)}.`;
    showAgregarProgramasStatus(message, 'success');
    
    input.value = '';
    await loadCatalogo();
  } catch (err) {
    console.error(err);
    showAgregarProgramasStatus(`Error: ${err.message}`, 'danger');
  } finally {
    setTimeout(hideAgregarProgramasProgress, 600);
  }
}

// Agregar event listener para Agregar Programas
document.getElementById('uploadAgregarProgramasBtn')?.addEventListener('click', uploadAgregarProgramas);

// ===== FUNCIONES PARA REGISTRO CALIFICADO =====

function showRegistroStatus(message, type = 'secondary') {
  const el = document.getElementById('registroStatus');
  if (!el) return;
  el.innerHTML = `<div class="alert alert-${type} py-2 mb-0">${escapeHtml(message)}</div>`;
}

function setRegistroProgress(percent) {
  const container = document.getElementById('registroProgressContainer');
  const bar = document.getElementById('registroProgressBar');
  if (!container || !bar) return;
  container.style.display = 'block';
  const value = Math.max(0, Math.min(100, Number(percent) || 0));
  bar.style.width = `${value}%`;
  bar.setAttribute('aria-valuenow', String(value));
}

function hideRegistroProgress() {
  const container = document.getElementById('registroProgressContainer');
  if (container) container.style.display = 'none';
  setRegistroProgress(0);
}

function renderRegistroTable(rows) {
  const tbody = document.getElementById('registroTableBody');
  if (!tbody) return;

  if (!rows.length) {
    tbody.innerHTML = '<tr><td colspan="34" class="text-center text-muted py-4">Sin registros</td></tr>';
    return;
  }

  tbody.innerHTML = rows.map((row) => `
    <tr>
      <td><small>${escapeHtml(row.id || '—')}</small></td>
      <td><small>${escapeHtml(row.proceso || '—')}</small></td>
      <td><small>${escapeHtml(row.tipo_tramite || '—')}</small></td>
      <td><small>${row.fecha_radicado ? new Date(row.fecha_radicado).toLocaleDateString('es-CO') : '—'}</small></td>
      <td><small>${escapeHtml(row.numero_resolucion || '—')}</small></td>
      <td><small>${row.fecha_resolucion ? new Date(row.fecha_resolucion).toLocaleDateString('es-CO') : '—'}</small></td>
      <td><small>${escapeHtml(row.resuelve || '—')}</small></td>
      <td><small>${escapeHtml(row.decreto_ampara || '—')}</small></td>
      <td><small>${escapeHtml(row.snies || '—')}</small></td>
      <td><small>${escapeHtml(row.cobertura || '—')}</small></td>
      <td><small>${escapeHtml(row.resolucion_ampara_programa || '—')}</small></td>
      <td><small>${escapeHtml(row.resolucion_ampara || '—')}</small></td>
      <td><small>${escapeHtml(row.resolucion_ampara_fecha || '—')}</small></td>
      <td><small>${row.fecha_vencimiento ? new Date(row.fecha_vencimiento).toLocaleDateString('es-CO') : '—'}</small></td>
      <td><small>${escapeHtml(row.vigencia_rc || '—')}</small></td>
      <td><small>${escapeHtml(row.cod_programa || '—')}</small></td>
      <td><small>${escapeHtml(row.version || '—')}</small></td>
      <td><small><strong>${escapeHtml(row.nombre_programa || '—')}</strong></small></td>
      <td><small>${escapeHtml(row.nivel_formacion || '—')}</small></td>
      <td><small>${escapeHtml(row.red_conocimiento || '—')}</small></td>
      <td><small>${escapeHtml(row.modalidad || '—')}</small></td>
      <td><small>${escapeHtml(row.centro_formacion || '—')}</small></td>
      <td><small>${escapeHtml(row.nombre_sede || '—')}</small></td>
      <td><small>${escapeHtml(row.tipo_sede || '—')}</small></td>
      <td><small>${escapeHtml(row.municipio || '—')}</small></td>
      <td><small>${escapeHtml(row.lugar_desarrollo || '—')}</small></td>
      <td><small>${escapeHtml(row.direccion || '—')}</small></td>
      <td><small>${escapeHtml(row.regional || '—')}</small></td>
      <td><small>${escapeHtml(row.nombre_regional || '—')}</small></td>
      <td><small>${escapeHtml(row.observaciones || '—')}</small></td>
      <td><small>${escapeHtml(row.clasificacion_tramite || '—')}</small></td>
      <td><small>${escapeHtml(row.aprendices_primer_cohorte || '—')}</small></td>
      <td><small>${escapeHtml(row.lugar_desarrollo_resolucion || '—')}</small></td>
      <td><small>${row.fecha_registro ? new Date(row.fecha_registro).toLocaleDateString('es-CO') : '—'}</small></td>
    </tr>
  `).join('');
}

async function loadRegistroData() {
  try {
    showRegistroStatus('Cargando...', 'info');
    const resp = await fetch(`${API_BASE}/registro-calificado/data`);
    const data = await resp.json().catch(() => null);
    
    if (!resp.ok) {
      const msg = data && data.detail ? data.detail : `${resp.status} ${resp.statusText}`;
      throw new Error(msg);
    }

    const items = Array.isArray(data?.items) ? data.items : [];
    const total = data?.total || items.length || 0;
    
    document.getElementById('registroTotal').textContent = total;
    renderRegistroTable(items);
    
    if (total === 0) {
      showRegistroStatus('Sin registros', 'warning');
    } else {
      showRegistroStatus(`✓ ${total} registros cargados`, 'success');
    }
  } catch (error) {
    showRegistroStatus(`Error: ${error.message}`, 'danger');
    console.error('Error:', error);
  }
}

async function uploadRegistroExcel() {
  const fileInput = document.getElementById('registroFile');
  if (!fileInput.files.length) {
    showRegistroStatus('Selecciona un archivo', 'warning');
    return;
  }

  const file = fileInput.files[0];
  const formData = new FormData();
  formData.append('file', file);

  try {
    showRegistroStatus('Subiendo...', 'info');
    setRegistroProgress(30);

    const resp = await fetch(`${API_BASE}/registro-calificado/upload-excel`, {
      method: 'POST',
      body: formData
    });

    setRegistroProgress(80);

    const data = await resp.json().catch(() => null);
    if (!resp.ok) {
      const msg = data && data.detail ? data.detail : `${resp.status} ${resp.statusText}`;
      throw new Error(msg);
    }

    const registrosInsertados = data.rows_processed || 0;
    showRegistroStatus(`✓ Se ingresaron ${registrosInsertados} registros`, 'success');
    
    fileInput.value = '';
    setRegistroProgress(100);
    
    setTimeout(() => {
      hideRegistroProgress();
      loadRegistroData();
    }, 500);
  } catch (error) {
    showRegistroStatus(`Error: ${error.message}`, 'danger');
    hideRegistroProgress();
  }
}

document.addEventListener('DOMContentLoaded', () => {
  // ===== INICIALIZACIÓN: Declarar todas las variables del DOM =====
  const catalogoSection = document.getElementById('catalogoSection');
  const registroSection = document.getElementById('registroSection');
  const ofertaSection = document.getElementById('ofertaSection');
  const consolidadoSection = document.getElementById('consolidadoSection');
  const pe04Section = document.getElementById('pe04Section');
  const subidaArchivosNavbar = document.getElementById('subidaArchivosNavbar');
  const seguimientoMetasSection = document.getElementById('seguimientoMetasModuleSection');
  const toggleSubidaArchivosBtn = document.getElementById('toggleSubidaArchivosSection');
  const toggleSeguimientoMetasBtn = document.getElementById('toggleSeguimientoMetasModuleBtn');

  // ===== Ocultar todas las secciones de subida de archivos por defecto =====
  if (catalogoSection) catalogoSection.style.display = 'none';
  if (registroSection) registroSection.style.display = 'none';
  if (ofertaSection) ofertaSection.style.display = 'none';
  if (consolidadoSection) consolidadoSection.style.display = 'none';
  if (pe04Section) pe04Section.style.display = 'none';

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

  // ===== FUNCIÓN ANTERIOR (DEPRECATED) - Reemplazada por toggles exclusivos =====
  // Las funciones syncSectionToggle y bindSectionToggle fueron reemplazadas por 
  // la lógica de toggles exclusivos para mejor UX

  // ===== TOGGLES EXCLUSIVOS PARA SUBMÓDULOS DE SUBIDA =====
  const submoduleButtons = {
    toggleSectionBtn: 'catalogoSection',
    toggleRegistroSection: 'registroSection',
    toggleOfertaSection: 'ofertaSection',
    toggleConsolidadoSection: 'consolidadoSection',
    togglePe04Section: 'pe04Section'
  };

  Object.entries(submoduleButtons).forEach(([buttonId, sectionId]) => {
    const button = document.getElementById(buttonId);
    const section = document.getElementById(sectionId);
    
    if (button && section) {
      button.addEventListener('click', () => {
        // Ocultar todos los demás submódulos
        Object.entries(submoduleButtons).forEach(([otherId, otherSectionId]) => {
          const otherBtn = document.getElementById(otherId);
          const otherSection = document.getElementById(otherSectionId);
          
          if (otherSection && otherSectionId !== sectionId) {
            otherSection.style.display = 'none';
            if (otherBtn) {
              otherBtn.classList.remove('active');
              otherBtn.setAttribute('aria-pressed', 'false');
            }
          }
        });
        
        // Mostrar el submódulo seleccionado
        section.style.display = 'block';
        button.classList.add('active');
        button.setAttribute('aria-pressed', 'true');
      });
    }
  });

  loadCatalogo();
  
  // Event listeners para Registro Calificado
  document.getElementById('uploadRegistroBtn')?.addEventListener('click', uploadRegistroExcel);
  document.getElementById('reloadRegistroBtn')?.addEventListener('click', loadRegistroData);

  const toggleRegistroTableBtn = document.getElementById('toggleRegistroTableBtn');
  const registroTableContainer = document.getElementById('registroTableContainer');
  if (toggleRegistroTableBtn) {
    toggleRegistroTableBtn.addEventListener('click', () => {
      if (registroTableContainer.style.display === 'none') {
        registroTableContainer.style.display = 'block';
        toggleRegistroTableBtn.textContent = 'Ocultar tabla';
      } else {
        registroTableContainer.style.display = 'none';
        toggleRegistroTableBtn.textContent = 'Mostrar tabla';
      }
    });
  }

  // ===== MODO OSCURO =====
  const themeToggle = document.getElementById('themeToggle');
  if (themeToggle) {
    const savedTheme = localStorage.getItem('theme') || 'light';
    document.documentElement.setAttribute('data-bs-theme', savedTheme);
    updateThemeButton(savedTheme);

    themeToggle.addEventListener('click', () => {
      const currentTheme = document.documentElement.getAttribute('data-bs-theme') || 'light';
      const newTheme = currentTheme === 'light' ? 'dark' : 'light';
      document.documentElement.setAttribute('data-bs-theme', newTheme);
      localStorage.setItem('theme', newTheme);
      updateThemeButton(newTheme);
    });
  }

  // Cargar datos de Registro Calificado automáticamente al cargar la página
  loadRegistroData();

  // ===== MÓDULO SEGUIMIENTO METAS: OFERTA =====
  function renderOfertaTable(rows) {
    const tbody = document.getElementById('ofertaTableBody');
    if (!tbody) return;

    if (!rows || !rows.length) {
      tbody.innerHTML = '<tr><td colspan="20" class="text-center text-muted py-4">Sin registros</td></tr>';
      return;
    }

    tbody.innerHTML = rows.map((row) => `
      <tr>
        <td><small>${escapeHtml(row.id)}</small></td>
        <td><small>${escapeHtml(row.codigo_centro || '')}</small></td>
        <td><small>${escapeHtml(row.tipo_oferta || '')}</small></td>
        <td><small>${escapeHtml(row.denominacion_formacion || '')}</small></td>
        <td><small>${escapeHtml(row.modalidad || '')}</small></td>
        <td><small>${escapeHtml(row.codigo_programa || '')}</small></td>
        <td><small>${escapeHtml(row.version_programa || '')}</small></td>
        <td><small>${escapeHtml(row.grupos || '')}</small></td>
        <td><small>${escapeHtml(row.cupos || '')}</small></td>
        <td><small>${escapeHtml(row.duracion_meses || '')}</small></td>
        <td><small>${escapeHtml(row.municipio || '')}</small></td>
        <td><small>${escapeHtml(row.sede || '')}</small></td>
        <td class="codigo-indicativa-col"><small>${escapeHtml(row.codigo_indicativa || '')}</small></td>
        <td><small>${escapeHtml(row.horario_formacion || '')}</small></td>
        <td><small>${escapeHtml(row.estrategia || '')}</small></td>
        <td><small>${escapeHtml(row.fecha_inicio || '')}</small></td>
        <td><small>${escapeHtml(row.fecha_fin || '')}</small></td>
        <td><small>${escapeHtml(row.oferta || '')}</small></td>
        <td><small>${escapeHtml(row.verificado || '')}</small></td>
        <td><button class="btn btn-sm btn-outline-primary editVerBtn" data-id="${escapeHtml(row.id)}">Editar</button></td>
      </tr>
    `).join('');

    // Añadir listeners a botones de editar
    Array.from(document.getElementsByClassName('editVerBtn')).forEach(btn => {
      btn.removeEventListener('click', handleEditVerClick);
      btn.addEventListener('click', handleEditVerClick);
    });
  }

  function handleEditVerClick(e) {
    const id = e.currentTarget.getAttribute('data-id');
    if (!id) return;
    const current = e.currentTarget.closest('tr')?.querySelectorAll('td')[18]?.textContent || '';
    const newVal = prompt('Establece verificado: (VERIFICADO / NO VERIFICADO / VERIFICACION MANUAL / REGISTRO VENCIDO). Dejar vacío para NULL', current.trim() || '');
    if (newVal === null) return; // cancel
    const upper = newVal.trim().toUpperCase();
    const normalized = upper === '' ? null : (upper === 'VERIFICADO' ? 'VERIFICADO' : upper === 'VERIFICACION MANUAL' ? 'VERIFICACION MANUAL' : upper === 'REGISTRO VENCIDO' ? 'REGISTRO VENCIDO' : 'NO VERIFICADO');
    updateVerificado(id, normalized);
  }

  async function updateVerificado(id, verificado) {
    try {
      showOfertaStatus('Actualizando verificado...', 'info');
      const resp = await fetch(`${API_BASE}/seguimiento-metas/update-verificado`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ id: Number(id), verificado: verificado }),
      });
      const data = await resp.json().catch(() => null);
      if (!resp.ok) throw new Error((data && data.detail) || `${resp.status} ${resp.statusText}`);
      showOfertaStatus('Verificado actualizado', 'success');
      await loadOfertaData();
    } catch (err) {
      console.error(err);
      showOfertaStatus(`Error actualizando verificado: ${err.message}`, 'danger');
    }
  }

  async function loadOfertaData() {
    try {
      showOfertaStatus('Cargando OFERTA...', 'info');
      const filterVerificado = (document.getElementById('filterOfertaVerificado')?.value || '').trim();
      const params = new URLSearchParams();
      if (filterVerificado) params.set('verificado', filterVerificado);
      const url = filterVerificado ? `${API_BASE}/seguimiento-metas/data?${params.toString()}` : `${API_BASE}/seguimiento-metas/data`;
      const resp = await fetch(url);
      const data = await resp.json().catch(() => null);
      if (!resp.ok) throw new Error((data && data.detail) || `${resp.status} ${resp.statusText}`);
      const items = Array.isArray(data?.items) ? data.items : [];
      renderOfertaTable(items);
      const totalEl = document.getElementById('ofertaTotal'); if (totalEl) totalEl.textContent = String(data?.total || items.length || 0);
      showOfertaStatus('OFERTA cargada', 'success');
    } catch (err) {
      console.error(err);
      showOfertaStatus(`Error al cargar OFERTA: ${err.message}`, 'danger');
      renderOfertaTable([]);
    }
  }

  async function uploadOfertaExcel() {
    const input = document.getElementById('ofertaFile');
    const files = input && input.files ? Array.from(input.files) : [];
    if (!files.length) { alert('Selecciona un archivo Excel primero.'); return; }
    const fd = new FormData(); fd.append('file', files[0]);
    try {
      setProgress(10);
      showOfertaStatus('Subiendo Excel de OFERTA...', 'info');
      const resp = await fetch(`${API_BASE}/seguimiento-metas/upload-oferta`, { method: 'POST', body: fd });
      setProgress(75);
      const data = await resp.json().catch(() => null);
      if (!resp.ok) throw new Error((data && data.detail) || `${resp.status} ${resp.statusText}`);
      setProgress(100);
      showOfertaStatus(`Subida completada. Filas procesadas: ${data.rows_processed || 0}.`, 'success');
      input.value = '';
      await loadOfertaData();
    } catch (err) {
      console.error(err);
      showOfertaStatus(`Error al subir OFERTA: ${err.message}`, 'danger');
    } finally {
      setTimeout(hideProgress, 600);
    }
  }

  // Agregar event listeners de OFERTA
  document.getElementById('uploadOfertaBtn')?.addEventListener('click', uploadOfertaExcel);
  document.getElementById('reloadOfertaBtn')?.addEventListener('click', loadOfertaData);
  document.getElementById('filterOfertaVerificado')?.addEventListener('change', loadOfertaData);

  const toggleOfertaTableBtn = document.getElementById('toggleOfertaTableBtn');
  const ofertaTableContainer = document.getElementById('ofertaTableContainer');
  if (toggleOfertaTableBtn) {
    toggleOfertaTableBtn.addEventListener('click', () => {
      if (!ofertaTableContainer) return;
      if (ofertaTableContainer.style.display === 'none') {
        ofertaTableContainer.style.display = 'block';
        toggleOfertaTableBtn.textContent = 'Ocultar tabla OFERTA';
      } else {
        ofertaTableContainer.style.display = 'none';
        toggleOfertaTableBtn.textContent = 'Mostrar tabla OFERTA';
      }
    });
  }

  // Cargar OFERTA inicialmente
  loadOfertaData();

  // ===== FUNCIONES PARA CONSOLIDADO COLEGIOS =====

  function showConsolidadoStatus(message, type = 'secondary') {
    const el = document.getElementById('consolidadoStatus');
    if (!el) return;
    el.innerHTML = `<div class="alert alert-${type} py-2 mb-0">${escapeHtml(message)}</div>`;
  }

  function setConsolidadoProgress(percent) {
    const container = document.getElementById('consolidadoProgressContainer');
    const bar = document.getElementById('consolidadoProgressBar');
    if (!container || !bar) return;
    container.style.display = 'block';
    const value = Math.max(0, Math.min(100, Number(percent) || 0));
    bar.style.width = `${value}%`;
    bar.setAttribute('aria-valuenow', String(value));
  }

  function hideConsolidadoProgress() {
    const container = document.getElementById('consolidadoProgressContainer');
    if (container) container.style.display = 'none';
    setConsolidadoProgress(0);
  }

  function renderConsolidadoTable(rows) {
    const tbody = document.getElementById('consolidadoTableBody');
    if (!tbody) return;

    if (!rows.length) {
      tbody.innerHTML = '<tr><td colspan="6" class="text-center text-muted py-4">Sin registros</td></tr>';
      return;
    }

    tbody.innerHTML = rows.map((row) => `
      <tr>
        <td><small>${escapeHtml(row.id || '—')}</small></td>
        <td><small>${escapeHtml(row.nombre_real_institucion || '—')}</small></td>
        <td><small>${escapeHtml(row.nombres_sofia_plus || '—')}</small></td>
        <td><small>${escapeHtml(row.municipio || '—')}</small></td>
        <td><small>${escapeHtml(row.clasificacion || '—')}</small></td>
        <td><small>${row.fecha_registro ? new Date(row.fecha_registro).toLocaleDateString('es-CO') : '—'}</small></td>
      </tr>
    `).join('');
  }

  async function loadConsolidadoData() {
    try {
      showConsolidadoStatus('Cargando Consolidado Colegios...', 'info');
      const resp = await fetch(`${API_BASE}/consolidado-colegios/data`);
      const data = await resp.json().catch(() => null);

      if (!resp.ok) {
        const msg = data && data.detail ? data.detail : `${resp.status} ${resp.statusText}`;
        throw new Error(msg);
      }

      const items = Array.isArray(data?.items) ? data.items : [];
      const total = data?.total || items.length || 0;

      document.getElementById('consolidadoTotal').textContent = total;
      renderConsolidadoTable(items);

      if (total === 0) {
        showConsolidadoStatus('Sin registros aún', 'warning');
      } else {
        showConsolidadoStatus(`✓ ${total} registros cargados`, 'success');
      }
    } catch (error) {
      showConsolidadoStatus(`Error: ${error.message}`, 'danger');
      renderConsolidadoTable([]);
    }
  }

  async function uploadConsolidadoExcel() {
    const input = document.getElementById('consolidadoFile');
    const files = input && input.files ? Array.from(input.files) : [];
    if (!files.length) {
      alert('Selecciona un archivo Excel primero.');
      return;
    }

    const fd = new FormData();
    fd.append('file', files[0]);

    try {
      setConsolidadoProgress(10);
      showConsolidadoStatus('Subiendo Excel de Consolidado Colegios...', 'info');
      const resp = await fetch(`${API_BASE}/consolidado-colegios/upload-excel`, {
        method: 'POST',
        body: fd,
      });
      setConsolidadoProgress(75);
      const data = await resp.json().catch(() => null);
      if (!resp.ok) {
        const msg = data && data.detail ? data.detail : `${resp.status} ${resp.statusText}`;
        throw new Error(msg);
      }
      setConsolidadoProgress(100);
      showConsolidadoStatus(`Subida completada. Registros insertados: ${data.inserted || 0}.`, 'success');
      input.value = '';
      await loadConsolidadoData();
    } catch (err) {
      console.error(err);
      showConsolidadoStatus(`Error al subir Consolidado Colegios: ${err.message}`, 'danger');
    } finally {
      setTimeout(hideConsolidadoProgress, 600);
    }
  }

  // Agregar event listeners de Consolidado Colegios
  document.getElementById('uploadConsolidadoBtn')?.addEventListener('click', uploadConsolidadoExcel);
  document.getElementById('reloadConsolidadoBtn')?.addEventListener('click', loadConsolidadoData);

  const toggleConsolidadoTableBtn = document.getElementById('toggleConsolidadoTableBtn');
  const consolidadoTableContainer = document.getElementById('consolidadoTableContainer');
  if (toggleConsolidadoTableBtn) {
    toggleConsolidadoTableBtn.addEventListener('click', () => {
      if (!consolidadoTableContainer) return;
      if (consolidadoTableContainer.style.display === 'none') {
        consolidadoTableContainer.style.display = 'block';
        toggleConsolidadoTableBtn.textContent = 'Ocultar tabla Consolidado Colegios';
      } else {
        consolidadoTableContainer.style.display = 'none';
        toggleConsolidadoTableBtn.textContent = 'Mostrar tabla Consolidado Colegios';
      }
    });
  }

  // Cargar Consolidado Colegios inicialmente
  loadConsolidadoData();

  // ===== FUNCIONES DE EXPORTACIÓN A EXCEL =====
  
  function downloadExcel(url, filename) {
    fetch(url)
      .then(response => {
        if (!response.ok) {
          throw new Error(`Error en la descarga: ${response.status}`);
        }
        return response.blob();
      })
      .then(blob => {
        const link = document.createElement('a');
        const urlBlob = window.URL.createObjectURL(blob);
        link.href = urlBlob;
        link.download = filename;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        window.URL.revokeObjectURL(urlBlob);
      })
      .catch(err => {
        console.error('Error descargando archivo:', err);
        alert(`Error al descargar archivo: ${err.message}`);
      });
  }

  // Event listeners para botones de exportación
  document.getElementById('exportCatalogoBtn')?.addEventListener('click', () => {
    downloadExcel(`${API_BASE}/catalogo/exportar-excel`, 'catalogo.xlsx');
  });

  document.getElementById('exportRegistroBtn')?.addEventListener('click', () => {
    downloadExcel(`${API_BASE}/registro-calificado/exportar-excel`, 'registro_calificado.xlsx');
  });

  document.getElementById('exportOfertaBtn')?.addEventListener('click', () => {
    downloadExcel(`${API_BASE}/oferta/exportar-excel`, 'oferta.xlsx');
  });

  document.getElementById('exportConsolidadoBtn')?.addEventListener('click', () => {
    downloadExcel(`${API_BASE}/consolidado-colegios/exportar-excel`, 'consolidado_colegios.xlsx');
  });

  // ===== FUNCIONES PARA APRENDICES RESUMEN (SEGUIMIENTO A LAS METAS) =====
  async function loadModalidades(centroCentro = '') {
    try {
      let url = `${API_BASE}/pe04-seguimiento/resumen-modalidades`;
      if (centroCentro) url += `?centro=${encodeURIComponent(centroCentro)}`;
      
      const response = await fetch(url);
      const data = await response.json();

      if (!response.ok) {
        throw new Error(data.detail || 'Error al cargar datos');
      }

      const items = data.items || [];
      const centros = data.centros_disponibles || [];
      
      // Actualizar dropdown de centros
      const centroSelect = document.getElementById('filtroModalidadesCentro');
      if (centroSelect && centros.length > 0) {
        const currentValue = centroSelect.value;
        centroSelect.innerHTML = '<option value="">-- Todos los centros --</option>' + 
          centros.map(c => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join('');
        centroSelect.value = currentValue;
      }
      
      const totalEl = document.getElementById('totalModalidades');
      if (totalEl) totalEl.textContent = data.total_aprendices || 0;

      const tbody = document.getElementById('modalidadesTableBody');
      if (!tbody) return;

      if (!items.length) {
        tbody.innerHTML = '<tr><td colspan="5" class="text-center text-muted py-4">Sin registros</td></tr>';
        return;
      }

      tbody.innerHTML = items.map((row) => {
        let badgeClass = 'bg-secondary';
        const clasificacion = (row.clasificacion_programa_especial || 'NA').toUpperCase();
        
        if (clasificacion === 'SENATEC') badgeClass = 'bg-primary';
        else if (clasificacion === 'ACME') badgeClass = 'bg-info';
        else if (clasificacion === 'SER CAMPESENA') badgeClass = 'bg-warning text-dark';
        else if (clasificacion === 'SER') badgeClass = 'bg-success';
        else if (clasificacion === 'BILINGUISMO') badgeClass = 'bg-danger';
        else if (clasificacion === 'CAMPESENA RADIAL') badgeClass = 'bg-teal';
        
        return `
          <tr>
            <td><strong>${escapeHtml(row.centro_formacion || 'N/A')}</strong></td>
            <td><span class="badge bg-dark">${escapeHtml(row.modalidad_formacion || 'N/A')}</span></td>
            <td><span class="badge ${badgeClass}">${escapeHtml(row.clasificacion_programa_especial || 'NA')}</span></td>
            <td><strong>${row.total_fichas || 0}</strong></td>
            <td><span class="badge bg-success" style="font-size: 0.95rem;">${row.total_aprendices || 0}</span></td>
          </tr>
        `;
      }).join('');
    } catch (error) {
      const tbody = document.getElementById('modalidadesTableBody');
      if (tbody) tbody.innerHTML = `<tr><td colspan="5" class="text-center text-danger py-4">Error: ${error.message}</td></tr>`;
    }
  }

  async function loadEspeciales(centroCentro = '') {
    try {
      let url = `${API_BASE}/pe04-seguimiento/resumen-especiales`;
      if (centroCentro) url += `?centro=${encodeURIComponent(centroCentro)}`;
      
      const response = await fetch(url);
      const data = await response.json();

      if (!response.ok) {
        throw new Error(data.detail || 'Error al cargar datos');
      }

      const items = data.items || [];
      const centros = data.centros_disponibles || [];
      
      // Actualizar dropdown de centros
      const centroSelect = document.getElementById('filtroEspecialesCentro');
      if (centroSelect && centros.length > 0) {
        const currentValue = centroSelect.value;
        centroSelect.innerHTML = '<option value="">-- Todos los centros --</option>' + 
          centros.map(c => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join('');
        centroSelect.value = currentValue;
      }
      
      const totalEl = document.getElementById('totalEspeciales');
      if (totalEl) totalEl.textContent = data.total_aprendices || 0;

      const tbody = document.getElementById('especialesTableBody');
      if (!tbody) return;

      if (!items.length) {
        tbody.innerHTML = '<tr><td colspan="4" class="text-center text-muted py-4">Sin registros</td></tr>';
        return;
      }

      tbody.innerHTML = items.map((row) => {
        let badgeClass = 'bg-secondary';
        const clasificacion = (row.clasificacion_programa_especial || 'NA').toUpperCase();
        
        if (clasificacion === 'ECONOMIA POPULAR') badgeClass = 'bg-purple';
        else if (clasificacion === 'FIC') badgeClass = 'bg-pink';
        else if (clasificacion === 'CAMPESENA') badgeClass = 'bg-olive';
        
        return `
          <tr>
            <td><strong>${escapeHtml(row.centro_formacion || 'N/A')}</strong></td>
            <td><span class="badge ${badgeClass}">${escapeHtml(row.clasificacion_programa_especial || 'NA')}</span></td>
            <td><strong>${row.total_fichas || 0}</strong></td>
            <td><span class="badge bg-success" style="font-size: 0.95rem;">${row.total_aprendices || 0}</span></td>
          </tr>
        `;
      }).join('');
    } catch (error) {
      const tbody = document.getElementById('especialesTableBody');
      if (tbody) tbody.innerHTML = `<tr><td colspan="4" class="text-center text-danger py-4">Error: ${error.message}</td></tr>`;
    }
  }

  // ===== EVENTOS PARA NAVBAR PRINCIPAL =====
  
  // Toggle para "Subida de Archivos" - Muestra/Oculta el navbar secundario
  if (toggleSubidaArchivosBtn && subidaArchivosNavbar) {
    toggleSubidaArchivosBtn.addEventListener('click', () => {
      // Mostrar navbar secundario
      subidaArchivosNavbar.style.display = 'block';
      
      // Mostrar solo la primera sección (Catálogo) y ocultar las demás
      if (catalogoSection) catalogoSection.style.display = 'block';
      if (registroSection) registroSection.style.display = 'none';
      if (ofertaSection) ofertaSection.style.display = 'none';
      if (consolidadoSection) consolidadoSection.style.display = 'none';
      if (pe04Section) pe04Section.style.display = 'none';
      
      // Ocultar módulo de Seguimiento a Metas
      if (seguimientoMetasSection) seguimientoMetasSection.style.display = 'none';
      
      // Activar el botón de Catálogo en el navbar secundario
      const catalogoBtn = document.getElementById('toggleSectionBtn');
      if (catalogoBtn) {
        catalogoBtn.classList.add('active');
        catalogoBtn.setAttribute('aria-pressed', 'true');
      }
      
      // Desactivar otros botones del navbar secundario
      ['toggleRegistroSection', 'toggleOfertaSection', 'toggleConsolidadoSection', 'togglePe04Section'].forEach(id => {
        const btn = document.getElementById(id);
        if (btn) {
          btn.classList.remove('active');
          btn.setAttribute('aria-pressed', 'false');
        }
      });
      
      // Actualizar estados de botones del navbar principal
      toggleSubidaArchivosBtn.classList.add('active');
      toggleSubidaArchivosBtn.setAttribute('aria-pressed', 'true');
      if (toggleSeguimientoMetasBtn) {
        toggleSeguimientoMetasBtn.classList.remove('active');
        toggleSeguimientoMetasBtn.setAttribute('aria-pressed', 'false');
      }
    });
  }

  // Toggle para "Seguimiento a las Metas" - Muestra/Oculta el nuevo módulo
  if (toggleSeguimientoMetasBtn && seguimientoMetasSection) {
    toggleSeguimientoMetasBtn.addEventListener('click', () => {
      // Mostrar módulo de Seguimiento a Metas
      seguimientoMetasSection.style.display = 'block';
      
      // Ocultar navbar secundario y todas las secciones de subida de archivos
      subidaArchivosNavbar.style.display = 'none';
      if (catalogoSection) catalogoSection.style.display = 'none';
      if (registroSection) registroSection.style.display = 'none';
      if (ofertaSection) ofertaSection.style.display = 'none';
      if (consolidadoSection) consolidadoSection.style.display = 'none';
      if (pe04Section) pe04Section.style.display = 'none';
      
      // Cargar datos de aprendices al mostrar la sección
      loadModalidades();
      loadEspeciales();
      
      // Actualizar estados de botones
      toggleSeguimientoMetasBtn.classList.add('active');
      toggleSeguimientoMetasBtn.setAttribute('aria-pressed', 'true');
      if (toggleSubidaArchivosBtn) {
        toggleSubidaArchivosBtn.classList.remove('active');
        toggleSubidaArchivosBtn.setAttribute('aria-pressed', 'false');
      }
    });
  }

  // ===== FUNCIONES PARA EL NUEVO MÓDULO "SEGUIMIENTO A LAS METAS" =====
  
  window.loadMetasPorCumplir = function() {
    showStatus('Cargando metas por cumplir...', 'info');
    // Placeholder - Se expandirá cuando el backend esté listo
    setTimeout(() => {
      showStatus('Funcionalidad en desarrollo', 'warning');
    }, 1000);
  };

  window.loadMetasCumplidas = function() {
    showStatus('Cargando metas cumplidas...', 'info');
    // Placeholder - Se expandirá cuando el backend esté listo
    setTimeout(() => {
      showStatus('Funcionalidad en desarrollo', 'warning');
    }, 1000);
  };

  window.loadAvanceGeneral = function() {
    showStatus('Cargando avance general...', 'info');
    // Placeholder - Se expandirá cuando el backend esté listo
    setTimeout(() => {
      showStatus('Funcionalidad en desarrollo', 'warning');
    }, 1000);
  };

  // Event listener para cargar Excel de Seguimiento de Metas
  document.getElementById('uploadSeguimientoMetasBtn')?.addEventListener('click', async () => {
    const fileInput = document.getElementById('seguimientoMetasFile');
    if (!fileInput || !fileInput.files.length) {
      const statusEl = document.getElementById('seguimientoMetasStatus');
      if (statusEl) statusEl.innerHTML = '<div class="alert alert-warning py-2 mb-0">Selecciona un archivo Excel</div>';
      return;
    }

    const formData = new FormData();
    formData.append('file', fileInput.files[0]);

    try {
      const statusEl = document.getElementById('seguimientoMetasStatus');
      const progressContainer = document.getElementById('seguimientoMetasProgressContainer');
      const progressBar = document.getElementById('seguimientoMetasProgressBar');
      
      if (progressContainer) progressContainer.style.display = 'block';
      if (statusEl) statusEl.innerHTML = '<div class="alert alert-info py-2 mb-0">Cargando archivo...</div>';

      const response = await fetch(`${API_BASE}/seguimiento-metas/upload-excel`, {
        method: 'POST',
        body: formData
      });

      const data = await response.json();

      if (!response.ok) {
        throw new Error(data.detail || 'Error al cargar el archivo');
      }

      if (statusEl) {
        statusEl.innerHTML = `<div class="alert alert-success py-2 mb-0">✅ ${data.message || 'Archivo cargado exitosamente'}</div>`;
      }

      fileInput.value = '';
      setTimeout(() => {
        if (progressContainer) progressContainer.style.display = 'none';
      }, 500);
    } catch (error) {
      const statusEl = document.getElementById('seguimientoMetasStatus');
      if (statusEl) statusEl.innerHTML = `<div class="alert alert-danger py-2 mb-0">❌ Error: ${error.message}</div>`;
    }
  });

  // Toggle para mostrar/ocultar tabla de Seguimiento de Metas
  const toggleSeguimientoMetasTableBtn = document.getElementById('toggleSeguimientoMetasTableBtn');
  const seguimientoMetasTableContainer = document.getElementById('seguimientoMetasTableContainer');
  
  if (toggleSeguimientoMetasTableBtn && seguimientoMetasTableContainer) {
    toggleSeguimientoMetasTableBtn.addEventListener('click', () => {
      if (seguimientoMetasTableContainer.style.display === 'none') {
        seguimientoMetasTableContainer.style.display = 'block';
        toggleSeguimientoMetasTableBtn.textContent = 'Ocultar tabla';
      } else {
        seguimientoMetasTableContainer.style.display = 'none';
        toggleSeguimientoMetasTableBtn.textContent = 'Mostrar tabla';
      }
    });
  }

  // ===== FUNCIONES PARA PE_04 (MINI MÓDULO) =====

  async function loadPe04Data() {
    try {
      const response = await fetch(`${API_BASE}/pe04-seguimiento/data`);
      const data = await response.json();

      if (!response.ok) {
        throw new Error(data.detail || 'Error al cargar datos');
      }

      const items = data.items || [];
      const pe04Total = document.getElementById('pe04Total');
      if (pe04Total) pe04Total.textContent = items.length;

      // Renderizar tabla
      const tbody = document.getElementById('pe04TableBody');
      if (!tbody) return;

      if (!items.length) {
        tbody.innerHTML = '<tr><td colspan="19" class="text-center text-muted py-4">Sin registros</td></tr>';
        return;
      }

      tbody.innerHTML = items.map((row) => {
        // Determinar color de badge según clasificación
        let badgeClass = 'bg-secondary';
        const clasificacion = (row.clasificacion_programa_especial || 'NA').toUpperCase();
        
        if (clasificacion === 'SENATEC') badgeClass = 'bg-primary';
        else if (clasificacion === 'ACME') badgeClass = 'bg-info';
        else if (clasificacion === 'SER CAMPESENA') badgeClass = 'bg-warning text-dark';
        else if (clasificacion === 'SER') badgeClass = 'bg-success';
        else if (clasificacion === 'BILINGUISMO') badgeClass = 'bg-danger';
        else if (clasificacion === 'CAMPESENA') badgeClass = 'bg-olive';
        else if (clasificacion === 'ECONOMIA POPULAR') badgeClass = 'bg-purple';
        else if (clasificacion === 'CAMPESENA RADIAL') badgeClass = 'bg-teal';
        else if (clasificacion === 'FIC') badgeClass = 'bg-pink';
        
        return `
          <tr>
            <td>${escapeHtml(row.id)}</td>
            <td>${escapeHtml(row.centro_formacion || '—')}</td>
            <td><span class="badge ${badgeClass}">${escapeHtml(row.clasificacion_programa_especial || 'NA')}</span></td>
            <td>${escapeHtml(row.numero_ficha || '—')}</td>
            <td>${escapeHtml(row.ciudad_municipio || '—')}</td>
            <td>${escapeHtml(row.fecha_inicio || '—')}</td>
            <td>${escapeHtml(row.fecha_fin || '—')}</td>
            <td>${escapeHtml(row.nivel_formacion || '—')}</td>
            <td>${escapeHtml(row.denominacion_programa || '—')}</td>
            <td>${escapeHtml(row.estrategia_programa || '—')}</td>
            <td>${escapeHtml(row.convenio || '—')}</td>
            <td>${escapeHtml(row.cupos || '—')}</td>
            <td>${escapeHtml(row.aprendices_activos || '—')}</td>
            <td>${escapeHtml(row.certificado || '—')}</td>
            <td>${escapeHtml(row.tipo_formacion || '—')}</td>
            <td>${escapeHtml(row.modalidad_formacion || '—')}</td>
            <td>${escapeHtml(row.estado_curso || '—')}</td>
            <td>${escapeHtml(row.fecha_corte || '—')}</td>
            <td><small>${escapeHtml(row.fecha_carga || '—')}</small></td>
          </tr>
        `;
      }).join('');
    } catch (error) {
      const statusEl = document.getElementById('pe04Status');
      if (statusEl) statusEl.innerHTML = `<div class="alert alert-danger py-2 mb-0">Error: ${error.message}</div>`;
    }
  }

  async function uploadPe04Excel() {
    const fileInput = document.getElementById('pe04File');
    if (!fileInput || !fileInput.files.length) {
      const statusEl = document.getElementById('pe04Status');
      if (statusEl) statusEl.innerHTML = '<div class="alert alert-warning py-2 mb-0">Selecciona un archivo Excel</div>';
      return;
    }

    const formData = new FormData();
    formData.append('file', fileInput.files[0]);

    try {
      const statusEl = document.getElementById('pe04Status');
      const progressContainer = document.getElementById('pe04ProgressContainer');
      const progressBar = document.getElementById('pe04ProgressBar');
      
      if (progressContainer) progressContainer.style.display = 'block';
      if (statusEl) statusEl.innerHTML = '<div class="alert alert-info py-2 mb-0">Cargando archivo...</div>';

      const response = await fetch(`${API_BASE}/pe04-seguimiento/upload-excel`, {
        method: 'POST',
        body: formData
      });

      const data = await response.json();

      if (!response.ok) {
        throw new Error(data.detail || 'Error al cargar el archivo');
      }

      if (statusEl) {
        statusEl.innerHTML = `<div class="alert alert-success py-2 mb-0">✅ ${data.message || 'Archivo cargado exitosamente'}</div>`;
      }

      fileInput.value = '';
      
      // Cargar datos actualizado
      await loadPe04Data();
      
      setTimeout(() => {
        if (progressContainer) progressContainer.style.display = 'none';
      }, 500);
    } catch (error) {
      const statusEl = document.getElementById('pe04Status');
      if (statusEl) statusEl.innerHTML = `<div class="alert alert-danger py-2 mb-0">❌ Error: ${error.message}</div>`;
    }
  }

  // Event listeners para PE_04
  document.getElementById('uploadPe04Btn')?.addEventListener('click', uploadPe04Excel);
  
  document.getElementById('reloadPe04Btn')?.addEventListener('click', () => {
    loadPe04Data();
  });

  const togglePe04TableBtn = document.getElementById('togglePe04TableBtn');
  const pe04TableContainer = document.getElementById('pe04TableContainer');
  
  if (togglePe04TableBtn && pe04TableContainer) {
    togglePe04TableBtn.addEventListener('click', () => {
      if (pe04TableContainer.style.display === 'none') {
        pe04TableContainer.style.display = 'block';
        togglePe04TableBtn.textContent = 'Ocultar tabla';
        loadPe04Data();
      } else {
        pe04TableContainer.style.display = 'none';
        togglePe04TableBtn.textContent = 'Mostrar tabla';
      }
    });
  }

  document.getElementById('exportPe04Btn')?.addEventListener('click', () => {
    downloadExcel(`${API_BASE}/pe04-seguimiento/exportar-excel`, 'pe04_seguimiento.xlsx');
  });

  // Event listeners para tablas de modalidades y especiales
  const toggleModalidadesBtn = document.getElementById('toggleModalidadesBtn');
  const modalidadesTableContainer = document.getElementById('modalidadesTableContainer');
  const filtroModalidadesCentro = document.getElementById('filtroModalidadesCentro');
  
  if (toggleModalidadesBtn && modalidadesTableContainer) {
    toggleModalidadesBtn.addEventListener('click', () => {
      if (modalidadesTableContainer.style.display === 'none') {
        modalidadesTableContainer.style.display = 'block';
        toggleModalidadesBtn.textContent = 'Ocultar Tabla';
        loadModalidades(filtroModalidadesCentro?.value || '');
      } else {
        modalidadesTableContainer.style.display = 'none';
        toggleModalidadesBtn.textContent = 'Mostrar Tabla';
      }
    });
  }
  
  if (filtroModalidadesCentro) {
    filtroModalidadesCentro.addEventListener('change', () => {
      if (modalidadesTableContainer.style.display !== 'none') {
        loadModalidades(filtroModalidadesCentro.value || '');
      }
    });
  }

  const toggleEspecialesBtn = document.getElementById('toggleEspecialesBtn');
  const especialesTableContainer = document.getElementById('especialesTableContainer');
  const filtroEspecialesCentro = document.getElementById('filtroEspecialesCentro');
  
  if (toggleEspecialesBtn && especialesTableContainer) {
    toggleEspecialesBtn.addEventListener('click', () => {
      if (especialesTableContainer.style.display === 'none') {
        especialesTableContainer.style.display = 'block';
        toggleEspecialesBtn.textContent = 'Ocultar Tabla';
        loadEspeciales(filtroEspecialesCentro?.value || '');
      } else {
        especialesTableContainer.style.display = 'none';
        toggleEspecialesBtn.textContent = 'Mostrar Tabla';
      }
    });
  }
  
  if (filtroEspecialesCentro) {
    filtroEspecialesCentro.addEventListener('change', () => {
      if (especialesTableContainer.style.display !== 'none') {
        loadEspeciales(filtroEspecialesCentro.value || '');
      }
    });
  }

  // Cargar datos de PE_04 al iniciar si existen
  loadPe04Data();
});

