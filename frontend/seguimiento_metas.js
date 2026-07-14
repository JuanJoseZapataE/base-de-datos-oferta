const API_BASE = 'http://127.0.0.1:8001';

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
      <td>${escapeHtml(row.prf_duracion_maxima || 'ÔÇö')}</td>
      <td>${escapeHtml(row.prf_dur_etapa_lectiva || 'ÔÇö')}</td>
      <td>${escapeHtml(row.prf_dur_etapa_prod || 'ÔÇö')}</td>
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
        // si el usuario ya ten├¡a un valor seleccionado, mantenerlo si existe
        const prev = (selectedNivel || '');
        if (prev) { nivelSelect.value = prev; }
      }
    } catch (e) { console.warn(e); }

    // Renderizar items tal cual vienen del servidor (server-side paging)
    renderCatalogo(items);

    // mostrar total real de la base de datos (no el tama├▒o de la p├ígina)
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

// ===== MINI M├ôDULO: AGREGAR PROGRAMAS EN CAT├üLOGO =====

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
    showAgregarProgramasStatus('Procesando c├│digos...', 'info');
    
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
    const message = `Ô£ô Completado: ${data.updated} programas actualizados, ${data.not_found} no encontrados. Tipo: ${escapeHtml(data.tipo_programa)}.`;
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
      <td><small>${escapeHtml(row.id || 'ÔÇö')}</small></td>
      <td><small>${escapeHtml(row.proceso || 'ÔÇö')}</small></td>
      <td><small>${escapeHtml(row.tipo_tramite || 'ÔÇö')}</small></td>
      <td><small>${row.fecha_radicado ? new Date(row.fecha_radicado).toLocaleDateString('es-CO') : 'ÔÇö'}</small></td>
      <td><small>${escapeHtml(row.numero_resolucion || 'ÔÇö')}</small></td>
      <td><small>${row.fecha_resolucion ? new Date(row.fecha_resolucion).toLocaleDateString('es-CO') : 'ÔÇö'}</small></td>
      <td><small>${escapeHtml(row.resuelve || 'ÔÇö')}</small></td>
      <td><small>${escapeHtml(row.decreto_ampara || 'ÔÇö')}</small></td>
      <td><small>${escapeHtml(row.snies || 'ÔÇö')}</small></td>
      <td><small>${escapeHtml(row.cobertura || 'ÔÇö')}</small></td>
      <td><small>${escapeHtml(row.resolucion_ampara_programa || 'ÔÇö')}</small></td>
      <td><small>${escapeHtml(row.resolucion_ampara || 'ÔÇö')}</small></td>
      <td><small>${escapeHtml(row.resolucion_ampara_fecha || 'ÔÇö')}</small></td>
      <td><small>${row.fecha_vencimiento ? new Date(row.fecha_vencimiento).toLocaleDateString('es-CO') : 'ÔÇö'}</small></td>
      <td><small>${escapeHtml(row.vigencia_rc || 'ÔÇö')}</small></td>
      <td><small>${escapeHtml(row.cod_programa || 'ÔÇö')}</small></td>
      <td><small>${escapeHtml(row.version || 'ÔÇö')}</small></td>
      <td><small><strong>${escapeHtml(row.nombre_programa || 'ÔÇö')}</strong></small></td>
      <td><small>${escapeHtml(row.nivel_formacion || 'ÔÇö')}</small></td>
      <td><small>${escapeHtml(row.red_conocimiento || 'ÔÇö')}</small></td>
      <td><small>${escapeHtml(row.modalidad || 'ÔÇö')}</small></td>
      <td><small>${escapeHtml(row.centro_formacion || 'ÔÇö')}</small></td>
      <td><small>${escapeHtml(row.nombre_sede || 'ÔÇö')}</small></td>
      <td><small>${escapeHtml(row.tipo_sede || 'ÔÇö')}</small></td>
      <td><small>${escapeHtml(row.municipio || 'ÔÇö')}</small></td>
      <td><small>${escapeHtml(row.lugar_desarrollo || 'ÔÇö')}</small></td>
      <td><small>${escapeHtml(row.direccion || 'ÔÇö')}</small></td>
      <td><small>${escapeHtml(row.regional || 'ÔÇö')}</small></td>
      <td><small>${escapeHtml(row.nombre_regional || 'ÔÇö')}</small></td>
      <td><small>${escapeHtml(row.observaciones || 'ÔÇö')}</small></td>
      <td><small>${escapeHtml(row.clasificacion_tramite || 'ÔÇö')}</small></td>
      <td><small>${escapeHtml(row.aprendices_primer_cohorte || 'ÔÇö')}</small></td>
      <td><small>${escapeHtml(row.lugar_desarrollo_resolucion || 'ÔÇö')}</small></td>
      <td><small>${row.fecha_registro ? new Date(row.fecha_registro).toLocaleDateString('es-CO') : 'ÔÇö'}</small></td>
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
  // ===== INICIALIZACI├ôN: Declarar todas las variables del DOM =====
  const catalogoSection = document.getElementById('catalogoSection');
  const registroSection = document.getElementById('registroSection');
  const ofertaSection = document.getElementById('ofertaSection');
  const consolidadoSection = document.getElementById('consolidadoSection');
  const pe04Section = document.getElementById('pe04Section');
  const subidaArchivosNavbar = document.getElementById('subidaArchivosNavbar');
  const seguimientoMetasSection = document.getElementById('seguimientoMetasModuleSection');
  const registroMetasSection = document.getElementById('registroMetasModuleSection');
  const toggleSubidaArchivosBtn = document.getElementById('toggleSubidaArchivosSection');
  const toggleSeguimientoMetasBtn = document.getElementById('toggleSeguimientoMetasModuleBtn');
  const toggleRegistroMetasBtn = document.getElementById('toggleRegistroMetasModuleBtn');

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
        // sincronizar visibilidad de la paginaci├│n
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

  // ===== FUNCI├ôN ANTERIOR (DEPRECATED) - Reemplazada por toggles exclusivos =====
  // Las funciones syncSectionToggle y bindSectionToggle fueron reemplazadas por 
  // la l├│gica de toggles exclusivos para mejor UX

  // ===== TOGGLES EXCLUSIVOS PARA SUBM├ôDULOS DE SUBIDA =====
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
        // Ocultar todos los dem├ís subm├│dulos
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
        
        // Mostrar el subm├│dulo seleccionado
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

  // Cargar datos de Registro Calificado autom├íticamente al cargar la p├ígina
  loadRegistroData();

  // ===== M├ôDULO SEGUIMIENTO METAS: OFERTA =====
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

    // A├▒adir listeners a botones de editar
    Array.from(document.getElementsByClassName('editVerBtn')).forEach(btn => {
      btn.removeEventListener('click', handleEditVerClick);
      btn.addEventListener('click', handleEditVerClick);
    });
  }

  function handleEditVerClick(e) {
    const id = e.currentTarget.getAttribute('data-id');
    if (!id) return;
    const current = e.currentTarget.closest('tr')?.querySelectorAll('td')[18]?.textContent || '';
    const newVal = prompt('Establece verificado: (VERIFICADO / NO VERIFICADO / VERIFICACION MANUAL / REGISTRO VENCIDO). Dejar vac├¡o para NULL', current.trim() || '');
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
        <td><small>${escapeHtml(row.id || 'ÔÇö')}</small></td>
        <td><small>${escapeHtml(row.nombre_real_institucion || 'ÔÇö')}</small></td>
        <td><small>${escapeHtml(row.nombres_sofia_plus || 'ÔÇö')}</small></td>
        <td><small>${escapeHtml(row.municipio || 'ÔÇö')}</small></td>
        <td><small>${escapeHtml(row.clasificacion || 'ÔÇö')}</small></td>
        <td><small>${row.fecha_registro ? new Date(row.fecha_registro).toLocaleDateString('es-CO') : 'ÔÇö'}</small></td>
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
        showConsolidadoStatus('Sin registros a├║n', 'warning');
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

  // ===== FUNCIONES DE EXPORTACI├ôN A EXCEL =====
  
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

  // Event listeners para botones de exportaci├│n
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

  function hideAllMainModuleSections() {
    if (subidaArchivosNavbar) subidaArchivosNavbar.style.display = 'none';
    if (catalogoSection) catalogoSection.style.display = 'none';
    if (registroSection) registroSection.style.display = 'none';
    if (ofertaSection) ofertaSection.style.display = 'none';
    if (consolidadoSection) consolidadoSection.style.display = 'none';
    if (pe04Section) pe04Section.style.display = 'none';
    if (seguimientoMetasSection) seguimientoMetasSection.style.display = 'none';
    if (registroMetasSection) registroMetasSection.style.display = 'none';
  }

  function deactivateMainNavButtons() {
    [toggleSubidaArchivosBtn, toggleSeguimientoMetasBtn, toggleRegistroMetasBtn].forEach((btn) => {
      if (btn) {
        btn.classList.remove('active');
        btn.setAttribute('aria-pressed', 'false');
      }
    });
  }
  
  // Toggle para "Subida de Archivos" - Muestra/Oculta el navbar secundario
  if (toggleSubidaArchivosBtn && subidaArchivosNavbar) {
    toggleSubidaArchivosBtn.addEventListener('click', () => {
      hideAllMainModuleSections();
      subidaArchivosNavbar.style.display = 'block';
      if (catalogoSection) catalogoSection.style.display = 'block';
      
      const catalogoBtn = document.getElementById('toggleSectionBtn');
      if (catalogoBtn) {
        catalogoBtn.classList.add('active');
        catalogoBtn.setAttribute('aria-pressed', 'true');
      }
      
      ['toggleRegistroSection', 'toggleOfertaSection', 'toggleConsolidadoSection', 'togglePe04Section'].forEach(id => {
        const btn = document.getElementById(id);
        if (btn) {
          btn.classList.remove('active');
          btn.setAttribute('aria-pressed', 'false');
        }
      });
      
      deactivateMainNavButtons();
      toggleSubidaArchivosBtn.classList.add('active');
      toggleSubidaArchivosBtn.setAttribute('aria-pressed', 'true');
    });
  }

  // Toggle para "Seguimiento a las Metas" - Muestra/Oculta el nuevo m├│dulo
  if (toggleSeguimientoMetasBtn && seguimientoMetasSection) {
    toggleSeguimientoMetasBtn.addEventListener('click', () => {
      hideAllMainModuleSections();
      seguimientoMetasSection.style.display = 'block';
      loadModalidades();
      loadEspeciales();
      deactivateMainNavButtons();
      toggleSeguimientoMetasBtn.classList.add('active');
      toggleSeguimientoMetasBtn.setAttribute('aria-pressed', 'true');
    });
  }

  // Toggle para "Registro Metas"
  if (toggleRegistroMetasBtn && registroMetasSection) {
    toggleRegistroMetasBtn.addEventListener('click', () => {
      hideAllMainModuleSections();
      registroMetasSection.style.display = 'block';
      loadRegistroMetasList();
      loadGruposMetasList();
      deactivateMainNavButtons();
      toggleRegistroMetasBtn.classList.add('active');
      toggleRegistroMetasBtn.setAttribute('aria-pressed', 'true');
    });
  }

  // ===== FUNCIONES PARA EL NUEVO M├ôDULO "SEGUIMIENTO A LAS METAS" =====
  
  window.loadMetasPorCumplir = function() {
    showStatus('Cargando metas por cumplir...', 'info');
    // Placeholder - Se expandir├í cuando el backend est├® listo
    setTimeout(() => {
      showStatus('Funcionalidad en desarrollo', 'warning');
    }, 1000);
  };

  window.loadMetasCumplidas = function() {
    showStatus('Cargando metas cumplidas...', 'info');
    // Placeholder - Se expandir├í cuando el backend est├® listo
    setTimeout(() => {
      showStatus('Funcionalidad en desarrollo', 'warning');
    }, 1000);
  };

  window.loadAvanceGeneral = function() {
    showStatus('Cargando avance general...', 'info');
    // Placeholder - Se expandir├í cuando el backend est├® listo
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
        statusEl.innerHTML = `<div class="alert alert-success py-2 mb-0">✓ ${data.message || 'Archivo cargado exitosamente'}</div>`;
      }

      fileInput.value = '';
      setTimeout(() => {
        if (progressContainer) progressContainer.style.display = 'none';
      }, 500);
    } catch (error) {
      const statusEl = document.getElementById('seguimientoMetasStatus');
      if (statusEl) statusEl.innerHTML = `<div class="alert alert-danger py-2 mb-0">Error: ${error.message}</div>`;
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

  // ===== FUNCIONES PARA PE_04 (MINI M├ôDULO) =====

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
        // Determinar color de badge seg├║n clasificaci├│n
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
            <td>${escapeHtml(row.centro_formacion || 'ÔÇö')}</td>
            <td><span class="badge ${badgeClass}">${escapeHtml(row.clasificacion_programa_especial || 'NA')}</span></td>
            <td>${escapeHtml(row.numero_ficha || 'ÔÇö')}</td>
            <td>${escapeHtml(row.ciudad_municipio || 'ÔÇö')}</td>
            <td>${escapeHtml(row.fecha_inicio || 'ÔÇö')}</td>
            <td>${escapeHtml(row.fecha_fin || 'ÔÇö')}</td>
            <td>${escapeHtml(row.nivel_formacion || 'ÔÇö')}</td>
            <td>${escapeHtml(row.denominacion_programa || 'ÔÇö')}</td>
            <td>${escapeHtml(row.estrategia_programa || 'ÔÇö')}</td>
            <td>${escapeHtml(row.convenio || 'ÔÇö')}</td>
            <td>${escapeHtml(row.cupos || 'ÔÇö')}</td>
            <td>${escapeHtml(row.aprendices_activos || 'ÔÇö')}</td>
            <td>${escapeHtml(row.certificado || 'ÔÇö')}</td>
            <td>${escapeHtml(row.tipo_formacion || 'ÔÇö')}</td>
            <td>${escapeHtml(row.modalidad_formacion || 'ÔÇö')}</td>
            <td>${escapeHtml(row.estado_curso || 'ÔÇö')}</td>
            <td>${escapeHtml(row.fecha_corte || 'ÔÇö')}</td>
            <td><small>${escapeHtml(row.fecha_carga || 'ÔÇö')}</small></td>
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
        statusEl.innerHTML = `<div class="alert alert-success py-2 mb-0">Ô£à ${data.message || 'Archivo cargado exitosamente'}</div>`;
      }

      fileInput.value = '';
      
      // Cargar datos actualizado
      await loadPe04Data();
      
      setTimeout(() => {
        if (progressContainer) progressContainer.style.display = 'none';
      }, 500);
    } catch (error) {
      const statusEl = document.getElementById('pe04Status');
      if (statusEl) statusEl.innerHTML = `<div class="alert alert-danger py-2 mb-0">ÔØî Error: ${error.message}</div>`;
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

  // ===== M├ôDULO REGISTRO METAS =====

  const CENTROS_LABELS = {
    '9308': '9308 - CENTRO DE COMERCIO Y SERVICIOS',
    '9121': '9121 - CENTRO ATENCION SECTOR AGROPECUARIO',
    '9223': '9223 - CENTRO DE DISE├æO E INNOVACI├ôN TECNOL├ôGICA INDUSTRIAL'
  };

  function showRegistroMetaFormStatus(message, type = 'secondary') {
    const el = document.getElementById('registroMetaFormStatus');
    if (!el) return;
    el.innerHTML = `<div class="alert alert-${type} py-2 mb-0">${escapeHtml(message)}</div>`;
  }

  function showRegistroMetasGrupoStatus(message, type = 'secondary') {
    const el = document.getElementById('registroMetasGrupoStatus');
    if (!el) return;
    el.innerHTML = `<div class="alert alert-${type} py-2 mb-0">${escapeHtml(message)}</div>`;
  }

  let registroMetasCache = [];
  let gruposMetasCache = [];

  function isNumericFieldValue(value) {
    return /^\d+$/.test(String(value || '').trim());
  }

  function isNumericList(value) {
  if (!value) return true;
  return /^\d+(\s*[,;\n]\s*\d+)*$/.test(String(value).trim());
  }

  function bindNumericOnlyInput(inputId) {
    const input = document.getElementById(inputId);
    if (!input) return;
    input.addEventListener('input', () => {
      input.value = input.value.replace(/\D/g, '');
    });
  }

  ['metaCodigoNivelFormacion', 'metaCupos']
  .forEach(bindNumericOnlyInput);

  function getSelectedMetaIds() {
    return Array.from(document.querySelectorAll('.meta-select-checkbox:checked'))
      .map((el) => Number(el.value))
      .filter((id) => Number.isFinite(id) && id > 0);
  }

  function updateCrearGrupoButtonState() {
    const btn = document.getElementById('crearGrupoMetasBtn');
    if (!btn) return;
    btn.disabled = getSelectedMetaIds().length === 0;
  }

  function renderRegistroMetasTable(rows) {
    const tbody = document.getElementById('registroMetasTableBody');
    const totalEl = document.getElementById('registroMetasTotal');
    if (!tbody) return;

    if (totalEl) totalEl.textContent = String(rows.length);

    if (!rows.length) {
      tbody.innerHTML = '<tr><td colspan="10" class="text-center text-muted py-4">Sin filtros registrados</td></tr>';
      updateCrearGrupoButtonState();
      return;
    }

    console.log(rows[0]);

    tbody.innerHTML = rows.map((row) => {
      const enGrupo = row.grupo_id ? true : false;
      const centroLabel = row.centro_formacion || CENTROS_LABELS[row.codigo_centro] || row.codigo_centro || '—';
      return `
        <tr>
          <td>
            <input
              type="checkbox"
              class="form-check-input meta-select-checkbox"
              value="${row.id}"
              ${enGrupo ? 'disabled title="Ya pertenece a un grupo"' : ''}
              style="width:22px; height:22px; border:2px solid #444; border-radius:4px; cursor:pointer; box-shadow:0 0 2px rgba(0,0,0,.4);"
            >
          </td>
          <td><strong>${row.id}</strong></td>
          <td>${escapeHtml(row.nombre_meta)}</td>
          <td>${escapeHtml(row.tipo_formacion)}</td>
          <td>${escapeHtml(row.codigo_nivel_formacion)}</td>
          <td>${escapeHtml(row.codigo_programa_especial)}</td>
          <td>${escapeHtml(row.codigo_convenio)}</td>
          <td><span class="badge bg-info">${escapeHtml(row.tipo_modalidad)}</span></td>
          <td><small>${escapeHtml(centroLabel)}</small></td>
          <td>${row.grupo_id ? `<span class="badge bg-primary">Grupo ${row.grupo_id}</span>` : '<span class="text-muted">—</span>'}</td>
          <td>
            <div class="d-flex flex-wrap gap-1">
              <button type="button" class="btn btn-sm btn-outline-danger registro-meta-action" data-action="delete" data-id="${row.id}">Eliminar</button>
            </div>
          </td>
        </tr>
      `;
    }).join('');

    document.querySelectorAll('.meta-select-checkbox').forEach((checkbox) => {
      checkbox.addEventListener('change', updateCrearGrupoButtonState);
    });
    updateCrearGrupoButtonState();
  }

  async function loadRegistroMetasList() {
    try {
      showRegistroMetasGrupoStatus('Cargando metas...', 'info');
      const resp = await fetch(`${API_BASE}/registro-metas/lista`);
      const data = await resp.json().catch(() => null);
      if (!resp.ok) {
        throw new Error(data?.detail || `${resp.status} ${resp.statusText}`);
      }
      registroMetasCache = Array.isArray(data?.items) ? data.items : [];
      renderRegistroMetasTable(registroMetasCache);
      showRegistroMetasGrupoStatus(`Se cargaron ${data?.total || 0} metas.`, 'success');
    } catch (error) {
      registroMetasCache = [];
      renderRegistroMetasTable([]);
      showRegistroMetasGrupoStatus(`Error al cargar metas: ${error.message}`, 'danger');
    }
  }

  function renderGruposMetasTable(rows) {
    const tbody = document.getElementById('gruposMetasTableBody');
    if (!tbody) return;

    if (!rows.length) {
      tbody.innerHTML = '<tr><td colspan="6" class="text-center text-muted py-4">Sin metas creadas</td></tr>';
      return;
    }

    tbody.innerHTML = rows.map((row) => `
      <tr>
        <td><strong>${row.id}</strong></td>
        <td>${escapeHtml(row.nombre_grupo)}</td>
        <td>${row.cantidad_metas || 0}</td>
        <td><span class="badge bg-success">${row.total_cupos || 0}</span></td>
        <td><small>${escapeHtml(row.fecha_creacion || '—')}</small></td>
        <td>
          <div class="d-flex flex-wrap gap-1">
            <button type="button" class="btn btn-sm btn-outline-danger grupo-meta-action" data-action="delete" data-id="${row.id}">Eliminar</button>
          </div>
        </td>
      </tr>
    `).join('');
  }

  async function loadGruposMetasList() {
    try {
      const resp = await fetch(`${API_BASE}/registro-metas/grupos`);
      const data = await resp.json().catch(() => null);
      if (!resp.ok) {
        throw new Error(data?.detail || `${resp.status} ${resp.statusText}`);
      }
      gruposMetasCache = Array.isArray(data?.items) ? data.items : [];
      renderGruposMetasTable(gruposMetasCache);
    } catch (error) {
      gruposMetasCache = [];
      renderGruposMetasTable([]);
    }
  }

  async function eliminarFiltroMeta(id) {
    if (!window.confirm('¿Deseas eliminar este filtro?')) return;
    try {
      showRegistroMetaFormStatus('Eliminando filtro...', 'info');
      const resp = await fetch(`${API_BASE}/registro-metas/filtros/${id}`, { method: 'DELETE' });
      const data = await resp.json().catch(() => null);
      if (!resp.ok) {
        throw new Error(data?.detail || `${resp.status} ${resp.statusText}`);
      }
      showRegistroMetaFormStatus(`Filtro ${id} eliminado.`, 'success');
      await loadRegistroMetasList();
      await loadGruposMetasList();
    } catch (error) {
      showRegistroMetaFormStatus(`Error al eliminar filtro: ${error.message}`, 'danger');
    }
  }

  async function eliminarGrupoMeta(id) {
    if (!window.confirm('¿Deseas eliminar esta meta? Los filtros quedarán sin meta asociada.')) return;
    try {
      showRegistroMetasGrupoStatus('Eliminando meta...', 'info');
      const resp = await fetch(`${API_BASE}/registro-metas/grupos/${id}`, { method: 'DELETE' });
      const data = await resp.json().catch(() => null);
      if (!resp.ok) {
        throw new Error(data?.detail || `${resp.status} ${resp.statusText}`);
      }
      showRegistroMetasGrupoStatus(`Meta ${id} eliminada.`, 'success');
      await loadRegistroMetasList();
      await loadGruposMetasList();
    } catch (error) {
      showRegistroMetasGrupoStatus(`Error al eliminar meta: ${error.message}`, 'danger');
    }
  }

  async function eliminarTodoRegistroMetas() {
    if (!window.confirm('¿Deseas eliminar TODO lo registrado en metas y filtros? Esta acción no se puede deshacer.')) return;
    try {
      showRegistroMetasGrupoStatus('Eliminando todo el registro de metas...', 'info');
      const resp = await fetch(`${API_BASE}/registro-metas/delete-all`, { method: 'DELETE' });
      const data = await resp.json().catch(() => null);
      if (!resp.ok) {
        throw new Error(data?.detail || `${resp.status} ${resp.statusText}`);
      }
      showRegistroMetasGrupoStatus(
        `Se eliminó todo el registro: ${data.deleted_filters || 0} filtros y ${data.deleted_groups || 0} metas.`,
        'success'
      );
      document.getElementById('registroMetaForm')?.reset();
      document.getElementById('nombreGrupoMetas') && (document.getElementById('nombreGrupoMetas').value = '');
      const selectAll = document.getElementById('selectAllMetasCheckbox');
      if (selectAll) selectAll.checked = false;
      await loadRegistroMetasList();
      await loadGruposMetasList();
    } catch (error) {
      showRegistroMetasGrupoStatus(`Error al eliminar todo: ${error.message}`, 'danger');
    }
  }

  async function submitRegistroMeta(event) {
    event.preventDefault();

    const tipoFormacion = document.getElementById('metaTipoFormacion')?.value.trim();
    const codigoNivel = document.getElementById('metaCodigoNivelFormacion')?.value.trim();
    const codigoPrograma = document.getElementById('metaCodigoProgramaEspecial')?.value.trim();
    const codigoConvenio = document.getElementById('metaCodigoConvenio')?.value.trim();
    const tipoModalidad = document.getElementById('metaTipoModalidad')?.value.trim();
    const nombreMeta = document.getElementById('metaNombre')?.value.trim();
    const codigoCentro = document.getElementById('metaCentroFormacion')?.value.trim();

    if (!tipoFormacion || !tipoModalidad || !nombreMeta || !codigoCentro) {
      showRegistroMetaFormStatus('Completa todos los campos obligatorios.', 'warning');
      return;
    }

    if (!isNumericFieldValue(codigoNivel)) {
      showRegistroMetaFormStatus(
        'El código de nivel de formación solo puede contener números.',
        'warning'
        );
        return;
    }

    if (!isNumericList(codigoPrograma)) {
      showRegistroMetaFormStatus(
        'Los códigos de programa especial deben ser números separados por comas, punto y coma o saltos de línea.',
        'warning'
      );
      return;
    }

    if (!isNumericList(codigoConvenio)) {
      showRegistroMetaFormStatus(
        'Los códigos de convenio deben ser números separados por comas, punto y coma o saltos de línea.',
        'warning'
      );
      return;
    }

    const formData = new FormData();
    formData.append('tipo_formacion', tipoFormacion);
    formData.append('codigo_nivel_formacion', codigoNivel);
    formData.append('codigo_programa_especial', codigoPrograma);
    formData.append('codigo_convenio', codigoConvenio);
    formData.append('tipo_modalidad', tipoModalidad);
    formData.append('nombre_meta', nombreMeta);
    formData.append('codigo_centro', codigoCentro);
    formData.append('centro_formacion', CENTROS_LABELS[codigoCentro] || codigoCentro);

    try {
      showRegistroMetaFormStatus('Guardando meta...', 'info');
      const resp = await fetch(`${API_BASE}/registro-metas/crear`, { method: 'POST', body: formData });
      const data = await resp.json().catch(() => null);
      if (!resp.ok) {
        throw new Error(data?.detail || `${resp.status} ${resp.statusText}`);
      }
      showRegistroMetaFormStatus(`Meta registrada correctamente (ID ${data.id}).`, 'success');
      document.getElementById('registroMetaForm')?.reset();
      await loadRegistroMetasList();
    } catch (error) {
      showRegistroMetaFormStatus(`Error al registrar meta: ${error.message}`, 'danger');
    }
  }

  async function crearGrupoMetas() {
  const selectedIds = getSelectedMetaIds();
  if (!selectedIds.length) {
    showRegistroMetasGrupoStatus('Selecciona al menos una meta para crear el grupo.', 'warning');
    return;
  }

  const nombreGrupo = (document.getElementById('nombreGrupoMetas')?.value || '').trim();
  if (!nombreGrupo) {
    showRegistroMetasGrupoStatus('Escribe un nombre para el grupo.', 'warning');
    return;
  }

  const totalCupos = parseInt(
    document.getElementById('metaTotalCupos')?.value || '',
    10
  );

  if (isNaN(totalCupos) || totalCupos < 0) {
    showRegistroMetasGrupoStatus('El total de cupos debe ser un número válido.', 'warning');
    return;
  }

  try {
    showRegistroMetasGrupoStatus('Creando grupo...', 'info');

    const resp = await fetch(`${API_BASE}/registro-metas/crear-grupo`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        nombre_grupo: nombreGrupo,
        total_cupos: totalCupos,
        meta_ids: selectedIds
      })
    });

    const data = await resp.json().catch(() => null);

    if (!resp.ok) {
      throw new Error(data?.detail || `${resp.status} ${resp.statusText}`);
    }

    showRegistroMetasGrupoStatus(
      `Grupo "${nombreGrupo}" creado. Metas: ${data.cantidad_metas}, total cupos: ${data.total_cupos}.`,
      'success'
    );

    document.getElementById('nombreGrupoMetas').value = '';
    document.getElementById('metaTotalCupos').value = '';

    const selectAll = document.getElementById('selectAllMetasCheckbox');
    if (selectAll) selectAll.checked = false;

    await loadRegistroMetasList();
    await loadGruposMetasList();

  } catch (error) {
    showRegistroMetasGrupoStatus(`Error al crear grupo: ${error.message}`, 'danger');
  }
}

  document.getElementById('registroMetaForm')?.addEventListener('submit', submitRegistroMeta);
  document.getElementById('reloadRegistroMetasBtn')?.addEventListener('click', loadRegistroMetasList);
  document.getElementById('reloadGruposMetasBtn')?.addEventListener('click', loadGruposMetasList);
  document.getElementById('crearGrupoMetasBtn')?.addEventListener('click', crearGrupoMetas);
  document.getElementById('deleteAllRegistroMetasBtn')?.addEventListener('click', eliminarTodoRegistroMetas);

  document.getElementById('registroMetasTableBody')?.addEventListener('click', (event) => {
    const button = event.target.closest('button.registro-meta-action');
    if (!button) return;
    const action = button.getAttribute('data-action');
    const id = Number(button.getAttribute('data-id'));
    if (action === 'delete') {
      eliminarFiltroMeta(id);
    }
  });

  document.getElementById('gruposMetasTableBody')?.addEventListener('click', (event) => {
    const button = event.target.closest('button.grupo-meta-action');
    if (!button) return;
    const action = button.getAttribute('data-action');
    const id = Number(button.getAttribute('data-id'));
    if (action === 'delete') {
      eliminarGrupoMeta(id);
    }
  });

  document.getElementById('selectAllMetasCheckbox')?.addEventListener('change', (event) => {
    const checked = event.target.checked;
    document.querySelectorAll('.meta-select-checkbox:not(:disabled)').forEach((checkbox) => {
      checkbox.checked = checked;
    });
    updateCrearGrupoButtonState();
  });
});

