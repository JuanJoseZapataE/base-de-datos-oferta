const API_BASE = 'http://127.0.0.1:8001';

function escapeHtml(v) {
  if (v === null || v === undefined) return '';
  return String(v).replace(/[&<>\"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}

function normalizeMultiNumericInput(value) {
  const parts = String(value || '')
    .split(/[\n,;]+/)
    .map((part) => part.trim())
    .filter(Boolean);
  return Array.from(new Set(parts)).join(',');
}

function formatMultiNumericValue(value) {
  const text = String(value || '').trim();
  return text ? escapeHtml(text.replace(/,/g, ', ')) : '—';
}

function cloneNodeWithoutListeners(id) {
  const node = document.getElementById(id);
  if (!node || !node.parentNode) return node;
  const clone = node.cloneNode(true);
  node.parentNode.replaceChild(clone, node);
  return clone;
}

document.addEventListener('DOMContentLoaded', () => {
  const filterForm = cloneNodeWithoutListeners('registroMetaForm');
  const reloadFiltersBtn = cloneNodeWithoutListeners('reloadRegistroMetasBtn');
  const reloadMetasBtn = cloneNodeWithoutListeners('reloadGruposMetasBtn');
  const createMetaBtn = cloneNodeWithoutListeners('crearGrupoMetasBtn');
  const selectAllCheckbox = cloneNodeWithoutListeners('selectAllMetasCheckbox');
  const filtersTbody = cloneNodeWithoutListeners('registroMetasTableBody');
  const metasTbody = cloneNodeWithoutListeners('gruposMetasTableBody');
  const cancelFilterEditBtn = document.getElementById('cancelRegistroMetaEditBtn');
  const cancelMetaEditBtn = document.getElementById('cancelRegistroGrupoEditBtn');
  const detailModal = document.getElementById('metaDetalleModal');
  const detailCloseBtn = document.getElementById('metaDetalleCloseBtn');

  let filtroEditId = null;
  let metaEditId = null;
  let filtrosCache = [];
  let metasCache = [];
  let metaDetalleActual = null;

  function showFilterStatus(message, type = 'secondary') {
    const el = document.getElementById('registroMetaFormStatus');
    if (el) el.innerHTML = `<div class="alert alert-${type} py-2 mb-0">${escapeHtml(message)}</div>`;
  }

  function showMetaStatus(message, type = 'secondary') {
    const el = document.getElementById('registroMetasGrupoStatus');
    if (el) el.innerHTML = `<div class="alert alert-${type} py-2 mb-0">${escapeHtml(message)}</div>`;
  }

  function setFilterFormMode() {
    const btn = document.getElementById('agregarMetaBtn');
    if (btn) btn.textContent = filtroEditId ? 'Actualizar filtro' : 'Agregar filtro';
    if (cancelFilterEditBtn) cancelFilterEditBtn.style.display = filtroEditId ? 'inline-flex' : 'none';
  }

  function setMetaFormMode() {
    const btn = document.getElementById('crearGrupoMetasBtn');
    if (btn) btn.textContent = metaEditId ? 'Actualizar meta' : 'Crear meta';
    if (cancelMetaEditBtn) cancelMetaEditBtn.style.display = metaEditId ? 'inline-flex' : 'none';
  }

  function resetFilterForm(preserveGroupId = '') {
    filterForm?.reset();
    const groupSelect = document.getElementById('metaGrupoId');
    if (groupSelect) groupSelect.value = preserveGroupId ? String(preserveGroupId) : '';
    filtroEditId = null;
    setFilterFormMode();
  }

  function resetMetaForm() {
    const nameInput = document.getElementById('nombreGrupoMetas');
    const totalInput = document.getElementById('metaTotalCupos');
    if (nameInput) nameInput.value = '';
    if (totalInput) totalInput.value = '';
    metaEditId = null;
    setMetaFormMode();
  }

  function populateGroupSelect(selectedValue = '') {
    const select = document.getElementById('metaGrupoId');
    if (!select) return;
    const current = selectedValue !== '' ? String(selectedValue) : select.value;
    select.innerHTML = '<option value="">-- Sin meta --</option>';
    metasCache.forEach((meta) => {
      const option = document.createElement('option');
      option.value = String(meta.id);
      option.textContent = `${meta.id} - ${meta.nombre_grupo}`;
      select.appendChild(option);
    });
    if (current) select.value = current;
  }

  function renderFilters(rows) {
    if (!filtersTbody) return;
    if (!rows.length) {
      filtersTbody.innerHTML = '<tr><td colspan="9" class="text-center text-muted py-4">Sin filtros registrados</td></tr>';
      return;
    }

    filtersTbody.innerHTML = rows.map((row) => {
      const centerLabel = row.centro_formacion || row.codigo_centro || '—';
      const groupLabel = row.grupo_id ? `Meta ${escapeHtml(row.grupo_nombre || row.grupo_id)}` : '—';
      return `
        <tr>
          <td>
            <input type="checkbox" class="form-check-input meta-select-checkbox" value="${row.id}" ${row.grupo_id ? 'disabled title="Ya pertenece a una meta"' : ''} style="width:22px; height:22px; border:2px solid #444; border-radius:4px; cursor:pointer; box-shadow:0 0 2px rgba(0,0,0,.4);">
          </td>
          <td><strong>${escapeHtml(row.id)}</strong></td>
          <td>${escapeHtml(row.nombre_meta)}</td>
          <td>${escapeHtml(row.tipo_formacion)}</td>
          <td><span class="badge bg-info">${escapeHtml(row.tipo_modalidad)}</span></td>
          <td><small>${formatMultiNumericValue(row.codigo_programa_especial)}</small></td>
          <td><small>${formatMultiNumericValue(row.codigo_convenio)}</small></td>
          <td><small>${escapeHtml(centerLabel)}</small></td>
          <td>${row.grupo_id ? `<span class="badge bg-primary">${groupLabel}</span>` : '<span class="text-muted">—</span>'}</td>
          <td>
            <div class="d-flex flex-wrap gap-1">
              <button type="button" class="btn btn-sm btn-outline-primary filtro-action" data-action="edit" data-id="${row.id}">Editar</button>
              <button type="button" class="btn btn-sm btn-outline-danger filtro-action" data-action="delete" data-id="${row.id}">Eliminar</button>
            </div>
          </td>
        </tr>
      `;
    }).join('');

    document.querySelectorAll('.meta-select-checkbox').forEach((checkbox) => {
      checkbox.addEventListener('change', updateCreateMetaButtonState);
    });
    updateCreateMetaButtonState();
  }

  function renderMetas(rows) {
    if (!metasTbody) return;
    if (!rows.length) {
      metasTbody.innerHTML = '<tr><td colspan="6" class="text-center text-muted py-4">Sin metas creadas</td></tr>';
      return;
    }

    metasTbody.innerHTML = rows.map((row) => `
      <tr>
        <td><strong>${escapeHtml(row.id)}</strong></td>
        <td>${escapeHtml(row.nombre_grupo)}</td>
        <td><span class="badge bg-info text-dark">${escapeHtml(row.cantidad_metas || 0)}</span></td>
        <td><span class="badge bg-success">${escapeHtml(row.total_cupos || 0)}</span></td>
        <td><small>${escapeHtml(row.fecha_creacion || '—')}</small></td>
        <td>
          <div class="d-flex flex-wrap gap-1">
            <button type="button" class="btn btn-sm btn-outline-secondary meta-action" data-action="detail" data-id="${row.id}">Ver filtros</button>
            <button type="button" class="btn btn-sm btn-outline-primary meta-action" data-action="edit" data-id="${row.id}">Editar</button>
            <button type="button" class="btn btn-sm btn-outline-danger meta-action" data-action="delete" data-id="${row.id}">Eliminar</button>
          </div>
        </td>
      </tr>
    `).join('');
  }

  function updateCreateMetaButtonState() {
    if (!createMetaBtn) return;
    createMetaBtn.disabled = metaEditId ? false : getSelectedFilterIds().length === 0;
  }

  function getSelectedFilterIds() {
    return Array.from(document.querySelectorAll('.meta-select-checkbox:checked'))
      .map((el) => Number(el.value))
      .filter((id) => Number.isFinite(id) && id > 0);
  }

  async function loadFilters() {
    try {
      showMetaStatus('Cargando filtros...', 'info');
      const resp = await fetch(`${API_BASE}/registro-metas/lista`);
      const data = await resp.json().catch(() => null);
      if (!resp.ok) throw new Error(data?.detail || `${resp.status} ${resp.statusText}`);
      filtrosCache = Array.isArray(data?.items) ? data.items : [];
      renderFilters(filtrosCache);
      showMetaStatus(`Se cargaron ${data?.total || 0} filtros.`, 'success');
    } catch (error) {
      filtrosCache = [];
      renderFilters([]);
      showMetaStatus(`Error al cargar filtros: ${error.message}`, 'danger');
    }
  }

  async function loadMetas() {
    try {
      const resp = await fetch(`${API_BASE}/registro-metas/grupos`);
      const data = await resp.json().catch(() => null);
      if (!resp.ok) throw new Error(data?.detail || `${resp.status} ${resp.statusText}`);
      metasCache = Array.isArray(data?.items) ? data.items : [];
      renderMetas(metasCache);
      populateGroupSelect();
    } catch (error) {
      metasCache = [];
      renderMetas([]);
      populateGroupSelect();
    }
  }

  function openDetailModal(meta) {
    metaDetalleActual = meta;
    const title = document.getElementById('metaDetalleTitle');
    const metaId = document.getElementById('metaDetalleGrupoId');
    const body = document.getElementById('metaDetalleBody');
    if (title) title.textContent = `Meta ${meta.id} - ${meta.nombre_grupo}`;
    if (metaId) metaId.textContent = String(meta.id);
    if (body) {
      const filtros = Array.isArray(meta.filtros) ? meta.filtros : [];
      body.innerHTML = `
        <div class="row g-3 mb-3">
          <div class="col-md-4"><div class="border rounded-3 p-3 h-100"><div class="text-muted small">Total cupos</div><div class="h4 mb-0">${escapeHtml(meta.total_cupos ?? 0)}</div></div></div>
          <div class="col-md-4"><div class="border rounded-3 p-3 h-100"><div class="text-muted small">Filtros</div><div class="h4 mb-0">${escapeHtml(meta.total_filtros ?? filtros.length)}</div></div></div>
          <div class="col-md-4"><div class="border rounded-3 p-3 h-100"><div class="text-muted small">Creación</div><div class="fw-semibold">${escapeHtml(meta.fecha_creacion || '—')}</div></div></div>
        </div>
        <div class="d-flex justify-content-between align-items-center gap-2 mb-2 flex-wrap">
          <div>
            <div class="fw-semibold">Filtros dentro de la meta</div>
            <div class="text-muted small">Puedes abrir el formulario para agregar otro filtro a esta meta.</div>
          </div>
          <button type="button" class="btn btn-sm btn-success" id="metaDetalleAgregarFiltroBtn">Agregar filtro a esta meta</button>
        </div>
        <div class="table-responsive">
          <table class="table table-sm table-hover align-middle mb-0 table-catalog">
            <thead>
              <tr>
                <th>ID</th>
                <th>Filtro</th>
                <th>Tipo</th>
                <th>Modalidad</th>
                <th>Programa especial</th>
                <th>Convenio</th>
                <th>Centro</th>
              </tr>
            </thead>
            <tbody>
              ${filtros.length ? filtros.map((filtro) => `
                <tr>
                  <td><strong>${escapeHtml(filtro.id)}</strong></td>
                  <td>${escapeHtml(filtro.nombre_meta)}</td>
                  <td>${escapeHtml(filtro.tipo_formacion)}</td>
                  <td><span class="badge bg-info text-dark">${escapeHtml(filtro.tipo_modalidad)}</span></td>
                  <td>${formatMultiNumericValue(filtro.codigo_programa_especial)}</td>
                  <td>${formatMultiNumericValue(filtro.codigo_convenio)}</td>
                  <td>${escapeHtml(filtro.centro_formacion || filtro.codigo_centro || '—')}</td>
                </tr>
              `).join('') : '<tr><td colspan="7" class="text-center text-muted py-4">Esta meta no tiene filtros asociados</td></tr>'}
            </tbody>
          </table>
        </div>
      `;
      document.getElementById('metaDetalleAgregarFiltroBtn')?.addEventListener('click', () => {
        const select = document.getElementById('metaGrupoId');
        if (select) select.value = String(meta.id);
        document.getElementById('registroMetaForm')?.scrollIntoView({ behavior: 'smooth', block: 'start' });
        closeDetailModal();
      });
    }
    detailModal?.classList.add('show');
  }

  function closeDetailModal() {
    detailModal?.classList.remove('show');
    metaDetalleActual = null;
  }

  async function loadMetaDetail(metaId) {
    try {
      showMetaStatus('Cargando detalle de la meta...', 'info');
      const resp = await fetch(`${API_BASE}/registro-metas/grupos/${metaId}`);
      const data = await resp.json().catch(() => null);
      if (!resp.ok) throw new Error(data?.detail || `${resp.status} ${resp.statusText}`);
      openDetailModal(data);
      showMetaStatus(`Meta ${data.nombre_grupo} cargada.`, 'success');
    } catch (error) {
      showMetaStatus(`Error al cargar la meta: ${error.message}`, 'danger');
    }
  }

  async function saveFilter(event) {
    event.preventDefault();
    const tipoFormacion = document.getElementById('metaTipoFormacion')?.value.trim();
    const codigoNivel = document.getElementById('metaCodigoNivelFormacion')?.value.trim();
    const codigoPrograma = document.getElementById('metaCodigoProgramaEspecial')?.value.trim();
    const codigoConvenio = document.getElementById('metaCodigoConvenio')?.value.trim();
    const tipoModalidad = document.getElementById('metaTipoModalidad')?.value.trim();
    const nombreMeta = document.getElementById('metaNombre')?.value.trim();
    const grupoId = document.getElementById('metaGrupoId')?.value.trim();
    const codigoCentro = document.getElementById('metaCentroFormacion')?.value.trim();

    if (!tipoFormacion || !tipoModalidad || !nombreMeta || !codigoCentro) {
      showFilterStatus('Completa todos los campos obligatorios.', 'warning');
      return;
    }
    if (!/^\d+$/.test(codigoNivel)) {
      showFilterStatus('El código del nivel de formación solo debe contener números.', 'warning');
      return;
    }
    if (!codigoPrograma || !codigoConvenio) {
      showFilterStatus('Completa los campos de programa especial y convenio.', 'warning');
      return;
    }

    const formData = new FormData();
    formData.append('tipo_formacion', tipoFormacion);
    formData.append('codigo_nivel_formacion', codigoNivel);
    formData.append('codigo_programa_especial', normalizeMultiNumericInput(codigoPrograma));
    formData.append('codigo_convenio', normalizeMultiNumericInput(codigoConvenio));
    formData.append('tipo_modalidad', tipoModalidad);
    formData.append('nombre_meta', nombreMeta);
    formData.append('codigo_centro', codigoCentro);
    formData.append('centro_formacion', document.getElementById('metaCentroFormacion')?.selectedOptions?.[0]?.textContent || codigoCentro);
    if (grupoId) formData.append('grupo_id', grupoId);

    try {
      const editing = Boolean(filtroEditId);
      showFilterStatus(editing ? 'Actualizando filtro...' : 'Guardando filtro...', 'info');
      const url = editing ? `${API_BASE}/registro-metas/filtros/${filtroEditId}` : `${API_BASE}/registro-metas/crear`;
      const resp = await fetch(url, { method: editing ? 'PUT' : 'POST', body: formData });
      const data = await resp.json().catch(() => null);
      if (!resp.ok) throw new Error(data?.detail || `${resp.status} ${resp.statusText}`);
      showFilterStatus(editing ? `Filtro ${data.id} actualizado correctamente.` : `Filtro registrado correctamente (ID ${data.id}).`, 'success');
      resetFilterForm(grupoId || '');
      await loadFilters();
      await loadMetas();
    } catch (error) {
      showFilterStatus(`Error al guardar filtro: ${error.message}`, 'danger');
    }
  }

  async function saveMeta() {
    const nombreGrupo = (document.getElementById('nombreGrupoMetas')?.value || '').trim();
    const totalCupos = (document.getElementById('metaTotalCupos')?.value || '').trim();
    if (!nombreGrupo) {
      showMetaStatus('Escribe un nombre para la meta.', 'warning');
      return;
    }
    if (!/^\d+$/.test(totalCupos)) {
      showMetaStatus('El total de cupos debe contener solo números.', 'warning');
      return;
    }

    const selectedIds = metaEditId ? [] : getSelectedFilterIds();
    if (!metaEditId && !selectedIds.length) {
      showMetaStatus('Selecciona al menos un filtro para crear la meta.', 'warning');
      return;
    }

    try {
      const editing = Boolean(metaEditId);
      showMetaStatus(editing ? 'Actualizando meta...' : 'Creando meta...', 'info');
      const payload = { nombre_grupo: nombreGrupo, total_cupos: Number(totalCupos) };
      if (!editing) payload.meta_ids = selectedIds;
      const url = editing ? `${API_BASE}/registro-metas/grupos/${metaEditId}` : `${API_BASE}/registro-metas/crear-grupo`;
      const resp = await fetch(url, {
        method: editing ? 'PUT' : 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      const data = await resp.json().catch(() => null);
      if (!resp.ok) throw new Error(data?.detail || `${resp.status} ${resp.statusText}`);
      showMetaStatus(editing ? `Meta "${nombreGrupo}" actualizada. Total cupos: ${data.total_cupos}.` : `Meta "${nombreGrupo}" creada. Filtros: ${data.cantidad_metas}, total cupos: ${data.total_cupos}.`, 'success');
      resetMetaForm();
      document.getElementById('selectAllMetasCheckbox') && (document.getElementById('selectAllMetasCheckbox').checked = false);
      await loadFilters();
      await loadMetas();
    } catch (error) {
      showMetaStatus(`Error al guardar meta: ${error.message}`, 'danger');
    }
  }

  function openFilterEdit(row) {
    if (!row) return;
    filtroEditId = Number(row.id);
    document.getElementById('metaTipoFormacion').value = row.tipo_formacion || '';
    document.getElementById('metaCodigoNivelFormacion').value = row.codigo_nivel_formacion || '';
    document.getElementById('metaCodigoProgramaEspecial').value = String(row.codigo_programa_especial || '').replace(/,/g, ', ');
    document.getElementById('metaCodigoConvenio').value = String(row.codigo_convenio || '').replace(/,/g, ', ');
    document.getElementById('metaTipoModalidad').value = row.tipo_modalidad || '';
    document.getElementById('metaNombre').value = row.nombre_meta || '';
    document.getElementById('metaGrupoId').value = row.grupo_id ? String(row.grupo_id) : '';
    document.getElementById('metaCentroFormacion').value = row.codigo_centro || '';
    showFilterStatus(`Editando filtro ${row.id}.`, 'info');
    setFilterFormMode();
    document.getElementById('registroMetaForm')?.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  function openMetaEdit(row) {
    if (!row) return;
    metaEditId = Number(row.id);
    document.getElementById('nombreGrupoMetas').value = row.nombre_grupo || '';
    document.getElementById('metaTotalCupos').value = row.total_cupos ?? '';
    showMetaStatus(`Editando meta ${row.id}.`, 'info');
    setMetaFormMode();
    document.getElementById('crearGrupoMetasBtn')?.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  async function deleteFilter(id) {
    if (!window.confirm('¿Deseas eliminar este filtro?')) return;
    try {
      showFilterStatus('Eliminando filtro...', 'info');
      const resp = await fetch(`${API_BASE}/registro-metas/filtros/${id}`, { method: 'DELETE' });
      const data = await resp.json().catch(() => null);
      if (!resp.ok) throw new Error(data?.detail || `${resp.status} ${resp.statusText}`);
      showFilterStatus(`Filtro ${id} eliminado.`, 'success');
      await loadFilters();
      await loadMetas();
    } catch (error) {
      showFilterStatus(`Error al eliminar filtro: ${error.message}`, 'danger');
    }
  }

  async function deleteMeta(id) {
    if (!window.confirm('¿Deseas eliminar esta meta? Los filtros quedarán sin meta asociada.')) return;
    try {
      showMetaStatus('Eliminando meta...', 'info');
      const resp = await fetch(`${API_BASE}/registro-metas/grupos/${id}`, { method: 'DELETE' });
      const data = await resp.json().catch(() => null);
      if (!resp.ok) throw new Error(data?.detail || `${resp.status} ${resp.statusText}`);
      showMetaStatus(`Meta ${id} eliminada.`, 'success');
      if (metaEditId && Number(metaEditId) === Number(id)) resetMetaForm();
      closeDetailModal();
      await loadFilters();
      await loadMetas();
    } catch (error) {
      showMetaStatus(`Error al eliminar meta: ${error.message}`, 'danger');
    }
  }

  function updateMetaSelectionButton() {
    if (!createMetaBtn) return;
    createMetaBtn.disabled = metaEditId ? false : getSelectedFilterIds().length === 0;
  }

  function getSelectedFilterIds() {
    return Array.from(document.querySelectorAll('.meta-select-checkbox:checked'))
      .map((el) => Number(el.value))
      .filter((id) => Number.isFinite(id) && id > 0);
  }

  document.addEventListener('click', (event) => {
    const button = event.target.closest('button.filtro-action, button.meta-action');
    if (!button) return;
    const id = Number(button.getAttribute('data-id'));
    const action = button.getAttribute('data-action');

    if (button.classList.contains('filtro-action')) {
      const row = filtrosCache.find((item) => Number(item.id) === id);
      if (action === 'edit') openFilterEdit(row);
      if (action === 'delete') deleteFilter(id);
    }

    if (button.classList.contains('meta-action')) {
      const row = metasCache.find((item) => Number(item.id) === id);
      if (action === 'detail') loadMetaDetail(id);
      if (action === 'edit') openMetaEdit(row);
      if (action === 'delete') deleteMeta(id);
    }
  });

  filterForm?.addEventListener('submit', saveFilter);
  reloadFiltersBtn?.addEventListener('click', loadFilters);
  reloadMetasBtn?.addEventListener('click', loadMetas);
  createMetaBtn?.addEventListener('click', saveMeta);
  cancelFilterEditBtn?.addEventListener('click', () => {
    resetFilterForm(document.getElementById('metaGrupoId')?.value || '');
    showFilterStatus('Edición cancelada.', 'secondary');
  });
  cancelMetaEditBtn?.addEventListener('click', () => {
    resetMetaForm();
    showMetaStatus('Edición cancelada.', 'secondary');
  });
  detailCloseBtn?.addEventListener('click', closeDetailModal);
  detailModal?.addEventListener('click', (event) => {
    if (event.target && event.target.id === 'metaDetalleModal') {
      closeDetailModal();
    }
  });
  selectAllCheckbox?.addEventListener('change', (event) => {
    const checked = event.target.checked;
    document.querySelectorAll('.meta-select-checkbox:not(:disabled)').forEach((checkbox) => {
      checkbox.checked = checked;
    });
    updateMetaSelectionButton();
  });

  setFilterFormMode();
  setMetaFormMode();
  loadFilters();
  loadMetas();
});
