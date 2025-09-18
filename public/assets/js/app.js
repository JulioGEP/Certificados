(() => {
  const STORAGE_KEY = 'gep-certificados/session/v1';
  const TABLE_COLUMNS = [
    { field: 'presupuesto', label: 'Presupuesto', type: 'text', placeholder: 'ID del deal' },
    { field: 'nombre', label: 'Nombre', type: 'text', placeholder: 'Nombre del alumno' },
    { field: 'apellido', label: 'Apellidos', type: 'text', placeholder: 'Apellidos del alumno' },
    { field: 'dni', label: 'DNI / NIE', type: 'text', placeholder: 'Documento' },
    { field: 'fecha', label: 'Fecha', type: 'date' },
    { field: 'lugar', label: 'Lugar', type: 'text', placeholder: 'Sede de la formación' },
    { field: 'duracion', label: 'Duración', type: 'number', placeholder: 'Horas' },
    { field: 'cliente', label: 'Cliente', type: 'text', placeholder: 'Empresa' },
    { field: 'formacion', label: 'Formación', type: 'text', placeholder: 'Título de la formación' }
  ];

  const TABLE_COLUMN_LOOKUP = TABLE_COLUMNS.reduce((lookup, column) => {
    lookup.set(column.field, column);
    return lookup;
  }, new Map());

  const OPEN_TRAINING_DURATION_ENTRIES = [
    ['Pack Emergencias', '6'],
    ['Trabajos en Altura', '8'],
    ['Trabajos Verticales', '12'],
    ['Carretilla elevadora', '8'],
    ['Espacios Confinados', '8'],
    ['Operaciones Telco', '6'],
    ['Riesgo Eléctrico Telco', '6'],
    ['Espacios Confinados Telco', '6'],
    ['Trabajos en altura Telco', '6'],
    ['Basico de Fuego', '4'],
    ['Avanzado de Fuego', '5'],
    ['Avanzado y Casa de Humo', '6'],
    ['Riesgo Químico', '4'],
    ['Primeros Auxilios', '4'],
    ['SVD y DEA', '6'],
    ['Implantación de PAU', '6'],
    ['Jefes de Emergencias', '8'],
    ['Curso de ERA\'s', '8'],
    ['Andamios', '8'],
    ['Renovación Bombero de Empresa', '20'],
    ['Bombero de Empresa Inicial', '350']
  ];

  const TRAINING_DURATION_LOOKUP = OPEN_TRAINING_DURATION_ENTRIES.reduce((lookup, [name, hours]) => {
    const normalisedName = normaliseTrainingName(name);
    if (normalisedName && !lookup.has(normalisedName)) {
      lookup.set(normalisedName, hours);
    }
    return lookup;
  }, new Map());

  const elements = {
    budgetForm: document.getElementById('budget-form'),
    budgetInput: document.getElementById('budget-input'),
    fillButton: document.getElementById('fill-button'),
    addManualRow: document.getElementById('add-manual-row'),
    clearStorage: document.getElementById('clear-storage'),
    tableBody: document.getElementById('table-body'),
    alertContainer: document.getElementById('alert-container')
  };

  const state = {
    rows: [],
    isLoading: false
  };

  document.addEventListener('DOMContentLoaded', () => {
    hydrateFromStorage();
    renderTable();
    bindEvents();
  });

  function bindEvents() {
    elements.budgetForm.addEventListener('submit', handleBudgetSubmit);
    elements.addManualRow.addEventListener('click', addEmptyRow);
    elements.clearStorage.addEventListener('click', clearAllRows);
  }

  function hydrateFromStorage() {
    try {
      const raw = window.sessionStorage.getItem(STORAGE_KEY);
      if (!raw) return;
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed)) {
        state.rows = parsed.map((row) => {
          const hydratedRow = { ...createEmptyRow(), ...row };
          if (hydratedRow.duracion === undefined || hydratedRow.duracion === null || hydratedRow.duracion === '') {
            const duration = getTrainingDuration(hydratedRow.formacion);
            if (duration) {
              hydratedRow.duracion = duration;
            }
          }
          return hydratedRow;
        });
      }
    } catch (error) {
      console.error('No se ha podido recuperar la información guardada', error);
    }
  }

  function persistRows() {
    try {
      window.sessionStorage.setItem(STORAGE_KEY, JSON.stringify(state.rows));
    } catch (error) {
      console.error('No se ha podido guardar la información', error);
      showAlert('warning', 'No se ha podido guardar la información en esta sesión.');
    }
  }

  function createEmptyRow() {
    return {
      presupuesto: '',
      nombre: '',
      apellido: '',
      dni: '',
      documentType: '',
      fecha: '',
      lugar: '',
      duracion: '',
      cliente: '',
      formacion: '',
      personaContacto: '',
      correoContacto: ''
    };
  }

  function renderTable() {
    const { tableBody } = elements;
    tableBody.innerHTML = '';

    if (!state.rows.length) {
      const emptyRow = document.createElement('tr');
      emptyRow.classList.add('empty-state');
      const emptyCell = document.createElement('td');
      emptyCell.colSpan = TABLE_COLUMNS.length + 1;
      emptyCell.textContent = 'Todavía no has añadido ningún alumno. Recupera un presupuesto o añade una fila manual.';
      emptyRow.appendChild(emptyCell);
      tableBody.appendChild(emptyRow);
      updateClearButtonState();
      return;
    }

    state.rows.forEach((row, index) => {
      const tr = document.createElement('tr');

      TABLE_COLUMNS.forEach((column) => {
        const td = document.createElement('td');
        td.dataset.label = column.label;

        if (column.field === 'dni') {
          const wrapper = document.createElement('div');
          wrapper.className = 'document-wrapper';

          const input = createInput(column, row[column.field], index);
          const badge = document.createElement('span');
          badge.className = 'badge rounded-pill text-bg-info-subtle document-badge';
          updateDocumentBadge(badge, row.documentType);

          input.addEventListener('input', (event) => {
            const value = event.target.value;
            updateRowValue(index, 'dni', value);
            updateCellTooltip(td, input, column, value);
            const documentType = detectDocumentType(value);
            updateRowValue(index, 'documentType', documentType);
            updateDocumentBadge(badge, documentType);
          });

          wrapper.appendChild(input);
          wrapper.appendChild(badge);
          td.appendChild(wrapper);
          updateCellTooltip(td, input, column, input.value);
        } else {
          const input = createInput(column, row[column.field], index);
          input.addEventListener('input', (event) => {
            const { value } = event.target;
            updateRowValue(index, column.field, value);
            updateCellTooltip(td, input, column, value);
            if (column.field === 'formacion') {
              applyTrainingDuration(index, event.target.value);
            }
          });
          if (column.field === 'fecha') {
            input.addEventListener('change', (event) => {
              const { value } = event.target;
              updateRowValue(index, column.field, value);
              updateCellTooltip(td, input, column, value);
            });
          }
          td.appendChild(input);
          updateCellTooltip(td, input, column, input.value);
        }

        tr.appendChild(td);
      });

      const actionsTd = document.createElement('td');
      actionsTd.className = 'text-center';
      const actionsWrapper = document.createElement('div');
      actionsWrapper.className = 'd-flex flex-column flex-sm-row gap-2 justify-content-center';

      const pdfButton = document.createElement('button');
      pdfButton.type = 'button';
      pdfButton.className = 'btn btn-primary btn-sm';
      pdfButton.textContent = 'Generar PDF';
      pdfButton.addEventListener('click', () => handleGeneratePdf(index, pdfButton));

      const deleteButton = document.createElement('button');
      deleteButton.type = 'button';
      deleteButton.className = 'btn btn-outline-danger btn-sm';
      deleteButton.textContent = 'Eliminar';
      deleteButton.addEventListener('click', () => removeRow(index));

      actionsWrapper.appendChild(pdfButton);
      actionsWrapper.appendChild(deleteButton);
      actionsTd.appendChild(actionsWrapper);
      tr.appendChild(actionsTd);

      tableBody.appendChild(tr);
    });

    updateClearButtonState();
  }

  function resolveTooltipText(column, rawValue) {
    if (!column) {
      return '';
    }

    if (column.type === 'date') {
      const normalisedDate = normaliseDateValue(rawValue);
      if (normalisedDate) {
        return normalisedDate;
      }
    }

    if (rawValue === undefined || rawValue === null) {
      return column.placeholder || '';
    }

    const value = typeof rawValue === 'string' ? rawValue.trim() : String(rawValue).trim();
    if (value !== '') {
      return value;
    }

    return column.placeholder || '';
  }

  function updateCellTooltip(cell, input, column, rawValue) {
    if (!cell || !input) {
      return;
    }

    const tooltipText = resolveTooltipText(column, rawValue);
    if (tooltipText) {
      cell.dataset.tooltip = tooltipText;
      input.title = tooltipText;
    } else {
      delete cell.dataset.tooltip;
      input.removeAttribute('title');
    }
  }

  function createInput(column, value, index) {
    const input = document.createElement('input');
    input.className = 'form-control';
    input.type = column.type === 'number' ? 'number' : column.type;
    input.placeholder = column.placeholder || '';
    input.value = column.type === 'date' ? normaliseDateValue(value) : value || '';
    input.dataset.field = column.field;
    input.dataset.index = String(index);
    if (column.label) {
      input.setAttribute('aria-label', column.label);
    }
    if (column.type === 'number') {
      input.min = '0';
      input.step = '0.5';
      input.inputMode = 'decimal';
    }

    return input;
  }

  function updateRowValue(index, field, value, reRender = false) {
    state.rows[index][field] = value;
    persistRows();
    if (reRender === true) {
      // Mantener la posición del cursor no es crítico en la mayoría de campos, por lo que re-renderizamos para asegurar la coherencia.
      renderTable();
    }
  }

  function detectDocumentType(value) {
    if (!value) return '';
    const clean = value.trim().toUpperCase();
    const dniPattern = /^[0-9]{8}[A-Z]$/;
    const niePattern = /^[XYZ][0-9]{7}[A-Z]$/;

    if (dniPattern.test(clean)) {
      return 'DNI';
    }

    if (niePattern.test(clean)) {
      return 'NIE';
    }

    if (/^[A-Z]/.test(clean) && /[A-Z]$/.test(clean)) {
      return 'NIE';
    }

    if (/[A-Z]$/.test(clean)) {
      return 'DNI';
    }

    return '';
  }

  function updateDocumentBadge(badge, documentType) {
    if (!documentType) {
      badge.classList.add('d-none');
      badge.textContent = '';
      return;
    }

    badge.classList.remove('d-none');
    badge.textContent = documentType;
  }

  function handleGeneratePdf(rowIndex, triggerButton) {
    const row = state.rows[rowIndex];
    if (!row) {
      showAlert('danger', 'No se ha encontrado la fila seleccionada.');
      return;
    }

    if (!window.certificatePdf || typeof window.certificatePdf.generate !== 'function') {
      showAlert('danger', 'La librería de certificados no está disponible.');
      return;
    }

    let originalLabel = '';
    if (triggerButton instanceof HTMLButtonElement) {
      originalLabel = triggerButton.textContent;
      triggerButton.disabled = true;
      triggerButton.textContent = 'Generando...';
    }

    window.certificatePdf
      .generate(row)
      .then(() => {
        showAlert('success', 'Certificado generado correctamente.');
      })
      .catch((error) => {
        console.error('No se ha podido generar el certificado PDF', error);
        showAlert('danger', 'No se ha podido generar el certificado en PDF. Revisa los datos e inténtalo de nuevo.');
      })
      .finally(() => {
        if (triggerButton instanceof HTMLButtonElement) {
          triggerButton.disabled = false;
          triggerButton.textContent = originalLabel || 'Generar PDF';
        }
      });
  }

  async function handleBudgetSubmit(event) {
    event.preventDefault();
    const dealIds = getBudgetIdsFromInput(elements.budgetInput.value);

    if (!dealIds.length) {
      showAlert('warning', 'Introduce uno o varios números de presupuesto válidos separados por comas.');
      return;
    }

    const { existingDealIds, newDealIds } = splitDealIdsByExistingState(dealIds);

    if (existingDealIds.length > 0) {
      const joinedIds = existingDealIds.join(', ');
      const label = existingDealIds.length > 1 ? 'Los presupuestos' : 'El presupuesto';
      const suffix = existingDealIds.length > 1 ? 'ya están añadidos' : 'ya está añadido';
      showAlert('warning', `${label} ${joinedIds} ${suffix} en el listado.`);
    }

    if (!newDealIds.length) {
      elements.budgetInput.focus();
      return;
    }

    setLoading(true);

    try {
      const results = await Promise.allSettled(newDealIds.map((dealId) => fetchBudgetData(dealId)));

      const successfulDeals = [];
      const failedDeals = [];

      results.forEach((result, index) => {
        const dealId = newDealIds[index];
        if (result.status === 'fulfilled') {
          successfulDeals.push(dealId);
          addDealToTable(dealId, result.value);
        } else {
          const message = result.reason && result.reason.message ? result.reason.message : 'No se ha podido recuperar la información del presupuesto.';
          const erroredDealId = (result.reason && result.reason.dealId) || dealId;
          const logMessage = erroredDealId
            ? `No se ha podido recuperar la información del presupuesto ${erroredDealId}`
            : 'No se ha podido recuperar la información del presupuesto';
          console.error(logMessage, result.reason);
          failedDeals.push({ dealId: erroredDealId, message });
        }
      });

      if (successfulDeals.length > 0) {
        const joinedIds = successfulDeals.join(', ');
        const label = successfulDeals.length > 1 ? 'los presupuestos' : 'el presupuesto';
        showAlert('success', `Información de ${label} ${joinedIds} añadida correctamente.`);
        elements.budgetInput.value = '';
        elements.budgetInput.focus();
      }

      if (failedDeals.length > 0) {
        const joinedIds = failedDeals
          .map((item) => item.dealId)
          .filter(Boolean)
          .join(', ');
        const label = failedDeals.length > 1 ? 'los presupuestos' : 'el presupuesto';
        const detail = failedDeals[0].message || 'No se ha podido recuperar la información solicitada.';
        const suffix = joinedIds ? ` ${joinedIds}` : '';
        showAlert('danger', `No se ha podido recuperar la información de ${label}${suffix}. ${detail}`);
      }
    } catch (error) {
      console.error(error);
      showAlert('danger', 'Ha ocurrido un error al recuperar la información de los presupuestos. Inténtalo de nuevo.');
    } finally {
      setLoading(false);
    }
  }

  function getBudgetIdsFromInput(rawValue) {
    if (!rawValue) return [];
    const segments = rawValue
      .split(',')
      .map((segment) => segment.trim())
      .filter((segment) => segment.length > 0);
    return Array.from(new Set(segments));
  }

  function splitDealIdsByExistingState(dealIds) {
    return dealIds.reduce(
      (acc, dealId) => {
        if (isDealAlreadyInState(dealId)) {
          acc.existingDealIds.push(dealId);
        } else {
          acc.newDealIds.push(dealId);
        }
        return acc;
      },
      { existingDealIds: [], newDealIds: [] }
    );
  }

  function isDealAlreadyInState(dealId) {
    const normalisedDealId = normaliseDealId(dealId);
    if (!normalisedDealId) return false;
    return state.rows.some((row) => normaliseDealId(row.presupuesto) === normalisedDealId);
  }

  function normaliseDealId(value) {
    if (value === undefined || value === null) {
      return '';
    }
    return String(value).trim().toLowerCase();
  }

  async function fetchBudgetData(dealId) {
    try {
      const response = await fetch(`/.netlify/functions/fetch-deal?dealId=${encodeURIComponent(dealId)}`);
      const payload = await response.json();

      if (!response.ok || payload.success === false) {
        const message = payload && payload.message ? payload.message : 'No se ha podido recuperar la información del presupuesto.';
        const error = new Error(message);
        error.dealId = dealId;
        throw error;
      }

      return payload.data;
    } catch (error) {
      const enrichedError = error instanceof Error ? error : new Error('Ha ocurrido un error al recuperar la información del presupuesto.');
      if (!enrichedError.message || enrichedError.message === 'Failed to fetch') {
        enrichedError.message = 'Ha ocurrido un error al conectar con el servicio. Inténtalo de nuevo.';
      }
      if (!enrichedError.dealId) {
        enrichedError.dealId = dealId;
      }
      throw enrichedError;
    }
  }

  function addDealToTable(dealId, data) {
    const baseRow = {
      ...createEmptyRow(),
      presupuesto: dealId,
      fecha: normaliseDateValue(data.trainingDate),
      lugar: data.trainingLocation || '',
      duracion: getTrainingDuration(data.trainingName),
      cliente: data.clientName || '',
      formacion: data.trainingName || '',
      personaContacto: data.contactName || '',
      correoContacto: data.contactEmail || ''
    };

    const students = Array.isArray(data.students) ? data.students : [];

    if (!students.length) {
      state.rows.push({ ...baseRow });
    } else {
      students.forEach((student) => {
        state.rows.push({
          ...baseRow,
          nombre: student.name || '',
          apellido: student.surname || '',
          dni: student.document || '',
          documentType: detectDocumentType(student.document || '')
        });
      });
    }

    persistRows();
    renderTable();
  }

  function addEmptyRow() {
    state.rows.push(createEmptyRow());
    persistRows();
    renderTable();
    showAlert('info', 'Fila manual añadida. Completa los datos cuando quieras.');
  }

  function removeRow(index) {
    state.rows.splice(index, 1);
    persistRows();
    renderTable();
    showAlert('info', 'Fila eliminada.');
  }

  function clearAllRows() {
    if (!state.rows.length) {
      showAlert('info', 'No hay filas que eliminar.');
      return;
    }

    const confirmation = window.confirm('¿Quieres vaciar por completo el listado de alumnos? Esta acción no se puede deshacer.');
    if (!confirmation) return;

    state.rows = [];
    persistRows();
    renderTable();
    showAlert('success', 'Listado vaciado correctamente.');
  }

  function updateClearButtonState() {
    elements.clearStorage.disabled = state.rows.length === 0;
  }

  function getTrainingDuration(trainingName) {
    const normalisedName = normaliseTrainingName(trainingName);
    if (!normalisedName) {
      return '';
    }
    return TRAINING_DURATION_LOOKUP.get(normalisedName) || '';
  }

  function normaliseTrainingName(value) {
    if (!value) return '';
    return value
      .toString()
      .trim()
      .toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '');
  }

  function applyTrainingDuration(rowIndex, trainingValue) {
    if (!state.rows[rowIndex]) return;
    const duration = getTrainingDuration(trainingValue);
    if (duration === '' || state.rows[rowIndex].duracion === duration) {
      return;
    }
    updateRowValue(rowIndex, 'duracion', duration);
    const durationInput = elements.tableBody.querySelector(
      `input[data-index="${rowIndex}"][data-field="duracion"]`
    );
    if (durationInput && durationInput.value !== duration) {
      durationInput.value = duration;
      const durationCell = durationInput.closest('td');
      const durationColumn = TABLE_COLUMN_LOOKUP.get('duracion');
      if (durationCell && durationColumn) {
        updateCellTooltip(durationCell, durationInput, durationColumn, duration);
      }
    }
  }

  function normaliseDateValue(value) {
    if (!value) return '';
    if (/^\d{4}-\d{2}-\d{2}$/.test(value)) {
      return value;
    }

    const parsed = new Date(value);
    if (Number.isNaN(parsed.getTime())) {
      return '';
    }

    return parsed.toISOString().split('T')[0];
  }

  function setLoading(isLoading) {
    state.isLoading = isLoading;
    const spinner = elements.fillButton.querySelector('.spinner-border');
    const buttonText = elements.fillButton.querySelector('.button-text');

    elements.fillButton.disabled = isLoading;
    elements.budgetInput.disabled = isLoading;

    if (isLoading) {
      spinner.classList.remove('d-none');
      buttonText.textContent = 'Cargando';
    } else {
      spinner.classList.add('d-none');
      buttonText.textContent = 'Añadir alumn@/s';
    }
  }

  function showAlert(type, message) {
    if (!elements.alertContainer) return;

    const alert = document.createElement('div');
    alert.className = `alert alert-${type} alert-dismissible fade show shadow-sm`;
    alert.setAttribute('role', 'alert');
    alert.textContent = message;

    const closeButton = document.createElement('button');
    closeButton.type = 'button';
    closeButton.className = 'btn-close';
    closeButton.setAttribute('data-bs-dismiss', 'alert');
    closeButton.setAttribute('aria-label', 'Cerrar');
    alert.appendChild(closeButton);

    elements.alertContainer.appendChild(alert);

    window.setTimeout(() => {
      try {
        if (window.bootstrap && window.bootstrap.Alert) {
          const instance = window.bootstrap.Alert.getOrCreateInstance(alert);
          instance.close();
        } else {
          alert.classList.remove('show');
          alert.addEventListener('transitionend', () => alert.remove(), { once: true });
        }
      } catch (error) {
        alert.remove();
      }
    }, 6000);
  }
})();
