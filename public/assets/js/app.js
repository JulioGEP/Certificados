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
        state.rows = parsed.map((row) => ({ ...createEmptyRow(), ...row }));
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
            const documentType = detectDocumentType(value);
            updateRowValue(index, 'documentType', documentType);
            updateDocumentBadge(badge, documentType);
          });

          wrapper.appendChild(input);
          wrapper.appendChild(badge);
          td.appendChild(wrapper);
        } else {
          const input = createInput(column, row[column.field], index);
          input.addEventListener('input', (event) => {
            updateRowValue(index, column.field, event.target.value);
          });
          td.appendChild(input);
        }

        tr.appendChild(td);
      });

      const actionsTd = document.createElement('td');
      actionsTd.className = 'text-center';
      const deleteButton = document.createElement('button');
      deleteButton.type = 'button';
      deleteButton.className = 'btn btn-outline-danger btn-sm';
      deleteButton.innerHTML = '<span aria-hidden="true">&times;</span> Eliminar';
      deleteButton.addEventListener('click', () => removeRow(index));
      actionsTd.appendChild(deleteButton);
      tr.appendChild(actionsTd);

      tableBody.appendChild(tr);
    });

    updateClearButtonState();
  }

  function createInput(column, value, index) {
    const input = document.createElement('input');
    input.className = 'form-control';
    input.type = column.type === 'number' ? 'number' : column.type;
    input.placeholder = column.placeholder || '';
    input.value = column.type === 'date' ? normaliseDateValue(value) : value || '';
    if (column.type === 'number') {
      input.min = '0';
      input.step = '0.5';
      input.inputMode = 'decimal';
    }

    if (column.field === 'fecha') {
      input.addEventListener('change', (event) => {
        updateRowValue(index, 'fecha', event.target.value);
      });
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
      buttonText.textContent = 'Rellenar';
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
