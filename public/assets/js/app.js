(() => {
  const STORAGE_KEY = 'gep-certificados/session/v1';
  const CONTACTS_STORAGE_KEY = 'gep-certificados/deal-contacts/v1';
  const trainingTemplates = window.trainingTemplates || null;
  const IRATA_TRAINERS = [
    'Cristobal Corredor Navarro: 1/248468',
    'Laura Medina Matías : 1/247658',
    'Luís Vicent Pérez - Irata: 1/175398',
    'Isaac El Allaoui Algaba: Técnico Irata: 1/248469'
  ];
  const TRABAJOS_VERTICALES_KEY = 'trabajos verticales';
  const TABLE_COLUMNS = [
    { field: 'presupuesto', label: 'Presu', type: 'text', placeholder: 'ID del deal' },
    { field: 'nombre', label: 'Nombre', type: 'text', placeholder: 'Nombre del alumno' },
    { field: 'apellido', label: 'Apellidos', type: 'text', placeholder: 'Apellidos del alumno' },
    { field: 'dni', label: 'DNI / NIE', type: 'text', placeholder: 'Documento' },
    { field: 'fecha', label: 'Fecha', type: 'date' },
    { field: 'segundaFecha', label: '2ª Fecha', type: 'date' },
    { field: 'lugar', label: 'Lugar', type: 'text', placeholder: 'Sede de la formación' },
    { field: 'duracion', label: 'Horas', type: 'text', placeholder: 'Ej. 6h' },
    { field: 'cliente', label: 'Cliente', type: 'text', placeholder: 'Empresa' },
    { field: 'formacion', label: 'Formación', type: 'text', placeholder: 'Título de la formación' },
    {
      field: 'irata',
      label: 'IRATA',
      type: 'select',
      placeholder: 'Selecciona formador/a',
      options: IRATA_TRAINERS
    }
  ];

  const TABLE_COLUMN_LOOKUP = TABLE_COLUMNS.reduce((lookup, column) => {
    lookup.set(column.field, column);
    return lookup;
  }, new Map());

  const CERTIFICATE_BUTTON_LABEL = 'Certificado';
  const GENERATE_ALL_CERTIFICATES_LABEL = 'Generar Todos los Certificados';

  const elements = {
    budgetForm: document.getElementById('budget-form'),
    budgetInput: document.getElementById('budget-input'),
    fillButton: document.getElementById('fill-button'),
    addManualRow: document.getElementById('add-manual-row'),
    clearStorage: document.getElementById('clear-storage'),
    generateAllCertificates: document.getElementById('generate-all-certificates'),
    tableBody: document.getElementById('table-body'),
    alertContainer: document.getElementById('alert-container'),
    manageTemplatesButton: document.getElementById('manage-templates-button'),
    templatesModal: document.getElementById('training-templates-modal'),
    templateForm: document.getElementById('training-template-form'),
    templateSelector: document.getElementById('template-selector'),
    deleteTemplateButton: document.getElementById('delete-template-button'),
    templateNameInput: document.getElementById('template-name'),
    templateTitleInput: document.getElementById('template-title'),
    templateDurationInput: document.getElementById('template-duration'),
    addTheoryPoint: document.getElementById('add-theory-point'),
    addPracticePoint: document.getElementById('add-practice-point'),
    theoryList: document.getElementById('theory-list'),
    practiceList: document.getElementById('practice-list')
  };

  const state = {
    rows: [],
    isLoading: false,
    dealContacts: new Map()
  };

  const templateState = {
    currentTemplateId: '',
    lastSelectedTemplateId: '',
    modalInstance: null,
    unsubscribe: null
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
    elements.generateAllCertificates.addEventListener('click', handleGenerateAllCertificates);
    setupTemplateManagement();
  }

  function setupTemplateManagement() {
    if (!trainingTemplates) {
      return;
    }

    const {
      manageTemplatesButton,
      templatesModal,
      templateForm,
      templateSelector,
      deleteTemplateButton,
      templateNameInput,
      templateTitleInput,
      templateDurationInput,
      addTheoryPoint,
      addPracticePoint,
      theoryList,
      practiceList
    } = elements;

    if (
      !manageTemplatesButton ||
      !templatesModal ||
      !templateForm ||
      !templateSelector ||
      !templateNameInput ||
      !templateTitleInput ||
      !templateDurationInput ||
      !theoryList ||
      !practiceList
    ) {
      return;
    }

    const bootstrapModal = window.bootstrap && window.bootstrap.Modal ? window.bootstrap.Modal : null;
    if (!bootstrapModal) {
      console.warn('Bootstrap Modal no disponible para la gestión de plantillas.');
      return;
    }

    if (!templateState.modalInstance) {
      templateState.modalInstance = bootstrapModal.getOrCreateInstance(templatesModal);
    }

    manageTemplatesButton.addEventListener('click', openTemplatesModal);
    templateSelector.addEventListener('change', handleTemplateSelectionChange);
    templateForm.addEventListener('submit', handleTemplateFormSubmit);

    if (deleteTemplateButton) {
      deleteTemplateButton.addEventListener('click', handleDeleteTemplateClick);
    }

    if (addTheoryPoint) {
      addTheoryPoint.addEventListener('click', () => addTemplatePoint('theory'));
    }

    if (addPracticePoint) {
      addPracticePoint.addEventListener('click', () => addTemplatePoint('practice'));
    }

    theoryList.addEventListener('click', (event) => handleTemplatePointListClick(event));
    practiceList.addEventListener('click', (event) => handleTemplatePointListClick(event));

    if (templateState.unsubscribe) {
      templateState.unsubscribe();
      templateState.unsubscribe = null;
    }

    templateState.unsubscribe = trainingTemplates.subscribe(() => {
      handleTemplateLibraryUpdated();
    });

    populateTemplateSelector(templateState.lastSelectedTemplateId);
    handleTemplateLibraryUpdated();
    updateDeleteTemplateButtonState();
  }

  function openTemplatesModal() {
    populateTemplateSelector(templateState.lastSelectedTemplateId);

    if (templateState.lastSelectedTemplateId) {
      const selectedTemplate = trainingTemplates.getTemplateById(templateState.lastSelectedTemplateId);
      if (selectedTemplate) {
        templateState.currentTemplateId = selectedTemplate.id;
        loadTemplateForm(selectedTemplate);
        elements.templateSelector.value = selectedTemplate.id;
      } else {
        resetTemplateForm();
      }
    } else {
      resetTemplateForm();
    }

    if (templateState.modalInstance) {
      templateState.modalInstance.show();
    }
  }

  function populateTemplateSelector(selectedId = '') {
    if (!trainingTemplates || !elements.templateSelector) {
      return;
    }

    const templates = trainingTemplates.listTemplates();
    const selector = elements.templateSelector;
    selector.innerHTML = '';

    const placeholder = document.createElement('option');
    placeholder.value = '';
    placeholder.textContent = 'Selecciona una plantilla…';
    placeholder.disabled = true;
    if (!selectedId) {
      placeholder.selected = true;
    }
    selector.appendChild(placeholder);

    templates.forEach((template) => {
      const option = document.createElement('option');
      option.value = template.id;
      option.textContent = template.name;
      if (template.id === selectedId) {
        option.selected = true;
      }
      selector.appendChild(option);
    });

    const newOption = document.createElement('option');
    newOption.value = '__new__';
    newOption.textContent = 'Crear nueva plantilla';
    if (selectedId === '__new__') {
      newOption.selected = true;
    }
    selector.appendChild(newOption);

    if (selectedId && selectedId !== '__new__') {
      selector.value = selectedId;
    } else if (selectedId === '__new__') {
      selector.value = '__new__';
    } else {
      selector.selectedIndex = 0;
    }
  }

  function resetTemplateForm() {
    templateState.currentTemplateId = '';
    const emptyTemplate = trainingTemplates ? trainingTemplates.createEmptyTemplate() : null;
    loadTemplateForm(emptyTemplate);
    if (elements.templateSelector) {
      elements.templateSelector.selectedIndex = 0;
    }
  }

  function loadTemplateForm(template) {
    const {
      templateNameInput,
      templateTitleInput,
      templateDurationInput,
      theoryList,
      practiceList
    } = elements;

    const payload = template || (trainingTemplates ? trainingTemplates.createEmptyTemplate() : null) || {
      name: '',
      title: '',
      duration: '',
      theory: [],
      practice: []
    };

    templateNameInput.value = payload.name || '';
    templateTitleInput.value = payload.title || payload.name || '';
    templateDurationInput.value = payload.duration || '';

    renderTemplatePoints(theoryList, Array.isArray(payload.theory) ? payload.theory : []);
    renderTemplatePoints(practiceList, Array.isArray(payload.practice) ? payload.practice : []);

    updateDeleteTemplateButtonState();
  }

  function renderTemplatePoints(container, items) {
    if (!container) {
      return;
    }
    container.innerHTML = '';
    const entries = Array.isArray(items) ? items : [];
    entries.forEach((value) => {
      const point = createTemplatePointInput(value);
      container.appendChild(point);
    });
  }

  function addTemplatePoint(type) {
    const container = type === 'practice' ? elements.practiceList : elements.theoryList;
    if (!container) {
      return;
    }
    const point = createTemplatePointInput();
    container.appendChild(point);
    const input = point.querySelector('input');
    if (input) {
      input.focus();
    }
  }

  function createTemplatePointInput(value = '') {
    const wrapper = document.createElement('div');
    wrapper.className = 'input-group template-point';

    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'form-control';
    input.placeholder = 'Descripción del contenido';
    input.value = value;

    const removeButton = document.createElement('button');
    removeButton.type = 'button';
    removeButton.className = 'btn btn-outline-danger';
    removeButton.dataset.action = 'remove-point';
    removeButton.innerHTML = '<span aria-hidden="true">&times;</span>';
    removeButton.setAttribute('aria-label', 'Eliminar punto');

    wrapper.appendChild(input);
    wrapper.appendChild(removeButton);
    return wrapper;
  }

  function handleTemplatePointListClick(event) {
    const target = event.target.closest('[data-action="remove-point"]');
    if (!target) {
      return;
    }
    const wrapper = target.closest('.template-point');
    if (wrapper && wrapper.parentElement) {
      wrapper.parentElement.removeChild(wrapper);
    }
  }

  function collectTemplatePoints(type) {
    const container = type === 'practice' ? elements.practiceList : elements.theoryList;
    if (!container) {
      return [];
    }
    return Array.from(container.querySelectorAll('input'))
      .map((input) => input.value.trim())
      .filter((text) => text !== '');
  }

  function handleTemplateSelectionChange() {
    if (!trainingTemplates) {
      return;
    }
    const { templateSelector } = elements;
    const selectedValue = templateSelector.value;

    if (!selectedValue || selectedValue === '') {
      resetTemplateForm();
      return;
    }

    if (selectedValue === '__new__') {
      templateState.currentTemplateId = '';
      loadTemplateForm(trainingTemplates.createEmptyTemplate());
      return;
    }

    const selectedTemplate = trainingTemplates.getTemplateById(selectedValue);
    if (selectedTemplate) {
      templateState.currentTemplateId = selectedTemplate.id;
      templateState.lastSelectedTemplateId = selectedTemplate.id;
      loadTemplateForm(selectedTemplate);
    }
  }

  function updateDeleteTemplateButtonState() {
    const { deleteTemplateButton } = elements;
    if (!deleteTemplateButton) {
      return;
    }

    const currentId = templateState.currentTemplateId;
    const canDelete = Boolean(
      currentId &&
        trainingTemplates &&
        typeof trainingTemplates.isCustomTemplateId === 'function' &&
        trainingTemplates.isCustomTemplateId(currentId)
    );

    deleteTemplateButton.disabled = !canDelete;
    deleteTemplateButton.setAttribute('aria-disabled', String(!canDelete));
    if (canDelete) {
      deleteTemplateButton.removeAttribute('title');
    } else {
      deleteTemplateButton.setAttribute(
        'title',
        'Solo se pueden eliminar las plantillas personalizadas.'
      );
    }
  }

  function handleDeleteTemplateClick() {
    if (!trainingTemplates) {
      return;
    }

    const templateId = templateState.currentTemplateId;
    if (!templateId || !trainingTemplates.isCustomTemplateId(templateId)) {
      return;
    }

    const template = trainingTemplates.getTemplateById(templateId);
    const templateName = template && template.name ? template.name : '';
    const confirmationMessage = templateName
      ? `¿Seguro que quieres eliminar la plantilla "${templateName}"? Esta acción no se puede deshacer.`
      : '¿Seguro que quieres eliminar la plantilla seleccionada? Esta acción no se puede deshacer.';

    if (!window.confirm(confirmationMessage)) {
      return;
    }

    const deleted = trainingTemplates.deleteTemplate(templateId);
    if (!deleted) {
      showAlert('danger', 'No se ha podido eliminar la plantilla.');
      return;
    }

    if (templateState.lastSelectedTemplateId === templateId) {
      templateState.lastSelectedTemplateId = '';
    }

    templateState.currentTemplateId = '';
    populateTemplateSelector('');
    resetTemplateForm();
    showAlert('success', 'Plantilla eliminada correctamente.');
  }

  function handleTemplateFormSubmit(event) {
    event.preventDefault();
    if (!trainingTemplates) {
      showAlert('danger', 'No se ha podido guardar la plantilla.');
      return;
    }

    const nameValue = elements.templateNameInput.value.trim();
    const titleValue = elements.templateTitleInput.value.trim();
    const durationValue = elements.templateDurationInput.value.trim();

    if (!nameValue) {
      elements.templateNameInput.focus();
      showAlert('danger', 'Introduce el nombre de la formación.');
      return;
    }

    const templatePayload = {
      id: templateState.currentTemplateId,
      name: nameValue,
      title: titleValue || nameValue,
      duration: durationValue,
      theory: collectTemplatePoints('theory'),
      practice: collectTemplatePoints('practice')
    };

    try {
      const savedTemplate = trainingTemplates.saveTemplate(templatePayload);
      if (!savedTemplate) {
        throw new Error('La plantilla no se ha podido guardar.');
      }
      templateState.currentTemplateId = savedTemplate.id;
      templateState.lastSelectedTemplateId = savedTemplate.id;
      populateTemplateSelector(savedTemplate.id);
      loadTemplateForm(savedTemplate);
      elements.templateSelector.value = savedTemplate.id;
      showAlert('success', 'Plantilla guardada correctamente.');
    } catch (error) {
      console.error('No se ha podido guardar la plantilla', error);
      const message = error && error.message ? error.message : 'No se ha podido guardar la plantilla.';
      showAlert('danger', message);
    }
  }

  function handleTemplateLibraryUpdated() {
    if (!trainingTemplates || !state.rows.length) {
      return;
    }

    let hasChanges = false;
    const updatedRows = state.rows.map((row) => {
      const updatedDuration = trainingTemplates.getTrainingDuration(row.formacion);
      if (updatedDuration && row.duracion !== updatedDuration) {
        hasChanges = true;
        return { ...row, duracion: updatedDuration };
      }
      return row;
    });

    if (hasChanges) {
      state.rows = updatedRows;
      persistRows();
      renderTable();
    }
  }

  function hydrateFromStorage() {
    hydrateRowsFromStorage();
    hydrateDealContactsFromStorage();
    pruneStoredDealContacts();
  }

  function hydrateRowsFromStorage() {
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
          applyTrainingSpecificRulesToRowObject(hydratedRow);
          return hydratedRow;
        });
      }
    } catch (error) {
      console.error('No se ha podido recuperar la información guardada', error);
    }
  }

  function hydrateDealContactsFromStorage() {
    try {
      const raw = window.sessionStorage.getItem(CONTACTS_STORAGE_KEY);
      if (!raw) return;
      const parsed = JSON.parse(raw);
      if (!Array.isArray(parsed)) {
        return;
      }

      const hydratedContacts = new Map();

      parsed.forEach((entry) => {
        if (!Array.isArray(entry) || entry.length < 2) {
          return;
        }
        const [storedDealId, storedContact] = entry;
        const normalisedDealId = normaliseDealId(storedDealId);
        if (!normalisedDealId || typeof storedContact !== 'object' || storedContact === null) {
          return;
        }

        hydratedContacts.set(normalisedDealId, {
          dealId: String(storedContact.dealId || storedDealId || ''),
          contactName: storedContact.contactName || '',
          contactEmail: storedContact.contactEmail || '',
          contactPersonId: storedContact.contactPersonId || ''
        });
      });

      state.dealContacts = hydratedContacts;
    } catch (error) {
      console.error('No se ha podido recuperar la información de contacto guardada', error);
      state.dealContacts = new Map();
    }
  }

  function pruneStoredDealContacts() {
    if (!(state.dealContacts instanceof Map)) {
      state.dealContacts = new Map();
    }

    const validDealIds = new Set(
      state.rows
        .map((row) => normaliseDealId(row.presupuesto))
        .filter((dealId) => dealId)
    );

    const updatedContacts = new Map();

    state.dealContacts.forEach((contact, dealIdKey) => {
      if (validDealIds.has(dealIdKey)) {
        updatedContacts.set(dealIdKey, contact);
      }
    });

    state.rows.forEach((row) => {
      const normalisedDealId = normaliseDealId(row.presupuesto);
      if (!normalisedDealId || updatedContacts.has(normalisedDealId)) {
        return;
      }

      if (!row.personaContacto && !row.correoContacto) {
        return;
      }

      updatedContacts.set(normalisedDealId, {
        dealId: String(row.presupuesto || ''),
        contactName: row.personaContacto || '',
        contactEmail: row.correoContacto || '',
        contactPersonId: row.contactPersonId || ''
      });
    });

    const hasChanges =
      updatedContacts.size !== state.dealContacts.size ||
      Array.from(updatedContacts.entries()).some(([dealIdKey, contact]) => {
        const previous = state.dealContacts.get(dealIdKey);
        if (!previous) {
          return true;
        }

        return (
          (previous.contactName || '') !== (contact.contactName || '') ||
          (previous.contactEmail || '') !== (contact.contactEmail || '') ||
          (previous.contactPersonId || '') !== (contact.contactPersonId || '') ||
          (previous.dealId || '') !== (contact.dealId || '')
        );
      });

    state.dealContacts = updatedContacts;

    if (hasChanges) {
      if (state.dealContacts.size === 0) {
        clearStoredDealContacts();
      } else {
        persistDealContacts();
      }
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

  function persistDealContacts() {
    try {
      if (!(state.dealContacts instanceof Map)) {
        state.dealContacts = new Map();
      }

      if (state.dealContacts.size === 0) {
        window.sessionStorage.removeItem(CONTACTS_STORAGE_KEY);
        return;
      }

      const serialisableContacts = Array.from(state.dealContacts.entries()).map(([dealIdKey, contact]) => [
        dealIdKey,
        {
          dealId: contact && contact.dealId ? String(contact.dealId) : String(dealIdKey || ''),
          contactName: contact && contact.contactName ? contact.contactName : '',
          contactEmail: contact && contact.contactEmail ? contact.contactEmail : '',
          contactPersonId: contact && contact.contactPersonId ? contact.contactPersonId : ''
        }
      ]);

      window.sessionStorage.setItem(CONTACTS_STORAGE_KEY, JSON.stringify(serialisableContacts));
    } catch (error) {
      console.error('No se ha podido guardar la información de contacto', error);
      showAlert('warning', 'No se ha podido guardar la información de contacto en esta sesión.');
    }
  }

  function storeDealContact(dealId, contactData = {}) {
    const normalisedDealId = normaliseDealId(dealId);
    if (!normalisedDealId) {
      return;
    }

    if (!(state.dealContacts instanceof Map)) {
      state.dealContacts = new Map();
    }

    const contactRecord = {
      dealId: String(dealId || ''),
      contactName: contactData.contactName || '',
      contactEmail: contactData.contactEmail || '',
      contactPersonId: contactData.contactPersonId || ''
    };

    state.dealContacts.set(normalisedDealId, contactRecord);
    persistDealContacts();
  }

  function removeStoredDealContactIfUnused(dealId, { skipRowIndex = null } = {}) {
    const normalisedDealId = normaliseDealId(dealId);
    if (!normalisedDealId || !(state.dealContacts instanceof Map) || state.dealContacts.size === 0) {
      return;
    }

    const stillReferenced = state.rows.some((row, index) => {
      if (skipRowIndex !== null && index === skipRowIndex) {
        return false;
      }
      return normaliseDealId(row.presupuesto) === normalisedDealId;
    });

    if (!stillReferenced && state.dealContacts.has(normalisedDealId)) {
      state.dealContacts.delete(normalisedDealId);
      persistDealContacts();
    }
  }

  function clearStoredDealContacts() {
    state.dealContacts = new Map();
    try {
      window.sessionStorage.removeItem(CONTACTS_STORAGE_KEY);
    } catch (error) {
      console.error('No se ha podido limpiar la información de contacto', error);
    }
  }

  function handleDealContactBudgetChange(previousBudgetId, newBudgetId, rowIndex) {
    const previousDealId = normaliseDealId(previousBudgetId);
    const newDealId = normaliseDealId(newBudgetId);

    if (!previousDealId) {
      return;
    }

    if (previousDealId === newDealId) {
      return;
    }

    removeStoredDealContactIfUnused(previousDealId, { skipRowIndex: rowIndex });
  }

  function createEmptyRow() {
    return {
      presupuesto: '',
      nombre: '',
      apellido: '',
      dni: '',
      documentType: '',
      fecha: '',
      segundaFecha: '',
      lugar: '',
      duracion: '',
      cliente: '',
      formacion: '',
      irata: '',
      personaContacto: '',
      correoContacto: '',
      contactPersonId: ''
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
      updateActionButtonsState();
      return;
    }

    state.rows.forEach((row, index) => {
      const tr = document.createElement('tr');

      TABLE_COLUMNS.forEach((column) => {
        const td = document.createElement('td');
        if (column.field === 'duracion') {
          td.classList.add('column-hours');
        }
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
          const handleValueChange = (event) => {
            const { value } = event.target;
            updateRowValue(index, column.field, value);
            updateCellTooltip(td, input, column, value);
            if (column.field === 'formacion') {
              applyTrainingDuration(index, event.target.value);
            }
            if (column.field === 'formacion' || column.field === 'fecha') {
              applyTrainingSpecificLogic(index);
            }
          };
          if (column.type === 'select') {
            input.addEventListener('change', handleValueChange);
            input.addEventListener('input', handleValueChange);
          } else {
            input.addEventListener('input', handleValueChange);
            if (column.type === 'date') {
              input.addEventListener('change', handleValueChange);
            }
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
      pdfButton.textContent = CERTIFICATE_BUTTON_LABEL;
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

    updateActionButtonsState();
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
    if (column.type === 'select') {
      const select = document.createElement('select');
      select.className = 'form-select';
      select.dataset.field = column.field;
      select.dataset.index = String(index);
      if (column.label) {
        select.setAttribute('aria-label', column.label);
      }

      const placeholderOption = document.createElement('option');
      placeholderOption.value = '';
      placeholderOption.textContent = column.placeholder || 'Selecciona una opción';
      select.appendChild(placeholderOption);

      const options = Array.isArray(column.options) ? column.options : [];
      options.forEach((option) => {
        if (option === null || option === undefined) {
          return;
        }
        const optionElement = document.createElement('option');
        if (typeof option === 'object') {
          const optionValue = option.value === undefined || option.value === null ? '' : String(option.value);
          optionElement.value = optionValue;
          optionElement.textContent = option.label || optionValue;
        } else {
          const optionValue = String(option);
          optionElement.value = optionValue;
          optionElement.textContent = optionValue;
        }
        select.appendChild(optionElement);
      });

      const normalisedValue = value === undefined || value === null ? '' : String(value);
      select.value = normalisedValue;
      if (select.value !== normalisedValue) {
        select.value = '';
      }

      return select;
    }

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
    if (column.field === 'duracion') {
      input.autocomplete = 'off';
      input.inputMode = 'decimal';
      input.pattern = '^\\d+(?:[.,]\\d{1,2})?h?$';
      input.spellcheck = false;
    } else if (column.type === 'number') {
      input.min = '0';
      input.step = '0.5';
      input.inputMode = 'decimal';
    }

    return input;
  }

  function updateRowValue(index, field, value, reRender = false) {
    if (!state.rows[index]) {
      return;
    }

    const row = state.rows[index];
    const previousValue = row[field];
    row[field] = value;
    persistRows();

    if (field === 'presupuesto' && previousValue !== value) {
      handleDealContactBudgetChange(previousValue, value, index);
    }

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
          triggerButton.textContent = originalLabel || CERTIFICATE_BUTTON_LABEL;
        }
      });
  }

  async function handleGenerateAllCertificates() {
    if (!state.rows.length) {
      showAlert('info', 'Añade al menos un alumn@ antes de generar los certificados.');
      return;
    }

    if (!window.certificatePdf || typeof window.certificatePdf.generate !== 'function') {
      showAlert('danger', 'La librería de certificados no está disponible.');
      return;
    }

    const triggerButton = elements.generateAllCertificates;
    let originalLabel = '';

    if (triggerButton instanceof HTMLButtonElement) {
      originalLabel = triggerButton.textContent || '';
      triggerButton.disabled = true;
      triggerButton.textContent = 'Generando...';
    }

    const rowsSnapshot = [...state.rows];
    const failedRows = [];

    for (let index = 0; index < rowsSnapshot.length; index += 1) {
      const row = rowsSnapshot[index];
      try {
        await window.certificatePdf.generate(row);
      } catch (error) {
        console.error(`No se ha podido generar el certificado PDF de la fila ${index + 1}`, error);
        failedRows.push(index + 1);
      }
    }

    if (failedRows.length === 0) {
      showAlert('success', 'Todos los certificados se han generado y descargado correctamente.');
    } else if (failedRows.length === rowsSnapshot.length) {
      showAlert('danger', 'No se ha podido generar ningún certificado. Revisa los datos e inténtalo de nuevo.');
    } else {
      const failedSummary = failedRows.join(', ');
      const suffix = failedSummary ? ` (filas: ${failedSummary})` : '';
      showAlert('warning', `Algunos certificados no se han podido generar${suffix}. Inténtalo de nuevo.`);
    }

    if (triggerButton instanceof HTMLButtonElement) {
      triggerButton.disabled = state.rows.length === 0;
      triggerButton.textContent = originalLabel || GENERATE_ALL_CERTIFICATES_LABEL;
    }

    updateActionButtonsState();
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
      segundaFecha: normaliseDateValue(data.secondaryTrainingDate),
      lugar: data.trainingLocation || '',
      duracion: getTrainingDuration(data.trainingName),
      cliente: data.clientName || '',
      formacion: data.trainingName || '',
      personaContacto: data.contactName || '',
      correoContacto: data.contactEmail || '',
      contactPersonId: data.contactPersonId || ''
    };

    storeDealContact(dealId, {
      contactName: data.contactName || '',
      contactEmail: data.contactEmail || '',
      contactPersonId: data.contactPersonId || ''
    });

    const students = Array.isArray(data.students) ? data.students : [];

    if (!students.length) {
      const newRow = { ...baseRow };
      applyTrainingSpecificRulesToRowObject(newRow);
      state.rows.push(newRow);
    } else {
      students.forEach((student) => {
        const newRow = {
          ...baseRow,
          nombre: student.name || '',
          apellido: student.surname || '',
          dni: student.document || '',
          documentType: detectDocumentType(student.document || '')
        };
        applyTrainingSpecificRulesToRowObject(newRow);
        state.rows.push(newRow);
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
    if (index < 0 || index >= state.rows.length) {
      return;
    }

    const [removedRow] = state.rows.splice(index, 1);
    persistRows();
    if (removedRow && removedRow.presupuesto) {
      removeStoredDealContactIfUnused(removedRow.presupuesto);
    }
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
    clearStoredDealContacts();
    renderTable();
    showAlert('success', 'Listado vaciado correctamente.');
  }

  function updateActionButtonsState() {
    const hasRows = state.rows.length > 0;
    elements.clearStorage.disabled = !hasRows;
    elements.generateAllCertificates.disabled = !hasRows;
  }

  function getTrainingDuration(trainingName) {
    if (!trainingTemplates) {
      return '';
    }
    return trainingTemplates.getTrainingDuration(trainingName) || '';
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

  function applyTrainingSpecificRulesToRowObject(row) {
    if (!row) {
      return row;
    }

    const desiredSecondaryDate = computeTrabajosVerticalesSecondDate(row);
    if (desiredSecondaryDate) {
      row.segundaFecha = desiredSecondaryDate;
    }

    return row;
  }

  function applyTrainingSpecificLogic(rowIndex) {
    if (!applyTrabajosVerticalesSecondDate(rowIndex)) {
      return;
    }
  }

  function applyTrabajosVerticalesSecondDate(rowIndex) {
    const row = state.rows[rowIndex];
    if (!row) {
      return false;
    }

    const desiredSecondaryDate = computeTrabajosVerticalesSecondDate(row);
    if (!desiredSecondaryDate) {
      return false;
    }

    updateRowValue(rowIndex, 'segundaFecha', desiredSecondaryDate);
    syncTableInputValue(rowIndex, 'segundaFecha', desiredSecondaryDate);
    return true;
  }

  function computeTrabajosVerticalesSecondDate(row) {
    if (!row || !isTrabajosVerticalesTraining(row.formacion)) {
      return null;
    }

    const primaryDate = normaliseDateValue(row.fecha);
    if (!primaryDate) {
      return null;
    }

    const desiredSecondaryDate = addDaysToIsoDate(primaryDate, 1);
    if (!desiredSecondaryDate) {
      return null;
    }

    const currentSecondaryNormalised = normaliseDateValue(row.segundaFecha);
    if (currentSecondaryNormalised === desiredSecondaryDate && row.segundaFecha === desiredSecondaryDate) {
      return null;
    }

    return desiredSecondaryDate;
  }

  function isTrabajosVerticalesTraining(trainingValue) {
    const normalisedName = normaliseTrainingName(trainingValue);
    if (!normalisedName) {
      return false;
    }
    return normalisedName.startsWith(TRABAJOS_VERTICALES_KEY);
  }

  function normaliseTrainingName(value) {
    if (trainingTemplates && typeof trainingTemplates.normaliseName === 'function') {
      return trainingTemplates.normaliseName(value);
    }
    if (value === undefined || value === null) {
      return '';
    }
    return String(value)
      .toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/[^a-z0-9]+/g, ' ')
      .trim()
      .replace(/\s+/g, ' ');
  }

  function addDaysToIsoDate(isoDate, daysToAdd) {
    if (!isoDate) {
      return '';
    }

    const date = new Date(`${isoDate}T00:00:00Z`);
    if (Number.isNaN(date.getTime())) {
      return '';
    }

    const increment = Number(daysToAdd || 0);
    date.setUTCDate(date.getUTCDate() + increment);
    return date.toISOString().split('T')[0];
  }

  function syncTableInputValue(rowIndex, field, value) {
    if (!elements.tableBody) {
      return;
    }

    const selector = `input[data-index="${rowIndex}"][data-field="${field}"]`;
    const input = elements.tableBody.querySelector(selector);
    if (!input) {
      return;
    }

    if (input.value !== value) {
      input.value = value;
    }

    const column = TABLE_COLUMN_LOOKUP.get(field);
    const cell = input.closest('td');
    if (cell && column) {
      updateCellTooltip(cell, input, column, value);
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
