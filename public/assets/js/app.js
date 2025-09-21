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
  const ACCOUNTING_EMAIL = 'contabilidad@gepgroup.es';
  const EMAIL_SEPARATOR = ';';
  const EMAIL_LIST_DELIMITER_PATTERN = /[;,]+/;
  const EMAIL_VALIDATION_PATTERN = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
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

  const CERTIFICATE_BUTTON_LABEL = 'PDF';
  const GENERATE_ALL_CERTIFICATES_LABEL = "Todos PDF's";
  const SVG_NS = 'http://www.w3.org/2000/svg';

  const elements = {
    budgetForm: document.getElementById('budget-form'),
    budgetInput: document.getElementById('budget-input'),
    fillButton: document.getElementById('fill-button'),
    addManualRow: document.getElementById('add-manual-row'),
    clearStorage: document.getElementById('clear-storage'),
    generateAllCertificates: document.getElementById('generate-all-certificates'),
    sendAllToDrive: document.getElementById('send-all-to-drive'),
    sendAllByEmail: document.getElementById('send-all-by-email'),
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
    practiceList: document.getElementById('practice-list'),
    emailModal: document.getElementById('email-modal'),
    emailForm: document.getElementById('email-form'),
    emailToInput: document.getElementById('email-to'),
    emailCcInput: document.getElementById('email-cc'),
    emailBccInput: document.getElementById('email-bcc'),
    emailBodyInput: document.getElementById('email-body'),
    emailSubjectInput: document.getElementById('email-subject'),
    emailSubjectPreview: document.getElementById('email-subject-preview'),
    emailStatus: document.getElementById('email-status'),
    emailBulkCounter: document.getElementById('email-bulk-counter'),
    emailSendButton: document.getElementById('email-send-button')
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

  const emailModalState = {
    modalInstance: null,
    currentRowIndex: -1,
    isSending: false,
    sendButtonOriginalLabel: '',
    autoCloseTimeoutId: null,
    initialised: false,
    isPreparingLink: false,
    isBulkMode: false,
    bulkQueue: [],
    bulkCurrentIndex: -1,
    bulkResults: [],
    bulkCurrentData: null
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
    if (elements.sendAllToDrive) {
      elements.sendAllToDrive.addEventListener('click', handleSendAllToDrive);
    }
    if (elements.sendAllByEmail) {
      elements.sendAllByEmail.addEventListener('click', handleSendAllByEmail);
    }
    setupTemplateManagement();
    setupEmailModal();
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
      contactPersonId: '',
      driveFileId: '',
      driveFileName: '',
      driveFileUrl: '',
      driveFileDownloadUrl: '',
      driveClientFolderId: '',
      driveTrainingFolderId: '',
      driveUploadedAt: ''
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
      actionsWrapper.className = 'd-flex flex-column flex-sm-row flex-sm-wrap gap-2 justify-content-center';

      const pdfButton = document.createElement('button');
      pdfButton.type = 'button';
      pdfButton.className = 'btn btn-primary btn-sm';
      pdfButton.textContent = CERTIFICATE_BUTTON_LABEL;
      pdfButton.addEventListener('click', () => handleGeneratePdf(index, pdfButton));

      const driveButton = document.createElement('button');
      driveButton.type = 'button';
      driveButton.className = 'btn btn-outline-success btn-sm btn-icon';
      driveButton.setAttribute('aria-label', 'Enviar a Google Drive');
      driveButton.title = 'Enviar a Google Drive';
      driveButton.appendChild(createGoogleDriveIcon());
      driveButton.addEventListener('click', () => handleRowGoogleDrive(index, driveButton));

      const emailButton = document.createElement('button');
      emailButton.type = 'button';
      emailButton.className = 'btn btn-outline-secondary btn-sm btn-icon';
      emailButton.setAttribute('aria-label', 'Enviar por correo');
      emailButton.title = 'Enviar por correo';
      emailButton.appendChild(createEnvelopeIcon());
      emailButton.addEventListener('click', () => handleRowEmail(index));

      const deleteButton = document.createElement('button');
      deleteButton.type = 'button';
      deleteButton.className = 'btn btn-outline-danger btn-sm btn-icon';
      deleteButton.setAttribute('aria-label', 'Eliminar fila');
      deleteButton.title = 'Eliminar fila';
      deleteButton.appendChild(createTrashIcon());
      deleteButton.addEventListener('click', () => removeRow(index));

      actionsWrapper.appendChild(pdfButton);
      actionsWrapper.appendChild(driveButton);
      actionsWrapper.appendChild(emailButton);
      actionsWrapper.appendChild(deleteButton);
      actionsTd.appendChild(actionsWrapper);
      tr.appendChild(actionsTd);

      tableBody.appendChild(tr);
    });

    updateActionButtonsState();
  }

  function createSvgElement(tagName) {
    return document.createElementNS(SVG_NS, tagName);
  }

  function createGoogleDriveIcon() {
    const svg = createSvgElement('svg');
    svg.setAttribute('viewBox', '0 0 64 64');
    svg.classList.add('icon', 'icon-google-drive');
    svg.setAttribute('aria-hidden', 'true');
    svg.setAttribute('focusable', 'false');

    const top = createSvgElement('polygon');
    top.setAttribute('fill', '#FBBC04');
    top.setAttribute('points', '32 6 44 26 20 26');
    svg.appendChild(top);

    const left = createSvgElement('polygon');
    left.setAttribute('fill', '#0F9D58');
    left.setAttribute('points', '6 54 20 26 32 46 18 54');
    svg.appendChild(left);

    const right = createSvgElement('polygon');
    right.setAttribute('fill', '#4285F4');
    right.setAttribute('points', '58 54 44 26 32 46 46 54');
    svg.appendChild(right);

    return svg;
  }

  function createEnvelopeIcon() {
    const svg = createSvgElement('svg');
    svg.setAttribute('viewBox', '0 0 24 24');
    svg.classList.add('icon', 'icon-envelope');
    svg.setAttribute('aria-hidden', 'true');
    svg.setAttribute('focusable', 'false');

    const path = createSvgElement('path');
    path.setAttribute('fill', 'currentColor');
    path.setAttribute(
      'd',
      'M4 4h16a2 2 0 0 1 2 2v12a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2zm0 2v.4l8 5.2 8-5.2V6H4zm16 2.8-6.92 4.49a3 3 0 0 1-3.16 0L4 8.8V18h16V8.8z'
    );
    svg.appendChild(path);

    return svg;
  }

  function createTrashIcon() {
    const svg = createSvgElement('svg');
    svg.setAttribute('viewBox', '0 0 24 24');
    svg.classList.add('icon', 'icon-trash');
    svg.setAttribute('aria-hidden', 'true');
    svg.setAttribute('focusable', 'false');

    const lid = createSvgElement('path');
    lid.setAttribute('fill', 'currentColor');
    lid.setAttribute('d', 'M9 4V3a2 2 0 0 1 2-2h2a2 2 0 0 1 2 2v1h5v2H4V4h5zm2-1v1h2V3h-2z');
    svg.appendChild(lid);

    const bin = createSvgElement('path');
    bin.setAttribute('fill', 'currentColor');
    bin.setAttribute('d', 'M5 7h14l-1.05 13.143A2 2 0 0 1 15.97 22H8.03a2 2 0 0 1-1.98-1.857L5 7zm5 3v9h2v-9h-2zm4 0v9h2v-9h-2z');
    svg.appendChild(bin);

    return svg;
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

  function showUpcomingFeatureMessage(label) {
    showAlert('info', `${label} estará disponible próximamente.`);
  }

  function resolveCertificateModule() {
    const certificate = window.certificatePdf;
    if (!certificate || typeof certificate.generate !== 'function') {
      return {
        certificate: null,
        error: {
          type: 'danger',
          message: 'La librería de certificados no está disponible.'
        }
      };
    }

    return { certificate };
  }

  function resolveGoogleDriveIntegration() {
    const drive = window.googleDrive;
    if (!drive || typeof drive.uploadCertificate !== 'function') {
      return {
        drive: null,
        error: {
          type: 'danger',
          message: 'La integración con Google Drive no está disponible en esta página.'
        }
      };
    }

    if (typeof drive.isConfigured === 'function' && !drive.isConfigured()) {
      return {
        drive: null,
        error: {
          type: 'warning',
          message: 'Configura el identificador de cliente de Google antes de enviar certificados a Drive.'
        }
      };
    }

    return { drive };
  }

  function buildDriveUploadDetails(row, certificate) {
    if (!row) {
      return {
        error: {
          type: 'danger',
          message: 'No se ha encontrado la fila seleccionada.'
        }
      };
    }

    const clientName = sanitiseDriveComponent(row.cliente, '');
    if (!clientName) {
      return {
        error: {
          type: 'warning',
          message: 'Introduce el nombre del cliente antes de enviar el certificado a Drive.'
        }
      };
    }

    const trainingTitle = getResolvedTrainingTitle(row);
    if (!trainingTitle) {
      return {
        error: {
          type: 'warning',
          message: 'Introduce el título de la formación antes de enviar el certificado a Drive.'
        }
      };
    }

    const primaryDateIso = normaliseDateValue(row.fecha);
    if (!primaryDateIso) {
      return {
        error: {
          type: 'warning',
          message: 'Indica la fecha principal de la formación antes de enviar el certificado a Drive.'
        }
      };
    }

    const sanitisedTrainingTitle = sanitiseDriveComponent(trainingTitle, 'Formación sin título');
    const trainingFolderDate = formatDateForFolderName(primaryDateIso);
    const trainingFolderName = sanitiseDriveComponent(
      `${sanitisedTrainingTitle} - ${trainingFolderDate}`,
      'Formación sin título'
    );

    const clientFolderName = sanitiseDriveComponent(clientName, 'Cliente sin nombre');

    const fallbackStudentName = sanitiseDriveComponent(buildStudentFullName(row), 'Alumno/a');
    const fallbackDateLabel = formatDateForFileLabel(row.fecha);
    const fallbackFileName = `${sanitisedTrainingTitle} - ${fallbackStudentName} - ${fallbackDateLabel}.pdf`;

    const fileName =
      certificate && typeof certificate.buildFileName === 'function'
        ? certificate.buildFileName(row)
        : fallbackFileName;

    return {
      clientFolderName,
      trainingFolderName,
      fileName: fileName || fallbackFileName
    };
  }

  async function handleRowGoogleDrive(rowIndex, triggerButton) {
    const row = state.rows[rowIndex];
    if (!row) {
      showAlert('danger', 'No se ha encontrado la fila seleccionada.');
      return;
    }

    const { certificate, error: certificateError } = resolveCertificateModule();
    if (!certificate) {
      if (certificateError) {
        showAlert(certificateError.type || 'danger', certificateError.message);
      }
      return;
    }

    const { drive, error: driveError } = resolveGoogleDriveIntegration();
    if (!drive) {
      if (driveError) {
        showAlert(driveError.type || 'danger', driveError.message);
      }
      return;
    }

    let originalHtml = '';
    if (triggerButton instanceof HTMLButtonElement) {
      originalHtml = triggerButton.innerHTML;
      triggerButton.disabled = true;
      triggerButton.setAttribute('aria-busy', 'true');
      triggerButton.innerHTML =
        '<span class="spinner-border spinner-border-sm" role="status" aria-hidden="true"></span>';
    }

    try {
      const result = await uploadRowCertificateToDrive(rowIndex, { drive, certificate });
      if (result && result.warning && result.warning.message) {
        showAlert(result.warning.type || 'warning', result.warning.message);
      }
      showAlert('success', 'Certificado guardado correctamente en Google Drive.');
    } catch (error) {
      console.error('No se ha podido guardar el certificado en Google Drive', error);
      const userFacing = error && error.userFacing ? error.userFacing : null;
      if (userFacing && userFacing.message) {
        showAlert(userFacing.type || 'danger', userFacing.message);
      } else {
        showAlert('danger', translateGoogleDriveError(error));
      }
    } finally {
      if (triggerButton instanceof HTMLButtonElement) {
        triggerButton.disabled = false;
        triggerButton.innerHTML = originalHtml || '';
        triggerButton.removeAttribute('aria-busy');
      }
    }
  }

  function handleRowEmail(rowIndex) {
    const row = state.rows[rowIndex];
    if (!row) {
      showAlert('danger', 'No se ha encontrado la fila seleccionada.');
      return;
    }

    if (!emailModalState.initialised || !emailModalState.modalInstance) {
      showAlert('warning', 'El envío por correo no está disponible en este momento.');
      return;
    }

    if (emailModalState.autoCloseTimeoutId) {
      window.clearTimeout(emailModalState.autoCloseTimeoutId);
      emailModalState.autoCloseTimeoutId = null;
    }

    emailModalState.currentRowIndex = rowIndex;
    setEmailModalLoading(false);
    setEmailModalStatus('');
    populateEmailModalFields(row);
    emailModalState.modalInstance.show();
    prepareEmailModalCertificateLink(rowIndex);
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

  async function handleSendAllToDrive() {
    if (!state.rows.length) {
      showAlert('info', 'Añade al menos un alumn@ antes de usar esta acción.');
      return;
    }

    const { certificate, error: certificateError } = resolveCertificateModule();
    if (!certificate) {
      if (certificateError) {
        showAlert(certificateError.type || 'danger', certificateError.message);
      }
      return;
    }

    const { drive, error: driveError } = resolveGoogleDriveIntegration();
    if (!drive) {
      if (driveError) {
        showAlert(driveError.type || 'danger', driveError.message);
      }
      return;
    }

    const triggerButton = elements.sendAllToDrive;
    let originalHtml = '';

    if (triggerButton instanceof HTMLButtonElement) {
      originalHtml = triggerButton.innerHTML;
      triggerButton.disabled = true;
      triggerButton.setAttribute('aria-busy', 'true');
      triggerButton.innerHTML =
        '<span class="spinner-border spinner-border-sm" role="status" aria-hidden="true"></span>';
    }

    const rowIndexes = state.rows.map((_, index) => index);
    const failedRows = [];
    const warningRows = [];

    for (const index of rowIndexes) {
      try {
        const result = await uploadRowCertificateToDrive(index, { drive, certificate });
        if (result && result.warning && result.warning.message) {
          warningRows.push({ index, message: result.warning.message, type: result.warning.type || 'warning' });
        }
      } catch (error) {
        console.error(
          `No se ha podido guardar el certificado en Google Drive de la fila ${index + 1}`,
          error
        );
        const userFacing = error && error.userFacing ? error.userFacing : null;
        const message =
          (userFacing && userFacing.message) || translateGoogleDriveError(error) || 'Error desconocido.';
        failedRows.push({ index, message });
      }
    }

    const totalRows = rowIndexes.length;
    const failedCount = failedRows.length;

    if (failedCount === 0) {
      showAlert('success', 'Todos los certificados se han guardado correctamente en Google Drive.');
    } else {
      const detailMessages = failedRows
        .map((item) => {
          const message = normaliseTextValue(item.message) || 'Error desconocido.';
          return `Fila ${item.index + 1}: ${message}`;
        })
        .join(' ');

      if (failedCount === totalRows) {
        const summary = detailMessages
          ? `No se ha podido guardar ningún certificado en Google Drive. ${detailMessages}`
          : 'No se ha podido guardar ningún certificado en Google Drive. Revisa los datos e inténtalo de nuevo.';
        showAlert('danger', summary);
      } else {
        const summary = detailMessages
          ? `Algunos certificados no se han podido guardar en Google Drive. ${detailMessages}`
          : 'Algunos certificados no se han podido guardar en Google Drive. Inténtalo de nuevo.';
        showAlert('warning', summary);
      }
    }

    if (warningRows.length) {
      const warningSummary = warningRows
        .map((item) => {
          const message = normaliseTextValue(item.message) || 'Revisa los permisos del archivo en Google Drive.';
          return `Fila ${item.index + 1}: ${message}`;
        })
        .join(' ');
      showAlert('warning', warningSummary);
    }

    if (triggerButton instanceof HTMLButtonElement) {
      triggerButton.disabled = false;
      triggerButton.innerHTML = originalHtml || '';
      triggerButton.removeAttribute('aria-busy');
    }

    updateActionButtonsState();
  }

  async function handleSendAllByEmail() {
    if (!state.rows.length) {
      showAlert('info', 'Añade al menos un alumn@ antes de usar esta acción.');
      return;
    }

    if (!emailModalState.initialised || !emailModalState.modalInstance) {
      showAlert('warning', 'El envío por correo no está disponible en este momento.');
      return;
    }

    if (emailModalState.isSending || emailModalState.isPreparingLink || emailModalState.isBulkMode) {
      showAlert('info', 'Ya hay un envío de correos en curso. Espera a que finalice.');
      return;
    }

    const queue = buildBulkEmailQueue();
    if (!queue.length) {
      showAlert('info', 'No hay presupuestos con información suficiente para enviar correos.');
      return;
    }

    const gmailResolution = resolveGoogleGmailIntegration();
    if (!gmailResolution.gmail) {
      const gmailError = gmailResolution.error;
      if (gmailError) {
        showAlert(gmailError.type || 'danger', gmailError.message);
      } else {
        showAlert('danger', 'La integración con Gmail no está disponible.');
      }
      return;
    }

    const triggerButton = elements.sendAllByEmail;
    let originalHtml = '';

    if (triggerButton instanceof HTMLButtonElement) {
      originalHtml = triggerButton.innerHTML;
      triggerButton.disabled = true;
      triggerButton.setAttribute('aria-busy', 'true');
      triggerButton.innerHTML =
        '<span class="spinner-border spinner-border-sm" role="status" aria-hidden="true"></span>';
    }

    try {
      await startBulkEmailManualFlow(queue);
    } catch (error) {
      console.error('No se ha podido preparar el envío masivo de correos.', error);
      showAlert('danger', 'No se ha podido preparar el envío masivo de correos. Inténtalo de nuevo.');
    } finally {
      if (triggerButton instanceof HTMLButtonElement) {
        triggerButton.disabled = state.rows.length === 0;
        triggerButton.innerHTML = originalHtml || '';
        triggerButton.removeAttribute('aria-busy');
      }
      updateActionButtonsState();
    }
  }

  async function startBulkEmailManualFlow(queue) {
    emailModalState.isBulkMode = true;
    emailModalState.bulkQueue = queue;
    emailModalState.bulkResults = [];
    emailModalState.bulkCurrentIndex = -1;
    emailModalState.bulkCurrentData = null;
    emailModalState.currentRowIndex = -1;
    emailModalState.isPreparingLink = false;

    setEmailModalStatus('');
    setEmailBulkCounter(0, queue.length);
    setEmailModalLoading(false);
    populateEmailModalFieldsForBulk({
      toValue: '',
      ccValue: normaliseEmailInput(ACCOUNTING_EMAIL),
      bccValue: '',
      subject: '',
      body: ''
    });
    updateEmailSendButtonState();

    emailModalState.modalInstance.show();
    await advanceToNextBulkEmail(0);
  }

  async function advanceToNextBulkEmail(startIndex) {
    if (!emailModalState.isBulkMode) {
      return;
    }

    const queue = Array.isArray(emailModalState.bulkQueue) ? emailModalState.bulkQueue : [];
    const total = queue.length;

    emailModalState.bulkCurrentData = null;
    emailModalState.isPreparingLink = true;
    updateEmailSendButtonState();

    for (let index = startIndex; index < total; index += 1) {
      if (!emailModalState.isBulkMode) {
        return;
      }

      const queueItem = queue[index];
      const counterCurrent = index + 1;
      setEmailBulkCounter(counterCurrent, total);

      const preparingMessage =
        total > 1 ? `(${counterCurrent}/${total}) Preparando correo…` : 'Preparando correo…';
      setEmailModalStatus(preparingMessage, 'info');

      const preparation = await prepareBulkQueueItem(queueItem, {
        index,
        total,
        onStatus: (message, type = 'info') => {
          if (!emailModalState.isBulkMode) {
            return;
          }
          const prefixed = total > 1 && message ? `(${counterCurrent}/${total}) ${message}` : message;
          setEmailModalStatus(prefixed || '', type);
        }
      });

      if (!emailModalState.isBulkMode) {
        return;
      }

      if (!preparation.success) {
        const failureMessage = preparation.statusMessage || 'No se ha podido preparar el correo.';
        const prefixedFailure =
          total > 1 ? `(${counterCurrent}/${total}) ${failureMessage}` : failureMessage;
        setEmailModalStatus(prefixedFailure, preparation.statusType || 'danger');
        emailModalState.bulkResults.push({ ...preparation, queueItem });
        const summaryMessage = buildBulkSummaryMessage(queueItem, failureMessage);
        if (summaryMessage) {
          showAlert(preparation.statusType || 'danger', summaryMessage);
        }
        await delay(600);
        continue;
      }

      const { data, warnings } = preparation;
      emailModalState.bulkCurrentIndex = index;
      emailModalState.bulkCurrentData = {
        queueItem,
        toValue: data.toValue,
        ccValue: data.ccValue,
        bccValue: data.bccValue,
        subject: data.subject,
        body: data.body,
        warnings,
        studentEntries: data.studentEntries,
        rowIndexes: data.rowIndexes,
        contact: data.contact
      };

      populateEmailModalFieldsForBulk({
        toValue: data.toValue,
        ccValue: data.ccValue,
        bccValue: data.bccValue,
        subject: data.subject,
        body: data.body
      });

      const readyMessage =
        preparation.statusMessage ||
        (warnings.length
          ? 'Correo preparado con avisos. Revisa y envía manualmente.'
          : 'Correo preparado. Revisa y envía manualmente.');
      const prefixedReady = total > 1 ? `(${counterCurrent}/${total}) ${readyMessage}` : readyMessage;

      emailModalState.isPreparingLink = false;
      setEmailModalStatus(
        prefixedReady,
        preparation.statusType || (warnings.length ? 'warning' : 'success')
      );
      updateEmailSendButtonState();
      return;
    }

    emailModalState.isPreparingLink = false;
    updateEmailSendButtonState();
    await finishBulkEmailFlow({ cancelled: false });
  }

  function buildBulkFailureResult(queueItem, message, type = 'danger', warnings = []) {
    return {
      success: false,
      statusMessage: message,
      statusType: type,
      summaryMessage: buildBulkSummaryMessage(queueItem, message),
      warnings
    };
  }

  async function prepareBulkQueueItem(queueItem, { index, total, onStatus } = {}) {
    if (!queueItem) {
      return buildBulkFailureResult(null, 'No se ha encontrado la información del correo.', 'danger');
    }

    const rowIndexes = Array.isArray(queueItem.rowIndexes) ? queueItem.rowIndexes : [];
    const rows = rowIndexes
      .map((rowIndex) => ({ rowIndex, row: state.rows[rowIndex] }))
      .filter((entry) => entry.row);

    if (!rows.length) {
      return buildBulkFailureResult(
        queueItem,
        'Los datos del presupuesto ya no están disponibles.',
        'danger'
      );
    }

    const contact = resolveGroupContact(queueItem);
    const toValue = normaliseEmailInput(contact.contactEmail || '');

    if (!toValue) {
      return buildBulkFailureResult(queueItem, 'No hay correo de contacto definido.', 'warning');
    }

    if (!hasValidEmailAddresses(toValue)) {
      return buildBulkFailureResult(queueItem, 'El correo de contacto tiene un formato no válido.', 'warning');
    }

    const ccValue = normaliseEmailInput(ACCOUNTING_EMAIL);
    const bccValue = '';
    const subject = buildEmailSubjectForRows(
      rows.map((entry) => entry.row),
      queueItem
    );

    const warningDetails = [];
    const studentEntries = [];

    for (let studentIndex = 0; studentIndex < rows.length; studentIndex += 1) {
      const { rowIndex, row } = rows[studentIndex];
      const studentName = buildStudentFullName(row) || `Alumno/a ${studentIndex + 1}`;

      if (typeof onStatus === 'function') {
        onStatus(`Preparando certificado · ${studentName}`, 'info');
      }

      const ensureResult = await ensureRowHasDriveFile(rowIndex, {
        onStatusChange: (message, type) => {
          if (typeof onStatus !== 'function') {
            return;
          }
          if (!message) {
            onStatus('', type);
            return;
          }
          const contextualMessage = rows.length > 1 ? `${message} · ${studentName}` : message;
          onStatus(contextualMessage, type);
        }
      });

      if (ensureResult.error) {
        const errorMessage = ensureResult.error.message || 'No se ha podido preparar el certificado.';
        return buildBulkFailureResult(
          queueItem,
          errorMessage,
          ensureResult.error.type || 'danger',
          warningDetails
        );
      }

      if (ensureResult.warning && ensureResult.warning.message) {
        warningDetails.push({
          message: ensureResult.warning.message,
          type: ensureResult.warning.type || 'warning'
        });
      }

      const link = ensureResult.link;
      if (!link) {
        return buildBulkFailureResult(
          queueItem,
          'No se ha podido obtener un enlace al certificado.',
          'danger',
          warningDetails
        );
      }

      studentEntries.push({ rowIndex, row, link });
    }

    if (!studentEntries.length) {
      return buildBulkFailureResult(
        queueItem,
        'No hay certificados disponibles para enviar.',
        'danger',
        warningDetails
      );
    }

    let body = '';

    if (studentEntries.length === 1) {
      const singleRow = studentEntries[0].row;
      const singleLink = studentEntries[0].link;
      let emailBody = buildEmailBody(singleRow);
      const ensuredBody = ensureEmailBodyHasLink(emailBody, singleLink);
      if (ensuredBody.didUpdate) {
        emailBody = ensuredBody.updatedBody;
      }
      body = emailBody;
    } else {
      body = buildMultiStudentEmailBody(studentEntries);
    }

    return {
      success: true,
      statusMessage:
        studentEntries.length === 1
          ? 'Certificado preparado. Revisa y envía manualmente.'
          : 'Certificados preparados. Revisa y envía manualmente.',
      statusType: warningDetails.length ? 'warning' : 'success',
      summaryMessage: null,
      warnings: warningDetails,
      data: {
        toValue,
        ccValue,
        bccValue,
        subject,
        body,
        rowIndexes: studentEntries.map((entry) => entry.rowIndex),
        studentEntries,
        contact
      }
    };
  }

  async function finishBulkEmailFlow({ cancelled } = {}) {
    const queue = Array.isArray(emailModalState.bulkQueue) ? emailModalState.bulkQueue : [];
    const results = Array.isArray(emailModalState.bulkResults) ? emailModalState.bulkResults : [];

    if (queue.length === 0) {
      setEmailBulkCounter(0, 0);
    } else {
      setEmailBulkCounter(queue.length, queue.length);
    }

    if (cancelled) {
      setEmailModalStatus('Proceso interrumpido.', 'warning');
    } else if (!results.length) {
      setEmailModalStatus('No hay correos para enviar.', 'warning');
    } else {
      const hasFailures = results.some((item) => item && item.success === false);
      const hasWarnings = results.some(
        (item) => item && Array.isArray(item.warnings) && item.warnings.length > 0
      );

      if (!hasFailures && !hasWarnings) {
        setEmailModalStatus('Proceso completado.', 'success');
      } else if (!hasFailures && hasWarnings) {
        setEmailModalStatus('Proceso completado con avisos.', 'warning');
      } else {
        setEmailModalStatus('Proceso completado con incidencias.', 'warning');
      }
    }

    emailModalState.isPreparingLink = false;
    emailModalState.isBulkMode = false;
    emailModalState.bulkCurrentIndex = -1;
    emailModalState.bulkCurrentData = null;
    emailModalState.bulkResults = results;
    updateEmailSendButtonState();

    const summary = summariseBulkEmailResults(results, queue, { cancelled });
    if (summary && summary.autoClose && emailModalState.modalInstance) {
      emailModalState.autoCloseTimeoutId = window.setTimeout(() => {
        if (emailModalState.modalInstance) {
          emailModalState.modalInstance.hide();
        }
      }, 1600);
    }
  }

  async function handleBulkEmailFormSubmission() {
    if (!emailModalState.isBulkMode) {
      return;
    }

    const currentData = emailModalState.bulkCurrentData;
    if (!currentData) {
      setEmailModalStatus('No hay ningún correo pendiente para enviar.', 'warning');
      return;
    }

    const toValue = normaliseEmailInput(elements.emailToInput ? elements.emailToInput.value : '');
    const ccValue = normaliseEmailInput(elements.emailCcInput ? elements.emailCcInput.value : '');
    const bccValue = normaliseEmailInput(elements.emailBccInput ? elements.emailBccInput.value : '');

    if (elements.emailToInput) {
      elements.emailToInput.value = toValue;
    }
    if (elements.emailCcInput) {
      elements.emailCcInput.value = ccValue;
    }
    if (elements.emailBccInput) {
      elements.emailBccInput.value = bccValue;
    }

    if (!toValue) {
      setEmailModalStatus('Introduce al menos un destinatario en el campo "Para".', 'warning');
      if (elements.emailToInput) {
        elements.emailToInput.focus();
      }
      return;
    }

    if (!hasValidEmailAddresses(toValue)) {
      setEmailModalStatus('Revisa las direcciones de correo del campo "Para".', 'warning');
      if (elements.emailToInput) {
        elements.emailToInput.focus();
      }
      return;
    }

    if (!hasValidEmailAddresses(ccValue)) {
      setEmailModalStatus('Revisa las direcciones de correo del campo "CC".', 'warning');
      if (elements.emailCcInput) {
        elements.emailCcInput.focus();
      }
      return;
    }

    if (!hasValidEmailAddresses(bccValue)) {
      setEmailModalStatus('Revisa las direcciones de correo del campo "CCO".', 'warning');
      if (elements.emailBccInput) {
        elements.emailBccInput.focus();
      }
      return;
    }

    const subjectValue =
      elements.emailSubjectInput && elements.emailSubjectInput.value
        ? elements.emailSubjectInput.value.trim()
        : currentData.subject || '';

    let bodyValue = elements.emailBodyInput ? elements.emailBodyInput.value : currentData.body || '';

    if (currentData.studentEntries && currentData.studentEntries.length === 1) {
      const singleLink = currentData.studentEntries[0]?.link || '';
      const ensured = ensureEmailBodyHasLink(bodyValue, singleLink);
      if (ensured.didUpdate) {
        bodyValue = ensured.updatedBody;
        if (elements.emailBodyInput) {
          elements.emailBodyInput.value = bodyValue;
        }
      }
    }

    setEmailModalLoading(true);
    setEmailModalStatus('Enviando correo…', 'info');

    const { gmail, error: gmailError } = resolveGoogleGmailIntegration();
    if (!gmail) {
      if (gmailError) {
        setEmailModalStatus(gmailError.message, gmailError.type || 'danger');
      } else {
        setEmailModalStatus('La integración con Gmail no está disponible.', 'danger');
      }
      setEmailModalLoading(false);
      return;
    }

    try {
      await gmail.sendEmail({
        to: formatEmailRecipientsForSending(toValue),
        cc: formatEmailRecipientsForSending(ccValue),
        bcc: formatEmailRecipientsForSending(bccValue),
        subject: subjectValue,
        body: bodyValue
      });
    } catch (error) {
      console.error('No se ha podido enviar el correo masivo.', error);
      setEmailModalStatus(translateGmailError(error), 'danger');
      setEmailModalLoading(false);
      return;
    }

    const resultStatusType =
      currentData.warnings && currentData.warnings.length ? 'warning' : 'success';

    const result = {
      success: true,
      statusMessage: 'Correo enviado correctamente.',
      statusType: resultStatusType,
      summaryMessage: null,
      warnings: currentData.warnings || []
    };

    const queueItem = currentData.queueItem;

    emailModalState.bulkResults.push({ ...result, queueItem });

    const queueRowIndexes = Array.isArray(queueItem.rowIndexes) ? queueItem.rowIndexes : [];
    queueRowIndexes.forEach((rowIndex) => {
      if (state.rows[rowIndex]) {
        updateRowValue(rowIndex, 'correoContacto', toValue);
      }
    });

    if (queueItem.dealId) {
      const contactToStore = {
        contactName: currentData.contact?.contactName || '',
        contactEmail: toValue,
        contactPersonId: currentData.contact?.contactPersonId || ''
      };
      storeDealContact(queueItem.dealId, contactToStore);
    }

    setEmailModalStatus('Correo enviado correctamente.', resultStatusType);
    setEmailModalLoading(false);

    const nextIndex = emailModalState.bulkCurrentIndex + 1;
    emailModalState.bulkCurrentData = null;
    emailModalState.bulkCurrentIndex = -1;

    await delay(300);
    await advanceToNextBulkEmail(nextIndex);
  }

  function buildBulkEmailQueue() {
    if (!Array.isArray(state.rows) || state.rows.length === 0) {
      return [];
    }

    const queue = [];
    const groupedByDealId = new Map();

    state.rows.forEach((row, index) => {
      if (!row) {
        return;
      }

      const normalisedDealId = normaliseDealId(row.presupuesto);
      if (normalisedDealId) {
        if (!groupedByDealId.has(normalisedDealId)) {
          const labelDealId = normaliseTextValue(row.presupuesto);
          groupedByDealId.set(normalisedDealId, {
            dealId: row.presupuesto || '',
            dealIdKey: normalisedDealId,
            rowIndexes: [],
            label: labelDealId ? `Presupuesto ${labelDealId}` : 'Presupuesto sin identificar'
          });
          queue.push(groupedByDealId.get(normalisedDealId));
        }

        const group = groupedByDealId.get(normalisedDealId);
        if (group) {
          group.rowIndexes.push(index);
        }
      } else {
        const studentName = buildStudentFullName(row);
        queue.push({
          dealId: '',
          dealIdKey: `row-${index}`,
          rowIndexes: [index],
          label: studentName ? `Fila ${index + 1} · ${studentName}` : `Fila ${index + 1}`
        });
      }
    });

    return queue;
  }

  function resolveGroupContact(queueItem) {
    const contact = {
      contactEmail: '',
      contactName: '',
      contactPersonId: ''
    };

    if (!queueItem) {
      return contact;
    }

    const dealIdKey = queueItem.dealIdKey;
    if (dealIdKey && state.dealContacts instanceof Map && state.dealContacts.has(dealIdKey)) {
      const storedContact = state.dealContacts.get(dealIdKey);
      if (storedContact) {
        contact.contactEmail = normaliseEmailInput(storedContact.contactEmail || '');
        contact.contactName = normaliseTextValue(storedContact.contactName);
        contact.contactPersonId = normaliseTextValue(storedContact.contactPersonId);
      }
    }

    const rowIndexes = Array.isArray(queueItem.rowIndexes) ? queueItem.rowIndexes : [];
    rowIndexes.forEach((rowIndex) => {
      const row = state.rows[rowIndex];
      if (!row) {
        return;
      }

      if (!contact.contactEmail) {
        contact.contactEmail = normaliseEmailInput(row.correoContacto || '');
      }

      if (!contact.contactName) {
        contact.contactName = normaliseTextValue(row.personaContacto);
      }

      if (!contact.contactPersonId) {
        contact.contactPersonId = normaliseTextValue(row.contactPersonId);
      }
    });

    return contact;
  }

  function populateEmailModalFieldsForBulk({ toValue, ccValue, bccValue, subject, body }) {
    if (elements.emailToInput) {
      elements.emailToInput.value = toValue || '';
    }
    if (elements.emailCcInput) {
      elements.emailCcInput.value = ccValue || '';
    }
    if (elements.emailBccInput) {
      elements.emailBccInput.value = bccValue || '';
    }
    if (elements.emailSubjectInput) {
      elements.emailSubjectInput.value = subject || '';
    }
    if (elements.emailSubjectPreview) {
      elements.emailSubjectPreview.textContent = subject || '';
    }
    if (elements.emailBodyInput) {
      elements.emailBodyInput.value = body || '';
    }
  }

  function buildEmailSubjectForRows(rows, queueItem) {
    if (!Array.isArray(rows) || rows.length === 0) {
      return 'Certificado de formación';
    }

    if (rows.length === 1) {
      return buildEmailSubject(rows[0]);
    }

    const referenceRow = rows[0];
    const trainingTitle = getResolvedTrainingTitle(referenceRow) || 'formación';
    const clientName = normaliseTextValue(referenceRow?.cliente);
    const dealId = queueItem ? normaliseTextValue(queueItem.dealId) : '';

    if (clientName && dealId) {
      return `Certificados - ${trainingTitle} · ${clientName} · Presupuesto ${dealId}`;
    }

    if (clientName) {
      return `Certificados - ${trainingTitle} · ${clientName}`;
    }

    if (dealId) {
      return `Certificados - ${trainingTitle} · Presupuesto ${dealId}`;
    }

    return `Certificados - ${trainingTitle}`;
  }

  function buildMultiStudentEmailBody(studentEntries) {
    if (!Array.isArray(studentEntries) || studentEntries.length === 0) {
      return '';
    }

    const referenceRow = studentEntries[0].row || {};
    const location = normaliseTextValue(referenceRow.lugar) || 'Barcelona';
    const trainingTitle = getResolvedTrainingTitle(referenceRow) || 'la formación indicada';
    const formattedDate = formatDateForEmail(referenceRow.fecha);

    const lines = [
      'Hola',
      '',
      `Adjuntamos los Certificados de los alumnos/as que han superado la formación de ${trainingTitle} en ${location} a fecha ${formattedDate}:`,
      '',
      ...studentEntries.map(({ row, link }, index) => {
        const name = buildStudentFullName(row) || `Alumno/a ${index + 1}`;
        return link ? `- ${name} (${link})` : `- ${name}`;
      }),
      '',
      'Responsable Formativo'
    ];

    return lines.join('\n');
  }

  function getBulkQueueItemLabel(queueItem) {
    if (!queueItem) {
      return '';
    }

    if (queueItem.label) {
      return queueItem.label;
    }

    const dealId = normaliseTextValue(queueItem.dealId);
    if (dealId) {
      return `Presupuesto ${dealId}`;
    }

    if (Array.isArray(queueItem.rowIndexes) && queueItem.rowIndexes.length) {
      const indexes = queueItem.rowIndexes.map((index) => index + 1).join(', ');
      return indexes ? `Filas ${indexes}` : '';
    }

    return '';
  }

  function buildBulkSummaryMessage(queueItem, message) {
    const label = getBulkQueueItemLabel(queueItem);
    const normalisedMessage = normaliseTextValue(message);

    if (label && normalisedMessage) {
      return `${label}: ${normalisedMessage}`;
    }

    return normalisedMessage || label || '';
  }

  function summariseBulkEmailResults(results, queue, { cancelled } = {}) {
    const processedResults = Array.isArray(results) ? results : [];
    const total = Array.isArray(queue) ? queue.length : processedResults.length;

    if (cancelled && processedResults.length < total) {
      const remaining = total - processedResults.length;
      const message =
        remaining > 0
          ? `El envío de correos se ha interrumpido. Quedaban ${remaining} correo(s) por enviar.`
          : 'El envío de correos se ha interrumpido antes de completarse.';
      showAlert('warning', message);
    }

    if (!processedResults.length) {
      return { autoClose: false };
    }

    const successResults = processedResults.filter((item) => item && item.success);
    const failureResults = processedResults.filter((item) => item && item.success === false);

    const failureMessages = [];
    failureResults.forEach((result) => {
      const summary = buildBulkSummaryMessage(result.queueItem, result.summaryMessage || result.statusMessage);
      const normalised = normaliseTextValue(summary);
      if (normalised && !failureMessages.includes(normalised)) {
        failureMessages.push(normalised);
      }
    });

    if (failureResults.length === 0) {
      const successCount = successResults.length;
      const message =
        successCount === 1
          ? 'Se ha enviado 1 correo correctamente.'
          : `Se han enviado correctamente ${successCount} correos.`;
      showAlert('success', message);
    } else if (successResults.length === 0) {
      const message =
        failureMessages.length > 0
          ? `No se ha podido enviar ningún correo. ${failureMessages.join(' ')}`
          : 'No se ha podido enviar ningún correo.';
      showAlert('danger', message);
    } else {
      const baseMessage = `Se han enviado ${successResults.length} correo${
        successResults.length === 1 ? '' : 's'
      }, pero ${failureResults.length} no se han podido enviar.`;
      const detail = failureMessages.length > 0 ? ` ${failureMessages.join(' ')}` : '';
      showAlert('warning', `${baseMessage}${detail}`);
    }

    const warningMessages = collectBulkWarningMessages(processedResults);
    warningMessages.forEach((warning) => {
      showAlert('warning', warning);
    });

    const shouldAutoClose = failureResults.length === 0 && !cancelled;
    return { autoClose: shouldAutoClose };
  }

  function collectBulkWarningMessages(results) {
    if (!Array.isArray(results)) {
      return [];
    }

    const seen = new Set();
    results.forEach((result) => {
      if (!result || !Array.isArray(result.warnings)) {
        return;
      }
      result.warnings.forEach((warning) => {
        const text = normaliseTextValue(warning && warning.message ? warning.message : warning);
        if (text && !seen.has(text)) {
          seen.add(text);
        }
      });
    });

    return Array.from(seen);
  }

  function setEmailBulkCounter(current, total) {
    const { emailBulkCounter } = elements;
    if (!emailBulkCounter) {
      return;
    }

    if (!total || total <= 0) {
      emailBulkCounter.textContent = '';
      emailBulkCounter.classList.add('d-none');
      return;
    }

    const safeCurrent = Math.max(0, Math.min(current, total));
    emailBulkCounter.textContent = `${safeCurrent}/${total}`;
    emailBulkCounter.classList.remove('d-none');
  }

  function delay(ms) {
    const duration = Number.isFinite(ms) ? Math.max(0, ms) : 0;
    return new Promise((resolve) => {
      window.setTimeout(resolve, duration);
    });
  }

  async function handleEmailFormSubmit(event) {
    event.preventDefault();

    if (emailModalState.isSending) {
      return;
    }

    if (emailModalState.isPreparingLink) {
      setEmailModalStatus('Espera a que el certificado esté listo para enviar.', 'info');
      return;
    }

    if (emailModalState.isBulkMode) {
      await handleBulkEmailFormSubmission();
      return;
    }

    const rowIndex = emailModalState.currentRowIndex;
    const row = state.rows[rowIndex];
    if (!row) {
      setEmailModalStatus('La fila seleccionada ya no está disponible.', 'danger');
      setEmailModalLoading(false);
      return;
    }

    const toValue = normaliseEmailInput(elements.emailToInput ? elements.emailToInput.value : '');
    const ccValue = normaliseEmailInput(elements.emailCcInput ? elements.emailCcInput.value : '');
    const bccValue = normaliseEmailInput(elements.emailBccInput ? elements.emailBccInput.value : '');

    if (elements.emailToInput) {
      elements.emailToInput.value = toValue;
    }
    if (elements.emailCcInput) {
      elements.emailCcInput.value = ccValue;
    }
    if (elements.emailBccInput) {
      elements.emailBccInput.value = bccValue;
    }
    const subjectValue =
      elements.emailSubjectInput && elements.emailSubjectInput.value
        ? elements.emailSubjectInput.value.trim()
        : buildEmailSubject(row);
    let bodyValue = elements.emailBodyInput ? elements.emailBodyInput.value : '';

    if (!toValue) {
      setEmailModalStatus('Introduce al menos un destinatario en el campo "Para".', 'warning');
      if (elements.emailToInput) {
        elements.emailToInput.focus();
      }
      return;
    }

    if (!hasValidEmailAddresses(toValue)) {
      setEmailModalStatus('Revisa las direcciones de correo del campo "Para".', 'warning');
      if (elements.emailToInput) {
        elements.emailToInput.focus();
      }
      return;
    }

    if (!hasValidEmailAddresses(ccValue)) {
      setEmailModalStatus('Revisa las direcciones de correo del campo "CC".', 'warning');
      if (elements.emailCcInput) {
        elements.emailCcInput.focus();
      }
      return;
    }

    if (!hasValidEmailAddresses(bccValue)) {
      setEmailModalStatus('Revisa las direcciones de correo del campo "CCO".', 'warning');
      if (elements.emailBccInput) {
        elements.emailBccInput.focus();
      }
      return;
    }

    setEmailModalLoading(true);
    setEmailModalStatus('Preparando el certificado…', 'info');

    const ensureResult = await ensureRowHasDriveFile(rowIndex, {
      onStatusChange: (message, type) => {
        if (message) {
          setEmailModalStatus(message, type);
        }
      }
    });

    if (ensureResult.error) {
      setEmailModalStatus(ensureResult.error.message, ensureResult.error.type || 'danger');
      setEmailModalLoading(false);
      return;
    }

    if (ensureResult.warning && ensureResult.warning.message) {
      setEmailModalStatus(ensureResult.warning.message, ensureResult.warning.type || 'warning');
    }

    const publicLink = ensureResult.link;
    if (!publicLink) {
      setEmailModalStatus(
        'No se ha podido obtener un enlace público al certificado. Revisa Google Drive e inténtalo de nuevo.',
        'danger'
      );
      setEmailModalLoading(false);
      return;
    }

    const updatedBody = ensureEmailBodyHasLink(bodyValue, publicLink);
    if (updatedBody.didUpdate) {
      bodyValue = updatedBody.updatedBody;
      if (elements.emailBodyInput) {
        elements.emailBodyInput.value = bodyValue;
      }
    }

    const { gmail, error: gmailError } = resolveGoogleGmailIntegration();
    if (!gmail) {
      if (gmailError) {
        setEmailModalStatus(gmailError.message, gmailError.type || 'danger');
      } else {
        setEmailModalStatus('La integración con Gmail no está disponible.', 'danger');
      }
      setEmailModalLoading(false);
      return;
    }

    setEmailModalStatus('Enviando correo…', 'info');

    const toForSending = formatEmailRecipientsForSending(toValue);
    const ccForSending = formatEmailRecipientsForSending(ccValue);
    const bccForSending = formatEmailRecipientsForSending(bccValue);

    try {
      await gmail.sendEmail({
        to: toForSending,
        cc: ccForSending || '',
        bcc: bccForSending || '',
        subject: subjectValue,
        body: bodyValue
      });

      setEmailModalStatus('Correo enviado correctamente.', 'success');

      if (toValue && row.correoContacto !== toValue) {
        updateRowValue(rowIndex, 'correoContacto', toValue);
      }

      if (row.presupuesto) {
        storeDealContact(row.presupuesto, {
          contactName: row.personaContacto || '',
          contactEmail: toValue,
          contactPersonId: row.contactPersonId || ''
        });
      }

      emailModalState.autoCloseTimeoutId = window.setTimeout(() => {
        if (emailModalState.modalInstance) {
          emailModalState.modalInstance.hide();
        }
      }, 1600);
    } catch (error) {
      console.error('No se ha podido enviar el correo', error);
      setEmailModalStatus(translateGmailError(error), 'danger');
      setEmailModalLoading(false);
      return;
    }

    setEmailModalLoading(false);
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

  function setupEmailModal() {
    if (emailModalState.initialised) {
      return;
    }

    const {
      emailModal,
      emailForm,
      emailToInput,
      emailCcInput,
      emailBccInput,
      emailBodyInput,
      emailSendButton,
      emailSubjectPreview
    } = elements;

    if (!emailModal || !emailForm || !emailToInput || !emailBodyInput || !emailSendButton || !emailSubjectPreview) {
      return;
    }

    const bootstrapModal = window.bootstrap && window.bootstrap.Modal ? window.bootstrap.Modal : null;
    if (!bootstrapModal) {
      console.warn('Bootstrap Modal no disponible para el envío de correos.');
      return;
    }

    emailModalState.modalInstance = bootstrapModal.getOrCreateInstance(emailModal);
    emailModalState.initialised = true;
    emailModalState.sendButtonOriginalLabel = emailSendButton.innerHTML;

    emailForm.addEventListener('submit', handleEmailFormSubmit);

    emailModal.addEventListener('shown.bs.modal', () => {
      window.setTimeout(() => {
        if (elements.emailToInput) {
          elements.emailToInput.focus();
          elements.emailToInput.select();
        }
      }, 150);
    });

    emailModal.addEventListener('hidden.bs.modal', () => {
      resetEmailModalState();
    });

    if (emailToInput) {
      emailToInput.addEventListener('input', () => {
        updateEmailSendButtonState();
      });
      emailToInput.addEventListener('blur', (event) => {
        const normalised = normaliseEmailInput(event.target.value);
        if (event.target.value !== normalised) {
          event.target.value = normalised;
        }
        updateEmailSendButtonState();
      });
    }

    [emailCcInput, emailBccInput].forEach((input) => {
      if (!input) {
        return;
      }
      input.addEventListener('blur', (event) => {
        event.target.value = normaliseEmailInput(event.target.value);
      });
    });

    updateEmailSendButtonState();
  }

  function resetEmailModalState() {
    if (emailModalState.autoCloseTimeoutId) {
      window.clearTimeout(emailModalState.autoCloseTimeoutId);
      emailModalState.autoCloseTimeoutId = null;
    }

    emailModalState.currentRowIndex = -1;
    emailModalState.isSending = false;
    emailModalState.isPreparingLink = false;
    emailModalState.isBulkMode = false;
    emailModalState.bulkQueue = [];
    emailModalState.bulkResults = [];
    emailModalState.bulkCurrentIndex = -1;
    emailModalState.bulkCurrentData = null;

    if (elements.emailForm) {
      elements.emailForm.reset();
    }

    if (elements.emailSubjectPreview) {
      elements.emailSubjectPreview.textContent = '';
    }

    if (elements.emailCcInput) {
      elements.emailCcInput.value = normaliseEmailInput(ACCOUNTING_EMAIL);
    }

    setEmailModalStatus('');
    setEmailBulkCounter(0, 0);
    setEmailModalLoading(false);
  }

  function populateEmailModalFields(row) {
    if (!row) {
      return;
    }

    if (elements.emailToInput) {
      elements.emailToInput.value = normaliseEmailInput(row.correoContacto || '');
    }

    if (elements.emailCcInput) {
      elements.emailCcInput.value = normaliseEmailInput(ACCOUNTING_EMAIL);
    }

    if (elements.emailBccInput) {
      elements.emailBccInput.value = '';
    }

    const subject = buildEmailSubject(row);
    if (elements.emailSubjectInput) {
      elements.emailSubjectInput.value = subject;
    }
    if (elements.emailSubjectPreview) {
      elements.emailSubjectPreview.textContent = subject;
    }

    const body = buildEmailBody(row);
    if (elements.emailBodyInput) {
      elements.emailBodyInput.value = body;
    }

    updateEmailSendButtonState();
  }

  async function prepareEmailModalCertificateLink(rowIndex) {
    if (emailModalState.currentRowIndex !== rowIndex) {
      return;
    }

    const row = state.rows[rowIndex];
    if (!row) {
      return;
    }

    emailModalState.isPreparingLink = true;
    const { emailSendButton, emailBodyInput } = elements;
    if (emailSendButton) {
      emailSendButton.setAttribute('aria-busy', 'true');
    }
    updateEmailSendButtonState();

    const updateStatus = (message, type) => {
      if (emailModalState.currentRowIndex !== rowIndex) {
        return;
      }
      if (!message) {
        setEmailModalStatus('');
        return;
      }
      setEmailModalStatus(message, type);
    };

    updateStatus('Preparando el certificado…', 'info');

    try {
      const ensureResult = (await ensureRowHasDriveFile(rowIndex, {
        onStatusChange: (message, type) => {
          if (emailModalState.currentRowIndex !== rowIndex) {
            return;
          }
          if (!message) {
            setEmailModalStatus('');
            return;
          }
          setEmailModalStatus(message, type);
        }
      })) || {};

      if (emailModalState.currentRowIndex !== rowIndex) {
        return;
      }

      if (ensureResult.error) {
        setEmailModalStatus(ensureResult.error.message, ensureResult.error.type || 'danger');
        return;
      }

      const publicLink = ensureResult.link || '';
      if (emailBodyInput) {
        const currentBody = emailBodyInput.value || '';
        const updatedBody = ensureEmailBodyHasLink(currentBody, publicLink);
        if (updatedBody.didUpdate) {
          emailBodyInput.value = updatedBody.updatedBody;
        }
      }

      if (ensureResult.warning && ensureResult.warning.message) {
        setEmailModalStatus(ensureResult.warning.message, ensureResult.warning.type || 'warning');
        return;
      }

      if (publicLink) {
        setEmailModalStatus('Certificado listo para enviar.', 'success');
      } else {
        setEmailModalStatus('');
      }
    } catch (error) {
      console.error('No se ha podido preparar el enlace del certificado', error);
      if (emailModalState.currentRowIndex === rowIndex) {
        setEmailModalStatus('No se ha podido preparar el enlace del certificado. Inténtalo de nuevo.', 'danger');
      }
    } finally {
      if (emailModalState.currentRowIndex === rowIndex) {
        emailModalState.isPreparingLink = false;
        if (emailSendButton) {
          emailSendButton.removeAttribute('aria-busy');
        }
        updateEmailSendButtonState();
      }
    }
  }

  function setEmailModalStatus(message, type = 'info') {
    const { emailStatus } = elements;
    if (!emailStatus) {
      return;
    }

    const normalisedMessage = normaliseTextValue(message);
    if (!normalisedMessage) {
      emailStatus.className = 'alert d-none';
      emailStatus.textContent = '';
      return;
    }

    emailStatus.className = `alert alert-${type} mb-0`;
    emailStatus.textContent = normalisedMessage;
  }

  function setEmailModalLoading(isLoading) {
    emailModalState.isSending = Boolean(isLoading);

    const controls = [
      elements.emailToInput,
      elements.emailCcInput,
      elements.emailBccInput,
      elements.emailBodyInput,
      elements.emailSubjectInput
    ];

    controls.forEach((control) => {
      if (!control) {
        return;
      }
      control.disabled = Boolean(isLoading);
    });

    const { emailSendButton } = elements;
    if (emailSendButton) {
      if (isLoading) {
        emailSendButton.disabled = true;
        emailSendButton.setAttribute('aria-busy', 'true');
        emailSendButton.innerHTML =
          '<span class="spinner-border spinner-border-sm me-2" role="status" aria-hidden="true"></span>Enviando…';
      } else {
        emailSendButton.disabled = false;
        emailSendButton.removeAttribute('aria-busy');
        emailSendButton.innerHTML = emailModalState.sendButtonOriginalLabel || 'Enviar correo';
        updateEmailSendButtonState();
      }
    }
  }

  function updateEmailSendButtonState() {
    const { emailSendButton, emailToInput } = elements;
    if (!emailSendButton) {
      return;
    }

    if (emailModalState.isSending || emailModalState.isPreparingLink) {
      emailSendButton.disabled = true;
      return;
    }

    if (emailModalState.isBulkMode && !emailModalState.bulkCurrentData) {
      emailSendButton.disabled = true;
      return;
    }

    const hasRecipient = normaliseTextValue(emailToInput ? emailToInput.value : '');
    emailSendButton.disabled = !hasRecipient;
  }

  function splitEmailRecipients(value) {
    const text = normaliseTextValue(value);
    if (!text) {
      return [];
    }

    return text
      .split(EMAIL_LIST_DELIMITER_PATTERN)
      .map((part) => part.trim())
      .filter((part) => part.length > 0);
  }

  function normaliseEmailInput(value) {
    const recipients = splitEmailRecipients(value);
    if (!recipients.length) {
      return '';
    }

    return recipients.join(`${EMAIL_SEPARATOR} `);
  }

  function hasValidEmailAddresses(value) {
    const recipients = splitEmailRecipients(value);
    if (!recipients.length) {
      return true;
    }

    return recipients.every((recipient) => EMAIL_VALIDATION_PATTERN.test(recipient));
  }

  function formatEmailRecipientsForSending(value) {
    const recipients = splitEmailRecipients(value);
    if (!recipients.length) {
      return '';
    }

    return recipients.join(', ');
  }

  function buildEmailSubject(row) {
    if (!row) {
      return 'Certificado de formación';
    }

    const trainingTitle = getResolvedTrainingTitle(row) || 'formación';
    const studentName = buildStudentFullName(row) || 'Alumno/a';

    return `Certificado - ${trainingTitle} · ${studentName}`;
  }

  function buildEmailBody(row) {
    if (!row) {
      return '';
    }

    const studentName = buildStudentFullName(row) || 'Alumno/a';
    const formattedDate = formatDateForEmail(row.fecha);
    const location = normaliseTextValue(row.lugar) || 'Barcelona';
    const trainingTitle = getResolvedTrainingTitle(row) || 'la formación indicada';
    const link = normaliseTextValue(row.driveFileUrl) || '';

    const lines = [
      'Hola',
      '',
      `Adjuntamos Certificado del alumno/a ${studentName} quien en fecha ${formattedDate} y en ${location} ha superado la formación de ${trainingTitle}`,
      '',
      'Responsable Formativo',
      '',
      link ? `Descargar certificado: ${link}` : 'Descargar certificado:'
    ];

    return lines.join('\n');
  }

  function formatDateForEmail(value) {
    const iso = normaliseDateValue(value);
    if (iso) {
      const date = new Date(`${iso}T00:00:00Z`);
      if (!Number.isNaN(date.getTime())) {
        return new Intl.DateTimeFormat('es-ES', {
          day: 'numeric',
          month: 'long',
          year: 'numeric'
        }).format(date);
      }
    }

    const fallback = normaliseTextValue(value);
    return fallback || '[fecha pendiente]';
  }

  function ensureEmailBodyHasLink(body, link) {
    const normalisedLink = normaliseTextValue(link);
    const normalisedBody = body || '';

    if (!normalisedLink) {
      return { updatedBody: normalisedBody, didUpdate: false };
    }

    if (normalisedBody.includes(normalisedLink)) {
      return { updatedBody: normalisedBody, didUpdate: false };
    }

    const linkLabelPattern = /(Descargar certificado:\s*)(.*)/i;
    if (linkLabelPattern.test(normalisedBody)) {
      const updatedBody = normalisedBody.replace(linkLabelPattern, `$1${normalisedLink}`);
      return { updatedBody, didUpdate: true };
    }

    const updatedBody = `${normalisedBody.trim()}

Descargar certificado: ${normalisedLink}`.trim();
    return { updatedBody, didUpdate: true };
  }

  async function ensureRowHasDriveFile(rowIndex, { onStatusChange } = {}) {
    const row = state.rows[rowIndex];
    if (!row) {
      return {
        error: {
          type: 'danger',
          message: 'No se ha encontrado la fila seleccionada.'
        }
      };
    }

    const rawDriveFileUrl = row.driveFileUrl === undefined || row.driveFileUrl === null ? '' : String(row.driveFileUrl);
    const rawDriveFileDownloadUrl =
      row.driveFileDownloadUrl === undefined || row.driveFileDownloadUrl === null
        ? ''
        : String(row.driveFileDownloadUrl);
    const rawDriveFileId = row.driveFileId === undefined || row.driveFileId === null ? '' : String(row.driveFileId);

    const normalisedDriveFileUrl = normaliseTextValue(rawDriveFileUrl);
    const normalisedDriveFileDownloadUrl = normaliseTextValue(rawDriveFileDownloadUrl);
    const normalisedDriveFileId = normaliseTextValue(rawDriveFileId);

    const metadataChanged =
      normalisedDriveFileUrl !== rawDriveFileUrl ||
      normalisedDriveFileDownloadUrl !== rawDriveFileDownloadUrl ||
      normalisedDriveFileId !== rawDriveFileId;

    if (metadataChanged) {
      Object.assign(row, {
        driveFileUrl: normalisedDriveFileUrl,
        driveFileDownloadUrl: normalisedDriveFileDownloadUrl,
        driveFileId: normalisedDriveFileId
      });
      persistRows();
    }

    if (normalisedDriveFileUrl) {
      return {
        link: normalisedDriveFileUrl,
        metadata: {
          driveFileUrl: normalisedDriveFileUrl,
          driveFileDownloadUrl: normalisedDriveFileDownloadUrl,
          driveFileId: normalisedDriveFileId || ''
        },
        warning: null
      };
    }

    const { drive, error: driveError } = resolveGoogleDriveIntegration();
    if (!drive) {
      if (normalisedDriveFileDownloadUrl) {
        return {
          link: normalisedDriveFileDownloadUrl,
          metadata: {
            driveFileUrl: '',
            driveFileDownloadUrl: normalisedDriveFileDownloadUrl,
            driveFileId: normalisedDriveFileId || ''
          },
          warning: null
        };
      }

      return {
        error:
          driveError || {
            type: 'danger',
            message: 'La integración con Google Drive no está disponible.'
          }
      };
    }

    if (normalisedDriveFileId && typeof drive.ensurePublicFileAccess === 'function') {
      try {
        const shared = await drive.ensurePublicFileAccess({
          fileId: normalisedDriveFileId,
          webViewLink: normalisedDriveFileUrl || ''
        });
        const link = shared?.webViewLink || shared?.downloadLink || '';
        if (link) {
          const metadata = {
            driveFileUrl: shared.webViewLink || link,
            driveFileDownloadUrl: shared.downloadLink || normalisedDriveFileDownloadUrl || '',
            driveFileId: normalisedDriveFileId
          };
          Object.assign(row, metadata);
          persistRows();
          return {
            link: metadata.driveFileUrl || metadata.driveFileDownloadUrl || link,
            metadata,
            warning: null
          };
        }
      } catch (error) {
        console.warn('No se ha podido actualizar el enlace público del archivo existente en Google Drive.', error);
      }
    }

    const { certificate, error: certificateError } = resolveCertificateModule();
    if (!certificate) {
      return {
        error:
          certificateError || {
            type: 'danger',
            message: 'La librería de certificados no está disponible.'
          }
      };
    }

    if (typeof onStatusChange === 'function') {
      onStatusChange('Generando el certificado en PDF…', 'info');
    }

    try {
      const result = await uploadRowCertificateToDrive(rowIndex, { drive, certificate });
      if (result && result.warning && typeof onStatusChange === 'function') {
        onStatusChange(result.warning.message, result.warning.type || 'warning');
      }

      const metadata = result ? result.metadata : null;
      const link =
        metadata && (metadata.driveFileUrl || metadata.driveFileDownloadUrl)
          ? metadata.driveFileUrl || metadata.driveFileDownloadUrl
          : '';

      return {
        link,
        metadata,
        warning: result ? result.warning : null
      };
    } catch (error) {
      const userFacing = error && error.userFacing ? error.userFacing : null;
      return {
        error:
          userFacing || {
            type: 'danger',
            message: translateGoogleDriveError(error)
          }
      };
    }
  }

  async function uploadRowCertificateToDrive(rowIndex, { drive, certificate }) {
    const row = state.rows[rowIndex];
    if (!row) {
      const error = new Error('No se ha encontrado la fila seleccionada.');
      error.userFacing = {
        type: 'danger',
        message: 'No se ha encontrado la fila seleccionada.'
      };
      throw error;
    }

    const driveDetails = buildDriveUploadDetails(row, certificate);
    if (driveDetails.error) {
      const error = new Error(driveDetails.error.message);
      error.userFacing = driveDetails.error;
      throw error;
    }

    const { clientFolderName, trainingFolderName, fileName } = driveDetails;

    let generated;
    try {
      generated = await certificate.generate(row, { download: false });
    } catch (error) {
      const enriched = error instanceof Error ? error : new Error('No se ha podido generar el certificado en PDF.');
      if (!enriched.userFacing) {
        enriched.userFacing = {
          type: 'danger',
          message: 'No se ha podido generar el certificado en PDF. Revisa los datos e inténtalo de nuevo.'
        };
      }
      throw enriched;
    }

    const blob = generated && generated.blob ? generated.blob : null;
    if (!(blob instanceof Blob)) {
      const error = new Error('No se ha podido generar el archivo PDF del certificado.');
      error.userFacing = {
        type: 'danger',
        message: 'No se ha podido generar el certificado en PDF. Revisa los datos e inténtalo de nuevo.'
      };
      throw error;
    }

    let uploadResult;
    try {
      uploadResult = await drive.uploadCertificate({
        clientFolderName,
        trainingFolderName,
        fileName,
        blob
      });
    } catch (error) {
      const enriched = error instanceof Error ? error : new Error('No se ha podido guardar el certificado en Google Drive.');
      enriched.userFacing = {
        type: 'danger',
        message: translateGoogleDriveError(error)
      };
      throw enriched;
    }

    let publicMetadata = null;
    let warning = null;
    if (uploadResult && uploadResult.fileId && typeof drive.ensurePublicFileAccess === 'function') {
      try {
        publicMetadata = await drive.ensurePublicFileAccess({
          fileId: uploadResult.fileId,
          webViewLink: uploadResult.webViewLink || ''
        });
      } catch (error) {
        console.warn('No se ha podido asegurar el acceso público al archivo de Google Drive.', error);
        warning = {
          type: 'warning',
          message:
            'El certificado se ha guardado en Google Drive, pero no se ha podido actualizar el enlace público automáticamente. Revisa los permisos del archivo antes de compartirlo.'
        };
      }
    }

    const downloadLink =
      (publicMetadata && publicMetadata.downloadLink) ||
      (uploadResult && uploadResult.fileId && typeof drive.buildPublicDownloadLink === 'function'
        ? drive.buildPublicDownloadLink(uploadResult.fileId)
        : '');
    const publicLink =
      (publicMetadata && publicMetadata.webViewLink) ||
      (uploadResult && uploadResult.webViewLink) ||
      downloadLink;

    const metadata = {
      driveFileId: (uploadResult && uploadResult.fileId) || '',
      driveFileName: (uploadResult && uploadResult.fileName) || '',
      driveFileUrl: warning ? '' : publicLink || '',
      driveFileDownloadUrl: downloadLink || '',
      driveClientFolderId: (uploadResult && uploadResult.clientFolderId) || '',
      driveTrainingFolderId: (uploadResult && uploadResult.trainingFolderId) || '',
      driveUploadedAt: new Date().toISOString()
    };

    Object.assign(row, metadata);
    persistRows();

    return { metadata, warning };
  }

  function resolveGoogleGmailIntegration() {
    const gmail = window.googleGmail;
    if (!gmail || typeof gmail.sendEmail !== 'function') {
      return {
        gmail: null,
        error: {
          type: 'danger',
          message: 'La integración con Gmail no está disponible en esta página.'
        }
      };
    }

    if (typeof gmail.isConfigured === 'function' && !gmail.isConfigured()) {
      return {
        gmail: null,
        error: {
          type: 'danger',
          message: 'Falta configurar la integración con Gmail.'
        }
      };
    }

    return { gmail };
  }

  function translateGmailError(error) {
    if (!error) {
      return 'No se ha podido enviar el correo. Inténtalo de nuevo.';
    }

    if (error.code === 'access_denied') {
      return 'Se ha cancelado la autorización de Gmail.';
    }

    if (error instanceof TypeError) {
      return 'No se ha podido conectar con Gmail. Comprueba tu conexión e inténtalo de nuevo.';
    }

    if (typeof error.status === 'number' && error.status === 401) {
      return 'Gmail ha rechazado la autorización. Vuelve a iniciar sesión e inténtalo de nuevo.';
    }

    const message = normaliseTextValue(error.message);
    if (message) {
      return message;
    }

    return 'No se ha podido enviar el correo. Inténtalo de nuevo.';
  }

  function updateActionButtonsState() {
    const hasRows = state.rows.length > 0;
    elements.clearStorage.disabled = !hasRows;
    elements.generateAllCertificates.disabled = !hasRows;
    if (elements.sendAllToDrive) {
      elements.sendAllToDrive.disabled = !hasRows;
    }
    if (elements.sendAllByEmail) {
      elements.sendAllByEmail.disabled = !hasRows;
    }
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

  function normaliseTextValue(value) {
    if (value === undefined || value === null) {
      return '';
    }
    return String(value).trim();
  }

  function sanitiseDriveComponent(value, fallback) {
    const cleaned = normaliseTextValue(value)
      .replace(/[\n\r]/g, ' ')
      .replace(/[\\/:*?"<>|]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
    return cleaned || fallback;
  }

  function getResolvedTrainingTitle(row) {
    if (!row) {
      return '';
    }

    if (window.certificatePdf && typeof window.certificatePdf.resolveTrainingTitle === 'function') {
      const resolved = window.certificatePdf.resolveTrainingTitle(row);
      const normalisedResolved = normaliseTextValue(resolved);
      if (normalisedResolved) {
        return normalisedResolved;
      }
    }

    if (trainingTemplates && typeof trainingTemplates.getTrainingTitle === 'function') {
      const templateTitle = trainingTemplates.getTrainingTitle(row.formacion);
      const normalisedTemplate = normaliseTextValue(templateTitle);
      if (normalisedTemplate) {
        return normalisedTemplate;
      }
    }

    return normaliseTextValue(row.formacion);
  }

  function buildStudentFullName(row) {
    if (!row) {
      return '';
    }
    const name = normaliseTextValue(row.nombre);
    const surname = normaliseTextValue(row.apellido);
    return [name, surname].filter(Boolean).join(' ').trim();
  }

  function formatDateForFolderName(value) {
    const iso = normaliseDateValue(value);
    if (iso) {
      return iso;
    }
    const fallback = normaliseTextValue(value);
    return fallback || 'Fecha sin definir';
  }

  function formatDateForFileLabel(value) {
    if (window.certificatePdf && typeof window.certificatePdf.formatDateForFileName === 'function') {
      const formatted = window.certificatePdf.formatDateForFileName(value);
      const normalisedFormatted = normaliseTextValue(formatted);
      if (normalisedFormatted) {
        return normalisedFormatted;
      }
    }

    const iso = normaliseDateValue(value);
    if (iso) {
      const [year, month, day] = iso.split('-');
      return `${day}-${month}-${year}`;
    }

    return normaliseTextValue(value) || 'Fecha sin definir';
  }

  function translateGoogleDriveError(error) {
    if (!error) {
      return 'No se ha podido guardar el certificado en Google Drive. Inténtalo de nuevo.';
    }

    if (error.code === 'access_denied') {
      return 'Se ha cancelado la autorización de Google Drive.';
    }

    if (error instanceof TypeError) {
      return 'No se ha podido conectar con Google Drive. Comprueba tu conexión e inténtalo de nuevo.';
    }

    if (typeof error.status === 'number' && error.status === 401) {
      return 'Google Drive ha rechazado la autorización. Vuelve a iniciar sesión e inténtalo de nuevo.';
    }

    const message = normaliseTextValue(error.message);
    if (message) {
      return message;
    }

    return 'No se ha podido guardar el certificado en Google Drive. Inténtalo de nuevo.';
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
