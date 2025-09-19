(function (global) {
  const STORAGE_KEY = 'gep-certificados/training-templates/v1';

  const DEFAULT_DURATION_ENTRIES = [
    ['Pack Emergencias', '6h'],
    ['Trabajos en Altura', '8h'],
    ['Trabajos Verticales', '12h'],
    ['Carretilla elevadora', '8h'],
    ['Espacios Confinados', '8h'],
    ['Operaciones Telco', '6h'],
    ['Riesgo Eléctrico Telco', '6h'],
    ['Espacios Confinados Telco', '6h'],
    ['Trabajos en altura Telco', '6h'],
    ['Basico de Fuego', '4h'],
    ['Avanzado de Fuego', '5h'],
    ['Avanzado y Casa de Humo', '6h'],
    ['Riesgo Químico', '4h'],
    ['Primeros Auxilios', '4h'],
    ['SVD y DEA', '6h'],
    ['Implantación de PAU', '6h'],
    ['Jefes de Emergencias', '8h'],
    ["Curso de ERA's", '8h'],
    ['Andamios', '8h'],
    ['Renovación Bombero de Empresa', '20h'],
    ['Bombero de Empresa Inicial', '350h']
  ];

  const DEFAULT_DETAILS_ENTRIES = [
    [
      'Pack Emergencias',
      {
        theory: [
          'Proceso de la combustión.',
          'Clases de fuego.',
          'Clasificación de combustible.',
          'Propagación de un incendio.',
          'Transferencia de calor.',
          'Agentes extintores.',
          'Mecanismos de extinción.',
          'Protocolo de actuación (P.A.S.).',
          'Valoración primaria.',
          'Posiciones de espera y traslado (P.L.S.).',
          'O.V.A.C.E. obstrucción vía aérea (Maniobra de Heimlich).'
        ],
        practice: [
          'Reconocer diferentes tipos de extintores.',
          'Trabajo en interior con extintor de polvo.',
          'Extinción con CO₂ en armario eléctrico.',
          'Extinción y cubrimiento con extintor hídrico.',
          'Ejercicio/simulacro con diferentes grados de dificultad, corte de suministro, víctima consciente e inconsciente.',
          'Apertura de puertas y comprobación de temperatura.'
        ]
      }
    ],
    [
      'Trabajos en Altura',
      {
        theory: [
          'Legislación y normativa vigente.',
          'Medidas de protección preventiva.',
          'Conocimientos generales de seguridad en altura.',
          'Equipos de protección individual y colectiva.',
          'Instalaciones horizontales y verticales.',
          'Actuación en caso de emergencia.'
        ],
        practice: [
          'Los nudos básicos y su realización.',
          'Utilización de equipos de protección individual.',
          'Utilización de los equipos de protección colectiva.',
          'Instalación de líneas de vida horizontales y verticales.',
          'Puntos de anclaje.',
          'Técnicas de acceso y posicionamiento en altura.',
          'Rescate de emergencia.'
        ]
      }
    ],
    [
      'Trabajos Verticales',
      {
        theory: [
          'Legislación y normativa vigente.',
          'Medidas de protección preventiva.',
          'Procedimientos de actuación y rescate.',
          'Conocimientos generales sobre seguridad.',
          'Equipos de protección individual y colectiva.',
          'Sistemas de instalación vertical y horizontal.',
          'Actuación en caso de emergencia y protocolo P.A.S.'
        ],
        practice: [
          'Taller de nudos y anclajes.',
          'Utilización y pruebas con los equipos de protección individual.',
          'Montaje de instalaciones verticales.',
          'Protección de las instalaciones.',
          'Paso de nudos.',
          'Ascenso y descenso con cuerdas, paso de nudos y cambios de cuerda.',
          'Rescate de víctimas y evacuación segura.'
        ]
      }
    ],
    [
      'Carretilla elevadora',
      {
        theory: [
          'Introducción a las carretillas.',
          'Legislación y normativa vigente.',
          'Tipos y componentes de la carretilla.',
          'El operador de carretillas.',
          'Riesgos y medidas preventivas.',
          'Normas básicas de utilización y manejo.',
          'Mantenimiento y seguridad de la carretilla.',
          'Conceptos sobre la manipulación de cargas.'
        ],
        practice: [
          'Presentación física de la carretilla.',
          'Normas básicas de utilización.',
          'Control y mantenimiento de la máquina.',
          'Comprobaciones previas de seguridad.',
          'La puesta en funcionamiento.',
          'Maniobras básicas de circulación.',
          'Manipulación y movimiento de cargas.'
        ]
      }
    ],
    [
      'Andamios',
      {
        theory: [
          'Legislación vigente.',
          'Generalidades de Prevención de Riesgos Laborales.',
          'Equipos de Protección Individual (EPI) y herramientas.',
          'Acceso por estructuras.',
          'Líneas de Vida horizontales y verticales.',
          'Elevación de cargas.',
          'Manipulación de cargas.',
          'Actuación en caso de emergencia.',
          'Primeros Auxilios y RCP (Reanimación Cardiopulmonar).'
        ],
        practice: [
          'Montaje y desmontaje de estructuras (ej. bastidas / andamios tipo Lyher y LITEC).',
          'Elevación de cargas.',
          'Manipulación de cargas.',
          'Trabajos en suspensión, posicionamiento y uso de polipastos, truss, etc.',
          'Técnicas de rescate, evacuación de víctima.',
          'Simulacro de emergencia.'
        ]
      }
    ],
    [
      'Operaciones Telco',
      {
        theory: [
          'Definición de los trabajos.',
          'Técnicas preventivas específicas.',
          'Evaluación de riesgos y plan de seguridad y salud.',
          'Protecciones colectivas (colocación, usos, obligaciones y mantenimiento).',
          'Protecciones individuales (colocación, usos, obligaciones y mantenimiento).',
          'Riesgo eléctrico.',
          'Caída de personas al mismo nivel.',
          'Espacios confinados.',
          'Riesgos de protecciones de partículas, golpes y cortes.',
          'Riesgos biológicos.',
          'Riesgos químicos.',
          'Iluminación, ruido, vibraciones, condiciones climatológicas.',
          'Radiaciones no ionizantes (campos electromagnéticos en telefonía móvil).',
          'Trabajos de fusionado de fibra óptica.',
          'Trabajos en salas de equipos.',
          'Manejo manual de cargas.',
          'Trabajos en vía pública y seguridad vial.',
          'Consignas de actuación ante emergencias y accidentes.',
          'Protocolo PAS (P.A.S.).'
        ],
        practice: [
          'Especialidad a elegir según necesidades reales del operario.',
          'Aplicación práctica de identificaciones de origen y causas de riesgos.',
          'Uso práctico de protecciones colectivas.',
          'Uso práctico de protecciones individuales.'
        ]
      }
    ],
    [
      'Riesgo Eléctrico Telco',
      {
        theory: [
          'Definición de los trabajos.',
          'Técnicas preventivas específicas.',
          'Evaluación de riesgos y plan de seguridad y salud.',
          'Protecciones colectivas (colocación, usos, obligaciones y mantenimiento).',
          'Protecciones individuales (colocación, usos, obligaciones y mantenimiento).',
          'Riesgo eléctrico.',
          'Caída de personas al mismo nivel.',
          'Caída de personas a distinto nivel.',
          'Espacios confinados.',
          'Riesgos de protecciones de partículas, golpes y cortes.',
          'Riesgos biológicos.',
          'Riesgos químicos.',
          'Iluminación, ruido, vibraciones, condiciones climatológicas.',
          'Radiaciones no ionizantes (campos electromagnéticos en telefonía móvil).',
          'Trabajos de fusionado de fibra óptica.',
          'Trabajos en salas de equipos.',
          'Manejo manual de cargas.',
          'Trabajos en vía pública y seguridad vial.',
          'Consignas de actuación ante emergencias y accidentes.',
          'Protocolo PAS.'
        ],
        practice: [
          'Detectar situaciones con riesgo eléctrico.',
          'Operar sin sobreexposición eléctrica, aplicando medidas correctas según tensión presente.'
        ]
      }
    ]
  ];

  function normaliseName(value) {
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

  function sanitiseText(value) {
    if (value === undefined || value === null) {
      return '';
    }
    return String(value).trim();
  }

  function sanitiseList(items) {
    if (!Array.isArray(items)) {
      return [];
    }
    return items
      .map((item) => sanitiseText(item))
      .filter((text) => text !== '');
  }

  function cloneTemplate(template) {
    if (!template) {
      return null;
    }
    return {
      id: template.id,
      name: template.name,
      title: template.title,
      duration: template.duration,
      theory: [...template.theory],
      practice: [...template.practice]
    };
  }

  function createCustomId() {
    return `custom-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
  }

  function sanitiseTemplate(template) {
    if (!template || typeof template !== 'object') {
      return {
        id: '',
        name: '',
        title: '',
        duration: '',
        theory: [],
        practice: []
      };
    }

    const name = sanitiseText(template.name);
    const title = sanitiseText(template.title) || name;

    return {
      id: template.id ? String(template.id) : '',
      name,
      title,
      duration: sanitiseText(template.duration),
      theory: sanitiseList(template.theory),
      practice: sanitiseList(template.practice)
    };
  }

  function buildDefaultTemplates() {
    const templatesByName = new Map();
    const durationLookup = new Map();

    DEFAULT_DURATION_ENTRIES.forEach(([name, duration]) => {
      const normalisedName = normaliseName(name);
      if (!normalisedName) {
        return;
      }
      const sanitisedDuration = sanitiseText(duration);
      if (sanitisedDuration) {
        durationLookup.set(normalisedName, sanitisedDuration);
      }
      const existing = templatesByName.get(normalisedName) || {
        id: `default-${normalisedName}`,
        name: sanitiseText(name),
        title: sanitiseText(name),
        duration: '',
        theory: [],
        practice: []
      };
      existing.duration = sanitisedDuration;
      templatesByName.set(normalisedName, existing);
    });

    DEFAULT_DETAILS_ENTRIES.forEach(([name, details]) => {
      const normalisedName = normaliseName(name);
      if (!normalisedName) {
        return;
      }
      const existing = templatesByName.get(normalisedName) || {
        id: `default-${normalisedName}`,
        name: sanitiseText(name),
        title: sanitiseText(name),
        duration: '',
        theory: [],
        practice: []
      };
      existing.name = sanitiseText(name) || existing.name;
      existing.title = sanitiseText(details && details.title) || existing.name || existing.title;
      existing.theory = sanitiseList(details && details.theory);
      existing.practice = sanitiseList(details && details.practice);
      if (!existing.duration) {
        let matchedDuration = durationLookup.get(normalisedName) || '';
        if (!matchedDuration) {
          for (const [durationKey, durationValue] of durationLookup.entries()) {
            if (durationKey && normalisedName.startsWith(durationKey)) {
              matchedDuration = durationValue;
              break;
            }
          }
        }
        if (matchedDuration) {
          existing.duration = matchedDuration;
        }
      }
      templatesByName.set(normalisedName, existing);
    });

    return Array.from(templatesByName.entries()).map(([key, template]) => {
      const sanitised = sanitiseTemplate(template);
      sanitised.id = template.id || `default-${key}`;
      if (!sanitised.title) {
        sanitised.title = sanitised.name;
      }
      return sanitised;
    });
  }

  function mergeTemplates(defaultTemplates, customTemplates) {
    const templatesByName = new Map();

    defaultTemplates.forEach((template) => {
      const normalisedName = normaliseName(template.name);
      if (!normalisedName) {
        return;
      }
      templatesByName.set(normalisedName, {
        ...template,
        id: template.id || `default-${normalisedName}`,
        theory: [...template.theory],
        practice: [...template.practice]
      });
    });

    customTemplates.forEach((template) => {
      const sanitised = sanitiseTemplate(template);
      const normalisedName = normaliseName(sanitised.name);
      if (!normalisedName) {
        return;
      }
      const customId = sanitised.id || createCustomId();
      templatesByName.set(normalisedName, {
        ...sanitised,
        id: customId,
        theory: [...sanitised.theory],
        practice: [...sanitised.practice]
      });
    });

    return Array.from(templatesByName.values());
  }

  function loadCustomTemplates() {
    try {
      if (!global.localStorage) {
        return [];
      }
      const raw = global.localStorage.getItem(STORAGE_KEY);
      if (!raw) {
        return [];
      }
      const parsed = JSON.parse(raw);
      if (!Array.isArray(parsed)) {
        return [];
      }
      return parsed.map((template) => {
        const sanitised = sanitiseTemplate(template);
        if (!sanitised.id) {
          sanitised.id = createCustomId();
        }
        if (!sanitised.title) {
          sanitised.title = sanitised.name;
        }
        return sanitised;
      });
    } catch (error) {
      console.warn('No se han podido cargar las plantillas personalizadas', error);
      return [];
    }
  }

  function persistCustomTemplates(customTemplates) {
    try {
      if (!global.localStorage) {
        return;
      }
      const serialisable = customTemplates.map((template) => ({
        id: template.id,
        name: template.name,
        title: template.title,
        duration: template.duration,
        theory: [...template.theory],
        practice: [...template.practice]
      }));
      global.localStorage.setItem(STORAGE_KEY, JSON.stringify(serialisable));
    } catch (error) {
      console.warn('No se han podido guardar las plantillas personalizadas', error);
    }
  }

  let customTemplates = loadCustomTemplates();
  let templates = [];
  let sortedTemplates = [];
  let templatesByName = new Map();
  let templatesById = new Map();
  const subscribers = new Set();

  function refreshTemplates() {
    const defaultTemplates = buildDefaultTemplates();
    templates = mergeTemplates(defaultTemplates, customTemplates);

    templatesByName = new Map();
    templatesById = new Map();

    templates.forEach((template) => {
      const normalisedName = normaliseName(template.name);
      const id = template.id ? String(template.id) : '';
      if (normalisedName) {
        templatesByName.set(normalisedName, template);
      }
      if (id) {
        templatesById.set(id, template);
      }
    });

    sortedTemplates = templates
      .slice()
      .sort((a, b) => a.name.localeCompare(b.name, 'es', { sensitivity: 'base' }));
  }

  function notifySubscribers() {
    subscribers.forEach((callback) => {
      try {
        callback();
      } catch (error) {
        console.error('Error al notificar los cambios de plantillas', error);
      }
    });
  }

  refreshTemplates();

  function listTemplates() {
    return sortedTemplates.map((template) => cloneTemplate(template));
  }

  function getTemplateByName(name) {
    const normalisedName = normaliseName(name);
    if (!normalisedName) {
      return null;
    }
    const template = templatesByName.get(normalisedName);
    return cloneTemplate(template);
  }

  function getTemplateById(id) {
    if (!id) {
      return null;
    }
    const template = templatesById.get(String(id));
    return cloneTemplate(template);
  }

  function getTrainingDuration(name) {
    const template = getTemplateByName(name);
    return template ? template.duration : '';
  }

  function getTrainingDetails(name) {
    const template = getTemplateByName(name);
    if (!template) {
      return null;
    }
    return {
      theory: [...template.theory],
      practice: [...template.practice]
    };
  }

  function getTrainingTitle(name) {
    const template = getTemplateByName(name);
    if (!template) {
      return sanitiseText(name);
    }
    return template.title || template.name;
  }

  function saveTemplate(template) {
    const sanitised = sanitiseTemplate(template);
    if (!sanitised.name) {
      throw new Error('El nombre de la formación es obligatorio.');
    }
    if (!sanitised.title) {
      sanitised.title = sanitised.name;
    }
    const normalisedName = normaliseName(sanitised.name);
    if (!normalisedName) {
      throw new Error('El nombre de la formación es obligatorio.');
    }

    if (!sanitised.id) {
      const existing = customTemplates.find((item) => normaliseName(item.name) === normalisedName);
      sanitised.id = existing ? existing.id : createCustomId();
    }

    let stored = false;
    customTemplates = customTemplates.map((item) => {
      if (item.id === sanitised.id) {
        stored = true;
        return { ...sanitised };
      }
      if (normaliseName(item.name) === normalisedName) {
        stored = true;
        return { ...sanitised, id: item.id };
      }
      return item;
    });

    if (!stored) {
      customTemplates.push({ ...sanitised });
    }

    persistCustomTemplates(customTemplates);
    refreshTemplates();
    notifySubscribers();

    return getTemplateById(sanitised.id) || getTemplateByName(sanitised.name);
  }

  function createEmptyTemplate() {
    return {
      id: '',
      name: '',
      title: '',
      duration: '',
      theory: [],
      practice: []
    };
  }

  function isCustomTemplateId(id) {
    if (!id) {
      return false;
    }
    const stringId = String(id);
    return customTemplates.some((template) => template.id === stringId);
  }

  function deleteTemplate(id) {
    if (!id) {
      return false;
    }

    const stringId = String(id);
    const templateToRemove = customTemplates.find((template) => template.id === stringId);
    if (!templateToRemove) {
      return false;
    }

    customTemplates = customTemplates.filter((template) => template.id !== stringId);
    persistCustomTemplates(customTemplates);
    refreshTemplates();
    notifySubscribers();

    return true;
  }

  function subscribe(callback) {
    if (typeof callback !== 'function') {
      return () => {};
    }
    subscribers.add(callback);
    return () => {
      subscribers.delete(callback);
    };
  }

  global.trainingTemplates = {
    listTemplates,
    getTemplateByName,
    getTemplateById,
    getTrainingDuration,
    getTrainingDetails,
    getTrainingTitle,
    saveTemplate,
    createEmptyTemplate,
    deleteTemplate,
    isCustomTemplateId,
    subscribe,
    normaliseName
  };
})(typeof window !== 'undefined' ? window : this);
