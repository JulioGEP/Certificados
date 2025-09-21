const TRAINING_DATE_FIELD = '98f072a788090ac2ae52017daaf9618c3a189033';
const TRAINING_LOCATION_FIELD = '676d6bd51e52999c582c01f67c99a35ed30bf6ae';
const TRAINING_NAME_FIELD = 'c99554c188c3f63ad9bc8b2cf7b50cbd145455ab';

const HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'Content-Type',
  'Access-Control-Allow-Methods': 'GET,OPTIONS'
};

const fieldOptionsCache = new Map();
let cachedDealFields = null;
let dealFieldsPromise = null;

exports.handler = async (event) => {
  if (event.httpMethod === 'OPTIONS') {
    return {
      statusCode: 204,
      headers: HEADERS,
      body: ''
    };
  }

  if (event.httpMethod !== 'GET') {
    return jsonResponse(405, { success: false, message: 'Método no permitido.' });
  }

  const { dealId } = event.queryStringParameters || {};
  if (!dealId) {
    return jsonResponse(400, { success: false, message: 'Debes indicar un número de presupuesto.' });
  }

  const baseUrl = process.env.PIPEDRIVE_API_URL;
  const apiToken = process.env.PIPEDRIVE_API_TOKEN;

  if (!baseUrl || !apiToken) {
    return jsonResponse(500, {
      success: false,
      message: 'Configuración incompleta. Revisa las variables de entorno de Pipedrive.'
    });
  }

  try {
    const dealResponse = await pipedriveRequest(baseUrl, apiToken, `/deals/${encodeURIComponent(dealId)}`);
    const deal = dealResponse && dealResponse.data ? dealResponse.data : null;

    if (!deal) {
      return jsonResponse(404, {
        success: false,
        message: 'No hemos encontrado el presupuesto solicitado.'
      });
    }

    const { primaryDate: trainingDate, secondaryDate: trainingSecondDate } = normaliseTrainingDates(
      deal[TRAINING_DATE_FIELD]
    );
    const rawTrainingLocation = await resolveDealFieldOptionValue(
      baseUrl,
      apiToken,
      TRAINING_LOCATION_FIELD,
      deal[TRAINING_LOCATION_FIELD]
    );
    const trainingLocation = mapTrainingLocation(rawTrainingLocation);
    const trainingName = await resolveDealFieldOptionValue(
      baseUrl,
      apiToken,
      TRAINING_NAME_FIELD,
      deal[TRAINING_NAME_FIELD]
    );

    const organisationId = extractEntityId(deal.org_id);
    const personId = extractEntityId(deal.person_id);
    const contactPersonId =
      personId !== null && personId !== undefined ? String(personId) : '';

    let clientName = deal.org_name || '';
    let contactName = deal.person_name || '';
    let contactEmail = '';

    if (!clientName && organisationId) {
      const organisationResponse = await pipedriveRequest(baseUrl, apiToken, `/organizations/${organisationId}`);
      clientName = organisationResponse?.data?.name || '';
    }

    if (personId) {
      const personResponse = await pipedriveRequest(baseUrl, apiToken, `/persons/${personId}`);
      const personData = personResponse?.data;
      contactName = contactName || personData?.name || '';
      contactEmail = extractPrimaryEmail(personData);
    }

    const notesResponse = await pipedriveRequest(baseUrl, apiToken, `/deals/${encodeURIComponent(dealId)}/notes`, {
      start: 0,
      limit: 100,
      sort_by: 'add_time',
      sort_order: 'desc'
    });

    const notes = Array.isArray(notesResponse?.data) ? notesResponse.data : [];
    const students = extractStudentsFromNotes(notes);

    return jsonResponse(200, {
      success: true,
      data: {
        trainingDate,
        trainingLocation,
        trainingName,
        clientName,
        contactName,
        contactEmail,
        contactPersonId,
        students,
        secondaryTrainingDate: trainingSecondDate
      }
    });
  } catch (error) {
    console.error('Error recuperando información del deal', error);
    const status = error.status || 500;
    const translatedMessage = translateErrorMessage(status, error.message, error.details);
    return jsonResponse(status, { success: false, message: translatedMessage });
  }
};

function jsonResponse(statusCode, body) {
  return {
    statusCode,
    headers: HEADERS,
    body: JSON.stringify(body)
  };
}

async function pipedriveRequest(baseUrl, apiToken, path, params = {}) {
  const url = buildUrl(baseUrl, path, apiToken, params);

  const response = await fetch(url, {
    headers: {
      Accept: 'application/json'
    }
  });

  let payload = null;
  try {
    payload = await response.json();
  } catch (error) {
    payload = null;
  }

  if (!response.ok || (payload && payload.success === false)) {
    const message = extractErrorMessage(payload) || `Error ${response.status} en la petición a ${path}`;
    const error = new Error(message);
    error.status = response.status;
    error.details = payload;
    throw error;
  }

  return payload;
}

function buildUrl(baseUrl, path, apiToken, params) {
  const normalisedBase = baseUrl.endsWith('/') ? baseUrl : `${baseUrl}/`;
  const normalisedPath = path.startsWith('/') ? path.slice(1) : path;
  const url = new URL(normalisedPath, normalisedBase);
  url.searchParams.set('api_token', apiToken);

  Object.entries(params || {}).forEach(([key, value]) => {
    if (value !== undefined && value !== null && value !== '') {
      url.searchParams.set(key, value);
    }
  });

  return url.toString();
}

function extractEntityId(entityField) {
  if (!entityField) return null;
  if (typeof entityField === 'object') {
    return entityField.value || entityField.id || null;
  }
  return entityField;
}

function extractPrimaryEmail(person) {
  if (!person) return '';
  const emailField = person.email;
  if (!emailField) return '';

  if (Array.isArray(emailField)) {
    const primary = emailField.find((item) => item && (item.primary || item.label === 'work')) || emailField[0];
    if (primary && typeof primary === 'object') {
      return primary.value || primary.email || '';
    }
  }

  if (typeof emailField === 'string') {
    return emailField;
  }

  return '';
}

async function resolveDealFieldOptionValue(baseUrl, apiToken, fieldKey, rawValue) {
  if (rawValue === null || rawValue === undefined) {
    return '';
  }

  if (Array.isArray(rawValue)) {
    const firstValue = rawValue.find((item) => item !== null && item !== undefined);
    if (firstValue !== undefined) {
      return resolveDealFieldOptionValue(baseUrl, apiToken, fieldKey, firstValue);
    }
    return '';
  }

  if (typeof rawValue === 'object') {
    const labelLike = rawValue.label || rawValue.name;
    if (labelLike) {
      return labelLike;
    }

    if (rawValue.value !== undefined || rawValue.id !== undefined) {
      const nestedValue = rawValue.value !== undefined ? rawValue.value : rawValue.id;
      return resolveDealFieldOptionValue(baseUrl, apiToken, fieldKey, nestedValue);
    }

    return '';
  }

  const stringValue = String(rawValue).trim();
  if (!stringValue) {
    return '';
  }

  if (!/^\d+$/.test(stringValue)) {
    return stringValue;
  }

  try {
    const options = await fetchDealFieldOptions(baseUrl, apiToken, fieldKey);
    const match = options.find((option) => {
      const optionIdentifier =
        option.id !== undefined ? option.id : option.value !== undefined ? option.value : option.key;
      return optionIdentifier !== undefined && String(optionIdentifier) === stringValue;
    });

    if (match) {
      return match.label || match.name || match.value || stringValue;
    }
  } catch (error) {
    console.error(`No se ha podido resolver el valor del campo ${fieldKey}`, error);
  }

  return stringValue;
}

async function fetchDealFieldOptions(baseUrl, apiToken, fieldKey) {
  if (fieldOptionsCache.has(fieldKey)) {
    return fieldOptionsCache.get(fieldKey);
  }

  try {
    const allFields = await fetchAllDealFields(baseUrl, apiToken);
    const match = allFields.find((field) => {
      if (!field) return false;
      const fieldId = field.id !== undefined ? String(field.id) : null;
      return field.key === fieldKey || fieldId === String(fieldKey);
    });

    const options = Array.isArray(match?.options) ? match.options : [];
    fieldOptionsCache.set(fieldKey, options);
    return options;
  } catch (error) {
    console.error(`No se han podido recuperar las opciones del campo ${fieldKey}`, error);
    fieldOptionsCache.set(fieldKey, []);
    return [];
  }
}

async function fetchAllDealFields(baseUrl, apiToken) {
  if (cachedDealFields) {
    return cachedDealFields;
  }

  if (dealFieldsPromise) {
    return dealFieldsPromise;
  }

  dealFieldsPromise = (async () => {
    try {
      const response = await pipedriveRequest(baseUrl, apiToken, '/dealFields', {
        start: 0,
        limit: 500
      });

      const fields = Array.isArray(response?.data) ? response.data : [];
      cachedDealFields = fields;

      fields.forEach((field) => {
        const options = Array.isArray(field?.options) ? field.options : [];
        if (field?.key) {
          fieldOptionsCache.set(field.key, options);
        }
        if (field?.id !== undefined) {
          fieldOptionsCache.set(String(field.id), options);
        }
      });

      return fields;
    } catch (error) {
      console.error('No se ha podido recuperar el listado de campos de deal', error);
      cachedDealFields = null;
      throw error;
    } finally {
      dealFieldsPromise = null;
    }
  })();

  return dealFieldsPromise;
}

function normaliseTrainingDates(rawValue) {
  const result = { primaryDate: '', secondaryDate: '' };
  const visited = new Set();

  function addDate(dateString) {
    if (!dateString) {
      return;
    }
    if (result.primaryDate === dateString || result.secondaryDate === dateString) {
      return;
    }
    if (!result.primaryDate) {
      result.primaryDate = dateString;
      return;
    }
    if (!result.secondaryDate) {
      result.secondaryDate = dateString;
    }
  }

  function collect(value) {
    if (result.primaryDate && result.secondaryDate) {
      return;
    }

    if (value === null || value === undefined) {
      return;
    }

    if (typeof value === 'string') {
      const trimmed = value.trim();
      if (!trimmed) {
        return;
      }

      const isoMatches = trimmed.match(/\d{4}-\d{2}-\d{2}/g);
      if (isoMatches && isoMatches.length) {
        isoMatches.forEach(addDate);
        const leftover = trimmed.replace(/\d{4}-\d{2}-\d{2}/g, '').trim();
        if (!leftover) {
          return;
        }
      }

      const europeanMatches = trimmed.match(/\d{2}\/\d{2}\/\d{4}/g);
      if (europeanMatches && europeanMatches.length) {
        europeanMatches.forEach((match) => {
          const iso = convertEuropeanDateToIso(match);
          if (iso) {
            addDate(iso);
          }
        });
        const leftover = trimmed.replace(/\d{2}\/\d{2}\/\d{4}/g, '').trim();
        if (!leftover) {
          return;
        }
      }

      const parsed = parseDateInput(trimmed);
      if (parsed) {
        addDate(parsed);
      }
      return;
    }

    if (value instanceof Date || (typeof value === 'number' && Number.isFinite(value))) {
      const parsed = parseDateInput(value);
      if (parsed) {
        addDate(parsed);
      }
      return;
    }

    if (Array.isArray(value)) {
      value.forEach(collect);
      return;
    }

    if (typeof value === 'object') {
      if (visited.has(value)) {
        return;
      }
      visited.add(value);

      const primaryKeys = ['start_date', 'startDate', 'start', 'from', 'initial', 'first', 'primary', 'value', 'date'];
      const secondaryKeys = ['end_date', 'endDate', 'end', 'to', 'final', 'second', 'finish'];

      primaryKeys.forEach((key) => {
        if (Object.prototype.hasOwnProperty.call(value, key)) {
          collect(value[key]);
        }
      });

      secondaryKeys.forEach((key) => {
        if (Object.prototype.hasOwnProperty.call(value, key)) {
          collect(value[key]);
        }
      });

      Object.values(value).forEach(collect);
    }
  }

  collect(rawValue);

  return result;
}

function parseDateInput(value) {
  if (value instanceof Date) {
    if (Number.isNaN(value.getTime())) {
      return '';
    }
    return value.toISOString().split('T')[0];
  }

  if (typeof value === 'number' && Number.isFinite(value)) {
    return formatDateFromTimestamp(value);
  }

  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (!trimmed) {
      return '';
    }

    if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) {
      return trimmed;
    }

    if (/^\d{2}\/\d{2}\/\d{4}$/.test(trimmed)) {
      return convertEuropeanDateToIso(trimmed);
    }

    const numericValue = Number(trimmed);
    if (!Number.isNaN(numericValue) && Number.isFinite(numericValue)) {
      const numericDate = formatDateFromTimestamp(numericValue);
      if (numericDate) {
        return numericDate;
      }
    }

    const parsed = new Date(trimmed);
    if (Number.isNaN(parsed.getTime())) {
      return '';
    }
    return parsed.toISOString().split('T')[0];
  }

  return '';
}

function formatDateFromTimestamp(value) {
  if (!Number.isFinite(value)) {
    return '';
  }

  const milliseconds = Math.abs(value) < 1e12 ? value * 1000 : value;
  const parsed = new Date(milliseconds);
  if (Number.isNaN(parsed.getTime())) {
    return '';
  }
  return parsed.toISOString().split('T')[0];
}

function convertEuropeanDateToIso(value) {
  const match = value.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!match) {
    return '';
  }
  const [, day, month, year] = match;
  return `${year}-${month}-${day}`;
}

function mapTrainingLocation(rawLocation) {
  if (!rawLocation) return '';
  const normalised = String(rawLocation).trim().toLowerCase();
  if (normalised === 'c/ primavera, 1, 28500, arganda del rey, madrid') {
    return 'Madrid';
  }
  if (normalised === 'c/ moratín, 100, 08206 sabadell, barcelona') {
    return 'Barcelona';
  }
  return String(rawLocation);
}

function extractStudentsFromNotes(notes) {
  if (!Array.isArray(notes) || !notes.length) {
    return [];
  }

  const sortedNotes = [...notes]
    .map((note) => ({
      added: note.add_time ? new Date(note.add_time) : new Date(0),
      content: sanitiseNoteContent(note.content)
    }))
    .filter((note) => note.content.toLowerCase().includes('alumnos del deal'))
    .sort((a, b) => b.added.getTime() - a.added.getTime());

  if (!sortedNotes.length) {
    return [];
  }

  const referenceNote = sortedNotes[0].content;
  const [, afterKeyword = ''] = referenceNote.split(/"?alumnos del deal"?/i);
  const cleaned = afterKeyword.replace(/^[:\-\s]+/, '').trim();
  if (!cleaned) {
    return [];
  }

  return cleaned
    .split(';')
    .map((entry) => entry.trim())
    .filter(Boolean)
    .map((entry) => {
      const rawParts = entry.split('|').map((part) => part.trim());
      const hasContent = rawParts.some((part) => part);

      if (!hasContent) {
        return null;
      }

      const [name = '', surname = '', document = ''] = rawParts;

      return {
        name: normalisePersonNameSegment(name),
        surname: normalisePersonNameSegment(surname),
        document,
        documentType: detectDocumentType(document)
      };
    })
    .filter(Boolean);
}

function normalisePersonNameSegment(value) {
  if (!value) return '';

  const repaired = repairEncodingArtifacts(String(value));
  const cleaned = repaired
    .replace(/[\u2018\u2019\u0060]/g, "'")
    .replace(/[\u201C\u201D]/g, '"')
    .replace(/[^\p{L}\p{M}\s'\-]/gu, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  if (!cleaned) {
    return '';
  }

  const corrected = applyNameCorrections(cleaned);
  return corrected.toLocaleUpperCase('es-ES');
}

function applyNameCorrections(value) {
  return value.replace(/\p{L}+/gu, (match) => {
    const key = match.toLocaleUpperCase('es-ES');
    if (NAME_CORRECTIONS.has(key)) {
      return NAME_CORRECTIONS.get(key);
    }
    return match;
  });
}

function repairEncodingArtifacts(value) {
  let output = value;

  if (/[ÃÂ]/.test(output)) {
    try {
      const decoded = Buffer.from(output, 'latin1').toString('utf8');
      if (!decoded.includes('�')) {
        output = decoded;
      }
    } catch (error) {
      // Ignoramos el error y continuamos con las sustituciones manuales.
    }
  }

  ENCODING_REPLACEMENTS.forEach(({ pattern, replacement }) => {
    output = output.replace(pattern, replacement);
  });

  return output;
}

const NAME_CORRECTIONS = new Map([
  ['JOSE', 'JOSÉ'],
  ['MARIA', 'MARÍA'],
  ['ANDRES', 'ANDRÉS'],
  ['ANGEL', 'ÁNGEL'],
  ['ANGELA', 'ÁNGELA'],
  ['JESUS', 'JESÚS'],
  ['RAUL', 'RAÚL'],
  ['JULIAN', 'JULIÁN'],
  ['MARTIN', 'MARTÍN'],
  ['ADRIAN', 'ADRIÁN'],
  ['RUBEN', 'RUBÉN'],
  ['AARON', 'AARÓN'],
  ['ALVARO', 'ÁLVARO'],
  ['IVAN', 'IVÁN'],
  ['OSCAR', 'ÓSCAR'],
  ['VICTOR', 'VÍCTOR'],
  ['ROCIO', 'ROCÍO'],
  ['MONICA', 'MÓNICA'],
  ['VERONICA', 'VERÓNICA'],
  ['SOFIA', 'SOFÍA'],
  ['LUCIA', 'LUCÍA'],
  ['ANAIS', 'ANAÏS'],
  ['INIGO', 'ÍÑIGO'],
  ['NOEMI', 'NOEMÍ'],
  ['NEREA', 'NEREA'],
  ['PAULA', 'PAULA'],
  ['MUNOZ', 'MUÑOZ'],
  ['NUNEZ', 'NUÑEZ'],
  ['PENA', 'PEÑA'],
  ['CANETE', 'CAÑETE'],
  ['CANEDO', 'CAÑEDO'],
  ['PINEIRO', 'PIÑEIRO'],
  ['PINERO', 'PIÑERO'],
  ['SANCHEZ', 'SÁNCHEZ'],
  ['GONZALEZ', 'GONZÁLEZ'],
  ['MARTINEZ', 'MARTÍNEZ'],
  ['HERNANDEZ', 'HERNÁNDEZ'],
  ['ALVAREZ', 'ÁLVAREZ'],
  ['FERNANDEZ', 'FERNÁNDEZ'],
  ['JIMENEZ', 'JIMÉNEZ'],
  ['GOMEZ', 'GÓMEZ'],
  ['LOPEZ', 'LÓPEZ'],
  ['RODRIGUEZ', 'RODRÍGUEZ'],
  ['RAMIREZ', 'RAMÍREZ'],
  ['BENITEZ', 'BENÍTEZ'],
  ['VELAZQUEZ', 'VELÁZQUEZ'],
  ['VALDES', 'VALDÉS'],
  ['SUAREZ', 'SUÁREZ'],
  ['PEREZ', 'PÉREZ'],
  ['DOMINGUEZ', 'DOMÍNGUEZ'],
  ['MENDEZ', 'MÉNDEZ'],
  ['GUTIERREZ', 'GUTIÉRREZ'],
  ['VASQUEZ', 'VÁSQUEZ'],
  ['ESPANA', 'ESPAÑA'],
  ['CASTANO', 'CASTAÑO'],
  ['MALAGA', 'MÁLAGA'],
  ['CORDOBA', 'CÓRDOBA'],
  ['LEON', 'LEÓN'],
  ['AVILA', 'ÁVILA']
]);

const ENCODING_REPLACEMENTS = [
  { pattern: /\u00C3\u00A1/g, replacement: 'á' },
  { pattern: /\u00C3\u0081/g, replacement: 'Á' },
  { pattern: /\u00C3\u00A9/g, replacement: 'é' },
  { pattern: /\u00C3\u0089/g, replacement: 'É' },
  { pattern: /\u00C3\u00AD/g, replacement: 'í' },
  { pattern: /\u00C3\u008D/g, replacement: 'Í' },
  { pattern: /\u00C3\u00B3/g, replacement: 'ó' },
  { pattern: /\u00C3\u0093/g, replacement: 'Ó' },
  { pattern: /\u00C3\u00BA/g, replacement: 'ú' },
  { pattern: /\u00C3\u009A/g, replacement: 'Ú' },
  { pattern: /\u00C3\u00BC/g, replacement: 'ü' },
  { pattern: /\u00C3\u009C/g, replacement: 'Ü' },
  { pattern: /\u00C3\u00B1/g, replacement: 'ñ' },
  { pattern: /\u00C3\u0091/g, replacement: 'Ñ' },
  { pattern: /\u00C2\u00B4/g, replacement: '' },
  { pattern: /\u00C2\u00A8/g, replacement: '' },
  { pattern: /\u00C2\u00BA/g, replacement: 'º' },
  { pattern: /\u00C2\u00AA/g, replacement: 'ª' }
];

function sanitiseNoteContent(content) {
  if (!content) return '';
  const withoutBreaks = String(content)
    .replace(/<\s*br\s*\/?\s*>/gi, '\n')
    .replace(/<\/?p>/gi, '\n')
    .replace(/<li[^>]*>/gi, '\n- ')
    .replace(/<\/?ul>/gi, '\n');
  const withoutHtml = withoutBreaks.replace(/<[^>]*>/g, ' ');
  const decoded = decodeHtmlEntities(withoutHtml);
  return decoded.replace(/\s+/g, ' ').trim();
}

function decodeHtmlEntities(text) {
  return text
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'")
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>');
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

function extractErrorMessage(payload) {
  if (!payload) return '';
  if (typeof payload.error === 'string') return payload.error;
  if (typeof payload.error_info === 'string') return payload.error_info;
  if (typeof payload.error_message === 'string') return payload.error_message;
  if (payload.data && typeof payload.data === 'string') return payload.data;
  return '';
}

function translateErrorMessage(status, originalMessage = '', details) {
  if (status === 404) {
    return 'No hemos encontrado el presupuesto solicitado.';
  }
  if (status === 401) {
    return 'Acceso no autorizado. Revisa la clave de la API de Pipedrive.';
  }
  if (status === 403) {
    return 'Acceso denegado. Comprueba los permisos del token de Pipedrive.';
  }
  if (status === 429) {
    return 'Hemos alcanzado el límite de peticiones de Pipedrive. Inténtalo de nuevo en unos minutos.';
  }

  const message = (originalMessage || '').toLowerCase();
  if (message.includes('not found')) {
    return 'No hemos encontrado el presupuesto solicitado.';
  }
  if (message.includes('token')) {
    return 'Hay un problema con el token de Pipedrive. Revisa su configuración.';
  }
  if (message.includes('rate limit')) {
    return 'Hemos alcanzado el límite de peticiones de Pipedrive. Inténtalo de nuevo en unos minutos.';
  }

  if (details && typeof details === 'object' && typeof details.error_info === 'string') {
    return `Pipedrive respondió: ${details.error_info}`;
  }

  return originalMessage
    ? `No se ha podido completar la operación. Detalle: ${originalMessage}`
    : 'No se ha podido completar la operación con Pipedrive.';
}
