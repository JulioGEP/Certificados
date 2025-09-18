const TRAINING_DATE_FIELD = '98f072a788090ac2ae52017daaf9618c3a189033';
const TRAINING_LOCATION_FIELD = '676d6bd51e52999c582c01f67c99a35ed30bf6ae';
const TRAINING_NAME_FIELD = 'c99554c188c3f63ad9bc8b2cf7b50cbd145455ab';

const HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'Content-Type',
  'Access-Control-Allow-Methods': 'GET,OPTIONS'
};

const fieldOptionsCache = new Map();

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

    const trainingDate = deal[TRAINING_DATE_FIELD] || '';
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
        students
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
    const fieldResponse = await pipedriveRequest(baseUrl, apiToken, `/dealFields/${fieldKey}`);
    const options = Array.isArray(fieldResponse?.data?.options) ? fieldResponse.data.options : [];
    fieldOptionsCache.set(fieldKey, options);
    return options;
  } catch (error) {
    console.error(`No se han podido recuperar las opciones del campo ${fieldKey}`, error);
    fieldOptionsCache.set(fieldKey, []);
    return [];
  }
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
        name,
        surname,
        document,
        documentType: detectDocumentType(document)
      };
    })
    .filter(Boolean);
}

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
