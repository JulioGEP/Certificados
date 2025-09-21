const HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'Content-Type',
  'Access-Control-Allow-Methods': 'POST,OPTIONS'
};

const MAX_FILE_SIZE = 10 * 1024 * 1024; // 10 MB
const SUPPORTED_MIME_TYPES = new Set([
  'application/pdf',
  'application/vnd.ms-excel',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  'application/vnd.oasis.opendocument.spreadsheet',
  'text/csv',
  'image/jpeg',
  'image/png',
  'image/webp',
  'image/gif',
  'image/bmp',
  'image/tiff',
  'image/heic',
  'image/heif'
]);

const EXTENSION_MIME_MAP = {
  pdf: 'application/pdf',
  xls: 'application/vnd.ms-excel',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  csv: 'text/csv',
  ods: 'application/vnd.oasis.opendocument.spreadsheet',
  jpeg: 'image/jpeg',
  jpg: 'image/jpeg',
  png: 'image/png',
  webp: 'image/webp',
  gif: 'image/gif',
  bmp: 'image/bmp',
  tif: 'image/tiff',
  tiff: 'image/tiff',
  heic: 'image/heic',
  heif: 'image/heif'
};

exports.handler = async (event) => {
  if (event.httpMethod === 'OPTIONS') {
    return {
      statusCode: 204,
      headers: HEADERS,
      body: ''
    };
  }

  if (event.httpMethod !== 'POST') {
    return jsonResponse(405, { success: false, message: 'Método no permitido.' });
  }

  let payload = null;
  try {
    payload = JSON.parse(event.body || '{}');
  } catch (error) {
    return jsonResponse(400, { success: false, message: 'El cuerpo de la solicitud no es válido.' });
  }

  const budgetId = normaliseText(payload?.budgetId);
  const fileName = normaliseFileName(payload?.fileName);
  const fileType = normaliseText(payload?.fileType).toLowerCase();
  const fileSize = Number(payload?.fileSize) || 0;
  const fileData = typeof payload?.fileData === 'string' ? payload.fileData : '';

  if (!budgetId) {
    return jsonResponse(400, { success: false, message: 'Debes indicar el número de presupuesto.' });
  }

  if (!fileName || !fileData) {
    return jsonResponse(400, {
      success: false,
      message: 'Debes adjuntar un archivo en formato PDF, Excel o imagen.'
    });
  }

  const buffer = decodeBase64File(fileData);
  if (!buffer || !buffer.length) {
    return jsonResponse(400, { success: false, message: 'El archivo recibido no es válido.' });
  }

  if (buffer.length > MAX_FILE_SIZE || fileSize > MAX_FILE_SIZE) {
    return jsonResponse(413, {
      success: false,
      message: 'El archivo supera el tamaño máximo permitido (10 MB).'
    });
  }

  const { valid, mimeType, message: validationMessage } = resolveFileMimeType(fileName, fileType);
  if (!valid) {
    return jsonResponse(400, {
      success: false,
      message: validationMessage || 'El tipo de archivo no es compatible. Usa un PDF, Excel o imagen.'
    });
  }

  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) {
    return jsonResponse(500, {
      success: false,
      message: 'Falta configurar la variable de entorno OPENAI_API_KEY en Netlify.'
    });
  }

  const authorisationHeader = `Bearer ${apiKey}`;
  let uploadedFileId = '';

  try {
    uploadedFileId = await uploadFileToOpenAi(buffer, fileName, mimeType, authorisationHeader);
    const responsePayload = await requestOpenAiExtraction(uploadedFileId, authorisationHeader);
    const outputText = extractResponseText(responsePayload);
    const parsed = parseResponsePayload(outputText);
    const students = normaliseStudents(parsed.students);
    const warnings = normaliseWarnings(parsed.warnings);

    return jsonResponse(200, {
      success: true,
      data: {
        students,
        warnings
      }
    });
  } catch (error) {
    console.error('Error procesando el documento con OpenAI', error);
    const status = error?.status && Number.isInteger(error.status) ? error.status : 500;
    const message = translateOpenAiError(error);
    return jsonResponse(status, { success: false, message });
  } finally {
    if (uploadedFileId) {
      try {
        await fetch(`https://api.openai.com/v1/files/${encodeURIComponent(uploadedFileId)}`, {
          method: 'DELETE',
          headers: {
            Authorization: authorisationHeader
          }
        });
      } catch (cleanupError) {
        console.warn('No se ha podido eliminar el archivo temporal de OpenAI', cleanupError);
      }
    }
  }
};

async function uploadFileToOpenAi(buffer, fileName, mimeType, authorisationHeader) {
  const formData = new FormData();
  formData.append('purpose', 'assistants');
  formData.append('file', new File([buffer], fileName, { type: mimeType || 'application/octet-stream' }));

  const response = await fetch('https://api.openai.com/v1/files', {
    method: 'POST',
    headers: {
      Authorization: authorisationHeader
    },
    body: formData
  });

  const payload = await safeJson(response);

  if (!response.ok || !payload?.id) {
    const error = buildOpenAiError(response.status, payload);
    throw error;
  }

  return payload.id;
}

async function requestOpenAiExtraction(fileId, authorisationHeader) {
  const body = {
    model: 'gpt-4.1-mini',
    temperature: 0,
    input: [
      {
        role: 'system',
        content: [
          {
            type: 'text',
            text:
              'Eres un asistente que ayuda a un equipo de formación a extraer listados de alumnado. '
              + 'Cuando recibas un documento deberás identificar a todos los alumnos y alumnas, '
              + 'separar el nombre de los apellidos y recuperar su DNI o NIE si está disponible. '
              + 'Si algún dato no aparece, devuelve el campo vacío.'
          }
        ]
      },
      {
        role: 'user',
        content: [
          {
            type: 'input_text',
            text:
              'Analiza el archivo adjunto y devuelve un JSON con el listado de alumnos. '
              + 'Cada elemento debe incluir "nombre", "apellido" y "dni". '
              + 'Si detectas dudas o incidencias importantes añádelas en un campo "warnings".'
          },
          {
            type: 'input_file',
            file_id: fileId
          }
        ]
      }
    ],
    response_format: {
      type: 'json_schema',
      json_schema: {
        name: 'students_payload',
        schema: {
          type: 'object',
          additionalProperties: false,
          properties: {
            students: {
              type: 'array',
              items: {
                type: 'object',
                additionalProperties: false,
                properties: {
                  nombre: { type: 'string', description: 'Nombre de pila del alumno o alumna.' },
                  apellido: { type: 'string', description: 'Apellidos del alumno o alumna.' },
                  dni: { type: 'string', description: 'Documento identificativo (DNI, NIE, etc.).' }
                },
                required: ['nombre', 'apellido', 'dni']
              }
            },
            warnings: {
              type: 'array',
              items: { type: 'string' }
            }
          },
          required: ['students']
        },
        strict: true
      }
    }
  };

  const response = await fetch('https://api.openai.com/v1/responses', {
    method: 'POST',
    headers: {
      Authorization: authorisationHeader,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(body)
  });

  const payload = await safeJson(response);

  if (!response.ok) {
    const error = buildOpenAiError(response.status, payload);
    throw error;
  }

  return payload;
}

function jsonResponse(statusCode, body) {
  return {
    statusCode,
    headers: HEADERS,
    body: JSON.stringify(body)
  };
}

function decodeBase64File(data) {
  if (!data) {
    return null;
  }

  const normalised = data.includes('base64,') ? data.slice(data.indexOf('base64,') + 7) : data;

  try {
    return Buffer.from(normalised, 'base64');
  } catch (error) {
    return null;
  }
}

function resolveFileMimeType(fileName, providedMime) {
  const mime = providedMime && SUPPORTED_MIME_TYPES.has(providedMime) ? providedMime : '';
  if (mime) {
    return { valid: true, mimeType: mime };
  }

  const extension = extractExtension(fileName);
  if (extension && EXTENSION_MIME_MAP[extension]) {
    return {
      valid: true,
      mimeType: EXTENSION_MIME_MAP[extension]
    };
  }

  return {
    valid: false,
    mimeType: '',
    message: 'El tipo de archivo no es compatible. Usa un PDF, Excel o imagen.'
  };
}

function extractExtension(fileName = '') {
  const normalised = normaliseText(fileName).toLowerCase();
  const lastDot = normalised.lastIndexOf('.');
  if (lastDot === -1 || lastDot === normalised.length - 1) {
    return '';
  }
  return normalised.slice(lastDot + 1);
}

function parseResponsePayload(text) {
  if (!text) {
    return { students: [], warnings: [] };
  }

  try {
    const parsed = JSON.parse(text);
    if (parsed && typeof parsed === 'object') {
      return parsed;
    }
  } catch (error) {
    console.warn('No se ha podido interpretar la respuesta de OpenAI como JSON.', error);
  }

  return { students: [], warnings: [] };
}

function extractResponseText(response) {
  if (!response) {
    return '';
  }

  if (typeof response.output_text === 'string') {
    return response.output_text;
  }

  const outputItems = Array.isArray(response.output) ? response.output : [];
  for (const item of outputItems) {
    const contentBlocks = Array.isArray(item?.content) ? item.content : [];
    for (const block of contentBlocks) {
      if (typeof block?.text === 'string') {
        return block.text;
      }
      if (typeof block?.output_text === 'string') {
        return block.output_text;
      }
    }
  }

  return '';
}

function normaliseStudents(students) {
  if (!Array.isArray(students)) {
    return [];
  }

  return students
    .map((student) => ({
      nombre: normaliseText(student?.nombre),
      apellido: normaliseText(student?.apellido),
      dni: normaliseDocument(student?.dni)
    }))
    .filter((student) => student.nombre || student.apellido || student.dni);
}

function normaliseWarnings(warnings) {
  if (!Array.isArray(warnings)) {
    return [];
  }

  return warnings
    .map((warning) => normaliseText(warning))
    .filter((warning) => warning.length > 0);
}

function translateOpenAiError(error) {
  if (!error) {
    return 'No se ha podido procesar el documento con OpenAI. Inténtalo de nuevo.';
  }

  if (error.status === 401) {
    return 'OpenAI ha rechazado la solicitud. Revisa la configuración de la API.';
  }

  if (error.status === 429) {
    return 'Se ha alcanzado el límite de peticiones a OpenAI. Inténtalo de nuevo en unos minutos.';
  }

  if (error.status === 413) {
    return 'El archivo es demasiado grande para procesarlo. Reduce su tamaño e inténtalo de nuevo.';
  }

  if (error?.message) {
    return normaliseText(error.message) || 'No se ha podido procesar el documento en OpenAI.';
  }

  if (error instanceof Error && error.message) {
    return error.message;
  }

  return 'No se ha podido procesar el documento en OpenAI.';
}

function buildOpenAiError(status, payload) {
  const error = new Error(payload?.error?.message || payload?.message || 'Error al comunicar con OpenAI.');
  error.status = status || 500;
  return error;
}

async function safeJson(response) {
  try {
    return await response.json();
  } catch (error) {
    return null;
  }
}

function normaliseText(value) {
  if (value === undefined || value === null) {
    return '';
  }

  return String(value)
    .replace(/\s+/g, ' ')
    .trim();
}

function normaliseDocument(value) {
  const text = normaliseText(value);
  if (!text) {
    return '';
  }

  return text.replace(/\s+/g, '').toUpperCase();
}

function normaliseFileName(value) {
  const text = normaliseText(value);
  if (!text) {
    return 'documento';
  }

  return text.split(/[/\\]/).pop();
}
