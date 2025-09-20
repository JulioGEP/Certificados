(function (global) {
  const DEFAULT_ROOT_FOLDER_ID = '15FclFHgqFha76Y51OnzJxypxxvgI-Pr6';
  const FOLDER_MIME_TYPE = 'application/vnd.google-apps.folder';
  const PDF_MIME_TYPE = 'application/pdf';
  const DRIVE_API_BASE = 'https://www.googleapis.com/drive/v3';
  const DRIVE_UPLOAD_BASE = 'https://www.googleapis.com/upload/drive/v3';
  const TOKEN_EXPIRY_SAFETY_MARGIN = 60 * 1000; // 60 seconds
  const GOOGLE_IDENTITY_SCRIPT_SRC = 'https://accounts.google.com/gsi/client';
  const DRIVE_SCOPES = [
    'https://www.googleapis.com/auth/drive.file',
    'https://www.googleapis.com/auth/drive'
  ].join(' ');

  const config = {
    clientId: '',
    rootFolderId: DEFAULT_ROOT_FOLDER_ID,
    scope: DRIVE_SCOPES
  };

  const state = {
    identityScriptPromise: null,
    tokenClient: null,
    accessToken: null,
    tokenExpiresAt: 0
  };

  function configure(options) {
    if (!options || typeof options !== 'object') {
      return;
    }

    if (typeof options.clientId === 'string') {
      config.clientId = options.clientId.trim();
    }

    if (typeof options.rootFolderId === 'string' && options.rootFolderId.trim()) {
      config.rootFolderId = options.rootFolderId.trim();
    }

    if (typeof options.scope === 'string' && options.scope.trim()) {
      config.scope = options.scope.trim();
    }
  }

  function isConfigured() {
    return Boolean(config.clientId && config.rootFolderId);
  }

  function getRootFolderId() {
    return config.rootFolderId;
  }

  function getClientId() {
    return config.clientId;
  }

  function ensureIdentityServicesLoaded() {
    if (state.identityScriptPromise) {
      return state.identityScriptPromise;
    }

    if (global.google && global.google.accounts && global.google.accounts.oauth2) {
      state.identityScriptPromise = Promise.resolve();
      return state.identityScriptPromise;
    }

    state.identityScriptPromise = new Promise((resolve, reject) => {
      const existingScript = document.querySelector(
        `script[src="${GOOGLE_IDENTITY_SCRIPT_SRC}"]`
      );
      if (existingScript && existingScript.hasAttribute('data-loaded')) {
        resolve();
        return;
      }

      const script = existingScript || document.createElement('script');
      script.src = GOOGLE_IDENTITY_SCRIPT_SRC;
      script.async = true;
      script.defer = true;
      script.onload = () => {
        script.setAttribute('data-loaded', 'true');
        resolve();
      };
      script.onerror = () => {
        reject(new Error('No se ha podido cargar la librería de identidad de Google.'));
      };

      if (!existingScript) {
        document.head.appendChild(script);
      }
    });

    return state.identityScriptPromise;
  }

  function hasValidToken() {
    return Boolean(
      state.accessToken &&
        state.tokenExpiresAt &&
        Date.now() + TOKEN_EXPIRY_SAFETY_MARGIN < state.tokenExpiresAt
    );
  }

  async function ensureAccessToken(options = {}) {
    if (!config.clientId) {
      throw new Error(
        'Falta configurar el identificador de cliente de Google (clientId) para utilizar Google Drive.'
      );
    }

    await ensureIdentityServicesLoaded();

    if (!state.tokenClient) {
      state.tokenClient = global.google.accounts.oauth2.initTokenClient({
        client_id: config.clientId,
        scope: config.scope,
        callback: () => {}
      });
    }

    if (!options.force && hasValidToken()) {
      return state.accessToken;
    }

    return new Promise((resolve, reject) => {
      state.tokenClient.callback = (response) => {
        if (!response) {
          reject(new Error('No se ha recibido respuesta del servicio de autenticación de Google.'));
          return;
        }

        if (response.error) {
          const error = new Error(
            response.error_description ||
              response.error ||
              'No se ha podido obtener autorización para usar Google Drive.'
          );
          error.code = response.error;
          reject(error);
          return;
        }

        const { access_token: accessToken, expires_in: expiresIn } = response;
        if (!accessToken) {
          reject(new Error('No se ha recibido el token de acceso de Google.'));
          return;
        }

        state.accessToken = accessToken;
        const expiresInMs = typeof expiresIn === 'number' ? expiresIn * 1000 : 0;
        state.tokenExpiresAt = Date.now() + Math.max(0, expiresInMs);
        resolve(accessToken);
      };

      try {
        state.tokenClient.requestAccessToken({
          prompt: options.force || !state.accessToken ? 'consent' : ''
        });
      } catch (error) {
        reject(error);
      }
    });
  }

  function buildDriveQuery(params) {
    const parts = [];
    if (params.parentId) {
      parts.push(`'${params.parentId}' in parents`);
    }
    if (params.mimeType) {
      parts.push(`mimeType = '${params.mimeType}'`);
    }
    if (params.name) {
      const escaped = params.name.replace(/'/g, "\\'");
      parts.push(`name = '${escaped}'`);
    }
    parts.push('trashed = false');
    return parts.join(' and ');
  }

  async function driveRequest(path, options = {}) {
    const accessToken = await ensureAccessToken();
    const url = path instanceof URL ? path : new URL(path, DRIVE_API_BASE);
    const headers = new Headers(options.headers || {});
    headers.set('Authorization', `Bearer ${accessToken}`);
    if (!headers.has('Accept')) {
      headers.set('Accept', 'application/json');
    }

    const response = await fetch(url.toString(), { ...options, headers });

    if (!response.ok) {
      let errorPayload = null;
      try {
        errorPayload = await response.json();
      } catch (error) {
        errorPayload = null;
      }

      const message =
        errorPayload?.error?.message ||
        `Error ${response.status} en la petición a Google Drive.`;
      const error = new Error(message);
      error.status = response.status;
      error.details = errorPayload;
      throw error;
    }

    if (response.status === 204) {
      return null;
    }

    try {
      return await response.json();
    } catch (error) {
      return null;
    }
  }

  async function searchItemByName({ name, parentId, mimeType }) {
    const query = buildDriveQuery({ name, parentId, mimeType });
    const params = new URLSearchParams({
      q: query,
      pageSize: '10',
      fields: 'files(id,name,mimeType,parents,webViewLink)',
      includeItemsFromAllDrives: 'true',
      supportsAllDrives: 'true',
      spaces: 'drive'
    });
    const data = await driveRequest(`/files?${params.toString()}`, { method: 'GET' });
    if (!data || !Array.isArray(data.files) || !data.files.length) {
      return null;
    }
    return data.files[0];
  }

  function sanitiseDriveName(value, fallback) {
    if (value === undefined || value === null) {
      return fallback;
    }
    const trimmed = String(value)
      .replace(/[\n\r]/g, ' ')
      .replace(/[\\/:*?"<>|]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
    return trimmed || fallback;
  }

  async function ensureFolder(name, parentId) {
    const safeName = sanitiseDriveName(name, 'Carpeta sin nombre');
    const parent = parentId || config.rootFolderId;
    if (!parent) {
      throw new Error('No se ha configurado la carpeta raíz de Google Drive.');
    }

    const existing = await searchItemByName({
      name: safeName,
      parentId: parent,
      mimeType: FOLDER_MIME_TYPE
    });
    if (existing) {
      return existing;
    }

    const metadata = {
      name: safeName,
      mimeType: FOLDER_MIME_TYPE,
      parents: [parent]
    };

    const params = new URLSearchParams({
      fields: 'id,name,parents',
      supportsAllDrives: 'true'
    });

    const response = await driveRequest(`/files?${params.toString()}`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json; charset=UTF-8'
      },
      body: JSON.stringify(metadata)
    });

    return response;
  }

  function buildMultipartBody(metadata, fileBlob, mimeType) {
    const boundary = `-------gepdrive-${Date.now().toString(16)}`;
    const metaPart = `--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${JSON.stringify(
      metadata
    )}\r\n`;
    const fileHeader = `--${boundary}\r\nContent-Type: ${mimeType}\r\n\r\n`;
    const footer = `\r\n--${boundary}--`;

    const body = new Blob([metaPart, fileHeader, fileBlob, footer], {
      type: `multipart/related; boundary=${boundary}`
    });

    return { body, boundary };
  }

  async function uploadOrUpdateFile({ name, parents, blob, mimeType = PDF_MIME_TYPE }) {
    if (!(blob instanceof Blob)) {
      throw new Error('El archivo a subir a Google Drive no es válido.');
    }

    const safeName = sanitiseDriveName(name, 'certificado.pdf');
    const parentId = Array.isArray(parents) ? parents[0] : parents;

    const existing = await searchItemByName({
      name: safeName,
      parentId,
      mimeType: PDF_MIME_TYPE
    });

    const metadata = {
      name: safeName,
      parents: Array.isArray(parents) ? parents : [parentId]
    };

    const { body, boundary } = buildMultipartBody(metadata, blob, mimeType);

    const params = new URLSearchParams({
      uploadType: 'multipart',
      fields: 'id,name,parents,webViewLink',
      supportsAllDrives: 'true'
    });

    const path = existing
      ? `${DRIVE_UPLOAD_BASE}/files/${existing.id}?${params.toString()}`
      : `${DRIVE_UPLOAD_BASE}/files?${params.toString()}`;

    const method = existing ? 'PATCH' : 'POST';

    const accessToken = await ensureAccessToken();
    const response = await fetch(path, {
      method,
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': `multipart/related; boundary=${boundary}`,
        Accept: 'application/json'
      },
      body
    });

    if (!response.ok) {
      let errorPayload = null;
      try {
        errorPayload = await response.json();
      } catch (error) {
        errorPayload = null;
      }

      const message =
        errorPayload?.error?.message ||
        `Error ${response.status} al ${existing ? 'actualizar' : 'subir'} el archivo en Google Drive.`;
      const error = new Error(message);
      error.status = response.status;
      error.details = errorPayload;
      throw error;
    }

    return response.json();
  }

  async function uploadCertificate({
    clientFolderName,
    trainingFolderName,
    fileName,
    blob,
    mimeType = PDF_MIME_TYPE
  }) {
    if (!config.rootFolderId) {
      throw new Error('No se ha configurado la carpeta raíz de Google Drive.');
    }

    const clientFolder = await ensureFolder(clientFolderName, config.rootFolderId);
    const trainingFolder = await ensureFolder(trainingFolderName, clientFolder.id);

    const file = await uploadOrUpdateFile({
      name: fileName,
      parents: [trainingFolder.id],
      blob,
      mimeType
    });

    return {
      clientFolderId: clientFolder.id,
      trainingFolderId: trainingFolder.id,
      fileId: file.id,
      fileName: file.name,
      webViewLink: file.webViewLink || null
    };
  }

  function resetToken() {
    state.accessToken = null;
    state.tokenExpiresAt = 0;
  }

  const initialConfig = global.GOOGLE_DRIVE_CONFIG;
  if (initialConfig && typeof initialConfig === 'object') {
    configure(initialConfig);
  }

  global.googleDrive = {
    configure,
    isConfigured,
    getRootFolderId,
    getClientId,
    ensureAccessToken,
    hasValidToken,
    uploadCertificate,
    resetToken
  };
})(window);
