(function (global) {
  const GMAIL_API_BASE = 'https://gmail.googleapis.com/gmail/v1';
  const TOKEN_EXPIRY_SAFETY_MARGIN = 60 * 1000; // 60 seconds
  const GOOGLE_IDENTITY_SCRIPT_SRC = 'https://accounts.google.com/gsi/client';
  const GMAIL_SCOPE = 'https://www.googleapis.com/auth/gmail.send';

  const STORAGE_PREFIX = 'certificados.gmail';
  const STORAGE_KEYS = {
    accessToken: (clientId) => `${STORAGE_PREFIX}.${clientId}.accessToken`,
    tokenExpiresAt: (clientId) => `${STORAGE_PREFIX}.${clientId}.tokenExpiresAt`
  };

  const config = {
    clientId: '',
    scope: GMAIL_SCOPE
  };

  const state = {
    identityScriptPromise: null,
    tokenClient: null,
    accessToken: null,
    tokenExpiresAt: 0
  };

  function getLocalStorage() {
    try {
      return global.localStorage || null;
    } catch (error) {
      return null;
    }
  }

  function clearTokenFromStorage(clientId = config.clientId) {
    if (!clientId) {
      return;
    }

    const storage = getLocalStorage();
    if (!storage) {
      return;
    }

    try {
      storage.removeItem(STORAGE_KEYS.accessToken(clientId));
      storage.removeItem(STORAGE_KEYS.tokenExpiresAt(clientId));
    } catch (error) {
      // Ignoramos los errores de almacenamiento para no bloquear el flujo principal.
    }
  }

  function persistTokenInStorage(clientId, token, expiresAt) {
    if (!clientId || !token) {
      return;
    }

    const storage = getLocalStorage();
    if (!storage) {
      return;
    }

    try {
      storage.setItem(STORAGE_KEYS.accessToken(clientId), token);
      storage.setItem(STORAGE_KEYS.tokenExpiresAt(clientId), String(expiresAt || 0));
    } catch (error) {
      // Si no podemos persistir el token seguimos funcionando con el estado en memoria.
    }
  }

  function restoreTokenFromStorageIfNeeded() {
    if (!config.clientId) {
      return;
    }

    if (state.accessToken && state.tokenExpiresAt && hasValidToken()) {
      return;
    }

    const storage = getLocalStorage();
    if (!storage) {
      return;
    }

    let storedAccessToken = null;
    let storedExpiresAt = 0;

    try {
      storedAccessToken = storage.getItem(STORAGE_KEYS.accessToken(config.clientId));
      const storedExpiresAtValue = storage.getItem(
        STORAGE_KEYS.tokenExpiresAt(config.clientId)
      );
      storedExpiresAt = Number(storedExpiresAtValue);
    } catch (error) {
      storedAccessToken = null;
      storedExpiresAt = 0;
    }

    if (!storedAccessToken || !Number.isFinite(storedExpiresAt)) {
      clearTokenFromStorage(config.clientId);
      return;
    }

    state.accessToken = storedAccessToken;
    state.tokenExpiresAt = storedExpiresAt;
  }

  function configure(options) {
    if (!options || typeof options !== 'object') {
      return;
    }

    if (typeof options.clientId === 'string') {
      const nextClientId = options.clientId.trim();
      if (nextClientId !== config.clientId) {
        if (config.clientId) {
          clearTokenFromStorage(config.clientId);
        }
        config.clientId = nextClientId;
        state.accessToken = null;
        state.tokenExpiresAt = 0;
      } else {
        config.clientId = nextClientId;
      }
    }

    if (typeof options.scope === 'string' && options.scope.trim()) {
      config.scope = options.scope.trim();
    }

    restoreTokenFromStorageIfNeeded();
  }

  function isConfigured() {
    return Boolean(config.clientId);
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
      const existingScript = document.querySelector(`script[src="${GOOGLE_IDENTITY_SCRIPT_SRC}"]`);
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

  function clearStoredToken() {
    state.accessToken = null;
    state.tokenExpiresAt = 0;
    clearTokenFromStorage();
  }

  function normaliseExpiresIn(expiresIn) {
    const expiresInNumber = Number(expiresIn);
    if (!Number.isFinite(expiresInNumber) || expiresInNumber <= 0) {
      return 0;
    }
    return expiresInNumber * 1000;
  }

  const INTERACTION_REQUIRED_ERRORS = new Set([
    'consent_required',
    'interaction_required',
    'login_required'
  ]);

  function requestAccessTokenWithPrompt(prompt) {
    return new Promise((resolve, reject) => {
      state.tokenClient.callback = (response) => {
        if (!response) {
          clearStoredToken();
          reject(new Error('No se ha recibido respuesta del servicio de autenticación de Google.'));
          return;
        }

        if (response.error) {
          const error = new Error(
            response.error_description || response.error || 'No se ha podido obtener autorización para enviar correos con Gmail.'
          );
          error.code = response.error;
          clearStoredToken();
          reject(error);
          return;
        }

        const { access_token: accessToken, expires_in: expiresIn } = response;
        if (!accessToken) {
          clearStoredToken();
          reject(new Error('No se ha recibido el token de acceso de Google.'));
          return;
        }

        state.accessToken = accessToken;
        const expiresInMs = normaliseExpiresIn(expiresIn);
        state.tokenExpiresAt = Date.now() + Math.max(0, expiresInMs);
        persistTokenInStorage(config.clientId, state.accessToken, state.tokenExpiresAt);
        resolve(accessToken);
      };

      try {
        state.tokenClient.requestAccessToken({
          prompt,
          include_granted_scopes: true
        });
      } catch (error) {
        clearStoredToken();
        reject(error);
      }
    });
  }

  async function ensureAccessToken(options = {}) {
    if (!config.clientId) {
      throw new Error('Falta configurar el identificador de cliente de Google (clientId) para utilizar Gmail.');
    }

    await ensureIdentityServicesLoaded();
    restoreTokenFromStorageIfNeeded();

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

    const prompt = options.force || !state.accessToken ? 'consent' : 'none';

    try {
      return await requestAccessTokenWithPrompt(prompt);
    } catch (error) {
      if (!options.force && error && INTERACTION_REQUIRED_ERRORS.has(error.code)) {
        return ensureAccessToken({ force: 'consent' });
      }
      throw error;
    }
  }

  function encodeBase64Url(input) {
    const encoder = new TextEncoder();
    const bytes = encoder.encode(input);
    let binary = '';
    bytes.forEach((byte) => {
      binary += String.fromCharCode(byte);
    });
    const base64 = global.btoa(binary);
    return base64.replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/g, '');
  }

  function encodeMimeWord(value) {
    if (value === undefined || value === null) {
      return '';
    }

    const stringValue = String(value);
    if (!stringValue) {
      return '';
    }

    if (/^[\x00-\x7F]*$/.test(stringValue)) {
      return stringValue;
    }

    const encoder = new TextEncoder();
    const bytes = encoder.encode(stringValue);
    let binary = '';
    bytes.forEach((byte) => {
      binary += String.fromCharCode(byte);
    });
    const base64 = global.btoa(binary);
    return `=?UTF-8?B?${base64}?=`;
  }

  function buildEmailPayload({ to, cc, bcc, subject, body, from }) {
    const headers = [];
    if (from) {
      headers.push(`From: ${from}`);
    }
    if (to) {
      headers.push(`To: ${to}`);
    }
    if (cc) {
      headers.push(`Cc: ${cc}`);
    }
    if (bcc) {
      headers.push(`Bcc: ${bcc}`);
    }

    const finalSubject = subject && String(subject).trim() ? String(subject).trim() : 'Certificado de formación';
    headers.push(`Subject: ${encodeMimeWord(finalSubject)}`);
    headers.push('MIME-Version: 1.0');
    headers.push('Content-Type: text/plain; charset="UTF-8"');
    headers.push('Content-Transfer-Encoding: 7bit');

    const normalisedBody = (body || '').replace(/\r?\n/g, '\r\n');
    const message = `${headers.join('\r\n')}\r\n\r\n${normalisedBody}`;
    return encodeBase64Url(message);
  }

  async function sendEmail({ to, cc, bcc, subject, body, from }) {
    if (!to || !String(to).trim()) {
      throw new Error('No se ha especificado el destinatario del correo.');
    }

    const rawMessage = buildEmailPayload({
      to: String(to).trim(),
      cc: cc && String(cc).trim(),
      bcc: bcc && String(bcc).trim(),
      subject,
      body,
      from
    });

    const accessToken = await ensureAccessToken();

    const response = await fetch(`${GMAIL_API_BASE}/users/me/messages/send`, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ raw: rawMessage })
    });

    if (!response.ok) {
      let errorPayload = null;
      try {
        errorPayload = await response.json();
      } catch (error) {
        errorPayload = null;
      }

      const message =
        errorPayload?.error?.message || `Error ${response.status} al enviar el correo con Gmail.`;
      const error = new Error(message);
      error.status = response.status;
      error.details = errorPayload;
      throw error;
    }

    try {
      return await response.json();
    } catch (error) {
      return null;
    }
  }

  function resetToken() {
    clearStoredToken();
  }

  const initialConfig = global.GOOGLE_GMAIL_CONFIG;
  if (initialConfig && typeof initialConfig === 'object') {
    configure(initialConfig);
  }

  global.googleGmail = {
    configure,
    isConfigured,
    ensureAccessToken,
    hasValidToken,
    sendEmail,
    resetToken
  };
})(window);
