(function (global) {
  const TRANSPARENT_PIXEL =
    'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQI12NkYGD4DwABBAEAi5JBSwAAAABJRU5ErkJggg==';

  const ASSET_PATHS = {
    background: 'assets/certificados/fondo-certificado.png',
    leftSidebar: 'assets/certificados/lateral-izquierdo.png',
    footer: 'assets/certificados/pie-firma.png',
    logo: 'assets/certificados/logo-certificado.png'
  };

  const IMAGE_ASPECT_RATIOS = {
    background: 1328 / 839,
    footer: 153 / 853,
    logo: 382 / 827
  };

  const FONT_SOURCES = {
    'Poppins-Regular.ttf':
      'https://cdn.jsdelivr.net/npm/@fontsource/poppins@5.0.17/files/poppins-latin-400-normal.ttf',
    'Poppins-Italic.ttf':
      'https://cdn.jsdelivr.net/npm/@fontsource/poppins@5.0.17/files/poppins-latin-400-italic.ttf',
    'Poppins-SemiBold.ttf':
      'https://cdn.jsdelivr.net/npm/@fontsource/poppins@5.0.17/files/poppins-latin-600-normal.ttf',
    'Poppins-SemiBoldItalic.ttf':
      'https://cdn.jsdelivr.net/npm/@fontsource/poppins@5.0.17/files/poppins-latin-600-italic.ttf'
  };

  let poppinsFontPromise = null;

  const assetCache = new Map();

  function getCachedAsset(key) {
    if (assetCache.has(key)) {
      return assetCache.get(key);
    }
    const promise = loadImageAsDataUrl(ASSET_PATHS[key]).catch((error) => {
      console.warn(`No se ha podido cargar el recurso "${key}" (${ASSET_PATHS[key]}).`, error);
      return TRANSPARENT_PIXEL;
    });
    assetCache.set(key, promise);
    return promise;
  }

  async function loadImageAsDataUrl(path) {
    if (!path) {
      return TRANSPARENT_PIXEL;
    }
    const response = await fetch(path);
    if (!response.ok) {
      throw new Error(`Respuesta ${response.status} al cargar ${path}`);
    }
    const blob = await response.blob();
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onloadend = () => resolve(reader.result);
      reader.onerror = () => reject(new Error(`No se ha podido leer el archivo ${path}`));
      reader.readAsDataURL(blob);
    });
  }

  async function loadFontAsBase64(url) {
    const response = await fetch(url);
    if (!response.ok) {
      throw new Error(`Respuesta ${response.status} al cargar ${url}`);
    }
    const buffer = await response.arrayBuffer();
    return arrayBufferToBase64(buffer);
  }

  function arrayBufferToBase64(buffer) {
    let binary = '';
    const bytes = new Uint8Array(buffer);
    const chunkSize = 0x8000;
    for (let i = 0; i < bytes.length; i += chunkSize) {
      const chunk = bytes.subarray(i, i + chunkSize);
      binary += String.fromCharCode.apply(null, chunk);
    }
    if (typeof global.btoa === 'function') {
      return global.btoa(binary);
    }
    if (typeof Buffer !== 'undefined') {
      return Buffer.from(binary, 'binary').toString('base64');
    }
    throw new Error('No se puede convertir el buffer en base64 en este entorno.');
  }

  async function ensurePoppinsFont() {
    const { pdfMake } = global;
    if (!pdfMake) {
      return;
    }
    if (pdfMake.fonts && pdfMake.fonts.Poppins) {
      return;
    }
    if (!poppinsFontPromise) {
      poppinsFontPromise = (async () => {
        try {
          const fontEntries = await Promise.all(
            Object.entries(FONT_SOURCES).map(async ([name, url]) => {
              try {
                const data = await loadFontAsBase64(url);
                return [name, data];
              } catch (error) {
                console.warn(`No se ha podido cargar la fuente Poppins (${name}).`, error);
                return null;
              }
            })
          );

          const validEntries = fontEntries.filter(Boolean);
          if (validEntries.length) {
            pdfMake.vfs = pdfMake.vfs || {};
            validEntries.forEach(([name, data]) => {
              pdfMake.vfs[name] = data;
            });
          }

          const existingFonts = pdfMake.fonts || {};
          const roboto = existingFonts.Roboto || {};
          const previousPoppins = existingFonts.Poppins || {};
          const availableFontNames = new Set(validEntries.map(([name]) => name));

          pdfMake.fonts = {
            ...existingFonts,
            Poppins: {
              normal: availableFontNames.has('Poppins-Regular.ttf')
                ? 'Poppins-Regular.ttf'
                : previousPoppins.normal || roboto.normal || 'Roboto-Regular.ttf',
              bold: availableFontNames.has('Poppins-SemiBold.ttf')
                ? 'Poppins-SemiBold.ttf'
                : previousPoppins.bold || roboto.bold || 'Roboto-Medium.ttf',
              italics: availableFontNames.has('Poppins-Italic.ttf')
                ? 'Poppins-Italic.ttf'
                : previousPoppins.italics || roboto.italics || 'Roboto-Italic.ttf',
              bolditalics: availableFontNames.has('Poppins-SemiBoldItalic.ttf')
                ? 'Poppins-SemiBoldItalic.ttf'
                : previousPoppins.bolditalics || roboto.bolditalics || 'Roboto-BoldItalic.ttf'
            }
          };
        } catch (error) {
          console.warn('No se ha podido preparar la tipografía Poppins.', error);
        }
      })();
    }

    try {
      await poppinsFontPromise;
    } catch (error) {
      console.warn('No se ha podido cargar la tipografía Poppins para el certificado.', error);
    }
  }

  function normaliseText(value) {
    if (value === undefined || value === null) {
      return '';
    }
    return String(value).trim();
  }

  function buildFullName(row) {
    const name = normaliseText(row.nombre);
    const surname = normaliseText(row.apellido);
    return [name, surname].filter(Boolean).join(' ').trim() || 'Nombre del alumno/a';
  }

  function buildDocumentSentence(row) {
    const documentType = normaliseText(row.documentType).toUpperCase();
    const documentNumber = normaliseText(row.dni);
    if (!documentType && !documentNumber) {
      return 'con documento de identidad';
    }
    if (!documentType) {
      return `con documento ${documentNumber}`;
    }
    if (!documentNumber) {
      return `con ${documentType}`;
    }
    return `con ${documentType} ${documentNumber}`;
  }

  function formatTrainingDate(value) {
    const normalised = normaliseText(value);
    if (!normalised) {
      return '________';
    }
    const parsed = new Date(normalised);
    if (Number.isNaN(parsed.getTime())) {
      return normalised;
    }
    const formatter = new Intl.DateTimeFormat('es-ES', {
      day: 'numeric',
      month: 'long',
      year: 'numeric'
    });
    return formatter.format(parsed);
  }

  function formatLocation(value) {
    return normaliseText(value) || '________';
  }

  function formatDuration(value) {
    if (value === undefined || value === null || value === '') {
      return '____';
    }
    const numberValue = Number(value);
    if (Number.isNaN(numberValue)) {
      return normaliseText(value);
    }
    return numberValue % 1 === 0 ? String(numberValue) : numberValue.toLocaleString('es-ES');
  }

  function formatTrainingName(value) {
    return normaliseText(value) || 'Nombre de la formación';
  }

  function buildFileName(row) {
    const name = buildFullName(row).toLowerCase();
    const safeName = name
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/[^a-z0-9]+/g, '-')
      .replace(/(^-|-$)/g, '');
    const dealId = normaliseText(row.presupuesto);
    const suffix = dealId ? `-${dealId}` : '';
    const base = safeName || 'certificado';
    return `${base}${suffix}.pdf`;
  }

  function buildDocStyles() {
    return {
      bodyText: {
        fontSize: 12,
        lineHeight: 1.4
      },
      introText: {
        fontSize: 12,
        lineHeight: 1.45,
        margin: [0, 0, 0, 10]
      },
      certificateTitle: {
        fontSize: 30,
        bold: true,
        color: '#c4143c',
        letterSpacing: 3,
        margin: [0, 8, 0, 10]
      },
      highlighted: {
        fontSize: 14,
        bold: true,
        margin: [0, 8, 0, 8]
      },
      trainingName: {
        fontSize: 18,
        bold: true,
        margin: [0, 14, 0, 0]
      }
    };
  }

  async function buildDocDefinition(row) {
    await ensurePoppinsFont();
    const [backgroundImage, sidebarImage, footerImage, logoImage] = await Promise.all([
      getCachedAsset('background'),
      getCachedAsset('leftSidebar'),
      getCachedAsset('footer'),
      getCachedAsset('logo')
    ]);

    const pageMargins = [105, 40, 160, 100];
    const fullName = buildFullName(row);
    const documentSentence = buildDocumentSentence(row);
    const trainingDate = formatTrainingDate(row.fecha);
    const location = formatLocation(row.lugar);
    const duration = formatDuration(row.duracion);
    const trainingName = formatTrainingName(row.formacion);

    const contentStack = [
      {
        text: 'Sr. Lluís Vicent Pérez,\nDirector de la escuela GEPCO Formación\nexpide el presente:',
        style: 'introText'
      },
      { text: 'CERTIFICADO', style: 'certificateTitle' },
      { text: `A nombre del alumno/a ${fullName}`, style: 'bodyText' },
      {
        text: `${documentSentence}, quien en fecha ${trainingDate} y en ${location}`,
        style: 'bodyText'
      },
      {
        text: `ha superado, con una duración total de ${duration} horas, la formación de:`,
        style: 'bodyText'
      },
      { text: trainingName, style: 'trainingName' },
      { text: '\n\n', margin: [0, 0, 0, 0] }
    ];

    const docDefinition = {
      pageOrientation: 'landscape',
      pageSize: 'A4',
      pageMargins,
      background: function (currentPage, pageSize) {
        const pageWidth = pageSize.width || 841.89;
        const pageHeight = pageSize.height || 595.28;
        const baseSidebarWidth = Math.min(70, pageWidth * 0.08);
        const sidebarWidth = baseSidebarWidth * 0.85;
        const backgroundWidth = Math.min(320, pageWidth * 0.35);
        const backgroundX = pageWidth - backgroundWidth + backgroundWidth * 0.12;
        const backgroundHeight = backgroundWidth * IMAGE_ASPECT_RATIOS.background;
        const backgroundY = (pageHeight - backgroundHeight) / 2;
        const footerBaseWidth = Math.min(pageWidth - 40, 780);
        const footerMinLeft = Math.max(0, sidebarWidth + 18);
        const footerMaxWidth = Math.max(0, pageWidth - footerMinLeft - 30);
        const footerWidth = Math.min(footerBaseWidth * 0.8, footerMaxWidth);
        const footerHeight = footerWidth * IMAGE_ASPECT_RATIOS.footer;
        const bottomLift = pageMargins[3] * 0.1;
        const footerY = Math.max(0, pageHeight - footerHeight - bottomLift);
        const logoWidth = Math.min(backgroundWidth * 0.6, 200);
        const logoHeight = logoWidth * IMAGE_ASPECT_RATIOS.logo;
        const logoY = Math.max(0, (pageHeight - logoHeight) / 2);
        const logoX = backgroundX + backgroundWidth - logoWidth + Math.min(logoWidth * 0.1, 20);

        return [
          {
            image: sidebarImage,
            width: sidebarWidth,
            height: pageHeight,
            absolutePosition: { x: 0, y: 0 }
          },
          {
            image: backgroundImage,
            width: backgroundWidth,
            absolutePosition: { x: backgroundX, y: backgroundY }
          },
          {
            image: logoImage,
            width: logoWidth,
            absolutePosition: { x: logoX, y: logoY }
          },
          {
            image: footerImage,
            width: footerWidth,
            absolutePosition: { x: footerMinLeft, y: footerY }
          }
        ];
      },
      content: [
        {
          margin: [0, 12, 0, 0],
          stack: contentStack
        }
      ],
      styles: buildDocStyles(),
      defaultStyle: {
        fontSize: 12,
        lineHeight: 1.45,
        color: '#1f274d',
        font: 'Poppins'
      },
      info: {
        title: `Certificado - ${fullName}`,
        author: 'GEPCO Formación',
        subject: trainingName
      }
    };

    return docDefinition;
  }

  function triggerDownload(blob, fileName) {
    if (typeof Blob !== 'undefined' && !(blob instanceof Blob)) {
      throw new Error('No se ha podido generar el archivo PDF.');
    }

    const { document: doc, URL: urlApi, navigator } = global;

    if (!blob) {
      throw new Error('El certificado generado está vacío.');
    }

    if (navigator && typeof navigator.msSaveOrOpenBlob === 'function') {
      navigator.msSaveOrOpenBlob(blob, fileName);
      return;
    }

    if (!doc || !urlApi || typeof urlApi.createObjectURL !== 'function') {
      throw new Error('El navegador no soporta la descarga automática de archivos.');
    }

    const downloadUrl = urlApi.createObjectURL(blob);
    const link = doc.createElement('a');
    link.href = downloadUrl;
    link.download = fileName;
    link.rel = 'noopener';
    link.style.display = 'none';
    doc.body.appendChild(link);
    link.click();
    doc.body.removeChild(link);

    setTimeout(() => {
      urlApi.revokeObjectURL(downloadUrl);
    }, 0);
  }

  async function generateCertificate(row) {
    if (!global.pdfMake || typeof global.pdfMake.createPdf !== 'function') {
      throw new Error('pdfMake no está disponible.');
    }
    const docDefinition = await buildDocDefinition(row || {});
    const fileName = buildFileName(row || {});

    return new Promise((resolve, reject) => {
      let pdfDocument;
      try {
        pdfDocument = global.pdfMake.createPdf(docDefinition);
      } catch (error) {
        reject(error);
        return;
      }

      try {
        pdfDocument.getBlob((blob) => {
          try {
            triggerDownload(blob, fileName);
            resolve({ fileName });
          } catch (error) {
            reject(error);
          }
        });
      } catch (error) {
        reject(error);
      }
    });
  }

  global.certificatePdf = {
    generate: generateCertificate,
    buildDocDefinition
  };
})(window);
