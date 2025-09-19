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

  const DATE_PART_FORMATTERS = {
    day: new Intl.DateTimeFormat('es-ES', { day: 'numeric' }),
    month: new Intl.DateTimeFormat('es-ES', { month: 'long' }),
    year: new Intl.DateTimeFormat('es-ES', { year: 'numeric' })
  };

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

  function extractDateParts(value) {
    const normalised = normaliseText(value);
    if (!normalised) {
      return null;
    }

    const parsed = new Date(normalised);
    if (Number.isNaN(parsed.getTime())) {
      return null;
    }

    return {
      day: DATE_PART_FORMATTERS.day.format(parsed),
      month: DATE_PART_FORMATTERS.month.format(parsed),
      year: DATE_PART_FORMATTERS.year.format(parsed)
    };
  }

  function buildTrainingDateText(primaryValue, secondaryValue) {
    const primaryText = formatTrainingDate(primaryValue);
    const secondaryNormalised = normaliseText(secondaryValue);

    if (!secondaryNormalised) {
      return primaryText;
    }

    const secondaryText = formatTrainingDate(secondaryValue);
    const primaryParts = extractDateParts(primaryValue);
    const secondaryParts = extractDateParts(secondaryValue);

    if (primaryParts && secondaryParts) {
      const sameMonth = primaryParts.month === secondaryParts.month && primaryParts.year === secondaryParts.year;
      const sameDay = sameMonth && primaryParts.day === secondaryParts.day;

      if (sameDay) {
        return primaryText;
      }

      if (sameMonth) {
        return `${primaryParts.day} y ${secondaryParts.day} de ${primaryParts.month} de ${primaryParts.year}`;
      }

      return `${primaryText} y ${secondaryText}`;
    }

    return `${primaryText} y ${secondaryText}`;
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
        fontSize: 14,
        lineHeight: 1.4
      },
      introText: {
        fontSize: 14,
        lineHeight: 1.5,
        margin: [0, 0, 0, 12]
      },
      certificateTitle: {
        fontSize: 34,
        bold: true,
        color: '#c4143c',
        letterSpacing: 3,
        margin: [0, 8, 0, 12]
      },
      highlighted: {
        fontSize: 16,
        bold: true,
        margin: [0, 8, 0, 8]
      },
      trainingName: {
        fontSize: 20,
        bold: true,
        margin: [0, 16, 0, 0]
      }
    };
  }

  async function buildDocDefinition(row) {
    const [backgroundImage, sidebarImage, footerImage, logoImage] = await Promise.all([
      getCachedAsset('background'),
      getCachedAsset('leftSidebar'),
      getCachedAsset('footer'),
      getCachedAsset('logo')
    ]);

    const pageMargins = [140, 70, 170, 110];
    const fullName = buildFullName(row);
    const documentSentence = buildDocumentSentence(row);
    const trainingDateText = buildTrainingDateText(row.fecha, row.segundaFecha);
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
        text: `${documentSentence}, quien en fecha ${trainingDateText} y en ${location}`,
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
        const sidebarWidth = Math.min(70, pageWidth * 0.08);
        const backgroundWidth = Math.min(320, pageWidth * 0.35);
        const backgroundX = pageWidth - backgroundWidth;
        const backgroundHeight = backgroundWidth * IMAGE_ASPECT_RATIOS.background;
        const backgroundY = 0;
        const backgroundBottom = backgroundY + backgroundHeight;
        const visibleBackgroundTop = Math.max(0, backgroundY);
        const visibleBackgroundBottom = Math.min(pageHeight, backgroundBottom);
        const availableBackgroundHeight = Math.max(
          0,
          visibleBackgroundBottom - visibleBackgroundTop
        );
        const footerWidth = Math.min(pageWidth - 40, 780);
        const footerHeight = footerWidth * IMAGE_ASPECT_RATIOS.footer;
        const footerX = Math.max(20, (pageWidth - footerWidth) / 2);
        const footerY = pageHeight - footerHeight;
        const logoWidth = Math.min(backgroundWidth * 0.7, 220);
        const logoHeight = logoWidth * IMAGE_ASPECT_RATIOS.logo;
        const logoYWithinBackground = availableBackgroundHeight
          ? visibleBackgroundTop + Math.max(0, (availableBackgroundHeight - logoHeight) / 2)
          : visibleBackgroundTop;
        const logoY = Math.min(pageHeight - logoHeight, Math.max(0, logoYWithinBackground));
        const logoX = backgroundX + (backgroundWidth - logoWidth) / 2;

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
            absolutePosition: { x: footerX, y: footerY }
          }
        ];
      },
      content: [
        {
          margin: [0, 40, 0, 0],
          stack: contentStack
        }
      ],
      styles: buildDocStyles(),
      defaultStyle: {
        fontSize: 14,
        lineHeight: 1.5,
        color: '#1f274d'
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
