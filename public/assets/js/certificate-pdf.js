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

  const SIDEBAR_WIDTH_REDUCTION = 0.55;
  const FOOTER_WIDTH_REDUCTION = 0.816; // 15% size reduction applied to previous value
  const FOOTER_LEFT_ADDITIONAL_OFFSET = 20;
  const FOOTER_BASELINE_PADDING = 18;
  const BACKGROUND_MARGIN_BLEED = 20;
  const LOGO_VERTICAL_SHIFT_RATIO = 0.05;
  const FONT_SIZE_ADJUSTMENT = 2;
  const LINE_HEIGHT_REDUCTION = 0.5;
  const MIN_LINE_HEIGHT = 0.7;
  const PRACTICE_COLUMN_SHIFT_RATIO = 0.1;
  const BACKGROUND_HORIZONTAL_SHIFT_RATIO = 0.15;
  const BACKGROUND_VERTICAL_SHIFT_RATIO = 0.05;
  const LOGO_HORIZONTAL_SHIFT_RATIO = 0.30;

  const PAGE_DIMENSIONS = {
    width: 841.89,
    height: 595.28
  };

  const FONT_SOURCES = {
    'Poppins-Regular.ttf': [
      'assets/certificados/Poppins-Regular.ttf',
      'https://cdn.jsdelivr.net/npm/@fontsource/poppins@5.0.17/files/poppins-latin-400-normal.ttf'
    ],
    'Poppins-Italic.ttf': [
      'assets/certificados/Poppins-Italic.ttf',
      'https://cdn.jsdelivr.net/npm/@fontsource/poppins@5.0.17/files/poppins-latin-400-italic.ttf'
    ],
    'Poppins-SemiBold.ttf': [
      'assets/certificados/Poppins-SemiBold.ttf',
      'https://cdn.jsdelivr.net/npm/@fontsource/poppins@5.0.17/files/poppins-latin-600-normal.ttf'
    ],
    'Poppins-SemiBoldItalic.ttf': [
      'assets/certificados/Poppins-SemiBoldItalic.ttf',
      'https://cdn.jsdelivr.net/npm/@fontsource/poppins@5.0.17/files/poppins-latin-600-italic.ttf'
    ]
  };

  let poppinsFontPromise = null;

  const assetCache = new Map();
  const trainingTemplates = global.trainingTemplates || null;

  function adjustFontSize(size) {
    return typeof size === 'number' ? size + FONT_SIZE_ADJUSTMENT : size;
  }

  function adjustLineHeight(value) {
    if (typeof value !== 'number') {
      return value;
    }
    return Math.max(MIN_LINE_HEIGHT, value - LINE_HEIGHT_REDUCTION);
  }

  function buildTrainingDetailsContent(details, options = {}) {
    if (!details) {
      return [];
    }

    const theoryItems = Array.isArray(details.theory) ? details.theory : [];
    const practiceItems = Array.isArray(details.practice) ? details.practice : [];
    const columns = [];
    const practiceColumnShift =
      typeof options.practiceColumnShift === 'number' ? options.practiceColumnShift : 0;

    if (theoryItems.length) {
      columns.push({
        width: '*',
        stack: [
          { text: 'Parte teórica', style: 'sectionHeading' },
          {
            ul: theoryItems.map((item) => ({ text: item, style: 'theoryListItem' })),
            margin: [0, 2, 0, 0]
          }
        ]
      });
    }

    if (practiceItems.length) {
      columns.push({
        width: '*',
        margin: practiceColumnShift ? [-practiceColumnShift, 0, 0, 0] : [0, 0, 0, 0],
        stack: [
          { text: 'Parte práctica', style: 'sectionHeading' },
          {
            ul: practiceItems.map((item) => ({ text: item, style: 'listItem' })),
            margin: [0, 2, 0, 0]
          }
        ]
      });
    }

    if (!columns.length) {
      return [];
    }

    if (columns.length === 1) {
      columns.push({ width: '*', text: '', margin: [0, 0, 0, 0] });
    }

    return [
      {
        columns,
        columnGap: 12,
        margin: [0, 4, 0, 12]
      }
    ];
  }

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
            Object.entries(FONT_SOURCES).map(async ([name, sources]) => {
              const sourceList = Array.isArray(sources) ? sources : [sources];
              let lastError = null;

              for (const source of sourceList) {
                try {
                  const data = await loadFontAsBase64(source);
                  if (source !== sourceList[0]) {
                    console.info(
                      `Se está usando la fuente Poppins (${name}) desde el origen alternativo: ${source}`
                    );
                  }
                  return [name, data];
                } catch (error) {
                  lastError = error;
                }
              }

              console.warn(`No se ha podido cargar la fuente Poppins (${name}).`, lastError);
              return null;
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

  function parseDateValue(value) {
    const normalised = normaliseText(value);
    if (!normalised) {
      return null;
    }
    const parsed = new Date(normalised);
    if (Number.isNaN(parsed.getTime())) {
      return null;
    }
    return parsed;
  }

  function formatTrainingDateRange(primaryValue, secondaryValue) {
    const hasSecondary = normaliseText(secondaryValue) !== '';
    if (!hasSecondary) {
      return formatTrainingDate(primaryValue);
    }

    const hasPrimary = normaliseText(primaryValue) !== '';
    if (!hasPrimary) {
      return formatTrainingDate(secondaryValue);
    }

    const primaryDate = parseDateValue(primaryValue);
    const secondaryDate = parseDateValue(secondaryValue);

    if (primaryDate && secondaryDate) {
      const sameYear = primaryDate.getFullYear() === secondaryDate.getFullYear();
      if (sameYear) {
        const yearFormatter = new Intl.DateTimeFormat('es-ES', { year: 'numeric' });
        const yearLabel = yearFormatter.format(primaryDate);

        if (primaryDate.getMonth() === secondaryDate.getMonth()) {
          const dayFormatter = new Intl.DateTimeFormat('es-ES', { day: 'numeric' });
          const monthFormatter = new Intl.DateTimeFormat('es-ES', { month: 'long' });
          const firstDay = dayFormatter.format(primaryDate);
          const secondDay = dayFormatter.format(secondaryDate);
          const monthLabel = monthFormatter.format(primaryDate);
          return `${firstDay} y ${secondDay} de ${monthLabel} de ${yearLabel}`;
        }

        const dayMonthFormatter = new Intl.DateTimeFormat('es-ES', { day: 'numeric', month: 'long' });
        const firstDayMonth = dayMonthFormatter.format(primaryDate);
        const secondDayMonth = dayMonthFormatter.format(secondaryDate);
        return `${firstDayMonth} y ${secondDayMonth} de ${yearLabel}`;
      }

      const fullFormatter = new Intl.DateTimeFormat('es-ES', {
        day: 'numeric',
        month: 'long',
        year: 'numeric'
      });
      return `${fullFormatter.format(primaryDate)} y ${fullFormatter.format(secondaryDate)}`;
    }

    const formattedPrimary = formatTrainingDate(primaryValue);
    const formattedSecondary = formatTrainingDate(secondaryValue);

    if (formattedPrimary === '________') {
      return formattedSecondary;
    }
    if (formattedSecondary === '________') {
      return formattedPrimary;
    }

    return `${formattedPrimary} y ${formattedSecondary}`;
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

  function calculateSidebarWidth(pageWidth) {
    const baseSidebarWidth = Math.min(70, pageWidth * 0.08);
    return baseSidebarWidth * 0.85 * SIDEBAR_WIDTH_REDUCTION;
  }

  function calculateFooterGeometry(pageWidth, pageHeight, pageMargins) {
    const sidebarWidth = calculateSidebarWidth(pageWidth);
    const footerBaseWidth = Math.min(pageWidth - 40, 780);
    const footerMinLeft = Math.max(
      0,
      sidebarWidth + FOOTER_BASELINE_PADDING + FOOTER_LEFT_ADDITIONAL_OFFSET
    );
    const footerMaxWidth = Math.max(0, pageWidth - footerMinLeft - 30);
    const footerWidthBase = Math.min(footerBaseWidth * 0.8, footerMaxWidth);
    const footerWidth = footerWidthBase * FOOTER_WIDTH_REDUCTION;
    const footerHeight = footerWidth * IMAGE_ASPECT_RATIOS.footer;
    const bottomLift = pageMargins[3] * 0.1;
    const footerY = Math.max(0, pageHeight - footerHeight - bottomLift);
    return { footerY, footerMinLeft, footerWidth, footerHeight };
  }

  function buildTrainerBlock(row, geometry, pageMargins) {
    const trainer = normaliseText(row.irata);
    if (!trainer) {
      return null;
    }

    const label = `Formador: ${trainer}`;
    const x = Math.max(geometry.footerMinLeft + 10, pageMargins[0]);
    const y = Math.max(0, geometry.footerY - 18);

    return {
      text: label,
      fontSize: adjustFontSize(9),
      color: '#1f274d',
      absolutePosition: { x, y },
      margin: [0, 0, 0, 0]
    };
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
        fontSize: adjustFontSize(10),
        lineHeight: adjustLineHeight(1.35)
      },
      introText: {
        fontSize: adjustFontSize(10),
        lineHeight: adjustLineHeight(1.35),
        margin: [0, 0, 0, 8]
      },
      certificateTitle: {
        fontSize: adjustFontSize(26),
        bold: true,
        color: '#c4143c',
        letterSpacing: 3,
        margin: [0, 8, 0, 8]
      },
      highlighted: {
        fontSize: adjustFontSize(12),
        bold: true,
        margin: [0, 6, 0, 6]
      },
      trainingName: {
        fontSize: adjustFontSize(16),
        bold: true,
        margin: [0, 12, 0, 0]
      },
      contentSectionTitle: {
        fontSize: adjustFontSize(12),
        bold: true,
        margin: [0, 12, 0, 6]
      },
      sectionHeading: {
        fontSize: adjustFontSize(11),
        bold: true,
        margin: [0, 0, 0, 4]
      },
      listItem: {
        fontSize: adjustFontSize(8),
        lineHeight: adjustLineHeight(1.2),
        margin: [0, 0, 0, 3]
      },
      theoryListItem: {
        fontSize: adjustFontSize(7.5),
        lineHeight: adjustLineHeight(1.15),
        margin: [0, 0, 0, 3]
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

    const referencePageMargins = [105, 40, 160, 100];
    const pageWidth = PAGE_DIMENSIONS.width;
    const pageHeight = PAGE_DIMENSIONS.height;
    const leftMargin = Math.max(
      0,
      calculateSidebarWidth(pageWidth) + FOOTER_BASELINE_PADDING + FOOTER_LEFT_ADDITIONAL_OFFSET
    );
    const baseContentWidth = pageWidth - referencePageMargins[0] - referencePageMargins[2];
    const contentWidthAdjustment = 15;
    const adjustedContentWidth = Math.max(0, baseContentWidth - contentWidthAdjustment);
    const minimumRightMargin = referencePageMargins[2] + 40;
    const rightMargin = Math.max(
      minimumRightMargin,
      pageWidth - leftMargin - adjustedContentWidth
    );
    const pageMargins = [
      leftMargin,
      referencePageMargins[1],
      rightMargin,
      referencePageMargins[3]
    ];
    const fullName = buildFullName(row);
    const documentSentence = buildDocumentSentence(row);
    const trainingDate = formatTrainingDateRange(row.fecha, row.segundaFecha);
    const location = formatLocation(row.lugar);
    const duration = formatDuration(row.duracion);
    const trainingTitle = trainingTemplates
      ? trainingTemplates.getTrainingTitle(row.formacion)
      : row.formacion;
    const trainingName = formatTrainingName(trainingTitle);
    const trainingDetails = trainingTemplates
      ? trainingTemplates.getTrainingDetails(row.formacion)
      : null;
    const footerGeometry = calculateFooterGeometry(pageWidth, pageHeight, pageMargins);
    const trainerBlock = buildTrainerBlock(row, footerGeometry, pageMargins);

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
      { text: trainingName, style: 'trainingName' }
    ];

    const availableContentWidth = Math.max(0, pageWidth - pageMargins[0] - pageMargins[2]);
    const trainingDetailsContent = buildTrainingDetailsContent(trainingDetails, {
      practiceColumnShift: availableContentWidth * PRACTICE_COLUMN_SHIFT_RATIO
    });
    if (trainingDetailsContent.length) {
      contentStack.push({ text: 'Contenidos de la formación', style: 'contentSectionTitle' });
      contentStack.push(...trainingDetailsContent);
    }
    contentStack.push({ text: '\n', margin: [0, 0, 0, 0] });

    const docDefinition = {
      pageOrientation: 'landscape',
      pageSize: 'A4',
      pageMargins,
      background: function (currentPage, pageSize) {
        const pageWidth = pageSize.width || PAGE_DIMENSIONS.width;
        const pageHeight = pageSize.height || PAGE_DIMENSIONS.height;
        const sidebarWidth = calculateSidebarWidth(pageWidth);
        const backgroundWidth = Math.min(320, pageWidth * 0.35);
        const backgroundBaseX = pageWidth - backgroundWidth + backgroundWidth * 0.12;
        const backgroundX = backgroundBaseX + backgroundWidth * BACKGROUND_HORIZONTAL_SHIFT_RATIO;
        const backgroundBaseHeight = backgroundWidth * IMAGE_ASPECT_RATIOS.background;
        const topBleed = (pageMargins[1] || 0) + BACKGROUND_MARGIN_BLEED;
        const bottomBleed = (pageMargins[3] || 0) + BACKGROUND_MARGIN_BLEED;
        const backgroundHeight = backgroundBaseHeight + topBleed + bottomBleed;
        const backgroundBaseY = (pageHeight - backgroundBaseHeight) / 2 - topBleed;
        const backgroundY = backgroundBaseY - backgroundHeight * BACKGROUND_VERTICAL_SHIFT_RATIO;
        const { footerMinLeft, footerWidth, footerY } = calculateFooterGeometry(
          pageWidth,
          pageHeight,
          pageMargins
        );
        const logoWidth = Math.min(backgroundWidth * 0.6, 200);
        const logoHeight = logoWidth * IMAGE_ASPECT_RATIOS.logo;
        const logoVerticalShift = pageHeight * LOGO_VERTICAL_SHIFT_RATIO;
        const logoY = Math.max(0, (pageHeight - logoHeight) / 2 + logoVerticalShift);
        const logoBaseX = backgroundX + (backgroundWidth - logoWidth) / 2;
        const logoX = logoBaseX - logoWidth * LOGO_HORIZONTAL_SHIFT_RATIO;

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
            height: backgroundHeight,
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
          margin: [0, 8, 0, 0],
          stack: contentStack
        }
      ],
      styles: buildDocStyles(),
      defaultStyle: {
        fontSize: adjustFontSize(10),
        lineHeight: adjustLineHeight(1.35),
        color: '#1f274d',
        font: 'Poppins'
      },
      info: {
        title: `Certificado - ${fullName}`,
        author: 'GEPCO Formación',
        subject: trainingName
      }
    };

    if (trainerBlock) {
      docDefinition.content.push(trainerBlock);
    }

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
