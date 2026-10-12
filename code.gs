
/**
 * RTBrok — Sistema de certificados
 *
 * Funciones:
 * - Generar certificados desde el formulario o mediante doPost().
 * - Crear PDF y guardarlo en Google Drive.
 * - Evitar duplicados cuando se encuentra un certificado existente.
 * - Verificar un certificado por ID mediante ?id=...
 * - Consultar el registro público mediante ?action=list
 */

// ============================================================
// CONFIGURACIÓN
// ============================================================

const CERTIFICADOS_FOLDER_ID = '';

const CERTIFICADO_BASE_URL =
  'https://diegoturijancontacto-sudo.github.io/Certificados/';

// ============================================================
// GENERACIÓN DE ID
// ============================================================

function generarID16() {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  let id = '';

  for (let i = 0; i < 16; i++) {
    id += chars.charAt(Math.floor(Math.random() * chars.length));
  }

  return id;
}

// ============================================================
// TRIGGER DE GOOGLE FORMS
// ============================================================

function onFormSubmit(e) {
  const sheet = e.range.getSheet();
  const row = e.range.getRow();

  const id = generarID16();

  // Columna L: ID del certificado.
  sheet.getRange(row, 12).setValue(id);

  // Columna M: URL de verificación.
  sheet.getRange(row, 13).setValue(
    CERTIFICADO_BASE_URL + id
  );

  // Hacer pública la imagen de la obra.
  const imagenUrl = e.namedValues['Imagen de la obra']?.[0];

  if (imagenUrl) {
    const driveId = extractDriveId(imagenUrl);

    if (driveId) {
      DriveApp.getFileById(driveId).setSharing(
        DriveApp.Access.ANYONE_WITH_LINK,
        DriveApp.Permission.VIEW
      );
    }
  }
}

// ============================================================
// UTILIDADES
// ============================================================

function extractDriveId(url) {
  if (!url) return null;

  const match = String(url).match(
    /(?:id=|\/d\/|file\/d\/)([a-zA-Z0-9_-]+)/
  );

  return match ? match[1] : null;
}

function jsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Respuesta JSONP para consultar el registro desde GitHub Pages.
 */
function jsonpResponse_(callback, payload) {
  if (
    !/^[a-zA-Z_$][0-9a-zA-Z_$]*(?:\.[a-zA-Z_$][0-9a-zA-Z_$]*)*$/
      .test(callback)
  ) {
    return jsonResponse_({
      ok: false,
      message: 'Callback no válido'
    });
  }

  const json = JSON.stringify(payload)
    .replace(/</g, '\\u003c')
    .replace(/>/g, '\\u003e')
    .replace(/&/g, '\\u0026');

  return ContentService
    .createTextOutput(callback + '(' + json + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function formatCertificateValue_(value) {
  return value === null || value === undefined || value === ''
    ? 'No especificado'
    : String(value);
}

// ============================================================
// NORMALIZACIÓN DE DATOS
// ============================================================

function normalizeCertificateData_(data) {
  const obra = data || {};
  const adjuntos = Array.isArray(obra.adjuntos)
    ? obra.adjuntos
    : [];

  const original =
    adjuntos.find(adjunto =>
      adjunto && adjunto.typeCode === 'RW'
    ) ||
    adjuntos.find(adjunto =>
      adjunto && (
        String(adjunto.mimeType || '').indexOf('image/') === 0 ||
        /\.(jpg|jpeg|png|gif|webp)$/i.test(adjunto.name || '')
      )
    );

  const imageUrl =
    obra.imagenUrl ||
    (original && original.url) ||
    '';

  const imageFileId =
    obra.imagenFileId ||
    (original && original.id) ||
    extractDriveId(imageUrl);

  return {
    obraId: String(obra.id || obra.obraId || '').trim(),
    categoria: String(
      obra.categoria || obra.tipo_obra || ''
    ).trim(),
    soporte: 'desconocido',
    alto: obra.alto ?? '',
    ancho: obra.ancho ?? '',
    profundo: obra.largo ?? obra.profundo ?? '',
    nombreObra: String(
      obra.nombre_obra || obra.nombreObra || ''
    ).trim(),
    autor: String(obra.autor || '').trim(),
    serie: String(obra.clave || obra.serie || '').trim(),
    imageUrl: imageUrl,
    imageFileId: imageFileId
      ? String(imageFileId).trim()
      : ''
  };
}

// ============================================================
// HOJA DE CERTIFICADOS
// ============================================================

function getCertificateSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  if (!spreadsheet) {
    throw new Error(
      'No se encontró la hoja de cálculo vinculada al proyecto.'
    );
  }

  const sheet = spreadsheet.getSheets()[0];

  const lastColumn = Math.max(sheet.getLastColumn(), 1);

  const headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getValues()[0]
    .map(value => String(value || '').trim().toLowerCase());

  const ensureHeader = name => {
    const normalized = name.toLowerCase();
    const existing = headers.indexOf(normalized);

    if (existing >= 0) {
      return existing + 1;
    }

    const column = sheet.getLastColumn() + 1;

    sheet.getRange(1, column).setValue(name);
    headers.push(normalized);

    return column;
  };

  return {
    sheet: sheet,
    obraIdColumn: ensureHeader('Obra ID'),
    pdfColumn: ensureHeader('PDF URL')
  };
}

// ============================================================
// BUSCAR CERTIFICADOS EXISTENTES
// ============================================================

function findExistingCertificate_(certificate) {
  const table = getCertificateSheet_();
  const values = table.sheet.getDataRange().getValues();

  const obraId = String(certificate.obraId || '').trim();
  const serie = String(certificate.serie || '').trim();
  const name = String(
    certificate.nombreObra || ''
  ).trim().toLowerCase();

  for (let row = 1; row < values.length; row++) {
    const current = values[row];

    const sameObra =
      obraId &&
      String(
        current[table.obraIdColumn - 1] || ''
      ).trim() === obraId;

    const sameLegacyRecord =
      serie &&
      String(current[9] || '').trim() === serie &&
      (
        !name ||
        String(current[7] || '')
          .trim()
          .toLowerCase() === name
      );

    if (sameObra || sameLegacyRecord) {
      const id = String(current[11] || '').trim();

      const certificateUrl =
        String(current[12] || '').trim() ||
        (id ? CERTIFICADO_BASE_URL + id : '');

      const pdfUrl = String(
        current[table.pdfColumn - 1] ||
        current[13] ||
        ''
      ).trim();

      if (id && pdfUrl) {
        return {
          id: id,
          certificateUrl: certificateUrl,
          pdfUrl: pdfUrl
        };
      }
    }
  }

  return null;
}

// ============================================================
// CARPETA DE CERTIFICADOS
// ============================================================

function getCertificatesFolder_() {
  if (CERTIFICADOS_FOLDER_ID) {
    return DriveApp.getFolderById(CERTIFICADOS_FOLDER_ID);
  }

  return DriveApp.getRootFolder();
}

// ============================================================
// GENERAR PDF
// ============================================================

function createCertificatePdf_(certificate, certificateId) {
  const title = certificate.nombreObra || 'Obra sin nombre';

  const document = DocumentApp.create(
    'Certificado - ' + title
  );

  const body = document.getBody();

  body
    .setMarginTop(42)
    .setMarginBottom(42)
    .setMarginLeft(54)
    .setMarginRight(54);

  const heading = body.appendParagraph('CERTIFICADO');

  heading.setAlignment(
    DocumentApp.HorizontalAlignment.CENTER
  );

  heading.editAsText()
    .setBold(true)
    .setFontSize(22);

  const subtitle = body.appendParagraph(
    'Certificado de autenticidad y registro de obra'
  );

  subtitle.setAlignment(
    DocumentApp.HorizontalAlignment.CENTER
  );

  subtitle.editAsText()
    .setFontSize(10)
    .setForegroundColor('#666666');

  body.appendHorizontalRule();

  if (certificate.imageFileId) {
    try {
      const image = DriveApp
        .getFileById(certificate.imageFileId)
        .getBlob();

      const imageParagraph = body.appendParagraph('');

      imageParagraph.setAlignment(
        DocumentApp.HorizontalAlignment.CENTER
      );

      const insertedImage =
        imageParagraph.appendInlineImage(image);

      insertedImage.setWidth(300);

      if (insertedImage.getHeight() > 360) {
        insertedImage.setHeight(360);
      }
    } catch (error) {
      console.log(
        'No se pudo insertar la imagen del certificado: ' +
        error
      );
    }
  }

  const table = body.appendTable([
    [
      'Nombre de la obra',
      formatCertificateValue_(certificate.nombreObra)
    ],
    [
      'Autor',
      formatCertificateValue_(certificate.autor)
    ],
    [
      'Categoría',
      formatCertificateValue_(certificate.categoria)
    ],
    [
      'Soporte',
      formatCertificateValue_(certificate.soporte)
    ],
    [
      'Alto',
      formatCertificateValue_(certificate.alto) + ' cm'
    ],
    [
      'Ancho',
      formatCertificateValue_(certificate.ancho) + ' cm'
    ],
    [
      'Profundo',
      formatCertificateValue_(certificate.profundo) + ' cm'
    ],
    [
      'Serie / WORK ID',
      formatCertificateValue_(certificate.serie)
    ]
  ]);

  table.setBorderWidth(0.5);

  for (let row = 0; row < table.getNumRows(); row++) {
    table.getCell(row, 0)
      .editAsText()
      .setBold(true);
  }

  body.appendParagraph('');

  const issue = body.appendParagraph(
    'ID de certificado: ' + certificateId +
    '\nFecha de emisión: ' +
    new Date().toLocaleDateString('es-MX')
  );

  issue.editAsText()
    .setFontSize(9)
    .setForegroundColor('#666666');

  const folder = getCertificatesFolder_();

  document.saveAndClose();

  const pdfBlob = DriveApp
    .getFileById(document.getId())
    .getAs(MimeType.PDF)
    .setName(
      'Certificado-' +
      (certificate.serie || certificateId) +
      '.pdf'
    );

  const pdfFile = folder.createFile(pdfBlob);

  pdfFile.setSharing(
    DriveApp.Access.ANYONE_WITH_LINK,
    DriveApp.Permission.VIEW
  );

  // El documento de Google Docs temporal se elimina.
  DriveApp.getFileById(document.getId()).setTrashed(true);

  return {
    fileId: pdfFile.getId(),
    pdfUrl:
      'https://drive.google.com/uc?export=download&id=' +
      pdfFile.getId(),
    viewUrl: pdfFile.getUrl()
  };
}

// ============================================================
// GUARDAR CERTIFICADO EN GOOGLE SHEETS
// ============================================================

function appendCertificateRow_(
  certificate,
  certificateId,
  pdfUrl
) {
  const table = getCertificateSheet_();
  const sheet = table.sheet;

  const row = Array(
    Math.max(
      sheet.getLastColumn(),
      table.obraIdColumn,
      table.pdfColumn,
      13
    )
  ).fill('');

  row[2] = certificate.categoria;
  row[3] = certificate.soporte;
  row[4] = certificate.alto;
  row[5] = certificate.ancho;
  row[6] = certificate.profundo;
  row[7] = certificate.nombreObra;
  row[8] = certificate.autor;
  row[9] = certificate.serie;
  row[10] = certificate.imageUrl || '';
  row[11] = certificateId;
  row[12] = CERTIFICADO_BASE_URL + certificateId;

  row[table.pdfColumn - 1] = pdfUrl;
  row[table.obraIdColumn - 1] = certificate.obraId;

  sheet.appendRow(row);
}

// ============================================================
// API POST: GENERAR CERTIFICADO
// ============================================================

function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    const body = JSON.parse(
      (e && e.postData && e.postData.contents) || '{}'
    );

    if (body.action !== 'certificado_generar') {
      return jsonResponse_({
        ok: false,
        message: 'Acción no válida'
      });
    }

    const certificate =
      normalizeCertificateData_(body.data);

    if (!certificate.nombreObra) {
      throw new Error('Falta el nombre de la obra');
    }

    const existing =
      findExistingCertificate_(certificate);

    if (existing) {
      return jsonResponse_({
        ok: true,
        created: false,
        id: existing.id,
        pdfUrl: existing.pdfUrl,
        viewUrl: existing.pdfUrl,
        certificateUrl: existing.certificateUrl,
        data: certificate
      });
    }

    const certificateId = generarID16();

    const pdf = createCertificatePdf_(
      certificate,
      certificateId
    );

    appendCertificateRow_(
      certificate,
      certificateId,
      pdf.pdfUrl
    );

    return jsonResponse_({
      ok: true,
      created: true,
      id: certificateId,
      pdfUrl: pdf.pdfUrl,
      viewUrl: pdf.viewUrl,
      certificateUrl:
        CERTIFICADO_BASE_URL + certificateId,
      data: certificate
    });

  } catch (error) {
    return jsonResponse_({
      ok: false,
      message: error.message || String(error)
    });

  } finally {
    try {
      lock.releaseLock();
    } catch (_) {}
  }
}

// ============================================================
// LISTADO PÚBLICO DE CERTIFICADOS
// ============================================================

function listarCertificadosPublicos_() {
  const table = getCertificateSheet_();
  const sheet = table.sheet;
  const values = sheet.getDataRange().getValues();

  if (values.length <= 1) {
    return {
      ok: true,
      total: 0,
      certificados: []
    };
  }

  const certificados = [];

  for (let i = 1; i < values.length; i++) {
    const fila = values[i];

    // Columna L: ID del certificado.
    const id = String(fila[11] || '').trim();

    // Ignorar filas sin certificado.
    if (!id) continue;

    // Columna K: imagen original.
    const imagenOriginal = String(fila[10] || '').trim();
    const driveId = extractDriveId(imagenOriginal);

    // Obtener enlaces desde las columnas configuradas.
    const certificadoUrl =
      String(fila[12] || '').trim() ||
      CERTIFICADO_BASE_URL + id;

    const pdfUrl = String(
      fila[table.pdfColumn - 1] || ''
    ).trim();

    certificados.push({
      id: id,
      nombreObra: String(fila[7] || '').trim(),
      autor: String(fila[8] || '').trim(),
      categoria: String(fila[2] || '').trim(),
      soporte: String(fila[3] || '').trim(),
      alto: String(fila[4] || '').trim(),
      ancho: String(fila[5] || '').trim(),
      profundo: String(fila[6] || '').trim(),
      serie: String(fila[9] || '').trim(),
      imagen: driveId
        ? 'https://lh3.googleusercontent.com/d/' + driveId
        : '',
      certificadoUrl: certificadoUrl,
      pdfUrl: pdfUrl
    });
  }

  // Certificados más recientes primero.
  certificados.reverse();

  return {
    ok: true,
    total: certificados.length,
    certificados: certificados
  };
}

// ============================================================
// API GET: LISTADO PÚBLICO Y VERIFICACIÓN INDIVIDUAL
// ============================================================

function doGet(e) {
  const params = (e && e.parameter) || {};

  // ----------------------------------------------------------
  // MODO 1: listado de todos los certificados
  // URL: /exec?action=list
  // JSONP: /exec?action=list&callback=rtbCertificadosCallback
  // ----------------------------------------------------------

  if (params.action === 'list') {
    try {
      const resultado = listarCertificadosPublicos_();

      if (params.callback) {
        return jsonpResponse_(
          params.callback,
          resultado
        );
      }

      return jsonResponse_(resultado);

    } catch (error) {
      const resultado = {
        ok: false,
        message: error.message || String(error),
        total: 0,
        certificados: []
      };

      if (params.callback) {
        return jsonpResponse_(
          params.callback,
          resultado
        );
      }

      return jsonResponse_(resultado);
    }
  }

  // ----------------------------------------------------------
  // MODO 2: verificación individual existente
  // URL: /exec?id=ID_DEL_CERTIFICADO
  // ----------------------------------------------------------

  const idBuscado = String(params.id || '').trim();

  if (!idBuscado) {
    return jsonResponse_({
      ok: false,
      message:
        'Indica un ID de certificado o utiliza action=list'
    });
  }

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheets()[0];

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const fila = data[i];

    // Columna K: imagen del formulario.
    const imagenForm = fila[10];

    // Columna L: ID del certificado.
    const idFila = String(fila[11] || '').trim();

    if (idFila !== idBuscado) continue;

    const driveId = extractDriveId(imagenForm);

    const imagenURL = driveId
      ? 'https://lh3.googleusercontent.com/d/' + driveId
      : null;

    return jsonResponse_({
      ok: true,
      categoria: fila[2],
      soporte: fila[3],
      alto: fila[4],
      ancho: fila[5],
      profundo: fila[6],
      nombreObra: fila[7],
      autor: fila[8],
      serie: fila[9],
      imagen: imagenURL,
      id: idFila
    });
  }

  return jsonResponse_({
    ok: false,
    message: 'Certificado no encontrado'
  });
}
