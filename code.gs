/**
 * Genera un código alfanumérico de 16 caracteres
 */
function generarID16() {
  const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
  let id = "";
  for (let i = 0; i < 16; i++) {
    id += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return id;
}

/**
 * Se ejecuta automáticamente cuando se envía el formulario
 */
function onFormSubmit(e) {
    const sheet = e.range.getSheet();
      const row = e.range.getRow();

        const id = generarID16();

          // L = ID
            sheet.getRange(row, 12).setValue(id);

              // M = URL
                sheet.getRange(row, 13).setValue(
                `https://diegoturijancontacto-sudo.github.io/Certificados/${id}`);


                  // 🔥 HACER PÚBLICA LA IMAGEN
                    const imagenUrl = e.namedValues["Imagen de la obra"]?.[0];
                      if (imagenUrl) {
                          const driveId = extractDriveId(imagenUrl);
                              DriveApp.getFileById(driveId).setSharing(
                                    DriveApp.Access.ANYONE_WITH_LINK,
                                          DriveApp.Permission.VIEW
                                              );
                                                }
                                                }

function extractDriveId(url) {
    if (!url) return null;
      const match = String(url).match(/(?:id=|\/d\/|file\/d\/)([a-zA-Z0-9_-]+)/);
        return match ? match[1] : null;
        }

const CERTIFICADOS_FOLDER_ID = '';
const CERTIFICADO_BASE_URL = 'https://diegoturijancontacto-sudo.github.io/Certificados/';

function jsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function normalizeCertificateData_(data) {
  const obra = data || {};
  const adjuntos = Array.isArray(obra.adjuntos) ? obra.adjuntos : [];
  const original = adjuntos.find(adjunto => adjunto && adjunto.typeCode === 'RW') ||
    adjuntos.find(adjunto => adjunto && (
      String(adjunto.mimeType || '').indexOf('image/') === 0 ||
      /\.(jpg|jpeg|png|gif|webp)$/i.test(adjunto.name || '')
    ));
  const imageUrl = obra.imagenUrl || (original && original.url) || '';
  const imageFileId = obra.imagenFileId || (original && original.id) || extractDriveId(imageUrl);

  return {
    obraId: String(obra.id || obra.obraId || '').trim(),
    categoria: String(obra.categoria || obra.tipo_obra || '').trim(),
    soporte: 'desconocido',
    alto: obra.alto ?? '',
    ancho: obra.ancho ?? '',
    profundo: obra.largo ?? obra.profundo ?? '',
    nombreObra: String(obra.nombre_obra || obra.nombreObra || '').trim(),
    autor: String(obra.autor || '').trim(),
    serie: String(obra.clave || obra.serie || '').trim(),
    imageUrl: imageUrl,
    imageFileId: imageFileId ? String(imageFileId).trim() : ''
  };
}

function getCertificateSheet_() {
  const sheet = SpreadsheetApp.getActive().getSheets()[0];
  const headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0]
    .map(value => String(value || '').trim().toLowerCase());
  const ensureHeader = (name) => {
    const normalized = name.toLowerCase();
    const existing = headers.indexOf(normalized);
    if (existing >= 0) return existing + 1;
    const column = sheet.getLastColumn() + 1;
    sheet.getRange(1, column).setValue(name);
    headers.push(normalized);
    return column;
  };
  return {
    sheet,
    obraIdColumn: ensureHeader('Obra ID'),
    pdfColumn: ensureHeader('PDF URL')
  };
}

function findExistingCertificate_(certificate) {
  const table = getCertificateSheet_();
  const values = table.sheet.getDataRange().getValues();
  const obraId = String(certificate.obraId || '').trim();
  const serie = String(certificate.serie || '').trim();
  const name = String(certificate.nombreObra || '').trim().toLowerCase();

  for (let row = 1; row < values.length; row++) {
    const current = values[row];
    const sameObra = obraId && String(current[table.obraIdColumn - 1] || '').trim() === obraId;
    const sameLegacyRecord = serie && String(current[9] || '').trim() === serie &&
      (!name || String(current[7] || '').trim().toLowerCase() === name);
    if (sameObra || sameLegacyRecord) {
      const id = String(current[11] || '').trim();
      const certificateUrl = String(current[12] || '').trim() || (id ? CERTIFICADO_BASE_URL + id : '');
      const pdfUrl = String(current[table.pdfColumn - 1] || current[13] || '').trim();
      if (id && pdfUrl) {
        return { id, certificateUrl, pdfUrl };
      }
    }
  }
  return null;
}

function getCertificatesFolder_() {
  if (CERTIFICADOS_FOLDER_ID) return DriveApp.getFolderById(CERTIFICADOS_FOLDER_ID);
  return DriveApp.getRootFolder();
}

function formatCertificateValue_(value) {
  return value === null || value === undefined || value === '' ? 'No especificado' : String(value);
}

function createCertificatePdf_(certificate, certificateId) {
  const title = certificate.nombreObra || 'Obra sin nombre';
  const document = DocumentApp.create('Certificado - ' + title);
  const body = document.getBody();
  body.setMarginTop(42).setMarginBottom(42).setMarginLeft(54).setMarginRight(54);

  const heading = body.appendParagraph('CERTIFICADO');
  heading.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  heading.editAsText().setBold(true).setFontSize(22);

  const subtitle = body.appendParagraph('Certificado de autenticidad y registro de obra');
  subtitle.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  subtitle.editAsText().setFontSize(10).setForegroundColor('#666666');
  body.appendHorizontalRule();

  if (certificate.imageFileId) {
    try {
      const image = DriveApp.getFileById(certificate.imageFileId).getBlob();
      const imageParagraph = body.appendParagraph('');
      imageParagraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      const insertedImage = imageParagraph.appendInlineImage(image);
      insertedImage.setWidth(300);
      if (insertedImage.getHeight() > 360) insertedImage.setHeight(360);
    } catch (error) {
      console.log('No se pudo insertar la imagen del certificado: ' + error);
    }
  }

  const table = body.appendTable([
    ['Nombre de la obra', formatCertificateValue_(certificate.nombreObra)],
    ['Autor', formatCertificateValue_(certificate.autor)],
    ['Categoría', formatCertificateValue_(certificate.categoria)],
    ['Soporte', formatCertificateValue_(certificate.soporte)],
    ['Alto', formatCertificateValue_(certificate.alto) + ' cm'],
    ['Ancho', formatCertificateValue_(certificate.ancho) + ' cm'],
    ['Profundo', formatCertificateValue_(certificate.profundo) + ' cm'],
    ['Serie / WORK ID', formatCertificateValue_(certificate.serie)]
  ]);
  table.setBorderWidth(0.5);
  for (let row = 0; row < table.getNumRows(); row++) {
    table.getCell(row, 0).editAsText().setBold(true);
  }

  body.appendParagraph('');
  const issue = body.appendParagraph('ID de certificado: ' + certificateId + '\nFecha de emisión: ' + new Date().toLocaleDateString('es-MX'));
  issue.editAsText().setFontSize(9).setForegroundColor('#666666');
  const folder = getCertificatesFolder_();
  document.saveAndClose();
  const pdfBlob = DriveApp.getFileById(document.getId()).getAs(MimeType.PDF)
    .setName('Certificado-' + (certificate.serie || certificateId) + '.pdf');
  const pdfFile = folder.createFile(pdfBlob);
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  DriveApp.getFileById(document.getId()).setTrashed(true);
  return {
    fileId: pdfFile.getId(),
    pdfUrl: 'https://drive.google.com/uc?export=download&id=' + pdfFile.getId(),
    viewUrl: pdfFile.getUrl()
  };
}

function appendCertificateRow_(certificate, certificateId, pdfUrl) {
  const table = getCertificateSheet_();
  const sheet = table.sheet;
  const row = Array(Math.max(sheet.getLastColumn(), table.obraIdColumn, table.pdfColumn)).fill('');
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

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.action !== 'certificado_generar') {
      return jsonResponse_({ ok: false, message: 'Acción no válida' });
    }

    const certificate = normalizeCertificateData_(body.data);
    if (!certificate.nombreObra) throw new Error('Falta el nombre de la obra');

    const existing = findExistingCertificate_(certificate);
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
    const pdf = createCertificatePdf_(certificate, certificateId);
    appendCertificateRow_(certificate, certificateId, pdf.pdfUrl);

    return jsonResponse_({
      ok: true,
      created: true,
      id: certificateId,
      pdfUrl: pdf.pdfUrl,
      viewUrl: pdf.viewUrl,
      certificateUrl: CERTIFICADO_BASE_URL + certificateId,
      data: certificate
    });
  } catch (error) {
    return jsonResponse_({ ok: false, message: error.message || String(error) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function doGet(e) {
    const idBuscado = e.parameter.id;
      const sheet = SpreadsheetApp.getActive().getSheets()[0];
        const data = sheet.getDataRange().getValues();

          for (let i = 1; i < data.length; i++) {
              const fila = data[i];

                  const imagenForm = fila[10]; // K → imagen del Form
                      const idFila = fila[11];     // L → ID

                          if (String(idFila).trim() === String(idBuscado).trim()) {

                                const driveId = extractDriveId(imagenForm);
                                      const imagenURL = driveId
                                              ? `https://lh3.googleusercontent.com/d/${driveId}`
                                                      : null;

                                                            return ContentService
                                                                    .createTextOutput(JSON.stringify({
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
                                                                                                                                                                                          }))
                                                                                                                                                                                                  .setMimeType(ContentService.MimeType.JSON);
                                                                                                                                                                                                      }
                                                                                                                                                                                                        }

                                                                                                                                                                                                          return ContentService
                                                                                                                                                                                                              .createTextOutput(JSON.stringify({ ok: false }))
                                                                                                                                                                                                                  .setMimeType(ContentService.MimeType.JSON);
                                                                                                                                                                                                                  }
