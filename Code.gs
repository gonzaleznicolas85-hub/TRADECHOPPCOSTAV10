/**
 * SISTEMA DE COMODATOS - BACKEND (Google Apps Script)
 * Manejo de datos dinámicos, fotos, generación de PDFs y Docs con imágenes integradas.
 */

const CONFIG = {
  SHEET_NAME: 'Comodatos_Prueba',
  NUMBER_PREFIX: 'TCC-',
  NUMBER_DIGITS: 5,
  DEPLOY_VERSION: 'FINAL_V6_MIS_COMODATOS',

  // --- REEMPLAZAR CON TUS IDs DE CARPETAS Y DOCUMENTOS ---
  SIGNATURE_FOLDER_ID: '104KUonvvkDTGRFglhZfimoU10pCAeEw5',
  PDF_FOLDER_ID: '1CBA5rVdj5sAJT3XSFDA5aFQXdLLezeeQ',
  TEMPLATE_DOC_ID: '1ts2l2YABwrr3_3XK3S0-X-oSAC54fOM-oxEZWNi8FSw',

  // SOLUCIÓN: Reemplaza este texto por tu ID real
  PHOTOS_PARENT_FOLDER_ID: 'PEGAR_AQUI_EL_ID_DE_LA_CARPETA_DE_FOTOS'
};

// Encabezados exactos de la hoja de cálculo
const HEADERS = [
  'Timestamp', 'ComodatoNumero', 'Fecha', 'Tecnico', 'Distribuidor', 'CodCliente',
  'CUIT', 'NombreFantasia', 'RazonSocial', 'Domicilio', 'Localidad',
  'EquiposDetalle',
  'RegCornelius', 'RegMafridis', 'LlaveMixta', 'CabezalMM', 'MangueraCerveza',
  'MangueraPython', 'CanillaNiquelada', 'CanillaAgua', 'Transformador',
  'SeparadoresCanilla', 'ConectorVasera', 'CantPicos', 'Medallones', 'TuboGas',
  'VaseraRinser', 'Rinser', 'MangueraDesague', 'Handle', 'Celli', 'Vasera',
  'PilonesDetalle',
  'Descripcion', 'FotosEquiposUrls', 'FotosPilonesUrls',
  'Aclaracion', 'DNI',
  'FirmaFileId', 'FirmaUrl', 'DocFileId', 'DocUrl', 'PdfFileId', 'PdfUrl', 'AceptaTerminos'
];

/**
 * 1. CONFIGURACIÓN INICIAL Y MENÚ
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚀 Sistema Comodatos')
    .addItem('Generar PDFs Pendientes', 'generarPdfDesdeBoton')
    .addToUi();
}

function crearEstructuraPrueba() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
  } else {
    sheet.clear();
  }

  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);

  const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
  headerRange.setBackground('#003366');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, HEADERS.length);

  Logger.log('¡Estructura creada con éxito en la pestaña: ' + CONFIG.SHEET_NAME + '!');
}

/**
 * 2. RECEPCIÓN DE PETICIONES GET
 *    - action=nextNumber          -> próximo número de comodato
 *    - action=listTecnicos        -> lista de técnicos ya cargados en la hoja
 *    - action=comodatosPorTecnico -> historial de comodatos de un técnico (param "tecnico")
 */
function doGet(e) {
  try {
    const action = e.parameter.action;

    if (action === 'nextNumber') {
      const nextNumber = getNextComodatoNumber_();
      return jsonOutput_({ ok: true, comodatoNumero: nextNumber });
    }

    if (action === 'listTecnicos') {
      return jsonOutput_({ ok: true, tecnicos: getDistinctTecnicos_() });
    }

    if (action === 'comodatosPorTecnico') {
      const tecnico = e.parameter.tecnico ? String(e.parameter.tecnico).trim() : '';
      if (!tecnico) return jsonOutput_({ ok: false, message: 'Falta el parámetro tecnico.' });
      return jsonOutput_({ ok: true, comodatos: getComodatosByTecnico_(tecnico) });
    }

    return jsonOutput_({ ok: false, message: 'Acción no válida' });
  } catch (error) {
    return jsonOutput_({ ok: false, message: error.message });
  }
}

/**
 * 3. RECEPCIÓN DE DATOS DESDE EL HTML (WEB APP)
 */
function doPost(e) {
  try {
    const sheet = getSheet_();
    if (!e || !e.postData || !e.postData.contents) throw new Error('No llegaron datos.');

    const data = JSON.parse(e.postData.contents);
    const comodatoNumero = data.comodatoNumero && String(data.comodatoNumero).trim() ? String(data.comodatoNumero).trim() : getNextComodatoNumero_();

    // Convertimos los arrays de equipos y pilones a texto
    const txtEquipos = formatEquipos_(data.equipos);
    const txtPilones = formatPilones_(data.pilones);
    data.equiposDetalle = txtEquipos;
    data.pilonesDetalle = txtPilones;

    // Creamos la carpeta de fotos
    const parentFolder = DriveApp.getFolderById(CONFIG.PHOTOS_PARENT_FOLDER_ID);
    const comodatoPhotosFolder = parentFolder.createFolder('Fotos_' + comodatoNumero);

    // Guardamos archivos en Drive
    const urlsEquipos = savePhotos_(comodatoPhotosFolder, data.fotosEquipos, 'EQ_' + comodatoNumero);
    const urlsPilones = savePhotos_(comodatoPhotosFolder, data.fotosPilones, 'PIL_' + comodatoNumero);
    const signatureFile = saveSignature_(data.firmaDataUrl, comodatoNumero);

    // Generamos Documento PDF y DOC con fotos incrustadas
    const archivosGenerados = generatePdfFromTemplate_(data, comodatoNumero, signatureFile);
    const pdfFile = archivosGenerados.pdfFile;
    const docFile = archivosGenerados.docFile;

    // Armamos la fila del Excel
    const row = [
      new Date(), comodatoNumero, safe_(data.fecha), safe_(data.tecnico), safe_(data.distribuidor),
      safe_(data.codCliente), safe_(data.cuit), safe_(data.nombreFantasia), safe_(data.razonSocial),
      safe_(data.domicilio), safe_(data.localidad),
      txtEquipos,
      toNumber_(data.regCornelius), toNumber_(data.regMafridis), toNumber_(data.llaveMixta),
      toNumber_(data.cabezalMM), toNumber_(data.mangueraCerveza), toNumber_(data.mangueraPython),
      toNumber_(data.canillaNiquelada), toNumber_(data.canillaAgua), toNumber_(data.transformador),
      toNumber_(data.separadoresCanilla), toNumber_(data.conectorVasera), toNumber_(data.cantPicos),
      toNumber_(data.medallones), toNumber_(data.tuboGas), toNumber_(data.vaseraRinser),
      toNumber_(data.rinser), toNumber_(data.mangueraDesague),
      safe_(data.handle), safe_(data.celli), safe_(data.vasera),
      txtPilones,
      safe_(data.descripcion),
      urlsEquipos, urlsPilones,
      safe_(data.aclaracion), safe_(data.dni),
      signatureFile.getId(), signatureFile.getUrl(),
      docFile.getId(), docFile.getUrl(),
      pdfFile.getId(), pdfFile.getUrl(),
      data.aceptaTerminos === true ? 'SI' : 'NO'
    ];

    sheet.appendRow(row);

    return jsonOutput_({
      ok: true,
      message: 'Guardado con éxito',
      comodatoNumero: comodatoNumero,
      nombreFantasia: safe_(data.nombreFantasia),
      pdfUrl: pdfFile.getUrl()
    });
  } catch (error) {
    return jsonOutput_({ ok: false, message: error.message });
  }
}

/**
 * 4. LÓGICA DE PDF Y DOCS (PLANTILLA)
 */
function generatePdfFromTemplate_(data, comodatoNumero, signatureFile) {
  const templateFile = DriveApp.getFileById(CONFIG.TEMPLATE_DOC_ID);
  const pdfFolder = DriveApp.getFolderById(CONFIG.PDF_FOLDER_ID);

  const nombreFantasia = safe_(data.nombreFantasia).replace(/[/\\?%*:|"<>]/g, '-');
  const nombreArchivoPdf = 'Comodato_' + comodatoNumero + '_' + nombreFantasia;

  const tempCopy = templateFile.makeCopy('TMP_' + comodatoNumero, pdfFolder);
  const doc = DocumentApp.openById(tempCopy.getId());
  const body = doc.getBody();

  // Diccionario de reemplazos de texto
  const replacements = {
    '{{COMODATO_NUMERO}}': comodatoNumero,
    '{{FECHA}}': safe_(data.fecha),
    '{{TECNICO}}': safe_(data.tecnico),
    '{{DISTRIBUIDOR}}': safe_(data.distribuidor),
    '{{COD_CLIENTE}}': safe_(data.codCliente),
    '{{CUIT}}': safe_(data.cuit),
    '{{NOMBRE_FANTASIA}}': safe_(data.nombreFantasia),
    '{{RAZON_SOCIAL}}': safe_(data.razonSocial),
    '{{DOMICILIO}}': safe_(data.domicilio),
    '{{LOCALIDAD}}': safe_(data.localidad),
    '{{EQUIPOS_DETALLE}}': safe_(data.equiposDetalle),
    '{{PILONES_DETALLE}}': safe_(data.pilonesDetalle),
    '{{REG_CORNELIUS}}': String(toNumber_(data.regCornelius)),
    '{{REG_MAFRIDIS}}': String(toNumber_(data.regMafridis)),
    '{{LLAVE_MIXTA}}': String(toNumber_(data.llaveMixta)),
    '{{CABEZAL_MM}}': String(toNumber_(data.cabezalMM)),
    '{{MANGUERA_CERVEZA}}': String(toNumber_(data.mangueraCerveza)),
    '{{MANGUERA_PYTHON}}': String(toNumber_(data.mangueraPython)),
    '{{CANILLA_NIQUELADA}}': String(toNumber_(data.canillaNiquelada)),
    '{{CANILLA_AGUA}}': String(toNumber_(data.canillaAgua)),
    '{{TRANSFORMADOR}}': String(toNumber_(data.transformador)),
    '{{SEPARADORES_CANILLA}}': String(toNumber_(data.separadoresCanilla)),
    '{{CONECTOR_VASERA}}': String(toNumber_(data.conectorVasera)),
    '{{CANT_PICOS}}': String(toNumber_(data.cantPicos)),
    '{{MEDALLONES}}': String(toNumber_(data.medallones)),
    '{{TUBO_GAS}}': String(toNumber_(data.tuboGas)),
    '{{VASERA_RINSER}}': String(toNumber_(data.vaseraRinser)),
    '{{RINSER}}': String(toNumber_(data.rinser)),
    '{{MANGUERA_DESAGUE}}': String(toNumber_(data.mangueraDesague)),
    '{{HANDLE}}': safe_(data.handle),
    '{{CELLI}}': safe_(data.celli),
    '{{VASERA}}': safe_(data.vasera),
    '{{DESCRIPCION}}': safe_(data.descripcion),
    '{{ACLARACION}}': safe_(data.aclaracion),
    '{{DNI}}': safe_(data.dni),
    '{{ACEPTA_TERMINOS}}': (data.aceptaTerminos === true || data.aceptaTerminos === 'SI') ? 'SI' : 'NO'
  };

  // Reemplazamos todos los textos
  Object.keys(replacements).forEach(function(key) {
    body.replaceText(escapeForReplaceText_(key), replacements[key]);
  });

  // Pegamos la firma
  insertSignatureIntoDoc_(body, signatureFile);

  // Pegamos las fotos en el documento
  insertPhotosIntoDoc_(body, data.fotosEquipos, '\\{\\{FOTOS_EQUIPOS\\}\\}');
  insertPhotosIntoDoc_(body, data.fotosPilones, '\\{\\{FOTOS_PILONES\\}\\}');

  // Guardamos cambios
  doc.saveAndClose();

  // Generamos el PDF
  const pdfBlob = tempCopy.getBlob().getAs(MimeType.PDF).setName(nombreArchivoPdf + '.pdf');
  const pdfFile = pdfFolder.createFile(pdfBlob);

  // Renombramos y conservamos el DOC temporal
  tempCopy.setName(nombreArchivoPdf + '_DOC');

  return {
    pdfFile: pdfFile,
    docFile: tempCopy
  };
}

function generarPdfDesdeBoton() {
  const sheet = getSheet_();
  const data = sheet.getDataRange().getValues();

  const idxNum = HEADERS.indexOf('ComodatoNumero');
  const idxPdfUrl = HEADERS.indexOf('PdfUrl');
  const idxFirmaId = HEADERS.indexOf('FirmaFileId');
  const idxPdfId = HEADERS.indexOf('PdfFileId');
  const idxDocId = HEADERS.indexOf('DocFileId');
  const idxDocUrl = HEADERS.indexOf('DocUrl');

  let procesados = 0;

  for (let i = 1; i < data.length; i++) {
    const fila = data[i];

    if (fila[idxNum] && (!fila[idxPdfUrl] || String(fila[idxPdfUrl]).trim() === "")) {
      try {
        const payload = {};
        HEADERS.forEach((h, index) => {
          let key = h.charAt(0).toLowerCase() + h.slice(1);
          payload[key] = fila[index];
        });

        if (!fila[idxFirmaId]) throw new Error("No hay ID de firma.");
        const signatureFile = DriveApp.getFileById(fila[idxFirmaId]);

        // En este proceso manual (desde el botón), no tenemos los arrays de fotos en Base64
        // por lo que las etiquetas de fotos simplemente se borrarán o mostrarán un mensaje de falta de fotos.
        const archivosGenerados = generatePdfFromTemplate_(payload, fila[idxNum], signatureFile);

        sheet.getRange(i + 1, idxPdfId + 1).setValue(archivosGenerados.pdfFile.getId());
        sheet.getRange(i + 1, idxPdfUrl + 1).setValue(archivosGenerados.pdfFile.getUrl());
        sheet.getRange(i + 1, idxDocId + 1).setValue(archivosGenerados.docFile.getId());
        sheet.getRange(i + 1, idxDocUrl + 1).setValue(archivosGenerados.docFile.getUrl());

        procesados++;
      } catch (e) {
        console.error("Error en fila " + (i + 1) + ": " + e.message);
      }
    }
  }
  SpreadsheetApp.getUi().alert(procesados > 0 ? "Se generaron " + procesados + " PDFs y Docs." : "No hay pendientes.");
}

/**
 * 5. FUNCIONES AUXILIARES DE ARCHIVOS Y MULTIMEDIA
 */
function formatEquipos_(equiposArray) {
  if (!equiposArray || equiposArray.length === 0) return "No se registraron equipos.";
  return equiposArray.map((eq, i) =>
    `Equipo #${i + 1} -> Marca: ${eq.marca} | Modelo: ${eq.modelo} | AF: ${eq.activoFijo} | Serie: ${eq.serie}`
  ).join('\n');
}

function formatPilones_(pilonesArray) {
  if (!pilonesArray || pilonesArray.length === 0) return "No se registraron pilones.";
  return pilonesArray.map((pil, i) => {
    let tipo = pil.tipo === 'OTROS' ? `OTROS (${pil.otros})` : pil.tipo;
    return `Pilón #${i + 1} -> Tipo: ${tipo} | Cabezal: ${pil.cabezal}`;
  }).join('\n');
}

function savePhotos_(folder, photosArray, prefix) {
  if (!photosArray || photosArray.length === 0) return '';
  let urls = [];
  photosArray.forEach((photo, index) => {
    const base64 = photo.data;
    const blob = Utilities.newBlob(Utilities.base64Decode(base64), photo.tipo, prefix + '_' + (index + 1) + '_' + photo.nombre);
    const file = folder.createFile(blob);
    urls.push(file.getUrl());
  });
  return urls.join('\n');
}

function saveSignature_(dataUrl, comodatoNumero) {
  const folder = DriveApp.getFolderById(CONFIG.SIGNATURE_FOLDER_ID);
  const base64 = String(dataUrl).split(',')[1];
  const blob = Utilities.newBlob(Utilities.base64Decode(base64), 'image/png', 'Firma_' + comodatoNumero + '.png');
  return folder.createFile(blob);
}

function insertSignatureIntoDoc_(body, signatureFile) {
  const found = body.findText('\\{\\{FIRMA\\}\\}');
  if (!found) return;
  const text = found.getElement().asText();
  text.deleteText(found.getStartOffset(), found.getEndOffsetInclusive());
  const parent = text.getParent();
  if (parent.getType() === DocumentApp.ElementType.PARAGRAPH) {
    parent.asParagraph().appendInlineImage(signatureFile.getBlob()).setWidth(180);
  }
}

/**
 * Función para insertar un array de fotos en el documento de Google
 */
function insertPhotosIntoDoc_(body, photosData, regexTag) {
  const found = body.findText(regexTag);
  if (!found) return;

  const textElement = found.getElement().asText();
  const parent = textElement.getParent();

  // Borramos la etiqueta de texto
  textElement.deleteText(found.getStartOffset(), found.getEndOffsetInclusive());

  // Verificamos si hay fotos para pegar
  if (Array.isArray(photosData) && photosData.length > 0) {
    if (parent.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const paragraph = parent.asParagraph();

      photosData.forEach(photo => {
        try {
          const base64 = photo.data;
          const blob = Utilities.newBlob(Utilities.base64Decode(base64), photo.tipo, photo.nombre);
          const img = paragraph.appendInlineImage(blob);

          // Redimensionamos la imagen a 200px de ancho
          const ratio = 200 / img.getWidth();
          img.setWidth(200);
          img.setHeight(img.getHeight() * ratio);

          // Agregamos espacio
          paragraph.appendText('   ');
        } catch (e) {
          console.error("Error insertando imagen: " + e.message);
        }
      });
    }
  } else {
    // Mensaje si no hay fotos
    if (parent.getType() === DocumentApp.ElementType.PARAGRAPH) {
      parent.asParagraph().appendText('No se adjuntaron fotografías en este registro.');
    }
  }
}

/**
 * 6. FUNCIONES DE UTILIDAD GENERAL
 */
function getSheet_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error('No existe la hoja ' + CONFIG.SHEET_NAME);
  return sheet;
}

function getNextComodatoNumber_() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return formatComodatoNumber_(1);
  const values = sheet.getRange(2, 2, lastRow - 1, 1).getValues().flat();
  let max = 0;
  values.forEach(v => {
    const match = String(v).match(/(\d+)$/);
    if (match) max = Math.max(max, Number(match[1]));
  });
  return formatComodatoNumber_(max + 1);
}

function formatComodatoNumber_(num) {
  return CONFIG.NUMBER_PREFIX + Utilities.formatString('%0' + CONFIG.NUMBER_DIGITS + 'd', num);
}

/**
 * 7. LISTADO "MIS COMODATOS" POR TÉCNICO
 */
function getDistinctTecnicos_() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const idxTecnico = HEADERS.indexOf('Tecnico');
  const values = sheet.getRange(2, idxTecnico + 1, lastRow - 1, 1).getValues().flat();

  const vistos = {};
  const tecnicos = [];
  values.forEach(v => {
    const nombre = String(v).trim();
    if (nombre && !vistos[nombre.toLowerCase()]) {
      vistos[nombre.toLowerCase()] = true;
      tecnicos.push(nombre);
    }
  });

  return tecnicos.sort((a, b) => a.localeCompare(b, 'es'));
}

function getComodatosByTecnico_(tecnico) {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();

  const idx = {
    timestamp: HEADERS.indexOf('Timestamp'),
    comodatoNumero: HEADERS.indexOf('ComodatoNumero'),
    fecha: HEADERS.indexOf('Fecha'),
    tecnico: HEADERS.indexOf('Tecnico'),
    distribuidor: HEADERS.indexOf('Distribuidor'),
    nombreFantasia: HEADERS.indexOf('NombreFantasia'),
    razonSocial: HEADERS.indexOf('RazonSocial'),
    localidad: HEADERS.indexOf('Localidad'),
    domicilio: HEADERS.indexOf('Domicilio'),
    pdfUrl: HEADERS.indexOf('PdfUrl'),
    docUrl: HEADERS.indexOf('DocUrl')
  };

  const tecnicoLower = String(tecnico).trim().toLowerCase();
  const resultados = [];

  data.forEach(fila => {
    if (String(fila[idx.tecnico]).trim().toLowerCase() === tecnicoLower) {
      const ts = fila[idx.timestamp];
      resultados.push({
        comodatoNumero: safe_(fila[idx.comodatoNumero]),
        fecha: safe_(fila[idx.fecha]),
        distribuidor: safe_(fila[idx.distribuidor]),
        nombreFantasia: safe_(fila[idx.nombreFantasia]),
        razonSocial: safe_(fila[idx.razonSocial]),
        localidad: safe_(fila[idx.localidad]),
        domicilio: safe_(fila[idx.domicilio]),
        pdfUrl: safe_(fila[idx.pdfUrl]),
        docUrl: safe_(fila[idx.docUrl]),
        timestamp: ts instanceof Date ? ts.getTime() : 0
      });
    }
  });

  resultados.sort((a, b) => b.timestamp - a.timestamp);
  return resultados;
}

function escapeForReplaceText_(text) { return String(text).replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }
function toNumber_(v) { const n = Number(v); return isNaN(n) ? 0 : n; }
function safe_(v) { return v === undefined || v === null ? '' : String(v); }
function jsonOutput_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
