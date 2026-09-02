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

  // Opcional: si lo dejás vacío o con el texto de ejemplo, el script crea/reutiliza
  // automáticamente una carpeta "Fotos Comodatos" al lado de la carpeta de PDFs
  // y guarda su ID en las Propiedades del Script. Ver getPhotosParentFolder_().
  PHOTOS_PARENT_FOLDER_ID: '',
  PHOTOS_FOLDER_NAME: 'Fotos Comodatos',

  // --- MÓDULO SANITIZACIONES ---
  SANIT_SHEET_PREFIX: 'Sanit_',
  CLIENTES_SHEET_NAME: 'Clientes_Sanitizacion',
  SANIT_PHOTOS_FOLDER_NAME: 'Fotos Sanitizaciones',
  SANIT_CICLO_DIAS: 28,
  SANIT_AVISO_DIAS: 7,
  TECNICOS: [
    'Gaston del Rio',
    'Federico Barbutti',
    'Maximiliano Di Pietro',
    'Mariano Diaz',
    'Jose Caporaletti',
    'Ramon Lazarte'
  ]
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

// Encabezados de las hojas de sanitización (una por técnico)
const SANIT_HEADERS = [
  'SanitizacionId', 'Tecnico', 'Cliente', 'Direccion', 'Localidad',
  'ComodatoNumero', 'Origen',
  'CheckIn', 'CheckOut', 'MinutosEnPdv',
  'Observaciones', 'FotosUrls', 'Estado', 'ProximaSanitizacion', 'Timestamp'
];

// Encabezados del padrón manual de clientes a sanitizar
const CLIENTES_HEADERS = [
  'Tecnico', 'Cliente', 'Direccion', 'Localidad', 'Activo', 'FechaAlta',
  'Equipo', 'Pilon', 'CantPicos'
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

    if (action === 'listTecnicosFijos') {
      return jsonOutput_({ ok: true, tecnicos: CONFIG.TECNICOS });
    }

    if (action === 'sanitCartera') {
      const t = resolveTecnico_(e.parameter.tecnico);
      return jsonOutput_({ ok: true, tecnico: t, cartera: getCarteraSanitizacion_(t) });
    }

    if (action === 'sanitResumenGeneral') {
      return jsonOutput_({ ok: true, general: getResumenGeneral_() });
    }

    if (action === 'sanitResumen') {
      const t = resolveTecnico_(e.parameter.tecnico);
      return jsonOutput_({ ok: true, tecnico: t, resumen: getResumenSanitizacion_(t) });
    }

    if (action === 'sanitHistorial') {
      const t = resolveTecnico_(e.parameter.tecnico);
      return jsonOutput_({ ok: true, tecnico: t, historial: getHistorialSanitizacion_(t) });
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
    if (!e || !e.postData || !e.postData.contents) throw new Error('No llegaron datos.');

    const body = JSON.parse(e.postData.contents);

    // Enrutado por acción. Sin "action" se asume el alta de comodato (compatibilidad).
    switch (body.action) {
      case 'sanitCheckIn':    return jsonOutput_(sanitCheckIn_(body));
      case 'sanitCheckOut':   return jsonOutput_(sanitCheckOut_(body));
      case 'sanitAltaCliente':return jsonOutput_(sanitAltaCliente_(body));
      case 'sanitBajaCliente':return jsonOutput_(sanitBajaCliente_(body));
    }

    const sheet = getSheet_();
    const data = body;
    const comodatoNumero = esNumeroComodatoValido_(data.comodatoNumero) ? String(data.comodatoNumero).trim() : getNextComodatoNumber_();

    // Convertimos los arrays de equipos y pilones a texto
    const txtEquipos = formatEquipos_(data.equipos);
    const txtPilones = formatPilones_(data.pilones);
    data.equiposDetalle = txtEquipos;
    data.pilonesDetalle = txtPilones;

    // Creamos (o reutilizamos) la carpeta de fotos de este comodato
    const comodatoPhotosFolder = getComodatoPhotosFolder_(comodatoNumero);

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
 * Valida que el numero recibido tenga la forma TCC-00000. Evita que un
 * placeholder del formulario ("Cargando...", vacio, "Error") termine
 * guardado como numero de comodato.
 */
function esNumeroComodatoValido_(valor) {
  const txt = safe_(valor).trim();
  const prefijo = safe_(CONFIG.NUMBER_PREFIX);
  if (!txt || txt.indexOf(prefijo) !== 0) return false;
  const resto = txt.slice(prefijo.length);
  if (!resto.length) return false;
  for (let i = 0; i < resto.length; i++) {
    if (resto[i] < '0' || resto[i] > '9') return false;
  }
  return true;
}

/**
 * 3.b CARPETA DE FOTOS (resolución automática)
 *
 * Orden de resolución:
 *   1) ID guardado en Propiedades del Script (PHOTOS_PARENT_FOLDER_ID)
 *   2) CONFIG.PHOTOS_PARENT_FOLDER_ID, si es un ID real y accesible
 *   3) Carpeta "Fotos Comodatos" dentro de la carpeta que contiene los PDFs
 *      (se busca por nombre y, si no existe, se crea una sola vez)
 * El ID resuelto queda cacheado en Propiedades del Script.
 */
function getPhotosParentFolder_() {
  const props = PropertiesService.getScriptProperties();
  const candidatos = [props.getProperty('PHOTOS_PARENT_FOLDER_ID'), CONFIG.PHOTOS_PARENT_FOLDER_ID];

  for (let i = 0; i < candidatos.length; i++) {
    const id = candidatos[i] ? String(candidatos[i]).trim() : '';
    if (!id || id.indexOf('PEGAR_AQUI') === 0) continue;
    try {
      return DriveApp.getFolderById(id);
    } catch (err) {
      // ID inválido o sin permisos: seguimos con el siguiente candidato
    }
  }

  const pdfFolder = DriveApp.getFolderById(CONFIG.PDF_FOLDER_ID);
  const parents = pdfFolder.getParents();
  const base = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();

  const nombre = CONFIG.PHOTOS_FOLDER_NAME || 'Fotos Comodatos';
  const existentes = base.getFoldersByName(nombre);
  const folder = existentes.hasNext() ? existentes.next() : base.createFolder(nombre);

  props.setProperty('PHOTOS_PARENT_FOLDER_ID', folder.getId());
  return folder;
}

/** Devuelve la subcarpeta Fotos_<numero>, reutilizándola si ya existe. */
function getComodatoPhotosFolder_(comodatoNumero) {
  const parent = getPhotosParentFolder_();
  const nombre = 'Fotos_' + comodatoNumero;
  const existentes = parent.getFoldersByName(nombre);
  return existentes.hasNext() ? existentes.next() : parent.createFolder(nombre);
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

/**
 * Lee la hoja de comodatos UNA sola vez y devuelve un mapa
 * { tecnicoNormalizado: [comodato, ...] }, cada lista ya ordenada de la mas
 * reciente a la mas vieja. Evita releer la hoja seis veces al armar el
 * resumen general.
 */
function leerComodatosAgrupados_() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  const mapa = {};
  if (lastRow < 2) return mapa;

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
    docUrl: HEADERS.indexOf('DocUrl'),
    equiposDetalle: HEADERS.indexOf('EquiposDetalle'),
    pilonesDetalle: HEADERS.indexOf('PilonesDetalle'),
    cantPicos: HEADERS.indexOf('CantPicos')
  };

  data.forEach(fila => {
    const key = norm_(fila[idx.tecnico]);
    if (!key) return;
    const ts = fila[idx.timestamp];
    if (!mapa[key]) mapa[key] = [];
    mapa[key].push({
      comodatoNumero: safe_(fila[idx.comodatoNumero]),
      fecha: safe_(fila[idx.fecha]),
      distribuidor: safe_(fila[idx.distribuidor]),
      nombreFantasia: safe_(fila[idx.nombreFantasia]),
      razonSocial: safe_(fila[idx.razonSocial]),
      localidad: safe_(fila[idx.localidad]),
      domicilio: safe_(fila[idx.domicilio]),
      pdfUrl: safe_(fila[idx.pdfUrl]),
      docUrl: safe_(fila[idx.docUrl]),
      equipo: resumirEquipos_(fila[idx.equiposDetalle]),
      pilon: resumirPilones_(fila[idx.pilonesDetalle]),
      cantPicos: toNumber_(fila[idx.cantPicos]),
      timestamp: ts instanceof Date ? ts.getTime() : 0
    });
  });

  Object.keys(mapa).forEach(k => mapa[k].sort((a, b) => b.timestamp - a.timestamp));
  return mapa;
}

function getComodatosByTecnico_(tecnico) {
  return leerComodatosAgrupados_()[norm_(tecnico)] || [];
}

/**
 * Lee el padron manual UNA sola vez, agrupado por tecnico y ya filtrado por
 * los clientes activos.
 */
function leerClientesManualesAgrupados_() {
  const sheet = getClientesSheet_();
  const lastRow = sheet.getLastRow();
  const mapa = {};
  if (lastRow < 2) return mapa;

  const rows = sheet.getRange(2, 1, lastRow - 1, CLIENTES_HEADERS.length).getValues();
  rows.forEach(f => {
    const key = norm_(f[0]);
    if (!key || norm_(f[4]) === 'no') return;
    if (!mapa[key]) mapa[key] = [];
    mapa[key].push(f);
  });
  return mapa;
}


function escapeForReplaceText_(text) { return String(text).replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }
function toNumber_(v) { const n = Number(v); return isNaN(n) ? 0 : n; }
function safe_(v) { return v === undefined || v === null ? '' : String(v); }
function jsonOutput_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }

/* ==========================================================================
 * 7. MÓDULO DE SANITIZACIONES (ciclo de 28 días)
 *
 * Modelo de datos:
 *   - Una hoja por técnico: "Sanit_<Nombre del técnico>" con SANIT_HEADERS.
 *   - Una hoja común "Clientes_Sanitizacion" con el padrón cargado a mano.
 *   - La cartera de cada técnico = clientes de sus comodatos + padrón manual.
 *
 * Ciclo: la próxima sanitización de un cliente es la fecha de su última
 * sanitización completada + CONFIG.SANIT_CICLO_DIAS. Si nunca se sanitizó,
 * se usa la fecha del comodato como punto de partida.
 * ========================================================================== */

/** Normaliza texto para comparar (sin acentos, sin mayúsculas, sin espacios dobles). */
function norm_(v) {
  return safe_(v)
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toLowerCase().replace(/\s+/g, ' ').trim();
}

/** Devuelve el nombre canónico del técnico según CONFIG.TECNICOS. Lanza si no existe. */
function resolveTecnico_(tecnico) {
  const objetivo = norm_(tecnico);
  if (!objetivo) throw new Error('Falta el técnico.');
  for (let i = 0; i < CONFIG.TECNICOS.length; i++) {
    if (norm_(CONFIG.TECNICOS[i]) === objetivo) return CONFIG.TECNICOS[i];
  }
  throw new Error('Técnico no reconocido: ' + tecnico);
}

/** Hoja de sanitizaciones del técnico; la crea con encabezados si no existe. */
function getSanitSheet_(tecnico) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nombre = CONFIG.SANIT_SHEET_PREFIX + tecnico;
  let sheet = ss.getSheetByName(nombre);
  if (!sheet) {
    sheet = ss.insertSheet(nombre);
    sheet.getRange(1, 1, 1, SANIT_HEADERS.length).setValues([SANIT_HEADERS]);
    const header = sheet.getRange(1, 1, 1, SANIT_HEADERS.length);
    header.setBackground('#003366').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/** Hoja del padrón manual de clientes; la crea si no existe. */
function getClientesSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.CLIENTES_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.CLIENTES_SHEET_NAME);
  }
  escribirCabecera_(sheet, CLIENTES_HEADERS);
  return sheet;
}

/**
 * Deja la fila 1 con los encabezados esperados. Si la hoja ya existia con
 * menos columnas (por ejemplo antes de sumar los datos de chopera), agrega
 * las que faltan sin tocar los datos ya cargados.
 */
function escribirCabecera_(sheet, headers) {
  const anchoActual = sheet.getLastColumn();
  if (anchoActual >= headers.length && sheet.getLastRow() >= 1) {
    const actuales = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
    let iguales = true;
    for (let i = 0; i < headers.length; i++) {
      if (safe_(actuales[i]).trim() !== headers[i]) { iguales = false; break; }
    }
    if (iguales) return sheet;
  }

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  const header = sheet.getRange(1, 1, 1, headers.length);
  header.setBackground('#003366').setFontColor('#ffffff').setFontWeight('bold');
  sheet.setFrozenRows(1);
  return sheet;
}

/** Crea de una las 6 hojas de técnico + el padrón. Se corre a mano desde el editor. */
function crearHojasSanitizacion() {
  CONFIG.TECNICOS.forEach(t => getSanitSheet_(t));
  getClientesSheet_();
  Logger.log('Hojas de sanitización listas.');
}

/** Carpeta de fotos de sanitizaciones, con subcarpeta por técnico. */
function getSanitPhotosFolder_(tecnico) {
  const base = getPhotosParentFolder_();
  const raiz = base.getFoldersByName(CONFIG.SANIT_PHOTOS_FOLDER_NAME);
  const parent = raiz.hasNext() ? raiz.next() : base.createFolder(CONFIG.SANIT_PHOTOS_FOLDER_NAME);
  const sub = parent.getFoldersByName(tecnico);
  return sub.hasNext() ? sub.next() : parent.createFolder(tecnico);
}

/**
 * Resume EquiposDetalle a algo corto y legible en la tarjeta del tecnico.
 * "Equipo #1 -> Marca: Celli | Modelo: T4 | AF: ... | Serie: ..." -> "Celli T4".
 * Si el texto no tiene el formato esperado se devuelve tal cual.
 */
function resumirEquipos_(texto) {
  const crudo = safe_(texto).trim();
  if (!crudo || crudo.indexOf('No se registraron') === 0) return '';

  const partes = crudo.split('\n').map(linea => {
    const marca = linea.match(/Marca:\s*([^|]*)/);
    const modelo = linea.match(/Modelo:\s*([^|]*)/);
    if (!marca && !modelo) return linea.trim();
    return [marca ? marca[1].trim() : '', modelo ? modelo[1].trim() : ''].filter(String).join(' ');
  }).filter(String);

  return partes.join(' · ');
}

/** Resume PilonesDetalle: se queda con el tipo de cada pilon. */
function resumirPilones_(texto) {
  const crudo = safe_(texto).trim();
  if (!crudo || crudo.indexOf('No se registraron') === 0) return '';

  const partes = crudo.split('\n').map(linea => {
    const tipo = linea.match(/Tipo:\s*([^|]*)/);
    return tipo ? tipo[1].trim() : linea.trim();
  }).filter(String);

  return partes.join(' · ');
}

/** Lee todas las filas de la hoja de un técnico como objetos. */
function readSanitRows_(tecnico) {
  const sheet = getSanitSheet_(tecnico);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, SANIT_HEADERS.length).getValues();
  return values.map((fila, i) => {
    const obj = { _row: i + 2 };
    SANIT_HEADERS.forEach((h, c) => { obj[h] = fila[c]; });
    return obj;
  });
}

function toDate_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  if (!v) return null;
  const d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

function addDays_(date, days) {
  const d = new Date(date.getTime());
  d.setDate(d.getDate() + days);
  return d;
}

function diffDays_(desde, hasta) {
  const MS = 24 * 60 * 60 * 1000;
  const a = new Date(desde.getFullYear(), desde.getMonth(), desde.getDate());
  const b = new Date(hasta.getFullYear(), hasta.getMonth(), hasta.getDate());
  return Math.round((b.getTime() - a.getTime()) / MS);
}

function isoDate_(d) {
  return d ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
}

function isoDateTime_(d) {
  return d ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : '';
}

/**
 * Cartera de clientes del técnico: comodatos cargados + padrón manual,
 * cruzada contra su historial de sanitizaciones.
 */
function getCarteraSanitizacion_(tecnico) {
  return construirCartera_(
    tecnico,
    leerComodatosAgrupados_()[norm_(tecnico)] || [],
    leerClientesManualesAgrupados_()[norm_(tecnico)] || [],
    readSanitRows_(tecnico)
  );
}

/**
 * Arma la cartera a partir de datos ya leidos. Separada de la carga para que
 * el resumen general pueda reutilizarla sin releer las hojas por cada tecnico.
 */
function construirCartera_(tecnico, comodatos, clientesManuales, sanitRows) {
  const mapa = {};

  const push_ = (cliente, direccion, localidad, origen, comodatoNumero, fechaBase, chopera) => {
    const key = norm_(cliente);
    if (!key) return;
    const eq = chopera || {};

    if (!mapa[key]) {
      mapa[key] = {
        key: key,
        cliente: safe_(cliente).trim(),
        direccion: safe_(direccion).trim(),
        localidad: safe_(localidad).trim(),
        origen: origen,
        comodatoNumero: safe_(comodatoNumero),
        fechaBase: fechaBase || null,
        equipo: safe_(eq.equipo).trim(),
        pilon: safe_(eq.pilon).trim(),
        cantPicos: toNumber_(eq.cantPicos)
      };
      return;
    }
    // Completar huecos y quedarse con la fecha base más reciente
    const actual = mapa[key];
    if (!actual.direccion) actual.direccion = safe_(direccion).trim();
    if (!actual.localidad) actual.localidad = safe_(localidad).trim();
    if (!actual.comodatoNumero) actual.comodatoNumero = safe_(comodatoNumero);
    if (!actual.equipo) actual.equipo = safe_(eq.equipo).trim();
    if (!actual.pilon) actual.pilon = safe_(eq.pilon).trim();
    if (!actual.cantPicos) actual.cantPicos = toNumber_(eq.cantPicos);
    if (fechaBase && (!actual.fechaBase || fechaBase > actual.fechaBase)) actual.fechaBase = fechaBase;
  };

  // A) Clientes que salen de los comodatos cargados
  comodatos.forEach(c => {
    push_(c.nombreFantasia || c.razonSocial, c.domicilio, c.localidad, 'COMODATO',
          c.comodatoNumero, toDate_(c.fecha) || (c.timestamp ? new Date(c.timestamp) : null),
          { equipo: c.equipo, pilon: c.pilon, cantPicos: c.cantPicos });
  });

  // B) Clientes del padrón manual
  clientesManuales.forEach(f => {
    push_(f[1], f[2], f[3], 'MANUAL', '', toDate_(f[5]),
          { equipo: f[6], pilon: f[7], cantPicos: f[8] });
  });

  // C) Cruce con el historial de sanitizaciones
  const ultima = {};
  const abierta = {};
  sanitRows.forEach(r => {
    const key = norm_(r.Cliente);
    if (!key) return;
    if (!mapa[key]) push_(r.Cliente, r.Direccion, r.Localidad, safe_(r.Origen) || 'MANUAL', r.ComodatoNumero, null, null);
    if (norm_(r.Estado) === 'en curso') {
      abierta[key] = { sanitizacionId: safe_(r.SanitizacionId), checkIn: toDate_(r.CheckIn) };
      return;
    }
    const fin = toDate_(r.CheckOut);
    if (fin && (!ultima[key] || fin > ultima[key])) ultima[key] = fin;
  });

  const hoy = new Date();
  const ciclo = CONFIG.SANIT_CICLO_DIAS;
  const aviso = CONFIG.SANIT_AVISO_DIAS;

  return Object.keys(mapa).map(key => {
    const c = mapa[key];
    const ult = ultima[key] || null;
    const base = ult || c.fechaBase;
    const prox = base ? addDays_(base, ciclo) : null;
    const dias = prox ? diffDays_(hoy, prox) : null;

    let estado;
    if (abierta[key]) estado = 'EN CURSO';
    else if (!prox) estado = 'SIN REGISTRO';
    else if (dias < 0) estado = 'VENCIDO';
    else if (dias <= aviso) estado = 'POR VENCER';
    else estado = 'AL DIA';

    return {
      cliente: c.cliente,
      direccion: c.direccion,
      localidad: c.localidad,
      origen: c.origen,
      comodatoNumero: c.comodatoNumero,
      equipo: c.equipo,
      pilon: c.pilon,
      cantPicos: c.cantPicos,
      ultimaSanitizacion: isoDate_(ult),
      proximaSanitizacion: isoDate_(prox),
      diasRestantes: dias,
      estado: estado,
      sanitizacionAbiertaId: abierta[key] ? abierta[key].sanitizacionId : '',
      checkInAbierto: abierta[key] ? isoDateTime_(abierta[key].checkIn) : ''
    };
  }).sort((a, b) => {
    const orden = { 'EN CURSO': 0, 'VENCIDO': 1, 'POR VENCER': 2, 'SIN REGISTRO': 3, 'AL DIA': 4 };
    if (orden[a.estado] !== orden[b.estado]) return orden[a.estado] - orden[b.estado];
    if (a.diasRestantes === null) return 1;
    if (b.diasRestantes === null) return -1;
    return a.diasRestantes - b.diasRestantes;
  });
}

/** Resumen de indicadores del técnico. */
function getResumenSanitizacion_(tecnico) {
  return calcularResumen_(getCarteraSanitizacion_(tecnico), readSanitRows_(tecnico));
}

/** Indicadores a partir de una cartera y un historial ya cargados. */
function calcularResumen_(cartera, filas) {
  const hoy = new Date();

  let completadas30 = 0, minutosTotal = 0, conMinutos = 0;
  filas.forEach(r => {
    if (norm_(r.Estado) !== 'completada') return;
    const fin = toDate_(r.CheckOut);
    if (fin && diffDays_(fin, hoy) <= 30) completadas30++;
    const m = Number(r.MinutosEnPdv);
    if (!isNaN(m) && m > 0) { minutosTotal += m; conMinutos++; }
  });

  const contar = estado => cartera.filter(c => c.estado === estado).length;

  return {
    totalClientes: cartera.length,
    alDia: contar('AL DIA'),
    porVencer: contar('POR VENCER'),
    vencidos: contar('VENCIDO'),
    sinRegistro: contar('SIN REGISTRO'),
    enCurso: contar('EN CURSO'),
    completadasUltimos30: completadas30,
    totalSanitizaciones: filas.filter(r => norm_(r.Estado) === 'completada').length,
    minutosPromedioPdv: conMinutos ? Math.round(minutosTotal / conMinutos) : 0
  };
}

/**
 * Resumen consolidado de los seis tecnicos, para cuando no hay ninguno
 * seleccionado. Lee la hoja de comodatos y el padron una sola vez y despues
 * arma la cartera de cada tecnico en memoria.
 */
function getResumenGeneral_() {
  const comodatos = leerComodatosAgrupados_();
  const clientes = leerClientesManualesAgrupados_();

  const totales = {
    totalClientes: 0, alDia: 0, porVencer: 0, vencidos: 0, sinRegistro: 0,
    enCurso: 0, completadasUltimos30: 0, totalSanitizaciones: 0, minutosPromedioPdv: 0
  };

  let minutosAcumulados = 0, sanitConMinutos = 0;

  const porTecnico = CONFIG.TECNICOS.map(tecnico => {
    const key = norm_(tecnico);
    const filas = readSanitRows_(tecnico);
    const cartera = construirCartera_(tecnico, comodatos[key] || [], clientes[key] || [], filas);
    const r = calcularResumen_(cartera, filas);

    totales.totalClientes += r.totalClientes;
    totales.alDia += r.alDia;
    totales.porVencer += r.porVencer;
    totales.vencidos += r.vencidos;
    totales.sinRegistro += r.sinRegistro;
    totales.enCurso += r.enCurso;
    totales.completadasUltimos30 += r.completadasUltimos30;
    totales.totalSanitizaciones += r.totalSanitizaciones;

    // El promedio general se pondera por sanitizacion, no por tecnico
    filas.forEach(f => {
      const m = Number(f.MinutosEnPdv);
      if (norm_(f.Estado) === 'completada' && !isNaN(m) && m > 0) {
        minutosAcumulados += m;
        sanitConMinutos++;
      }
    });

    // El PDV mas atrasado del tecnico, para saber que tan critico esta
    let peor = null;
    cartera.forEach(c => {
      if (c.diasRestantes === null || c.diasRestantes === undefined) return;
      if (peor === null || c.diasRestantes < peor) peor = c.diasRestantes;
    });

    return {
      tecnico: tecnico,
      totalClientes: r.totalClientes,
      alDia: r.alDia,
      porVencer: r.porVencer,
      vencidos: r.vencidos,
      sinRegistro: r.sinRegistro,
      enCurso: r.enCurso,
      completadasUltimos30: r.completadasUltimos30,
      minutosPromedioPdv: r.minutosPromedioPdv,
      diasMasAtrasado: peor
    };
  });

  totales.minutosPromedioPdv = sanitConMinutos ? Math.round(minutosAcumulados / sanitConMinutos) : 0;

  // Primero los que mas urgencia tienen
  porTecnico.sort((a, b) => {
    if (b.vencidos !== a.vencidos) return b.vencidos - a.vencidos;
    return b.porVencer - a.porVencer;
  });

  return { totales: totales, porTecnico: porTecnico };
}

/** Historial completo de sanitizaciones del técnico, de la más reciente a la más vieja. */
function getHistorialSanitizacion_(tecnico) {
  return readSanitRows_(tecnico).map(r => ({
    sanitizacionId: safe_(r.SanitizacionId),
    cliente: safe_(r.Cliente),
    direccion: safe_(r.Direccion),
    localidad: safe_(r.Localidad),
    checkIn: isoDateTime_(toDate_(r.CheckIn)),
    checkOut: isoDateTime_(toDate_(r.CheckOut)),
    minutosEnPdv: Number(r.MinutosEnPdv) || 0,
    observaciones: safe_(r.Observaciones),
    fotosUrls: safe_(r.FotosUrls),
    estado: safe_(r.Estado),
    proximaSanitizacion: isoDate_(toDate_(r.ProximaSanitizacion)),
    _orden: toDate_(r.CheckIn) ? toDate_(r.CheckIn).getTime() : 0
  })).sort((a, b) => b._orden - a._orden);
}

/** Registra la llegada del técnico al PDV. */
function sanitCheckIn_(body) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const tecnico = resolveTecnico_(body.tecnico);
    const cliente = safe_(body.cliente).trim();
    if (!cliente) throw new Error('Falta el cliente.');

    // No permitir dos check-in abiertos del mismo cliente
    const abiertaPrevia = readSanitRows_(tecnico).filter(
      r => norm_(r.Estado) === 'en curso' && norm_(r.Cliente) === norm_(cliente)
    );
    if (abiertaPrevia.length) {
      return {
        ok: false,
        message: 'Ya hay una sanitización en curso para este cliente. Cerrala con el check-out.',
        sanitizacionId: safe_(abiertaPrevia[0].SanitizacionId)
      };
    }

    const ahora = new Date();
    const id = 'SAN-' + ahora.getTime().toString(36).toUpperCase();

    getSanitSheet_(tecnico).appendRow([
      id, tecnico, cliente, safe_(body.direccion), safe_(body.localidad),
      safe_(body.comodatoNumero), safe_(body.origen) || 'MANUAL',
      ahora, '', '',
      '', '', 'EN CURSO', '', ahora
    ]);

    return { ok: true, sanitizacionId: id, checkIn: isoDateTime_(ahora), message: 'Check-in registrado.' };
  } finally {
    lock.releaseLock();
  }
}

/** Cierra la sanitización: hora de salida, tiempo en PDV, fotos y observaciones. */
function sanitCheckOut_(body) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const tecnico = resolveTecnico_(body.tecnico);
    const id = safe_(body.sanitizacionId).trim();
    if (!id) throw new Error('Falta el ID de la sanitización.');

    const sheet = getSanitSheet_(tecnico);
    const fila = readSanitRows_(tecnico).filter(r => safe_(r.SanitizacionId) === id)[0];
    if (!fila) throw new Error('No se encontró la sanitización ' + id);
    if (norm_(fila.Estado) === 'completada') throw new Error('Esa sanitización ya está cerrada.');

    const inicio = toDate_(fila.CheckIn) || new Date();
    const fin = new Date();
    const minutos = Math.max(1, Math.round((fin.getTime() - inicio.getTime()) / 60000));
    const proxima = addDays_(fin, CONFIG.SANIT_CICLO_DIAS);

    let urls = '';
    if (body.fotos && body.fotos.length) {
      urls = savePhotos_(getSanitPhotosFolder_(tecnico), body.fotos, id);
    }

    const col = h => SANIT_HEADERS.indexOf(h) + 1;
    sheet.getRange(fila._row, col('CheckOut')).setValue(fin);
    sheet.getRange(fila._row, col('MinutosEnPdv')).setValue(minutos);
    sheet.getRange(fila._row, col('Observaciones')).setValue(safe_(body.observaciones));
    sheet.getRange(fila._row, col('FotosUrls')).setValue(urls);
    sheet.getRange(fila._row, col('Estado')).setValue('COMPLETADA');
    sheet.getRange(fila._row, col('ProximaSanitizacion')).setValue(proxima);

    return {
      ok: true,
      message: 'Sanitización cerrada.',
      sanitizacionId: id,
      minutosEnPdv: minutos,
      checkOut: isoDateTime_(fin),
      proximaSanitizacion: isoDate_(proxima)
    };
  } finally {
    lock.releaseLock();
  }
}

/** Alta manual de un cliente en el padrón de sanitización. */
function sanitAltaCliente_(body) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const tecnico = resolveTecnico_(body.tecnico);
    const cliente = safe_(body.cliente).trim();
    if (!cliente) throw new Error('Falta el nombre del cliente.');

    const sheet = getClientesSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const rows = sheet.getRange(2, 1, lastRow - 1, CLIENTES_HEADERS.length).getValues();
      const duplicado = rows.some(f => norm_(f[0]) === norm_(tecnico) && norm_(f[1]) === norm_(cliente));
      if (duplicado) return { ok: false, message: 'Ese cliente ya está en tu cartera.' };
    }

    const cantPicos = toNumber_(body.cantPicos);
    sheet.appendRow([
      tecnico, cliente, safe_(body.direccion), safe_(body.localidad), 'SI', new Date(),
      safe_(body.equipo).trim(), safe_(body.pilon).trim(), cantPicos
    ]);
    return {
      ok: true,
      message: 'Cliente agregado a la cartera.',
      equipo: safe_(body.equipo).trim(),
      pilon: safe_(body.pilon).trim(),
      cantPicos: cantPicos
    };
  } finally {
    lock.releaseLock();
  }
}

/** Baja lógica de un cliente del padrón manual (Activo = NO). */
function sanitBajaCliente_(body) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const tecnico = resolveTecnico_(body.tecnico);
    const cliente = safe_(body.cliente).trim();
    const sheet = getClientesSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { ok: false, message: 'No hay clientes cargados a mano.' };

    const rows = sheet.getRange(2, 1, lastRow - 1, CLIENTES_HEADERS.length).getValues();
    for (let i = 0; i < rows.length; i++) {
      if (norm_(rows[i][0]) === norm_(tecnico) && norm_(rows[i][1]) === norm_(cliente)) {
        sheet.getRange(i + 2, CLIENTES_HEADERS.indexOf('Activo') + 1).setValue('NO');
        return { ok: true, message: 'Cliente dado de baja de la cartera.' };
      }
    }
    return { ok: false, message: 'Ese cliente no está en el padrón manual (viene de un comodato).' };
  } finally {
    lock.releaseLock();
  }
}

/* ==========================================================================
 * 8. LIMPIEZA DE LOS DATOS DE PRUEBA (uso unico)
 *
 * Borra exactamente los registros creados durante la prueba del 02/09/2026:
 *   - Comodato TCC-00092 (fila + PDF + Doc + firma + carpeta de fotos)
 *   - Sanitizacion SAN-MTKG67C6 en la hoja Sanit_Gaston del Rio (fila + foto)
 *   - Cliente "ZZZ TEST CLAUDE BORRAR" del padron Clientes_Sanitizacion
 *
 * Ejecutar UNA sola vez desde el editor (Ejecutar -> limpiarPruebasClaude).
 * No hace falta redeployar: las funciones del editor corren sobre el codigo
 * guardado, no sobre la implementacion publicada.
 *
 * Los archivos de Drive van a la papelera (setTrashed), no se destruyen: si
 * algo se borra de mas, se recupera desde la papelera.
 *
 * Cada borrado verifica primero que la fila sea realmente la de prueba. Si no
 * coincide, la saltea y lo informa en el log. Una vez ejecutada, esta seccion
 * se puede eliminar del archivo.
 * ========================================================================== */

const PRUEBA_CLAUDE = {
  MARCA: 'ZZZ TEST CLAUDE BORRAR',
  COMODATO: 'TCC-00092',
  SANITIZACION: 'SAN-MTKG67C6',
  TECNICO: 'Gaston del Rio'
};

/** Manda un archivo de Drive a la papelera, sin romper si ya no existe. */
function papeleraPorId_(id, etiqueta, log) {
  const limpio = safe_(id).trim();
  if (!limpio) return;
  try {
    DriveApp.getFileById(limpio).setTrashed(true);
    log.push('  papelera: ' + etiqueta + ' (' + limpio + ')');
  } catch (err) {
    log.push('  no se pudo mover ' + etiqueta + ': ' + err.message);
  }
}

/** Extrae el ID de archivo de una URL de Drive. */
function idDesdeUrlDrive_(url) {
  const m = safe_(url).match(/[-\w]{25,}/);
  return m ? m[0] : '';
}

function limpiarPruebasClaude() {
  const log = [];

  // 1) Comodato de prueba
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  let borroComodato = false;

  if (lastRow >= 2) {
    const datos = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
    const iNum = HEADERS.indexOf('ComodatoNumero');
    const iNombre = HEADERS.indexOf('NombreFantasia');

    for (let i = datos.length - 1; i >= 0; i--) {
      const fila = datos[i];
      if (safe_(fila[iNum]).trim() !== PRUEBA_CLAUDE.COMODATO) continue;

      if (safe_(fila[iNombre]).trim() !== PRUEBA_CLAUDE.MARCA) {
        log.push('OJO: ' + PRUEBA_CLAUDE.COMODATO + ' no es la fila de prueba (cliente: ' +
                 safe_(fila[iNombre]) + '). No se toca.');
        continue;
      }

      log.push('Comodato ' + PRUEBA_CLAUDE.COMODATO + ':');
      papeleraPorId_(fila[HEADERS.indexOf('PdfFileId')], 'PDF', log);
      papeleraPorId_(fila[HEADERS.indexOf('DocFileId')], 'Doc', log);
      papeleraPorId_(fila[HEADERS.indexOf('FirmaFileId')], 'firma', log);

      // Carpeta de fotos del comodato
      try {
        const carpetas = getPhotosParentFolder_().getFoldersByName('Fotos_' + PRUEBA_CLAUDE.COMODATO);
        while (carpetas.hasNext()) {
          const f = carpetas.next();
          f.setTrashed(true);
          log.push('  papelera: carpeta ' + f.getName());
        }
      } catch (err) {
        log.push('  no se pudo mover la carpeta de fotos: ' + err.message);
      }

      sheet.deleteRow(i + 2);
      log.push('  fila borrada de ' + CONFIG.SHEET_NAME);
      borroComodato = true;
      break;
    }
  }
  if (!borroComodato) log.push('Comodato ' + PRUEBA_CLAUDE.COMODATO + ': no se encontro (ya estaba limpio).');

  // 2) Sanitizacion de prueba
  const sanitSheet = getSanitSheet_(PRUEBA_CLAUDE.TECNICO);
  const filaSanit = readSanitRows_(PRUEBA_CLAUDE.TECNICO).filter(
    r => safe_(r.SanitizacionId) === PRUEBA_CLAUDE.SANITIZACION
  )[0];

  if (filaSanit && safe_(filaSanit.Cliente).trim() === PRUEBA_CLAUDE.MARCA) {
    log.push('Sanitizacion ' + PRUEBA_CLAUDE.SANITIZACION + ':');
    safe_(filaSanit.FotosUrls).split('\n').filter(String).forEach((url, i) => {
      papeleraPorId_(idDesdeUrlDrive_(url), 'foto ' + (i + 1), log);
    });
    sanitSheet.deleteRow(filaSanit._row);
    log.push('  fila borrada de ' + sanitSheet.getName());
  } else {
    log.push('Sanitizacion ' + PRUEBA_CLAUDE.SANITIZACION + ': no se encontro (ya estaba limpia).');
  }

  // 3) Cliente de prueba del padron manual
  const clientes = getClientesSheet_();
  const ultima = clientes.getLastRow();
  let borroCliente = false;

  if (ultima >= 2) {
    const filas = clientes.getRange(2, 1, ultima - 1, CLIENTES_HEADERS.length).getValues();
    for (let i = filas.length - 1; i >= 0; i--) {
      if (safe_(filas[i][1]).trim() !== PRUEBA_CLAUDE.MARCA) continue;
      clientes.deleteRow(i + 2);
      log.push('Cliente de prueba borrado de ' + CONFIG.CLIENTES_SHEET_NAME);
      borroCliente = true;
    }
  }
  if (!borroCliente) log.push('Cliente de prueba: no se encontro (ya estaba limpio).');

  log.push('');
  log.push('Listo. Proximo comodato disponible: ' + getNextComodatoNumber_());

  const texto = log.join('\n');
  Logger.log(texto);
  return texto;
}
