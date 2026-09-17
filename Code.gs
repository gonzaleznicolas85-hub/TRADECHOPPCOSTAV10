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
  INTERVENCIONES_SHEET_NAME: 'Intervenciones',

  // Por que no se pudo sanitizar. La visita queda registrada pero NO cuenta
  // como sanitizacion: el ciclo de 28 dias sigue corriendo desde la ultima
  // que si se hizo.
  MOTIVOS_FALLA: [
    'PDV cerrado',
    'Equipo fuera de servicio',
    'No autorizaron el ingreso',
    'Falta de repuesto',
    'Otro'
  ],

  TIPOS_INTERVENCION: [
    'Arreglo de equipo',
    'Cambio de manguera',
    'Cambio de canilla',
    'Cambio de pilón',
    'Cambio de tubo de gas',
    'Cambio de regulador',
    'Instalación',
    'Retiro de equipo',
    'Otro'
  ],

  // --- BAJA DE CLIENTES ---
  // Un cliente dado de baja sale de la cartera y de los indicadores. El
  // comodato y todo su historial quedan en las hojas; se puede reactivar.
  BAJAS_SHEET_NAME: 'Bajas_Clientes',
  BAJAS_PHOTOS_FOLDER_NAME: 'Fotos Bajas',
  MOTIVOS_BAJA: [
    'Cierre del local',
    'Cambio de proveedor',
    'Bajo consumo',
    'Falta de pago / deuda',
    'Decisión del cliente',
    'Otro'
  ],
  RETIRO_EQUIPO_OPCIONES: [
    'Retirado completo, en buen estado',
    'Retirado con daños',
    'Retirado con faltantes',
    'No se retiró (queda pendiente)',
    'El cliente no tenía equipo'
  ],

  // --- MÓDULO HELADERAS (tickets que llegan desde un Google Form) ---
  // Hoja de respuestas del Form. Si el Form esta vinculado a ESTE mismo Sheet,
  // HELADERAS_FORM_SPREADSHEET_ID va vacio. Si el nombre de la pestaña no
  // coincide, se usa la primera que empiece con "Respuestas de formulario".
  // Planilla "Servicio TECNICO SMK (Respuestas)". La cuenta dueña de este
  // script tiene que tener acceso de EDITOR a esa planilla.
  HELADERAS_FORM_SPREADSHEET_ID: '1nkK-3soVT2aJFujxQ0zv6LcOIJ-904vyWyGE8LcV4Lk',
  // Quienes pueden tomar tickets de heladeras. Vacio = los mismos de CONFIG.TECNICOS.
  HELADERAS_TECNICOS: [],
  HELADERAS_FORM_SHEET_NAME: 'Respuestas de formulario 1',
  // Estado de gestion de cada ticket. La hoja del Form no se toca nunca.
  HELADERAS_SHEET_NAME: 'Tickets_Heladeras',
  HELADERAS_PHOTOS_FOLDER_NAME: 'Fotos Heladeras',
  // Los cerrados mas viejos que esto no se mandan a la app (siguen en la hoja)
  HELADERAS_DIAS_CERRADOS_VISIBLES: 60,
  // --- AVISOS POR WHATSAPP (ticket nuevo de heladera) ---
  // 'callmebot' (gratis, cada destinatario activa su clave), 'meta' (API oficial
  // de WhatsApp Business, con plantilla aprobada) o '' para apagarlos.
  // Los destinatarios se cargan en la hoja Avisos_WhatsApp, no aca.
  WHATSAPP_PROVEEDOR: 'callmebot',
  WHATSAPP_SHEET_NAME: 'Avisos_WhatsApp',
  WHATSAPP_META_PLANTILLA: 'nuevo_ticket_heladera',
  WHATSAPP_META_IDIOMA: 'es_AR',
  // URL publica de la app (Netlify). Si esta, el aviso trae el link al ticket.
  APP_URL: 'https://comodatoschoppcosta.netlify.app/',

  HELADERAS_MOTIVOS_NO_RESUELTO: [
    'Falta de repuesto',
    'Requiere retiro a taller',
    'PDV cerrado',
    'No autorizaron el ingreso',
    'Equipo no localizado',
    'No era falla del equipo',
    'Otro'
  ],

  // Nombres con los que un tecnico quedo cargado antes de un cambio de nombre.
  // Los comodatos viejos siguen teniendo el nombre anterior y no se reescriben
  // (son el registro firmado), asi que el cruce tiene que reconocer los dos.
  NOMBRES_ANTERIORES: {
    'José Alejandro Caporaletti': ['Jose Caporaletti']
  },

  TECNICOS: [
    'Gaston del Rio',
    'Federico Barbutti',
    'Maximiliano Di Pietro',
    'Mariano Diaz',
    'José Alejandro Caporaletti',
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

// Encabezados del registro de arreglos y cambios sobre las choperas
const INTERVENCIONES_HEADERS = [
  'IntervencionId', 'Fecha', 'Tecnico', 'Cliente', 'Tipo',
  'Detalle', 'Repuestos', 'FotosUrls', 'Estado', 'Timestamp'
];

// Encabezados del registro de bajas de clientes (una fila por baja)
const BAJAS_HEADERS = [
  'BajaId', 'FechaBaja', 'Tecnico', 'Cliente', 'ComodatoNumero',
  'Motivo', 'Observaciones', 'RetiroEquipo', 'DetalleRetiro',
  'Equipo', 'Pilon', 'CantPicos', 'FotosUrls',
  'Aclaracion', 'DNI', 'FirmaFileId', 'PdfFileId', 'PdfUrl', 'DocUrl',
  'Estado', 'ReactivadoEl', 'Timestamp'
];

// Encabezados de la gestion de tickets de heladeras (una fila por ticket tocado)
const HELADERAS_HEADERS = [
  'TicketId', 'Cliente', 'Estado', 'Tecnico', 'TomadoEl', 'CerradoEl',
  'TrabajoRealizado', 'Repuestos', 'MotivoNoResuelto', 'NotaEspera',
  'FotosUrls', 'Historial', 'Actualizado'
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
 *    - action=comodatosGeneral    -> todos los comodatos de todos los técnicos
 *    - action=choperasGeneral     -> todas las choperas de todos los técnicos
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

    if (action === 'choperas') {
      const t = resolveTecnico_(e.parameter.tecnico);
      const choperas = getChoperas_(t);
      return jsonOutput_({
        ok: true,
        tecnico: t,
        choperas: choperas,
        bajas: getBajasVigentes_(t, choperas),
        motivos: CONFIG.MOTIVOS_FALLA,
        tiposIntervencion: CONFIG.TIPOS_INTERVENCION,
        motivosBaja: CONFIG.MOTIVOS_BAJA,
        opcionesRetiro: CONFIG.RETIRO_EQUIPO_OPCIONES
      });
    }

    if (action === 'sanitHistorial') {
      const t = resolveTecnico_(e.parameter.tecnico);
      return jsonOutput_({ ok: true, tecnico: t, historial: getHistorialSanitizacion_(t) });
    }

    if (action === 'heladeras') {
      return jsonOutput_({
        ok: true,
        tickets: getTicketsHeladeras_(e.parameter.historico === '1'),
        motivosNoResuelto: CONFIG.HELADERAS_MOTIVOS_NO_RESUELTO,
        tecnicos: tecnicosHeladeras_()
      });
    }

    if (action === 'comodatosGeneral') {
      return jsonOutput_({ ok: true, general: getComodatosGeneral_() });
    }

    if (action === 'choperasGeneral') {
      return jsonOutput_({ ok: true, general: getChoperasGeneral_() });
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
      case 'sanitVisitaFallida': return jsonOutput_(sanitVisitaFallida_(body));
      case 'guardarChopera':  return jsonOutput_(guardarChopera_(body));
      case 'nuevaIntervencion': return jsonOutput_(nuevaIntervencion_(body));
      case 'cerrarIntervencion': return jsonOutput_(cerrarIntervencion_(body));
      case 'sanitBajaCliente':return jsonOutput_(sanitBajaCliente_(body));
      case 'bajaCliente':     return jsonOutput_(bajaCliente_(body));
      case 'reactivarCliente':return jsonOutput_(reactivarCliente_(body));
      case 'heladeraAccion':  return jsonOutput_(heladeraAccion_(body));
      case 'heladeraFinalizar': return jsonOutput_(heladeraFinalizar_(body));
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
 * Deja el archivo visible para cualquiera que tenga el link.
 *
 * Sin esto, el link que ve el tecnico en el celular lo manda a iniciar sesion
 * con la cuenta dueña del Sheet: los archivos que crea el script nacen
 * privados de esa cuenta. Es el mismo permiso que necesita el boton de
 * WhatsApp, que le manda el PDF al cliente.
 *
 * No corta el guardado si falla: algunas configuraciones de Workspace no
 * permiten compartir fuera del dominio, y perder el comodato entero por eso
 * seria peor que tener un link que hay que abrir logueado.
 */
function compartirPorLink_(file) {
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return true;
  } catch (err) {
    Logger.log('No se pudo compartir ' + file.getName() + ': ' + err.message);
    return false;
  }
}

/**
 * Comparte los archivos que ya estaban creados desde antes del fix.
 * Se corre a mano desde el editor, una sola vez. Es idempotente: volver a
 * correrla no rompe nada.
 */
function compartirArchivosExistentes() {
  const log = [];
  let ok = 0, fallaron = 0;

  const compartirPorId_ = (id, etiqueta) => {
    const limpio = safe_(id).trim();
    if (!limpio) return;
    try {
      if (compartirPorLink_(DriveApp.getFileById(limpio))) ok++;
      else fallaron++;
    } catch (err) {
      fallaron++;
      log.push('  no se pudo abrir ' + etiqueta + ' (' + limpio + '): ' + err.message);
    }
  };

  // Comodatos: PDF y Doc
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const datos = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
    const iPdf = HEADERS.indexOf('PdfFileId');
    const iPdfUrl = HEADERS.indexOf('PdfUrl');
    const iDoc = HEADERS.indexOf('DocFileId');
    const iFotosEq = HEADERS.indexOf('FotosEquiposUrls');
    const iFotosPil = HEADERS.indexOf('FotosPilonesUrls');

    datos.forEach(fila => {
      compartirPorId_(fila[iPdf], 'PDF');
      // En las filas viejas el PDF esta en PdfUrl (ver urlArchivoDrive_)
      compartirPorId_(idDesdeUrlDrive_(fila[iPdfUrl]), 'PDF');
      compartirPorId_(fila[iDoc], 'Doc');
      [fila[iFotosEq], fila[iFotosPil]].forEach(celda => {
        safe_(celda).split('\n').filter(String).forEach(url => {
          compartirPorId_(idDesdeUrlDrive_(url), 'foto de comodato');
        });
      });
    });
    log.push('Comodatos revisados: ' + datos.length);
  }

  // Fotos de sanitizaciones
  CONFIG.TECNICOS.forEach(t => {
    readSanitRows_(t).forEach(r => {
      safe_(r.FotosUrls).split('\n').filter(String).forEach(url => {
        compartirPorId_(idDesdeUrlDrive_(url), 'foto de sanitización');
      });
    });
  });

  // Fotos de intervenciones
  readIntervenciones_('').forEach(i => {
    safe_(i.FotosUrls).split('\n').filter(String).forEach(url => {
      compartirPorId_(idDesdeUrlDrive_(url), 'foto de intervención');
    });
  });

  log.push('Archivos compartidos: ' + ok);
  log.push('Archivos que no se pudieron compartir: ' + fallaron);

  const texto = log.join('\n');
  Logger.log(texto);
  return texto;
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
  compartirPorLink_(pdfFile);
  compartirPorLink_(tempCopy);

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
    compartirPorLink_(file);
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
    codCliente: HEADERS.indexOf('CodCliente'),
    cuit: HEADERS.indexOf('CUIT'),
    aclaracion: HEADERS.indexOf('Aclaracion'),
    dni: HEADERS.indexOf('DNI'),
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
    const key = keyTecnico_(fila[idx.tecnico]);
    if (!key) return;
    const ts = fila[idx.timestamp];
    if (!mapa[key]) mapa[key] = [];
    mapa[key].push({
      comodatoNumero: safe_(fila[idx.comodatoNumero]),
      // Nombre tal como quedo cargado en la fila: el tablero general necesita
      // saber de quien es cada comodato, la vista por tecnico lo ignora.
      tecnico: safe_(fila[idx.tecnico]),
      fecha: safe_(fila[idx.fecha]),
      distribuidor: safe_(fila[idx.distribuidor]),
      codCliente: safe_(fila[idx.codCliente]),
      cuit: safe_(fila[idx.cuit]),
      aclaracion: safe_(fila[idx.aclaracion]),
      dni: safe_(fila[idx.dni]),
      nombreFantasia: safe_(fila[idx.nombreFantasia]),
      razonSocial: safe_(fila[idx.razonSocial]),
      localidad: safe_(fila[idx.localidad]),
      domicilio: safe_(fila[idx.domicilio]),
      pdfUrl: urlArchivoDrive_(fila[idx.pdfUrl]),
      docUrl: safe_(fila[idx.docUrl]).indexOf('http') === 0 ? safe_(fila[idx.docUrl]) : '',
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
  return leerComodatosAgrupados_()[keyTecnico_(tecnico)] || [];
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
    const key = keyTecnico_(f[0]);
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

/**
 * Clave con la que se agrupa a un tecnico. Es norm_() salvo que el valor sea
 * un nombre anterior del mismo tecnico, en cuyo caso devuelve la clave del
 * nombre actual. Todo lo que agrupe o filtre por tecnico tiene que usar esto,
 * no norm_(), o los registros cargados con el nombre viejo se pierden.
 */
function keyTecnico_(valor) {
  const n = norm_(valor);
  if (!n) return '';

  const mapa = CONFIG.NOMBRES_ANTERIORES || {};
  for (const canonico in mapa) {
    if (norm_(canonico) === n) return n;
    const previos = mapa[canonico] || [];
    for (let i = 0; i < previos.length; i++) {
      if (norm_(previos[i]) === n) return norm_(canonico);
    }
  }
  return n;
}

/** Devuelve el nombre canónico del técnico según CONFIG.TECNICOS. Lanza si no existe. */
function resolveTecnico_(tecnico) {
  const objetivo = keyTecnico_(tecnico);
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

  // Si el tecnico cambio de nombre, su hoja anterior se renombra en vez de
  // crear una vacia al lado y dejar el historial huerfano.
  if (!sheet) {
    const previos = (CONFIG.NOMBRES_ANTERIORES || {})[tecnico] || [];
    for (let i = 0; i < previos.length; i++) {
      const vieja = ss.getSheetByName(CONFIG.SANIT_SHEET_PREFIX + previos[i]);
      if (vieja) { vieja.setName(nombre); sheet = vieja; break; }
    }
  }

  if (!sheet) sheet = ss.insertSheet(nombre);
  escribirCabecera_(sheet, SANIT_HEADERS);
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
  getIntervencionesSheet_();
  getBajasSheet_();
  getHeladerasSheet_();
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
    leerComodatosAgrupados_()[keyTecnico_(tecnico)] || [],
    leerClientesManualesAgrupados_()[keyTecnico_(tecnico)] || [],
    readSanitRows_(tecnico),
    leerBajasVigentesAgrupadas_()[keyTecnico_(tecnico)] || {}
  );
}

/**
 * Arma la cartera a partir de datos ya leidos. Separada de la carga para que
 * el resumen general pueda reutilizarla sin releer las hojas por cada tecnico.
 */
function construirCartera_(tecnico, comodatos, clientesManuales, sanitRows, bajas) {
  const mapa = {};
  bajas = bajas || {};

  // ficha: datos del comodato (razon social, CUIT, PDF...) si el cliente tiene uno
  const push_ = (cliente, direccion, localidad, origen, comodatoNumero, fechaBase, chopera, ficha) => {
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
        cantPicos: toNumber_(eq.cantPicos),
        ficha: ficha || null
      };
      return;
    }
    // Completar huecos y quedarse con la fecha base más reciente
    const actual = mapa[key];
    if (!actual.direccion) actual.direccion = safe_(direccion).trim();
    if (!actual.localidad) actual.localidad = safe_(localidad).trim();
    if (!actual.comodatoNumero) actual.comodatoNumero = safe_(comodatoNumero);
    if (!actual.ficha && ficha) actual.ficha = ficha;
    if (!actual.equipo) actual.equipo = safe_(eq.equipo).trim();
    if (!actual.pilon) actual.pilon = safe_(eq.pilon).trim();
    if (!actual.cantPicos) actual.cantPicos = toNumber_(eq.cantPicos);
    if (fechaBase && (!actual.fechaBase || fechaBase > actual.fechaBase)) actual.fechaBase = fechaBase;
  };

  // A) Clientes que salen de los comodatos cargados
  comodatos.forEach(c => {
    push_(c.nombreFantasia || c.razonSocial, c.domicilio, c.localidad, 'COMODATO',
          c.comodatoNumero, toDate_(c.fecha) || (c.timestamp ? new Date(c.timestamp) : null),
          { equipo: c.equipo, pilon: c.pilon, cantPicos: c.cantPicos },
          {
            comodatoNumero: c.comodatoNumero,
            fecha: isoDate_(toDate_(c.fecha)) || c.fecha,
            distribuidor: c.distribuidor,
            codCliente: c.codCliente,
            cuit: c.cuit,
            nombreFantasia: c.nombreFantasia,
            razonSocial: c.razonSocial,
            domicilio: c.domicilio,
            localidad: c.localidad,
            aclaracion: c.aclaracion,
            dni: c.dni,
            pdfUrl: c.pdfUrl
          });
  });

  // B) Clientes del padrón manual
  clientesManuales.forEach(f => {
    push_(f[1], f[2], f[3], 'MANUAL', '', toDate_(f[5]),
          { equipo: f[6], pilon: f[7], cantPicos: f[8] });
  });

  // C) Cruce con el historial de sanitizaciones
  const ultima = {};
  const abierta = {};
  const fallidas = {};
  sanitRows.forEach(r => {
    const key = norm_(r.Cliente);
    if (!key) return;
    if (!mapa[key]) push_(r.Cliente, r.Direccion, r.Localidad, safe_(r.Origen) || 'MANUAL', r.ComodatoNumero, null, null);

    const estado = norm_(r.Estado);

    if (estado === 'en curso') {
      abierta[key] = { sanitizacionId: safe_(r.SanitizacionId), checkIn: toDate_(r.CheckIn) };
      return;
    }

    // Una visita fallida deja constancia de que el tecnico fue, pero NO es una
    // sanitizacion: el ciclo sigue corriendo desde la ultima que si se hizo.
    if (estado === 'no realizada') {
      const cuando = toDate_(r.CheckOut) || toDate_(r.CheckIn);
      if (!fallidas[key]) fallidas[key] = { cantidad: 0, ultima: null, motivo: '' };
      fallidas[key].cantidad++;
      if (cuando && (!fallidas[key].ultima || cuando > fallidas[key].ultima)) {
        fallidas[key].ultima = cuando;
        fallidas[key].motivo = safe_(r.Observaciones);
      }
      return;
    }

    if (estado !== 'completada') return;

    const fin = toDate_(r.CheckOut);
    if (fin && (!ultima[key] || fin > ultima[key])) ultima[key] = fin;
  });

  // D) Clientes dados de baja: salen de la cartera. Si despues de la baja se
  // firmo un comodato nuevo con el mismo nombre, el cliente volvio y se muestra.
  const ultimoComodato = {};
  comodatos.forEach(c => {
    const key = norm_(c.nombreFantasia || c.razonSocial);
    if (key && c.timestamp > (ultimoComodato[key] || 0)) ultimoComodato[key] = c.timestamp;
  });
  const deBaja = key => !!bajas[key] && !((ultimoComodato[key] || 0) > bajas[key].timestamp);

  const hoy = new Date();
  const ciclo = CONFIG.SANIT_CICLO_DIAS;
  const aviso = CONFIG.SANIT_AVISO_DIAS;

  return Object.keys(mapa).filter(key => !deBaja(key)).map(key => {
    const c = mapa[key];
    const ult = ultima[key] || null;
    const falla = fallidas[key] || null;
    // Solo interesan las fallidas posteriores a la ultima sanitizacion hecha
    const fallaVigente = falla && falla.ultima && (!ult || falla.ultima > ult) ? falla : null;
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
      ficha: c.ficha,
      ultimaSanitizacion: isoDate_(ult),
      proximaSanitizacion: isoDate_(prox),
      diasRestantes: dias,
      estado: estado,
      sanitizacionAbiertaId: abierta[key] ? abierta[key].sanitizacionId : '',
      checkInAbierto: abierta[key] ? isoDateTime_(abierta[key].checkIn) : '',
      visitasFallidas: fallaVigente ? fallaVigente.cantidad : 0,
      ultimaVisitaFallida: fallaVigente ? isoDate_(fallaVigente.ultima) : '',
      motivoUltimaFalla: fallaVigente ? fallaVigente.motivo : ''
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
  const bajas = leerBajasVigentesAgrupadas_();

  const totales = {
    totalClientes: 0, alDia: 0, porVencer: 0, vencidos: 0, sinRegistro: 0,
    enCurso: 0, completadasUltimos30: 0, totalSanitizaciones: 0, minutosPromedioPdv: 0
  };

  let minutosAcumulados = 0, sanitConMinutos = 0;

  const porTecnico = CONFIG.TECNICOS.map(tecnico => {
    const key = keyTecnico_(tecnico);
    const filas = readSanitRows_(tecnico);
    const cartera = construirCartera_(tecnico, comodatos[key] || [], clientes[key] || [], filas, bajas[key] || {});
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

    // Volver a cargar a mano un cliente dado de baja lo reactiva
    const reactivadas = reactivarBajasDe_(tecnico, cliente);

    const sheet = getClientesSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const rows = sheet.getRange(2, 1, lastRow - 1, CLIENTES_HEADERS.length).getValues();
      const duplicado = rows.some(f => keyTecnico_(f[0]) === keyTecnico_(tecnico) && norm_(f[1]) === norm_(cliente));
      if (duplicado && reactivadas) return { ok: true, message: 'El cliente estaba dado de baja: quedó reactivado en la cartera.' };
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
      if (keyTecnico_(rows[i][0]) === keyTecnico_(tecnico) && norm_(rows[i][1]) === norm_(cliente)) {
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

/**
 * Link para abrir un archivo de Drive. Los comodatos viejos (hasta TCC-00056
 * aprox.) se cargaron con otro orden de columnas y en PdfUrl quedo el ID del
 * PDF en vez del link: sin esto, "Ver PDF" abria una direccion invalida.
 */
function urlArchivoDrive_(valor) {
  const txt = safe_(valor).trim();
  if (!txt || txt.indexOf('http') === 0) return txt;
  const id = idDesdeUrlDrive_(txt);
  return id === txt ? 'https://drive.google.com/file/d/' + id + '/view' : '';
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

/* ==========================================================================
 * 9. VISITAS FALLIDAS, CHOPERAS E INTERVENCIONES
 * ========================================================================== */

/** Hoja unica de arreglos y cambios, con columna Tecnico. */
function getIntervencionesSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.INTERVENCIONES_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(CONFIG.INTERVENCIONES_SHEET_NAME);
  escribirCabecera_(sheet, INTERVENCIONES_HEADERS);
  return sheet;
}

/** Lee las intervenciones como objetos, opcionalmente filtradas por tecnico. */
function readIntervenciones_(tecnico) {
  const sheet = getIntervencionesSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const filtro = tecnico ? keyTecnico_(tecnico) : '';
  const values = sheet.getRange(2, 1, lastRow - 1, INTERVENCIONES_HEADERS.length).getValues();

  const filas = [];
  values.forEach((fila, i) => {
    const obj = { _row: i + 2 };
    INTERVENCIONES_HEADERS.forEach((h, c) => { obj[h] = fila[c]; });
    if (filtro && keyTecnico_(obj.Tecnico) !== filtro) return;
    filas.push(obj);
  });
  return filas;
}

/**
 * Registra que el tecnico fue al PDV y no pudo sanitizar.
 *
 * Queda como una fila mas del historial, con Estado NO REALIZADA. Al no ser
 * COMPLETADA, getCarteraSanitizacion_ no la toma como sanitizacion y el ciclo
 * de 28 dias sigue corriendo desde la ultima que si se hizo. Eso es todo el
 * punto: sin esto, el PDV que nunca se deja atender es indistinguible del
 * tecnico que no fue.
 *
 * Si venia de un check-in abierto, cierra esa fila. Si no, crea una nueva.
 */
function sanitVisitaFallida_(body) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const tecnico = resolveTecnico_(body.tecnico);
    const cliente = safe_(body.cliente).trim();
    const motivo = safe_(body.motivo).trim();
    if (!cliente) throw new Error('Falta el cliente.');
    if (!motivo) throw new Error('Hay que indicar por qué no se pudo sanitizar.');

    const detalle = safe_(body.detalle).trim();
    const observacion = detalle ? motivo + ' — ' + detalle : motivo;
    const sheet = getSanitSheet_(tecnico);
    const ahora = new Date();
    const id = safe_(body.sanitizacionId).trim();

    if (id) {
      const fila = readSanitRows_(tecnico).filter(r => safe_(r.SanitizacionId) === id)[0];
      if (!fila) throw new Error('No se encontró la sanitización ' + id);
      if (norm_(fila.Estado) !== 'en curso') throw new Error('Esa sanitización ya está cerrada.');

      const inicio = toDate_(fila.CheckIn) || ahora;
      const col = h => SANIT_HEADERS.indexOf(h) + 1;
      sheet.getRange(fila._row, col('CheckOut')).setValue(ahora);
      sheet.getRange(fila._row, col('MinutosEnPdv')).setValue(
        Math.max(1, Math.round((ahora.getTime() - inicio.getTime()) / 60000)));
      sheet.getRange(fila._row, col('Observaciones')).setValue(observacion);
      sheet.getRange(fila._row, col('Estado')).setValue('NO REALIZADA');

      return { ok: true, message: 'Visita registrada como no realizada.', sanitizacionId: id };
    }

    const nuevoId = 'SAN-' + ahora.getTime().toString(36).toUpperCase();
    sheet.appendRow([
      nuevoId, tecnico, cliente, safe_(body.direccion), safe_(body.localidad),
      safe_(body.comodatoNumero), safe_(body.origen) || 'MANUAL',
      ahora, ahora, 0,
      observacion, '', 'NO REALIZADA', '', ahora
    ]);

    return { ok: true, message: 'Visita registrada como no realizada.', sanitizacionId: nuevoId };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Vista "Mis Choperas": la cartera del tecnico con el estado de comodato, el
 * de sanitizacion, los datos del equipo y las intervenciones de cada PDV.
 */
function getChoperas_(tecnico) {
  return armarChoperas_(getCarteraSanitizacion_(tecnico), readIntervenciones_(tecnico));
}

/**
 * Cruza una cartera ya armada con su historial de intervenciones. Separada de
 * la carga para que el tablero general la reutilice sin releer las hojas una
 * vez por tecnico.
 */
function armarChoperas_(cartera, intervenciones) {
  const porCliente = {};
  intervenciones.forEach(i => {
    const key = norm_(i.Cliente);
    if (!key) return;
    if (!porCliente[key]) porCliente[key] = [];
    porCliente[key].push({
      intervencionId: safe_(i.IntervencionId),
      fecha: isoDate_(toDate_(i.Fecha)),
      tipo: safe_(i.Tipo),
      detalle: safe_(i.Detalle),
      repuestos: safe_(i.Repuestos),
      fotosUrls: safe_(i.FotosUrls),
      estado: safe_(i.Estado) || 'RESUELTO',
      _orden: toDate_(i.Fecha) ? toDate_(i.Fecha).getTime() : 0
    });
  });

  Object.keys(porCliente).forEach(k => porCliente[k].sort((a, b) => b._orden - a._orden));

  return cartera.map(c => {
    const lista = porCliente[norm_(c.cliente)] || [];
    const pendientes = lista.filter(i => norm_(i.estado) === 'pendiente');

    return {
      cliente: c.cliente,
      direccion: c.direccion,
      localidad: c.localidad,
      origen: c.origen,
      tieneComodato: !!c.comodatoNumero,
      comodatoNumero: c.comodatoNumero,
      pdfUrl: c.ficha ? c.ficha.pdfUrl : '',
      ficha: c.ficha,
      equipo: c.equipo,
      pilon: c.pilon,
      cantPicos: c.cantPicos,
      estadoSanitizacion: c.estado,
      sanitizado: c.estado === 'AL DIA',
      ultimaSanitizacion: c.ultimaSanitizacion,
      proximaSanitizacion: c.proximaSanitizacion,
      diasRestantes: c.diasRestantes,
      visitasFallidas: c.visitasFallidas,
      motivoUltimaFalla: c.motivoUltimaFalla,
      intervenciones: lista,
      intervencionesPendientes: pendientes.length,
      intervencionesTotal: lista.length
    };
  });
}

/**
 * Alta o edicion de los datos de una chopera.
 *
 * Todo se guarda en Clientes_Sanitizacion, que funciona como capa de estado
 * actual: lo que se cargue ahi pisa lo que traiga el comodato. El comodato no
 * se toca nunca, es el registro firmado.
 *
 * Renombrar solo se permite en clientes sin comodato. Si el cliente vino de un
 * comodato, el nombre lo manda el comodato: cambiarlo aca partiria la cartera
 * en dos (el nombre viejo seguiria llegando desde la hoja de comodatos).
 */
function guardarChopera_(body) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const tecnico = resolveTecnico_(body.tecnico);
    const original = safe_(body.clienteOriginal).trim();
    const cliente = safe_(body.cliente).trim();
    if (!cliente) throw new Error('El nombre del cliente no puede quedar vacío.');

    const renombra = original && norm_(original) !== norm_(cliente);

    if (renombra) {
      const comodatos = leerComodatosAgrupados_()[keyTecnico_(tecnico)] || [];
      const tieneComodato = comodatos.some(
        c => norm_(c.nombreFantasia || c.razonSocial) === norm_(original));
      if (tieneComodato) {
        return {
          ok: false,
          message: 'Ese cliente viene de un comodato: el nombre sale de ahí y no se puede cambiar desde acá. ' +
                   'Si está mal escrito, hay que corregirlo en la hoja de comodatos.'
        };
      }
    }

    const sheet = getClientesSheet_();
    const lastRow = sheet.getLastRow();
    const buscado = norm_(original || cliente);
    let fila = 0;

    if (lastRow >= 2) {
      const rows = sheet.getRange(2, 1, lastRow - 1, CLIENTES_HEADERS.length).getValues();
      for (let i = 0; i < rows.length; i++) {
        if (keyTecnico_(rows[i][0]) === keyTecnico_(tecnico) && norm_(rows[i][1]) === buscado) { fila = i + 2; break; }
      }
    }

    const valores = [
      tecnico, cliente, safe_(body.direccion).trim(), safe_(body.localidad).trim(), 'SI',
      new Date(), safe_(body.equipo).trim(), safe_(body.pilon).trim(), toNumber_(body.cantPicos)
    ];

    if (fila) {
      // La fecha de alta original no se pisa: es el punto de partida del ciclo
      const altaPrevia = sheet.getRange(fila, CLIENTES_HEADERS.indexOf('FechaAlta') + 1).getValue();
      if (altaPrevia) valores[5] = altaPrevia;
      sheet.getRange(fila, 1, 1, CLIENTES_HEADERS.length).setValues([valores]);
    } else {
      sheet.appendRow(valores);
    }

    if (renombra) propagarRenombre_(tecnico, original, cliente);
    // Alta a mano de un cliente dado de baja: vuelve a la cartera
    if (!original) reactivarBajasDe_(tecnico, cliente);

    return { ok: true, message: 'Chopera actualizada.' };
  } finally {
    lock.releaseLock();
  }
}

/** Al renombrar un cliente sin comodato, su historial tiene que seguirlo. */
function propagarRenombre_(tecnico, viejo, nuevo) {
  const sanitSheet = getSanitSheet_(tecnico);
  const colCliente = SANIT_HEADERS.indexOf('Cliente') + 1;
  readSanitRows_(tecnico).forEach(r => {
    if (norm_(r.Cliente) === norm_(viejo)) sanitSheet.getRange(r._row, colCliente).setValue(nuevo);
  });

  const intSheet = getIntervencionesSheet_();
  const colIntCliente = INTERVENCIONES_HEADERS.indexOf('Cliente') + 1;
  readIntervenciones_(tecnico).forEach(i => {
    if (norm_(i.Cliente) === norm_(viejo)) intSheet.getRange(i._row, colIntCliente).setValue(nuevo);
  });
}

/** Registra un arreglo o cambio sobre la chopera de un PDV. */
function nuevaIntervencion_(body) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const tecnico = resolveTecnico_(body.tecnico);
    const cliente = safe_(body.cliente).trim();
    const tipo = safe_(body.tipo).trim();
    if (!cliente) throw new Error('Falta el cliente.');
    if (!tipo) throw new Error('Falta el tipo de intervención.');

    const ahora = new Date();
    const id = 'INT-' + ahora.getTime().toString(36).toUpperCase();
    const estado = norm_(body.estado) === 'pendiente' ? 'PENDIENTE' : 'RESUELTO';

    let urls = '';
    if (body.fotos && body.fotos.length) {
      urls = savePhotos_(getIntervencionPhotosFolder_(tecnico), body.fotos, id);
    }

    getIntervencionesSheet_().appendRow([
      id, ahora, tecnico, cliente, tipo,
      safe_(body.detalle).trim(), safe_(body.repuestos).trim(), urls, estado, ahora
    ]);

    return { ok: true, message: 'Intervención registrada.', intervencionId: id, estado: estado };
  } finally {
    lock.releaseLock();
  }
}

/** Marca como resuelta una intervencion que habia quedado pendiente. */
function cerrarIntervencion_(body) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const tecnico = resolveTecnico_(body.tecnico);
    const id = safe_(body.intervencionId).trim();
    if (!id) throw new Error('Falta el ID de la intervención.');

    const fila = readIntervenciones_(tecnico).filter(i => safe_(i.IntervencionId) === id)[0];
    if (!fila) throw new Error('No se encontró la intervención ' + id);
    if (norm_(fila.Estado) === 'resuelto') return { ok: false, message: 'Esa intervención ya estaba resuelta.' };

    getIntervencionesSheet_()
      .getRange(fila._row, INTERVENCIONES_HEADERS.indexOf('Estado') + 1)
      .setValue('RESUELTO');

    return { ok: true, message: 'Intervención marcada como resuelta.' };
  } finally {
    lock.releaseLock();
  }
}

/** Carpeta de fotos de intervenciones, con subcarpeta por tecnico. */
function getIntervencionPhotosFolder_(tecnico) {
  const base = getPhotosParentFolder_();
  const raiz = base.getFoldersByName('Fotos Intervenciones');
  const parent = raiz.hasNext() ? raiz.next() : base.createFolder('Fotos Intervenciones');
  const sub = parent.getFoldersByName(tecnico);
  return sub.hasNext() ? sub.next() : parent.createFolder(tecnico);
}

/* ==========================================================================
 * 9.b BAJA DE CLIENTES
 *
 * La baja no borra nada: agrega una fila en Bajas_Clientes con el motivo, el
 * retiro de la chopera, fotos, firma y un PDF de acta. Mientras la fila este
 * ACTIVA el cliente no aparece en la cartera ni cuenta en los indicadores.
 * Reactivar pasa la fila a REACTIVADA y el cliente vuelve con su historial.
 * ========================================================================== */

function getBajasSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.BAJAS_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(CONFIG.BAJAS_SHEET_NAME);
  escribirCabecera_(sheet, BAJAS_HEADERS);
  return sheet;
}

/** Lee las bajas como objetos, opcionalmente filtradas por tecnico. */
function readBajas_(tecnico) {
  const sheet = getBajasSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const filtro = tecnico ? keyTecnico_(tecnico) : '';
  const values = sheet.getRange(2, 1, lastRow - 1, BAJAS_HEADERS.length).getValues();

  const filas = [];
  values.forEach((fila, i) => {
    const obj = { _row: i + 2 };
    BAJAS_HEADERS.forEach((h, c) => { obj[h] = fila[c]; });
    if (filtro && keyTecnico_(obj.Tecnico) !== filtro) return;
    filas.push(obj);
  });
  return filas;
}

/** { tecnicoKey: { clienteKey: { bajaId, timestamp } } } con las bajas activas. */
function leerBajasVigentesAgrupadas_() {
  const mapa = {};
  readBajas_().forEach(b => {
    if (norm_(b.Estado) !== 'activa') return;
    const tk = keyTecnico_(b.Tecnico);
    const ck = norm_(b.Cliente);
    if (!tk || !ck) return;
    const ts = toDate_(b.Timestamp) ? toDate_(b.Timestamp).getTime() : 0;
    if (!mapa[tk]) mapa[tk] = {};
    if (!mapa[tk][ck] || ts > mapa[tk][ck].timestamp) {
      mapa[tk][ck] = { bajaId: safe_(b.BajaId), timestamp: ts };
    }
  });
  return mapa;
}

/**
 * Bajas activas del tecnico que efectivamente dejan al cliente fuera de la
 * cartera (si firmo un comodato nuevo despues, ya no figura como baja).
 */
function getBajasVigentes_(tecnico, cartera) {
  const enCartera = {};
  (cartera || getCarteraSanitizacion_(tecnico)).forEach(c => { enCartera[norm_(c.cliente)] = true; });

  return readBajas_(tecnico)
    .filter(b => norm_(b.Estado) === 'activa' && !enCartera[norm_(b.Cliente)])
    .map(b => ({
      bajaId: safe_(b.BajaId),
      cliente: safe_(b.Cliente),
      fechaBaja: isoDate_(toDate_(b.FechaBaja)) || safe_(b.FechaBaja),
      comodatoNumero: safe_(b.ComodatoNumero),
      motivo: safe_(b.Motivo),
      observaciones: safe_(b.Observaciones),
      retiroEquipo: safe_(b.RetiroEquipo),
      detalleRetiro: safe_(b.DetalleRetiro),
      equipo: safe_(b.Equipo),
      pilon: safe_(b.Pilon),
      fotosUrls: safe_(b.FotosUrls),
      pdfUrl: safe_(b.PdfUrl),
      _orden: toDate_(b.Timestamp) ? toDate_(b.Timestamp).getTime() : 0
    }))
    .sort((a, b) => b._orden - a._orden);
}

/** "2026-09-15" -> Date local (new Date("2026-09-15") lo toma en UTC y corre un dia). */
function fechaLocal_(txt) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(safe_(txt).trim());
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : new Date();
}

/** Da de baja un cliente: fotos, firma, PDF de acta y fila en Bajas_Clientes. */
function bajaCliente_(body) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const tecnico = resolveTecnico_(body.tecnico);
    const cliente = safe_(body.cliente).trim();
    const motivo = safe_(body.motivo).trim();
    const retiro = safe_(body.retiroEquipo).trim();
    const aclaracion = safe_(body.aclaracion).trim();
    const dni = safe_(body.dni).trim();

    if (!cliente) throw new Error('Falta el cliente.');
    if (!motivo) throw new Error('Indicá el motivo de la baja.');
    if (!retiro) throw new Error('Indicá qué pasó con la chopera.');
    if (!body.firmaDataUrl) throw new Error('Falta la firma.');
    if (!aclaracion || !dni) throw new Error('Completá aclaración y DNI de quien firma.');

    const item = getCarteraSanitizacion_(tecnico).filter(c => norm_(c.cliente) === norm_(cliente))[0];
    if (!item) {
      return { ok: false, message: 'Ese cliente no está en la cartera (¿ya estaba dado de baja?).' };
    }

    const ahora = new Date();
    const id = 'BAJ-' + ahora.getTime().toString(36).toUpperCase();
    const fechaBaja = fechaLocal_(body.fecha);
    const ficha = item.ficha || {};

    const urls = (body.fotos && body.fotos.length)
      ? savePhotos_(getBajaPhotosFolder_(tecnico), body.fotos, id)
      : '';
    const firma = saveSignature_(body.firmaDataUrl, id);

    const archivos = generarPdfBaja_({
      bajaId: id,
      fecha: Utilities.formatDate(fechaBaja, Session.getScriptTimeZone(), 'dd/MM/yyyy'),
      tecnico: tecnico,
      cliente: item.cliente,
      razonSocial: ficha.razonSocial,
      cuit: ficha.cuit,
      codCliente: ficha.codCliente,
      distribuidor: ficha.distribuidor,
      domicilio: ficha.domicilio || item.direccion,
      localidad: ficha.localidad || item.localidad,
      comodatoNumero: item.comodatoNumero,
      fechaComodato: ficha.fecha,
      motivo: motivo,
      observaciones: safe_(body.observaciones).trim(),
      retiroEquipo: retiro,
      detalleRetiro: safe_(body.detalleRetiro).trim(),
      equipo: item.equipo,
      pilon: item.pilon,
      cantPicos: item.cantPicos,
      aclaracion: aclaracion,
      dni: dni
    }, body.fotos, firma);

    getBajasSheet_().appendRow([
      id, fechaBaja, tecnico, item.cliente, safe_(item.comodatoNumero),
      motivo, safe_(body.observaciones).trim(), retiro, safe_(body.detalleRetiro).trim(),
      item.equipo, item.pilon, toNumber_(item.cantPicos), urls,
      aclaracion, dni, firma.getId(),
      archivos.pdfFile.getId(), archivos.pdfFile.getUrl(), archivos.docFile.getUrl(),
      'ACTIVA', '', ahora
    ]);

    return { ok: true, message: 'Cliente dado de baja.', bajaId: id, pdfUrl: archivos.pdfFile.getUrl() };
  } finally {
    lock.releaseLock();
  }
}

/** Vuelve a poner en la cartera un cliente dado de baja. */
function reactivarCliente_(body) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const tecnico = resolveTecnico_(body.tecnico);
    const id = safe_(body.bajaId).trim();
    if (!id) throw new Error('Falta el ID de la baja.');

    const fila = readBajas_(tecnico).filter(b => safe_(b.BajaId) === id)[0];
    if (!fila) throw new Error('No se encontró la baja ' + id);
    if (norm_(fila.Estado) !== 'activa') return { ok: false, message: 'Ese cliente ya estaba reactivado.' };

    // Reactiva todas las bajas activas del cliente, no solo esta
    reactivarBajasDe_(tecnico, fila.Cliente);
    return { ok: true, message: 'Cliente reactivado.', cliente: safe_(fila.Cliente) };
  } finally {
    lock.releaseLock();
  }
}

/** Pasa a REACTIVADA las bajas activas de un cliente. Sin lock: lo toma quien llama. */
function reactivarBajasDe_(tecnico, cliente) {
  const sheet = getBajasSheet_();
  const colEstado = BAJAS_HEADERS.indexOf('Estado') + 1;
  const colReact = BAJAS_HEADERS.indexOf('ReactivadoEl') + 1;
  const ahora = new Date();
  let n = 0;
  readBajas_(tecnico).forEach(b => {
    if (norm_(b.Estado) !== 'activa' || norm_(b.Cliente) !== norm_(cliente)) return;
    sheet.getRange(b._row, colEstado).setValue('REACTIVADA');
    sheet.getRange(b._row, colReact).setValue(ahora);
    n++;
  });
  return n;
}

function getBajaPhotosFolder_(tecnico) {
  const base = getPhotosParentFolder_();
  const raiz = base.getFoldersByName(CONFIG.BAJAS_PHOTOS_FOLDER_NAME);
  const parent = raiz.hasNext() ? raiz.next() : base.createFolder(CONFIG.BAJAS_PHOTOS_FOLDER_NAME);
  const sub = parent.getFoldersByName(tecnico);
  return sub.hasNext() ? sub.next() : parent.createFolder(tecnico);
}

/**
 * Arma el acta de baja en un Doc nuevo (no usa plantilla) y lo exporta a PDF
 * en la misma carpeta que los PDFs de comodatos.
 */
function generarPdfBaja_(d, fotos, signatureFile) {
  const pdfFolder = DriveApp.getFolderById(CONFIG.PDF_FOLDER_ID);
  const nombre = 'Baja_' + d.bajaId + '_' + safe_(d.cliente).replace(/[/\\?%*:|"<>]/g, '-');

  const doc = DocumentApp.create(nombre + '_DOC');
  const docFile = DriveApp.getFileById(doc.getId());
  docFile.moveTo(pdfFolder);

  const body = doc.getBody();
  const azul = '#003366';

  const titulo = body.getParagraphs()[0];
  titulo.setText('ACTA DE BAJA DE CLIENTE')
    .setHeading(DocumentApp.ParagraphHeading.HEADING1)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  titulo.editAsText().setForegroundColor(azul).setBold(true);
  body.appendParagraph('Trade Marketing Chopp Costa · N° ' + d.bajaId + ' · ' + d.fecha)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  const seccion = texto => {
    const p = body.appendParagraph(texto).setHeading(DocumentApp.ParagraphHeading.HEADING3);
    p.editAsText().setForegroundColor(azul).setBold(true);
  };
  const tabla = filas => {
    const datos = filas.filter(f => safe_(f[1]).trim()).map(f => [f[0], safe_(f[1])]);
    if (!datos.length) return;
    const t = body.appendTable(datos);
    t.setBorderColor('#cbd5e1');
    for (let i = 0; i < t.getNumRows(); i++) {
      t.getRow(i).getCell(0).setWidth(150).editAsText().setBold(true);
    }
  };

  seccion('Cliente');
  tabla([
    ['Nombre de fantasía', d.cliente],
    ['Razón social', d.razonSocial],
    ['CUIT', d.cuit],
    ['Cód. cliente', d.codCliente],
    ['Distribuidor', d.distribuidor],
    ['Domicilio', [d.domicilio, d.localidad].filter(Boolean).join(', ')],
    ['Comodato', d.comodatoNumero ? d.comodatoNumero + (d.fechaComodato ? ' (' + d.fechaComodato + ')' : '') : 'Sin comodato cargado']
  ]);

  seccion('Baja');
  tabla([
    ['Fecha de baja', d.fecha],
    ['Técnico', d.tecnico],
    ['Motivo', d.motivo],
    ['Observaciones', d.observaciones]
  ]);

  seccion('Retiro de chopera');
  tabla([
    ['Estado del retiro', d.retiroEquipo],
    ['Detalle / faltantes', d.detalleRetiro],
    ['Equipo', d.equipo],
    ['Pilón', d.pilon],
    ['Cantidad de picos', d.cantPicos ? String(d.cantPicos) : '']
  ]);

  seccion('Fotos');
  const pFotos = body.appendParagraph('');
  if (Array.isArray(fotos) && fotos.length) {
    fotos.forEach(photo => {
      try {
        const blob = Utilities.newBlob(Utilities.base64Decode(photo.data), photo.tipo, photo.nombre);
        const img = pFotos.appendInlineImage(blob);
        const ratio = 200 / img.getWidth();
        img.setWidth(200).setHeight(Math.round(img.getHeight() * ratio));
        pFotos.appendText('   ');
      } catch (e) {
        console.error('Error insertando foto de baja: ' + e.message);
      }
    });
  } else {
    pFotos.setText('No se adjuntaron fotografías.');
  }

  seccion('Conformidad');
  body.appendParagraph(
    'El cliente deja constancia de la baja del servicio y, en caso de corresponder, de la ' +
    'restitución del equipo y los materiales entregados en comodato, en el estado detallado en ' +
    'la presente acta. La firma digital incorporada tiene validez como constancia de conformidad.'
  );

  const pFirma = body.appendParagraph('');
  const firmaImg = pFirma.appendInlineImage(signatureFile.getBlob());
  const ratioFirma = 180 / firmaImg.getWidth();
  firmaImg.setWidth(180).setHeight(Math.round(firmaImg.getHeight() * ratioFirma));
  body.appendParagraph('Aclaración: ' + d.aclaracion + '    DNI: ' + d.dni);

  doc.saveAndClose();

  const pdfFile = pdfFolder.createFile(docFile.getAs(MimeType.PDF).setName(nombre + '.pdf'));
  compartirPorLink_(pdfFile);
  compartirPorLink_(docFile);

  return { pdfFile: pdfFile, docFile: docFile };
}

/* ==========================================================================
 * 10. MÓDULO HELADERAS (pool de tickets de heladeras rotas)
 *
 * Los tickets los carga cualquiera desde un Google Form. La hoja de
 * respuestas del Form es de solo lectura para el script: si se le agregaran
 * columnas, el Form las pisa o las corre al sumar una pregunta nueva.
 *
 * La gestion (quien lo tomo, estado, cierre, fotos) vive en la hoja
 * Tickets_Heladeras, una fila por ticket que alguien toco. Un ticket sin fila
 * ahi esta SIN ASIGNAR: es el pool.
 *
 * Estados: SIN ASIGNAR -> ASIGNADO <-> EN ESPERA -> RESUELTO | NO RESUELTO.
 * Un cerrado se puede reabrir y vuelve al pool.
 * ========================================================================== */

const HEL_ESTADOS_ACTIVOS = ['ASIGNADO', 'EN ESPERA'];
const HEL_ESTADOS_CERRADOS = ['RESUELTO', 'NO RESUELTO'];

/**
 * Como se reconoce cada dato en los encabezados del Form. Se busca por
 * palabra clave, en este orden, y cada columna se usa una sola vez: por eso
 * "fotos" y "email" van antes que "falla" o "cliente" (una pregunta
 * "Foto de la falla" es foto, "Codigo de cliente" es codigo).
 * Las preguntas que no matcheen igual se muestran en el detalle del ticket.
 */
const HEL_CAMPOS_FORM = [
  ['fotos', ['foto', 'imagen', 'adjunt']],
  ['email', ['correo', 'email', 'mail']],
  ['falla', ['falla', 'problema', 'que le pasa', 'inconveniente', 'observacion', 'descripcion', 'detalle', 'motivo']],
  ['codigo', ['numero de boca', 'boca', 'codigo', 'cod cliente', 'nro de cliente', 'numero de cliente', 'n de cliente']],
  ['cliente', ['nombre del pdv', 'nombre del cliente', 'tienda', 'pdv', 'cliente', 'nombre fantasia', 'comercio', 'razon social', 'nombre del local']],
  ['direccion', ['direccion', 'domicilio', 'calle']],
  ['localidad', ['localidad', 'ciudad', 'zona']],
  ['telefono', ['telefono', 'celular', 'whatsapp']],
  ['negocio', ['negocio', 'canal']],
  ['equipo', ['heladera', 'equipo', 'modelo', 'activo fijo', 'serie', 'marca']],
  ['supervisor', ['supervisor']],
  ['solicitante', ['repositor', 'solicitante', 'quien carga', 'vendedor', 'nombre y apellido', 'tu nombre', 'nombre']]
];

/**
 * Técnicos que pueden tomar tickets de heladeras. El servicio SMK puede no
 * ser la misma gente que las choperas: por eso tiene su propia lista.
 */
function tecnicosHeladeras_() {
  const propios = (CONFIG.HELADERAS_TECNICOS || []).filter(t => safe_(t).trim());
  return propios.length ? propios : CONFIG.TECNICOS;
}

function resolveTecnicoHeladera_(tecnico) {
  const objetivo = keyTecnico_(tecnico);
  if (!objetivo) throw new Error('Elegí tu nombre arriba.');
  const lista = tecnicosHeladeras_();
  for (let i = 0; i < lista.length; i++) {
    if (keyTecnico_(lista[i]) === objetivo) return lista[i];
  }
  throw new Error('Técnico no reconocido: ' + tecnico);
}

/** Hoja de respuestas del Form. Lanza con un mensaje claro si no la encuentra. */
function getHeladerasFormSheet_() {
  const idExterno = safe_(CONFIG.HELADERAS_FORM_SPREADSHEET_ID).trim();
  const ss = idExterno ? SpreadsheetApp.openById(idExterno) : SpreadsheetApp.getActiveSpreadsheet();

  const porNombre = ss.getSheetByName(CONFIG.HELADERAS_FORM_SHEET_NAME);
  if (porNombre) return porNombre;

  const candidata = ss.getSheets().filter(s => {
    const n = norm_(s.getName());
    return n.indexOf('respuestas de formulario') === 0 || n.indexOf('form responses') === 0;
  })[0];
  if (candidata) return candidata;

  throw new Error('No encuentro la hoja de respuestas del formulario de heladeras (' +
                  CONFIG.HELADERAS_FORM_SHEET_NAME + '). Revisá CONFIG.HELADERAS_FORM_SHEET_NAME.');
}

/**
 * Planilla de comodatos (la que contiene este script).
 *
 * No alcanza con getActiveSpreadsheet(): el disparador del Form esta
 * instalado sobre la planilla SMK, y cuando corre desde ahi la planilla
 * "activa" puede ser la del Form. Las hojas de gestion y de avisos se
 * crearian del lado equivocado. activarAvisosWhatsApp guarda el ID correcto.
 */
function ssPrincipal_() {
  const id = PropertiesService.getScriptProperties().getProperty('SS_PRINCIPAL_ID');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (err) { /* sigue con la activa */ }
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

/** Hoja de gestion de tickets; la crea con encabezados si no existe. */
function getHeladerasSheet_() {
  const ss = ssPrincipal_();
  let sheet = ss.getSheetByName(CONFIG.HELADERAS_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(CONFIG.HELADERAS_SHEET_NAME);
  escribirCabecera_(sheet, HELADERAS_HEADERS);
  return sheet;
}

/**
 * Decide que columna del Form corresponde a cada dato. Devuelve
 * { campo: indice } y, en mapa.extra, { campo: [indices] } cuando la misma
 * pregunta esta repetida: el Form SMK tiene tres "Repositor" (una por
 * supervisor) y solo una viene completa en cada respuesta.
 */
function mapearColumnasForm_(encabezados) {
  const normalizados = encabezados.map(h => norm_(h));
  const usadas = {};
  const mapa = {};

  const iTs = normalizados.findIndex(h => h === 'marca temporal' || h === 'timestamp');
  mapa.timestamp = iTs === -1 ? 0 : iTs;
  usadas[mapa.timestamp] = true;

  HEL_CAMPOS_FORM.forEach(par => {
    const campo = par[0];
    const claves = par[1];
    for (let k = 0; k < claves.length && mapa[campo] === undefined; k++) {
      for (let i = 0; i < normalizados.length; i++) {
        if (usadas[i]) continue;
        if (normalizados[i].indexOf(claves[k]) !== -1) {
          mapa[campo] = i;
          usadas[i] = true;
          break;
        }
      }
    }

    if (mapa[campo] === undefined) return;
    const repetidas = [];
    for (let i = 0; i < normalizados.length; i++) {
      if (!usadas[i] && normalizados[i] === normalizados[mapa[campo]]) {
        repetidas.push(i);
        usadas[i] = true;
      }
    }
    if (repetidas.length) {
      mapa.extra = mapa.extra || {};
      mapa.extra[campo] = repetidas;
    }
  });
  return mapa;
}

/**
 * Fecha de la marca temporal. El Form la guarda como fecha, pero si alguien
 * la pego como texto ("25/8/2026 12:49:32") new Date() la leeria con dia y
 * mes invertidos.
 */
function fechaForm_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  const m = safe_(v).trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0));
  return toDate_(v);
}

/** Hash corto y estable de un texto (4 caracteres), para desempatar IDs. */
function hashCorto_(texto) {
  let h = 5381;
  for (let i = 0; i < texto.length; i++) h = ((h * 33) ^ texto.charCodeAt(i)) >>> 0;
  return h.toString(36).toUpperCase().slice(-4);
}

function textoCeldaForm_(v) {
  if (v instanceof Date) return isoDateTime_(v);
  // Resto de una opcion sin renombrar en el Form: "Opción 1Eric Hock" -> "Eric Hock"
  return safe_(v).trim().replace(/^opci[oó]n\s*\d+\s*/i, '');
}

/**
 * Lee las respuestas del Form y les asigna un ID estable.
 *
 * El ID sale de la marca temporal (HEL-260914-103512), no del numero de fila:
 * si alguien ordena o borra filas en la hoja del Form, los tickets no se
 * mezclan con la gestion de otro. Dos respuestas en el mismo segundo llevan
 * un sufijo sacado del contenido de la fila (no del orden), por la misma razon.
 */
function leerTicketsForm_() {
  const sheet = getHeladerasFormSheet_();
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  const valores = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const encabezados = valores[0].map(h => safe_(h).trim());
  const mapa = mapearColumnasForm_(encabezados);
  const tz = Session.getScriptTimeZone();
  const tickets = [];

  // Numero de fila en la hoja de cada respuesta, para ubicar la que dispara el aviso
  const numerosFila = [];
  const filas = valores.slice(1).filter((f, i) => {
    if (f.every(v => safe_(v).trim() === '')) return false;
    numerosFila.push(i + 2);
    return true;
  });
  const bases = filas.map(f => {
    const ts = fechaForm_(f[mapa.timestamp]);
    return ts ? 'HEL-' + Utilities.formatDate(ts, tz, 'yyMMdd-HHmmss') : '';
  });
  const repetidas = {};
  bases.forEach(b => { repetidas[b] = (repetidas[b] || 0) + 1; });

  for (let r = 0; r < filas.length; r++) {
    const fila = filas[r];
    const ts = fechaForm_(fila[mapa.timestamp]);
    let id = bases[r];
    if (!id) id = 'HEL-X' + hashCorto_(fila.map(textoCeldaForm_).join('|'));
    else if (repetidas[id] > 1) id += '-' + hashCorto_(fila.map(textoCeldaForm_).join('|'));

    // Primer valor no vacio entre la columna del campo y sus repetidas
    const campo = nombre => {
      if (mapa[nombre] === undefined) return '';
      const indices = [mapa[nombre]].concat((mapa.extra && mapa.extra[nombre]) || []);
      for (let i = 0; i < indices.length; i++) {
        const v = textoCeldaForm_(fila[indices[i]]);
        if (v) return v;
      }
      return '';
    };

    const respuestas = [];
    encabezados.forEach((h, c) => {
      if (c === mapa.timestamp) return;
      const valor = textoCeldaForm_(fila[c]);
      if (valor) respuestas.push({ pregunta: h, respuesta: valor });
    });

    // Las subidas de archivo del Form llegan como links de Drive separados por coma
    const fotos = [];
    fila.forEach(v => {
      (safe_(v).match(/https?:\/\/drive\.google\.com[^\s,]+/g) || []).forEach(u => {
        if (fotos.indexOf(u) === -1) fotos.push(u);
      });
    });

    tickets.push({
      id: id,
      filaForm: numerosFila[r],
      creado: ts ? isoDateTime_(ts) : '',
      cliente: campo('cliente'),
      codigo: campo('codigo'),
      // Como se llama el codigo en el Form ("Numero de Boca"), para mostrarlo igual
      codigoEtiqueta: mapa.codigo === undefined ? '' : encabezados[mapa.codigo],
      direccion: campo('direccion'),
      localidad: campo('localidad'),
      telefono: campo('telefono'),
      negocio: campo('negocio'),
      equipo: campo('equipo'),
      falla: campo('falla'),
      supervisor: campo('supervisor'),
      solicitante: campo('solicitante'),
      email: campo('email'),
      fotosForm: fotos,
      respuestas: respuestas
    });
  }
  return tickets;
}

/** Filas de gestion como mapa { TicketId: objeto con _row }. */
function leerGestionHeladeras_() {
  const sheet = getHeladerasSheet_();
  const lastRow = sheet.getLastRow();
  const mapa = {};
  if (lastRow < 2) return mapa;

  sheet.getRange(2, 1, lastRow - 1, HELADERAS_HEADERS.length).getValues().forEach((fila, i) => {
    const obj = { _row: i + 2 };
    HELADERAS_HEADERS.forEach((h, c) => { obj[h] = fila[c]; });
    const id = safe_(obj.TicketId).trim();
    if (id) mapa[id] = obj;
  });
  return mapa;
}

function estadoHeladera_(gestion) {
  const e = gestion ? safe_(gestion.Estado).trim().toUpperCase() : '';
  return e || 'SIN ASIGNAR';
}

/** Ticket del Form + su gestion, listo para la app. */
function armarTicketHeladera_(t, g) {
  return Object.assign({}, t, {
    estado: estadoHeladera_(g),
    tecnico: g ? safe_(g.Tecnico) : '',
    tomadoEl: g ? isoDateTime_(toDate_(g.TomadoEl)) : '',
    cerradoEl: g ? isoDateTime_(toDate_(g.CerradoEl)) : '',
    trabajo: g ? safe_(g.TrabajoRealizado) : '',
    repuestos: g ? safe_(g.Repuestos) : '',
    motivoNoResuelto: g ? safe_(g.MotivoNoResuelto) : '',
    notaEspera: g ? safe_(g.NotaEspera) : '',
    fotosCierre: g ? safe_(g.FotosUrls).split('\n').filter(String) : [],
    historial: g ? safe_(g.Historial).split('\n').filter(String) : []
  });
}

/**
 * Todos los tickets abiertos y los cerrados recientes. Con "historico" van
 * tambien los cerrados viejos: lo usa el tablero general.
 */
function getTicketsHeladeras_(historico) {
  const gestion = leerGestionHeladeras_();
  const limite = addDays_(new Date(), -CONFIG.HELADERAS_DIAS_CERRADOS_VISIBLES);

  return leerTicketsForm_()
    .map(t => armarTicketHeladera_(t, gestion[t.id]))
    .filter(t => {
      if (historico || HEL_ESTADOS_CERRADOS.indexOf(t.estado) === -1) return true;
      const g = gestion[t.id];
      const cierre = g ? toDate_(g.CerradoEl) : null;
      return !cierre || cierre >= limite;
    });
}

function lineaHistorial_(tecnico, texto) {
  const cuando = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  return cuando + ' · ' + tecnico + ' · ' + texto;
}

/** Escribe (o crea) la fila de gestion del ticket con los cambios dados. */
function guardarGestionHeladera_(ticket, previa, cambios, textoHistorial, tecnico) {
  const sheet = getHeladerasSheet_();
  const base = {};
  HELADERAS_HEADERS.forEach(h => { base[h] = previa ? previa[h] : ''; });

  Object.assign(base, cambios);
  base.TicketId = ticket.id;
  base.Cliente = ticket.cliente;
  base.Actualizado = new Date();
  base.Historial = [safe_(previa ? previa.Historial : '').trim(), lineaHistorial_(tecnico, textoHistorial)]
    .filter(String).join('\n');

  const fila = HELADERAS_HEADERS.map(h => base[h] === undefined || base[h] === null ? '' : base[h]);
  if (previa) {
    sheet.getRange(previa._row, 1, 1, HELADERAS_HEADERS.length).setValues([fila]);
  } else {
    sheet.appendRow(fila);
  }
}

/** Busca el ticket y su gestion actual. Lanza si el ID no existe en el Form. */
function buscarTicketHeladera_(id) {
  const limpio = safe_(id).trim();
  if (!limpio) throw new Error('Falta el ID del ticket.');
  const ticket = leerTicketsForm_().filter(t => t.id === limpio)[0];
  if (!ticket) throw new Error('No encontré el ticket ' + limpio + '. Actualizá la lista.');
  const previa = leerGestionHeladeras_()[limpio] || null;
  return { ticket: ticket, previa: previa, estado: estadoHeladera_(previa) };
}

/** Controla que el ticket este activo y lo tenga el tecnico que opera. */
function exigirTicketPropio_(actual, tecnico) {
  if (HEL_ESTADOS_ACTIVOS.indexOf(actual.estado) === -1) {
    throw new Error('El ticket está ' + actual.estado.toLowerCase() + '. Actualizá la lista.');
  }
  if (keyTecnico_(actual.previa.Tecnico) !== keyTecnico_(tecnico)) {
    throw new Error('Este ticket lo tiene ' + safe_(actual.previa.Tecnico) + '.');
  }
}

/**
 * Acciones sobre un ticket: tomar, liberar, espera, retomar, reabrir.
 * Todo bajo lock: dos tecnicos tocando "Tomar" a la vez no pueden quedarse
 * los dos con el mismo ticket.
 */
function heladeraAccion_(body) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const tecnico = resolveTecnicoHeladera_(body.tecnico);
    const accion = safe_(body.accion).trim();
    const nota = safe_(body.nota).trim();
    const actual = buscarTicketHeladera_(body.ticketId);
    const t = actual.ticket;

    if (accion === 'tomar') {
      if (actual.estado !== 'SIN ASIGNAR') {
        const quien = safe_(actual.previa && actual.previa.Tecnico);
        if (HEL_ESTADOS_ACTIVOS.indexOf(actual.estado) !== -1 && keyTecnico_(quien) === keyTecnico_(tecnico)) {
          return { ok: true, message: 'Ya tenías este ticket.' };
        }
        return { ok: false, message: quien ? 'Llegaste tarde: lo tomó ' + quien + '.' : 'El ticket ya no está disponible.' };
      }
      guardarGestionHeladera_(t, actual.previa, {
        Estado: 'ASIGNADO', Tecnico: tecnico, TomadoEl: new Date(), CerradoEl: '',
        TrabajoRealizado: '', Repuestos: '', MotivoNoResuelto: '', NotaEspera: '', FotosUrls: ''
      }, 'Tomó el ticket', tecnico);
      return { ok: true, message: 'Ticket asignado a vos.' };
    }

    if (accion === 'liberar') {
      exigirTicketPropio_(actual, tecnico);
      guardarGestionHeladera_(t, actual.previa, {
        Estado: 'SIN ASIGNAR', Tecnico: '', TomadoEl: '', NotaEspera: ''
      }, 'Lo devolvió al pool' + (nota ? ': ' + nota : ''), tecnico);
      return { ok: true, message: 'Ticket devuelto al pool.' };
    }

    if (accion === 'espera') {
      exigirTicketPropio_(actual, tecnico);
      if (!nota) throw new Error('Contá qué se está esperando (repuesto, turno, etc.).');
      guardarGestionHeladera_(t, actual.previa, { Estado: 'EN ESPERA', NotaEspera: nota },
        'En espera: ' + nota, tecnico);
      return { ok: true, message: 'Ticket en espera.' };
    }

    if (accion === 'retomar') {
      exigirTicketPropio_(actual, tecnico);
      guardarGestionHeladera_(t, actual.previa, { Estado: 'ASIGNADO', NotaEspera: '' },
        'Lo retomó', tecnico);
      return { ok: true, message: 'Ticket retomado.' };
    }

    if (accion === 'reabrir') {
      if (HEL_ESTADOS_CERRADOS.indexOf(actual.estado) === -1) throw new Error('Solo se reabre un ticket cerrado.');
      if (!nota) throw new Error('Contá por qué se reabre.');
      guardarGestionHeladera_(t, actual.previa, {
        Estado: 'SIN ASIGNAR', Tecnico: '', TomadoEl: '', CerradoEl: '', NotaEspera: ''
      }, 'Reabrió el ticket (' + actual.estado.toLowerCase() + '): ' + nota, tecnico);
      return { ok: true, message: 'Ticket reabierto: volvió al pool.' };
    }

    throw new Error('Acción no válida: ' + accion);
  } finally {
    lock.releaseLock();
  }
}

/** Cierra el ticket como RESUELTO o NO RESUELTO, con detalle y fotos. */
function heladeraFinalizar_(body) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const tecnico = resolveTecnicoHeladera_(body.tecnico);
    const actual = buscarTicketHeladera_(body.ticketId);
    exigirTicketPropio_(actual, tecnico);

    const resuelto = safe_(body.resultado).trim().toUpperCase() === 'RESUELTO';
    const trabajo = safe_(body.trabajo).trim();
    const motivo = safe_(body.motivo).trim();
    if (resuelto && !trabajo) throw new Error('Contá qué se le hizo al equipo.');
    if (!resuelto && !motivo) throw new Error('Elegí por qué no se pudo resolver.');

    let urls = '';
    if (body.fotos && body.fotos.length) {
      urls = savePhotos_(getHeladeraPhotosFolder_(actual.ticket.id), body.fotos, actual.ticket.id);
    }

    const estado = resuelto ? 'RESUELTO' : 'NO RESUELTO';
    guardarGestionHeladera_(actual.ticket, actual.previa, {
      Estado: estado,
      CerradoEl: new Date(),
      TrabajoRealizado: trabajo,
      Repuestos: safe_(body.repuestos).trim(),
      MotivoNoResuelto: resuelto ? '' : motivo,
      NotaEspera: '',
      FotosUrls: urls
    }, resuelto ? 'Lo cerró como resuelto' : 'Lo cerró sin resolver: ' + motivo, tecnico);

    return { ok: true, message: resuelto ? 'Ticket resuelto.' : 'Ticket cerrado como no resuelto.', estado: estado };
  } finally {
    lock.releaseLock();
  }
}

/** Carpeta de fotos de heladeras, con subcarpeta por ticket. */
function getHeladeraPhotosFolder_(ticketId) {
  const base = getPhotosParentFolder_();
  const raiz = base.getFoldersByName(CONFIG.HELADERAS_PHOTOS_FOLDER_NAME);
  const parent = raiz.hasNext() ? raiz.next() : base.createFolder(CONFIG.HELADERAS_PHOTOS_FOLDER_NAME);
  const sub = parent.getFoldersByName(ticketId);
  return sub.hasNext() ? sub.next() : parent.createFolder(ticketId);
}

/**
 * Prueba desde el editor: muestra como se mapearon las columnas del Form y
 * los ultimos tickets. Correrla una vez despues de vincular el Form.
 */
function probarHeladeras() {
  PropertiesService.getScriptProperties()
    .setProperty('SS_PRINCIPAL_ID', SpreadsheetApp.getActiveSpreadsheet().getId());
  const sheet = getHeladerasFormSheet_();
  const encabezados = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const mapa = mapearColumnasForm_(encabezados);
  const log = ['Planilla del Form: ' + sheet.getParent().getName() + ' / ' + sheet.getName(), 'Columnas detectadas:'];
  Object.keys(mapa).forEach(k => {
    if (k === 'extra') return;
    const repetidas = mapa.extra && mapa.extra[k] ? ' (+' + mapa.extra[k].length + ' repetida/s)' : '';
    log.push('  ' + k + ' -> ' + encabezados[mapa[k]] + repetidas);
  });
  log.push('Técnicos de heladeras: ' + tecnicosHeladeras_().join(', '));
  leerTicketsForm_().slice(-3).forEach(t => log.push(t.id + ' · ' + t.cliente + ' · ' + t.falla));
  Logger.log(log.join('\n'));
}

/* ==========================================================================
 * 11. AVISOS POR WHATSAPP
 *
 * Cuando entra una respuesta al Form de heladeras, un disparador instalable
 * (onFormSubmit) le manda un WhatsApp a cada destinatario activo de la hoja
 * Avisos_WhatsApp. El resultado de cada envio queda escrito en esa misma
 * hoja, porque un disparador que falla no se ve en ningun otro lado.
 *
 * Puesta en marcha (una vez, desde el editor):
 *   1) Cargar los destinatarios en Avisos_WhatsApp (se crea sola).
 *   2) Correr activarAvisosWhatsApp  -> instala el disparador.
 *   3) Correr probarWhatsApp         -> manda un mensaje de prueba.
 *
 * Proveedores:
 *   - callmebot: cada destinatario le manda "I allow callmebot to send me
 *     messages" al +34 644 66 32 62 desde su WhatsApp y recibe su apikey, que
 *     se pega en la columna ApiKey. Gratis, pensado para uso personal.
 *   - meta: API oficial de WhatsApp Cloud. Necesita en Propiedades del Script
 *     WHATSAPP_META_TOKEN y WHATSAPP_META_PHONE_ID, y una plantilla aprobada
 *     con 5 variables: {{1}} ID, {{2}} cliente, {{3}} ubicacion, {{4}} falla,
 *     {{5}} link. Los tokens NO van en el codigo: el repo esta en GitHub.
 * ========================================================================== */

const WHATSAPP_HEADERS = ['Nombre', 'Telefono', 'ApiKey', 'Activo', 'UltimoEnvio', 'UltimoResultado'];

function getWhatsAppSheet_() {
  const ss = ssPrincipal_();
  let sheet = ss.getSheetByName(CONFIG.WHATSAPP_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(CONFIG.WHATSAPP_SHEET_NAME);
  escribirCabecera_(sheet, WHATSAPP_HEADERS);
  return sheet;
}

/** Destinatarios activos, con la fila para anotarles el resultado. */
function leerDestinatariosWhatsApp_() {
  const sheet = getWhatsAppSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, WHATSAPP_HEADERS.length).getValues()
    .map((f, i) => ({
      fila: i + 2,
      nombre: safe_(f[0]).trim(),
      telefono: safe_(f[1]).replace(/[^\d]/g, ''),
      apikey: safe_(f[2]).trim(),
      activo: norm_(f[3]) !== 'no'
    }))
    .filter(d => d.activo && d.telefono);
}

function linkTicketApp_(ticketId) {
  const base = safe_(CONFIG.APP_URL).trim();
  if (!base) return '';
  return base + (base.indexOf('?') === -1 ? '?' : '&') + 'vista=heladeras&ticket=' + encodeURIComponent(ticketId);
}

/** "TOLEDO · Boca 903045": el nombre del PDV con su codigo, si lo hay. */
function nombreConCodigo_(t) {
  const etiqueta = norm_(t.codigoEtiqueta).indexOf('boca') !== -1 ? 'Boca' : (t.codigoEtiqueta || 'Cód.');
  return (t.cliente || 'PDV sin nombre') + (t.codigo ? ' · ' + etiqueta + ' ' + t.codigo : '');
}

function textoAvisoTicket_(t) {
  const ubicacion = [t.direccion, t.localidad].filter(String).join(', ');
  const link = linkTicketApp_(t.id);
  const quien = [t.solicitante, t.supervisor ? 'sup. ' + t.supervisor : ''].filter(String).join(' · ');
  return [
    '🧊 *Nuevo ticket de heladera*' + (t.negocio ? ' (' + t.negocio + ')' : ''),
    t.id,
    '🏪 ' + nombreConCodigo_(t),
    ubicacion ? '📍 ' + ubicacion : '',
    t.falla ? '⚠️ ' + t.falla : '',
    quien ? '👤 ' + quien : '',
    link ? '\nTomalo acá: ' + link : 'Tomalo desde la solapa Heladeras de la app.'
  ].filter(String).join('\n');
}

/** Arma el pedido HTTP de un envio segun el proveedor configurado. */
function pedidoWhatsApp_(destinatario, ticket, textoLibre) {
  const proveedor = safe_(CONFIG.WHATSAPP_PROVEEDOR).trim().toLowerCase();

  if (proveedor === 'callmebot') {
    if (!destinatario.apikey) throw new Error('falta la ApiKey de CallMeBot');
    return {
      url: 'https://api.callmebot.com/whatsapp.php?phone=%2B' + destinatario.telefono +
           '&text=' + encodeURIComponent(textoLibre) + '&apikey=' + encodeURIComponent(destinatario.apikey),
      method: 'get',
      muteHttpExceptions: true
    };
  }

  if (proveedor === 'meta') {
    const props = PropertiesService.getScriptProperties();
    const token = props.getProperty('WHATSAPP_META_TOKEN');
    const phoneId = props.getProperty('WHATSAPP_META_PHONE_ID');
    if (!token || !phoneId) throw new Error('faltan WHATSAPP_META_TOKEN / WHATSAPP_META_PHONE_ID en Propiedades del Script');

    // Las variables de plantilla no aceptan saltos de linea ni tiras de espacios
    const limpio = v => safe_(v).replace(/\s+/g, ' ').trim().slice(0, 900) || '-';
    const t = ticket || {};
    return {
      url: 'https://graph.facebook.com/v21.0/' + phoneId + '/messages',
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true,
      payload: JSON.stringify({
        messaging_product: 'whatsapp',
        to: destinatario.telefono,
        type: 'template',
        template: {
          name: CONFIG.WHATSAPP_META_PLANTILLA,
          language: { code: CONFIG.WHATSAPP_META_IDIOMA },
          components: [{
            type: 'body',
            parameters: [
              t.id, nombreConCodigo_(t),
              [t.direccion, t.localidad].filter(String).join(', '),
              t.falla, linkTicketApp_(t.id || '')
            ].map(v => ({ type: 'text', text: limpio(v) }))
          }]
        }
      })
    };
  }

  throw new Error('proveedor de WhatsApp no configurado');
}

/** CallMeBot responde 200 aun con errores; el detalle viene en el texto. */
function envioOk_(respuesta) {
  const codigo = respuesta.getResponseCode();
  const cuerpo = safe_(respuesta.getContentText());
  if (codigo < 200 || codigo >= 300) return false;
  return !/error|invalid|not (been )?activated/i.test(cuerpo.slice(0, 2000));
}

/**
 * Manda el aviso de un ticket a todos los destinatarios, en paralelo, y
 * anota el resultado de cada uno en la hoja. Devuelve cuantos salieron bien.
 */
function enviarAvisoWhatsApp_(ticket, textoLibre) {
  if (!safe_(CONFIG.WHATSAPP_PROVEEDOR).trim()) return { enviados: 0, fallidos: 0, apagado: true };

  const sheet = getWhatsAppSheet_();
  const destinatarios = leerDestinatariosWhatsApp_();
  const texto = textoLibre || textoAvisoTicket_(ticket);
  const ahora = new Date();

  const pedidos = [];
  const errores = {};
  destinatarios.forEach((d, i) => {
    try {
      pedidos.push({ i: i, pedido: pedidoWhatsApp_(d, ticket, texto) });
    } catch (err) {
      errores[i] = err.message;
    }
  });

  const respuestas = pedidos.length ? UrlFetchApp.fetchAll(pedidos.map(p => p.pedido)) : [];
  const resultado = {};
  pedidos.forEach((p, n) => {
    const r = respuestas[n];
    resultado[p.i] = envioOk_(r)
      ? 'OK'
      : 'ERROR ' + r.getResponseCode() + ': ' + safe_(r.getContentText()).replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').slice(0, 200);
  });

  let enviados = 0, fallidos = 0;
  destinatarios.forEach((d, i) => {
    const txt = errores[i] ? 'ERROR: ' + errores[i] : resultado[i];
    if (txt === 'OK') enviados++; else fallidos++;
    sheet.getRange(d.fila, WHATSAPP_HEADERS.indexOf('UltimoEnvio') + 1, 1, 2)
      .setValues([[ahora, (ticket ? ticket.id + ' · ' : '') + txt]]);
  });

  return { enviados: enviados, fallidos: fallidos };
}

/**
 * Disparador onFormSubmit. Solo reacciona a la hoja del Form de heladeras.
 *
 * Google a veces dispara onFormSubmit dos veces para la misma respuesta: el
 * CacheService evita mandar el mismo WhatsApp dos veces.
 */
function alEnviarFormHeladera(e) {
  if (!e || !e.range) return;

  const hojaForm = getHeladerasFormSheet_();
  const hojaEvento = e.range.getSheet();
  if (hojaEvento.getSheetId() !== hojaForm.getSheetId() ||
      hojaEvento.getParent().getId() !== hojaForm.getParent().getId()) return;

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const fila = e.range.getRow();
    const ticket = leerTicketsForm_().filter(t => t.filaForm === fila)[0];
    if (!ticket) return;

    const cache = CacheService.getScriptCache();
    const clave = 'wa_aviso_' + ticket.id;
    if (cache.get(clave)) return;
    cache.put(clave, '1', 21600);

    compartirFotosTicket_(ticket);
    const r = enviarAvisoWhatsApp_(ticket);
    Logger.log('Aviso ' + ticket.id + ': ' + JSON.stringify(r));
  } finally {
    lock.releaseLock();
  }
}

/**
 * Las fotos que se suben por el Form nacen privadas del dueño del Form: sin
 * esto, el tecnico que toca "Foto del reclamo" en el celular ve "Solicitar
 * acceso". Si la cuenta del script no puede compartirlas, no corta nada.
 */
function compartirFotosTicket_(ticket) {
  let ok = 0;
  (ticket.fotosForm || []).forEach(url => {
    const id = idDesdeUrlDrive_(url);
    if (!id) return;
    try {
      if (compartirPorLink_(DriveApp.getFileById(id))) ok++;
    } catch (err) {
      Logger.log('No se pudo abrir la foto ' + id + ': ' + err.message);
    }
  });
  return ok;
}

/** Comparte las fotos de todos los tickets ya cargados. Uso unico desde el editor. */
function compartirFotosFormHeladeras() {
  let ok = 0, total = 0;
  leerTicketsForm_().forEach(t => {
    total += (t.fotosForm || []).length;
    ok += compartirFotosTicket_(t);
  });
  Logger.log('Fotos del Form compartidas: ' + ok + ' de ' + total);
}

/** Instala (o reinstala) el disparador del Form. Correr una vez desde el editor. */
function activarAvisosWhatsApp() {
  // Corriendo desde el editor, la activa es la planilla de comodatos: se guarda
  // para que el disparador (que corre desde la planilla del Form) la encuentre.
  PropertiesService.getScriptProperties()
    .setProperty('SS_PRINCIPAL_ID', SpreadsheetApp.getActiveSpreadsheet().getId());

  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'alEnviarFormHeladera')
    .forEach(t => ScriptApp.deleteTrigger(t));

  const ssForm = getHeladerasFormSheet_().getParent();
  ScriptApp.newTrigger('alEnviarFormHeladera').forSpreadsheet(ssForm).onFormSubmit().create();
  getWhatsAppSheet_();

  Logger.log('Disparador instalado sobre "' + ssForm.getName() + '". Destinatarios activos: ' +
             leerDestinatariosWhatsApp_().length);
}

/** Manda un mensaje de prueba a todos los destinatarios activos. */
function probarWhatsApp() {
  const ultimo = leerTicketsForm_().slice(-1)[0];
  const ticket = ultimo || {
    id: 'HEL-PRUEBA', cliente: 'PDV de prueba', direccion: '', localidad: '',
    falla: 'Mensaje de prueba del sistema', solicitante: ''
  };
  const texto = CONFIG.WHATSAPP_PROVEEDOR === 'meta' ? null : '✅ Prueba de avisos de heladeras\n\n' + textoAvisoTicket_(ticket);
  const r = enviarAvisoWhatsApp_(ticket, texto);
  Logger.log('Prueba de WhatsApp: ' + JSON.stringify(r) + '. Mirá la hoja ' + CONFIG.WHATSAPP_SHEET_NAME + ' para el detalle.');
}

/* ==========================================================================
 * 12. TABLEROS GENERALES (todos los tecnicos en una sola tabla)
 *
 * "Mis Comodatos" y "Mis Choperas" muestran la cartera de UN tecnico. Estas
 * dos funciones arman la misma informacion para los seis juntos: una lista
 * plana con el tecnico como columna, mas un resumen por tecnico para comparar.
 *
 * Las hojas se leen una sola vez y la cartera de cada tecnico se arma en
 * memoria, igual que getResumenGeneral_: leer por tecnico multiplicaba por
 * seis las llamadas al Sheet y se iba de tiempo.
 * ========================================================================== */

/**
 * Nombre canonico del tecnico si esta en CONFIG.TECNICOS; si no, el nombre tal
 * como quedo cargado. A diferencia de resolveTecnico_ no lanza: los tableros
 * generales tienen que poder mostrar tambien filas de tecnicos que ya no estan
 * en la lista.
 */
function nombreTecnico_(valor) {
  const key = keyTecnico_(valor);
  if (!key) return '';
  for (let i = 0; i < CONFIG.TECNICOS.length; i++) {
    if (keyTecnico_(CONFIG.TECNICOS[i]) === key) return CONFIG.TECNICOS[i];
  }
  return safe_(valor).trim();
}

/** Dias transcurridos desde una fecha, o null si no se puede leer. */
function diasDesdeFecha_(valor) {
  const d = toDate_(valor);
  return d ? diffDays_(d, new Date()) : null;
}

/** Fila vacia del resumen de comodatos de un tecnico. */
function filaResumenComodatos_(tecnico) {
  return {
    tecnico: tecnico, comodatos: 0, ultimos30: 0, ultimos90: 0, picos: 0,
    clientes: 0, ultimoNumero: '', ultimaFecha: '', diasSinCargar: null, _clientes: {}
  };
}

/**
 * Tablero general de comodatos: todos los comodatos cargados, de todos los
 * tecnicos, mas un resumen por tecnico.
 */
function getComodatosGeneral_() {
  const mapa = leerComodatosAgrupados_();

  const lista = [];
  const resumen = {};
  const clientesTotales = {};

  // Los seis tecnicos siempre aparecen, aunque no tengan comodatos cargados
  CONFIG.TECNICOS.forEach(t => { resumen[keyTecnico_(t)] = filaResumenComodatos_(t); });

  Object.keys(mapa).forEach(key => {
    mapa[key].forEach(c => {
      const tecnico = nombreTecnico_(c.tecnico) || nombreTecnico_(key);
      const cliente = c.nombreFantasia || c.razonSocial || '';

      const fila = {};
      for (const p in c) fila[p] = c[p];
      fila.tecnico = tecnico;
      fila.cliente = cliente;
      lista.push(fila);

      if (cliente) clientesTotales[norm_(cliente)] = true;

      if (!resumen[key]) resumen[key] = filaResumenComodatos_(tecnico);
      const r = resumen[key];
      r.comodatos++;
      r.picos += toNumber_(c.cantPicos);
      if (cliente) r._clientes[norm_(cliente)] = true;

      const dias = diasDesdeFecha_(c.fecha);
      if (dias !== null && dias >= 0) {
        if (dias <= 30) r.ultimos30++;
        if (dias <= 90) r.ultimos90++;
      }
    });

    // mapa[key] ya viene de la mas reciente a la mas vieja
    const ultimo = mapa[key][0];
    const r = resumen[key];
    if (ultimo && r) {
      r.ultimoNumero = ultimo.comodatoNumero;
      r.ultimaFecha = isoDate_(toDate_(ultimo.fecha)) || safe_(ultimo.fecha);
      r.diasSinCargar = diasDesdeFecha_(ultimo.fecha);
    }
  });

  lista.sort((a, b) => b.timestamp - a.timestamp);

  const porTecnico = Object.keys(resumen).map(k => {
    const r = resumen[k];
    r.clientes = Object.keys(r._clientes).length;
    delete r._clientes;
    return r;
  }).sort((a, b) => b.comodatos - a.comodatos);

  return {
    comodatos: lista,
    porTecnico: porTecnico,
    totales: {
      comodatos: lista.length,
      tecnicos: porTecnico.filter(t => t.comodatos > 0).length,
      clientes: Object.keys(clientesTotales).length,
      picos: porTecnico.reduce((a, t) => a + t.picos, 0),
      ultimos30: porTecnico.reduce((a, t) => a + t.ultimos30, 0),
      ultimos90: porTecnico.reduce((a, t) => a + t.ultimos90, 0)
    }
  };
}

/**
 * Tablero general de choperas: la cartera de los seis tecnicos en una sola
 * lista, con estado de comodato y de sanitizacion, mas un resumen por tecnico.
 *
 * La lista no trae la ficha del comodato ni el detalle de cada intervencion
 * (solo los contadores): son seis carteras juntas y el payload se iba a varios
 * MB. Para el detalle de un cliente esta la vista por tecnico.
 */
function getChoperasGeneral_() {
  const comodatos = leerComodatosAgrupados_();
  const clientes = leerClientesManualesAgrupados_();
  const bajas = leerBajasVigentesAgrupadas_();

  // Una sola lectura de Intervenciones para los seis, agrupada por tecnico
  const intervencionesPorTecnico = {};
  readIntervenciones_('').forEach(i => {
    const k = keyTecnico_(i.Tecnico);
    if (!k) return;
    if (!intervencionesPorTecnico[k]) intervencionesPorTecnico[k] = [];
    intervencionesPorTecnico[k].push(i);
  });

  const lista = [];

  const porTecnico = CONFIG.TECNICOS.map(tecnico => {
    const key = keyTecnico_(tecnico);
    const filas = readSanitRows_(tecnico);
    const cartera = construirCartera_(
      tecnico, comodatos[key] || [], clientes[key] || [], filas, bajas[key] || {}
    );
    const suyas = armarChoperas_(cartera, intervencionesPorTecnico[key] || []);

    suyas.forEach(c => {
      lista.push({
        tecnico: tecnico,
        cliente: c.cliente,
        direccion: c.direccion,
        localidad: c.localidad,
        origen: c.origen,
        tieneComodato: c.tieneComodato,
        comodatoNumero: c.comodatoNumero,
        pdfUrl: c.pdfUrl,
        equipo: c.equipo,
        pilon: c.pilon,
        cantPicos: c.cantPicos,
        estadoSanitizacion: c.estadoSanitizacion,
        sanitizado: c.sanitizado,
        ultimaSanitizacion: c.ultimaSanitizacion,
        proximaSanitizacion: c.proximaSanitizacion,
        diasRestantes: c.diasRestantes,
        visitasFallidas: c.visitasFallidas,
        motivoUltimaFalla: c.motivoUltimaFalla,
        intervencionesPendientes: c.intervencionesPendientes,
        intervencionesTotal: c.intervencionesTotal
      });
    });

    // Bajas vigentes: las del tecnico que hoy no estan en su cartera
    const enCartera = {};
    cartera.forEach(c => { enCartera[norm_(c.cliente)] = true; });
    const bajasVigentes = Object.keys(bajas[key] || {}).filter(ck => !enCartera[ck]).length;

    const contar = filtro => suyas.filter(filtro).length;
    let peor = null;
    suyas.forEach(c => {
      if (c.diasRestantes === null || c.diasRestantes === undefined) return;
      if (peor === null || c.diasRestantes < peor) peor = c.diasRestantes;
    });

    return {
      tecnico: tecnico,
      choperas: suyas.length,
      sinComodato: contar(c => !c.tieneComodato),
      conComodato: contar(c => c.tieneComodato),
      sanitizadas: contar(c => c.sanitizado),
      sinSanitizar: contar(c => !c.sanitizado),
      vencidas: contar(c => c.estadoSanitizacion === 'VENCIDO'),
      porVencer: contar(c => c.estadoSanitizacion === 'POR VENCER'),
      sinRegistro: contar(c => c.estadoSanitizacion === 'SIN REGISTRO'),
      pendientes: contar(c => c.intervencionesPendientes > 0),
      picos: suyas.reduce((a, c) => a + toNumber_(c.cantPicos), 0),
      bajas: bajasVigentes,
      diasMasAtrasado: peor
    };
  }).sort((a, b) => (b.sinSanitizar - a.sinSanitizar) || (b.choperas - a.choperas));

  const sumar = campo => porTecnico.reduce((a, t) => a + t[campo], 0);

  return {
    choperas: lista,
    porTecnico: porTecnico,
    totales: {
      choperas: lista.length,
      sinComodato: sumar('sinComodato'),
      conComodato: sumar('conComodato'),
      sanitizadas: sumar('sanitizadas'),
      sinSanitizar: sumar('sinSanitizar'),
      vencidas: sumar('vencidas'),
      porVencer: sumar('porVencer'),
      sinRegistro: sumar('sinRegistro'),
      pendientes: sumar('pendientes'),
      picos: sumar('picos'),
      bajas: sumar('bajas')
    }
  };
}
