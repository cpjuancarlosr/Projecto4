/**
 * @OnlyCurrentDoc
 *
 * Versión Ligera · Sistema contable en Google Sheets (CFDI MX)
 * Autor: JC
 * Fecha: 26/09/2025
 * Meta: Plantilla simple, pocas hojas, operaciones clave completas (CFDI→Pólizas, CxC/CxP, banco, IVA) con buena velocidad.
 */

/**
 * Crea el menú personalizado en la UI de Google Sheets al abrir el archivo.
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Jefatura Contable')
    .addItem('Cargar XML…','JC_showPicker')
    .addItem('Generar Previa','JC_buildPreview')
    .addItem('Emitir Pólizas','JC_emitirPolizas')
    .addSeparator()
    .addItem('Importar Banco (CSV/PDF)','JC_importBanco')
    .addItem('Conciliación','JC_runConciliacion')
    .addSeparator()
    .addItem('Refrescar Reporte','JC_refresh')
    .addToUi();
}

// --- PLACEHOLDERS PARA FUNCIONES DEL MENÚ ---

function JC_showPicker() {
  // Esta función debería mostrar el Google Picker para seleccionar archivos XML de Drive.
  // Por ahora, solo muestra una alerta.
  SpreadsheetApp.getUi().alert('Función "Cargar XML" no implementada. Aquí se abriría el selector de archivos.');
  // Ejemplo de implementación futura:
  // const html = HtmlService.createHtmlOutputFromFile('picker.html').setWidth(600).setHeight(425).setSandboxMode(HtmlService.SandboxMode.IFRAME);
  // SpreadsheetApp.getUi().showModalDialog(html, 'Seleccionar archivos XML');
}

function JC_buildPreview() {
  SpreadsheetApp.getUi().alert('Función "Generar Previa" no implementada. Aquí se crearían los asientos en estado "Borrador".');
}

function JC_emitirPolizas() {
  SpreadsheetApp.getUi().alert('Función "Emitir Pólizas" no implementada. Aquí se cambiaría el estado de "Borrador" a "Emitida".');
}

function JC_importBanco() {
  SpreadsheetApp.getUi().alert('Función "Importar Banco" no implementada. Aquí se procesaría un archivo CSV o PDF.');
}

function JC_runConciliacion() {
  SpreadsheetApp.getUi().alert('Función "Conciliación" no implementada. Aquí se cruzarían los movimientos de banco con las pólizas.');
}

function JC_refresh() {
  SpreadsheetApp.getUi().alert('Función "Refrescar Reporte" no implementada. Forzaría el recálculo de las fórmulas de la hoja Reporte.');
  // Nota: Generalmente las fórmulas se actualizan solas. Esto podría ser para scripts o procesos más complejos.
}


// --- LÓGICA DE PROCESAMIENTO DE CFDI (PROPORCIONADA EN LA ESPECIFICACIÓN) ---

/**
 * Procesa una lista de IDs de archivos de Google Drive, extrayendo datos de los XML.
 * @param {string[]} ids - Un array de IDs de archivos de Drive.
 */
function JC_parseFilesById_(ids){
  const sh = ss_('CFDI');
  if (!sh) {
    SpreadsheetApp.getUi().alert('Error: No se encuentra la hoja "CFDI". Por favor, créala.');
    return;
  }
  const rows = [];
  ids.forEach(id=>{
    try {
      const file = DriveApp.getFileById(id);
      if (file.getMimeType().indexOf('xml')===-1) return;
      const xml = XmlService.parse(file.getBlob().getDataAsString('UTF-8'));
      const d = JC_extractCFDI_(xml.getRootElement());
      d.ArchivoXML_ID = id; // Guardar el ID del archivo
      rows.push(JC_toRow_(d));
    } catch (e) {
      console.error(`Error procesando archivo ${id}: ${e.toString()}`);
    }
  });
  if(rows.length) {
    sh.getRange(sh.getLastRow()+1, 1, rows.length, rows[0].length).setValues(rows);
    SpreadsheetApp.getUi().alert(`${rows.length} CFDI(s) importado(s) correctamente.`);
  } else {
    SpreadsheetApp.getUi().alert('No se importaron nuevos CFDI. Verifique que los archivos seleccionados sean XML válidos.');
  }
}

/**
 * Extrae los datos principales de un elemento raíz de un XML de CFDI 4.0.
 * @param {XmlService.Element} root - El elemento raíz del documento XML.
 * @return {Object} Un objeto con los datos extraídos.
 */
function JC_extractCFDI_(root){
  const ns = {
    cfdi: XmlService.getNamespace('cfdi','http://www.sat.gob.mx/cfd/4'),
    tfd:  XmlService.getNamespace('tfd','http://www.sat.gob.mx/TimbreFiscalDigital'),
    p20:  XmlService.getNamespace('pago20','http://www.sat.gob.mx/Pagos20')
  };
  const g = (el,att)=> el?.getAttribute(att)?.getValue()||'';

  const comp = root;
  const em = comp.getChild('Emisor',ns.cfdi);
  const re = comp.getChild('Receptor',ns.cfdi);
  const complemento = comp.getChild('Complemento', ns.cfdi);
  const tim = complemento?.getChild('TimbreFiscalDigital',ns.tfd);
  const conceptos = comp.getChild('Conceptos',ns.cfdi)?.getChildren('Concepto',ns.cfdi)||[];
  const c0 = conceptos[0];

  const data = {
    Tipo: g(comp,'TipoDeComprobante'),
    UUID: tim? g(tim,'UUID'): '',
    Serie: g(comp,'Serie'), Folio: g(comp,'Folio'), Fecha: g(comp,'Fecha'),
    RFC_Emisor: g(em,'Rfc'), Nombre_Emisor: g(em,'Nombre'),
    RFC_Receptor: g(re,'Rfc'), Nombre_Receptor: g(re,'Nombre'),
    UsoCFDI: g(re,'UsoCFDI'), Método: g(comp,'MetodoPago'), Forma_Pago: g(comp,'FormaPago'),
    Moneda: g(comp,'Moneda')||'MXN', Tipo_Cambio: g(comp,'TipoCambio')||'1',
    Subtotal: +g(comp,'SubTotal')||0, Descuento: +g(comp,'Descuento')||0, Total: +g(comp,'Total')||0,
    Concepto_Principal: c0? (g(c0,'ClaveProdServ')+': '+(g(c0,'Descripcion')||'')) : ''
  };

  // Impuestos compactos
  const imp = comp.getChild('Impuestos',ns.cfdi);
  if (imp){
    const tras = imp.getChild('Traslados',ns.cfdi)?.getChildren('Traslado',ns.cfdi)||[];
    tras.forEach(t=>{
      const tasa = parseFloat(g(t,'TasaOCuota'))||0;
      const impi = parseFloat(g(t,'Importe'))||0;
      if (Math.abs(tasa-0.16)<1e-6) data.IVA_16 = (data.IVA_16||0)+impi;
      else if (Math.abs(tasa-0.08)<1e-6) data.IVA_08 = (data.IVA_08||0)+impi;
      else if (g(t, 'TipoFactor') === 'Tasa' && tasa === 0) data.IVA_00 = (data.IVA_00||0)+impi;
    });
    const rets = imp.getChild('Retenciones',ns.cfdi)?.getChildren('Retencion',ns.cfdi)||[];
    rets.forEach(r=>{
      const im = g(r,'Impuesto');
      const impi = parseFloat(g(r,'Importe'))||0;
      if (im==='001') data.Ret_ISR = (data.Ret_ISR||0)+impi;
      if (im==='002') data.Ret_IVA = (data.Ret_IVA||0)+impi;
    });
  }

  // Complemento de pagos 2.0 → marcar Tipo=P si aplica
  const pagos = complemento?.getChild('Pagos',ns.p20);
  if (pagos) data.Tipo = 'P';

  // Relacionados (para Notas de Crédito y Pagos)
  const relacionados = comp.getChild('CfdiRelacionados', ns.cfdi);
  if (relacionados) {
    data.Tipo_Relacion = g(relacionados, 'TipoRelacion');
    const relacionado = relacionados.getChild('CfdiRelacionado', ns.cfdi);
    data.UUID_Relacionado = g(relacionado, 'UUID');
  }

  return data;
}

/**
 * Convierte un objeto de datos de CFDI en un array para ser insertado como fila en la hoja.
 * @param {Object} d - El objeto de datos del CFDI.
 * @return {Array} Un array con los valores en el orden correcto de las columnas.
 */
function JC_toRow_(d){
  // El orden debe coincidir EXACTAMENTE con las columnas de la hoja "CFDI"
  return [
    d.Tipo, d.UUID, d.Fecha,
    d.Fecha ? Utilities.formatDate(new Date(d.Fecha), Session.getScriptTimeZone(),'yyyy-MM'):'',
    d.Serie, d.Folio,
    d.RFC_Emisor, d.Nombre_Emisor, d.RFC_Receptor, d.Nombre_Receptor,
    d.UsoCFDI, d.Método, d.Forma_Pago, d.Moneda, d.Tipo_Cambio,
    d.Concepto_Principal,
    d.Subtotal, d.Descuento, d.IVA_16||0, d.IVA_08||0, d.IVA_00||0, d.Ret_ISR||0, d.Ret_IVA||0,
    d.IEPS||0, // IEPS no está en el extractor, se asume 0
    d.Total,
    d.UUID_Relacionado || '', d.Tipo_Relacion || '',
    d.ArchivoXML_ID ? `https://drive.google.com/file/d/${d.ArchivoXML_ID}/view` : '', // Link_XML
    '', // Link_PDF (se deja vacío)
    d.ArchivoXML_ID || '',
    'No', // Con_Póliza
    'No', // Conciliada
    'Sí'  // Incluida_Reportes
  ];
}

/**
 * Función de utilidad para obtener una hoja por su nombre.
 * @param {string} n - El nombre de la hoja.
 * @return {GoogleAppsScript.Spreadsheet.Sheet|null} El objeto de la hoja o null si no se encuentra.
 */
function ss_(n){
  return SpreadsheetApp.getActive().getSheetByName(n);
}