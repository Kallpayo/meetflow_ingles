function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ MeetFlow CRM')
    .addItem('Open Main Panel', 'abrirPanelWeb')
    .addToUi();
}

function abrirPanelWeb() {
  // Esto abre tu interfaz HTML en una ventana modal dentro de la misma hoja
  var html = HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('MeetFlow CRM')
      .setWidth(1200)
      .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'MeetFlow CRM');
}
const SHEETS_CONFIG = {
  USUARIO: {
    sheetName: 'USUARIO', idPrefix: 'USU',
    headerMap: { 
      'UserID': 'usuarioID', 
      'EmployeeID': 'numeroEmpleado',
      'FullName': 'nombreCompleto', 
      'CorporateEmail': 'correoCorporativo', 
      'PhoneExtension': 'telefonoContacto', 
      'Gender': 'genero', 
      'Position': 'cargo', 
      'Department': 'departamento', 
      'EmploymentType': 'tipoContrato',
      'WorkMode': 'modalidadTrabajo',
      'PrivacyConsent': 'consentimientoGDPR',
      'RegistrationDate': 'fechaRegistro', 
      'LastUpdate': 'ultimaActualizacion' 
    }
  },
  CONSULTA: {
    sheetName: 'CONSULTA', idPrefix: 'CON',
    headerMap: { 
      'InteractionID': 'consultaID', 
      'Department': 'departamento',
      'InteractionChannel': 'medioAtencion', 
      'InteractionDate': 'fechaConsulta', 
      'StartTime': 'horaVisita', 
      'EndTime': 'horaFin', 
      'TicketCode': 'codigoTicket',   
      'InteractionType': 'tipoConsulta', 
      'AgendaSubject': 'temaOrdenDelDia', 
      'ResolutionStatus': 'tipoSolucion', 
      'Commitments': 'acuerdosCompromisos', 
      'AssignedOfficer': 'nombreFuncionarioQueAtendio', 
      'PreScheduled': 'agendadaPreviamente' 
    }
  },
  PARTICIPANTES: {
    sheetName: 'PARTICIPANTES',
    headerMap: { 'InteractionID': 'consultaID', 'UserID': 'usuarioID', 'Role': 'rol' }
  },
  CONFIG: {
    sheetName: 'CONFIG',
    headerMap: { 'ListType': 'ListType', 'ItemValue': 'ItemValue' }
  }
};

function logActivity(level, message) {
  Logger.log(`[${level}] ${new Date().toISOString()} - ${message}`);
}

function doGet(e) {
  const page = e.parameter.page || 'index';
  return HtmlService.createTemplateFromFile(page)
    .evaluate()
    .setTitle('MeetFlow - Executive Portal')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function getTemplateContent(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (e) {
    return `<div class="error">Error: Module ${filename}.html not found.</div>`;
  }
}

function parseSheetValue(value) {
  if (value == null || value === undefined) return '';
  if (typeof value === 'string' && value.trim() === '') return '';
  if (value instanceof Date) {
    const isTimeOnly = value.getFullYear() === 1899;
    const format = isTimeOnly ? "HH:mm" : "yyyy-MM-dd";
    return Utilities.formatDate(value, Session.getScriptTimeZone(), format);
  }
  return String(value).trim();
}

function getSheetData(sheetNameFromConfig) {
  Logger.log(`=== getSheetData: ${sheetNameFromConfig} ===`);
  try {
    const config = SHEETS_CONFIG[sheetNameFromConfig];
    if (!config) {
      Logger.log(`ERROR: No existe configuración para ${sheetNameFromConfig}`);
      return JSON.stringify([]);
    }
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.sheetName);
    if (!sheet) {
      Logger.log(`ERROR: Hoja ${config.sheetName} no existe`);
      return JSON.stringify([]);
    }
    
    const values = sheet.getDataRange().getValues();
    if (values.length < 2) {
      Logger.log(`La hoja ${config.sheetName} está vacía`);
      return JSON.stringify([]);
    }

    const headers = values[0].map(h => String(h).trim());
    Logger.log(`Headers de ${config.sheetName}: ${JSON.stringify(headers)}`);
    
    const data = values.slice(1).map((row, rowIndex) => {
      const obj = {};
      Object.entries(config.headerMap).forEach(([sheetHeader, jsKey]) => {
        let index = headers.indexOf(sheetHeader);
        
        if (index === -1) {
          const normalizedSearch = sheetHeader.replace(/\s+/g, '');
          for (let i = 0; i < headers.length; i++) {
            if (headers[i].replace(/\s+/g, '') === normalizedSearch) {
              index = i;
              break;
            }
          }
        }
        
        obj[jsKey] = index !== -1 ? parseSheetValue(row[index]) : '';
      });
      return obj;
    });
    
    Logger.log(`Datos extraídos: ${data.length} registros de ${config.sheetName}`);
    if (data.length > 0) {
      Logger.log(`Primer registro: ${JSON.stringify(data[0])}`);
    }
    
    return JSON.stringify(data);
  } catch (e) {
    Logger.log(`ERROR CRÍTICO en getSheetData: ${e.message}`);
    return JSON.stringify([]);
  }
}

function generateNextId(sheetNameFromConfig, idPrefix) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const config = SHEETS_CONFIG[sheetNameFromConfig];
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.sheetName);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return `${idPrefix}001`;
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    const numericIds = ids
      .filter(id => id && String(id).startsWith(idPrefix))
      .map(id => parseInt(String(id).substring(idPrefix.length), 10))
      .filter(num => !isNaN(num));
    const max = numericIds.length > 0 ? Math.max(...numericIds) : 0;
    return `${idPrefix}${String(max + 1).padStart(3, '0')}`;
  } finally {
    lock.releaseLock();
  }
}

function getDropdownData(tipoLista) {
  try {
    const data = JSON.parse(getSheetData('CONFIG'));
    Logger.log(`Buscando listType: ${tipoLista}`);
    
    if (!data || data.length === 0) {
      Logger.log("No hay datos en CONFIG");
      return JSON.stringify([]);
    }
    
    const unicos = [...new Set(
      data
        .filter(d => d.ListType && d.ListType.trim() === tipoLista.trim())
        .map(d => d.ItemValue)
        .filter(Boolean)
    )];
    
    Logger.log(`Valores únicos para ${tipoLista}: ${JSON.stringify(unicos)}`);
    return JSON.stringify(unicos);
  } catch (e) {
    Logger.log(`Error en getDropdownData para ${tipoLista}: ${e.message}`);
    return JSON.stringify([]);
  }
}

function validateUsuarioData(data) {
  if (!data) throw new Error("Empty data received.");
  if (!data.nombreCompleto || data.nombreCompleto.trim().length < 3) throw new Error("Name is too short.");
  if (!data.correoCorporativo || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(data.correoCorporativo)) throw new Error("Invalid email format.");
  if (!data.genero || !data.cargo || !data.departamento) throw new Error("Gender, Position, and Department are required.");
  return true;
}

function validateConsultaData(data) {
  if (!data) throw new Error("Empty operational data.");
  if (!data.medioAtencion) throw new Error("Interaction channel is required.");
  if (!data.fechaConsulta) throw new Error("Interaction date is required.");
  if (!data.tipoConsulta) throw new Error("Interaction type is required.");
  if (!data.nombreFuncionarioQueAtendio) throw new Error("Assigned officer is required.");

  if (data.medioAtencion === 'Ticket') {
    if (!data.codigoTicket) throw new Error("Ticket code is mandatory for this channel.");
    if (!data.departamento) throw new Error("Department is required for technical tickets.");
    return true;
  }
    // CORREGIDO: Expresión regular que acepta minutos 00-59
  const timeRegex = /^([01]?[0-9]|2[0-3]):[0-5][0-9]$/;
  if (!timeRegex.test(data.horaVisita)) throw new Error("Invalid start time format. Use HH:MM (e.g., 09:30)");
  if (!timeRegex.test(data.horaFin)) throw new Error("Invalid end time format. Use HH:MM (e.g., 17:30)");
  
  // Validar que horaFin sea mayor que horaVisita
  if (data.horaFin <= data.horaVisita) throw new Error("End time must be after the start time.");
  if (!data.temaOrdenDelDia) throw new Error("Subject or agenda details are required.");
  
  if (data.medioAtencion === 'Reunión' && !data.departamento) {
    throw new Error("Department is mandatory for meetings.");
  }
  return true;
}

function createUsuario(usuarioData) {
  try {
    validateUsuarioData(usuarioData);
    const config = SHEETS_CONFIG.USUARIO;
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.sheetName);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const newId = generateNextId('USUARIO', config.idPrefix);
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

    const mappedData = {
      'UserID': newId,
      'EmployeeID': usuarioData.numeroEmpleado || 'N/A',
      'FullName': usuarioData.nombreCompleto.toUpperCase(),
      'CorporateEmail': usuarioData.correoCorporativo.toLowerCase(),
      'PhoneExtension': usuarioData.telefonoContacto || 'N/A',
      'Gender': usuarioData.genero,
      'Department': usuarioData.departamento,
      'Position': usuarioData.cargo,
      'EmploymentType': usuarioData.tipoContrato,
      'WorkMode': usuarioData.modalidadTrabajo,
      'PrivacyConsent': usuarioData.consentimientoGDPR || 'NO',
      'RegistrationDate': now,
      'LastUpdate': now
    };

    const rowData = headers.map(h => {
      const value = mappedData[h];
      if (value === undefined || value === null) {
        Logger.log(`ADVERTENCIA: Header "${h}" no tiene valor mapeado`);
        return '';
      }
      return value;
    });
    
    Logger.log(`Guardando usuario: ${JSON.stringify(rowData)}`);
    sheet.appendRow(rowData);
    return { success: true, message: "Record successfully created with ID: " + newId };
  } catch (e) {
    Logger.log(`ERROR en createUsuario: ${e.message}`);
    return { success: false, message: e.message };
  }
}

function searchUsuarios(searchQuery) {
  try {
    if (!searchQuery || searchQuery.trim().length < 3) return JSON.stringify([]);
    const query = searchQuery.toLowerCase().trim();
    const usersData = getSheetData('USUARIO');
    const users = JSON.parse(usersData);
    
    Logger.log(`Buscando: "${query}" en ${users.length} usuarios`);
    if (users.length > 0) {
      Logger.log(`Primer usuario: ${JSON.stringify(users[0])}`);
    }
    
    const results = users.filter(u => {
      if (!u || typeof u !== 'object') return false;
      const nombre = (u.nombreCompleto || '').toLowerCase();
      const email = (u.correoCorporativo || '').toLowerCase();
      const empId = (u.numeroEmpleado || '').toLowerCase();
      return nombre.includes(query) || email.includes(query) || empId.includes(query);
    });
    
    Logger.log(`Encontrados: ${results.length} resultados`);
    
    return JSON.stringify(results.map(u => ({
      usuarioID: u.usuarioID || '',
      nombreCompleto: u.nombreCompleto || 'Sin nombre',
      correoCorporativo: u.correoCorporativo || '',
      numeroEmpleado: u.numeroEmpleado || 'No ID'
    })));
  } catch (e) {
    Logger.log(`Error en searchUsuarios: ${e.message}`);
    return JSON.stringify([]);
  }
}

function createConsulta(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    validateConsultaData(payload.consultaData);
    if (!payload.participantes || payload.participantes.length === 0) {
      throw new Error("At least one attendee or requester must be linked.");
    }

    const configCons = SHEETS_CONFIG.CONSULTA;
    const sheetCons = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(configCons.sheetName);
    const headersCons = sheetCons.getRange(1, 1, 1, sheetCons.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const newId = generateNextId('CONSULTA', configCons.idPrefix);
    const data = payload.consultaData;

    const mappedConsulta = {
      'InteractionID': newId,
      'Department': data.departamento || '',
      'InteractionChannel': data.medioAtencion || '',
      'InteractionDate': data.fechaConsulta || '',
      'StartTime': data.horaVisita || '',
      'EndTime': data.horaFin || '',
      'TicketCode': data.codigoTicket || '',
      'AgendaSubject': data.temaOrdenDelDia || '',
      'ResolutionStatus': data.tipoSolucion || '',
      'InteractionType': data.tipoConsulta || '',
      'Commitments': data.acuerdosCompromisos || '',
      'AssignedOfficer': data.nombreFuncionarioQueAtendio || '',
      'PreScheduled': data.agendadaPreviamente || ''
    };

    const rowDataCons = headersCons.map(h => mappedConsulta[h] || '');
    Logger.log(`Guardando consulta: ${JSON.stringify(rowDataCons)}`);
    sheetCons.appendRow(rowDataCons);

    const configPart = SHEETS_CONFIG.PARTICIPANTES;
    const sheetPart = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(configPart.sheetName);
    const headersPart = sheetPart.getRange(1, 1, 1, sheetPart.getLastColumn()).getValues()[0].map(h => String(h).trim());

    payload.participantes.forEach(part => {
      const mappedPart = {
        'InteractionID': newId,
        'UserID': part.usuarioID || '',
        'Role': part.rol || ''
      };
      const rowDataPart = headersPart.map(h => mappedPart[h] || '');
      Logger.log(`Guardando participante: ${JSON.stringify(rowDataPart)}`);
      sheetPart.appendRow(rowDataPart);
    });

    return { success: true, message: "Transaction successfully logged." };
  } catch (e) {
    logActivity('ERROR', `createConsulta: ${e.message}`);
    return { success: false, message: "Logging error: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

function getDatosDashboard() {
  try {
    const consultas = JSON.parse(getSheetData('CONSULTA'));
    const usuarios = JSON.parse(getSheetData('USUARIO'));
    const participantes = JSON.parse(getSheetData('PARTICIPANTES'));

    const usuariosMap = usuarios.reduce((map, u) => {
      map[u.usuarioID] = u.nombreCompleto;
      return map;
    }, {});

    const conteoPorTipo = {};
    const conteoPorMedio = {};

    consultas.forEach(c => {
      const tipo = c.tipoConsulta || 'General';
      const medio = c.medioAtencion || 'On-site';
      conteoPorTipo[tipo] = (conteoPorTipo[tipo] || 0) + 1;
      conteoPorMedio[medio] = (conteoPorMedio[medio] || 0) + 1;
    });

    const ultimas = consultas.slice(-5).reverse().map(c => {
      const parts = participantes.filter(p => p.consultaID === c.consultaID);
      const nombresPartes = parts.map(p => `${usuariosMap[p.usuarioID] || 'Corporate Employee'} (${p.rol})`).join(', ');
      return {
        fecha: c.fechaConsulta,
        funcionario: c.nombreFuncionarioQueAtendio,
        tipo: c.tipoConsulta,
        medio: c.medioAtencion,
        participantes: nombresPartes || 'Direct log',
        resolucion: c.tipoSolucion
      };
    });

    const totalConsultas = consultas.length;
    const resueltas = consultas.filter(c => c.tipoSolucion && (c.tipoSolucion.toUpperCase().includes('RESOLVED'))).length;
    const pendientes = consultas.filter(c => c.tipoSolucion &&(c.tipoSolucion.toUpperCase().includes('PENDING'))).length;
    const cerradas = consultas.filter(c => c.tipoSolucion &&(c.tipoSolucion.toUpperCase().includes('CLOSED'))).length;

    return {
      success: true,
      data: {
        kpis: {
          totalConsultas: totalConsultas,
          totalResueltas: resueltas,
          totalPendientes: pendientes,
          totalCerradas: cerradas,
          totalUsuarios: usuarios.length,
          promedioDuracion: calcularPromedioDuracion(consultas)
        },
        conteoPorTipo: conteoPorTipo,
        conteoPorMedio: conteoPorMedio,
        ultimasActividades: ultimas
      }
    };
  } catch (e) {
    logActivity('ERROR', `getDatosDashboard: ${e.message}`);
    return { success: false, message: "Metrics calculation exception: " + e.message };
  }
}

function calcularPromedioDuracion(consultas) {
  const validas = consultas.filter(c => 
    c.horaVisita && c.horaFin && c.horaVisita !== '00:00' &&
    /^([1]?[1-9]|2[1-3]):[1-5][1-9]$/.test(c.horaVisita) &&
    /^([1]?[1-9]|2[1-3]):[1-5][1-9]$/.test(c.horaFin)
  );
  if (!validas.length) return 0;
  const total = validas.reduce((sum, c) => {
    const [h1, m1] = c.horaVisita.split(':').map(Number);
    const [h2, m2] = c.horaFin.split(':').map(Number);
    return sum + ((h2 * 60 + m2) - (h1 * 60 + m1));
  }, 0);
  return Math.round(total / validas.length);
}

function generarReporteConsultasSheet(funcionarioFiltro) {
  try {
    const carpetaDestino = getMeetFlowFolder("MeetFlow_Reports");
    let consultas = JSON.parse(getSheetData('CONSULTA'));
    if (consultas.length === 0) return { success: false, message: "No data available." };

    if (funcionarioFiltro && funcionarioFiltro !== 'Todos') {
      consultas = consultas.filter(c => c.nombreFuncionarioQueAtendio === funcionarioFiltro);
      if (consultas.length === 0) return { success: false, message: `No records found for ${funcionarioFiltro}.` };
    }

    const usuarios = JSON.parse(getSheetData('USUARIO'));
    const participantes = JSON.parse(getSheetData('PARTICIPANTES'));
    const usuariosMap = usuarios.reduce((map, user) => { map[user.usuarioID] = user.nombreCompleto; return map; }, {});

    const sufix = funcionarioFiltro && funcionarioFiltro !== 'Todos' ? `_${funcionarioFiltro.replace(/\s/g, '')}` : '_General';
    const nombreArchivo = `MeetFlow_Audit${sufix}_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd")}`;
    const nuevaHoja = SpreadsheetApp.create(nombreArchivo);
    const hoja = nuevaHoja.getActiveSheet();

    const headers = ["ID", "Department", "Channel", "Date", "Start", "End", "Ticket Code", "Type", "Resolution Status", "Subject", "Commitments", "Officer", "Participants", "Scheduled"];
    const datos = consultas.map(c => {
      const parts = participantes.filter(p => p.consultaID === c.consultaID);
      const partNames = parts.map(p => `${usuariosMap[p.usuarioID] || 'N/A'} (${p.rol})`).join('; ');
      return [c.consultaID, c.departamento, c.medioAtencion, c.fechaConsulta, c.horaVisita, c.horaFin, c.codigoTicket, c.tipoConsulta, c.tipoSolucion, c.temaOrdenDelDia, c.acuerdosCompromisos, c.nombreFuncionarioQueAtendio, partNames, c.agendadaPreviamente];
    });

    const outputData = [headers, ...datos];
    hoja.getRange(1, 1, outputData.length, headers.length).setValues(outputData);
    hoja.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#e2e8f0");
    hoja.autoResizeColumns(1, headers.length);
    hoja.setFrozenRows(1);

    const archivo = DriveApp.getFileById(nuevaHoja.getId());
    archivo.moveTo(carpetaDestino);
    return { success: true, fileUrl: archivo.getUrl(), message: "Excel Audit Report generated successfully." };
  } catch (error) {
    return { success: false, message: `System Error: ${error.message}` };
  }
}

function getReunionesPorDepartamento(departamento) {
  try {
    const consultas = JSON.parse(getSheetData('CONSULTA'));
    const reuniones = consultas.filter(c => 
      (c.medioAtencion === 'Reunión' || c.medioAtencion === 'Ticket' || c.medioAtencion === 'Presencial' || c.medioAtencion === 'Virtual') &&
      (departamento === 'Todas' || c.departamento === departamento)
    );
    return JSON.stringify(reuniones.map(r => ({
      id: r.consultaID,
      fecha: r.fechaConsulta,
      hora: r.horaVisita,
      tema: r.temaOrdenDelDia,
      funcionario: r.nombreFuncionarioQueAtendio,
      canal: r.medioAtencion,
      ticket: r.codigoTicket
    })));
  } catch (e) {
    logActivity('ERROR', `getReunionesPorDepartamento: ${e.message}`);
    return JSON.stringify([]);
  }
}

function generarActaReunionPDF(consultaID) {
  try {
    const consultas = JSON.parse(getSheetData('CONSULTA'));
    const consulta = consultas.find(c => c.consultaID === consultaID);
    if (!consulta) throw new Error("Interaction record not found.");

    const nombreEmpresa = getEmpresaNombre();
    const docName = `Meeting_report_${consulta.consultaID}`;
    const doc = DocumentApp.create(docName);
    const body = doc.getBody();

    body.appendParagraph(nombreEmpresa).setAlignment(DocumentApp.HorizontalAlignment.CENTER).setBold(true).setFontSize(14);
    body.appendParagraph("EXECUTIVE SESSION MINUTES").setAlignment(DocumentApp.HorizontalAlignment.CENTER).setBold(true).setFontSize(12);
    body.appendParagraph(`Minutes No. ${consulta.consultaID}`).setAlignment(DocumentApp.HorizontalAlignment.RIGHT).setBold(true).setFontSize(10);
    body.appendParagraph("____________________________________________________________________");
    body.appendParagraph(`Date: ${consulta.fechaConsulta} | Time: ${consulta.horaVisita} - ${consulta.horaFin}`).setBold(true);
    body.appendParagraph(`Department: ${consulta.departamento} | Assigned Officer: ${consulta.nombreFuncionarioQueAtendio}`).setBold(true);
    
    body.appendParagraph("\nATTENDEES:").setBold(true).setFontSize(11);
    const usuarios = JSON.parse(getSheetData('USUARIO'));
    const participantes = JSON.parse(getSheetData('PARTICIPANTES'));
    const usuariosMap = usuarios.reduce((map, u) => { map[u.usuarioID] = u.nombreCompleto; return map; }, {});
    const parts = participantes.filter(p => p.consultaID === consultaID);
    parts.forEach(p => body.appendParagraph(`• ${usuariosMap[p.usuarioID] || 'N/A'} - ${p.rol}`));

    body.appendParagraph("\nSUBJECT / AGENDA:").setBold(true).setFontSize(11);
    body.appendParagraph(consulta.temaOrdenDelDia);
    body.appendParagraph("\nACTION ITEMS & COMMITMENTS:").setBold(true).setFontSize(11);
    body.appendParagraph(consulta.acuerdosCompromisos);

    let tableData = [];
    let row = [];
    parts.forEach((p, i) => {
      const nombre = usuariosMap[p.usuarioID] || 'Guest';
      const rol = p.rol || 'Attendee';
      row.push(`${nombre}\n${rol}`);
      if ((i + 1) % 3 === 0) { tableData.push(row); row = []; }
    });
    if (row.length > 0) {
      while (row.length < 3) { row.push(""); }
      tableData.push(row);
    }

    body.appendParagraph("\n\n\n\n");
    const tablaFirmas = body.appendTable(tableData);
    tablaFirmas.setBorderWidth(0);
    tableData.forEach((r, rowIndex) => {
      const rowObj = tablaFirmas.getRow(rowIndex);
      for(let colIndex = 0; colIndex < 3; colIndex++) {
        const cell = rowObj.getCell(colIndex);
        if(cell && cell.getText().trim() !== "") {
          cell.appendParagraph("\n____________________\nSignature").setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        }
      }
    });

    doc.saveAndClose();
    const archivoDoc = DriveApp.getFileById(doc.getId());
    const carpetaDestino = getMeetFlowFolder("MeetFlow_Minutes");
    const url = `https://docs.google.com/document/d/${doc.getId()}/export?format=pdf`;
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token } });
    const blob = response.getBlob().setName(`${docName}.pdf`);
    const pdfFile = carpetaDestino.createFile(blob);
    archivoDoc.setTrashed(true);

    return { success: true, fileUrl: pdfFile.getUrl(), message: "Minutes PDF generated successfully." };
  } catch (error) {
    return { success: false, message: "PDF Error: " + error.message };
  }
}

function getEmpresaNombre() {
  try {
    const config = JSON.parse(getSheetData('CONFIG'));
    const empConfig = config.find(c => c.ListType === 'Empresa');
    return empConfig ? empConfig.ItemValue : "YOUR COMPANY NAME";
  } catch (e) { 
    Logger.log(`Error en getEmpresaNombre: ${e.message}`);
    return "YOUR COMPANY NAME"; 
  }
}

function getMeetFlowFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(folderName);
}

function getFileForDownload(fileUrl) {
  try {
    const match = fileUrl.match(/[-\w]{25,}/);
    if (!match) throw new Error("Invalid URL.");
    const fileId = match;
    const file = DriveApp.getFileById(fileId);
    const mimeType = file.getMimeType();
    let blob;
    let fileName = file.getName();
    let finalMimeType = mimeType;

    if (mimeType === MimeType.GOOGLE_SHEETS) {
      const url = `https://docs.google.com/spreadsheets/d/${fileId}/export?format=xlsx`;
      const token = ScriptApp.getOAuthToken();
      const response = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token } });
      blob = response.getBlob();
      fileName += '.xlsx';
      finalMimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    } else {
      blob = file.getBlob();
    }

    const bytes = blob.getBytes();
    const base64 = Utilities.base64Encode(bytes);
    return { success: true, base64: base64, mimeType: finalMimeType, fileName: fileName };
  } catch (error) {
    return { success: false, error: "Download preparation error: " + error.message };
  }
}

/** ESPACIO CREADO PARA TEST DE EJECUCIONES. BORRAR LUEGO DE PROBAR VERSION FINAL DE CONFIGURACION MODULO IMPLEMENTADO */

function testConfig() {
  try {
    Logger.log("=== TEST CONFIGURACIÓN ===");
    const data = JSON.parse(getSheetData('CONFIG'));
    Logger.log("Datos CONFIG crudos:");
    data.forEach(row => {
      Logger.log(`ListType: ${row.ListType}, ItemValue: ${row.ItemValue}`);
    });
    
    const tipos = ['Departamentos', 'Cargos', 'TipoConsulta', 'Generos', 'Agendado', 'NomEmp'];
    tipos.forEach(tipo => {
      const resultado = getDropdownData(tipo);
      Logger.log(`Dropdown para ${tipo}: ${resultado}`);
    });
    
    return "Test completado. Revisa los logs.";
  } catch (e) {
    Logger.log(`Error en testConfig: ${e.message}`);
    return `Error: ${e.message}`;
  }
}

function normalizarConfig() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CONFIG');
    const data = sheet.getDataRange().getValues();
    
    const headers = data[0];
    Logger.log(`Encabezados CONFIG: ${JSON.stringify(headers)}`);
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      Logger.log(`Fila ${i}: ListType="${row[0]}", ItemValue="${row[1]}"`);
    }
    
    return "Configuración normalizada";
  } catch (e) {
    return `Error: ${e.message}`;
  }
}

function diagnosticarDropdowns() {
  const tipos = ['Departamentos', 'Cargos', 'TipoConsulta', 'Generos', 'Agendado', 'NomEmp'];
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CONFIG');
  const data = sheet.getDataRange().getValues();
  
  Logger.log("=== DIAGNÓSTICO ===");
  Logger.log(`Headers: ${JSON.stringify(data[0])}`);
  
  tipos.forEach(tipo => {
    const items = data.filter(row => row[0] === tipo).map(row => row[1]);
    Logger.log(`${tipo}: ${items.length} items - ${JSON.stringify(items)}`);
  });
  
  tipos.forEach(tipo => {
    const resultado = getDropdownData(tipo);
    Logger.log(`getDropdownData("${tipo}"): ${resultado}`);
  });
}

function autorizarPermisos() {
  SpreadsheetApp.getActiveSpreadsheet();
  DriveApp.getFoldersByName("Test").hasNext();
  UrlFetchApp.fetch("https://www.google.com");
  const doc = DocumentApp.create("Temp");
  DriveApp.getFileById(doc.getId()).setTrashed(true);
}

function diagnosticarDatos() {
  Logger.log("=== DIAGNÓSTICO DE DATOS ===");
  
  const usuarios = JSON.parse(getSheetData('USUARIO'));
  Logger.log(`Usuarios: ${usuarios.length}`);
  if (usuarios.length > 0) {
    Logger.log(`Primer usuario: ${JSON.stringify(usuarios[0])}`);
  }
  
  const consultas = JSON.parse(getSheetData('CONSULTA'));
  Logger.log(`Consultas: ${consultas.length}`);
  if (consultas.length > 0) {
    Logger.log(`Primera consulta: ${JSON.stringify(consultas[0])}`);
  }
  
  const participantes = JSON.parse(getSheetData('PARTICIPANTES'));
  Logger.log(`Participantes: ${participantes.length}`);
  if (participantes.length > 0) {
    Logger.log(`Primer participante: ${JSON.stringify(participantes[0])}`);
  }
  
  const busqueda = searchUsuarios('MILO');
  Logger.log(`Búsqueda 'MILO': ${busqueda}`);
  
  return "Diagnóstico completado. Revisa los logs.";
}

/** FIN ESPACIO DE PRUEBAS BORRAR SI TODO VA BIEN */

// --- INICIO MÓDULO DE CONFIGURACIÓN DINÁMICA ---

function getAllConfigData() {
  try {
    return getSheetData('CONFIG');
  } catch (e) {
    logActivity('ERROR', 'getAllConfigData: ' + e.message);
    return JSON.stringify([]);
  }
}

function updateConfigParameter(tipo, valorNuevo) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const config = SHEETS_CONFIG.CONFIG;
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.sheetName);
    const data = sheet.getDataRange().getValues();
    
    // Recorremos la hoja buscando el parámetro (ej: 'Empresa')
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === tipo) {
        sheet.getRange(i + 1, 2).setValue(valorNuevo);
        return { success: true, message: "Parameter'" + tipo + "'updated successfully" };
      }
    }
    
    // Si no lo encuentra, lo crea al final
    sheet.appendRow([tipo, valorNuevo]);
    return { success: true, message: "Parameter'" + tipo + "'created successfully" };
  } catch (e) {
    logActivity('ERROR', 'updateConfigParameter: ' + e.message);
    return { success: false, message: "Failed to update configuration: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

function addConfigItem(tipo, valor) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (!tipo || !valor) throw new Error("List type and value are required.");
    const config = SHEETS_CONFIG.CONFIG;
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.sheetName);
    
    sheet.appendRow([tipo, valor]);
    return { success: true, message: "Item added '" + valor + "' to the list " + tipo + "." };
  } catch (e) {
    logActivity('ERROR', 'addConfigItem: ' + e.message);
    return { success: false, message: "Error adding item: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteConfigItem(tipo, valor) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const config = SHEETS_CONFIG.CONFIG;
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.sheetName);
    const data = sheet.getDataRange().getValues();
    
    let deleted = false;
    // Buscamos desde abajo hacia arriba para que los índices no cambien al eliminar filas
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]).trim() === String(tipo).trim() && String(data[i][1]).trim() === String(valor).trim()) {
        sheet.deleteRow(i + 1);
        deleted = true;
        break; // Solo elimina la primera coincidencia exacta
      }
    }
    
    if (deleted) {
      return { success: true, message: "the item '" + valor + "' Successfully deleted." };
    } else {
      return { success: false, message: "Item not found in database." };
    }
  } catch (e) {
    logActivity('ERROR', 'deleteConfigItem: ' + e.message);
    return { success: false, message: "Error deleting item: " + e.message };
  } finally {
    lock.releaseLock();
  }
}