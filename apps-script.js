var SCRIPT_URL = 'PEGA_AQUI_LA_URL_DE_TU_APP_SCRIPT';

function doGet(e) {
  var action = e.parameter.action;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (action === 'solicitudes') {
    var cedula = e.parameter.cedula;
    var soloPendientes = e.parameter.pendientes === 'true';
    var sheet = ss.getSheetByName('solicitudes');
    if (!sheet) return ContentService.createTextOutput(JSON.stringify({data: []})).setMimeType(ContentService.MimeType.JSON);
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var results = [];
    for (var i = 1; i < data.length; i++) {
      var match = true;
      if (cedula && data[i][headers.indexOf('Cedula')] != cedula) match = false;
      if (soloPendientes && data[i][headers.indexOf('Estado')] != 'Pendiente') match = false;
      if (match) {
        var obj = {};
        for (var j = 0; j < headers.length; j++) {
          obj[headers[j]] = data[i][j] instanceof Date ? Utilities.formatDate(data[i][j], 'America/Bogota', 'dd/MM/yyyy') : data[i][j];
        }
        results.push(obj);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({data: results})).setMimeType(ContentService.MimeType.JSON);
  }
  
  if (action === 'disponibles') {
    var cedula = e.parameter.cedula;
    var sheet = ss.getSheetByName('empleados');
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    for (var i = 1; i < data.length; i++) {
      if (data[i][headers.indexOf('cedula')] == cedula) {
        var tomadas = parseInt(data[i][headers.indexOf('vacaciones')]) || 0;
        var ingreso = new Date(data[i][headers.indexOf('ingreso')]);
        var hoy = new Date();
        var meses = (hoy.getFullYear() - ingreso.getFullYear()) * 12 + (hoy.getMonth() - ingreso.getMonth());
        var acumuladas = Math.floor((15 / 12) * meses);
        var disponibles = Math.max(acumuladas - tomadas, 0);
        return ContentService.createTextOutput(JSON.stringify({disponibles: disponibles, tomadas: tomadas, acumuladas: acumuladas})).setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({disponibles: 0})).setMimeType(ContentService.MimeType.JSON);
  }
  
  return ContentService.createTextOutput(JSON.stringify({error: 'Accion no valida'})).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var body = JSON.parse(e.postData.contents);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (body.action === 'aprobar' || body.action === 'rechazar') {
    var sheet = ss.getSheetByName('solicitudes');
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == body.id) {
        var estado = body.action === 'aprobar' ? 'Aprobada' : 'Rechazada';
        sheet.getRange(i + 1, headers.indexOf('Estado') + 1).setValue(estado);
        sheet.getRange(i + 1, headers.indexOf('Aprobado Por') + 1).setValue(body.aprobadoPor || '');
        sheet.getRange(i + 1, headers.indexOf('Observaciones') + 1).setValue(body.observaciones || '');
        
        if (body.action === 'aprobar') {
          var empSheet = ss.getSheetByName('empleados');
          var empData = empSheet.getDataRange().getValues();
          var empHeaders = empData[0];
          for (var j = 1; j < empData.length; j++) {
            if (empData[j][empHeaders.indexOf('cedula')] == data[i][headers.indexOf('Cedula')]) {
              var vacIdx = empHeaders.indexOf('vacaciones');
              var actuales = parseInt(empData[j][vacIdx]) || 0;
              empSheet.getRange(j + 1, vacIdx + 1).setValue(actuales + parseInt(data[i][headers.indexOf('Dias')]));
              break;
            }
          }
        }
        
        return ContentService.createTextOutput(JSON.stringify({success: true})).setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({error: 'No encontrada'})).setMimeType(ContentService.MimeType.JSON);
  }
  
  var sheet = ss.getSheetByName('solicitudes');
  if (!sheet) {
    sheet = ss.insertSheet('solicitudes');
    sheet.appendRow(['ID', 'Cedula', 'Nombre', 'Fecha Solicitud', 'Fecha Inicio', 'Fecha Fin', 'Dias', 'Motivo', 'Estado', 'Aprobado Por', 'Observaciones']);
  }
  
  var id = 'SOL-' + new Date().getTime();
  var fechaSolicitud = Utilities.formatDate(new Date(), 'America/Bogota', 'dd/MM/yyyy');
  
  sheet.appendRow([id, body.cedula, body.nombre, fechaSolicitud, body.fechaInicio, body.fechaFin, body.dias, body.motivo, 'Pendiente', '', '']);
  
  return ContentService.createTextOutput(JSON.stringify({success: true, id: id})).setMimeType(ContentService.MimeType.JSON);
}
