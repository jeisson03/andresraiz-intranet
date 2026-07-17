const CLAVE_SECRETA = "andresraiz2026";

function doGet(e) {
  var callback = e.parameter.callback;

  function jsonResp(data) {
    var json = JSON.stringify(data);
    if (callback) {
      return ContentService.createTextOutput(callback + '(' + json + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
  }

  if (!e || !e.parameter || !e.parameter.clave || e.parameter.clave !== CLAVE_SECRETA) {
    return jsonResp({ error: "Acceso denegado" });
  }

  var action = e.parameter.action;
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ==================== COLILLAS ====================
  if (!action) {
    var usuario = e.parameter.usuario;
    var periodo = e.parameter.periodo;

    var sheet = ss.getSheetByName("colillas");
    if (!sheet) return jsonResp({ error: "Hoja no encontrada" });

    var data = sheet.getDataRange().getValues();
    var headers = data[0];

    for (let i = 1; i < data.length; i++) {
      const fila = data[i];
      const row = {};
      for (let j = 0; j < headers.length; j++) {
        row[headers[j]] = fila[j];
      }

      if (row.usuario === usuario && row.periodo === periodo) {
        return jsonResp({
          usuario: row.usuario,
          periodo: row.periodo,
          salario: row.salario,
          transporte: row.transporte,
          rodamiento: row.rodamiento,
          comisiones: row.comisiones
        });
      }
    }

    return jsonResp({ error: "Colilla no encontrada" });
  }

  // ==================== SOLICITUDES ====================
  if (action === 'solicitudes') {
    var cedula = e.parameter.cedula;
    var sheet = ss.getSheetByName('solicitudes');
    if (!sheet) return jsonResp({data: []});
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var cedulaIdx = -1;
    for (var h = 0; h < headers.length; h++) {
      if (headers[h].toString().toLowerCase().replace(/[o]/g,'o') === 'cedula' || headers[h].toString().normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase() === 'cedula') {
        cedulaIdx = h;
        break;
      }
    }
    var results = [];
    for (var i = 1; i < data.length; i++) {
      var match = true;
      if (cedula && cedulaIdx >= 0) {
        var cellVal = data[i][cedulaIdx].toString().replace(/[^0-9]/g, '');
        if (cellVal != cedula.toString().replace(/[^0-9]/g, '')) match = false;
      }
      if (match) {
        var obj = {};
        for (var j = 0; j < headers.length; j++) {
          obj[headers[j]] = data[i][j] instanceof Date ? Utilities.formatDate(data[i][j], 'America/Bogota', 'dd/MM/yyyy') : data[i][j];
        }
        results.push(obj);
      }
    }
    return jsonResp({data: results});
  }

  // ==================== DISPONIBLES ====================
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
        return jsonResp({disponibles: disponibles, tomadas: tomadas, acumuladas: acumuladas});
      }
    }
    return jsonResp({disponibles: 0});
  }

  // ==================== APROBAR / RECHAZAR ====================
  if (action === 'aprobar' || action === 'rechazar') {
    var sheet = ss.getSheetByName('solicitudes');
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var id = e.parameter.id;

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == id) {
        var estado = action === 'aprobar' ? 'Aprobada' : 'Rechazada';
        sheet.getRange(i + 1, headers.indexOf('Estado') + 1).setValue(estado);
        sheet.getRange(i + 1, headers.indexOf('Aprobado Por') + 1).setValue(e.parameter.aprobadoPor || '');
        sheet.getRange(i + 1, headers.indexOf('Observaciones') + 1).setValue(e.parameter.observaciones || '');

        if (action === 'aprobar') {
          var empSheet = ss.getSheetByName('empleados');
          var empData = empSheet.getDataRange().getValues();
          var empHeaders = empData[0];
          var cedulaColIdx = -1;
          for (var h = 0; h < empHeaders.length; h++) {
            if (empHeaders[h].toString().normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase() === 'cedula') {
              cedulaColIdx = h;
              break;
            }
          }
          var solCedulaIdx = -1;
          for (var h2 = 0; h2 < headers.length; h2++) {
            if (headers[h2].toString().normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase() === 'cedula') {
              solCedulaIdx = h2;
              break;
            }
          }
          if (cedulaColIdx >= 0 && solCedulaIdx >= 0) {
            for (var j = 1; j < empData.length; j++) {
              var empCed = empData[j][cedulaColIdx].toString().replace(/[^0-9]/g, '');
              var solCed = data[i][solCedulaIdx].toString().replace(/[^0-9]/g, '');
              if (empCed == solCed) {
                var vacIdx = empHeaders.indexOf('vacaciones');
                if (vacIdx < 0) {
                  for (var vh = 0; vh < empHeaders.length; vh++) {
                    if (empHeaders[vh].toString().normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase() === 'vacaciones') {
                      vacIdx = vh;
                      break;
                    }
                  }
                }
                var actuales = parseInt(empData[j][vacIdx]) || 0;
                var diasVal = 0;
                var diasKeys = ['Días Hábiles', 'Dias Habiles', 'Días', 'Dias'];
                for (var dk = 0; dk < diasKeys.length; dk++) {
                  var idx = headers.indexOf(diasKeys[dk]);
                  if (idx >= 0 && data[i][idx]) { diasVal = parseInt(data[i][idx]) || 0; break; }
                }
                empSheet.getRange(j + 1, vacIdx + 1).setValue(actuales + diasVal);
                break;
              }
            }
          }
        }

        return jsonResp({success: true});
      }
    }
    return jsonResp({error: 'No encontrada'});
  }

  // ==================== NUEVA SOLICITUD (GET) ====================
  if (action === 'nueva_solicitud') {
    var sheet = ss.getSheetByName('solicitudes');
    if (!sheet) {
      sheet = ss.insertSheet('solicitudes');
      sheet.appendRow(['ID', 'Cédula', 'Nombre', 'Fecha Solicitud', 'Fecha Inicio', 'Fecha Fin', 'Días Calendario', 'Días Hábiles', 'Motivo', 'Estado', 'Aprobado Por', 'Observaciones']);
    }

    var id = 'SOL-' + new Date().getTime();
    var fechaSolicitud = Utilities.formatDate(new Date(), 'America/Bogota', 'dd/MM/yyyy');

    sheet.appendRow([
      id,
      e.parameter.cedula || '',
      e.parameter.nombre || '',
      fechaSolicitud,
      e.parameter.fechaInicio || '',
      e.parameter.fechaFin || '',
      e.parameter.diasCalendario || '',
      e.parameter.diasHabiles || '',
      e.parameter.motivo || '',
      'Pendiente',
      '',
      ''
    ]);

    return jsonResp({success: true, id: id});
  }

  return jsonResp({error: 'Accion no valida'});
}

function doPost(e) {
  var params = {};

  if (e.parameter && Object.keys(e.parameter).length > 0) {
    params = e.parameter;
  } else if (e.postData && e.postData.contents) {
    var lines = e.postData.contents.split('\n');
    for (var i = 0; i < lines.length; i++) {
      var idx = lines[i].indexOf('=');
      if (idx > -1) {
        params[lines[i].substring(0, idx)] = lines[i].substring(idx + 1);
      }
    }
  }

  if (!params.clave || params.clave !== CLAVE_SECRETA) {
    return ContentService.createTextOutput(JSON.stringify({error: 'Acceso denegado'})).setMimeType(ContentService.MimeType.JSON);
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var action = params.action;

  if (action === 'nueva_solicitud') {
    var sheet = ss.getSheetByName('solicitudes');
    if (!sheet) {
      sheet = ss.insertSheet('solicitudes');
      sheet.appendRow(['ID', 'Cédula', 'Nombre', 'Fecha Solicitud', 'Fecha Inicio', 'Fecha Fin', 'Días Calendario', 'Días Hábiles', 'Motivo', 'Estado', 'Aprobado Por', 'Observaciones']);
    }

    var id = 'SOL-' + new Date().getTime();
    var fechaSolicitud = Utilities.formatDate(new Date(), 'America/Bogota', 'dd/MM/yyyy');

    sheet.appendRow([
      id,
      params.cedula || '',
      params.nombre || '',
      fechaSolicitud,
      params.fechaInicio || '',
      params.fechaFin || '',
      params.diasCalendario || '',
      params.diasHabiles || '',
      params.motivo || '',
      'Pendiente',
      '',
      ''
    ]);

    return ContentService.createTextOutput(JSON.stringify({success: true, id: id})).setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'aprobar' || action === 'rechazar') {
    var sheet = ss.getSheetByName('solicitudes');
    var data = sheet.getDataRange().getValues();
    var headers = data[0];

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == params.id) {
        var estado = action === 'aprobar' ? 'Aprobada' : 'Rechazada';
        sheet.getRange(i + 1, headers.indexOf('Estado') + 1).setValue(estado);
        sheet.getRange(i + 1, headers.indexOf('Aprobado Por') + 1).setValue(params.aprobadoPor || '');
        sheet.getRange(i + 1, headers.indexOf('Observaciones') + 1).setValue(params.observaciones || '');

        if (action === 'aprobar') {
          var empSheet = ss.getSheetByName('empleados');
          var empData = empSheet.getDataRange().getValues();
          var empHeaders = empData[0];
          for (var j = 1; j < empData.length; j++) {
            if (empData[j][empHeaders.indexOf('cedula')] == data[i][headers.indexOf('Cedula')]) {
              var vacIdx = empHeaders.indexOf('vacaciones');
              var actuales = parseInt(empData[j][vacIdx]) || 0;
              var diasVal2 = 0;
              var diasKeys2 = ['Días Hábiles', 'Dias Habiles', 'Días', 'Dias'];
              for (var dk2 = 0; dk2 < diasKeys2.length; dk2++) {
                var idx2 = headers.indexOf(diasKeys2[dk2]);
                if (idx2 >= 0 && data[i][idx2]) { diasVal2 = parseInt(data[i][idx2]) || 0; break; }
              }
              empSheet.getRange(j + 1, vacIdx + 1).setValue(actuales + diasVal2);
              break;
            }
          }
        }

        return ContentService.createTextOutput(JSON.stringify({success: true})).setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({error: 'No encontrada'})).setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(JSON.stringify({error: 'Accion no valida'})).setMimeType(ContentService.MimeType.JSON);
}
