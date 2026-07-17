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
    var results = [];
    for (var i = 1; i < data.length; i++) {
      var match = true;
      if (cedula && data[i][headers.indexOf('Cedula')] != cedula) match = false;
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

  // ==================== NUEVA SOLICITUD ====================
  if (action === 'nueva_solicitud') {
    var sheet = ss.getSheetByName('solicitudes');
    if (!sheet) {
      sheet = ss.insertSheet('solicitudes');
      sheet.appendRow(['ID', 'Cedula', 'Nombre', 'Fecha Solicitud', 'Fecha Inicio', 'Fecha Fin', 'Dias', 'Motivo', 'Estado', 'Aprobado Por', 'Observaciones']);
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
      e.parameter.dias || '',
      e.parameter.motivo || '',
      'Pendiente',
      '',
      ''
    ]);

    return jsonResp({success: true, id: id});
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
          for (var j = 1; j < empData.length; j++) {
            if (empData[j][empHeaders.indexOf('cedula')] == data[i][headers.indexOf('Cedula')]) {
              var vacIdx = empHeaders.indexOf('vacaciones');
              var actuales = parseInt(empData[j][vacIdx]) || 0;
              empSheet.getRange(j + 1, vacIdx + 1).setValue(actuales + parseInt(data[i][headers.indexOf('Dias')]));
              break;
            }
          }
        }

        return jsonResp({success: true});
      }
    }
    return jsonResp({error: 'No encontrada'});
  }

  return jsonResp({error: 'Accion no valida'});
}
