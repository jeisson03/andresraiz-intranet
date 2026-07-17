const CLAVE_SECRETA = "andresraiz2026";

function doGet(e) {
  if (!e || !e.parameter || !e.parameter.clave || e.parameter.clave !== CLAVE_SECRETA) {
    return ContentService.createTextOutput(
      JSON.stringify({ error: "Acceso denegado" })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  var action = e.parameter.action;
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ==================== COLILLAS ====================
  if (!action) {
    var usuario = e.parameter.usuario;
    var periodo = e.parameter.periodo;

    var sheet = ss.getSheetByName("colillas");
    if (!sheet) {
      return ContentService.createTextOutput(
        JSON.stringify({ error: "Hoja no encontrada" })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    var data = sheet.getDataRange().getValues();
    var headers = data[0];

    for (let i = 1; i < data.length; i++) {
      const fila = data[i];
      const row = {};
      for (let j = 0; j < headers.length; j++) {
        row[headers[j]] = fila[j];
      }

      if (row.usuario === usuario && row.periodo === periodo) {
        return ContentService.createTextOutput(
          JSON.stringify({
            usuario: row.usuario,
            periodo: row.periodo,
            salario: row.salario,
            transporte: row.transporte,
            rodamiento: row.rodamiento,
            comisiones: row.comisiones
          })
        ).setMimeType(ContentService.MimeType.JSON);
      }
    }

    return ContentService.createTextOutput(
      JSON.stringify({ error: "Colilla no encontrada" })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  // ==================== SOLICITUDES VACACIONES ====================
  if (action === 'solicitudes') {
    var cedula = e.parameter.cedula;
    var sheet = ss.getSheetByName('solicitudes');
    if (!sheet) return ContentService.createTextOutput(JSON.stringify({data: []})).setMimeType(ContentService.MimeType.JSON);
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

  if (!body.clave || body.clave !== CLAVE_SECRETA) {
    return ContentService.createTextOutput(
      JSON.stringify({ error: "Acceso denegado" })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ==================== APROBAR / RECHAZAR ====================
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

  // ==================== NUEVA SOLICITUD ====================
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
