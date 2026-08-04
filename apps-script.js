const CLAVE_SECRETA = "andresraiz2026";
const HOJA_ID = "19AYAQ6KkliEv9drdqEx77AXReXRRwZFwszLpl8NbFgE";

function norm(t) {
  return (t || '').toString().normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
}

function findSheetByName(ss, name) {
  var sheets = ss.getSheets();
  var target = norm(name);
  for (var s = 0; s < sheets.length; s++) {
    if (norm(sheets[s].getName()) === target) return sheets[s];
  }
  return null;
}

function sheetHasHeaders(sheet, required) {
  if (!sheet) return false;
  var headers = sheet.getDataRange().getValues()[0] || [];
  var normHeaders = [];
  for (var i = 0; i < headers.length; i++) normHeaders.push(norm(headers[i]));
  for (var r = 0; r < required.length; r++) {
    if (normHeaders.indexOf(norm(required[r])) < 0) return false;
  }
  return true;
}

function findSheetByHeaders(ss, required) {
  var sheets = ss.getSheets();
  for (var s = 0; s < sheets.length; s++) {
    if (sheetHasHeaders(sheets[s], required)) return sheets[s];
  }
  return null;
}

function doGet(e) {
  var callback = e.parameter.callback;

  function jsonResp(data) {
    var json = JSON.stringify(data);
    if (callback) {
      return ContentService.createTextOutput(callback + '(' + json + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
  }

  if (!e || !e.parameter) {
    return jsonResp({ error: "Acceso denegado" });
  }
  var secretoRecibido = e.parameter.clave || e.parameter.secret;
  if (!secretoRecibido || secretoRecibido !== CLAVE_SECRETA) {
    return jsonResp({ error: "Acceso denegado" });
  }

  var action = e.parameter.action;
  var ss = SpreadsheetApp.openById(HOJA_ID);

  // ==================== LOGIN ====================
  if (action === 'login') {
    var usuario = (e.parameter.usuario || '').toString().trim();
    var clave = (e.parameter.contrasena || e.parameter.password || '').toString().trim();
    var empSheet = findSheetByName(ss, 'empleados');
          if (!empSheet) empSheet = findSheetByHeaders(ss, ['cedula', 'vacaciones']);
    if (!empSheet) empSheet = findSheetByHeaders(ss, ['usuario', 'clave']);
    if (!empSheet) return jsonResp({ error: 'Hoja de empleados no encontrada' });

    var empData = empSheet.getDataRange().getValues();
    var empHeaders = empData[0];

    function colIdx(name) {
      var target = norm(name);
      for (var c = 0; c < empHeaders.length; c++) {
        if (norm(empHeaders[c]) === target) return c;
      }
      return -1;
    }

    var uIdx = colIdx('usuario');
    var cIdx = colIdx('clave');
    if (uIdx < 0 || cIdx < 0) return jsonResp({ error: 'Credenciales incorrectas' });

    for (var i = 1; i < empData.length; i++) {
      var empUser = (empData[i][uIdx] || '').toString().trim();
      var empClave = (empData[i][cIdx] || '').toString().trim();
      if (empUser === usuario && empClave === clave) {
        var nIdx = colIdx('nombre');
        var cedIdx = colIdx('cedula');
        var carIdx = colIdx('cargo');
        var ingIdx = colIdx('ingreso');
        var salIdx = colIdx('salario');
        var traIdx = colIdx('transporte');
        var rodIdx = colIdx('rodamiento');
        var comIdx = colIdx('comisiones');
        var vacIdx = colIdx('vacaciones');
        var admIdx = colIdx('admin');

        var ingreso = empData[i][ingIdx];
        if (ingreso instanceof Date) {
          ingreso = Utilities.formatDate(ingreso, 'America/Bogota', 'yyyy-MM-dd');
        }

        var admin = false;
        if (admIdx >= 0) {
          var admVal = (empData[i][admIdx] || '').toString().toUpperCase();
          admin = admVal === 'TRUE' || admVal === 'VERDADERO' || admVal === 'SI' || admVal === '1' || admVal === 'X';
        }

        return jsonResp({
          success: true,
          usuario: empUser,
          nombre: nIdx >= 0 ? empData[i][nIdx] : '',
          cedula: cedIdx >= 0 ? empData[i][cedIdx] : '',
          cargo: carIdx >= 0 ? empData[i][carIdx] : '',
          ingreso: ingreso,
          salario: salIdx >= 0 ? empData[i][salIdx] : 0,
          transporte: traIdx >= 0 ? empData[i][traIdx] : 0,
          rodamiento: rodIdx >= 0 ? empData[i][rodIdx] : 0,
          comisiones: comIdx >= 0 ? empData[i][comIdx] : 0,
          vacaciones: vacIdx >= 0 ? empData[i][vacIdx] : 0,
          admin: admin
        });
      }
    }
    return jsonResp({ error: 'Credenciales incorrectas' });
  }

  // ==================== COLILLAS ====================
  if (!action) {
    var usuario = e.parameter.usuario;
    var periodo = e.parameter.periodo;

    var sheet = findSheetByName(ss, 'colillas');
    if (!sheet) sheet = findSheetByHeaders(ss, ['usuario', 'periodo']);
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
    var sheet = findSheetByName(ss, 'solicitudes');
    if (!sheet) sheet = findSheetByHeaders(ss, ['ID', 'Estado', 'Motivo']);
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
    var sheet = findSheetByName(ss, 'empleados');
    if (!sheet) sheet = findSheetByHeaders(ss, ['cedula', 'vacaciones']);
    if (!sheet) return jsonResp({disponibles: 0});
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var cedIdx = -1, vacIdx = -1, ingIdx = -1;
    for (var hc = 0; hc < headers.length; hc++) {
      var hn = norm(headers[hc]);
      if (hn === 'cedula' && cedIdx < 0) cedIdx = hc;
      if (hn === 'vacaciones' && vacIdx < 0) vacIdx = hc;
      if (hn === 'ingreso' && ingIdx < 0) ingIdx = hc;
    }
    if (cedIdx < 0) return jsonResp({disponibles: 0});
    for (var i = 1; i < data.length; i++) {
      var empCed = (data[i][cedIdx] || '').toString().replace(/[^0-9]/g, '');
      if (empCed === (cedula || '').toString().replace(/[^0-9]/g, '')) {
        var tomadas = parseInt(data[i][vacIdx]) || 0;
        var ingreso = new Date(data[i][ingIdx]);
        var hoy = new Date();
        var meses = (hoy.getFullYear() - ingreso.getFullYear()) * 12 + (hoy.getMonth() - ingreso.getMonth());
        var acumuladas = Math.floor((15 / 12) * meses);
        var colectivas = 0;
        var vcSheet = findSheetByName(ss, 'vacaciones_colectivas');
        if (!vcSheet) vcSheet = findSheetByHeaders(ss, ['Fecha Inicio', 'Fecha Fin']);
        if (vcSheet) {
          var vcData = vcSheet.getDataRange().getValues();
          var vcHeaders = vcData[0];
          var fiIdx = -1, ffIdx = -1, diasIdx = -1;
          for (var hv = 0; hv < vcHeaders.length; hv++) {
            var hvn = norm(vcHeaders[hv]);
            if (hvn === 'fecha inicio' && fiIdx < 0) fiIdx = hv;
            if (hvn === 'fecha fin' && ffIdx < 0) ffIdx = hv;
            if (hvn === 'dias' && diasIdx < 0) diasIdx = hv;
          }
          for (var v = 1; v < vcData.length; v++) {
            var vcInicio = new Date(vcData[v][fiIdx]);
            var vcFin = new Date(vcData[v][ffIdx]);
            if (hoy >= vcInicio && hoy <= vcFin) {
              colectivas += parseInt(vcData[v][diasIdx]) || 0;
            }
          }
        }
        var disponibles = Math.max(acumuladas - tomadas - colectivas, 0);
        return jsonResp({disponibles: disponibles, tomadas: tomadas, acumuladas: acumuladas, colectivas: colectivas});
      }
    }
    return jsonResp({disponibles: 0});
  }

  // ==================== VACACIONES COLECTIVAS ====================
  if (action === 'vacaciones_colectivas') {
    var vcSheet = findSheetByName(ss, 'vacaciones_colectivas');
    if (!vcSheet) vcSheet = findSheetByHeaders(ss, ['Fecha Inicio', 'Fecha Fin']);
    if (!vcSheet) return jsonResp({data: []});
    var vcData = vcSheet.getDataRange().getValues();
    var vcHeaders = vcData[0];
    var results = [];
    for (var v = 1; v < vcData.length; v++) {
      var obj = {};
      for (var j = 0; j < vcHeaders.length; j++) {
        obj[vcHeaders[j]] = vcData[v][j] instanceof Date ? Utilities.formatDate(vcData[v][j], 'America/Bogota', 'dd/MM/yyyy') : vcData[v][j];
      }
      results.push(obj);
    }
    return jsonResp({data: results});
  }

  // ==================== APROBAR / RECHAZAR ====================
  if (action === 'aprobar' || action === 'rechazar') {
    var sheet = findSheetByName(ss, 'solicitudes');
    if (!sheet) sheet = findSheetByHeaders(ss, ['ID', 'Estado', 'Motivo']);
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
          var empSheet = findSheetByName(ss, 'empleados');
          if (!empSheet) empSheet = findSheetByHeaders(ss, ['cedula', 'vacaciones']);
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
                empSheet.getRange(j + 1, vacIdx + 1).setValue(1);
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
    var sheet = findSheetByName(ss, 'solicitudes');
    if (!sheet) sheet = findSheetByHeaders(ss, ['ID', 'Estado', 'Motivo']);
    if (!sheet) {
      sheet = ss.insertSheet('solicitudes');
      sheet.appendRow(['ID', 'Cédula', 'Nombre', 'Fecha Solicitud', 'Fecha Inicio', 'Fecha Fin', 'Días', 'Motivo', 'Estado', 'Aprobado Por', 'Observaciones']);
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
      e.parameter.dias || '7',
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

  var ss = SpreadsheetApp.openById(HOJA_ID);
  var action = params.action;

  if (action === 'nueva_solicitud') {
    var sheet = findSheetByName(ss, 'solicitudes');
    if (!sheet) sheet = findSheetByHeaders(ss, ['ID', 'Estado', 'Motivo']);
    if (!sheet) {
      sheet = ss.insertSheet('solicitudes');
      sheet.appendRow(['ID', 'Cédula', 'Nombre', 'Fecha Solicitud', 'Fecha Inicio', 'Fecha Fin', 'Días', 'Motivo', 'Estado', 'Aprobado Por', 'Observaciones']);
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
      params.dias || '7',
      params.motivo || '',
      'Pendiente',
      '',
      ''
    ]);

    return ContentService.createTextOutput(JSON.stringify({success: true, id: id})).setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'aprobar' || action === 'rechazar') {
    var sheet = findSheetByName(ss, 'solicitudes');
    if (!sheet) sheet = findSheetByHeaders(ss, ['ID', 'Estado', 'Motivo']);
    var data = sheet.getDataRange().getValues();
    var headers = data[0];

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == params.id) {
        var estado = action === 'aprobar' ? 'Aprobada' : 'Rechazada';
        sheet.getRange(i + 1, headers.indexOf('Estado') + 1).setValue(estado);
        sheet.getRange(i + 1, headers.indexOf('Aprobado Por') + 1).setValue(params.aprobadoPor || '');
        sheet.getRange(i + 1, headers.indexOf('Observaciones') + 1).setValue(params.observaciones || '');

        if (action === 'aprobar') {
          var empSheet = findSheetByName(ss, 'empleados');
          if (!empSheet) empSheet = findSheetByHeaders(ss, ['cedula', 'vacaciones']);
          var empData = empSheet.getDataRange().getValues();
          var empHeaders = empData[0];
          for (var j = 1; j < empData.length; j++) {
            if (empData[j][empHeaders.indexOf('cedula')] == data[i][headers.indexOf('Cedula')]) {
              var vacIdx = empHeaders.indexOf('vacaciones');
              empSheet.getRange(j + 1, vacIdx + 1).setValue(1);
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
