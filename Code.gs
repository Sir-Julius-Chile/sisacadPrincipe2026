// Code.gs - SLEP Marga Marga Sync v4.0
// Cada tipo de dato en su propia hoja (alumnos, notas, usuarios, etc.)
// IMPORTANTE: Crear desde Extensiones > Apps Script dentro de la hoja de calculo.
// Implementar como: Ejecutar como: Yo, Acceso: Cualquier usuario, incluso anonimo

var PREFIX = 'SLEP_'; // prefijo para todas las hojas de datos

function sheetName(name) { return PREFIX + name; }

function doGet(e) {
  try {
    var params = (e && e.parameter) || {};
    var action = params.action || '';
    var callback = params.callback || '';
    var result;
    if (action === 'load') { result = loadData(); }
    else if (action === 'test') { result = testSheet(); }
    else if (action === 'dump') { result = dumpDebug(); }
    else if (action === 'clear') { result = clearData(); }
    else { result = { success: true, message: 'Servidor activo v4.0', time: new Date().toISOString() }; }
    return jsonpOrJson(result, callback);
  } catch (err) {
    return jsonpOrJson({ success: false, error: err.toString() }, (e && e.parameter && e.parameter.callback) || '');
  }
}

function doPost(e) {
  try {
    var raw = null;
    if (e && e.parameter && e.parameter.payload) {
      raw = e.parameter.payload;
    } else if (e && e.postData && e.postData.contents && e.postData.type !== 'application/x-www-form-urlencoded') {
      raw = e.postData.contents;
    }
    if (!raw) {
      return respond({ success: false, error: 'No se recibieron datos. Use POST con payload en form-urlencoded.' });
    }
    var data = JSON.parse(raw);
    if (data && data.action === 'save' && data.data) {
      var result = saveDataSeparatSheets(data.data, data.timestamp);
      return respond(result);
    }
    if (data && data.action === 'save') {
      var result2 = saveDataSeparatSheets(data, data.timestamp);
      return respond(result2);
    }
    return respond({ success: false, error: 'Formato invalido. Envie {action:"save", data:{...}}' });
  } catch (err) {
    return respond({ success: false, error: 'Error: ' + err.toString() });
  }
}

// ========== TEST ==========

function testSheet() {
  try {
    var ss = getSpreadsheet();
    var meta = ss.getSheetByName(sheetName('meta'));
    var info = { existeMeta: !!meta, hojas: [] };
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getName().indexOf(PREFIX) === 0) {
        info.hojas.push({ nombre: sheets[i].getName(), filas: sheets[i].getLastRow(), columnas: sheets[i].getLastColumn() });
      }
    }
    var version = '';
    if (meta) {
      var r = meta.getRange('A1:B1').getValues();
      version = (r[0][0] || '') + ' | ' + (r[0][1] || '');
    }
    info.version = version;
    info.success = true;
    info.message = 'Spreadsheet: "' + ss.getName() + '"';
    info.tieneDatos = info.hojas.length > 1;
    return info;
  } catch (err) {
    return { success: false, error: err.toString(), ayuda: 'El script debe crearse desde Extensiones > Apps Script dentro de la hoja de calculo.' };
  }
}

// ========== SAVE (multi-sheet) ==========

function saveDataSeparatSheets(data, timestamp) {
  try {
    var ss = getSpreadsheet();
    var ts = timestamp || new Date().toISOString();
    var log = [];
    var errors = [];

    // 1. Meta
    try {
      var meta = getOrCreateSheet(ss, 'meta', ['version', 'updatedAt', 'gasUrl']);
      meta.getRange('A2:C2').setValues([['v4.0', ts, data.gasUrl || '']]);
    } catch(e) { errors.push('meta:' + e.toString().substring(0, 100)); }

    // 2. Usuarios
    try {
      if (data.users && data.users.length) {
        var usersSheet = getOrCreateSheet(ss, 'usuarios', ['id', 'username', 'passwordHash', 'role', 'permissions', 'headTeacherCourse']);
        var userRows = [];
        for (var ui = 0; ui < data.users.length; ui++) {
          var u = data.users[ui];
          userRows.push([u.id || '', u.username || '', u.passwordHash || '', u.role || '', JSON.stringify(u.permissions || {}), u.headTeacherCourse || '']);
        }
        clearAndFill(usersSheet, userRows);
        log.push('usuarios:' + userRows.length);
      }
    } catch(e) { errors.push('usuarios:' + e.toString().substring(0, 100)); }

    // 3. Alumnos
    try {
      if (data.students) {
        var alumHeaders = ['course', 'id', 'name', 'rut', 'ap1_nombre', 'ap1_rut', 'ap1_parentesco', 'ap1_fono', 'ap1_email', 'ap1_direccion', 'ap2_nombre', 'ap2_rut', 'ap2_parentesco', 'ap2_fono', 'ap2_email', 'condiciones', 'pie', 'pie_diagnostico'];
        var alumSheet = getOrCreateSheet(ss, 'alumnos', alumHeaders);
        // Migrate headers if sheet exists with old schema
        var alumExisting = ss.getSheetByName(sheetName('alumnos'));
        if (alumExisting) {
          var hRowAlum = alumExisting.getRange(1, 1, 1, alumHeaders.length).getValues()[0];
          if (hRowAlum.join('|') !== alumHeaders.join('|')) {
            ss.deleteSheet(alumExisting);
            alumSheet = ss.insertSheet(sheetName('alumnos'));
            alumSheet.getRange(1, 1, 1, alumHeaders.length).setValues([alumHeaders]);
          }
        }
        var alumRows = [];
        var courseKeys = Object.keys(data.students);
        for (var ci = 0; ci < courseKeys.length; ci++) {
          var course = courseKeys[ci];
          var list = data.students[course] || [];
          for (var si = 0; si < list.length; si++) {
            var s = list[si];
            alumRows.push([course, s.id || '', s.name || '', s.rut || '',
              s.ap1_nombre || '', s.ap1_rut || '', s.ap1_parentesco || '', s.ap1_fono || '', s.ap1_email || '', s.ap1_direccion || '',
              s.ap2_nombre || '', s.ap2_rut || '', s.ap2_parentesco || '', s.ap2_fono || '', s.ap2_email || '',
              s.condiciones || '', s.pie ? 'SI' : 'NO', s.pie_diagnostico || '']);
          }
        }
        clearAndFill(alumSheet, alumRows);
        log.push('alumnos:' + alumRows.length + ' (de ' + courseKeys.length + ' cursos)');
      }
    } catch(e) { errors.push('alumnos:' + e.toString().substring(0, 100)); }

    // 4. Notas
    try {
      if (data.grades) {
        var allKeys = Object.keys(data.grades);
        logDebug('notas total keys', allKeys.length + ' claves en data.grades');
        var notasSheet = getOrCreateSheet(ss, 'notas', ['course', 'studentId', 'subject', 'semestre', 'n1', 'n2', 'n3', 'n4', 'n5', 'n6', 'n7', 'n8', 'n9', 'n10']);
        var notaRows = [];
        for (var gki = 0; gki < allKeys.length; gki++) {
          var key = allKeys[gki];
          var parts = key.split('_');
          var course = parts[0];
          var studentId = parts[parts.length - 1];
          var subject = parts.slice(1, -1).join('_');
          var gradeData = data.grades[key] || {};
          var row1 = [course, studentId, subject, '1'];
          for (var ni = 1; ni <= 10; ni++) {
            var v1 = gradeData['s1n' + ni];
            row1.push(v1 !== undefined && v1 !== null ? v1 : '');
          }
          notaRows.push(row1);
          var row2 = [course, studentId, subject, '2'];
          for (var ni2 = 1; ni2 <= 10; ni2++) {
            var v2 = gradeData['s2n' + ni2];
            row2.push(v2 !== undefined && v2 !== null ? v2 : '');
          }
          notaRows.push(row2);
        }
        logDebug('notas filas generadas', notaRows.length + ' filas de ' + allKeys.length + ' claves');
        clearAndFill(notasSheet, notaRows);
        log.push('notas:' + notaRows.length);
      } else {
        logDebug('notas', 'data.grades es null/undefined');
      }
    } catch(e) { errors.push('notas:' + e.toString().substring(0, 200)); logDebug('notas ERROR', e.toString()); }

    // 5. Asistencia
    try {
      if (data.attendance) {
        var asisSheet = getOrCreateSheet(ss, 'asistencia', ['course', 'studentId', 'date', 'status']);
        var asisRows = [];
        var aCourseKeys = Object.keys(data.attendance);
        for (var aci = 0; aci < aCourseKeys.length; aci++) {
          var ac = aCourseKeys[aci];
          var courseAtt = data.attendance[ac] || {};
          var aStudentIds = Object.keys(courseAtt);
          for (var asi2 = 0; asi2 < aStudentIds.length; asi2++) {
            var asid = aStudentIds[asi2];
            var days = courseAtt[asid] || {};
            var dateKeys = Object.keys(days);
            for (var dki = 0; dki < dateKeys.length; dki++) {
              var dk = dateKeys[dki];
              asisRows.push([ac, asid, dk, days[dk]]);
            }
          }
        }
        clearAndFill(asisSheet, asisRows);
        log.push('asistencia:' + asisRows.length);
      }
    } catch(e) { errors.push('asistencia:' + e.toString().substring(0, 100)); }

    // 6. Observaciones
    try {
      if (data.observations) {
        var obsSheet = getOrCreateSheet(ss, 'observaciones', ['course', 'studentId', 'date', 'text', 'type']);
        var obsRows = [];
        var oCourseKeys = Object.keys(data.observations);
        for (var oci = 0; oci < oCourseKeys.length; oci++) {
          var oc = oCourseKeys[oci];
          var courseObs = data.observations[oc] || {};
          var oStudentIds = Object.keys(courseObs);
          for (var osi = 0; osi < oStudentIds.length; osi++) {
            var osid = oStudentIds[osi];
            var entries = courseObs[osid] || [];
            for (var ei = 0; ei < entries.length; ei++) {
              var e = entries[ei];
              obsRows.push([oc, osid, e.date || '', e.text || '', e.type || '']);
            }
          }
        }
        clearAndFill(obsSheet, obsRows);
        log.push('observaciones:' + obsRows.length);
      }
    } catch(e) { errors.push('observaciones:' + e.toString().substring(0, 100)); }

    // 7. Citaciones
    try {
      if (data._citations && data._citations.length) {
        var citHeaders = ['id', 'courseId', 'studentId', 'date', 'horaInicio', 'horaTermino', 'departamento', 'motivo', 'observacion', 'pie', 'asistio', 'createdBy', 'createdAt'];
        var citSheet = getOrCreateSheet(ss, 'citaciones', citHeaders);
        // Migrate headers if sheet exists with old schema
        if (ss.getSheetByName(sheetName('citaciones'))) {
          var hRow = citSheet.getRange(1, 1, 1, citHeaders.length).getValues()[0];
          if (hRow.join('|') !== citHeaders.join('|')) {
            ss.deleteSheet(citSheet);
            citSheet = ss.insertSheet(sheetName('citaciones'));
            citSheet.getRange(1, 1, 1, citHeaders.length).setValues([citHeaders]);
          }
        }
        var citRows = [];
        for (var cti = 0; cti < data._citations.length; cti++) {
          var c = data._citations[cti];
          citRows.push([
            c.id || '', c.courseId || '', c.studentId || '',
            c.date || '', c.horaInicio || '', c.horaTermino || '',
            c.departamento || '', c.motivo || '', c.observacion || '',
            c.pie ? 'SI' : 'NO',
            c.asistio === true ? 'SI' : (c.asistio === false ? 'NO' : ''),
            c.createdBy || '', c.createdAt || ''
          ]);
        }
        clearAndFill(citSheet, citRows);
        log.push('citaciones:' + citRows.length);
      }
    } catch(e) { errors.push('citaciones:' + e.toString().substring(0, 100)); }

    // 8. Grupos
    try {
      if (data._subjectGroups) {
        var grpSheet = getOrCreateSheet(ss, 'grupos', ['course', 'groupName', 'subjects']);
        var grpRows = [];
        var grpCourseKeys = Object.keys(data._subjectGroups);
        for (var grci = 0; grci < grpCourseKeys.length; grci++) {
          var grc = grpCourseKeys[grci];
          var courseGroups = data._subjectGroups[grc] || {};
          var groups = courseGroups.groups || [];
          for (var ggi = 0; ggi < groups.length; ggi++) {
            var g = groups[ggi];
            grpRows.push([grc, g.name || '', JSON.stringify(g.subjects || [])]);
          }
        }
        clearAndFill(grpSheet, grpRows);
        log.push('grupos:' + grpRows.length);
      }
    } catch(e) { errors.push('grupos:' + e.toString().substring(0, 100)); }

    // 9. Calendario
    try {
      if (data._schoolYear) {
        var calSheet = getOrCreateSheet(ss, 'calendario', ['key', 'value']);
        var calRows = [];
        var sy = data._schoolYear;
        calRows.push(['start', sy.start || '']);
        calRows.push(['end', sy.end || '']);
        calRows.push(['end4medio', sy.end4medio || '']);
        calRows.push(['winterStart', sy.winterStart || '']);
        calRows.push(['winterEnd', sy.winterEnd || '']);
        var holidays = sy.holidays || [];
        for (var hi = 0; hi < holidays.length; hi++) {
          calRows.push(['holiday', holidays[hi]]);
        }
        clearAndFill(calSheet, calRows);
        log.push('calendario:' + calRows.length);
      }
    } catch(e) { errors.push('calendario:' + e.toString().substring(0, 100)); }

    var resultMsg = log.join(', ');
    if (errors.length) resultMsg += ' | ERRORES: ' + errors.join('; ');
    return { success: true, message: 'Datos guardados en hojas separadas', detalle: log.join(', '), errores: errors, timestamp: ts };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ========== LOAD (multi-sheet) ==========

function loadData() {
  try {
    var ss = getSpreadsheet();
    var result = {};

    // Leer meta
    var meta = ss.getSheetByName(sheetName('meta'));
    if (!meta) return { success: false, error: 'No existe la hoja "' + sheetName('meta') + '". Debe subir datos primero.' };

    // Leer usuarios
    var usersSheet = ss.getSheetByName(sheetName('usuarios'));
    if (usersSheet && usersSheet.getLastRow() > 1) {
      result.users = readTable(usersSheet, ['id', 'username', 'passwordHash', 'role', 'permissions', 'headTeacherCourse']);
      for (var ui = 0; ui < result.users.length; ui++) {
        try { if (typeof result.users[ui].permissions === 'string') result.users[ui].permissions = JSON.parse(result.users[ui].permissions); } catch(e) { result.users[ui].permissions = {}; }
      }
    }

    // Leer alumnos
    result.students = {};
    var alumSheet = ss.getSheetByName(sheetName('alumnos'));
    if (alumSheet && alumSheet.getLastRow() > 1) {
      var alumRows = readTable(alumSheet, ['course', 'id', 'name', 'rut', 'ap1_nombre', 'ap1_rut', 'ap1_parentesco', 'ap1_fono', 'ap1_email', 'ap1_direccion', 'ap2_nombre', 'ap2_rut', 'ap2_parentesco', 'ap2_fono', 'ap2_email', 'condiciones', 'pie', 'pie_diagnostico']);
      for (var ai = 0; ai < alumRows.length; ai++) {
        var a = alumRows[ai];
        if (!result.students[a.course]) result.students[a.course] = [];
        result.students[a.course].push({
          id: a.id, name: a.name, rut: a.rut,
          ap1_nombre: a.ap1_nombre, ap1_rut: a.ap1_rut, ap1_parentesco: a.ap1_parentesco, ap1_fono: a.ap1_fono, ap1_email: a.ap1_email, ap1_direccion: a.ap1_direccion,
          ap2_nombre: a.ap2_nombre, ap2_rut: a.ap2_rut, ap2_parentesco: a.ap2_parentesco, ap2_fono: a.ap2_fono, ap2_email: a.ap2_email,
          condiciones: a.condiciones, pie: a.pie === 'SI', pie_diagnostico: a.pie_diagnostico
        });
      }
    }

    // Leer notas (formato plano: clave = course_subject_studentId, valores = s1n1..s2n10)
    result.grades = {};
    var notasSheet = ss.getSheetByName(sheetName('notas'));
    if (notasSheet && notasSheet.getLastRow() > 1) {
      var notaRows = readTable(notasSheet, ['course', 'studentId', 'subject', 'semestre', 'n1', 'n2', 'n3', 'n4', 'n5', 'n6', 'n7', 'n8', 'n9', 'n10']);
      var notasCount = 0;
      for (var ni = 0; ni < notaRows.length; ni++) {
        var nr = notaRows[ni];
        if (!nr.studentId || nr.studentId === '') continue;
        var key = nr.course + '_' + nr.subject + '_' + nr.studentId;
        if (!result.grades[key]) result.grades[key] = {};
        var semPrefix = 's' + nr.semestre;
        for (var ni2 = 1; ni2 <= 10; ni2++) {
          var v = nr['n' + ni2];
          result.grades[key][semPrefix + 'n' + ni2] = v !== '' ? Number(v) : null;
        }
        notasCount++;
      }
      logDebug('load notas', notasCount + ' filas leidas, ' + Object.keys(result.grades).length + ' claves unicas');
    }

    // Leer asistencia
    result.attendance = {};
    var asisSheet = ss.getSheetByName(sheetName('asistencia'));
    if (asisSheet && asisSheet.getLastRow() > 1) {
      var asisRows = readTable(asisSheet, ['course', 'studentId', 'date', 'status']);
      for (var asi = 0; asi < asisRows.length; asi++) {
        var ar = asisRows[asi];
        var dateKey = gsDateStr(ar.date, 'yyyy-MM-dd');
        if (!result.attendance[ar.course]) result.attendance[ar.course] = {};
        if (!result.attendance[ar.course][ar.studentId]) result.attendance[ar.course][ar.studentId] = {};
        result.attendance[ar.course][ar.studentId][dateKey] = ar.status;
      }
    }

    // Leer observaciones
    result.observations = {};
    var obsSheet = ss.getSheetByName(sheetName('observaciones'));
    if (obsSheet && obsSheet.getLastRow() > 1) {
      var obsRows = readTable(obsSheet, ['course', 'studentId', 'date', 'text', 'type']);
      for (var oi = 0; oi < obsRows.length; oi++) {
        var or_ = obsRows[oi];
        if (!result.observations[or_.course]) result.observations[or_.course] = {};
        if (!result.observations[or_.course][or_.studentId]) result.observations[or_.course][or_.studentId] = [];
        result.observations[or_.course][or_.studentId].push({ date: gsDateStr(or_.date, 'yyyy-MM-dd'), text: or_.text, type: or_.type });
      }
    }

    // Leer citaciones
    result._citations = [];
    var citSheet = ss.getSheetByName(sheetName('citaciones'));
    if (citSheet && citSheet.getLastRow() > 1) {
      var citRows = readTable(citSheet, ['id', 'courseId', 'studentId', 'date', 'horaInicio', 'horaTermino', 'departamento', 'motivo', 'observacion', 'pie', 'asistio', 'createdBy', 'createdAt']);
      for (var ci = 0; ci < citRows.length; ci++) {
        var cit = citRows[ci];
        result._citations.push({
          id: cit.id,
          courseId: cit.courseId,
          studentId: cit.studentId,
          date: gsDateStr(cit.date, 'yyyy-MM-dd'),
          horaInicio: gsDateStr(cit.horaInicio, 'HH:mm'),
          horaTermino: gsDateStr(cit.horaTermino, 'HH:mm'),
          departamento: cit.departamento,
          motivo: cit.motivo,
          observacion: cit.observacion,
          pie: cit.pie === 'SI',
          asistio: cit.asistio === 'SI' ? true : (cit.asistio === 'NO' ? false : null),
          createdBy: cit.createdBy,
          createdAt: gsDateStr(cit.createdAt, 'yyyy-MM-dd HH:mm:ss')
        });
      }
    }

    // Leer grupos
    result._subjectGroups = {};
    var grpSheet = ss.getSheetByName(sheetName('grupos'));
    if (grpSheet && grpSheet.getLastRow() > 1) {
      var grpRows = readTable(grpSheet, ['course', 'groupName', 'subjects']);
      for (var gi = 0; gi < grpRows.length; gi++) {
        var gr = grpRows[gi];
        if (!result._subjectGroups[gr.course]) result._subjectGroups[gr.course] = { groups: [] };
        try { result._subjectGroups[gr.course].groups.push({ name: gr.groupName, subjects: JSON.parse(gr.subjects) }); } catch(e) {}
      }
    }

    // Leer calendario
    result._schoolYear = { holidays: [] };
    var calSheet = ss.getSheetByName(sheetName('calendario'));
    if (calSheet && calSheet.getLastRow() > 1) {
      var calRows = readTable(calSheet, ['key', 'value']);
      for (var cali = 0; cali < calRows.length; cali++) {
        var cr = calRows[cali];
        var val = gsDateStr(cr.value, 'yyyy-MM-dd');
        if (cr.key === 'holiday') { result._schoolYear.holidays.push(val); }
        else { result._schoolYear[cr.key] = val; }
      }
    }

    // Leer gasUrl de meta
    var metaVals = meta.getRange('C2').getValue() || '';
    result.gasUrl = metaVals;

    // Cursos vacios para asistencia y observaciones
    var coursesList = Object.keys(result.students);
    for (var ci = 0; ci < coursesList.length; ci++) {
      if (!result.attendance[coursesList[ci]]) result.attendance[coursesList[ci]] = {};
      if (!result.observations[coursesList[ci]]) result.observations[coursesList[ci]] = {};
    }

    var ts = '';
    try { ts = meta.getRange('B2').getValue() || ''; } catch(e) {}
    return { success: true, data: result, timestamp: ts.toString().substring(0, 40) };
  } catch (err) {
    return { success: false, error: 'Error al cargar: ' + err.toString() };
  }
}

// ========== HELPERS ==========

function getOrCreateSheet(ss, name, headers) {
  var fullName = sheetName(name);
  var sheet = ss.getSheetByName(fullName);
  if (!sheet) {
    sheet = ss.insertSheet(fullName);
    if (headers && headers.length) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
  }
  return sheet;
}

function clearAndFill(sheet, rows) {
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
  if (rows && rows.length) {
    var numCols = rows[0].length;
    sheet.getRange(2, 1, rows.length, numCols).setValues(rows);
  }
}

function readTable(sheet, columns) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var data = sheet.getRange(2, 1, lastRow - 1, columns.length).getValues();
  var result = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (row.join('').trim() === '') continue; // skip empty rows
    var obj = {};
    for (var j = 0; j < columns.length; j++) {
      obj[columns[j]] = row[j];
    }
    result.push(obj);
  }
  return result;
}

// Convert Google Sheets Date object to formatted string (handles auto-date-conversion)
function gsDateStr(val, fmt) {
  if (val instanceof Date && !isNaN(val)) {
    return Utilities.formatDate(val, 'America/Santiago', fmt);
  }
  return String(val || '');
}

function clearData() {
  try {
    var ss = getSpreadsheet();
    var sheets = ss.getSheets();
    var names = [];
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getName().indexOf(PREFIX) === 0) {
        names.push(sheets[i].getName());
        ss.deleteSheet(sheets[i]);
      }
    }
    return { success: true, deleted: names };
  } catch(err) {
    return { success: false, error: err.toString() };
  }
}

function getSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('No active spreadsheet. Create script from Extensions > Apps Script inside the spreadsheet, not as standalone.');
  return ss;
}

// ========== DEBUG ==========

function logDebug(label, info) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;
    var sheet = ss.getSheetByName(sheetName('debug'));
    if (!sheet) { sheet = ss.insertSheet(sheetName('debug')); sheet.getRange('A1:C1').setValues([['time', 'label', 'info']]); }
    var val = typeof info === 'string' ? info : JSON.stringify(info);
    sheet.appendRow([new Date().toISOString(), label, val.substring(0, 30000)]);
  } catch(e) { /* silent */ }
}

function dumpDebug() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(sheetName('debug'));
    if (!sheet) return { success: false, error: 'No debug sheet' };
    var rows = sheet.getDataRange().getValues();
    var last20 = [];
    for (var i = Math.max(0, rows.length - 21); i < rows.length; i++) {
      last20.push({ time: rows[i][0], label: rows[i][1], info: rows[i][2] });
    }
    return { success: true, entries: last20, total: rows.length };
  } catch(err) {
    return { success: false, error: err.toString() };
  }
}

// ========== RESPONSE ==========

function jsonpOrJson(obj, callback) {
  var json = JSON.stringify(obj);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
