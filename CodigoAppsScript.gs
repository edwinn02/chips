/**
 * Backend de Control de Carrera.
 * 1. Cree un Google Sheet vacío y abra Extensiones > Apps Script.
 * 2. Pegue este archivo, guarde y ejecute setup() una vez.
 * 3. Implemente como Aplicación web: ejecutar como usted y acceso para quienes operarán la carrera.
 * 4. Pegue la URL /exec en Configuración del sistema.
 */
const TZ = 'America/Panama';
const RUNNERS_SHEET = 'Corredores';
const CONFIG_SHEET = 'Configuracion';
const LOG_SHEET = 'LlegadasLog';
const OPS_SHEET = 'Operaciones';
const RUNNER_HEADERS = [
  'bib','item','name','surname','gender','epc','country','city','id_card','birthday','age',
  'age_group','age_group_2','phone','email','team','team_index','type','area','status',
  'delivered_at','returned_at','arrival_time','arrival_timestamp','version','updated_at','updated_by'
];

function setup() {
  ensureSheet_(RUNNERS_SHEET, RUNNER_HEADERS);
  ensureSheet_(CONFIG_SHEET, ['key', 'value']);
  ensureSheet_(LOG_SHEET, ['id','bib','time','recorded_at','matched','was_official','device_id','operation_id']);
  ensureSheet_(OPS_SHEET, ['operation_id','action','device_id','processed_at']);
  SpreadsheetApp.getActive().setSpreadsheetTimeZone(TZ);
}

function doGet() {
  try {
    setup();
    return json_({
      ok: true,
      serverTime: panamaIso_(new Date()),
      runners: readObjects_(RUNNERS_SHEET),
      config: getConfig_(),
      arrivalLog: readObjects_(LOG_SHEET).slice(-80).reverse()
    });
  } catch (error) {
    return json_({ ok: false, error: error.message });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    setup();
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (!data.action) throw new Error('Acción requerida');
    if (data.operationId && operationExists_(data.operationId)) {
      return json_({ ok: true, duplicate: true, serverTime: panamaIso_(new Date()) });
    }
    const result = route_(data);
    if (data.operationId) recordOperation_(data);
    return json_(Object.assign({ ok: true, serverTime: panamaIso_(new Date()) }, result || {}));
  } catch (error) {
    return json_({ ok: false, error: error.message, serverTime: panamaIso_(new Date()) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function route_(data) {
  switch (data.action) {
    case 'mergeRunners': return mergeRunners_(data.runners || [], data.deviceId);
    case 'replaceAll': return mergeRunners_(data.runners || [], data.deviceId);
    case 'updateStatus': return updateRunner_(data, ['status','delivered_at','returned_at']);
    case 'updateArrival': return updateRunner_(data, ['arrival_time','arrival_timestamp']);
    case 'quickArrival': return quickArrival_(data);
    case 'correctQuickArrival': return correctQuickArrival_(data);
    case 'setConfig': return setConfig_(data);
    case 'reset': return reset_();
    default: throw new Error('Acción no reconocida: ' + data.action);
  }
}

function mergeRunners_(incoming, deviceId) {
  const current = readObjects_(RUNNERS_SHEET);
  const byBib = {};
  current.forEach(r => byBib[String(r.bib)] = r);
  let created = 0, updated = 0;
  incoming.forEach(raw => {
    const bib = String(raw.bib || '').trim();
    if (!bib) return;
    const old = byBib[bib];
    if (!old) {
      byBib[bib] = Object.assign({}, raw, {
        bib, status: raw.status || 'pendiente', version: 1,
        updated_at: panamaIso_(new Date()), updated_by: deviceId || ''
      });
      created++;
      return;
    }
    const operational = {
      status: old.status || 'pendiente', delivered_at: old.delivered_at || '',
      returned_at: old.returned_at || '', arrival_time: old.arrival_time || '',
      arrival_timestamp: old.arrival_timestamp || ''
    };
    byBib[bib] = Object.assign({}, old, raw, operational, {
      bib, version: Number(old.version || 0) + 1,
      updated_at: panamaIso_(new Date()), updated_by: deviceId || ''
    });
    updated++;
  });
  writeObjects_(RUNNERS_SHEET, RUNNER_HEADERS, Object.keys(byBib).map(key => byBib[key]));
  return { created, updated, total: Object.keys(byBib).length };
}

function updateRunner_(data, fields) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(RUNNERS_SHEET);
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const bibCol = headers.indexOf('bib');
  const versionCol = headers.indexOf('version');
  for (let row = 1; row < values.length; row++) {
    if (String(values[row][bibCol]) !== String(data.bib)) continue;
    const currentVersion = Number(values[row][versionCol] || 0);
    if (data.expectedVersion !== undefined && Number(data.expectedVersion) !== currentVersion) {
      return { conflict: true, message: 'El corredor fue modificado por otro dispositivo.', runner: rowObject_(headers, values[row]) };
    }
    fields.forEach(field => {
      const col = headers.indexOf(field);
      if (col >= 0) sheet.getRange(row + 1, col + 1).setValue(data[field] || '');
    });
    sheet.getRange(row + 1, versionCol + 1).setValue(currentVersion + 1);
    setCell_(sheet, row + 1, headers, 'updated_at', panamaIso_(new Date()));
    setCell_(sheet, row + 1, headers, 'updated_by', data.deviceId || '');
    return { version: currentVersion + 1 };
  }
  throw new Error('Dorsal no encontrado: ' + data.bib);
}

function quickArrival_(data) {
  const runners = readObjects_(RUNNERS_SHEET);
  const runner = runners.find(r => String(r.bib) === String(data.bib));
  const wasOfficial = !!runner && !runner.arrival_time;
  if (wasOfficial) {
    data.expectedVersion = Number(runner.version || 0);
    data.arrival_time = data.time;
    updateRunner_(data, ['arrival_time','arrival_timestamp']);
  }
  const id = Utilities.getUuid();
  appendObject_(LOG_SHEET, {
    id, bib: data.bib, time: data.time, recorded_at: data.arrival_timestamp || panamaIso_(new Date()),
    matched: !!runner, was_official: wasOfficial, device_id: data.deviceId || '', operation_id: data.operationId || ''
  });
  return {
    id, matched: !!runner, wasOfficial, existingTime: runner ? runner.arrival_time : '',
    runnerName: runner ? [runner.name, runner.surname].filter(Boolean).join(' ') : ''
  };
}

function correctQuickArrival_(data) {
  const runners = readObjects_(RUNNERS_SHEET);
  const oldRunner = runners.find(r => String(r.bib) === String(data.oldBib));
  const newRunner = runners.find(r => String(r.bib) === String(data.newBib));
  if (!newRunner) return { matched: false, wasOfficial: false };
  if (oldRunner && oldRunner.arrival_time === data.time) {
    updateRunner_({ bib: data.oldBib, arrival_time: '', arrival_timestamp: '', deviceId: data.deviceId }, ['arrival_time','arrival_timestamp']);
  }
  let wasOfficial = false;
  if (!newRunner.arrival_time) {
    updateRunner_({
      bib: data.newBib, arrival_time: data.time, arrival_timestamp: data.clientTime,
      deviceId: data.deviceId
    }, ['arrival_time','arrival_timestamp']);
    wasOfficial = true;
  }
  appendObject_(LOG_SHEET, {
    id: Utilities.getUuid(), bib: data.newBib, time: data.time, recorded_at: data.clientTime,
    matched: true, was_official: wasOfficial, device_id: data.deviceId || '', operation_id: data.operationId || ''
  });
  return { matched: true, wasOfficial, runnerName: [newRunner.name, newRunner.surname].filter(Boolean).join(' ') };
}

function getConfig_() {
  const rows = readObjects_(CONFIG_SHEET);
  const map = {};
  rows.forEach(r => map[r.key] = r.value);
  let heats = {};
  try { heats = JSON.parse(map.heatStartTimes || '{}'); } catch (_) {}
  return { raceDate: map.raceDate || '', heatStartTimes: heats };
}

function setConfig_(data) {
  writeObjects_(CONFIG_SHEET, ['key','value'], [
    { key: 'raceDate', value: data.raceDate || '' },
    { key: 'heatStartTimes', value: JSON.stringify(data.heatStartTimes || {}) }
  ]);
  return {};
}

function reset_() {
  writeObjects_(RUNNERS_SHEET, RUNNER_HEADERS, []);
  writeObjects_(LOG_SHEET, ['id','bib','time','recorded_at','matched','was_official','device_id','operation_id'], []);
  return {};
}

function ensureSheet_(name, headers) {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  const existing = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
  headers.forEach(header => {
    if (existing.indexOf(header) < 0) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header);
      existing.push(header);
    }
  });
  return sheet;
}

function readObjects_(name) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getDataRange().getValues();
  return values.slice(1).filter(row => row.some(v => v !== '')).map(row => rowObject_(values[0], row));
}

function rowObject_(headers, row) {
  const object = {};
  headers.forEach((header, index) => object[header] = row[index]);
  return object;
}

function writeObjects_(name, headers, objects) {
  const sheet = ensureSheet_(name, headers);
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (objects.length) {
    sheet.getRange(2, 1, objects.length, headers.length)
      .setValues(objects.map(object => headers.map(header => object[header] === undefined ? '' : object[header])));
  }
}

function appendObject_(name, object) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(name);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  sheet.appendRow(headers.map(header => object[header] === undefined ? '' : object[header]));
}

function setCell_(sheet, row, headers, field, value) {
  const col = headers.indexOf(field);
  if (col >= 0) sheet.getRange(row, col + 1).setValue(value);
}

function operationExists_(id) {
  if (!id) return false;
  return readObjects_(OPS_SHEET).some(op => String(op.operation_id) === String(id));
}

function recordOperation_(data) {
  appendObject_(OPS_SHEET, {
    operation_id: data.operationId, action: data.action, device_id: data.deviceId || '',
    processed_at: panamaIso_(new Date())
  });
}

function panamaIso_(date) {
  return Utilities.formatDate(date, TZ, "yyyy-MM-dd'T'HH:mm:ss'-05:00'");
}

function json_(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
