/**
 * Backend de Control de Carrera (Ultrarrápido v2).
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
  const ss = SpreadsheetApp.getActive();
  ensureSheet_(ss, RUNNERS_SHEET, RUNNER_HEADERS, true);
  ensureSheet_(ss, CONFIG_SHEET, ['key', 'value'], true);
  ensureSheet_(ss, LOG_SHEET, ['id','bib','time','recorded_at','matched','was_official','device_id','operation_id'], true);
  ensureSheet_(ss, OPS_SHEET, ['operation_id','action','device_id','processed_at'], true);
  ss.setSpreadsheetTimeZone(TZ);
}

function doGet() {
  try {
    const ss = SpreadsheetApp.getActive();
    return json_({
      ok: true,
      serverTime: panamaIso_(new Date()),
      runners: readObjects_(ss, RUNNERS_SHEET),
      config: getConfig_(ss),
      arrivalLog: readObjects_(ss, LOG_SHEET).slice(-80).reverse()
    });
  } catch (error) {
    return json_({ ok: false, error: error.message });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const ss = SpreadsheetApp.getActive();
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (!data.action) throw new Error('Acción requerida');
    if (data.operationId && operationExists_(ss, data.operationId)) {
      return json_({ ok: true, duplicate: true, serverTime: panamaIso_(new Date()) });
    }
    const result = route_(ss, data);
    if (data.operationId) recordOperation_(ss, data);
    return json_(Object.assign({ ok: true, serverTime: panamaIso_(new Date()) }, result || {}));
  } catch (error) {
    return json_({ ok: false, error: error.message, serverTime: panamaIso_(new Date()) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function route_(ss, data) {
  switch (data.action) {
    case 'mergeRunners': return mergeRunners_(ss, data.runners || [], data.deviceId);
    case 'replaceAll': return mergeRunners_(ss, data.runners || [], data.deviceId);
    case 'updateStatus': return updateRunner_(ss, data, ['status','delivered_at','returned_at']);
    case 'updateArrival': return updateRunner_(ss, data, ['arrival_time','arrival_timestamp']);
    case 'quickArrival': return quickArrival_(ss, data);
    case 'correctQuickArrival': return correctQuickArrival_(ss, data);
    case 'setConfig': return setConfig_(ss, data);
    case 'reset': return reset_(ss);
    default: throw new Error('Acción no reconocida: ' + data.action);
  }
}

function mergeRunners_(ss, incoming, deviceId) {
  const current = readObjects_(ss, RUNNERS_SHEET);
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
  writeObjects_(ss, RUNNERS_SHEET, RUNNER_HEADERS, Object.keys(byBib).map(key => byBib[key]));
  return { created, updated, total: Object.keys(byBib).length };
}

function updateRunner_(ss, data, fields) {
  const sheet = ss.getSheetByName(RUNNERS_SHEET);
  if (!sheet) throw new Error('Hoja de corredores no encontrada');
  const range = sheet.getDataRange();
  const values = range.getValues();
  if (values.length < 2) throw new Error('Sin datos en la hoja');
  const headers = values[0];
  const bibCol = headers.indexOf('bib');
  const versionCol = headers.indexOf('version');
  const updatedAtCol = headers.indexOf('updated_at');
  const updatedByCol = headers.indexOf('updated_by');

  for (let row = 1; row < values.length; row++) {
    if (String(values[row][bibCol]) !== String(data.bib)) continue;
    const currentVersion = Number(values[row][versionCol] || 0);
    if (data.expectedVersion !== undefined && Number(data.expectedVersion) !== currentVersion) {
      return { conflict: true, message: 'El corredor fue modificado por otro dispositivo.', runner: rowObject_(headers, values[row]) };
    }
    const rowValues = values[row];
    fields.forEach(field => {
      const col = headers.indexOf(field);
      if (col >= 0) rowValues[col] = data[field] || '';
    });
    if (versionCol >= 0) rowValues[versionCol] = currentVersion + 1;
    if (updatedAtCol >= 0) rowValues[updatedAtCol] = panamaIso_(new Date());
    if (updatedByCol >= 0) rowValues[updatedByCol] = data.deviceId || '';

    // Escritura ultrarrápida de fila completa en 1 sola llamada API (~100ms)
    sheet.getRange(row + 1, 1, 1, headers.length).setValues([rowValues]);
    return { version: currentVersion + 1 };
  }
  throw new Error('Dorsal no encontrado: ' + data.bib);
}

function quickArrival_(ss, data) {
  const runners = readObjects_(ss, RUNNERS_SHEET);
  const runner = runners.find(r => String(r.bib) === String(data.bib));
  const wasOfficial = !!runner && !runner.arrival_time;
  if (wasOfficial) {
    data.expectedVersion = Number(runner.version || 0);
    data.arrival_time = data.time;
    updateRunner_(ss, data, ['arrival_time','arrival_timestamp']);
  }
  const id = Utilities.getUuid();
  appendObjectDirect_(ss, LOG_SHEET, [
    id, data.bib, data.time, data.arrival_timestamp || panamaIso_(new Date()),
    !!runner, wasOfficial, data.deviceId || '', data.operationId || ''
  ]);
  return {
    id, matched: !!runner, wasOfficial, existingTime: runner ? runner.arrival_time : '',
    runnerName: runner ? [runner.name, runner.surname].filter(Boolean).join(' ') : ''
  };
}

function correctQuickArrival_(ss, data) {
  const runners = readObjects_(ss, RUNNERS_SHEET);
  const oldRunner = runners.find(r => String(r.bib) === String(data.oldBib));
  const newRunner = runners.find(r => String(r.bib) === String(data.newBib));
  if (!newRunner) return { matched: false, wasOfficial: false };
  if (oldRunner && oldRunner.arrival_time === data.time) {
    updateRunner_(ss, { bib: data.oldBib, arrival_time: '', arrival_timestamp: '', deviceId: data.deviceId }, ['arrival_time','arrival_timestamp']);
  }
  let wasOfficial = false;
  if (!newRunner.arrival_time) {
    updateRunner_(ss, {
      bib: data.newBib, arrival_time: data.time, arrival_timestamp: data.clientTime,
      deviceId: data.deviceId
    }, ['arrival_time','arrival_timestamp']);
    wasOfficial = true;
  }
  appendObjectDirect_(ss, LOG_SHEET, [
    Utilities.getUuid(), data.newBib, data.time, data.clientTime,
    true, wasOfficial, data.deviceId || '', data.operationId || ''
  ]);
  return { matched: true, wasOfficial, runnerName: [newRunner.name, newRunner.surname].filter(Boolean).join(' ') };
}

function getConfig_(ss) {
  const rows = readObjects_(ss, CONFIG_SHEET);
  const map = {};
  rows.forEach(r => map[r.key] = r.value);
  let heats = {};
  try { heats = JSON.parse(map.heatStartTimes || '{}'); } catch (_) {}
  return { raceDate: map.raceDate || '', heatStartTimes: heats };
}

function setConfig_(ss, data) {
  writeObjects_(ss, CONFIG_SHEET, ['key','value'], [
    { key: 'raceDate', value: data.raceDate || '' },
    { key: 'heatStartTimes', value: JSON.stringify(data.heatStartTimes || {}) }
  ]);
  return {};
}

function reset_(ss) {
  writeObjects_(ss, RUNNERS_SHEET, RUNNER_HEADERS, []);
  writeObjects_(ss, LOG_SHEET, ['id','bib','time','recorded_at','matched','was_official','device_id','operation_id'], []);
  return {};
}

function ensureSheet_(ss, name, headers, forceCheckHeaders) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return sheet;
  }
  if (forceCheckHeaders && sheet.getLastRow() > 0) {
    const existing = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
    headers.forEach(header => {
      if (existing.indexOf(header) < 0) {
        sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header);
        existing.push(header);
      }
    });
  }
  return sheet;
}

function readObjects_(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getDataRange().getValues();
  return values.slice(1).filter(row => row.some(v => v !== '')).map(row => rowObject_(values[0], row));
}

function rowObject_(headers, row) {
  const object = {};
  headers.forEach((header, index) => object[header] = row[index]);
  return object;
}

function writeObjects_(ss, name, headers, objects) {
  const sheet = ensureSheet_(ss, name, headers, true);
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (objects.length) {
    sheet.getRange(2, 1, objects.length, headers.length)
      .setValues(objects.map(object => headers.map(header => object[header] === undefined ? '' : object[header])));
  }
}

function appendObjectDirect_(ss, name, rowValues) {
  const sheet = ss.getSheetByName(name);
  if (sheet) sheet.appendRow(rowValues);
}

function operationExists_(ss, id) {
  if (!id) return false;
  const sheet = ss.getSheetByName(OPS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return false;
  const finder = sheet.getRange(1, 1, sheet.getLastRow(), 1).createTextFinder(String(id)).matchEntireCell(true);
  return finder.findNext() !== null;
}

function recordOperation_(ss, data) {
  appendObjectDirect_(ss, OPS_SHEET, [data.operationId, data.action, data.deviceId || '', panamaIso_(new Date())]);
}

function panamaIso_(date) {
  return Utilities.formatDate(date, TZ, "yyyy-MM-dd'T'HH:mm:ss'-05:00'");
}

function json_(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
