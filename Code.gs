const SHEETS = {
  Regions: ['id', 'name'],
  Salesmen: ['id', 'name', 'pin', 'role', 'active', 'defaultRegionId'],
  Teams: ['id', 'name'],
  Entries: ['id', 'salesmanId', 'date', 'collection', 'sales', 'visits'],
  Challenges: ['id', 'salesmanId', 'date', 'collectionTarget', 'salesTarget', 'notes'],
  Monthly: ['id', 'salesmanId', 'monthKey', 'regionId', 'teamId', 'collectionTarget', 'salesTarget'],
  Openings: ['id', 'salesmanId', 'monthKey', 'amount'],
  Slabs: ['id', 'regionMonthKey', 'name', 'target', 'incentivePct'],
  OffDays: ['date'],
  Settings: ['key', 'value']
};

function doGet(e) {
  initSheets();
  return json_({ ok: true, data: readAll_() });
}

function doPost(e) {
  initSheets();
  const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  const action = body.action || 'saveAll';
  if (action === 'saveAll') writeAll_(body.payload || {});
  if (action === 'setMonthly') setMonthly_(body.payload || {});
  if (action === 'setOpening') setOpening_(body.payload || {});
  if (action === 'addRow') addRow_(body.payload.sheet, body.payload.row);
  if (action === 'updateRow') updateRow_(body.payload.sheet, body.payload.id, body.payload.patch);
  if (action === 'deleteRow') deleteRow_(body.payload.sheet, body.payload.id);
  return json_({ ok: true, data: readAll_() });
}

function initSheets() {
  Object.keys(SHEETS).forEach(name => ensureHeaders_(name, SHEETS[name]));
}

function ensureHeaders_(name, headers) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const existing = sh.getRange(1, 1, 1, lastCol).getValues()[0].filter(String);
  if (!existing.length) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    return sh;
  }
  const missing = headers.filter(h => existing.indexOf(h) === -1);
  if (missing.length) sh.getRange(1, existing.length + 1, 1, missing.length).setValues([missing]);
  return sh;
}

function readSheet_(name) {
  const sh = ensureHeaders_(name, SHEETS[name]);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  return values.slice(1).filter(r => r.some(v => v !== '')).map(row => headers.reduce((obj, h, i) => {
    obj[h] = row[i];
    if (obj[h] === 'TRUE') obj[h] = true;
    if (obj[h] === 'FALSE') obj[h] = false;
    return obj;
  }, {}));
}

function writeSheet_(name, rows) {
  const sh = ensureHeaders_(name, SHEETS[name]);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].filter(String);
  sh.clearContents();
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows && rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows.map(r => headers.map(h => r[h] == null ? '' : r[h])));
}

function readAll_() {
  const settingsRows = readSheet_('Settings');
  return {
    regions: readSheet_('Regions'),
    salesmen: readSheet_('Salesmen'),
    teams: readSheet_('Teams'),
    entries: readSheet_('Entries'),
    challenges: readSheet_('Challenges'),
    monthly: readSheet_('Monthly'),
    openings: readSheet_('Openings'),
    slabs: readSheet_('Slabs'),
    offDays: readSheet_('OffDays').map(r => r.date),
    settings: settingsRows.reduce((o, r) => (o[r.key] = r.value, o), {})
  };
}

function writeAll_(data) {
  writeSheet_('Regions', data.regions || []);
  writeSheet_('Salesmen', data.salesmen || []);
  writeSheet_('Teams', data.teams || []);
  writeSheet_('Entries', data.entries || []);
  writeSheet_('Challenges', data.challenges || []);
  writeSheet_('Monthly', data.monthly || []);
  writeSheet_('Openings', data.openings || []);
  writeSheet_('Slabs', data.slabs || []);
  writeSheet_('OffDays', (data.offDays || []).map(date => ({ date })));
  writeSheet_('Settings', Object.keys(data.settings || {}).map(key => ({ key, value: data.settings[key] })));
}

function addRow_(sheet, row) {
  const sh = ensureHeaders_(sheet, SHEETS[sheet]);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].filter(String);
  const out = Object.assign({ id: Utilities.getUuid() }, row || {});
  sh.appendRow(headers.map(h => out[h] == null ? '' : out[h]));
  return out;
}

function updateRow_(sheet, id, patch) {
  const sh = ensureHeaders_(sheet, SHEETS[sheet]);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const idCol = headers.indexOf('id');
  const rowIndex = values.findIndex((r, i) => i && String(r[idCol]) === String(id));
  if (rowIndex < 1) return addRow_(sheet, Object.assign({ id: id || Utilities.getUuid() }, patch || {}));
  Object.keys(patch || {}).forEach(k => {
    const col = headers.indexOf(k);
    if (col >= 0) sh.getRange(rowIndex + 1, col + 1).setValue(patch[k]);
  });
}

function deleteRow_(sheet, id) {
  const sh = ensureHeaders_(sheet, SHEETS[sheet]);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const idCol = headers.indexOf('id');
  const rowIndex = values.findIndex((r, i) => i && String(r[idCol]) === String(id));
  if (rowIndex > 0) sh.deleteRow(rowIndex + 1);
}

function setMonthly_(row) {
  const existing = readSheet_('Monthly').find(r => String(r.salesmanId) === String(row.salesmanId) && String(r.monthKey) === String(row.monthKey));
  if (existing) updateRow_('Monthly', existing.id, row);
  else addRow_('Monthly', Object.assign({ id: Utilities.getUuid() }, row));
}

function setOpening_(row) {
  const existing = readSheet_('Openings').find(r => String(r.salesmanId) === String(row.salesmanId) && String(r.monthKey) === String(row.monthKey));
  if (existing) updateRow_('Openings', existing.id, row);
  else addRow_('Openings', Object.assign({ id: Utilities.getUuid() }, row));
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
