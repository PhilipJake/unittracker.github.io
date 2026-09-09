const SPREADSHEET_ID = '1kSpF64p6kyRRkEKrd5DrUjmHr1wkyNAkZTG-DlrBHeA';
const SHEETS = {accounts: 'Accounts', units: 'Units', branches: 'Branches'};
const UNIT_STATUSES = ['For observation', 'Released', 'For release', 'To be transfered', 'In warehouse'];
const WAREHOUSE = 'Warehouse';
// Temporary bootstrap access. Remove this after creating a permanent technician account.
const TEMP_TECHNICIAN = {name: 'Temporary Technician', username: 'temp.technician', password: 'UnitflowTemp2026!', role: 'technician', branch: 'All branches', status: 'Active'};

function doGet() { return json({ ok: true, service: 'unitflow' }); }

function doPost(event) {
  try {
    const request = JSON.parse(event.postData.contents);
    if (request.action === 'login') return json(login(request.username, request.password));
    const user = authenticate(request.token);
    if (!user) return json({ ok: false, error: 'Unauthorized' });
    if (request.action === 'list') return json({ ok: true, data: listUnits(user) });
    if (request.action === 'createUnit') { requireRole(user, ['technician', 'office', 'admin']); return json({ ok: true, data: createUnit(request, user) }); }
    if (request.action === 'updateUnit') { requireRole(user, ['technician', 'office', 'admin']); return json({ ok: true, data: updateUnit(request, user) }); }
    if (request.action === 'deleteUnit') { requireRole(user, ['technician']); return json({ ok: true, data: deleteUnit(request.unitCode) }); }
    if (request.action === 'createAccount') { requireRole(user, ['technician']); return json({ ok: true, data: createAccount(request) }); }
    if (request.action === 'createBranch') { requireRole(user, ['technician', 'office']); return json({ ok: true, data: createBranch(request.name) }); }
    if (request.action === 'deleteBranch') { requireRole(user, ['technician']); return json({ ok: true, data: deleteBranch(request.name) }); }
    return json({ ok: false, error: 'Unknown action' });
  } catch (error) { return json({ ok: false, error: error.message }); }
}

function login(username, password) {
  const rows = sheet(SHEETS.accounts).getDataRange().getValues();
  const headers = rows.shift();
  const account = rows.map(row => objectFrom(headers, row)).find(item => item.username === username && item.password === password && item.status !== 'Disabled') || (username === TEMP_TECHNICIAN.username && password === TEMP_TECHNICIAN.password ? {...TEMP_TECHNICIAN} : null);
  if (!account) return { ok: false, error: 'Invalid username or password' };
  const token = Utilities.getUuid(); CacheService.getScriptCache().put(token, JSON.stringify(account), 21600);
  delete account.password; return { ok: true, user: account, token: token };
}

function listUnits(user) { const values = sheet(SHEETS.units).getDataRange().getValues(); const headers = values.shift(); return values.map(row => objectFrom(headers, row)).filter(unit => user.role !== 'admin' || adminLocations(user).indexOf(unit.currentLocation) !== -1); }
function createUnit(request, user) {
  if (UNIT_STATUSES.indexOf(request.status) === -1) throw new Error('Invalid unit status');
  if (user.role === 'admin' && adminLocations(user).indexOf(request.currentLocation) === -1) throw new Error('Admins can only add units to their assigned branch, BNB Rosales branch, or Warehouse');
  const tab = sheet(SHEETS.units);
  tab.appendRow([request.unitCode, request.clientName, request.model, request.processor, request.ram, request.storage, request.unitPrice, request.status, request.currentLocation, request.dateReceived, request.releasedDate]);
  return request;
}
function updateUnit(request, user) {
  if (UNIT_STATUSES.indexOf(request.status) === -1) throw new Error('Invalid unit status');
  const tab = sheet(SHEETS.units);
  const values = tab.getDataRange().getValues();
  const headers = values[0];
  const codeColumn = columnIndex(headers, 'Unit Code');
  const locationColumn = columnIndex(headers, 'Current Location');
  const rowIndex = values.findIndex((row, index) => index > 0 && String(row[codeColumn]) === String(request.originalUnitCode));
  if (rowIndex < 1) throw new Error('Unit not found');
  if (user.role === 'admin' && (adminLocations(user).indexOf(String(values[rowIndex][locationColumn])) === -1 || adminLocations(user).indexOf(request.currentLocation) === -1)) throw new Error('Admins can only edit units in their assigned branch, BNB Rosales branch, or Warehouse');
  const updated = [request.unitCode, request.clientName, request.model, request.processor, request.ram, request.storage, request.unitPrice, request.status, request.currentLocation, request.dateReceived, request.releasedDate];
  tab.getRange(rowIndex + 1, 1, 1, updated.length).setValues([updated]);
  return request;
}
function createAccount(request) { sheet(SHEETS.accounts).appendRow([request.name, request.username, request.password, request.role, request.branch || 'All branches', 'Active']); return request.username; }
function createBranch(name) { sheet(SHEETS.branches).appendRow([name, new Date()]); return name; }
function deleteUnit(unitCode) {
  const tab = sheet(SHEETS.units);
  const values = tab.getDataRange().getValues();
  const rowIndex = values.findIndex((row, index) => index > 0 && String(row[0]) === String(unitCode));
  if (rowIndex < 1) throw new Error('Unit not found');
  tab.deleteRow(rowIndex + 1);
  return unitCode;
}
function deleteBranch(name) {
  const tab = sheet(SHEETS.branches);
  const values = tab.getDataRange().getValues();
  const rowIndex = values.findIndex((row, index) => index > 0 && String(row[0]) === String(name));
  if (rowIndex < 1) throw new Error('Branch not found');
  tab.deleteRow(rowIndex + 1);
  return name;
}
function adminLocations(user) { return [...new Set([user.branch, 'BNB Rosales branch', WAREHOUSE])]; }
function authenticate(token) { const raw = CacheService.getScriptCache().get(token); return raw ? JSON.parse(raw) : null; }
function requireRole(user, roles) { if (roles.indexOf(user.role) === -1) throw new Error('Insufficient permissions'); }
function sheet(name) { return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(name); }
function objectFrom(headers, row) {
  return headers.reduce((object, header, index) => {
    const key = String(header).toLowerCase().trim();
    const aliases = {'unit code':'unitCode', 'client name':'clientName', 'current location':'currentLocation', 'storage size':'storage', 'storage size (hdd/ssd/sd)':'storage', 'unit price':'unitPrice', 'date received':'dateReceived', 'released date':'releasedDate'};
    object[aliases[key] || key] = row[index];
    return object;
  }, {});
}
function columnIndex(headers, name) { const target = name.toLowerCase().trim(); const index = headers.findIndex(header => String(header).toLowerCase().trim() === target); if (index < 0) throw new Error('Missing sheet column: ' + name); return index; }
function json(data) { return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON); }