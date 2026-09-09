const SPREADSHEET_ID = '1kSpF64p6kyRRkEKrd5DrUjmHr1wkyNAkZTG-DlrBHeA';
const SHEETS = {accounts: 'Accounts', units: 'Units', branches: 'Branches'};
const UNIT_STATUSES = ['For observation', 'Released', 'For release', 'To be transfered', 'In warehouse'];
const WAREHOUSE = 'Warehouse';
const BRANCH_SHEETS = {
  'BNB Rosales branch': {tab: 'BNB Rosales', header: '#FFD966', text: '#000000'},
  'BNB Urdaneta branch': {tab: 'BNB Urdaneta', header: '#FFD966', text: '#000000'},
  'BNB Tayug branch': {tab: 'BNB Tayug', header: '#FFD966', text: '#000000'},
  'EZ San Carlos branch': {tab: 'EZ San Carlos', header: '#93C47D', text: '#000000'},
  'EZ San Jose branch': {tab: 'EZ San Jose', header: '#93C47D', text: '#000000'},
  '1LR Tarlac Branch': {tab: '1LR Tarlac', header: '#000000', text: '#FF0000'}
};
const UNIT_HEADERS = ['Unit Code', 'Client Name', 'Model', 'Processor', 'RAM', 'Storage Size (HDD/SSD/SD)', 'Unit Price', 'Status', 'Current Location', 'Date Received', 'Released Date'];
// Temporary bootstrap access. Remove this after creating a permanent technician account.
const TEMP_TECHNICIAN = {name: 'Temporary Technician', username: 'temp.technician', password: 'UnitflowTemp2026!', role: 'technician', branch: 'All branches', status: 'Active'};

function doGet() { ensureSheets(); return json({ ok: true, service: 'unitflow' }); }

function doPost(event) {
  try {
    const request = JSON.parse(event.postData.contents);
    ensureSheets();
    if (request.action === 'login') return json(login(request.username, request.password));
    const user = authenticate(request.token);
    if (!user) return json({ ok: false, error: 'Unauthorized' });
    if (request.action === 'list') return json({ ok: true, data: listUnits(user) });
    if (request.action === 'createUnit') { requireRole(user, ['technician', 'office', 'admin']); return json({ ok: true, data: createUnit(request, user) }); }
    if (request.action === 'updateUnit') { requireRole(user, ['technician', 'office', 'admin']); return json({ ok: true, data: updateUnit(request, user) }); }
    if (request.action === 'deleteUnit') { requireRole(user, ['technician']); return json({ ok: true, data: deleteUnit(request.unitCode) }); }
    if (request.action === 'createAccount') { requireRole(user, ['technician']); return json({ ok: true, data: createAccount(request) }); }
    if (request.action === 'deleteAccount') { requireRole(user, ['technician']); return json({ ok: true, data: deleteAccount(request.username, user) }); }
    if (request.action === 'createBranch') { requireRole(user, ['technician', 'office']); return json({ ok: true, data: createBranch(request.name) }); }
    if (request.action === 'deleteBranch') { requireRole(user, ['technician']); return json({ ok: true, data: deleteBranch(request.name) }); }
    return json({ ok: false, error: 'Unknown action' });
  } catch (error) { return json({ ok: false, error: error.message }); }
}

function login(username, password) {
  ensureSheets();
  const rows = sheet(SHEETS.accounts).getDataRange().getValues();
  const headers = rows.shift();
  const submittedUsername = String(username || '').trim().toLowerCase();
  const submittedPassword = String(password || '');
  const account = rows.map(row => objectFrom(headers, row)).find(item => String(item.username || '').trim().toLowerCase() === submittedUsername && String(item.password || '') === submittedPassword && String(item.status || 'Active').trim().toLowerCase() !== 'disabled');
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
  syncBranchSheets();
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
  syncBranchSheets();
  return request;
}
function createAccount(request) { sheet(SHEETS.accounts).appendRow([request.name, request.username, request.password, request.role, request.branch || 'All branches', 'Active']); return request.username; }
function deleteAccount(username, user) {
  if (String(username).trim().toLowerCase() === String(user.username).trim().toLowerCase()) throw new Error('You cannot delete your current account');
  const tab = sheet(SHEETS.accounts);
  const values = tab.getDataRange().getValues();
  const rowIndex = values.findIndex((row, index) => index > 0 && String(row[1]).trim().toLowerCase() === String(username).trim().toLowerCase());
  if (rowIndex < 1) throw new Error('Account not found');
  tab.deleteRow(rowIndex + 1);
  return username;
}
function createBranch(name) {
  const branchName = String(name || '').trim();
  if (!branchName) throw new Error('Branch name is required');
  const branchesTab = sheet(SHEETS.branches);
  const existing = branchesTab.getDataRange().getValues().slice(1).some(row => String(row[0]).trim().toLowerCase() === branchName.toLowerCase());
  if (existing) throw new Error('Branch already exists');
  branchesTab.appendRow([branchName, new Date()]);
  createBranchSheet(SpreadsheetApp.openById(SPREADSHEET_ID), branchName);
  syncBranchSheets();
  return branchName;
}
function deleteUnit(unitCode) {
  const tab = sheet(SHEETS.units);
  const values = tab.getDataRange().getValues();
  const rowIndex = values.findIndex((row, index) => index > 0 && String(row[0]) === String(unitCode));
  if (rowIndex < 1) throw new Error('Unit not found');
  tab.deleteRow(rowIndex + 1);
  syncBranchSheets();
  return unitCode;
}
function deleteBranch(name) {
  const tab = sheet(SHEETS.branches);
  const values = tab.getDataRange().getValues();
  const branchName = String(name || '').trim();
  const rowIndex = values.findIndex((row, index) => index > 0 && String(row[0]).trim().toLowerCase() === branchName.toLowerCase());
  if (rowIndex < 1) throw new Error('Branch not found');
  tab.deleteRow(rowIndex + 1);
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const branchTab = spreadsheet.getSheetByName((BRANCH_SHEETS[branchName] || branchSettings(branchName)).tab);
  if (branchTab) spreadsheet.deleteSheet(branchTab);
  return branchName;
}
function adminLocations(user) { return [...new Set([user.branch, 'BNB Rosales branch', WAREHOUSE])]; }
function authenticate(token) { const raw = CacheService.getScriptCache().get(token); return raw ? JSON.parse(raw) : null; }
function requireRole(user, roles) { if (roles.indexOf(user.role) === -1) throw new Error('Insufficient permissions'); }
function ensureSheets() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const definitions = {
    Accounts: ['name', 'username', 'password', 'role', 'branch', 'status'],
    Units: UNIT_HEADERS,
    Branches: ['name', 'created']
  };
  Object.keys(definitions).forEach(name => {
    let tab = spreadsheet.getSheetByName(name);
    if (!tab) tab = spreadsheet.insertSheet(name);
    if (tab.getLastRow() === 0) tab.getRange(1, 1, 1, definitions[name].length).setValues([definitions[name]]);
  });
  ensureDefaultBranches(spreadsheet.getSheetByName(SHEETS.branches));
  ensureTemporaryAccount(spreadsheet.getSheetByName(SHEETS.accounts));
  setupBranchSheets(spreadsheet);
}
function ensureDefaultBranches(tab) {
  if (tab.getLastRow() > 1) return;
  Object.keys(BRANCH_SHEETS).forEach(branch => tab.appendRow([branch, new Date()]));
}
function ensureTemporaryAccount(tab) {
  const rows = tab.getDataRange().getValues();
  const hasTemporaryAccount = rows.slice(1).some(row => String(row[1] || '').trim().toLowerCase() === TEMP_TECHNICIAN.username);
  if (!hasTemporaryAccount) tab.appendRow([TEMP_TECHNICIAN.name, TEMP_TECHNICIAN.username, TEMP_TECHNICIAN.password, TEMP_TECHNICIAN.role, TEMP_TECHNICIAN.branch, TEMP_TECHNICIAN.status]);
}
function setupSpreadsheet() { ensureSheets(); syncBranchSheets(); }
function setupBranchSheets(spreadsheet) {
  const branches = sheet(SHEETS.branches).getDataRange().getValues().slice(1).map(row => String(row[0]).trim()).filter(Boolean);
  branches.forEach(branch => createBranchSheet(spreadsheet, branch));
}
function createBranchSheet(spreadsheet, branch) {
  const settings = BRANCH_SHEETS[branch] || branchSettings(branch);
  let tab = spreadsheet.getSheetByName(settings.tab);
  if (!tab) tab = spreadsheet.insertSheet(settings.tab);
  tab.getRange(1, 1, 1, UNIT_HEADERS.length).setValues([UNIT_HEADERS]);
  tab.setFrozenRows(1);
  tab.getRange(1, 1, 1, UNIT_HEADERS.length).setBackground(settings.header).setFontColor(settings.text).setFontWeight('bold');
  tab.getRange(1, 1, 1, UNIT_HEADERS.length).setHorizontalAlignment('center');
  tab.autoResizeColumns(1, UNIT_HEADERS.length);
}
function branchSettings(branch) {
  const name = String(branch).trim();
  if (name.toUpperCase().startsWith('BNB')) return {tab: name.replace(/\s+branch$/i, ''), header: '#FFD966', text: '#000000'};
  if (name.toUpperCase().startsWith('EZ')) return {tab: name.replace(/\s+branch$/i, ''), header: '#93C47D', text: '#000000'};
  if (name.toUpperCase().startsWith('1LR')) return {tab: name.replace(/\s+branch$/i, ''), header: '#000000', text: '#FF0000'};
  return {tab: name.substring(0, 90), header: '#D9EAD3', text: '#000000'};
}
function syncBranchSheets() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  setupBranchSheets(spreadsheet);
  const master = sheet(SHEETS.units);
  const values = master.getDataRange().getValues();
  values.shift();
  const branchRows = sheet(SHEETS.branches).getDataRange().getValues().slice(1);
  const branches = [...new Set([...Object.keys(BRANCH_SHEETS), ...branchRows.map(row => String(row[0]).trim()).filter(Boolean)])];
  branches.forEach(branch => {
    createBranchSheet(spreadsheet, branch);
    const settings = BRANCH_SHEETS[branch] || branchSettings(branch);
    const tab = spreadsheet.getSheetByName(settings.tab);
    const rows = values.filter(row => String(row[8]) === branch);
    const rowCount = Math.max(tab.getLastRow() - 1, 0);
    if (rowCount) tab.getRange(2, 1, rowCount, UNIT_HEADERS.length).clearContent();
    if (rows.length) tab.getRange(2, 1, rows.length, UNIT_HEADERS.length).setValues(rows);
    tab.getRange(1, 1, 1, UNIT_HEADERS.length).setBackground(settings.header).setFontColor(settings.text).setFontWeight('bold');
  });
}
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