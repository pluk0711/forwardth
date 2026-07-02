// ════════════════════════════════════════════════════════════
// ForwardTH — Google Apps Script Backend  v4
// Deploy: Extensions → Apps Script → Deploy → New deployment → Web App
//   Execute as: Me  |  Who has access: Anyone
// Sheets: Customers, Shops, Lots, Shipments, Consolidation, Config
// ════════════════════════════════════════════════════════════

const SS = SpreadsheetApp.getActiveSpreadsheet();

function jsonOut(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════
// GET
// ════════════════════════════════════
function doGet(e) {
  try {
    const type = (e.parameter && e.parameter.type) || 'all';
    let result;
    if      (type === 'customers') result = { customers: getCustomers() };
    else if (type === 'shops')     result = { shops: getShops() };
    else if (type === 'lots')      result = { lots: getLots() };
    else if (type === 'ships')     result = { ships: getShips() };
    else if (type === 'config')    result = { config: getConfig() };
    else if (type === 'consolidations') result = { consolidations: getConsolidations() };
    else result = { customers: getCustomers(), shops: getShops(), lots: getLots(), ships: getShips(), consolidations: getConsolidations(), bookings: getBookings(), config: getConfig() };
    return jsonOut(result);
  } catch(err) {
    return jsonOut({ error: err.message });
  }
}

// ════════════════════════════════════
// POST
// ════════════════════════════════════
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action;
    let result   = { ok: true };

    if      (action === 'addCustomer')    addCustomer(body.data);
    else if (action === 'updateCustomer') updateCustomer(body.id, body.updates);
    else if (action === 'deleteCustomer') deleteCustomer(body.id);
    else if (action === 'addShop')        addShop(body.data);
    else if (action === 'updateShop')     updateShop(body.id, body.updates);
    else if (action === 'deleteShop')     deleteShop(body.id);
    else if (action === 'addLot')         addLot(body.data);
    else if (action === 'updateLot')      updateLot(body.id, body.updates);
    else if (action === 'addShip')        addShip(body.data);
    else if (action === 'updateShip')     updateShip(body.id, body.updates);
    else if (action === 'addEvent')       addEvent(body.id, body.event);
    else if (action === 'setConfig')         setConfigKey(body.key, body.value);
    else if (action === 'addConsolidation')    addConsolidation(body.data);
    else if (action === 'updateConsolidation') updateConsolidation(body.id, body.updates);
    else if (action === 'deleteConsolidation') { deleteConsolidation(body.id); deleteShip(body.id); }
    else if (action === 'addBooking')          addBooking(body.data);
    else if (action === 'updateBooking')       updateBooking(body.id, body.updates);
    else if (action === 'seed')           runSeed(body.customers, body.shops, body.lots, body.ships, body.config);
    else if (action === 'uploadImage')    result = { ok: true, url: uploadPassportImage(body.customerId, body.filename, body.base64, body.mimeType) };
    else if (action === 'uploadLotPhoto')  result = { ok: true, url: uploadLotPhoto(body.lotId, body.slot, body.base64, body.mimeType) };
    else if (action === 'deleteLotPhoto')  { deleteLotPhoto(body.lotId, body.slot); }
    else result = { ok: false, error: 'Unknown action: ' + action };

    return jsonOut(result);
  } catch(err) {
    return jsonOut({ ok: false, error: err.message });
  } finally {
    lock.releaseLock();
  }
}

// ════════════════════════════════════
// CUSTOMERS
// ════════════════════════════════════
const CUST_COLS = ['id','name','companyName','passportNumber','passportImageUrl','email','countryCode','phone','imType','imAccount','country','address','responsibleSale','note','aliases'];

function getCustomers() {
  const sh = getSheet('Customers');
  const n  = sh.getLastRow();
  if (n <= 1) return [];
  return sh.getRange(2, 1, n - 1, CUST_COLS.length).getValues().map(rowToCust);
}

function rowToCust(r) {
  const o = {};
  CUST_COLS.forEach((k, i) => o[k] = r[i]);
  o.createdAt = Number(o.createdAt) || 0;
  return o;
}

function custToRow(c) {
  return CUST_COLS.map(k => c[k] !== undefined ? c[k] : '');
}

function addCustomer(cust) {
  const sh = getSheet('Customers');
  ensureHeader(sh, CUST_COLS);
  sh.appendRow(custToRow(cust));
}

function updateCustomer(id, updates) {
  const sh   = getSheet('Customers');
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      Object.keys(updates).forEach(k => {
        const col = CUST_COLS.indexOf(k);
        if (col >= 0) sh.getRange(i + 1, col + 1).setValue(updates[k] ?? '');
      });
      return;
    }
  }
}

function deleteCustomer(id) {
  const sh   = getSheet('Customers');
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) { sh.deleteRow(i + 1); return; }
  }
}

// ════════════════════════════════════
// SHOPS
// ════════════════════════════════════
const SHOP_COLS = ['id','name','market','address','email','phone','contactType','contactValue','imAccount','productTypes','note','active'];

function getShops() {
  const sh = getSheet('Shops');
  const n  = sh.getLastRow();
  if (n <= 1) return [];
  return sh.getRange(2, 1, n - 1, SHOP_COLS.length).getValues().map(rowToShopItem);
}

function rowToShopItem(r) {
  const o = {};
  SHOP_COLS.forEach((k, i) => o[k] = r[i]);
  o.active = o.active === true || o.active === 'TRUE';
  return o;
}

function shopToRow(s) {
  return SHOP_COLS.map(k => s[k] !== undefined ? s[k] : '');
}

function addShop(shop) {
  const sh = getSheet('Shops');
  ensureHeader(sh, SHOP_COLS);
  sh.appendRow(shopToRow(shop));
}

function updateShop(id, updates) {
  const sh   = getSheet('Shops');
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      Object.keys(updates).forEach(k => {
        const col = SHOP_COLS.indexOf(k);
        if (col >= 0) {
          const val = k === 'customerIds' ? JSON.stringify(updates[k]) : (updates[k] ?? '');
          sh.getRange(i + 1, col + 1).setValue(val);
        }
      });
      return;
    }
  }
}

function deleteShop(id) {
  const sh   = getSheet('Shops');
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) { sh.deleteRow(i + 1); return; }
  }
}

// ════════════════════════════════════
// LOTS
// ════════════════════════════════════
const LOT_COLS = [
  'id','customerId','customerName','shopName','date','time',
  'pkgType','pkgUnit','pkgCount','actualWeight','volWeight',
  'rack','dn','note','photoCount','status','photoUrls'
];

function getLots() {
  const sh = getSheet('Lots');
  const n  = sh.getLastRow();
  if (n <= 1) return [];
  return sh.getRange(2, 1, n - 1, LOT_COLS.length).getValues().map(rowToLot);
}

function rowToLot(r) {
  const o = {};
  LOT_COLS.forEach((k, i) => o[k] = r[i]);
  o.pkgCount     = Number(o.pkgCount)     || 0;
  o.actualWeight = Number(o.actualWeight) || 0;
  o.volWeight    = (o.volWeight === '' || o.volWeight === null) ? null : Number(o.volWeight);
  o.photoCount   = Number(o.photoCount)   || 0;
  o.customerName = String(o.customerName || '');
  try { o.photoUrls = JSON.parse(o.photoUrls || '[]'); } catch(_) { o.photoUrls = []; }
  // Convert date: Date object → YYYY-MM-DD string
  if (o.date instanceof Date) {
    const tz = Session.getScriptTimeZone();
    o.date = Utilities.formatDate(o.date, tz, 'yyyy-MM-dd');
  } else {
    o.date = String(o.date || '');
  }
  // Convert time: Date object (epoch 1899-12-30) → HH:mm string
  if (o.time instanceof Date) {
    const h = String(o.time.getHours()).padStart(2,'0');
    const m = String(o.time.getMinutes()).padStart(2,'0');
    o.time = h + ':' + m;
  } else {
    o.time = String(o.time || '');
  }
  return o;
}

function lotToRow(l) {
  return LOT_COLS.map(k => {
    if (k === 'volWeight') return (l[k] === null || l[k] === undefined) ? '' : l[k];
    if (k === 'photoUrls') return JSON.stringify(Array.isArray(l[k]) ? l[k] : []);
    return l[k] !== undefined ? l[k] : '';
  });
}

function addLot(lot) {
  const sh = getSheet('Lots');
  ensureHeader(sh, LOT_COLS);
  sh.appendRow(lotToRow(lot));
}

function updateLot(id, updates) {
  const sh   = getSheet('Lots');
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      Object.keys(updates).forEach(k => {
        const col = LOT_COLS.indexOf(k);
        if (col >= 0) sh.getRange(i + 1, col + 1).setValue(updates[k] === null ? '' : updates[k]);
      });
      return;
    }
  }
}

// ════════════════════════════════════
// SHIPMENTS
// ════════════════════════════════════
const SHIP_COLS = [
  'id','customerId','lotIds','destination','destinationLabel','carrier','sendDate','note',
  'totalActualKg','totalVolKg','totalChargeableKg','totalPackages',
  'status','awb','dispatchDate','tracking','createdAt'
];

function getShips() {
  const sh = getSheet('Shipments');
  const n  = sh.getLastRow();
  if (n <= 1) return [];
  return sh.getRange(2, 1, n - 1, SHIP_COLS.length).getValues().map(rowToShip);
}

function rowToShip(r) {
  const o = {};
  SHIP_COLS.forEach((k, i) => o[k] = r[i]);
  try { o.lotIds   = JSON.parse(o.lotIds   || '[]'); } catch(_) { o.lotIds = []; }
  try { o.tracking = JSON.parse(o.tracking || '[]'); } catch(_) { o.tracking = []; }
  o.totalActualKg     = Number(o.totalActualKg)     || 0;
  o.totalVolKg        = Number(o.totalVolKg)        || 0;
  o.totalChargeableKg = Number(o.totalChargeableKg) || 0;
  o.totalPackages     = Number(o.totalPackages)     || 0;
  o.createdAt         = Number(o.createdAt)         || 0;
  return o;
}

function shipToRow(s) {
  return SHIP_COLS.map(k => {
    if (k === 'lotIds' || k === 'tracking') return JSON.stringify(s[k] || []);
    return s[k] !== undefined ? s[k] : '';
  });
}

function addShip(ship) {
  const sh = getSheet('Shipments');
  ensureHeader(sh, SHIP_COLS);
  sh.appendRow(shipToRow(ship));
}

function updateShip(id, updates) {
  const sh   = getSheet('Shipments');
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      Object.keys(updates).forEach(k => {
        const col = SHIP_COLS.indexOf(k);
        if (col >= 0) {
          const val = (k === 'lotIds' || k === 'tracking')
            ? JSON.stringify(updates[k])
            : (updates[k] === null ? '' : updates[k]);
          sh.getRange(i + 1, col + 1).setValue(val);
        }
      });
      return;
    }
  }
}

function addEvent(shipId, evt) {
  const ships = getShips();
  const ship  = ships.find(s => s.id === shipId);
  if (!ship) return;
  ship.tracking.push(evt);
  updateShip(shipId, { tracking: ship.tracking });
}

// ════════════════════════════════════
// CONSOLIDATION
// ════════════════════════════════════
const CONSOL_COLS = [
  'id','customerId','customerName','lotIds','destination','note',
  'totalActualKg','totalVolKg','totalChargeableKg','totalPackages',
  'status','cbmRows','cbmBoxes','cbmTotalWeight','cbmTotalCbm','cbmChargeable','cbmTransportType',
  'createdAt'
];

function getConsolidations() {
  const sh = getSheet('Consolidation');
  const n  = sh.getLastRow();
  if (n <= 1) return [];
  return sh.getRange(2, 1, n - 1, CONSOL_COLS.length).getValues().map(rowToConsol);
}

function rowToConsol(r) {
  const o = {};
  CONSOL_COLS.forEach((k, i) => o[k] = r[i]);
  try { o.lotIds  = JSON.parse(o.lotIds  || '[]'); } catch(_) { o.lotIds = []; }
  try { o.cbmRows = JSON.parse(o.cbmRows || '[]'); } catch(_) { o.cbmRows = []; }
  o.totalActualKg     = Number(o.totalActualKg)     || 0;
  o.totalVolKg        = Number(o.totalVolKg)        || 0;
  o.totalChargeableKg = Number(o.totalChargeableKg) || 0;
  o.totalPackages     = Number(o.totalPackages)     || 0;
  o.cbmBoxes          = Number(o.cbmBoxes)          || 0;
  o.cbmTotalWeight    = Number(o.cbmTotalWeight)    || 0;
  o.cbmTotalCbm       = Number(o.cbmTotalCbm)       || 0;
  o.cbmChargeable     = Number(o.cbmChargeable)     || 0;
  o.createdAt         = Number(o.createdAt)         || 0;
  return o;
}

function consolToRow(c) {
  return CONSOL_COLS.map(k => {
    if (k === 'lotIds')  return JSON.stringify(c[k] || []);
    if (k === 'cbmRows') return JSON.stringify(c[k] || []);
    return c[k] !== undefined ? c[k] : '';
  });
}

function addConsolidation(consol) {
  const sh = getSheet('Consolidation');
  ensureHeader(sh, CONSOL_COLS);
  // Check if already exists (avoid duplicate on retry)
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(consol.id)) {
      // Update instead
      updateConsolidation(consol.id, consol);
      return;
    }
  }
  sh.appendRow(consolToRow(consol));
}

function updateConsolidation(id, updates) {
  const sh   = getSheet('Consolidation');
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      Object.keys(updates).forEach(k => {
        const col = CONSOL_COLS.indexOf(k);
        if (col >= 0) {
          const val = (k === 'lotIds' || k === 'cbmRows')
            ? JSON.stringify(updates[k])
            : (updates[k] === null ? '' : updates[k]);
          sh.getRange(i + 1, col + 1).setValue(val);
        }
      });
      return;
    }
  }
  // Not found — insert new
  ensureHeader(sh, CONSOL_COLS);
  sh.appendRow(consolToRow({...updates, id}));
}

function deleteConsolidation(id) {
  const sh   = getSheet('Consolidation');
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) { sh.deleteRow(i + 1); return; }
  }
}

function deleteShip(id) {
  const sh   = getSheet('Shipments');
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) { sh.deleteRow(i + 1); return; }
  }
}

// ════════════════════════════════════
// CONFIG
// ════════════════════════════════════
function getConfig() {
  const sh   = getSheet('Config');
  const data = sh.getDataRange().getValues();
  const cfg  = { lot_ctr: 1, co_ctr: 1, cust_ctr: 1, bk_ctr: 1, seeded: false };
  data.forEach(r => {
    if (r[0] === 'lot_ctr')  cfg.lot_ctr  = Number(r[1]) || 1;
    if (r[0] === 'co_ctr')   cfg.co_ctr   = Number(r[1]) || 1;
    if (r[0] === 'cust_ctr') cfg.cust_ctr = Number(r[1]) || 1;
    if (r[0] === 'bk_ctr')  cfg.bk_ctr   = Number(r[1]) || 1;
    if (r[0] === 'seeded')   cfg.seeded   = r[1] === true || r[1] === 'TRUE' || r[1] === 'true';
  });
  return cfg;
}

function setConfigKey(key, value) {
  const sh   = getSheet('Config');
  const data = sh.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === key) { sh.getRange(i + 1, 2).setValue(value); return; }
  }
  sh.appendRow([key, value]);
}

// ════════════════════════════════════
// SEED
// ════════════════════════════════════
function runSeed(customers, shops, lots, ships, config) {
  const cfg = getConfig();
  if (cfg.seeded) return;

  const csh  = getSheet('Customers');
  const spsh = getSheet('Shops');
  const lsh  = getSheet('Lots');
  const ssh  = getSheet('Shipments');

  if (csh.getLastRow()  > 1) csh.getRange(2,1,csh.getLastRow()-1,CUST_COLS.length).clearContent();
  if (spsh.getLastRow() > 1) spsh.getRange(2,1,spsh.getLastRow()-1,SHOP_COLS.length).clearContent();
  if (lsh.getLastRow()  > 1) lsh.getRange(2,1,lsh.getLastRow()-1,LOT_COLS.length).clearContent();
  if (ssh.getLastRow()  > 1) ssh.getRange(2,1,ssh.getLastRow()-1,SHIP_COLS.length).clearContent();

  ensureHeader(csh,  CUST_COLS);
  ensureHeader(spsh, SHOP_COLS);
  ensureHeader(lsh,  LOT_COLS);
  ensureHeader(ssh,  SHIP_COLS);

  if (customers && customers.length) customers.forEach(c => csh.appendRow(custToRow(c)));
  if (shops     && shops.length)     shops.forEach(s     => spsh.appendRow(shopToRow(s)));
  if (lots      && lots.length)      lots.forEach(l      => lsh.appendRow(lotToRow(l)));
  if (ships     && ships.length)     ships.forEach(s     => ssh.appendRow(shipToRow(s)));

  if (config) {
    setConfigKey('lot_ctr',  config.lot_ctr  || 14);
    setConfigKey('co_ctr',   config.co_ctr   || 2);
    setConfigKey('cust_ctr', config.cust_ctr || 5);
    setConfigKey('shop_ctr', config.shop_ctr || 4);
  }
  setConfigKey('seeded', true);
}

// ════════════════════════════════════
// PASSPORT IMAGE UPLOAD
// ════════════════════════════════════
function uploadPassportImage(customerId, filename, base64, mimeType) {
  const folderName = 'ForwardTH-Passports';
  const folders = DriveApp.getFoldersByName(folderName);
  const folder  = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

  // Remove old passport files for this customer (keep it clean)
  const existing = folder.getFilesByName(customerId + '_passport');
  while (existing.hasNext()) existing.next().setTrashed(true);

  const decoded  = Utilities.base64Decode(base64);
  const blob     = Utilities.newBlob(decoded, mimeType || 'image/jpeg', customerId + '_passport');
  const file     = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // Return direct-view URL (works as <img src>)
  return 'https://drive.google.com/uc?export=view&id=' + file.getId();
}

// ════════════════════════════════════
// LOT PHOTO UPLOAD
// ════════════════════════════════════
function uploadLotPhoto(lotId, slot, base64, mimeType) {
  const folderName = 'ForwardTH-LotPhotos';
  const folders = DriveApp.getFoldersByName(folderName);
  const folder  = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

  const filename = lotId + '_slot' + slot;
  // Remove old file for this slot if exists
  const existing = folder.getFilesByName(filename);
  while (existing.hasNext()) existing.next().setTrashed(true);

  const decoded = Utilities.base64Decode(base64);
  const blob    = Utilities.newBlob(decoded, mimeType || 'image/jpeg', filename);
  const file    = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // Update photoUrls in Lots sheet
  const url = 'https://drive.google.com/uc?export=view&id=' + file.getId();
  const sh   = getSheet('Lots');
  const data = sh.getDataRange().getValues();
  const urlCol = LOT_COLS.indexOf('photoUrls');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(lotId)) {
      let urls = [];
      try { urls = JSON.parse(data[i][urlCol] || '[]'); } catch(_) {}
      while (urls.length < slot) urls.push('');
      urls[slot - 1] = url;
      sh.getRange(i + 1, urlCol + 1).setValue(JSON.stringify(urls));
      // Update photoCount too
      const cntCol = LOT_COLS.indexOf('photoCount');
      sh.getRange(i + 1, cntCol + 1).setValue(urls.filter(u => u).length);
      break;
    }
  }
  return url;
}

function deleteLotPhoto(lotId, slot) {
  const sh   = getSheet('Lots');
  const data = sh.getDataRange().getValues();
  const urlCol = LOT_COLS.indexOf('photoUrls');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(lotId)) {
      let urls = [];
      try { urls = JSON.parse(data[i][urlCol] || '[]'); } catch(_) {}
      if (urls[slot - 1]) urls[slot - 1] = '';
      sh.getRange(i + 1, urlCol + 1).setValue(JSON.stringify(urls));
      const cntCol = LOT_COLS.indexOf('photoCount');
      sh.getRange(i + 1, cntCol + 1).setValue(urls.filter(u => u).length);
      break;
    }
  }
}

// ════════════════════════════════════
// HELPERS
// ════════════════════════════════════
function getSheet(name) {
  return SS.getSheetByName(name) || SS.insertSheet(name);
}

function ensureHeader(sh, cols) {
  if (sh.getLastRow() === 0) {
    sh.appendRow(cols);
  } else {
    // Always overwrite row 1 to keep headers in sync with LOT_COLS / SHIP_COLS etc.
    sh.getRange(1, 1, 1, cols.length).setValues([cols]);
    // Clear any extra columns from old schema
    const totalCols = sh.getLastColumn();
    if (totalCols > cols.length) {
      sh.getRange(1, cols.length + 1, 1, totalCols - cols.length).clearContent();
    }
  }
  sh.getRange(1, 1, 1, cols.length)
    .setBackground('#1A9E6F')
    .setFontColor('white')
    .setFontWeight('bold');
  sh.setFrozenRows(1);
}

// Fix all sheet headers — run once from Apps Script editor after schema changes
// ════════════════════════════════════
// BOOKINGS
// ════════════════════════════════════
const BK_COLS = [
  'id','consolId','status',
  // Schedule (requested)
  'carrier','containerType','carrierName','vesselName','awb','pol','pod','etd','eta','cutoff',
  // Shipper / Consignee
  'shipperName','shipperAddr','consigneeName','consigneeAddr','notifyParty','incoterms',
  // Confirmation — docs
  'bookingRef','blNo','confirmedDate','containerNo','sealNo','containerSize',
  // Confirmation — load point
  'warehouse','gate','pot','cargoReceiving','siCutoff',
  // Confirmation — verified schedule
  'confirmedCarrier','confirmedVessel','confirmedEtd','confirmedEta','confirmedPol','confirmedPod',
  // Meta
  'agentNote','note','volDivisor','sentAt','confirmedAt','shippedAt','createdAt'
];

function bookingToRow(b) {
  return BK_COLS.map(k => {
    const v = b[k];
    return (v === undefined || v === null) ? '' : v;
  });
}

function rowToBooking(headers, row) {
  const b = {};
  headers.forEach((h, i) => { b[h] = row[i]; });
  return b;
}

function getBookings() {
  const sh = getSheet('Bookings');
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).filter(r => r[0]).map(r => rowToBooking(headers, r));
}

function addBooking(b) {
  const sh = getSheet('Bookings');
  ensureHeader(sh, BK_COLS);
  sh.appendRow(bookingToRow(b));
}

function updateBooking(id, updates) {
  const sh = getSheet('Bookings');
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('id');
  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === id) {
      Object.keys(updates).forEach(k => {
        const col = headers.indexOf(k);
        if (col >= 0) sh.getRange(i + 1, col + 1).setValue(updates[k]);
      });
      return;
    }
  }
}


function fixAllHeaders() {
  ensureHeader(getSheet('Customers'),     CUST_COLS);
  ensureHeader(getSheet('Shops'),         SHOP_COLS);
  ensureHeader(getSheet('Lots'),          LOT_COLS);
  ensureHeader(getSheet('Shipments'),     SHIP_COLS);
  ensureHeader(getSheet('Consolidation'), CONSOL_COLS);
  ensureHeader(getSheet('Bookings'),     BK_COLS);
  Logger.log('Headers updated on all sheets.');
}
                                                                                                                                                                                                                             