/** Science department inventory — bound to a Google Sheet. */
// Optimized edition: one spreadsheet/tab lookup per request, cached name signatures,
// lock released before the post-save reload, and static catalogs sent only on page load.
const TABLES = {
  Items: ['id','sku','name','category','kind','quantity','checkedOut','unit','location','storage','minStock','manufacturer','catalogNumber','hazard','expiration','notes','updatedAt','needsRepair','needsReplacement'],
  Loans: ['id','itemId','quantity','remaining','borrower','destination','checkedOutAt','dueDate','returnedAt','status','notes'],
  Activity: ['id','timestamp','action','itemId','itemName','quantity','actor','details'],
  Issues: ['id','itemId','type','quantity','remaining','reportedBy','details','reportedAt','status','resolvedBy','resolution','resolvedAt'],
  Locations: ['location'],
  Catalog: ['category','name','kind','unit','courses','notes'],
  PriceEstimates: ['id','name','kind','unit','averageUnitCost','quotesJson','checkedAt','matchType','notes','updatedBy'],
  PriceCatalog: ['name','kind','unit','averageUnitCost','basis','notes','checkedAt'],
  PriceJobs: ['id','key','itemId','name','kind','unit','manufacturer','catalogNumber','status','attempts','updatedAt','nextAttemptAt','candidatesJson','message']
};
const ROOM_OPTIONS = ['409','410','411','412','400','200','201','203','204','209','211'];
const SCHEMA_CACHE_KEY_ = 'inventory-schema-automatic-pricing-v2:';
const SCHEMA_CACHE_SECONDS_ = 21600; // 6 hours. Run runSetup() from the editor after changing tabs or columns.
const DONE_JOB_DAYS_ = 30;           // Finished price jobs older than this are not sent to browsers (they stay in the Sheet).

function doGet(e) {
  if(e && e.parameter && e.parameter.page === 'hub') {
    const allowed = ['home','meetings','updates','faculty','links'];
    const tab = allowed.includes(e.parameter.tab) ? e.parameter.tab : 'home';
    const html = HtmlService.createHtmlOutputFromFile('HubIndex').getContent().replace('window.EAHS_INITIAL_TAB = "home";', 'window.EAHS_INITIAL_TAB = ' + JSON.stringify(tab) + ';');
    return HtmlService.createHtmlOutput(html).setTitle('EAHS Science Department Hub');
  }
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('Science Inventory');
}

let executionBook_;
let executionReady_ = false;
const executionSheets_ = {};
function book_() {
  if (executionBook_) return executionBook_;
  const properties = PropertiesService.getScriptProperties();
  const id = properties.getProperty('INVENTORY_SHEET_ID');
  if (id) return executionBook_ = SpreadsheetApp.openById(id);
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) throw new Error('Open the project from the inventory Sheet and run getData once to connect it.');
  properties.setProperty('INVENTORY_SHEET_ID', active.getId());
  return executionBook_ = active;
}
function setup_() {
  if (executionReady_) return;
  const book = book_();
  Object.keys(TABLES).forEach(name => {
    let sheet = book.getSheetByName(name);
    if (!sheet) sheet = book.insertSheet(name);
    const headers = TABLES[name];
    // Safe migration for the initial app: add the two condition-count columns.
    if (name === 'Items' && sheet.getLastRow() > 0) {
      const prior = headers.slice(0, 17);
      if (sheet.getRange(1, 1, 1, prior.length).getValues()[0].join('|') === prior.join('|') && sheet.getRange(1, 18).getValue() === '') {
        sheet.getRange(1, 18, 1, 2).setValues([['needsRepair','needsReplacement']]);
      }
    }
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#123257').setFontColor('#ffffff');
      sheet.autoResizeColumns(1, headers.length);
    } else if (sheet.getRange(1, 1, 1, headers.length).getValues()[0].join('|') !== headers.join('|')) {
      throw new Error('The ' + name + ' sheet headers have changed. Restore the original column order before continuing.');
    }
    if (sheet.getLastRow() === 1 && name === 'Locations') sheet.getRange(2, 1, ROOM_OPTIONS.length, 1).setValues(ROOM_OPTIONS.map(x => [x]));
    if (sheet.getLastRow() === 1 && name === 'Catalog') sheet.getRange(2, 1, CATALOG_ROWS.length, headers.length).setValues(CATALOG_ROWS);
    if (sheet.getLastRow() === 1 && name === 'PriceCatalog') sheet.getRange(2, 1, CATALOG_PRICE_DEFAULTS.length, headers.length).setValues(CATALOG_PRICE_DEFAULTS);
    if (sheet.getLastRow() === 1 && name === 'PriceEstimates' && PRICE_SEED_ROWS.length) sheet.getRange(2, 1, PRICE_SEED_ROWS.length, headers.length).setValues(PRICE_SEED_ROWS);
  });
  executionReady_ = true;
}

function ensureReadReady_() {
  // Cache only schema readiness; inventory counts and prices are always read fresh.
  let cache;
  const key = SCHEMA_CACHE_KEY_ + book_().getId();
  try { cache = CacheService.getScriptCache(); if (cache.get(key)) return; } catch (e) {}
  setup_();
  try { if (cache) cache.put(key, "ready", SCHEMA_CACHE_SECONDS_); } catch (e) {}
}
// Run from the Apps Script editor after adding a tab or column, or after editing header rows.
function runSetup() {
  setup_();
  try { CacheService.getScriptCache().put(SCHEMA_CACHE_KEY_ + book_().getId(), 'ready', SCHEMA_CACHE_SECONDS_); } catch (e) {}
  return 'Inventory tabs and headers verified.';
}

function sheet_(name) { return executionSheets_[name] || (executionSheets_[name] = book_().getSheetByName(name)); }
function rows_(name, limit) {
  const sh = sheet_(name);
  if (!sh) throw new Error("Missing " + name + " tab. Run runSetup from the Apps Script editor to initialize the inventory.");
  const last = sh.getLastRow(), start = limit ? Math.max(2, last - limit + 1) : 2, count = last - start + 1;
  if (count < 1) return [];
  const keys = TABLES[name];
  const roomField = name === 'Items' ? 'location' : name === 'Loans' ? 'destination' : name === 'Locations' ? 'location' : '';
  let tz = null;
  return sh.getRange(start, 1, count, keys.length).getValues().map((values, index) => {
    const row = { _row: index + start };
    for (let i = 0; i < keys.length; i++) {
      const v = values[i];
      row[keys[i]] = v instanceof Date ? Utilities.formatDate(v, tz || (tz = Session.getScriptTimeZone()), 'yyyy-MM-dd') : v;
    }
    if (roomField) row[roomField] = clean_(row[roomField], 100);
    return row;
  });
}
function strip_(row) { const obj = Object.assign({}, row); delete obj._row; return obj; }
function clean_(value, length) {
  return String(value == null ? '' : value).trim().slice(0, length || 500).replace(/[\u0000-\u001f]/g, ' ');
}
// Google Sheets interprets user-entered leading formula characters. Keep them literal.
function safe_(value, length) {
  const s = clean_(value, length);
  return /^[=+\-@\t\r]/.test(s) ? "'" + s : s;
}
function positive_(n, label, allowZero) {
  if (n === '' || n === null || n === undefined) throw new Error(label + ' is required.');
  const value = Number(n);
  if (!Number.isSafeInteger(value) || value < (allowZero ? 0 : 1)) throw new Error(label + ' must be a ' + (allowZero ? 'nonnegative' : 'positive') + ' whole number.');
  return value;
}
function log_(action, item, quantity, actor, details) {
  append_('Activity', {id: Utilities.getUuid(), timestamp: new Date().toISOString(), action, itemId: item.id, itemName: item.name, quantity, actor: safe_(actor, 100), details: safe_(details, 500)});
}
function append_(name, obj) { sheet_(name).appendRow(TABLES[name].map(k => obj[k] == null ? '' : obj[k])); }
function write_(name, obj) { sheet_(name).getRange(obj._row, 1, 1, TABLES[name].length).setValues([TABLES[name].map(k => obj[k] == null ? '' : obj[k])]); }

// Ask locked_ to send fresh data AFTER the lock is released.
function reload_(message, extra) { return {__reload: true, message: message || '', extra: extra || null}; }
function locked_(fn) {
  // The schema check is not repeated on every save; page loads and runSetup() handle it.
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new Error('Inventory is busy. Try again in a moment.');
  let out;
  try {
    out = fn();
    SpreadsheetApp.flush(); // Commit writes before another request can take the lock.
  } finally { lock.releaseLock(); }
  if (!out || out.__reload !== true) return out; // Older callers that return their own data still work.
  const data = getLiveData_();
  if (out.message) data._message = out.message;
  return out.extra ? Object.assign({}, out.extra, {data: data}) : data;
}

function livePriceJobs_() {
  const cutoff = Date.now() - DONE_JOB_DAYS_ * 86400000;
  return rows_('PriceJobs').filter(j => j.status !== 'Done' || !(Date.parse(j.updatedAt) < cutoff)).map(strip_);
}
// Everything that changes during the day. The catalog and price catalog are sent once, by getData().
function getLiveData_() {
  return {items: rows_('Items').map(strip_), loans: rows_('Loans').filter(x => x.status === 'Open').map(strip_), issues: rows_('Issues').filter(x => x.status === 'Open').map(strip_), activity: rows_('Activity', 200).reverse().map(strip_), locations: rows_('Locations').map(x => x.location).filter(Boolean), prices: rows_('PriceEstimates').map(strip_), priceJobs: livePriceJobs_(), pricing: automaticPricingStatus_()};
}
function getData() {
  ensureReadReady_();
  const data = getLiveData_();
  data.catalog = catalogRowsCached_().map(strip_);
  data.priceCatalog = rows_('PriceCatalog').map(strip_);
  return data;
}
// Small background refresh used by the inventory page while supplier lookups are pending.
function getPricingUpdates() {
  return {prices: rows_('PriceEstimates').map(strip_), priceJobs: livePriceJobs_(), pricing: automaticPricingStatus_()};
}
function unavailable_(item) { return Number(item.checkedOut) + Number(item.needsRepair || 0) + Number(item.needsReplacement || 0); }
function nextAutoSku_(usedSkus) {
  const used = usedSkus || new Set(rows_('Items').filter(x => x.sku).map(x => String(x.sku).toUpperCase()));
  let highest = 0;
  used.forEach(sku => {
    const match = String(sku).match(/^EAHS-(\d+)$/i);
    if (match) highest = Math.max(highest, Number(match[1]));
  });
  let number = highest + 1;
  let sku = 'EAHS-' + String(number).padStart(6, '0');
  while (used.has(sku.toUpperCase())) { number += 1; sku = 'EAHS-' + String(number).padStart(6, '0'); }
  used.add(sku.toUpperCase());
  return sku;
}
// Same result as before; each distinct name is only normalized once per request.
const nameSignatureMemo_ = new Map();
function nameSignature_(value) {
  const key = value == null ? '' : String(value);
  let signature = nameSignatureMemo_.get(key);
  if (signature === undefined) { signature = nameSignatureRaw_(key); nameSignatureMemo_.set(key, signature); }
  return signature;
}
function nameSignatureRaw_(value) {
  const text = clean_(value, 160).toLowerCase().normalize('NFKD').replace(/[̀-ͯ]/g, '')
    .replace(/\b(\d+(?:\.\d+)?)\s*(?:milliliters?|millilitres?|ml)\b/g, (_, amount) => Number(amount) + 'ml')
    .replace(/\b(\d+(?:\.\d+)?)\s*(?:microliters?|microlitres?|µl|μl|ul)\b/g, (_, amount) => Number(amount) + 'ul')
    .replace(/\b(\d+(?:\.\d+)?)\s*(?:liters?|litres?|l)\b/g, (_, amount) => Number(amount) * 1000 + 'ml')
    .replace(/×/g, 'x');
  const singular = {beakers:'beaker', flasks:'flask', cylinders:'cylinder', pipettes:'pipette', pipets:'pipet', tubes:'tube', dishes:'dish', bottles:'bottle', microscopes:'microscope'};
  return text.replace(/[^a-z0-9]+/g, ' ').trim().split(/\s+/).filter(Boolean).map(word => singular[word] || word).sort().join('|');
}
function catalogAlias_(name) {
  const parts = name.split(' — ');
  return parts[0].split(',')[0] + (parts.length > 1 ? ' ' + parts.slice(1).join(' ') : '');
}
// The Catalog tab is read once per request and indexed by normalized name.
let catalogRowsMemo_ = null, catalogIndexMemo_ = null;
function catalogRowsCached_() { return catalogRowsMemo_ || (catalogRowsMemo_ = rows_('Catalog')); }
function catalogIndex_() {
  if (catalogIndexMemo_) return catalogIndexMemo_;
  const full = new Map(), alias = new Map();
  const add = (map, key, row) => { if (!map.has(key)) map.set(key, []); map.get(key).push(row); };
  catalogRowsCached_().forEach(row => {
    add(full, nameSignature_(row.name), row);
    add(alias, nameSignature_(catalogAlias_(String(row.name))), row);
  });
  return catalogIndexMemo_ = {full, alias};
}
function canonicalName_(rawName, kind, unit, existing) {
  const key = nameSignature_(rawName), wantedUnit = unit.toLowerCase(), index = catalogIndex_();
  // Standard catalog names take precedence over historical names, including exact historical matches.
  for (const wantedKind of [kind, 'Review']) {
    for (const map of [index.full, index.alias]) {
      const matches = [...new Set((map.get(key) || []).filter(row => row.kind === wantedKind && String(row.unit).toLowerCase() === wantedUnit).map(row => row.name))];
      if (matches.length === 1) return matches[0];
      if (matches.length > 1) throw new Error('Several catalog variants match this name. Choose the full name specifying material or type: ' + matches.slice(0, 3).join('; '));
    }
  }
  const known = (existing || []).find(item => nameSignature_(item.name) === key && item.kind === kind && String(item.unit).toLowerCase() === wantedUnit);
  if (known) return known.name;
  return rawName.replace(/\b(\d+(?:\.\d+)?)\s*ml\b/gi, '$1 mL').replace(/\b(\d+(?:\.\d+)?)\s*(?:µl|μl|ul)\b/gi, '$1 µL');
}

function itemKey_(item) { return nameSignature_(item.name) + '|' + clean_(item.location).toLowerCase() + '|' + clean_(item.kind).toLowerCase(); }
function sameStock_(existing, incoming) {
  if (existing.unit.toLowerCase() !== incoming.unit.toLowerCase()) throw new Error('Matching item has a different counting unit. Use its existing unit or give this material a distinct name.');
  if (existing.sku && incoming.sku && existing.sku.toLowerCase() !== incoming.sku.toLowerCase()) throw new Error('Matching item has a different SKU. Use Add stock on the correct entry or distinguish the item names.');
  if (existing.catalogNumber && incoming.catalogNumber && existing.catalogNumber.toLowerCase() !== incoming.catalogNumber.toLowerCase()) throw new Error('Matching item has a different catalog number. Check which entry should receive these units.');
}
// Categories are learned from saved inventory and catalog records; no new tab is needed.
const categoryKeyMemo_ = new Map();
function categoryKey_(value) {
  const key = value == null ? '' : String(value);
  let result = categoryKeyMemo_.get(key);
  if (result === undefined) { result = clean_(key, 1000).replace(/\s+/g, ' ').toLowerCase(); categoryKeyMemo_.set(key, result); }
  return result;
}
function categoryInput_(raw, knownItems) {
  const custom = categoryKey_(raw.category) === 'other' && raw.newCategory != null;
  const label = clean_(custom ? raw.newCategory : raw.category, 1000).replace(/\s+/g, ' ') || (custom ? '' : 'Other');
  if (!label) throw new Error('Enter a name for the new category.');
  if (label.length > 80) throw new Error('Category names must be 80 characters or fewer.');
  const safe = safe_(label, 80), wanted = categoryKey_(safe);
  const prior = (knownItems || []).find(row => categoryKey_(row.category) === wanted) || catalogRowsCached_().find(row => categoryKey_(row.category) === wanted);
  return prior ? prior.category : safe;
}
function itemInput_(raw, previous, allowedRooms, knownItems) {
  const rawName = safe_(raw.name, 160), location = safe_(raw.location, 100);
  const kind = clean_(raw.kind, 20), unit = safe_(raw.unit, 40) || 'each';
  const name = rawName ? canonicalName_(rawName, kind, unit, knownItems) : '';
  if (!name || !location) throw new Error('Item name and home classroom or prep area are required.');
  if (!(allowedRooms || new Set(rows_('Locations').map(x => x.location))).has(location)) throw new Error('Select a listed classroom or prep area.');
  if (!['Reusable','Consumable'].includes(kind)) throw new Error('Select Reusable or Consumable.');
  const quantity = positive_(raw.quantity, 'Quantity', true);
  if (previous && quantity < unavailable_(previous)) throw new Error('Quantity cannot be less than checked out or flagged units.');
  const expiration = clean_(raw.expiration, 10);
  if (expiration && !/^\d{4}-\d{2}-\d{2}$/.test(expiration)) throw new Error('Use YYYY-MM-DD for expiration.');
  return {
    id: previous ? previous.id : Utilities.getUuid(), sku: safe_(raw.sku, 80), name,
    category: categoryInput_(raw, knownItems), kind, quantity,
    checkedOut: previous ? Number(previous.checkedOut) : 0,
    unit, location, storage: safe_(raw.storage, 120),
    minStock: positive_(raw.minStock === '' || raw.minStock == null ? 0 : raw.minStock, 'Reorder level', true),
    manufacturer: safe_(raw.manufacturer, 120), catalogNumber: safe_(raw.catalogNumber, 100),
    hazard: safe_(raw.hazard, 120), expiration, notes: safe_(raw.notes, 500), updatedAt: new Date().toISOString(),
    needsRepair: previous ? Number(previous.needsRepair || 0) : 0,
    needsReplacement: previous ? Number(previous.needsReplacement || 0) : 0
  };
}
function saveItem(raw) {
  return locked_(() => {
    const existing = rows_('Items');
    const usedSkus = new Set(existing.filter(x => x.sku).map(x => String(x.sku).toUpperCase()));
    const old = raw.id ? existing.find(x => x.id === raw.id) : null;
    if (raw.id && !old) throw new Error('This item no longer exists. Refresh the inventory.');
    const enteredPrice = enteredUnitPrice_(raw.unitPrice);
    validatePriceSource_(raw.priceSource);
    const item = itemInput_(raw, old, null, existing);
    if (old && old.kind !== item.kind && Number(old.checkedOut) > 0) throw new Error('Return outstanding loans before changing item type.');
    if (old && existing.some(x => x.id !== old.id && itemKey_(x) === itemKey_(item))) throw new Error('An entry for this item already exists in that location. Use Add stock on the existing entry.');
    const match = old ? null : existing.find(x => itemKey_(x) === itemKey_(item));
    if (match) {
      sameStock_(match, item);
      if (item.sku && existing.some(x => x.id !== match.id && x.location === item.location && x.sku && x.sku.toLowerCase() === item.sku.toLowerCase())) throw new Error('This SKU belongs to another inventory entry in this room.');
      const received = positive_(item.quantity, 'Additional quantity');
      if (!match.sku) match.sku = nextAutoSku_(usedSkus);
      match.quantity = Number(match.quantity) + received;
      match.updatedAt = new Date().toISOString(); write_('Items', match);
      log_('RESTOCK', match, received, raw.actor || '', 'Matching item in ' + match.location + '; new total ' + match.quantity);
      ensureAutomaticPrice_(match, {...raw,unitPrice:enteredPrice});
      return reload_('Added ' + received + ' ' + match.unit + ' to ' + match.name + ' in ' + match.location + '. New total: ' + match.quantity + '.');
    }
    if (item.sku && existing.some(x => x.location === item.location && x.sku && x.sku.toLowerCase() === item.sku.toLowerCase() && x.id !== item.id)) throw new Error('This SKU belongs to another entry in this room. Use Add stock on that entry.');
    if (!item.sku) item.sku = old && old.sku ? old.sku : nextAutoSku_(usedSkus);
    if (old) { item._row = old._row; write_('Items', item); }
    else append_('Items', item);
    log_(old ? 'EDIT' : 'ADD', item, item.quantity, raw.actor || '', old ? 'Item details updated' : 'New item');
    ensureAutomaticPrice_(item, {...raw,unitPrice:enteredPrice});
    return reload_();
  });
}
function receiveStock(raw) {
  return locked_(() => {
    const item = rows_('Items').find(x => x.id === raw.itemId);
    if (!item) throw new Error('Inventory item not found. Refresh the page.');
    const quantity = positive_(raw.quantity, 'Quantity received');
    const by = safe_(raw.actor, 100);
    if (!by) throw new Error('Enter the staff member receiving the stock.');
    item.quantity = Number(item.quantity) + quantity;
    item.updatedAt = new Date().toISOString(); write_('Items', item);
    log_('RESTOCK', item, quantity, by, 'New total ' + item.quantity + (raw.notes ? '; ' + clean_(raw.notes, 250) : ''));
    return reload_('Added ' + quantity + ' ' + item.unit + ' to ' + item.name + ' in ' + item.location + '. New total: ' + item.quantity + '.');
  });
}
function importItems(rawItems, actor, addToExisting) {
  if (!Array.isArray(rawItems) || rawItems.length > 500) throw new Error('Import up to 500 rows at a time.');
  return locked_(() => {
    if (addToExisting && !safe_(actor, 100)) throw new Error('Enter the staff member receiving stock before importing.');
    const existing = rows_('Items');
    const usedSkus = new Set(existing.filter(x => x.sku).map(x => String(x.sku).toUpperCase()));
    const allowedRooms = new Set(rows_('Locations').map(x => x.location));
    const skuKey = x => x.sku.toLowerCase() + '|' + x.location;
    const bySku = new Map(existing.filter(x => x.sku).map(x => [skuKey(x), x]));
    const byKey = new Map(existing.map(x => [itemKey_(x), x]));
    const prepared = [], increments = new Map(), errors = [], skipped = [];
    rawItems.forEach((raw, i) => {
      try {
        const unitPrice=enteredUnitPrice_(raw.unitPrice);validatePriceSource_(raw.priceSource);
        const item = itemInput_(raw, null, allowedRooms, existing.concat(prepared));
        item._enteredPrice=unitPrice;item._priceSource=raw.priceSource;
        const sameName = byKey.get(itemKey_(item));
        const sameSku = item.sku && bySku.get(skuKey(item));
        if (sameSku && sameName && sameSku.id !== sameName.id) throw new Error('SKU and item name point to different records.');
        if (sameSku && itemKey_(sameSku) !== itemKey_(item)) {
          if (addToExisting) throw new Error('SKU belongs to another item or home location.');
          skipped.push(i + 2); return;
        }
        const match = sameName || sameSku;
        if (match) {
          if (!addToExisting) { skipped.push(i + 2); return; }
          sameStock_(match, item);
          const amount = positive_(item.quantity, 'Additional quantity');
          if (match._row) increments.set(match.id, {item: match, quantity: (increments.get(match.id)?.quantity || 0) + amount, unitPrice: unitPrice??increments.get(match.id)?.unitPrice, priceSource: raw.priceSource||increments.get(match.id)?.priceSource});
          else match.quantity = Number(match.quantity) + amount; // Merge repeated new entries inside this same CSV.
          return;
        }
        if (!item.sku) item.sku = nextAutoSku_(usedSkus);
        else usedSkus.add(item.sku.toUpperCase());
        byKey.set(itemKey_(item), item);
        if (item.sku) bySku.set(skuKey(item), item);
        prepared.push(item);
      } catch (e) { errors.push('CSV row ' + (i + 2) + ': ' + e.message); }
    });
    if (errors.length) throw new Error('No rows imported. ' + errors.slice(0, 8).join(' ') + (errors.length > 8 ? ' (and ' + (errors.length - 8) + ' more)' : ''));
    increments.forEach(({item, quantity}) => {
      item.quantity = Number(item.quantity) + quantity;
      item.updatedAt = new Date().toISOString(); write_('Items', item);
      log_('RESTOCK', item, quantity, actor, 'CSV stock receipt; new total ' + item.quantity);
    });
    if (prepared.length) {
      sheet_('Items').getRange(sheet_('Items').getLastRow() + 1, 1, prepared.length, TABLES.Items.length).setValues(prepared.map(x => TABLES.Items.map(k => x[k] == null ? '' : x[k])));
      const activity = prepared.map(item => [Utilities.getUuid(), new Date().toISOString(), 'IMPORT', item.id, item.name, item.quantity, safe_(actor, 100), 'CSV import']);
      sheet_('Activity').getRange(sheet_('Activity').getLastRow() + 1, 1, activity.length, TABLES.Activity.length).setValues(activity);
    }
    const priceBatch=automaticPriceBatch_(true);
    increments.forEach(({item,unitPrice,priceSource})=>ensureAutomaticPrice_(item,{unitPrice,priceSource,actor},priceBatch));
    prepared.forEach(item=>ensureAutomaticPrice_(item,{unitPrice:item._enteredPrice,priceSource:item._priceSource,actor},priceBatch));
    flushAutomaticPriceBatch_(priceBatch);
    return reload_('', {added: prepared.length, restocked: increments.size, skipped: skipped.length, skippedRows: skipped});
  });
}
function checkout(raw) {
  return locked_(() => {
    const item = rows_('Items').find(x => x.id === raw.itemId);
    if (!item) throw new Error('Item not found. Refresh the inventory.');
    const quantity = positive_(raw.quantity, 'Quantity');
    if (quantity > Number(item.quantity) - unavailable_(item)) throw new Error('Only ' + (Number(item.quantity) - unavailable_(item)) + ' ' + item.unit + ' available.');
    const borrower = safe_(raw.borrower, 100), destination = safe_(raw.destination, 100);
    if (!borrower || !destination) throw new Error('Staff member and destination are required.');
    if (!rows_('Locations').some(x => x.location === destination)) throw new Error('Choose a listed destination.');
    const dueDate = clean_(raw.dueDate, 10);
    if (dueDate && !/^\d{4}-\d{2}-\d{2}$/.test(dueDate)) throw new Error('Use YYYY-MM-DD for due date.');
    if (item.kind === 'Reusable') {
      item.checkedOut = Number(item.checkedOut) + quantity;
      append_('Loans', {id: Utilities.getUuid(), itemId: item.id, quantity, remaining: quantity, borrower, destination, checkedOutAt: new Date().toISOString(), dueDate, returnedAt: '', status: 'Open', notes: safe_(raw.notes, 300)});
      log_('CHECK OUT', item, quantity, borrower, 'To ' + destination + (dueDate ? '; due ' + dueDate : ''));
    } else {
      item.quantity = Number(item.quantity) - quantity;
      log_('USE', item, quantity, borrower, 'Used in ' + destination + (raw.notes ? '; ' + raw.notes : ''));
    }
    item.updatedAt = new Date().toISOString(); write_('Items', item);
    return reload_();
  });
}
function returnLoan(raw) {
  return locked_(() => {
    const loan = rows_('Loans').find(x => x.id === raw.loanId && x.status === 'Open');
    if (!loan) throw new Error('Checkout already closed or not found. Refresh the page.');
    const quantity = positive_(raw.quantity, 'Return quantity');
    if (quantity > Number(loan.remaining)) throw new Error('Only ' + loan.remaining + ' still checked out.');
    const item = rows_('Items').find(x => x.id === loan.itemId);
    if (!item) throw new Error('Inventory item for this checkout is missing.');
    loan.remaining = Number(loan.remaining) - quantity;
    if (loan.remaining === 0) { loan.status = 'Returned'; loan.returnedAt = new Date().toISOString(); }
    item.checkedOut = Number(item.checkedOut) - quantity;
    item.updatedAt = new Date().toISOString();
    write_('Loans', loan); write_('Items', item);
    log_('RETURN', item, quantity, raw.actor || loan.borrower, 'From ' + loan.destination + (raw.notes ? '; ' + raw.notes : ''));
    return reload_();
  });
}
function adjustStock(raw) {
  return locked_(() => {
    const item = rows_('Items').find(x => x.id === raw.itemId);
    if (!item) throw new Error('Item not found. Refresh the page.');
    const next = positive_(raw.quantity, 'New total quantity', true);
    if (next < unavailable_(item)) throw new Error('New total cannot be below checked out or flagged quantities.');
    const old = Number(item.quantity);
    item.quantity = next; item.updatedAt = new Date().toISOString(); write_('Items', item);
    log_('COUNT', item, next - old, raw.actor || '', 'Count corrected from ' + old + ' to ' + next + '; ' + clean_(raw.reason, 250));
    return reload_();
  });
}
function reportIssue(raw) {
  return locked_(() => {
    const item = rows_('Items').find(x => x.id === raw.itemId);
    if (!item) throw new Error('Item not found. Refresh the page.');
    const type = clean_(raw.type, 20);
    if (!['Repair','Replacement'].includes(type)) throw new Error('Select repair or replacement.');
    const quantity = positive_(raw.quantity, 'Affected quantity');
    if (quantity > Number(item.quantity) - unavailable_(item)) throw new Error('Only available units can be flagged. Return checked-out units first.');
    const by = safe_(raw.reportedBy, 100), details = safe_(raw.details, 500);
    if (!by || !details) throw new Error('Staff member and description are required.');
    append_('Issues', {id: Utilities.getUuid(), itemId: item.id, type, quantity, remaining: quantity, reportedBy: by, details, reportedAt: new Date().toISOString(), status: 'Open'});
    const field = type === 'Repair' ? 'needsRepair' : 'needsReplacement';
    item[field] = Number(item[field] || 0) + quantity;
    item.updatedAt = new Date().toISOString(); write_('Items', item);
    log_(type.toUpperCase() + ' NEEDED', item, quantity, by, details);
    return reload_();
  });
}
// Attach a staff-confirmed request to units already flagged in the Sheet.
// Recheck under the inventory lock; never flag those units a second time.
function recordExistingIssue(raw) {
  return locked_(() => {
    const item = rows_('Items').find(x => x.id === raw.itemId);
    if (!item) throw new Error('Item not found. Refresh the inventory.');
    const type = clean_(raw.type, 20);
    if (!['Repair','Replacement'].includes(type)) throw new Error('Select repair or replacement.');
    const field = type === 'Repair' ? 'needsRepair' : 'needsReplacement';
    const quantity = positive_(raw.quantity, 'Already flagged quantity');
    const recorded = rows_('Issues').filter(x => x.itemId === item.id && x.type === type && x.status === 'Open')
      .reduce((sum, x) => sum + Number(x.remaining || 0), 0);
    if (!Number.isSafeInteger(recorded) || recorded < 0) throw new Error('Open request quantities are inconsistent. Reconcile the Sheet before recording a request.');
    const unrecorded = positive_(item[field] || 0, 'Current flagged quantity', true) - recorded;
    if (quantity > unrecorded) throw new Error('These flags already have requests or the item changed. Refresh before recording a request.');
    const by = safe_(raw.reportedBy, 100), details = safe_(raw.details, 500);
    if (!by || !details) throw new Error('Staff member and description are required.');
    append_('Issues', {id: Utilities.getUuid(), itemId: item.id, type, quantity, remaining: quantity,
      reportedBy: by, details, reportedAt: new Date().toISOString(), status: 'Open'});
    log_('FLAG REQUEST RECORDED', item, quantity, by, type + ': ' + details + '; existing counts unchanged');
    return reload_();
  });
}

function resolveIssue(raw) {
  return locked_(() => {
    const issue = rows_('Issues').find(x => x.id === raw.issueId && x.status === 'Open');
    if (!issue) throw new Error('Request already closed or not found. Refresh the page.');
    const item = rows_('Items').find(x => x.id === issue.itemId);
    if (!item) throw new Error('Inventory item for this request is missing.');
    const quantity = positive_(raw.quantity, 'Resolution quantity');
    if (quantity > Number(issue.remaining)) throw new Error('Only ' + issue.remaining + ' unit(s) remain flagged.');
    const resolution = clean_(raw.resolution, 30);
    if (!['Back in service','Removed from inventory'].includes(resolution)) throw new Error('Select how the request was resolved.');
    const by = safe_(raw.resolvedBy, 100);
    if (!by) throw new Error('Staff member is required.');
    const field = issue.type === 'Repair' ? 'needsRepair' : 'needsReplacement';
    if (quantity > positive_(item[field] || 0, 'Current flagged quantity', true)) throw new Error('The item flag count no longer covers this request. Refresh and reconcile the Sheet before resolving it.');
    item[field] = Number(item[field] || 0) - quantity;
    if (resolution === 'Removed from inventory') item.quantity = Number(item.quantity) - quantity;
    issue.remaining = Number(issue.remaining) - quantity;
    if (issue.remaining === 0) { issue.status = 'Resolved'; issue.resolvedAt = new Date().toISOString(); }
    issue.resolvedBy = by; issue.resolution = resolution;
    item.updatedAt = new Date().toISOString(); write_('Issues', issue); write_('Items', item);
    log_('ISSUE RESOLVED', item, quantity, by, resolution + (raw.notes ? '; ' + clean_(raw.notes, 250) : ''));
    return reload_();
  });
}

function priceKey_(item) {
  return nameSignature_(item.name) + '|' + item.kind + '|' + String(item.unit || 'each').trim().toLowerCase();
}
function priceQuotes_(rawQuotes) {
  if (!Array.isArray(rawQuotes) || !rawQuotes.length || rawQuotes.length > 6) throw new Error('Enter one to six supplier quotes.');
  const seen = new Set();
  return rawQuotes.map(raw => {
    const supplier = clean_(raw.supplier, 150), url = clean_(raw.url, 1200);
    const price = Number(raw.price), packQuantity = Number(raw.packQuantity);
    if (!supplier || !/^https?:\/\/[^\s<>"']+$/i.test(url)) throw new Error('Each quote needs a supplier and a full http(s) source link.');
    if (!Number.isFinite(price) || price <= 0 || price > 100000000) throw new Error('Quote price must be greater than zero.');
    if (!Number.isSafeInteger(packQuantity) || packQuantity < 1) throw new Error('Inventory units per pack must be a positive whole number.');
    const sourceKey = url.toLowerCase();
    if (seen.has(sourceKey)) throw new Error('Use each supplier listing only once in an estimate.');
    seen.add(sourceKey);
    return {supplier, url, price, packQuantity};
  });
}
function savePriceEstimate(raw) {
  return locked_(() => {
    const item = rows_('Items').find(x => x.id === raw.itemId);
    if (!item) throw new Error('Choose an inventory item before pricing it.');
    const quotes = priceQuotes_(raw.quotes);
    const matchType = clean_(raw.matchType, 40);
    if (!['Exact model','Comparable item','Confirm model'].includes(matchType)) throw new Error('Choose how closely the quotes match the inventory item.');
    const checkedAt = clean_(raw.checkedAt, 10), actor = safe_(raw.actor, 100);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(checkedAt) || (!Number.isFinite(Date.parse(checkedAt)) || new Date(checkedAt).toISOString().slice(0,10) !== checkedAt)) throw new Error('Enter the date the quotes were checked.');
    if (!actor) throw new Error('Enter the staff member checking these prices.');
    const prior = rows_('PriceEstimates').find(x => priceKey_(x) === priceKey_(item));
    const estimate = {id: prior ? prior.id : Utilities.getUuid(), name: item.name, kind: item.kind, unit: item.unit,
      averageUnitCost: quotes.reduce((sum, q) => sum + q.price / q.packQuantity, 0) / quotes.length,
      quotesJson: JSON.stringify(quotes), checkedAt, matchType, notes: safe_(raw.notes, 1000), updatedBy: actor};
    if (prior) { estimate._row = prior._row; write_('PriceEstimates', estimate); }
    else append_('PriceEstimates', estimate);
    log_('PRICE ESTIMATE', item, 0, actor, 'USD ' + estimate.averageUnitCost.toFixed(4) + ' per ' + item.unit + '; ' + quotes.length + ' source(s); ' + matchType);
    finishPriceJobs_(item,'Done','Staff quote saved; automatic search will not overwrite it.');
    return reload_('Replacement price saved for matching names, types, and units across all rooms.');
  });
}

const PRICE_SEED_ROWS = [["price-6ea57478-8330-4460-b6aa-ecbf1d733cc4", "Compound light microscope", "Reusable", "each", 774.9, "[{\"supplier\":\"Ward's Science — Swift M3604C-4\",\"url\":\"https://www.wardsci.com/us/en/product/27154368/swift-m3600-series-compound-microscopes\",\"price\":774.9,\"packQuantity\":1}]", "2026-09-29", "Confirm model", "Representative Swift M3604C-4 price. Recorded M3600/M3700 family covers multiple configurations; confirm each exact submodel before ordering. New-equipment estimate.", "Initial supplier search"], ["price-b8de8bf3-dbc7-4ef0-8b49-b0dec9161ec9", "DNA Playground", "Reusable", "each", 592, "[{\"supplier\":\"Amino Labs\",\"url\":\"https://amino.bio/products/dna-playground\",\"price\":592,\"packQuantity\":1}]", "2026-09-29", "Confirm model", "Manufacturer listing displays USD 592 with Large/Small variants. Confirm Large Classroom selection and final quote before ordering.", "Initial supplier search"], ["price-fef08bc8-6715-4dbe-8c37-8b8e90d85bb6", "Bambu H2D", "Reusable", "each", 1549, "[{\"supplier\":\"Bambu Lab US — standalone H2D\",\"url\":\"https://us.store.bambulab.com/products/h2d\",\"price\":1549,\"packQuantity\":1}]", "2026-09-29", "Exact model", "Current standalone printer sale price. Excludes AMS and laser/cutting modules; AMS is priced separately. Typical listed printer price is USD 1749.", "Initial supplier search"], ["price-93c64550-59af-42bd-93ab-fc07ad87dd05", "Bambu AMS Pro 2", "Reusable", "each", 299, "[{\"supplier\":\"Bambu Lab US — AMS 2 Pro\",\"url\":\"https://us.store.bambulab.com/collections/ams-bundles\",\"price\":299,\"packQuantity\":1}]", "2026-09-29", "Confirm model", "Official listing starts at USD 299. Assumes AMS 2 Pro base unit; verify power adapter/upgrade kit requirements and selected bundle.", "Initial supplier search"], ["price-67368710-f53d-4978-9b18-9bca56882321", "Beaker, low form borosilicate — 50 mL", "Reusable", "each", 3.2333333333333334, "[{\"supplier\":\"Ward's Science — 50 mL standard borosilicate (470149-366)\",\"url\":\"https://www.wardsci.com/us/en/product/8892248/vwr-standard-standard-grade-beakers\",\"price\":38.1,\"packQuantity\":12},{\"supplier\":\"Carolina — 50 mL standard borosilicate\",\"url\":\"https://www.carolina.com/lab-beakers/carolina-beakers-standard-grade-50-ml-pack-of-12/731003.pr\",\"price\":39.5,\"packQuantity\":12}]", "2026-09-29", "Comparable item", "Preferred supplier review. Mean of Ward's and Carolina unit prices. Pack prices divided by 12. Confirm glass grade/brand before ordering; tax and shipping excluded. Flinn current price requires staff verification.", "Preferred supplier review"], ["price-527b420e-d7cb-4dca-84f2-7b5b6fb792df", "Beaker, low form borosilicate — 100 mL", "Reusable", "each", 3.4583333333333335, "[{\"supplier\":\"Ward's Science — 100 mL standard borosilicate (470149-364)\",\"url\":\"https://www.wardsci.com/us/en/product/8892248/vwr-standard-standard-grade-beakers\",\"price\":41.5,\"packQuantity\":12}]", "2026-09-29", "Comparable item", "Preferred supplier review. Ward's unit price; one verified preferred-supplier quote. Pack prices divided by 12. Confirm glass grade/brand before ordering; tax and shipping excluded. Flinn current price requires staff verification.", "Preferred supplier review"], ["price-06f8912a-6c4d-48bb-822a-b407879e66b9", "Beaker, low form borosilicate — 250 mL", "Reusable", "each", 4.558333333333334, "[{\"supplier\":\"Ward's Science — 250 mL standard borosilicate (470149-360)\",\"url\":\"https://www.wardsci.com/us/en/product/8892248/vwr-standard-standard-grade-beakers\",\"price\":66.4,\"packQuantity\":12},{\"supplier\":\"Carolina — 250 mL standard borosilicate\",\"url\":\"https://www.carolina.com/lab-beakers/carolina-beakers-standard-grade-250-ml-pack-of-12/731009.pr\",\"price\":43,\"packQuantity\":12}]", "2026-09-29", "Comparable item", "Preferred supplier review. Mean of Ward's and Carolina unit prices. Pack prices divided by 12. Confirm glass grade/brand before ordering; tax and shipping excluded. Flinn current price requires staff verification.", "Preferred supplier review"], ["price-cdd5b10d-bf6d-4a1b-94b2-cd80ef682ef1", "Beaker, low form borosilicate — 400 mL", "Reusable", "each", 9.5, "[{\"supplier\":\"Ward's Science — 400 mL standard borosilicate (470149-358)\",\"url\":\"https://www.wardsci.com/us/en/product/8892248/vwr-standard-standard-grade-beakers\",\"price\":114,\"packQuantity\":12}]", "2026-09-29", "Comparable item", "Preferred supplier review. Ward's unit price; one verified preferred-supplier quote. Pack prices divided by 12. Confirm glass grade/brand before ordering; tax and shipping excluded. Flinn current price requires staff verification.", "Preferred supplier review"], ["price-3ccc1542-e280-4054-9325-b54000943e23", "Beakers, 20mL", "Reusable", "each", 4.133333333333334, "[{\"supplier\":\"Ward's Science — 20 mL standard borosilicate (470149-370)\",\"url\":\"https://www.wardsci.com/us/en/product/8892248/vwr-standard-standard-grade-beakers\",\"price\":49.6,\"packQuantity\":12}]", "2026-09-29", "Comparable item", "Preferred supplier review. Ward's unit price; one verified preferred-supplier quote. Pack prices divided by 12. Confirm glass grade/brand before ordering; tax and shipping excluded. Flinn current price requires staff verification.", "Preferred supplier review"], ["price-fa951ab6-a5fb-4981-8e94-2e16b3da63e4", "Erlynmeyer Flask 125 mL", "Reusable", "each", 5.358333333333333, "[{\"supplier\":\"Ward's Science — 125 mL standard borosilicate flask (470149-312)\",\"url\":\"https://www.wardsci.com/us/en/product/8892250/vwr-standard-standard-grade-erlenmeyer-flasks\",\"price\":64.3,\"packQuantity\":12}]", "2026-09-29", "Comparable item", "Preferred supplier review. Ward's 125 mL standard borosilicate flask pack of 12, divided by 12 inventory units. Confirm neck/closure type. Single verified preferred-supplier quote; tax and shipping excluded. Flinn current price requires staff verification.", "Preferred supplier review"]];

// Return project-owned client JavaScript separately from the initial HTML shell.
// getRawContent avoids template evaluation and preserves the original code exactly.
function getInventoryClient() {
  try {
    return HtmlService.createTemplateFromFile('InventoryClient').getRawContent();
  } catch (e) {
    throw new Error('Cannot read InventoryClient.html. Create an HTML file named InventoryClient, paste the supplied content, and deploy a new version. ' + e.message);
  }
}
