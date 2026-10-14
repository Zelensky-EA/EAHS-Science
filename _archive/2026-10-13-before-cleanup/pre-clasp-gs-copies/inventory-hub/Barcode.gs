// Barcode lookup: saved identities, UPCitemdb, Open Facts (four databases), Open Library, FDA/NLM AccessGUDID.
// Free public read-only services; no API keys, subscriptions or automatic inventory saves.
// Optimized edition: the shared script lock is held only to reserve a lookup slot, never during a web request,
// and inventory checks read the Items tab instead of the full inventory payload.
function eahsBarcode_(raw) {
  var code = String(raw || '').trim();
  if (!/^(\d{8}|\d{12}|\d{13}|\d{14})$/.test(code)) throw new Error('Use an 8, 12, 13 or 14 digit UPC/EAN barcode.');
  var sum = 0;
  for (var i = code.length - 2, weight = 3; i >= 0; i--, weight = weight === 3 ? 1 : 3) sum += Number(code[i]) * weight;
  if ((10 - sum % 10) % 10 !== Number(code[code.length - 1])) throw new Error('Barcode check digit is invalid. Scan again.');
  return code.padStart(14, '0');
}
function eahsBarcodeLinks_(code) {
  return JSON.parse(PropertiesService.getScriptProperties().getProperty('EAHS_PRODUCT_BARCODE_' + code) || '[]');
}
function lookupInventoryBarcode(raw) {
  var code = eahsBarcode_(raw), items = rows_('Items').map(strip_), links = eahsBarcodeLinks_(code);
  var matches = items.filter(function(item) {
    return links.some(function(link) { return link.itemId === item.id; }) || [item.barcode, item.catalogNumber].some(function(value) {
      try { return eahsBarcode_(value) === code; } catch (e) { return false; }
    });
  }).map(function(item) {
    return {item: item, unitsPerPack: (links.find(function(link) { return link.itemId === item.id && link.unit === item.unit; }) || {}).unitsPerPack || ''};
  });
  // Confirmed inventory identities win. Unknown products use the free online service.
  return Object.assign({barcode: String(raw).trim(), matches: matches}, matches.length ? {products: [], lookupStatus: 'inventory'} : eahsLookupOnline_(raw));
}
function eahsLookupOnline_(raw) {
  var code = eahsBarcode_(raw), isbn = /^097[89]/.test(code) ? code.slice(1) : '';
  if (!isbn) return eahsOnlineBarcode_(raw);
  var book = eahsBookBarcode_(isbn);
  if (book.products.length) return book;
  var product = eahsOnlineBarcode_(raw);
  if (product.products.length) return product;
  if (book.lookupStatus !== 'not_found') return book;
  if (product.lookupStatus !== 'not_found') return product;
  return {products: [], lookupStatus: 'not_found', message: 'Neither the book catalog nor the product database found this exact ISBN. Verify the digits or enter the book details manually.'};
}
function eahsBarcodeText_(value, max) {
  return String(value || '').replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, max);
}
function eahsBookProduct_(isbn, book, source, basis) {
  var clean = eahsBarcodeText_, title = clean(book.title, 300), subtitle = clean(book.subtitle, 200);
  if (subtitle) title += ': ' + subtitle;
  var names = function(list) { return (Array.isArray(list) ? list : []).map(function(v) { return clean(typeof v === 'string' ? v : v.name, 100); }).filter(Boolean).join(', '); };
  var author = names(book.authors), publisher = names(book.publishers), date = clean(book.publish_date, 60);
  var pages = Number(book.number_of_pages), format = clean(book.physical_format, 60);
  var details = [author ? 'Author: ' + author : '', publisher ? 'Publisher: ' + publisher : '', date ? 'Published: ' + date : '', format, Number.isSafeInteger(pages) && pages > 0 ? pages + ' pages' : '', 'ISBN: ' + isbn].filter(Boolean).join(' · ');
  return {title: title, name: title.slice(0,160), manufacturer: publisher.slice(0,100), catalogNumber: isbn, description: details, size: format, sourceUrl: source, lookupProvider: source.indexOf('openlibrary.org') >= 0 ? 'Open Library' : 'Verified publisher record', unitsPerPack: 1, packs: 1, unit: 'each', kind: 'Reusable', category: 'Other', packBasis: 'Suggested counting unit: one book per scanned copy. Verify edition and number of copies received.', notes: (details + '. ' + basis + '. Source: ' + source).slice(0,500)};
}
function eahsVerifiedBook_(isbn) {
  // Publisher-verified fallback, not a live API response or a replacement-price quote.
  // Exact edition checked 2026-10-06; it also works when a catalog has no record.
  if (isbn !== '9781784705954') return null;
  return eahsBookProduct_(isbn, {title: 'The Inner Life of Animals', subtitle: 'Surprising Observations of a Hidden World', authors: [{name: 'Peter Wohlleben'}], publishers: [{name: 'Vintage'}], publish_date: '1 March 2018', physical_format: 'Paperback', number_of_pages: 288}, 'https://www.penguin.co.uk/books/433826/the-inner-life-of-animals-by-peter-wohlleben/9781784705954', 'Publisher-verified fallback record checked 2026-10-06');
}
function eahsISBN10to13_(raw) {
  var code = String(raw || '').replace(/[-\s]/g, '').toUpperCase();
  if (!/^\d{9}[\dX]$/.test(code)) return '';
  var sum = 0;
  for (var i = 0; i < 10; i++) sum += (10 - i) * (code[i] === 'X' ? 10 : Number(code[i]));
  if (sum % 11) return '';
  var base = '978' + code.slice(0,9), check = 0;
  for (var j = 0; j < 12; j++) check += Number(base[j]) * (j % 2 ? 3 : 1);
  return base + String((10 - check % 10) % 10);
}
function eahsBookBarcode_(isbn) {
  var fallback = eahsVerifiedBook_(isbn);
  var fail = function(status, message) { return fallback ? {products: [fallback], lookupStatus: 'found', lookupProvider: 'Verified publisher fallback'} : {products: [], lookupStatus: status, message: message}; };
  var cache = CacheService.getScriptCache(), key = 'EAHS_BOOK_LOOKUP_V1_' + isbn, cached = cache.get(key);
  if (cached) return JSON.parse(cached);
  // Reserve one request slot (brief lock), then fetch with no lock held so inventory saves are never blocked.
  if (!eahsProviderReservation_('OPENLIBRARY', 1100)) return fail('rate_limit', 'Book lookup is busy or temporarily rate limited. Wait a moment and scan this ISBN again.');
  try {
    // One human-triggered ISBN lookup, cached; free public catalog, no key or paid API.
    var url = 'https://openlibrary.org/api/books?bibkeys=' + encodeURIComponent('ISBN:' + isbn) + '&jscmd=data&format=json';
    var response = UrlFetchApp.fetch(url, {method: 'get', headers: {Accept: 'application/json', 'User-Agent': 'EAHS-Science-Inventory/1.1 (human-triggered ISBN lookup)'}, muteHttpExceptions: true, followRedirects: false});
    var status = response.getResponseCode();
    if (status === 429) { eahsProviderBackoff_('OPENLIBRARY'); return fail('rate_limit', 'The free book catalog reached its lookup limit. Try later or enter the details manually.'); }
    if (status !== 200 && status !== 404) return fail('unavailable', 'Book lookup is unavailable (HTTP ' + status + '). Try later or enter details manually.');
    var body = status === 200 ? JSON.parse(response.getContentText()) : {}, book = body['ISBN:' + isbn], products = [];
    if (book && typeof book.title === 'string' && book.title.trim()) {
      var ids = book.identifiers || {};
      var exact = (Array.isArray(ids.isbn_13) ? ids.isbn_13 : []).some(function(v) { try { return eahsBarcode_(String(v).replace(/[-\s]/g,'')) === '0' + isbn; } catch (e) { return false; } }) || (Array.isArray(ids.isbn_10) ? ids.isbn_10 : []).some(function(v) { return eahsISBN10to13_(v) === isbn; });
      // A title match alone cannot establish the edition; returned ISBN must agree.
      if (exact) products.push(eahsBookProduct_(isbn, book, 'https://openlibrary.org/isbn/' + isbn, 'Open Library exact ISBN suggestion'));
    }
    if (!products.length && fallback) return fail('not_found', '');
    var result = {products: products, lookupProvider: 'Open Library', lookupStatus: products.length ? 'found' : 'not_found', message: products.length ? '' : 'The book catalog has no verified match for this exact ISBN.'};
    cache.put(key, JSON.stringify(result), products.length ? 21600 : 300);
    return result;
  } catch (e) {
    return fail('unavailable', 'Book lookup could not connect. Check Apps Script external-service permission, then try again. Manual entry is available.');
  }
}
function eahsUPCBarcode_(raw) {
  var code = eahsBarcode_(raw), cache = CacheService.getScriptCache(), cacheKey = 'EAHS_UPC_LOOKUP_V2_' + code;
  var cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);
  // Free trial endpoint: at most one request per 10 seconds. The lock is held only while reserving the slot.
  if (!eahsProviderReservation_('UPCITEMDB', 10000)) return {products: [], lookupStatus: 'rate_limit', message: 'Online lookup is busy or temporarily rate limited. Wait a moment and scan again, or enter the details manually.'};
  try {
    // Free-only endpoint: no keys, accounts, subscriptions, paid fallback or charges.
    var response = UrlFetchApp.fetch('https://api.upcitemdb.com/prod/trial/lookup?upc=' + encodeURIComponent(String(raw).trim()), {method: 'get', headers: {Accept: 'application/json'}, muteHttpExceptions: true, followRedirects: false});
    var status = response.getResponseCode();
    if (status === 429) {
      eahsProviderBackoff_('UPCITEMDB');
      return {products: [], lookupStatus: 'rate_limit', message: 'The free product database reached its lookup limit. Try later or enter the product details manually.'};
    }
    if (status !== 200 && status !== 404) return {products: [], lookupStatus: 'unavailable', message: 'Online product lookup is unavailable (HTTP ' + status + '). Try again later or enter details manually.'};
    var body = JSON.parse(response.getContentText());
    if (status === 200 && body.code !== 'OK') return {products: [], lookupStatus: 'unavailable', message: 'The online product database did not return a valid lookup result. Try later or enter details manually.'};
    var products = (Array.isArray(body.items) ? body.items : []).filter(function(product) {
      return typeof product.title === 'string' && product.title.trim() && ['gtin','ean','upc'].some(function(key) {
        try { return eahsBarcode_(product[key]) === code; } catch (e) { return false; }
      });
    }).slice(0,5).map(function(product) {
      var clean = function(value, max) { return String(value || '').replace(/[\u0000-\u001f]/g, ' ').trim().slice(0,max); };
      var title = clean(product.title, 350);
      [clean(product.size,80),clean(product.color,50)].filter(Boolean).forEach(function(detail) { if (title.toLowerCase().indexOf(detail.toLowerCase()) < 0) title += ' — ' + detail; });
      var source = 'https://www.upcitemdb.com/upc/' + encodeURIComponent(String(raw).trim());
      return eahsProductSuggestion_({title: title, manufacturer: clean(product.brand, 100), catalogNumber: clean(product.model, 80), size: clean(product.size, 100), color: clean(product.color, 100), description: clean(product.description, 500), sourceUrl: source, lookupProvider: 'UPCitemdb'});
    });
    var result = {products: products, lookupStatus: products.length ? 'found' : 'not_found', message: products.length ? '' : 'No online product matched this exact barcode. Enter the details manually or scan another barcode.'};
    cache.put(cacheKey, JSON.stringify(result), products.length ? 21600 : 300);
    return result;
  } catch (e) {
    return {products: [], lookupStatus: 'unavailable', message: 'Online lookup could not connect. Verify Apps Script has permission to connect to external services, then try again. Manual entry is still available.'};
  }
}
// Count one complete scanned package unless staff explicitly choose another unit.
// A title's “100 count” may describe gloves, sheets, pairs, or nested packages.
function eahsProductSuggestion_(p) {
  var clean = eahsBarcodeText_, title = clean(p.title, 500);
  var consumable = p.kind === 'Consumable' || /\b(?:gloves?|salt|sodium chloride|paper towels?|tissues?|disposable|pipette tips?|cotton|swabs?|soap|saniti[sz]er|detergent|bleach|batteries|adhesive tape|trash bags?)\b/i.test(title);
  var unit = /\bbox(?:es)?\b/i.test(title) ? 'box' : /\b(?:pack|count)\b/i.test(title) ? 'pack' : 'each';
  return {title: title, name: title.slice(0,160), manufacturer: clean(p.manufacturer,100), catalogNumber: clean(p.catalogNumber,80), size: clean(p.size,100), color: clean(p.color,100), description: clean(p.description,500), sourceUrl: p.sourceUrl, lookupProvider: p.lookupProvider,
    unitsPerPack: 1, packs: 1, unit: unit, kind: consumable ? 'Consumable' : 'Reusable', category: 'Other',
    packBasis: 'Suggested count: one complete scanned package = 1 ' + unit + '. To count individual contents instead, change the unit and units per pack after checking the label. Weight/volume and count in the title are not a stock multiplier.',
    notes: clean('Product suggestion: ' + title + '. ' + (p.description || '') + ' Source: ' + p.lookupProvider + ' — ' + p.sourceUrl,500)};
}
function eahsOnlineBarcode_(raw) {
  var code = eahsBarcode_(raw), cache = CacheService.getScriptCache(), key = 'EAHS_MULTI_BARCODE_V1_' + code;
  var cached = cache.get(key); if (cached) return JSON.parse(cached);
  // Known inventory is handled before this function. A successful broad database
  // lookup stops the chain; another five databases are available on a miss/failure.
  var primary = eahsUPCBarcode_(raw);
  if (primary.products.length) { primary.lookupProvider = 'UPCitemdb'; cache.put(key,JSON.stringify(primary),21600); return primary; }
  // ISBN records belong to book catalogs, not the food/medical fallback.
  if (/^097[89]/.test(code)) return primary;
  var extra = eahsSupplementalBarcode_(raw);
  if (extra.products.length) { cache.put(key,JSON.stringify(extra),21600); return extra; }
  var results = [primary].concat(extra.providerResults || []);
  var failed = results.some(function(r) { return r.lookupStatus !== 'not_found'; });
  var result = {products: [], lookupStatus: failed ? 'partial_unavailable' : 'not_found',
    message: failed ? 'Some product databases are temporarily unavailable or rate limited. No exact match was returned by the available sources. Retry later or verify the product details manually; this barcode will be remembered after you confirm and save.' : 'No exact barcode match was found in the available product databases. Enter and verify the label details once; this barcode will be remembered after you confirm and save.'};
  cache.put(key,JSON.stringify(result),failed ? 30 : 300); return result;
}
function eahsLookupRequest_(url) {
  return {url: url, method: 'get', headers: {Accept: 'application/json', 'User-Agent': 'EAHS-Science-Inventory/1.2 (human-triggered barcode lookup)'}, muteHttpExceptions: true, followRedirects: false};
}
function eahsFetchOne_(url) {
  var options = eahsLookupRequest_(url); delete options.url;
  return UrlFetchApp.fetch(url,options);
}
function eahsProviderReservation_(provider, interval) {
  var lock = LockService.getScriptLock(); if (!lock.tryLock(1000)) return false;
  try {
    var props = PropertiesService.getScriptProperties(), key = 'EAHS_PROVIDER_AFTER_' + provider, now = Date.now();
    if (Number(props.getProperty(key) || 0) > now) return false;
    props.setProperty(key,String(now + interval)); return true;
  } finally { lock.releaseLock(); }
}
function eahsProviderBackoff_(provider) {
  PropertiesService.getScriptProperties().setProperty('EAHS_PROVIDER_AFTER_' + provider,String(Date.now()+60000));
}
function eahsSupplementalBarcode_(raw) {
  var code = eahsBarcode_(raw), digits = String(raw).trim();
  if (digits.length === 12) digits = '0' + digits;
  if (digits.length === 14 && digits[0] === '0') digits = digits.slice(1);
  var cache = CacheService.getScriptCache(), jobs = [
    {provider:'OPEN_FACTS', interval:8500, url:'https://world.openfoodfacts.org/api/v3/product/' + digits + '?product_type=all&fields=code,product_name,product_name_en,brands,quantity,product_type,generic_name,generic_name_en,packaging_text,ingredients_text'},
    {provider:'GUDID', interval:1100, url:'https://accessgudid.nlm.nih.gov/api/v3/devices/lookup.json?di=' + code}
  ], pending = [], results = [];
  jobs.forEach(function(job) {
    job.key = 'EAHS_PROVIDER_V1_' + job.provider + '_' + code;
    var cached = cache.get(job.key);
    if (cached) results.push(JSON.parse(cached));
    else if (eahsProviderReservation_(job.provider,job.interval)) pending.push(job);
    else results.push({products:[],lookupStatus:'rate_limit'});
  });
  var responses = [];
  if (pending.length) {
    try { responses = UrlFetchApp.fetchAll(pending.map(function(job) { return eahsLookupRequest_(job.url); })); }
    catch (e) {
      // A transport failure for one host must not lose the other provider.
      responses = pending.map(function(job) { try { return eahsFetchOne_(job.url); } catch (e) { return null; } });
    }
  }
  pending.forEach(function(job,index) {
    var response = responses[index], result;
    try {
      if (job.provider === 'OPEN_FACTS') response = eahsOpenFactsRedirect_(response);
      if (!response) throw new Error('Unavailable');
      var status = response.getResponseCode();
      if (status === 429 || status === 503) { eahsProviderBackoff_(job.provider); result = {products:[],lookupStatus:'rate_limit'}; }
      else if (status === 404) result = {products:[],lookupStatus:'not_found'};
      else if (status !== 200) result = {products:[],lookupStatus:'unavailable'};
      else {
        var body = JSON.parse(response.getContentText());
        var products = job.provider === 'OPEN_FACTS' ? eahsOpenFactsProduct_(body,code) : eahsGUDIDProduct_(body,code);
        result = {products:products,lookupStatus:products.length ? 'found' : 'not_found'};
      }
    } catch (e) { result = {products:[],lookupStatus:'unavailable'}; }
    cache.put(job.key,JSON.stringify(result),result.lookupStatus === 'found' ? 21600 : result.lookupStatus === 'not_found' ? 300 : 30);
    results.push(result);
  });
  // Keep differing records separate. Staff choose the exact variant; descriptions
  // and pack counts from separate sources are never silently merged.
  var products = results.reduce(function(all,r) { return all.concat(r.products); },[]);
  return {products:products,lookupStatus:products.length ? 'found' : 'not_found',providerResults:results};
}
function eahsOpenFactsRedirect_(response) {
  // The documented universal API routes to one of four Open Facts instances.
  // Only follow HTTPS product endpoints on those hosts, with no arbitrary URLs.
  for (var i = 0; response && i < 1; i++) {
    if ([301,302,303,307,308].indexOf(response.getResponseCode()) < 0) break;
    var headers = response.getAllHeaders(), location = headers.Location || headers.location;
    if (Array.isArray(location)) location = location[0];
    if (!/^https:\/\/world\.open(?:food|beauty|petfood|products)facts\.org\/api\/v3(?:\.\d+)?\/product\/\d{8,14}(?:\.json)?(?:\?[^\s]*)?$/.test(String(location || ''))) throw new Error('Untrusted provider redirect');
    response = eahsFetchOne_(location);
  }
  return response;
}
function eahsOpenFactsProduct_(body,code) {
  var p = body && body.product;
  if (!p || eahsBarcode_(p.code) !== code) return [];
  var clean = eahsBarcodeText_, title = clean(p.product_name_en || p.product_name,300);
  if (!title) return [];
  var types = {food:['Open Food Facts','food'],beauty:['Open Beauty Facts','beauty'],petfood:['Open Pet Food Facts','petfood'],product:['Open Products Facts','products']};
  var type = types[p.product_type || body.product_type || 'food']; if (!type) return [];
  var size = clean(p.quantity,100), description = clean([p.generic_name_en || p.generic_name,p.packaging_text,p.ingredients_text].filter(Boolean).join(' · '),500);
  // Product names often omit net contents. Preserve returned quantity in the draft.
  if (size && title.toLowerCase().indexOf(size.toLowerCase()) < 0) title += ' — ' + size;
  return [eahsProductSuggestion_({title:title,manufacturer:clean(p.brands,100),size:size,description:description,kind: type[1] === 'products' ? '' : 'Consumable',sourceUrl:'https://world.open' + type[1] + 'facts.org/product/' + encodeURIComponent(p.code),lookupProvider:type[0] + ' (ODbL)'})];
}
function eahsGUDIDProduct_(body,code) {
  var d = body && body.gudid && body.gudid.device; if (!d) return [];
  var identifiers = d.identifiers && d.identifiers.identifier;
  if (!Array.isArray(identifiers)) identifiers = identifiers ? [identifiers] : [];
  var exact = identifiers.find(function(id) { try { return id.deviceIdIssuingAgency === 'GS1' && eahsBarcode_(id.deviceId) === code; } catch (e) { return false; } });
  if (!exact) return [];
  var clean = eahsBarcodeText_, description = clean(d.deviceDescription,500), model = clean(d.catalogNumber || d.versionModelNumber,80);
  var title = clean(description || d.brandName,300); if (!title) return [];
  if (model && title.toLowerCase().indexOf(model.toLowerCase()) < 0) title += ' — ' + model;
  var contents = exact.pkgType ? 'Scanned packaging type: ' + clean(exact.pkgType,60) + '. ' : '';
  // Package quantity may count lower-level packages, not gloves or pairs.
  if (exact.pkgQuantity != null) contents += 'Reported package quantity: ' + clean(exact.pkgQuantity,30) + ' (verify packaging hierarchy). ';
  if (d.deviceCommDistributionStatus) contents += 'Distribution status: ' + clean(d.deviceCommDistributionStatus,80) + '. ';
  return [eahsProductSuggestion_({title:title,manufacturer:clean(d.companyName,100),catalogNumber:model,description:contents+description,kind:d.singleUse === true ? 'Consumable' : '',sourceUrl:'https://accessgudid.nlm.nih.gov/devices/' + encodeURIComponent(exact.deviceId),lookupProvider:'FDA/NLM AccessGUDID'})];
}

// Run once from the editor after installing to authorize the new external-fetch scope.
function authorizeBarcodeOnlineLookup() {
  var book = eahsBookBarcode_('9781784705954');
  var product = eahsOnlineBarcode_('615435163246');
  return 'Book lookup: ' + book.lookupStatus + ' (' + (book.lookupProvider || 'catalog') + '); product lookup: ' + product.lookupStatus;
}
function saveBarcodeInventory(p) {
  // UserLock is deliberately separate from the stock functions' ScriptLock.
  var lock = LockService.getUserLock(); lock.waitLock(30000);
  try {
    if (!p || p.confirmed !== true || !/^[a-zA-Z0-9-]{10,100}$/.test(p.requestId || '')) throw new Error('Review and confirm this scan before saving.');
    var props = PropertiesService.getUserProperties(), requestKey = 'EAHS_BARCODE_REQUEST_' + p.requestId;
    var prior = props.getProperty(requestKey);
    if (prior === 'Done') return {data: getData(), warning: 'This confirmed scan was already saved. Stock was not added again.'};
    if (prior) throw new Error('This save may already have changed stock. Refresh and check Activity before another addition.');
    var code = eahsBarcode_(p.barcode), packs = Number(p.packs), units = Number(p.unitsPerPack), quantity = packs * units;
    if (![packs, units, quantity].every(function(v) { return Number.isSafeInteger(v) && v > 0; })) throw new Error('Pack count and counted units per pack must be positive whole numbers.');
    // Only the Items tab is needed for these checks.
    var draft = p.item || {}, before = {items: rows_('Items').map(strip_)}, existing = before.items.find(function(i) { return i.id === p.itemId; });
    if (!String(draft.name || '').trim() || !String(draft.unit || '').trim() || !String(draft.actor || '').trim()) throw new Error('Item name, counted unit and staff member are required.');
    if (!Number.isSafeInteger(Number(draft.minStock)) || Number(draft.minStock) < 0) throw new Error('Reorder level must be a nonnegative whole number.');
    var data;
    if (p.mode === 'stock') {
      if (!existing) throw new Error('That inventory entry no longer exists. Scan again.');
      if (['name','unit','kind','location','storage','minStock','manufacturer','catalogNumber'].some(function(k) { return String(existing[k] == null ? '' : existing[k]) !== String(draft[k] == null ? '' : draft[k]); })) throw new Error('Item details changed while you reviewed. Scan again to confirm the current entry.');
      props.setProperty(requestKey, 'Saving');
      data = receiveStock({itemId: existing.id, quantity: quantity, actor: draft.actor, notes: 'Confirmed product barcode ' + String(p.barcode).trim() + ': ' + packs + ' pack(s) x ' + units + ' ' + draft.unit + ' per pack.'});
    } else if (p.mode === 'new') {
      var sameName = function(a,b) { return typeof nameSignature_ === 'function' ? nameSignature_(a) === nameSignature_(b) : String(a).trim().toLowerCase() === String(b).trim().toLowerCase(); };
      if (before.items.some(function(i) { return sameName(i.name, draft.name) && i.kind === draft.kind && String(i.location) === String(draft.location); })) throw new Error('This item already has an entry in that room. Use Add stock, or choose a different home location or a distinct model name.');
      props.setProperty(requestKey, 'Saving');
      data = saveItem(Object.assign({}, draft, {id: '', sku: '', quantity: quantity}));
      var added = data.items.filter(function(i) { return !before.items.some(function(old) { return old.id === i.id; }); });
      if (added.length !== 1) { props.setProperty(requestKey, 'Done'); return {data: data, warning: 'Stock was saved, but the new entry could not be identified for barcode linking. Check Activity; do not add the stock again.'}; }
      existing = added[0];
    } else throw new Error('Choose Add stock or Separate entry.');
    props.setProperty(requestKey, 'Done');
    var warning = '';
    try {
      var linkLock = LockService.getScriptLock(); linkLock.waitLock(30000);
      try {
        var links = eahsBarcodeLinks_(code).filter(function(link) { return link.itemId !== existing.id; });
        links.push({itemId: existing.id, unit: existing.unit, unitsPerPack: units});
        PropertiesService.getScriptProperties().setProperty('EAHS_PRODUCT_BARCODE_' + code, JSON.stringify(links));
      } finally { linkLock.releaseLock(); }
    } catch (e) { warning = 'Stock was saved, but barcode linking failed: ' + e.message + '. Do not add this stock again.'; }
    // Bound user-property storage. Keep recent receipts to suppress duplicate submissions.
    var receipts = JSON.parse(props.getProperty('EAHS_BARCODE_RECENT') || '[]');
    receipts.push(requestKey); while (receipts.length > 100) props.deleteProperty(receipts.shift());
    props.setProperty('EAHS_BARCODE_RECENT', JSON.stringify(receipts));
    return {data: data, warning: warning};
  } finally { lock.releaseLock(); }
}
