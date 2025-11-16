// Code.gs

const PRICING_SHEET_ADSELETO = 'Precificação - AdSeleto';
const PRICING_SHEET_IGOAL = 'Precificação - Igoal';
const BASE_SHEET_CANDIDATES = ['Base de Dados','Base de Dados - Exemplo','BD - GAM','BD – GAM','BD - GAM1'];
const AUTO_PRICING_DATA_SHEETS = ['BD - GAM','BD – GAM'];
const HEADER_ALIASES = {
  date: ['Data'],
  hour: ['Hora'],
  site: ['Site'],
  channel: ['Canal', 'utm_source', 'UTM Source'],
  url: ['URL'],
  adunit: ['Bloco de Anúncio', 'Bloco', 'Ad Unit'],
  impressions: ['Impressões', 'Impressoes'],
  clicks: ['Cliques'],
  ctr: ['CTR'],
  revenue: ['Receita (USD)', 'Receita', 'Receita Total', 'Revenue'],
  ecpm: ['eCPM', 'ECPM'],
  requests: ['Solicitação AD', 'Solicitações', 'Solicitacoes', 'Requests'],
  coverage: ['Cobertura', 'Cobertura (%)', 'Cobertura %', 'Coverage'],
  cpc: ['CPC'],
  viewability: ['Viewability']
};

const AUTO_PRICING_ALLOWED_ADUNIT_TOKENS = ['mob_top','mob_offerwall','mob_rewarded','desk_top','desk_offerwall','desk_rewarded'];
const AUTO_PRICING_CONFIG_SHEET = 'Config - Precificação';
const AUTO_PRICING_LOG_SHEET = 'Log - Precificação';
const AUTO_PRICING_INTERVAL_HOURS = 3;
const AUTO_PRICING_WINDOW_HOURS = 3;
const AUTO_PRICING_MIN_REQUESTS = 10;
const AUTO_PRICING_RANGE_HEADER = ['Cobertura Superior','Cobertura Inferior','Coeficiente'];
const AUTO_PRICING_SITE_HEADER = ['Site','utm_source (separe por ;)','Rede','Company ID / PMD ID','Ativo?'];
const AUTO_PRICING_BUCKET_HEADER = ['Site','Buckets disponíveis (JSON ou separados por ;)'];
const AUTO_PRICING_LOG_HEADERS = ['Timestamp','Disparo','Site','Rede','UTM Source','URL','Regra Atual','Nova Regra','Cobertura (%)','eCPM','Coeficiente','Bucket Aplicado','Solicitações','Observações','Payload enviado'];

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Dashboard de Operações')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getBaseSheet_() {
  const ss = SpreadsheetApp.getActive();
  for (var i = 0; i < BASE_SHEET_CANDIDATES.length; i++) {
    var sh = ss.getSheetByName(BASE_SHEET_CANDIDATES[i]);
    if (sh) return sh;
  }
  return ss.getSheets()[0];
}

function getAutoPricingDataSheet_() {
  const ss = SpreadsheetApp.getActive();
  for (var i = 0; i < AUTO_PRICING_DATA_SHEETS.length; i++) {
    var sh = ss.getSheetByName(AUTO_PRICING_DATA_SHEETS[i]);
    if (sh) return sh;
  }
  return getBaseSheet_();
}

function ensurePricingSheetAdSeleto_() {
  const ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(PRICING_SHEET_ADSELETO);
  if (!sh) {
    sh = ss.insertSheet(PRICING_SHEET_ADSELETO);
    sh.getRange(1, 1, 1, 7).setValues([
      ['spnprice_id', 'domain_id', 'price_rule', 'utm_source', 'slot_id', 'url', 'Sincronizar']
    ]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function ensurePricingSheetIgoal_() {
  const ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(PRICING_SHEET_IGOAL);
  if (!sh) {
    sh = ss.insertSheet(PRICING_SHEET_IGOAL);
    sh.getRange(1, 1, 1, 8).setValues([
      ['rule_id', 'dominio', 'utm_source', 'url', 'bloco', 'company_id', 'price_rule', 'Sincronizar']
    ]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function ensureAutoPricingConfigSheet_() {
  const ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(AUTO_PRICING_CONFIG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(AUTO_PRICING_CONFIG_SHEET);
    var defaultRows = [];
    var minCols = Math.max(
      AUTO_PRICING_RANGE_HEADER.length,
      AUTO_PRICING_SITE_HEADER.length,
      AUTO_PRICING_BUCKET_HEADER.length
    );
    function pushRow(arr) {
      var row = (arr || []).slice();
      while (row.length < minCols) row.push('');
      defaultRows.push(row);
    }
    pushRow(AUTO_PRICING_RANGE_HEADER);
    [[100,95,1.30],[95,90,1.20],[90,85,1.15],[85,78,1.10],[78,72,1.00],[72,65,0.90],[65,60,0.85],[60,55,0.80],[55,50,0.75],[50,45,0.70],[45,40,0.65],[40,35,0.60],[35,30,0.55],[30,0,0.50]].forEach(function(row){ pushRow(row); });
    pushRow([]);
    pushRow(AUTO_PRICING_SITE_HEADER);
    pushRow(['exemplo.com','utm_a;utm_b','Igoal','', 'TRUE']);
    pushRow([]);
    pushRow(AUTO_PRICING_BUCKET_HEADER);
    pushRow(['exemplo.com','["0.13","0.14","0.15","0.16"]']);
    if (defaultRows.length) {
      sh.getRange(1,1,defaultRows.length,defaultRows[0].length).setValues(defaultRows);
    }
    sh.setFrozenRows(1);
  }
  return sh;
}

function ensureAutoPricingLogSheet_() {
  const ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(AUTO_PRICING_LOG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(AUTO_PRICING_LOG_SHEET);
  }
  var neededCols = AUTO_PRICING_LOG_HEADERS.length;
  if (sh.getMaxColumns() < neededCols) {
    sh.insertColumnsAfter(sh.getMaxColumns(), neededCols - sh.getMaxColumns());
  }
  sh.getRange(1, 1, 1, neededCols).setValues([AUTO_PRICING_LOG_HEADERS]);
  sh.setFrozenRows(1);
  return sh;
}

function getHeaderIndexMap_(headers) {
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    var raw = String(headers[i] == null ? '' : headers[i]).trim();
    if (!raw) continue;
    if (!(raw in map)) {
      map[raw] = i;
    }
    var lower = raw.toLowerCase();
    if (!(lower in map)) {
      map[lower] = i;
    }
    if (typeof normalizeHeaderName_ === 'function') {
      var normalized = normalizeHeaderName_(raw);
      if (normalized && !(normalized in map)) {
        map[normalized] = i;
      }
    }
  }
  return map;
}

function normalizeHeaderName_(name) {
  return String(name == null ? '' : name)
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

function mapHeadersByAlias_(headers, aliasMap) {
  var normalizedHeaders = headers.map(function(name){ return normalizeHeaderName_(name); });
  var result = {};
  Object.keys(aliasMap || {}).forEach(function(key){
    var aliases = aliasMap[key];
    if (!Array.isArray(aliases)) aliases = [aliases];
    var normalizedAliases = aliases.map(function(name){ return normalizeHeaderName_(name); });
    for (var i = 0; i < normalizedHeaders.length; i++) {
      if (normalizedAliases.indexOf(normalizedHeaders[i]) !== -1) {
        result[key] = i;
        break;
      }
    }
  });
  return result;
}

function toNumber_(val) {
  if (typeof val === 'number') return val;
  if (val == null) return 0;
  var s = String(val).trim();
  if (!s) return 0;
  var sanitized = s.replace(/\s/g, '').replace(/%/g, '');
  var hasComma = sanitized.indexOf(',') !== -1;
  var hasDot = sanitized.indexOf('.') !== -1;
  if (hasComma && hasDot) {
    if (sanitized.lastIndexOf(',') > sanitized.lastIndexOf('.')) {
      sanitized = sanitized.replace(/\./g, '').replace(/,/g, '.');
    } else {
      sanitized = sanitized.replace(/,/g, '');
    }
  } else if (hasComma) {
    sanitized = sanitized.replace(/\./g, '').replace(/,/g, '.');
  } else {
    sanitized = sanitized.replace(/,/g, '');
  }
  var n = parseFloat(sanitized);
  return isNaN(n) ? 0 : n;
}

function parseCoverageValue_(val) {
  if (val == null) return 0;
  if (typeof val === 'number') {
    if (!isFinite(val)) return 0;
    if (val >= 0 && val <= 1) return val * 100;
    return val;
  }
  var raw = String(val).trim();
  if (!raw) return 0;
  var hasPercent = raw.indexOf('%') !== -1;
  var num = toNumber_(raw);
  if (!isFinite(num)) return 0;
  if (hasPercent) return num;
  if (num >= 0 && num <= 1) return num * 100;
  return num;
}

function formatPriceRuleIgoal_(value) {
  if (typeof normalizeIgoalPrice_ === 'function') {
    var normalized = normalizeIgoalPrice_(value);
    return normalized != null ? String(normalized) : '';
  }
  if (value === null || value === undefined) return '';
  var s = String(value).trim();
  if (!s) return '';
  s = s.replace(/,/g, '.');
  var n = parseFloat(s);
  if (isNaN(n)) return '';
  return n.toFixed(2);
}

function formatPriceRuleAdSeleto_(value) {
  if (typeof normalizePmdPrice_ === 'function') {
    var normalized = normalizePmdPrice_(value);
    if (normalized == null || isNaN(normalized)) return '';
    return Number(normalized).toFixed(2);
  }
  if (value === null || value === undefined) return '';
  var s = String(value).trim();
  if (!s) return '';
  s = s.replace(/,/g, '.');
  var n = parseFloat(s);
  if (isNaN(n)) return '';
  return n.toFixed(2);
}

function normUrlStrict_(u) {
  if (!u) return '';
  var s = String(u).trim();
  if (!s) return '';
  s = s.replace(/^https?:\/\//i, '');
  s = s.replace(/^www\./i, '');
  s = s.split('#')[0];
  s = s.split('?')[0];
  s = s.replace(/\\/g, '/');
  s = s.replace(/\/+/g, '/');
  var slashIndex = s.indexOf('/');
  var slug = slashIndex >= 0 ? s.substring(slashIndex + 1) : s;
  slug = slug.replace(/^\/+/, '').replace(/\/+$/, '');
  if (!slug) {
    // Se a URL representar apenas o domínio, não enviamos o domínio completo.
    return '';
  }
  if (slashIndex < 0 && /\./.test(slug)) {
    // Valores que ainda parecem um domínio inteiro são descartados.
    return '';
  }
  return slug.trim().toLowerCase();
}

function normSite_(val) {
  return String(val == null ? '' : val).trim().toLowerCase();
}

function normalizeIgoalUtmSourceValue_(value) {
  if (value === undefined || value === null) return '';
  var str = String(value).trim();
  if (!str) return '';
  var lower = str.toLowerCase();
  var markerIndex = lower.indexOf('utm_source');
  if (markerIndex !== -1) {
    if (markerIndex > 0) {
      var prevChar = lower.charAt(markerIndex - 1);
      if (['?', '&', '#', '='].indexOf(prevChar) === -1 && !/\s/.test(prevChar)) {
        return str;
      }
    }
    var slice = str.substring(markerIndex + 'utm_source'.length);
    slice = slice.replace(/^[^a-z0-9]+/i, '');
    var stop = slice.search(/[&#]/);
    if (stop !== -1) {
      slice = slice.substring(0, stop);
    }
    str = slice.trim();
  }
  return str;
}

function normUtm_(val) {
  return String(val == null ? '' : val).trim().toLowerCase();
}

function normNetwork_(val) {
  return String(val == null ? '' : val).trim().toLowerCase();
}

function normAdUnit_(val) {
  return String(val == null ? '' : val).trim().toLowerCase();
}

function generateSpnpriceId_() {
  var uuid = Utilities.getUuid();
  if (!uuid) {
    return 'auto-' + Math.floor(Math.random() * 1e9);
  }
  return 'auto-' + uuid.replace(/-/g, '').slice(0, 16).toLowerCase();
}

function coalesceForPayload_() {
  for (var i = 0; i < arguments.length; i++) {
    var value = arguments[i];
    if (value === undefined || value === null) continue;
    if (typeof value === 'number') {
      if (!isNaN(value)) return value;
      continue;
    }
    var str = String(value).trim();
    if (str) return str;
  }
  return '';
}

function buildPricingPayloadLog_(networkKey, entry, ruleInfo, formattedRuleValue, rawRuleValue) {
  networkKey = normNetwork_(networkKey);
  entry = entry || {};
  ruleInfo = ruleInfo || {};
  var slugFromRule = ruleInfo.url ? normUrlStrict_(ruleInfo.url) : '';
  var slugFromEntry = entry.normUrl || '';
  var slugFallback = entry.url ? normUrlStrict_(entry.url) : '';
  var priceRuleValue = rawRuleValue;
  if (!isFinite(priceRuleValue)) {
    var parsed = toNumber_(formattedRuleValue);
    if (isFinite(parsed)) {
      priceRuleValue = parsed;
    } else {
      priceRuleValue = formattedRuleValue != null ? formattedRuleValue : '';
    }
  }
  if (isFinite(priceRuleValue)) {
    priceRuleValue = Math.round(priceRuleValue * 100) / 100;
  }
  var payload;
  if (networkKey === 'adseleto') {
    payload = {
      domain_id: coalesceForPayload_(ruleInfo.domain_id, entry.domain_id, entry.identifier, entry.site),
      price_rule: priceRuleValue,
      utm_source: coalesceForPayload_(ruleInfo.utm_source, entry.utm_source),
      slot_id: coalesceForPayload_(ruleInfo.slot_id, entry.slot_id, entry.adunit),
      url: coalesceForPayload_(slugFromRule, slugFromEntry, slugFallback),
      spnprice_id: coalesceForPayload_(ruleInfo.spnprice_id, entry.spnprice_id)
    };
  } else if (networkKey === 'igoal') {
    var utmForIgoal = normalizeIgoalUtmSourceValue_(coalesceForPayload_(ruleInfo.utm_source, entry.utm_source));
    payload = {
      dominio: coalesceForPayload_(ruleInfo.dominio, ruleInfo.domain_id, entry.domain_id, entry.site),
      utm_source: utmForIgoal,
      url: coalesceForPayload_(slugFromRule, slugFromEntry, slugFallback),
      bloco: coalesceForPayload_(ruleInfo.bloco, ruleInfo.slot_id, entry.bloco, entry.adunit),
      company_id: coalesceForPayload_(ruleInfo.company_id, entry.company_id, entry.identifier),
      price_rule: priceRuleValue
    };
  } else {
    return '';
  }

  var normalizedPayload = {};
  Object.keys(payload).forEach(function(key){
    var value = payload[key];
    if (value === undefined || value === null) {
      normalizedPayload[key] = '';
    } else if (typeof value === 'number') {
      normalizedPayload[key] = isFinite(value) ? value : '';
    } else {
      normalizedPayload[key] = String(value).trim();
    }
  });

  try {
    return JSON.stringify(normalizedPayload);
  } catch (err) {
    return '';
  }
}

function matchesAllowedAdUnitTokens_(value) {
  var tokens = AUTO_PRICING_ALLOWED_ADUNIT_TOKENS || [];
  if (!tokens.length) return true;
  var normalized = normAdUnit_(value);
  if (!normalized) return false;
  for (var i = 0; i < tokens.length; i++) {
    if (normalized.indexOf(String(tokens[i] || '').toLowerCase()) !== -1) {
      return true;
    }
  }
  return false;
}

function parseDate_(v) {
  if (v instanceof Date) return v;
  var s = String(v || '').trim();
  if (!s) return null;
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return new Date(+m[1], +m[2]-1, +m[3]);
  m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (m) return new Date(+m[3], +m[2]-1, +m[1]);
  var d = new Date(s);
  return isNaN(d) ? null : d;
}

function parseAutoPricingBoolean_(value) {
  if (value === true) return true;
  if (value === false) return false;
  var s = String(value || '').trim().toLowerCase();
  if (!s) return true;
  if (['false','0','no','não','nao','off','inactive','desligado'].indexOf(s) !== -1) return false;
  return true;
}

function isAutoPricingBooleanToken_(value) {
  if (value === true || value === false) return true;
  var s = String(value || '').trim().toLowerCase();
  if (!s) return false;
  return ['true','false','0','1','sim','não','nao','yes','no','on','off','ativo','inativo','active','inactive','desligado','ligado'].indexOf(s) !== -1;
}

function parseBucketsValue_(value) {
  if (value == null) return [];
  if (Array.isArray(value)) {
    return value.map(function(v){ return String(v || '').trim(); }).filter(function(v){ return !!v; });
  }
  var s = String(value || '').trim();
  if (!s) return [];
  if (s.charAt(0) === '[') {
    try {
      var arr = JSON.parse(s);
      if (Array.isArray(arr)) {
        return arr.map(function(v){ return String(v || '').trim(); }).filter(function(v){ return !!v; });
      }
    } catch (err) {
      // fallback para split manual
    }
  }
  return s.split(/[,;\n]/).map(function(v){ return String(v || '').trim(); }).filter(function(v){ return !!v; });
}

function createNormalizedSet_(arr, normalizer) {
  if (!arr || !arr.length) return null;
  var set = {};
  arr.forEach(function(item){
    var norm = normalizer ? normalizer(item) : String(item || '').trim();
    if (norm) set[norm] = true;
  });
  return Object.keys(set).length ? set : null;
}

function formatHourLabelAuto_(ms) {
  if (ms == null) return '';
  var tz = Session.getScriptTimeZone();
  return Utilities.formatDate(new Date(ms), tz, "dd/MM HH'h'");
}

function chooseBucketValue_(targetValue, bucketEntry) {
  if (!bucketEntry || !bucketEntry.values || !bucketEntry.values.length) {
    return { value: null, label: null };
  }
  var values = bucketEntry.values;
  var labels = bucketEntry.labels;
  var best = null;
  for (var i = 0; i < values.length; i++) {
    var current = values[i];
    if (!isFinite(current)) continue;
    if (targetValue == null || isNaN(targetValue)) {
      if (best == null) best = current;
      continue;
    }
    if (current <= targetValue + 1e-9) {
      if (best == null || current > best) {
        best = current;
      }
    }
  }
  if (best == null) {
    best = values[0];
  }
  var label = null;
  if (labels && labels.length) {
    for (var j = 0; j < values.length; j++) {
      if (Math.abs(values[j] - best) <= 1e-9) {
        label = labels[j];
        break;
      }
    }
  }
  if (label == null) label = best != null ? String(best) : '';
  return { value: best, label: label };
}

function nearlyEqual_(a, b, tolerance) {
  if (!isFinite(a) && !isFinite(b)) return true;
  if (!isFinite(a) || !isFinite(b)) return false;
  var tol = tolerance == null ? 0.0001 : tolerance;
  return Math.abs(a - b) <= tol;
}

function loadAutoPricingConfig_() {
  var sh = ensureAutoPricingConfigSheet_();
  var values = sh.getDataRange().getValues();
  var config = { ranges: [], sites: [], buckets: [] };
  var section = '';
  values.forEach(function(row){
    if (!row) return;
    if (rowMatchesHeader_(row, AUTO_PRICING_RANGE_HEADER)) { section = 'ranges'; return; }
    if (rowMatchesHeader_(row, AUTO_PRICING_SITE_HEADER)) { section = 'sites'; return; }
    if (rowMatchesHeader_(row, AUTO_PRICING_BUCKET_HEADER)) { section = 'buckets'; return; }
    var isEmpty = row.every(function(cell){ return String(cell || '').trim() === ''; });
    if (isEmpty) return;
    if (section === 'ranges') {
      var upper = parseCoverageValue_(row[0]);
      var lower = parseCoverageValue_(row[1]);
      var coefficient = toNumber_(row[2]);
      if (isFinite(upper) && isFinite(lower) && isFinite(coefficient)) {
        config.ranges.push({ upper: upper, lower: lower, coefficient: coefficient });
      }
    } else if (section === 'sites') {
      var site = String(row[0] || '').trim();
      if (!site) return;
      var utmParts = parseBucketsValue_(row[1]);
      var network = '';
      var identifier = '';
      var activeCellValue;
      if (row.length >= 5) {
        network = String(row[2] || '').trim();
        identifier = String(row[3] || '').trim();
        activeCellValue = row[4];
      } else if (row.length >= 4) {
        network = String(row[2] || '').trim();
        if (isAutoPricingBooleanToken_(row[3])) {
          activeCellValue = row[3];
        } else {
          identifier = String(row[3] || '').trim();
        }
      } else if (row.length >= 3) {
        network = String(row[2] || '').trim();
      } else {
        activeCellValue = row[2];
      }
      var active = parseAutoPricingBoolean_(activeCellValue);
      config.sites.push({ site: site, utmSources: utmParts, network: network, identifier: identifier, active: active });
    } else if (section === 'buckets') {
      var bucketSite = String(row[0] || '').trim();
      if (!bucketSite) return;
      var bucketValues = parseBucketsValue_(row[1]);
      config.buckets.push({ site: bucketSite, buckets: bucketValues });
    }
  });
  return config;
}

function rowMatchesHeader_(row, header) {
  if (!row || !header) return false;
  for (var i = 0; i < header.length; i++) {
    if (String(row[i] || '').trim() !== String(header[i] || '').trim()) {
      return false;
    }
  }
  return true;
}

function buildAutoPricingContext_() {
  var config = loadAutoPricingConfig_();
  var ranges = (config.ranges || []).slice().filter(function(r){
    return isFinite(r.upper) && isFinite(r.lower) && isFinite(r.coefficient);
  });
  ranges.sort(function(a,b){ return (b.upper || 0) - (a.upper || 0); });

  var combosMap = {};
  (config.sites || []).forEach(function(entry){
    if (!entry || !entry.site) return;
    var siteName = String(entry.site).trim();
    var normSite = normSite_(siteName);
    if (!normSite) return;
    var utmList = Array.isArray(entry.utmSources) ? entry.utmSources : [];
    var active = parseAutoPricingBoolean_(entry.active);
    var network = String(entry.network || '').trim();
    var normNetwork = normNetwork_(network);
    var identifier = String(entry.identifier || '').trim();
    utmList.forEach(function(utm){
      var normUtm = normUtm_(utm);
      if (!normUtm) return;
      var key = normSite + '||' + normUtm;
      if (!combosMap[key]) {
        combosMap[key] = { site: siteName, normSite: normSite, utm_source: utm, normUtm: normUtm, active: active, network: network, normNetwork: normNetwork, identifier: identifier };
      } else {
        combosMap[key].active = combosMap[key].active || active;
        if (!combosMap[key].utm_source && utm) combosMap[key].utm_source = utm;
        if (!combosMap[key].network && network) {
          combosMap[key].network = network;
          combosMap[key].normNetwork = normNetwork;
        }
        if (!combosMap[key].identifier && identifier) {
          combosMap[key].identifier = identifier;
        }
      }
    });
  });

  var combosAll = Object.keys(combosMap).map(function(key){ return combosMap[key]; });
  combosAll.sort(function(a,b){
    var s = String(a.site || '').localeCompare(String(b.site || ''), 'pt-BR');
    if (s !== 0) return s;
    return String(a.utm_source || '').localeCompare(String(b.utm_source || ''), 'pt-BR');
  });

  var combosActive = combosAll.filter(function(item){
    return parseAutoPricingBoolean_(item && item.active);
  });

  var siteOptionsMap = {};
  var utmOptionsMap = {};
  combosAll.forEach(function(combo){
    if (!siteOptionsMap[combo.normSite]) siteOptionsMap[combo.normSite] = combo.site;
    if (!utmOptionsMap[combo.normUtm]) utmOptionsMap[combo.normUtm] = combo.utm_source;
  });

  var siteOptions = Object.keys(siteOptionsMap).map(function(key){
    return { site: siteOptionsMap[key], normSite: key };
  }).sort(function(a,b){
    return String(a.site || '').localeCompare(String(b.site || ''), 'pt-BR');
  });

  if (!siteOptions.length) {
    var seenSiteFallback = {};
    (config.sites || []).forEach(function(entry){
      if (!entry) return;
      var siteName = String(entry.site || '').trim();
      if (!siteName) return;
      var normSite = normSite_(siteName);
      if (!normSite || seenSiteFallback[normSite]) return;
      seenSiteFallback[normSite] = true;
      siteOptions.push({ site: siteName, normSite: normSite });
    });
    siteOptions.sort(function(a,b){
      return String(a.site || '').localeCompare(String(b.site || ''), 'pt-BR');
    });
  }

  var utmOptions = Object.keys(utmOptionsMap).map(function(key){
    return { utm_source: utmOptionsMap[key], normUtm: key };
  }).sort(function(a,b){
    return String(a.utm_source || '').localeCompare(String(b.utm_source || ''), 'pt-BR');
  });

  if (!utmOptions.length) {
    var seenUtmFallback = {};
    (config.sites || []).forEach(function(entry){
      if (!entry) return;
      var utms = Array.isArray(entry.utmSources) ? entry.utmSources : [];
      utms.forEach(function(utm){
        var label = String(utm || '').trim();
        if (!label) return;
        var norm = normUtm_(label);
        if (!norm || seenUtmFallback[norm]) return;
        seenUtmFallback[norm] = true;
        utmOptions.push({ utm_source: label, normUtm: norm });
      });
    });
    utmOptions.sort(function(a,b){
      return String(a.utm_source || '').localeCompare(String(b.utm_source || ''), 'pt-BR');
    });
  }

  var bucketMap = {};
  (config.buckets || []).forEach(function(entry){
    if (!entry || !entry.site) return;
    var siteName = String(entry.site).trim();
    var normSite = normSite_(siteName);
    if (!normSite) return;
    var buckets = Array.isArray(entry.buckets) ? entry.buckets : [];
    var pairs = [];
    buckets.forEach(function(value){
      var str = String(value || '').trim();
      if (!str) return;
      var num = toNumber_(str);
      if (!isFinite(num)) return;
      pairs.push({ label: str, value: num });
    });
    pairs.sort(function(a,b){ return a.value - b.value; });
    bucketMap[normSite] = {
      labels: pairs.map(function(p){ return p.label; }),
      values: pairs.map(function(p){ return p.value; })
    };
  });

  return {
    config: config,
    ranges: ranges,
    combos: combosActive,
    activeCombos: combosActive,
    allCombos: combosAll,
    siteOptions: siteOptions,
    utmOptions: utmOptions,
    bucketMap: bucketMap
  };
}

function getIgoalRuleContext_() {
  var sh = ensurePricingSheetIgoal_();
  var values = sh.getDataRange().getValues();
  var headers = values.shift() || [];
  var idx = getHeaderIndexMap_(headers);
  var map = {};
  if (values.length) {
    values.forEach(function(row, index){
      var dominio = row[idx['dominio']];
      var utmRaw = row[idx['utm_source']];
      var utm = normalizeIgoalUtmSourceValue_(utmRaw);
      var url = row[idx['url']];
      var bloco = idx['bloco'] != null ? row[idx['bloco']] : '';
      if (!matchesAllowedAdUnitTokens_(bloco)) return;
      var normSite = normSite_(dominio);
      var normAdunit = normAdUnit_(bloco);
      var normSlug = normUrlStrict_(url);
      var normUtmSanitized = normUtm_(utm);
      var normUtmRaw = normUtm_(utmRaw);
      var keyVariants = [];
      if (normUtmSanitized) keyVariants.push(normUtmSanitized);
      if (normUtmRaw && normUtmRaw !== normUtmSanitized) keyVariants.push(normUtmRaw);
      if (!keyVariants.length) keyVariants.push('');
      var info = {
        price_rule: toNumber_(row[idx['price_rule']]),
        rowIndex: index + 2,
        dominio: dominio,
        utm_source: utm,
        url: url,
        bloco: bloco,
        adunit: bloco,
        normAdUnit: normAdUnit_(bloco),
        company_id: row[idx['company_id']],
        rule_id: row[idx['rule_id']]
      };
      keyVariants.forEach(function(normUtmValue){
        var key = normSite + '||' + normUtmValue + '||' + normAdunit + '||' + normSlug;
        map[key] = info;
      });
    });
  }
  return {
    sheet: sh,
    headers: headers,
    idx: idx,
    map: map,
    syncCol: headers.indexOf('Sincronizar'),
    formatRule: formatPriceRuleIgoal_
  };
}

function getAdSeletoRuleContext_() {
  var sh = ensurePricingSheetAdSeleto_();
  var values = sh.getDataRange().getValues();
  var headers = values.shift() || [];
  var idx = getHeaderIndexMap_(headers);
  var map = {};
  if (values.length) {
    values.forEach(function(row, index){
      var domain = row[idx['domain_id']];
      var utm = row[idx['utm_source']];
      var url = row[idx['url']];
      var slot = idx['slot_id'] != null ? row[idx['slot_id']] : '';
      if (!matchesAllowedAdUnitTokens_(slot)) return;
      var key = normSite_(domain) + '||' + normUtm_(utm) + '||' + normAdUnit_(slot) + '||' + normUrlStrict_(url);
      map[key] = {
        price_rule: toNumber_(row[idx['price_rule']]),
        rowIndex: index + 2,
        domain_id: domain,
        utm_source: utm,
        url: url,
        slot_id: slot,
        adunit: slot,
        normAdUnit: normAdUnit_(slot),
        spnprice_id: row[idx['spnprice_id']]
      };
    });
  }
  return {
    sheet: sh,
    headers: headers,
    idx: idx,
    map: map,
    syncCol: headers.indexOf('Sincronizar'),
    formatRule: formatPriceRuleAdSeleto_
  };
}

function buildRuleContextsForCombos_(comboMap) {
  var needs = {};
  Object.keys(comboMap || {}).forEach(function(key){
    var combo = comboMap[key];
    if (!combo) return;
    var normNetwork = combo.normNetwork || normNetwork_(combo.network);
    if (!normNetwork) return;
    needs[normNetwork] = true;
  });
  var contexts = {};
  if (needs['igoal']) {
    contexts['igoal'] = getIgoalRuleContext_();
  }
  if (needs['adseleto']) {
    contexts['adseleto'] = getAdSeletoRuleContext_();
  }
  return contexts;
}

function findCoefficientForCoverage_(coverage, ranges) {
  var cov = isFinite(coverage) ? coverage : 0;
  for (var i = 0; i < ranges.length; i++) {
    var range = ranges[i] || {};
    var upper = isFinite(range.upper) ? range.upper : 0;
    var lower = isFinite(range.lower) ? range.lower : 0;
    if (cov <= upper + 1e-9 && cov >= lower - 1e-9) {
      return range.coefficient;
    }
  }
  return ranges.length ? ranges[0].coefficient : 1;
}

function planToState_(plan) {
  plan = plan || {};
  return {
    config: plan.context ? plan.context.config : {},
    ranges: plan.context ? plan.context.ranges : [],
    siteOptions: plan.context ? plan.context.siteOptions : [],
    utmOptions: plan.context ? plan.context.utmOptions : [],
    bucketMap: plan.context ? plan.context.bucketMap : {},
    intervalHours: AUTO_PRICING_INTERVAL_HOURS,
    windowHours: AUTO_PRICING_WINDOW_HOURS,
    minRequests: AUTO_PRICING_MIN_REQUESTS,
    analysis: plan.analysis || { entries: [], siteRows: [], utmRows: [], summary: { totalEntries: 0, eligibleUpdates: 0 }, windowHours: [], excludedHour: null, filters: { sites: [], utm_sources: [] } }
  };
}

function computeAutoPricingPlan_(params) {
  params = params || {};
  var context = buildAutoPricingContext_();
  var includeInactive = params.includeInactive === true;
  var combos = includeInactive ? (context.allCombos || context.combos || []) : (context.activeCombos || context.combos || []);
  var siteFiltersRaw = [];
  if (Array.isArray(params.sites)) siteFiltersRaw = params.sites.slice();
  else if (Array.isArray(params.siteFilter)) siteFiltersRaw = params.siteFilter.slice();
  var utmFiltersRaw = [];
  if (Array.isArray(params.utm_sources)) utmFiltersRaw = params.utm_sources.slice();
  else if (Array.isArray(params.utmSources)) utmFiltersRaw = params.utmSources.slice();
  if (!combos.length) {
    return {
      context: context,
      ruleContexts: {},
      analysis: {
        generatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'),
        entries: [],
        siteRows: [],
        utmRows: [],
        windowHours: [],
        excludedHour: null,
        summary: { totalEntries: 0, eligibleUpdates: 0 },
        filters: { sites: siteFiltersRaw.slice(), utm_sources: utmFiltersRaw.slice() }
      }
    };
  }

  var siteFilterSet = createNormalizedSet_(siteFiltersRaw, normSite_);
  var utmFilterSet = createNormalizedSet_(utmFiltersRaw, normUtm_);

  var allowedComboMap = {};
  combos.forEach(function(combo){
    if (siteFilterSet && !siteFilterSet[combo.normSite]) return;
    if (utmFilterSet && !utmFilterSet[combo.normUtm]) return;
    allowedComboMap[combo.normSite + '||' + combo.normUtm] = combo;
  });

  if (!Object.keys(allowedComboMap).length) {
    combos.forEach(function(combo){
      allowedComboMap[combo.normSite + '||' + combo.normUtm] = combo;
    });
  }

  var ruleContexts = buildRuleContextsForCombos_(allowedComboMap);

  var baseSheet = getAutoPricingDataSheet_();
  var data = baseSheet.getDataRange().getValues();
  if (!data.length) {
    return {
      context: context,
      ruleContexts: ruleContexts,
      analysis: {
        generatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'),
        entries: [],
        siteRows: [],
        utmRows: [],
        windowHours: [],
        excludedHour: null,
        summary: { totalEntries: 0, eligibleUpdates: 0 },
        filters: {
          sites: siteFiltersRaw.slice(),
          utm_sources: utmFiltersRaw.slice()
        }
      }
    };
  }

  var headers = data[0] || [];
  var headerIdx = mapHeadersByAlias_(headers, HEADER_ALIASES);
  var requiredKeys = ['date', 'hour', 'site', 'channel', 'url', 'adunit', 'requests', 'coverage', 'ecpm', 'revenue'];
  for (var c = 0; c < requiredKeys.length; c++) {
    var key = requiredKeys[c];
    if (headerIdx[key] == null) {
      var aliases = HEADER_ALIASES[key] || [];
      var label = aliases.length ? aliases[0] : key;
      throw new Error('Coluna obrigatória ausente na base de dados: ' + label);
    }
  }

  var hourSet = {};
  var rows = [];
  var values = data.slice(1);
  values.forEach(function(row){
    var dateVal = row[headerIdx.date];
    var hourVal = row[headerIdx.hour];
    var siteVal = row[headerIdx.site];
    var utmVal = row[headerIdx.channel];
    var urlVal = row[headerIdx.url];
    var adunitVal = headerIdx.adunit != null ? row[headerIdx.adunit] : '';
    if (!matchesAllowedAdUnitTokens_(adunitVal)) return;
    var normSite = normSite_(siteVal);
    var normUtm = normUtm_(utmVal);
    var comboKey = normSite + '||' + normUtm;
    var comboInfo = allowedComboMap[comboKey];
    if (!comboInfo) return;
    var dateObj = parseDate_(dateVal);
    if (!dateObj) return;
    var hourNumber = parseInt(hourVal, 10);
    if (isNaN(hourNumber)) return;
    var ms = new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate(), hourNumber, 0, 0, 0).getTime();
    hourSet[ms] = true;
    rows.push({
      ms: ms,
      site: String(siteVal || ''),
      normSite: normSite,
      utm_source: String(utmVal || ''),
      normUtm: normUtm,
      url: String(urlVal || ''),
      normUrl: normUrlStrict_(urlVal),
      adunit: String(adunitVal || ''),
      normAdUnit: normAdUnit_(adunitVal),
      network: comboInfo && comboInfo.network ? String(comboInfo.network) : '',
      normNetwork: comboInfo && comboInfo.normNetwork ? comboInfo.normNetwork : normNetwork_(comboInfo && comboInfo.network),
      identifier: comboInfo && comboInfo.identifier ? String(comboInfo.identifier) : '',
      requests: toNumber_(row[headerIdx.requests]),
      coverage: parseCoverageValue_(row[headerIdx.coverage]),
      ecpm: toNumber_(row[headerIdx.ecpm]),
      revenue: toNumber_(row[headerIdx.revenue])
    });
  });

  var sortedHours = Object.keys(hourSet).map(function(k){ return +k; }).sort(function(a,b){ return a - b; });
  if (!sortedHours.length) {
    return {
      context: context,
      ruleContexts: ruleContexts,
      analysis: {
        generatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'),
        entries: [],
        siteRows: [],
        utmRows: [],
        windowHours: [],
        excludedHour: null,
        summary: { totalEntries: 0, eligibleUpdates: 0 },
        filters: {
          sites: siteFiltersRaw.slice(),
          utm_sources: utmFiltersRaw.slice()
        }
      }
    };
  }

  var excludedHour = sortedHours.length ? sortedHours[sortedHours.length - 1] : null;
  var considered = sortedHours.slice(0, Math.max(0, sortedHours.length - 1));
  var windowHours = [];
  for (var h = considered.length - 1; h >= 0 && windowHours.length < AUTO_PRICING_WINDOW_HOURS; h--) {
    windowHours.unshift(considered[h]);
  }
  if (!windowHours.length) {
    windowHours = considered.slice(-AUTO_PRICING_WINDOW_HOURS);
  }
  var windowHourSet = {};
  windowHours.forEach(function(ms){ windowHourSet[ms] = true; });

  var comboAgg = {};
  rows.forEach(function(row){
    if (!windowHourSet[row.ms]) return;
    var key = row.normSite + '||' + row.normUtm + '||' + row.normAdUnit + '||' + row.normUrl;
    if (!comboAgg[key]) {
      comboAgg[key] = {
        site: row.site,
        normSite: row.normSite,
        utm_source: row.utm_source,
        normUtm: row.normUtm,
        adunit: row.adunit,
        normAdUnit: row.normAdUnit,
        network: row.network,
        normNetwork: row.normNetwork,
        url: row.url,
        normUrl: row.normUrl,
        identifier: row.identifier,
        requests: 0,
        coverageWeighted: 0,
        ecpmWeighted: 0,
        revenue: 0,
        hours: {}
      };
    }
    var agg = comboAgg[key];
    var req = row.requests || 0;
    agg.requests += req;
    if (req > 0) {
      agg.coverageWeighted += (row.coverage || 0) * req;
      agg.ecpmWeighted += (row.ecpm || 0) * req;
    }
    agg.revenue += row.revenue || 0;
    agg.hours[row.ms] = true;
  });

  var entries = [];
  var seenKeys = {};
  Object.keys(comboAgg).forEach(function(key){
    var agg = comboAgg[key];
    var requests = agg.requests || 0;
    var avgCoverage = requests > 0 ? (agg.coverageWeighted / requests) : 0;
    var avgEcpm = requests > 0 ? (agg.ecpmWeighted / requests) : 0;
    var coefficient = findCoefficientForCoverage_(avgCoverage, context.ranges);
    var computedRule = avgEcpm * coefficient;
    var bucketInfo = chooseBucketValue_(computedRule, context.bucketMap[agg.normSite]);
    var contextForNetwork = ruleContexts[agg.normNetwork];
    var ruleInfo = contextForNetwork ? contextForNetwork.map[key] : null;
    var currentRule = ruleInfo ? ruleInfo.price_rule : 0;
    var supportedNetwork = !!contextForNetwork;
    var canSelect = supportedNetwork && bucketInfo.value != null;
    var eligible = canSelect && requests >= AUTO_PRICING_MIN_REQUESTS;
    var shouldUpdate = eligible && !nearlyEqual_(currentRule, bucketInfo.value);
    var reason = '';
    if (!supportedNetwork) reason = 'Rede não suportada pela automação';
    else if (!requests) reason = 'Sem dados nas últimas horas';
    else if (requests < AUTO_PRICING_MIN_REQUESTS) reason = 'Solicitações insuficientes';
    else if (bucketInfo.value == null) reason = 'Sem bucket disponível';
    else if (!shouldUpdate) reason = 'Regra atualizada';
    entries.push({
      key: key,
      site: agg.site,
      normSite: agg.normSite,
      utm_source: agg.utm_source,
      normUtm: agg.normUtm,
      adunit: agg.adunit,
      normAdUnit: agg.normAdUnit,
      network: agg.network,
      normNetwork: agg.normNetwork,
      url: agg.url,
      normUrl: agg.normUrl,
      identifier: agg.identifier,
      slot_id: ruleInfo && ruleInfo.slot_id ? ruleInfo.slot_id : '',
      bloco: ruleInfo && ruleInfo.bloco ? ruleInfo.bloco : '',
      domain_id: ruleInfo && ruleInfo.domain_id ? ruleInfo.domain_id : (ruleInfo && ruleInfo.dominio ? ruleInfo.dominio : (agg.normNetwork === 'adseleto' ? agg.identifier : '')),
      company_id: ruleInfo && ruleInfo.company_id ? ruleInfo.company_id : (agg.normNetwork === 'igoal' ? agg.identifier : ''),
      rule_id: ruleInfo && ruleInfo.rule_id ? ruleInfo.rule_id : '',
      spnprice_id: ruleInfo && ruleInfo.spnprice_id ? ruleInfo.spnprice_id : '',
      requests: requests,
      coverage: avgCoverage,
      ecpm: avgEcpm,
      coefficient: coefficient,
      computedRule: computedRule,
      bucketRule: bucketInfo.value,
      bucketLabel: bucketInfo.label,
      currentRule: currentRule,
      eligible: eligible,
      shouldUpdate: shouldUpdate,
      reason: reason,
      canSelect: canSelect,
      hoursUsed: Object.keys(agg.hours).map(function(ms){ return formatHourLabelAuto_(+ms); }).sort(),
      windowHours: windowHours.map(formatHourLabelAuto_)
    });
    seenKeys[key] = true;
  });

  Object.keys(ruleContexts).forEach(function(networkKey){
    var ctx = ruleContexts[networkKey];
    if (!ctx) return;
    Object.keys(ctx.map).forEach(function(key){
      if (seenKeys[key]) return;
      var parts = key.split('||');
      if (parts.length < 4) return;
      var normSite = parts[0];
      var normUtm = parts[1];
      var normAdUnit = parts[2];
      var comboKey = normSite + '||' + normUtm;
      var comboInfo = allowedComboMap[comboKey];
      if (!comboInfo) return;
      var info = ctx.map[key];
      var adunitLabel = info.adunit || info.bloco || info.slot_id || '';
      var networkName = comboInfo.network || '';
      entries.push({
        key: key,
        site: info.dominio || info.domain_id || comboInfo.site,
        normSite: normSite,
        utm_source: info.utm_source || comboInfo.utm_source,
        normUtm: normUtm,
        adunit: adunitLabel,
        normAdUnit: normAdUnit,
        network: networkName,
        normNetwork: networkKey,
        url: info.url,
        normUrl: normUrlStrict_(info.url),
        identifier: comboInfo.identifier || '',
        slot_id: info.slot_id || '',
        bloco: info.bloco || '',
        domain_id: info.domain_id || info.dominio || (networkKey === 'adseleto' ? (comboInfo.identifier || '') : ''),
        company_id: info.company_id || (networkKey === 'igoal' ? (comboInfo.identifier || '') : ''),
        rule_id: info.rule_id || '',
        spnprice_id: info.spnprice_id || '',
        requests: 0,
        coverage: 0,
        ecpm: 0,
        coefficient: findCoefficientForCoverage_(0, context.ranges),
        computedRule: 0,
        bucketRule: null,
        bucketLabel: null,
        currentRule: info.price_rule,
        eligible: false,
        shouldUpdate: false,
        reason: 'Sem dados nas últimas horas',
        canSelect: false,
        hoursUsed: [],
        windowHours: windowHours.map(formatHourLabelAuto_)
      });
    });
  });

  entries.sort(function(a,b){
    var siteCmp = String(a.site || '').localeCompare(String(b.site || ''), 'pt-BR');
    if (siteCmp !== 0) return siteCmp;
    var utmCmp = String(a.utm_source || '').localeCompare(String(b.utm_source || ''), 'pt-BR');
    if (utmCmp !== 0) return utmCmp;
    var adunitCmp = String(a.adunit || '').localeCompare(String(b.adunit || ''), 'pt-BR');
    if (adunitCmp !== 0) return adunitCmp;
    return String(a.url || '').localeCompare(String(b.url || ''), 'pt-BR');
  });

  var siteAgg = {};
  var utmAgg = {};
  entries.forEach(function(entry){
    var ctx = ruleContexts[entry.normNetwork];
    if (!ctx) return;
    var keySite = entry.normSite;
    if (!siteAgg[keySite]) {
      siteAgg[keySite] = {
        site: entry.site,
        normSite: entry.normSite,
        requests: 0,
        coverageWeighted: 0,
        ecpmWeighted: 0,
        currentWeighted: 0,
        networks: {}
      };
    }
    var siteItem = siteAgg[keySite];
    var req = entry.requests || 0;
    siteItem.requests += req;
    if (req > 0) {
      siteItem.coverageWeighted += (entry.coverage || 0) * req;
      siteItem.ecpmWeighted += (entry.ecpm || 0) * req;
      siteItem.currentWeighted += (entry.currentRule || 0) * req;
    }
    if (entry.normNetwork) {
      siteItem.networks[entry.normNetwork] = entry.network || entry.normNetwork;
    }

    var keyUtm = entry.normSite + '||' + entry.normUtm;
    if (!utmAgg[keyUtm]) {
      utmAgg[keyUtm] = {
        site: entry.site,
        normSite: entry.normSite,
        utm_source: entry.utm_source,
        normUtm: entry.normUtm,
        requests: 0,
        coverageWeighted: 0,
        ecpmWeighted: 0,
        currentWeighted: 0,
        networks: {}
      };
    }
    var utmItem = utmAgg[keyUtm];
    utmItem.requests += req;
    if (req > 0) {
      utmItem.coverageWeighted += (entry.coverage || 0) * req;
      utmItem.ecpmWeighted += (entry.ecpm || 0) * req;
      utmItem.currentWeighted += (entry.currentRule || 0) * req;
    }
    if (entry.normNetwork) {
      utmItem.networks[entry.normNetwork] = entry.network || entry.normNetwork;
    }
  });

  var siteRows = Object.keys(siteAgg).map(function(key){
    var item = siteAgg[key];
    var requests = item.requests || 0;
    var coverage = requests > 0 ? (item.coverageWeighted / requests) : 0;
    var ecpm = requests > 0 ? (item.ecpmWeighted / requests) : 0;
    var coefficient = findCoefficientForCoverage_(coverage, context.ranges);
    var computedRule = ecpm * coefficient;
    var bucketInfo = chooseBucketValue_(computedRule, context.bucketMap[item.normSite]);
    var currentRule = requests > 0 ? (item.currentWeighted / requests) : 0;
    var eligible = requests >= AUTO_PRICING_MIN_REQUESTS && bucketInfo.value != null;
    var shouldUpdate = eligible && !nearlyEqual_(currentRule, bucketInfo.value);
    var reason = 'Sem alteração';
    if (!eligible) {
      if (requests < AUTO_PRICING_MIN_REQUESTS) reason = 'Solicitações insuficientes';
      else if (bucketInfo.value == null) reason = 'Sem bucket disponível';
    }
    return {
      site: item.site,
      normSite: item.normSite,
      requests: requests,
      coverage: coverage,
      ecpm: ecpm,
      coefficient: coefficient,
      computedRule: computedRule,
      bucketRule: bucketInfo.value,
      bucketLabel: bucketInfo.label,
      currentRule: currentRule,
      shouldUpdate: shouldUpdate,
      eligible: eligible,
      reason: reason,
      networks: Object.keys(item.networks || {}).map(function(n){ return item.networks[n]; })
    };
  }).sort(function(a,b){
    return String(a.site || '').localeCompare(String(b.site || ''), 'pt-BR');
  });

  var utmRows = Object.keys(utmAgg).map(function(key){
    var item = utmAgg[key];
    var requests = item.requests || 0;
    var coverage = requests > 0 ? (item.coverageWeighted / requests) : 0;
    var ecpm = requests > 0 ? (item.ecpmWeighted / requests) : 0;
    var coefficient = findCoefficientForCoverage_(coverage, context.ranges);
    var computedRule = ecpm * coefficient;
    var bucketInfo = chooseBucketValue_(computedRule, context.bucketMap[item.normSite]);
    var currentRule = requests > 0 ? (item.currentWeighted / requests) : 0;
    var eligible = requests >= AUTO_PRICING_MIN_REQUESTS && bucketInfo.value != null;
    var shouldUpdate = eligible && !nearlyEqual_(currentRule, bucketInfo.value);
    var reason = 'Sem alteração';
    if (!eligible) {
      if (requests < AUTO_PRICING_MIN_REQUESTS) reason = 'Solicitações insuficientes';
      else if (bucketInfo.value == null) reason = 'Sem bucket disponível';
    }
    return {
      site: item.site,
      normSite: item.normSite,
      utm_source: item.utm_source,
      normUtm: item.normUtm,
      requests: requests,
      coverage: coverage,
      ecpm: ecpm,
      coefficient: coefficient,
      computedRule: computedRule,
      bucketRule: bucketInfo.value,
      bucketLabel: bucketInfo.label,
      currentRule: currentRule,
      shouldUpdate: shouldUpdate,
      eligible: eligible,
      reason: reason,
      networks: Object.keys(item.networks || {}).map(function(n){ return item.networks[n]; })
    };
  }).sort(function(a,b){
    var siteCmp = String(a.site || '').localeCompare(String(b.site || ''), 'pt-BR');
    if (siteCmp !== 0) return siteCmp;
    return String(a.utm_source || '').localeCompare(String(b.utm_source || ''), 'pt-BR');
  });

  var eligibleUpdates = entries.filter(function(e){ return e.shouldUpdate; }).length;

  return {
    context: context,
    ruleContexts: ruleContexts,
    analysis: {
      generatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'),
      entries: entries,
      siteRows: siteRows,
      utmRows: utmRows,
      windowHours: windowHours.map(formatHourLabelAuto_),
      excludedHour: excludedHour != null ? formatHourLabelAuto_(excludedHour) : null,
      summary: {
        totalEntries: entries.length,
        eligibleUpdates: eligibleUpdates
      },
      filters: {
        sites: siteFiltersRaw.slice(),
        utm_sources: utmFiltersRaw.slice()
      }
    }
  };
}

function getAutoPricingState(params) {
  var base = (params && typeof params === 'object') ? params : {};
  var payload = {};
  Object.keys(base).forEach(function(key){ payload[key] = base[key]; });
  payload.includeInactive = true;
  var plan = computeAutoPricingPlan_(payload);
  return planToState_(plan);
}

function saveAutoPricingConfig(payload) {
  var data = payload || {};
  var ranges = Array.isArray(data.ranges) ? data.ranges : [];
  var sites = Array.isArray(data.sites) ? data.sites : [];
  var buckets = Array.isArray(data.buckets) ? data.buckets : [];

  var sh = ensureAutoPricingConfigSheet_();
  sh.clearContents();

  var rows = [];
  var minCols = Math.max(
    AUTO_PRICING_RANGE_HEADER.length,
    AUTO_PRICING_SITE_HEADER.length,
    AUTO_PRICING_BUCKET_HEADER.length
  );
  function push(arr) {
    var row = (arr || []).slice();
    while (row.length < minCols) row.push('');
    rows.push(row);
  }

  push(AUTO_PRICING_RANGE_HEADER);
  ranges.forEach(function(item){
    if (!item) return;
    push([item.upper, item.lower, item.coefficient]);
  });
  push([]);
  push(AUTO_PRICING_SITE_HEADER);
  sites.forEach(function(item){
    if (!item || !item.site) return;
    var utmList = Array.isArray(item.utmSources) ? item.utmSources : [];
    var network = String(item.network || '').trim();
    var identifier = String(item.identifier || '').trim();
    push([item.site, utmList.join(';'), network, identifier, parseAutoPricingBoolean_(item.active)]);
  });
  push([]);
  push(AUTO_PRICING_BUCKET_HEADER);
  buckets.forEach(function(item){
    if (!item || !item.site) return;
    var bucketList = Array.isArray(item.buckets) ? item.buckets : [];
    var formatted = '[' + bucketList.map(function(v){ return '"' + String(v || '').trim() + '"'; }).join(',') + ']';
    push([item.site, formatted]);
  });

  if (rows.length) {
    sh.getRange(1,1,rows.length,rows[0].length).setValues(rows);
  }

  return getAutoPricingState();
}

function ensureRuleRowForEntry_(networkKey, ctx, entry, formattedValue, rawValue) {
  networkKey = normNetwork_(networkKey);
  if (!ctx || !ctx.sheet) return null;
  var headers = ctx.headers || [];
  if (!headers.length) return null;
  var idx = ctx.idx || {};
  var syncCol = ctx.syncCol;
  var sheet = ctx.sheet;
  if (networkKey === 'adseleto') {
    var domainId = coalesceForPayload_(entry.domain_id, entry.identifier);
    var utm = coalesceForPayload_(entry.utm_source);
    var slot = coalesceForPayload_(entry.slot_id, entry.adunit);
    var slug = entry.normUrl || normUrlStrict_(entry.url);
    if (!domainId) {
      entry.reason = entry.reason || 'ID do domínio não configurado na aba de precificação';
      return null;
    }
    if (!slot) {
      entry.reason = entry.reason || 'Bloco/slot não configurado para a regra de precificação';
      return null;
    }
    if (!slug) {
      entry.reason = entry.reason || 'URL inválida para criação da regra';
      return null;
    }
    var spnId = entry.spnprice_id || generateSpnpriceId_();
    var rowValues = [];
    for (var i = 0; i < headers.length; i++) rowValues.push('');
    if (idx['spnprice_id'] != null) rowValues[idx['spnprice_id']] = spnId;
    if (idx['domain_id'] != null) rowValues[idx['domain_id']] = domainId;
    if (idx['price_rule'] != null) rowValues[idx['price_rule']] = formattedValue;
    if (idx['utm_source'] != null) rowValues[idx['utm_source']] = utm;
    if (idx['slot_id'] != null) rowValues[idx['slot_id']] = slot;
    if (idx['url'] != null) rowValues[idx['url']] = slug;
    if (syncCol != null && syncCol >= 0) rowValues[syncCol] = 'Sim';
    var rowIndex = sheet.getLastRow() + 1;
    sheet.getRange(rowIndex, 1, 1, headers.length).setValues([rowValues]);
    var info = {
      price_rule: rawValue,
      rowIndex: rowIndex,
      domain_id: domainId,
      utm_source: utm,
      url: slug,
      slot_id: slot,
      adunit: slot,
      normAdUnit: normAdUnit_(slot),
      spnprice_id: spnId,
      wasRuleCreated: true
    };
    ctx.map[entry.key] = info;
    entry.domain_id = domainId;
    entry.slot_id = slot;
    entry.spnprice_id = spnId;
    entry.wasRuleCreated = true;
    return info;
  } else if (networkKey === 'igoal') {
    var dominio = coalesceForPayload_(entry.site, entry.domain_id);
    var utmIgoal = normalizeIgoalUtmSourceValue_(coalesceForPayload_(entry.utm_source));
    var bloco = coalesceForPayload_(entry.bloco, entry.adunit);
    var slugIgoal = entry.normUrl || normUrlStrict_(entry.url);
    var companyId = coalesceForPayload_(entry.company_id, entry.identifier);
    if (!dominio) {
      entry.reason = entry.reason || 'Domínio não configurado para a Igoal';
      return null;
    }
    if (!companyId) {
      entry.reason = entry.reason || 'Company ID não configurado para a Igoal';
      return null;
    }
    if (!bloco) {
      entry.reason = entry.reason || 'Bloco não configurado para a Igoal';
      return null;
    }
    if (!slugIgoal) {
      entry.reason = entry.reason || 'URL inválida para criação da regra';
      return null;
    }
    var rowValuesIgoal = [];
    for (var j = 0; j < headers.length; j++) rowValuesIgoal.push('');
    if (idx['rule_id'] != null) rowValuesIgoal[idx['rule_id']] = '';
    if (idx['dominio'] != null) rowValuesIgoal[idx['dominio']] = dominio;
    if (idx['utm_source'] != null) rowValuesIgoal[idx['utm_source']] = utmIgoal;
    if (idx['url'] != null) rowValuesIgoal[idx['url']] = slugIgoal;
    if (idx['bloco'] != null) rowValuesIgoal[idx['bloco']] = bloco;
    if (idx['company_id'] != null) rowValuesIgoal[idx['company_id']] = companyId;
    if (idx['price_rule'] != null) rowValuesIgoal[idx['price_rule']] = formattedValue;
    if (syncCol != null && syncCol >= 0) rowValuesIgoal[syncCol] = 'Sim';
    var rowIndexIgoal = sheet.getLastRow() + 1;
    sheet.getRange(rowIndexIgoal, 1, 1, headers.length).setValues([rowValuesIgoal]);
    var infoIgoal = {
      price_rule: rawValue,
      rowIndex: rowIndexIgoal,
      dominio: dominio,
      utm_source: utmIgoal,
      url: slugIgoal,
      bloco: bloco,
      adunit: bloco,
      normAdUnit: normAdUnit_(bloco),
      company_id: companyId,
      wasRuleCreated: true
    };
    ctx.map[entry.key] = infoIgoal;
    entry.domain_id = dominio;
    entry.company_id = companyId;
    entry.utm_source = utmIgoal;
    entry.bloco = bloco;
    entry.wasRuleCreated = true;
    return infoIgoal;
  }
  return null;
}

function prepareRuleInfoForUpdate_(networkKey, ctx, entry, ruleInfo) {
  networkKey = normNetwork_(networkKey);
  if (!ctx || !ctx.sheet || !ruleInfo || !ruleInfo.rowIndex) return ruleInfo;
  var idx = ctx.idx || {};
  var sheet = ctx.sheet;
  if (networkKey === 'adseleto') {
    var domainValue = coalesceForPayload_(ruleInfo.domain_id, entry.domain_id, entry.identifier);
    if (domainValue && idx['domain_id'] != null && String(ruleInfo.domain_id || '').trim() !== String(domainValue).trim()) {
      sheet.getRange(ruleInfo.rowIndex, idx['domain_id'] + 1).setValue(domainValue);
      ruleInfo.domain_id = domainValue;
    }
    if (domainValue) entry.domain_id = domainValue;
    var slotValue = coalesceForPayload_(ruleInfo.slot_id, entry.slot_id, entry.adunit);
    if (slotValue && idx['slot_id'] != null && String(ruleInfo.slot_id || '').trim() !== String(slotValue).trim()) {
      sheet.getRange(ruleInfo.rowIndex, idx['slot_id'] + 1).setValue(slotValue);
      ruleInfo.slot_id = slotValue;
      ruleInfo.adunit = slotValue;
      ruleInfo.normAdUnit = normAdUnit_(slotValue);
    }
    if (slotValue) entry.slot_id = slotValue;
    var utmValue = coalesceForPayload_(ruleInfo.utm_source, entry.utm_source);
    if (utmValue && idx['utm_source'] != null && String(ruleInfo.utm_source || '').trim() !== String(utmValue).trim()) {
      sheet.getRange(ruleInfo.rowIndex, idx['utm_source'] + 1).setValue(utmValue);
      ruleInfo.utm_source = utmValue;
    }
    var slugValue = coalesceForPayload_(ruleInfo.url, entry.normUrl, normUrlStrict_(entry.url));
    if (slugValue && idx['url'] != null && String(ruleInfo.url || '').trim() !== String(slugValue).trim()) {
      sheet.getRange(ruleInfo.rowIndex, idx['url'] + 1).setValue(slugValue);
      ruleInfo.url = slugValue;
    }
    var spnValue = coalesceForPayload_(ruleInfo.spnprice_id, entry.spnprice_id);
    if (!spnValue) {
      spnValue = generateSpnpriceId_();
      if (idx['spnprice_id'] != null) {
        sheet.getRange(ruleInfo.rowIndex, idx['spnprice_id'] + 1).setValue(spnValue);
      }
      ruleInfo.spnprice_id = spnValue;
    }
    entry.spnprice_id = spnValue;
  } else if (networkKey === 'igoal') {
    var dominioValue = coalesceForPayload_(ruleInfo.dominio, entry.site, entry.domain_id);
    if (dominioValue && idx['dominio'] != null && String(ruleInfo.dominio || '').trim() !== String(dominioValue).trim()) {
      sheet.getRange(ruleInfo.rowIndex, idx['dominio'] + 1).setValue(dominioValue);
      ruleInfo.dominio = dominioValue;
    }
    if (dominioValue) entry.domain_id = dominioValue;
    var blocoValue = coalesceForPayload_(ruleInfo.bloco, ruleInfo.slot_id, entry.bloco, entry.adunit);
    if (blocoValue && idx['bloco'] != null && String(ruleInfo.bloco || '').trim() !== String(blocoValue).trim()) {
      sheet.getRange(ruleInfo.rowIndex, idx['bloco'] + 1).setValue(blocoValue);
      ruleInfo.bloco = blocoValue;
      ruleInfo.adunit = blocoValue;
      ruleInfo.normAdUnit = normAdUnit_(blocoValue);
    }
    if (blocoValue) entry.bloco = blocoValue;
    var utmValueIgoal = normalizeIgoalUtmSourceValue_(coalesceForPayload_(ruleInfo.utm_source, entry.utm_source));
    if (idx['utm_source'] != null && String(ruleInfo.utm_source || '').trim() !== String(utmValueIgoal).trim()) {
      sheet.getRange(ruleInfo.rowIndex, idx['utm_source'] + 1).setValue(utmValueIgoal);
      ruleInfo.utm_source = utmValueIgoal;
    }
    entry.utm_source = utmValueIgoal;
    var urlValue = coalesceForPayload_(ruleInfo.url, entry.normUrl, normUrlStrict_(entry.url));
    if (urlValue && idx['url'] != null && String(ruleInfo.url || '').trim() !== String(urlValue).trim()) {
      sheet.getRange(ruleInfo.rowIndex, idx['url'] + 1).setValue(urlValue);
      ruleInfo.url = urlValue;
    }
    var companyValue = coalesceForPayload_(ruleInfo.company_id, entry.company_id, entry.identifier);
    if (companyValue && idx['company_id'] != null && String(ruleInfo.company_id || '').trim() !== String(companyValue).trim()) {
      sheet.getRange(ruleInfo.rowIndex, idx['company_id'] + 1).setValue(companyValue);
      ruleInfo.company_id = companyValue;
    }
    if (companyValue) entry.company_id = companyValue;
  }
  if (ruleInfo.wasRuleCreated) {
    entry.wasRuleCreated = true;
  }
  return ruleInfo;
}

function runAutoPricing(params) {
  params = params || {};
  params.trigger = params.trigger || 'manual';
  var plan = computeAutoPricingPlan_(params);
  var selectedSet = null;
  if (Array.isArray(params.entryKeys) && params.entryKeys.length) {
    selectedSet = {};
    params.entryKeys.forEach(function(key){
      if (key == null) return;
      selectedSet[String(key)] = true;
    });
  }
  var ruleContexts = plan.ruleContexts || {};
  var updates = plan.analysis.entries.filter(function(entry){
    if (!entry || !entry.key) return false;
    var canSelect = entry.canSelect !== false;
    if (selectedSet) {
      if (!selectedSet[entry.key]) return false;
      if (!canSelect) return false;
      var ctx = ruleContexts[entry.normNetwork];
      if (!ctx || !ctx.map) return false;
      if (entry.bucketRule == null) return false;
      entry.forceUpdate = !entry.shouldUpdate;
      return true;
    }
    if (!entry.shouldUpdate) return false;
    if (!canSelect) return false;
    return true;
  });
  if (!updates.length) {
    var manualAttempt = selectedSet && Object.keys(selectedSet).length;
    return {
      ok: true,
      updated: 0,
      message: manualAttempt ? 'Nenhuma atualização válida para as URLs selecionadas.' : 'Nenhuma atualização necessária.',
      state: planToState_(plan)
    };
  }

  var triggeredNetworks = {};
  updates.forEach(function(entry){
    if (!entry || !entry.normNetwork) return;
    triggeredNetworks[entry.normNetwork] = true;
  });
  var tz = Session.getScriptTimeZone();
  var timestamp = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
  var logSheet = ensureAutoPricingLogSheet_();
  var logRows = [];

  Object.keys(ruleContexts).forEach(function(networkKey){
    var ctx = ruleContexts[networkKey];
    if (!ctx) return;
    var priceCol = ctx.idx ? ctx.idx['price_rule'] : null;
    if (priceCol == null) {
      throw new Error("A coluna 'price_rule' não foi encontrada na aba de precificação de " + networkKey + '.');
    }
    var syncCol = ctx.syncCol;
    updates.filter(function(entry){ return entry.normNetwork === networkKey; }).forEach(function(entry){
      var ruleInfo = ctx.map[entry.key];
      var newValue = entry.bucketRule;
      var formatted = ctx.formatRule ? ctx.formatRule(newValue) : newValue;
      if (!ruleInfo || !ruleInfo.rowIndex) {
        ruleInfo = ensureRuleRowForEntry_(networkKey, ctx, entry, formatted, newValue);
      }
      if (ruleInfo && ruleInfo.rowIndex) {
        ruleInfo = prepareRuleInfoForUpdate_(networkKey, ctx, entry, ruleInfo);
      }
      var payloadLog = buildPricingPayloadLog_(networkKey, entry, ruleInfo, formatted, newValue);
      if (!ruleInfo || !ruleInfo.rowIndex) {
        var failureReason = entry.reason || 'Regra não encontrada na aba de precificação (Bloco: ' + (entry.adunit || '-') + ')';
        logRows.push([timestamp, params.trigger, entry.site, entry.network, entry.utm_source, entry.url, entry.currentRule, entry.bucketRule, entry.coverage, entry.ecpm, entry.coefficient, entry.bucketLabel, entry.requests, failureReason, payloadLog]);
        entry.reason = failureReason;
        return;
      }
      var previousRule = entry.wasRuleCreated ? '' : ruleInfo.price_rule;
      ctx.sheet.getRange(ruleInfo.rowIndex, priceCol + 1).setValue(formatted);
      if (syncCol != null && syncCol >= 0) {
        ctx.sheet.getRange(ruleInfo.rowIndex, syncCol + 1).setValue('Sim');
      }
      ruleInfo.price_rule = newValue;
      var actionNote = entry.wasRuleCreated ? 'Regra criada via automação (Bloco: ' + (entry.adunit || '-') + ')' : 'Atualizado via automação (Bloco: ' + (entry.adunit || '-') + ')';
      if (entry.forceUpdate) {
        actionNote += ' (atualização manual forçada)';
      }
      logRows.push([timestamp, params.trigger, entry.site, entry.network, entry.utm_source, entry.url, previousRule, newValue, entry.coverage, entry.ecpm, entry.coefficient, entry.bucketLabel, entry.requests, actionNote, payloadLog]);
    });
  });

  if (logRows.length) {
    var range = logSheet.getRange(logSheet.getLastRow() + 1, 1, logRows.length, AUTO_PRICING_LOG_HEADERS.length);
    var payload = logRows.map(function(row){
      var r = row.slice();
      while (r.length < AUTO_PRICING_LOG_HEADERS.length) r.push('');
      return r;
    });
    range.setValues(payload);
  }

  var syncSummaries = [];
  function pushSyncSummary(label, result) {
    var header = '[' + label + '] ';
    if (!result) {
      syncSummaries.push(header + 'Sincronização concluída.');
      return;
    }
    var ok = result.ok !== false;
    var message = typeof result.message === 'string' ? result.message : '';
    if (!message) {
      message = ok ? 'Sincronização concluída.' : 'Falha ao sincronizar.';
    }
    syncSummaries.push(header + (ok ? 'OK' : 'ERRO') + '\n' + message);
  }

  function runNetworkSync(label, fn) {
    if (typeof fn !== 'function') {
      syncSummaries.push('[' + label + '] Função de sincronização indisponível.');
      return;
    }
    try {
      var result = fn();
      pushSyncSummary(label, result);
    } catch (err) {
      syncSummaries.push('[' + label + '] ERRO\n' + (err && err.message ? err.message : 'Falha desconhecida.'));
    }
  }

  if (triggeredNetworks['adseleto']) {
    runNetworkSync('AdSeleto', runAdSeletoPricingUpdate);
  }
  if (triggeredNetworks['igoal']) {
    runNetworkSync('Igoal', runIgoalPricingUpdate);
  }

  var refreshed = computeAutoPricingPlan_(params);
  var baseMessage = 'Precificação aplicada para ' + logRows.length + ' itens.';
  if (syncSummaries.length) {
    baseMessage += '\n\n' + syncSummaries.join('\n\n');
  }
  return {
    ok: true,
    updated: logRows.length,
    message: baseMessage,
    state: planToState_(refreshed)
  };
}