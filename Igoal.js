// PricingIgoal.gs

const IGOAL_API_BASE_URL = 'https://my.spun.com.br';
const IGOAL_API_ENDPOINT = '/api/prices/update';
const IGOAL_API_TOKEN = '8jwl4v1ZmBYQlwFzPPEHNkYC8IOvRxB3ino1665b93f36cd228';

function buildIgoalApiUrl_() {
  return IGOAL_API_BASE_URL.replace(/\/$/, '') + IGOAL_API_ENDPOINT;
}

function normalizeIgoalPrice_(value) {
  if (value == null) return '';
  if (typeof value === 'number') return isNaN(value) ? '' : value.toFixed(2);
  var str = String(value).trim();
  if (!str) return '';
  str = str.replace(/\s/g, '').replace(/%/g, '');
  var comma = str.indexOf(',');
  var dot = str.indexOf('.');
  if (comma !== -1 && dot !== -1) {
    if (comma > dot) {
      str = str.replace(/\./g, '').replace(/,/g, '.');
    } else {
      str = str.replace(/,/g, '');
    }
  } else if (comma !== -1) {
    str = str.replace(/\./g, '').replace(/,/g, '.');
  } else {
    str = str.replace(/,/g, '');
  }
  var numeric = parseFloat(str);
  if (isNaN(numeric)) return '';
  return numeric.toFixed(2);
}

function shouldSyncIgoalRow_(rowValue, syncColumnIndex) {
  if (syncColumnIndex === -1) return true;
  var raw = rowValue == null ? '' : String(rowValue).trim().toLowerCase();
  if (!raw) return true;
  return ['sim', 'yes', 'y'].indexOf(raw) !== -1;
}

function runIgoalPricingUpdate() {
  try {
    var sheet = typeof ensurePricingSheetIgoal_ === 'function'
      ? ensurePricingSheetIgoal_()
      : SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Precificação - Igoal');
    if (!sheet) {
      throw new Error('A aba "Precificação - Igoal" não foi encontrada.');
    }

    var data = sheet.getDataRange().getValues();
    if (!data || data.length < 2) {
      return { ok: true, message: 'Nenhuma regra para sincronizar na Igoal.' };
    }

    var headers = data[0] || [];
    var idx = getHeaderIndexMap_(headers);
    var requiredHeaders = ['dominio', 'url', 'company_id', 'price_rule'];
    var missingHeaders = requiredHeaders.filter(function(name) { return !(name in idx); });
    if (missingHeaders.length) {
      throw new Error('Colunas obrigatórias ausentes na planilha da Igoal: ' + missingHeaders.join(', '));
    }
    var syncIdx = -1;
    if ('Sincronizar' in idx) {
      syncIdx = idx['Sincronizar'];
    } else if ('sincronizar' in idx) {
      syncIdx = idx['sincronizar'];
    } else {
      syncIdx = headers.indexOf('Sincronizar');
    }

    var rowsToSync = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!shouldSyncIgoalRow_(syncIdx === -1 ? '' : row[syncIdx], syncIdx)) {
        continue;
      }
      rowsToSync.push({ row: row, rowNum: i + 1 });
    }

    if (!rowsToSync.length) {
      return { ok: true, message: 'Nenhuma regra válida para sincronizar na Igoal.' };
    }

    var logMessages = [];
    var url = buildIgoalApiUrl_();

    rowsToSync.forEach(function(item) {
      var row = item.row;
      var dominio = idx.hasOwnProperty('dominio') ? String(row[idx['dominio']] || '').trim() : '';
      var urlPath = idx.hasOwnProperty('url') ? String(row[idx['url']] || '').trim() : '';
      var companyIdRaw = idx.hasOwnProperty('company_id') ? String(row[idx['company_id']] || '').trim() : '';
      var priceRuleRaw = idx.hasOwnProperty('price_rule') ? row[idx['price_rule']] : '';
      var priceRule = normalizeIgoalPrice_(priceRuleRaw);
      var utmSourceRaw = idx.hasOwnProperty('utm_source') ? row[idx['utm_source']] : '';
      var utmSource = typeof normalizeIgoalUtmSourceValue_ === 'function'
        ? normalizeIgoalUtmSourceValue_(utmSourceRaw)
        : String(utmSourceRaw == null ? '' : utmSourceRaw).trim();
      var bloco = idx.hasOwnProperty('bloco') ? String(row[idx['bloco']] || '').trim() : '';

      if (!dominio || !urlPath || !companyIdRaw || !priceRule) {
        logMessages.push('- ' + (dominio || '(domínio vazio)') + '/' + (urlPath || '(url vazia)') + ': ERRO! Campos obrigatórios ausentes ou inválidos.');
        return;
      }

      var companyIdValue = parseInt(companyIdRaw, 10);
      var payload = {
        dominio: dominio,
        url: urlPath,
        company_id: isNaN(companyIdValue) ? companyIdRaw : companyIdValue,
        price_rule: priceRule,
      };
      if (utmSource) payload.utm_source = utmSource;
      if (bloco) payload.bloco = bloco;

      try {
        var response = UrlFetchApp.fetch(url, {
          method: 'post',
          contentType: 'application/json',
          muteHttpExceptions: true,
          payload: JSON.stringify(payload),
          headers: { Authorization: IGOAL_API_TOKEN },
        });

        var statusCode = response.getResponseCode();
        var raw = response.getContentText() || '';
        var json;
        try { json = raw ? JSON.parse(raw) : null; } catch (err) { json = null; }

        var success = (statusCode === 200) && (!!json && (json.status === true || json.message === '200' || json.success === true));
        var message = '';
        if (success) {
          message = json && (json.data || json.message) ? (json.data || json.message) : 'Preço atualizado com sucesso.';
          if (syncIdx !== -1) {
            sheet.getRange(item.rowNum, syncIdx + 1).clearContent();
          }
          if (
            idx.hasOwnProperty('utm_source') &&
            idx['utm_source'] != null &&
            String(row[idx['utm_source']] || '').trim() !== utmSource
          ) {
            sheet.getRange(item.rowNum, idx['utm_source'] + 1).setValue(utmSource);
            row[idx['utm_source']] = utmSource;
          }
          logMessages.push('- ' + dominio + '/' + urlPath + ': OK! ' + message);
        } else {
          message = json && (json.data || json.message) ? (json.data || json.message) : (raw || ('Erro HTTP ' + statusCode));
          logMessages.push('- ' + dominio + '/' + urlPath + ': ERRO! ' + message);
        }
      } catch (err) {
        logMessages.push('- ' + dominio + '/' + urlPath + ': ERRO! ' + err.message);
      }
    });

    return { ok: true, message: 'Resultado da Sincronização Igoal:\n\n' + logMessages.join('\n') };
  } catch (e) {
    Logger.log('[IGOAL][ERRO FATAL] ' + e.message);
    return { ok: false, message: 'Erro ao sincronizar com a Igoal: ' + e.message };
  }
}