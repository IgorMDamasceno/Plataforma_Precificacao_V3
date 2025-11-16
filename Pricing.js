// Pricing.gs

const PMD_API_URL = 'https://pmd.cloud.adseleto.com/api/domains/';
const PMD_API_TOKEN = '6fae52ca5fb3d41eca803d903a73e2d011bb96b5415f101c563025a8c3470d85b249b2103a3e2a4dc073c3f62aebd438f0ddf2b72d6bc0c9156a8e5801d3d73023600b64232f6e407f4cea1f2c97e18c1471978083dd12ac164d76b3689729d9649e7e325b33bc06714842ea19e3160a5a4a4dda11b29bcdb1a4b31ca562a6f4';

function normalizePmdPrice_(value) {
  if (value == null || value === '') return null;
  if (typeof value === 'number') return isNaN(value) ? null : value;
  var str = String(value).trim();
  if (!str) return null;
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
  var parsed = parseFloat(str);
  return isNaN(parsed) ? null : parsed;
}

/**
 * Helper para buscar as regras atuais diretamente da API PMD.
 */
function getPricingRuleUrl_(domainId) {
  try {
    const response = UrlFetchApp.fetch(PMD_API_URL + domainId, {
      method: 'GET',
      headers: { Authorization: 'Bearer ' + PMD_API_TOKEN },
      muteHttpExceptions: true,
    });
    if (response.getResponseCode() !== 200) {
      Logger.log(`[ERRO GET] Falha ao buscar regras para domínio ${domainId}. Resposta: ${response.getContentText()}`);
      return null; // Retorna null para indicar falha na busca
    }
    const data = JSON.parse(response.getContentText());
    return data.data?.attributes?.pricing_rule_url || [];
  } catch (e) {
    Logger.log(`[ERRO GET] Falha ao conectar na API para o domínio ${domainId}: ${e.message}`);
    return null; // Retorna null em caso de erro de conexão
  }
}

function shouldSyncRow_(rowValue, syncColumnIndex) {
  if (syncColumnIndex === -1) return true;
  var raw = rowValue == null ? '' : String(rowValue).trim().toLowerCase();
  if (!raw) return true;
  return ['sim', 'yes', 'y'].indexOf(raw) !== -1;
}

/**
 * Sincroniza de forma inteligente: busca, mescla e envia as regras.
 */
function runAdSeletoPricingUpdate() {
  try {
    const sheet = typeof ensurePricingSheetAdSeleto_ === 'function'
      ? ensurePricingSheetAdSeleto_()
      : SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Precificação - AdSeleto');

    if (!sheet) throw new Error('A aba "Precificação - AdSeleto" não foi encontrada.');

    const range = sheet.getDataRange();
    const allData = range.getValues();
    if (!allData || allData.length < 2) {
      return { ok: true, message: 'Nenhuma regra para sincronizar.' };
    }

    const headers = allData[0];
    const headerMap = typeof getHeaderIndexMap_ === 'function'
      ? getHeaderIndexMap_(headers)
      : headers.reduce(function(map, name, index) {
          map[String(name).trim()] = index;
          return map;
        }, {});

    const requiredHeaders = ['spnprice_id', 'domain_id', 'price_rule', 'utm_source', 'slot_id', 'url'];
    const missingHeaders = requiredHeaders.filter(function(name) { return !(name in headerMap); });
    if (missingHeaders.length) {
      throw new Error('Colunas obrigatórias ausentes na planilha da AdSeleto: ' + missingHeaders.join(', '));
    }

    const spnIdx = headerMap['spnprice_id'];
    const domainIdx = headerMap['domain_id'];
    const priceIdx = headerMap['price_rule'];
    const utmIdx = headerMap['utm_source'];
    const slotIdx = headerMap['slot_id'];
    const urlIdx = headerMap['url'];
    const syncColumnIndex = headerMap['Sincronizar'] != null ? headerMap['Sincronizar'] : headers.indexOf('Sincronizar');

    const rulesFromSheetByDomain = {};
    const skippedRows = [];

    for (let i = 1; i < allData.length; i++) {
      const row = allData[i];
      if (!shouldSyncRow_(syncColumnIndex === -1 ? '' : row[syncColumnIndex], syncColumnIndex)) {
        continue;
      }

      const domainId = row[domainIdx];
      const priceRuleValue = normalizePmdPrice_(row[priceIdx]);
      const utmSource = row[utmIdx];
      const slotId = row[slotIdx];
      const url = row[urlIdx];

      if (!domainId || priceRuleValue == null || !utmSource || !slotId || !url) {
        skippedRows.push(`Linha ${i + 1}: campos obrigatórios ausentes ou inválidos.`);
        continue;
      }

      if (!rulesFromSheetByDomain[domainId]) rulesFromSheetByDomain[domainId] = [];
      rulesFromSheetByDomain[domainId].push({
        spnprice_id: row[spnIdx],
        domain_id: domainId,
        price_rule: priceRuleValue,
        utm_source: utmSource,
        slot_id: slotId,
        url: url,
        _rowNum: i + 1,
      });
    }

    const domainsToUpdate = Object.keys(rulesFromSheetByDomain);
    if (!domainsToUpdate.length) {
      const baseMessage = 'Nenhuma regra válida encontrada para sincronizar.';
      return {
        ok: false,
        message: skippedRows.length
          ? baseMessage + '\n\nProblemas encontrados:\n- ' + skippedRows.join('\n- ')
          : baseMessage,
      };
    }

    const logMessages = [];

    domainsToUpdate.forEach(domainId => {
      Logger.log(`[INFO] Sincronizando domínio ${domainId}...`);

      // 1. GET - Busca regras existentes na plataforma
      const currentRulesFromAPI = getPricingRuleUrl_(domainId);
      if (currentRulesFromAPI === null) {
        logMessages.push(`- Domínio ${domainId}: ERRO! Não foi possível buscar as regras atuais da API. O domínio foi ignorado.`);
        return; // Pula para o próximo domínio
      }

      const finalRules = currentRulesFromAPI.slice(); // Cria uma cópia para trabalhar
      const rulesFromSheet = rulesFromSheetByDomain[domainId];

      // 2. MERGE - Mescla as regras da planilha com as da API
      rulesFromSheet.forEach(ruleFromSheet => {
        let found = false;
        for (let i = 0; i < finalRules.length; i++) {
          // Tenta encontrar por ID primeiro
          if (finalRules[i].spnprice_id && String(finalRules[i].spnprice_id) === String(ruleFromSheet.spnprice_id)) {
            finalRules[i].price_rule = ruleFromSheet.price_rule;
            found = true;
            break;
          }
        }
        if (!found) {
          for (let i = 0; i < finalRules.length; i++) {
            // Se não achou por ID, tenta pela combinação
            if (
              finalRules[i].url === ruleFromSheet.url &&
              finalRules[i].slot_id === ruleFromSheet.slot_id &&
              finalRules[i].utm_source === ruleFromSheet.utm_source
            ) {
              finalRules[i].price_rule = ruleFromSheet.price_rule;
              found = true;
              break;
            }
          }
        }
        if (!found) {
          // Se não encontrou de nenhuma forma, é uma regra nova
          finalRules.push({
            spnprice_id: ruleFromSheet.spnprice_id,
            domain_id: ruleFromSheet.domain_id,
            price_rule: ruleFromSheet.price_rule,
            utm_source: ruleFromSheet.utm_source,
            slot_id: ruleFromSheet.slot_id,
            url: ruleFromSheet.url,
          });
        }
      });

      const payload = { data: { pricing_rule_url: finalRules } };

      // 3. PUT - Envia a lista completa e mesclada de volta
      const response = UrlFetchApp.fetch(PMD_API_URL + domainId, {
        method: 'PUT',
        contentType: 'application/json',
        headers: { Authorization: 'Bearer ' + PMD_API_TOKEN },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true,
      });
      const responseCode = response.getResponseCode();
      if (responseCode >= 200 && responseCode < 300) {
        logMessages.push(`- Domínio ${domainId}: OK! Sincronizado com sucesso.`);
        if (syncColumnIndex !== -1) {
          rulesFromSheet.forEach(rule => {
            sheet.getRange(rule._rowNum, syncColumnIndex + 1).clearContent();
          });
        }
      } else {
        const errorText = response.getContentText();
        logMessages.push(`- Domínio ${domainId}: ERRO! (${responseCode}): ${errorText}`);
      }
    });

    if (skippedRows.length) {
      logMessages.push('Linhas ignoradas:\n- ' + skippedRows.join('\n- '));
    }

    return { ok: true, message: 'Resultado da Sincronização:\n\n' + logMessages.join('\n') };
  } catch (e) {
    Logger.log(`[ERRO FATAL] ${e.message}`);
    return { ok: false, message: `Ocorreu um erro fatal no script: ${e.message}` };
  }
}