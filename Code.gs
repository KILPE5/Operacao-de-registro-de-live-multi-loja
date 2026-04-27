/**
 * SOULBOM - SISTEMA DE APRESENTADORAS MULTI LOJA - V9
 * Arquivo único de backend: Code.gs
 *
 * Instalação:
 * 1. Cole este arquivo como Code.gs.
 * 2. Cole o outro arquivo como Index.html.
 * 3. Salve.
 * 4. Rode instalarSistemaApresentadorasMultiLoja().
 * 5. Implante como Web App.
 */

const APP = {
  DB_PROP: 'APRESENTADORAS_MULTI_LOJA_DATABASE_ID_V9',
  ADMIN_PIN_PROP: 'APRESENTADORAS_MULTI_LOJA_ADMIN_PIN_V9',
  ADMIN_PIN_DEFAULT: '1234',
  SNAPSHOT_FOLDER_NAME: 'Snapshots SoulBom Sistema Apresentadoras',
  CACHE_DASH_PREFIX: 'DASH_MULTI_V9_',
  CACHE_DASH_SECONDS: 20,
  TZ: Session.getScriptTimeZone() || 'America/Sao_Paulo',
  SHEETS: {
    STORES: 'Cfg_Lojas',
    HOSTS: 'Cfg_Apresentadoras',
    OPERATIONS: 'Operacoes_MultiLoja',
    SALES: 'Vendas_MultiLoja',
    AUDIT: 'Auditoria_MultiLoja',
    CONFIG: 'Configuracoes'
  }
};

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('SoulBom | Sistema de Apresentadoras')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function instalarSistemaApresentadorasMultiLoja() {
  setupSheets();
  return {
    ok: true,
    databaseUrl: getDb_().getUrl(),
    pinAdmin: getAdminPin_(),
    message: 'Sistema instalado com sucesso.'
  };
}

function setupSheets() {
  const ss = getDb_();

  createSheetIfNeeded_(ss, APP.SHEETS.CONFIG, ['CHAVE', 'VALOR', 'DESCRICAO', 'ATUALIZADO_EM']);
  createSheetIfNeeded_(ss, APP.SHEETS.STORES, ['ID', 'NOME', 'ATIVO', 'ORDEM', 'CRIADO_EM', 'ATUALIZADO_EM']);
  createSheetIfNeeded_(ss, APP.SHEETS.HOSTS, ['ID', 'NOME', 'ATIVO', 'CRIADO_EM', 'ATUALIZADO_EM']);
  createSheetIfNeeded_(ss, APP.SHEETS.OPERATIONS, [
    'ID', 'HOST', 'DATA_OPERACAO', 'INICIO', 'TERMINO', 'STATUS', 'MOTIVO_ENCERRAMENTO',
    'SNAPSHOT_URL', 'QTD_TOTAL', 'VALOR_TOTAL', 'LOJAS_JSON', 'CRIADO_EM', 'ATUALIZADO_EM'
  ]);
  createSheetIfNeeded_(ss, APP.SHEETS.SALES, [
    'ID_VENDA', 'OPERACAO_ID', 'DATA_OPERACAO', 'DATA_VENDA', 'HORA_VENDA',
    'LOJA', 'VALOR', 'HOST', 'OPERADOR', 'CRIADO_EM'
  ]);
  createSheetIfNeeded_(ss, APP.SHEETS.AUDIT, [
    'DATA_HORA', 'MODULO', 'ACAO', 'REGISTRO_ID', 'ANTES_JSON', 'DEPOIS_JSON', 'USUARIO'
  ]);

  ensureDefaultConfig_();
  ensureDefaultStores_();
  ensureDefaultHosts_();

  return true;
}

function getDb_() {
  const props = PropertiesService.getScriptProperties();
  const existingId = props.getProperty(APP.DB_PROP);

  if (existingId) {
    try {
      return SpreadsheetApp.openById(existingId);
    } catch (err) {
      props.deleteProperty(APP.DB_PROP);
    }
  }

  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active && active.getId()) {
      props.setProperty(APP.DB_PROP, active.getId());
      return active;
    }
  } catch (err) {}

  const created = SpreadsheetApp.create('Sistema Apresentadoras Multi Loja - Banco');
  props.setProperty(APP.DB_PROP, created.getId());
  return created;
}

function getSheet_(name) {
  let sheet = getDb_().getSheetByName(name);
  if (!sheet) {
    setupSheets();
    sheet = getDb_().getSheetByName(name);
  }
  return sheet;
}

function createSheetIfNeeded_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
    headers.forEach(function(header, i) {
      if (current[i] !== header) sheet.getRange(1, i + 1).setValue(header);
    });
  }

  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#111827')
    .setFontColor('#ffffff');

  sheet.setFrozenRows(1);
}

function ensureDefaultConfig_() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty(APP.ADMIN_PIN_PROP)) {
    props.setProperty(APP.ADMIN_PIN_PROP, APP.ADMIN_PIN_DEFAULT);
  }

  const sheet = getSheet_(APP.SHEETS.CONFIG);
  const rows = readRows_(sheet);
  if (!rows.some(function(r) { return String(r.CHAVE) === 'ADMIN_PIN'; })) {
    sheet.appendRow(['ADMIN_PIN', getAdminPin_(), 'PIN para acesso ao painel administrativo', new Date()]);
  }
}

function ensureDefaultStores_() {
  const defaults = ['ZAYA', 'Z.ZON'];
  const existing = getAllStores_();

  defaults.forEach(function(name, idx) {
    if (!existing.some(function(x) { return normalizeKey_(x.nome) === normalizeKey_(name); })) {
      const now = new Date();
      getSheet_(APP.SHEETS.STORES).appendRow([
        generateId_('STORE'), canonicalStoreName_(name), 'SIM', idx + 1, now, now
      ]);
    }
  });
}

function ensureDefaultHosts_() {
  const defaults = ['MORENA', 'YASMIN', 'BIA', 'ANA PAULA', 'VINI', 'RAPHAEL'];
  const existing = getAllHosts_();

  defaults.forEach(function(name) {
    if (!existing.some(function(x) { return normalizeKey_(x.nome) === normalizeKey_(name); })) {
      const now = new Date();
      getSheet_(APP.SHEETS.HOSTS).appendRow([
        generateId_('HOST'), canonicalHostName_(name), 'SIM', now, now
      ]);
    }
  });
}

function getAdminPin_() {
  return PropertiesService.getScriptProperties().getProperty(APP.ADMIN_PIN_PROP) || APP.ADMIN_PIN_DEFAULT;
}

function alterarPinAdmin(novoPin) {
  const pin = String(novoPin || '').trim();
  if (pin.length < 4) throw new Error('O PIN precisa ter pelo menos 4 caracteres.');

  PropertiesService.getScriptProperties().setProperty(APP.ADMIN_PIN_PROP, pin);

  const sheet = getSheet_(APP.SHEETS.CONFIG);
  const rows = readRowsWithIndex_(sheet);
  const found = rows.find(function(r) { return String(r.data.CHAVE) === 'ADMIN_PIN'; });

  if (found) {
    sheet.getRange(found.rowIndex, 2, 1, 3).setValues([[pin, 'PIN para acesso ao painel administrativo', new Date()]]);
  } else {
    sheet.appendRow(['ADMIN_PIN', pin, 'PIN para acesso ao painel administrativo', new Date()]);
  }

  return { ok: true, message: 'PIN alterado com sucesso.' };
}

function getBootstrapData() {
  try {
    setupSheets();
    return {
      ok: true,
      config: getAppConfig_(),
      operation: getOperationData(),
      serverTime: formatDateTime_(new Date(), APP.TZ),
      databaseUrl: getDb_().getUrl()
    };
  } catch (err) {
    return {
      ok: false,
      message: String(err && err.message ? err.message : err),
      config: fallbackConfig_(),
      operation: { active: null, recent: [], serverTime: formatDateTime_(new Date(), APP.TZ) },
      serverTime: formatDateTime_(new Date(), APP.TZ)
    };
  }
}

function fallbackConfig_() {
  return {
    stores: [
      { id: 'STORE-FALLBACK-1', nome: 'ZAYA', ativo: true, ordem: 1 },
      { id: 'STORE-FALLBACK-2', nome: 'Z.ZON', ativo: true, ordem: 2 }
    ],
    hosts: [
      { id: 'HOST-FALLBACK-1', nome: 'Morena', ativo: true },
      { id: 'HOST-FALLBACK-2', nome: 'Yasmin', ativo: true },
      { id: 'HOST-FALLBACK-3', nome: 'Bia', ativo: true }
    ],
    operationRefreshSeconds: 5,
    adminRefreshSeconds: 30,
    fallback: true
  };
}

function getAppConfig_() {
  return {
    stores: getActiveStores_(),
    hosts: getActiveHosts_(),
    operationRefreshSeconds: 5,
    adminRefreshSeconds: 30
  };
}

function adminLogin(pin) {
  if (String(pin || '') !== String(getAdminPin_())) {
    throw new Error('PIN inválido.');
  }
  return { ok: true };
}

function assertAdmin_(pin) {
  adminLogin(pin);
}

function getOperationData() {
  return {
    active: getActiveOperation_(),
    recent: getRecentClosedOperations_(15),
    serverTime: formatDateTime_(new Date(), APP.TZ)
  };
}

function startLive(data) {
  return withLock_('startLive', function() {
    setupSheets();

    if (getActiveOperation_()) throw new Error('Já existe uma live em andamento.');

    const host = canonicalHostName_(data && data.host);
    const opDate = String((data && data.opDate) || '').trim();

    if (!host) throw new Error('Selecione a apresentadora.');
    if (!opDate) throw new Error('Selecione a data.');

    const activeStores = getActiveStores_();
    if (!activeStores.length) throw new Error('Cadastre pelo menos uma loja ativa no admin.');

    const selectedStoreNames = Array.isArray(data.selectedStores)
      ? data.selectedStores.map(function(name) { return canonicalStoreName_(name); }).filter(Boolean)
      : [];

    const storesForThisLive = selectedStoreNames.length
      ? activeStores.filter(function(store) {
          return selectedStoreNames.some(function(name) {
            return normalizeKey_(name) === normalizeKey_(store.nome);
          });
        })
      : activeStores;

    if (!storesForThisLive.length) throw new Error('Selecione pelo menos uma loja para essa live.');

    const storeTotals = storesForThisLive.map(function(store) {
      return { id: store.id, nome: store.nome, quantidade: 0, valor: 0 };
    });

    const now = new Date();
    const row = [
      generateId_('LIVE'),
      host,
      opDate,
      Utilities.formatDate(now, APP.TZ, 'HH:mm'),
      '',
      'EM_ANDAMENTO',
      '',
      '',
      0,
      0,
      JSON.stringify(storeTotals),
      now,
      now
    ];

    getSheet_(APP.SHEETS.OPERATIONS).appendRow(row);
    registerAudit_('APRESENTADORAS', 'INICIAR_LIVE', row[0], null, rowToOperationObj_(row), getCurrentUserLabel_());
    clearDashCache_();

    return getOperationData();
  });
}

function registerSale(data) {
  return withLock_('registerSale', function() {
    const info = getActiveOperationRowInfo_();
    if (!info) throw new Error('Não existe live em andamento.');

    const storeName = canonicalStoreName_(data && data.storeName);
    const value = toNumber_(data && data.value);

    if (!storeName) throw new Error('Loja inválida.');
    if (value <= 0) throw new Error('Digite um valor de venda maior que zero.');

    const before = info.data;
    const now = new Date();
    const storeTotals = clone_(before.storeTotals || []);
    let found = false;

    storeTotals.forEach(function(item) {
      if (normalizeKey_(item.nome) === normalizeKey_(storeName)) {
        item.quantidade = Number(item.quantidade || 0) + 1;
        item.valor = round2_(Number(item.valor || 0) + value);
        found = true;
      }
    });

    if (!found) {
      storeTotals.push({ id: '', nome: storeName, quantidade: 1, valor: round2_(value) });
    }

    const totalQty = storeTotals.reduce(function(sum, item) { return sum + Number(item.quantidade || 0); }, 0);
    const totalValue = round2_(storeTotals.reduce(function(sum, item) { return sum + Number(item.valor || 0); }, 0));

    info.sheet.getRange(info.rowIndex, 9, 1, 5).setValues([[
      totalQty,
      totalValue,
      JSON.stringify(storeTotals),
      info.sheet.getRange(info.rowIndex, 12).getValue(),
      now
    ]]);

    getSheet_(APP.SHEETS.SALES).appendRow([
      generateId_('SALE'),
      before.id,
      before.opDate,
      before.opDate,
      Utilities.formatDate(now, APP.TZ, 'HH:mm'),
      storeName,
      value,
      before.host,
      getCurrentUserLabel_(),
      now
    ]);

    const after = findOperationById_(before.id).data;
    registerAudit_('APRESENTADORAS', 'REGISTRAR_VENDA', before.id, before, after, getCurrentUserLabel_());
    clearDashCache_();

    return getOperationData();
  });
}

function undoLastSale(data) {
  return withLock_('undoLastSale', function() {
    const info = getActiveOperationRowInfo_();
    if (!info) throw new Error('Não existe live em andamento.');

    const storeName = canonicalStoreName_(data && data.storeName);
    if (!storeName) throw new Error('Loja inválida.');

    const salesSheet = getSheet_(APP.SHEETS.SALES);
    const lastRow = salesSheet.getLastRow();
    if (lastRow <= 1) throw new Error('Não existe venda para desfazer.');

    const values = salesSheet.getRange(2, 1, lastRow - 1, 10).getValues();
    let foundRow = -1;
    let sale = null;

    for (let i = values.length - 1; i >= 0; i--) {
      if (String(values[i][1] || '') === String(info.data.id) &&
          normalizeKey_(values[i][5]) === normalizeKey_(storeName)) {
        foundRow = i + 2;
        sale = values[i];
        break;
      }
    }

    if (!sale) throw new Error('Nenhuma venda encontrada para essa loja.');

    const before = info.data;
    const saleValue = Number(sale[6] || 0);
    const storeTotals = clone_(before.storeTotals || []);

    storeTotals.forEach(function(item) {
      if (normalizeKey_(item.nome) === normalizeKey_(storeName)) {
        item.quantidade = Math.max(0, Number(item.quantidade || 0) - 1);
        item.valor = round2_(Math.max(0, Number(item.valor || 0) - saleValue));
      }
    });

    const totalQty = storeTotals.reduce(function(sum, item) { return sum + Number(item.quantidade || 0); }, 0);
    const totalValue = round2_(storeTotals.reduce(function(sum, item) { return sum + Number(item.valor || 0); }, 0));

    salesSheet.deleteRow(foundRow);
    info.sheet.getRange(info.rowIndex, 9, 1, 5).setValues([[
      totalQty,
      totalValue,
      JSON.stringify(storeTotals),
      info.sheet.getRange(info.rowIndex, 12).getValue(),
      new Date()
    ]]);

    const after = findOperationById_(before.id).data;
    registerAudit_('APRESENTADORAS', 'DESFAZER_VENDA', before.id, before, after, getCurrentUserLabel_());
    clearDashCache_();

    return getOperationData();
  });
}

function switchHost(data) {
  return withLock_('switchHost', function() {
    const active = getActiveOperation_();
    if (!active) throw new Error('Não existe live em andamento.');

    const newHost = canonicalHostName_(data && data.newHost);
    if (!newHost) throw new Error('Selecione a nova apresentadora.');

    closeActiveOperation_({
      reason: 'TROCA DE HOST',
      screenshotDataUrl: data && data.screenshotDataUrl
    });

    return startLive({
      host: newHost,
      opDate: active.opDate,
      selectedStores: (active.storeTotals || []).map(function(s) { return s.nome; })
    });
  });
}

function endLive(data) {
  return withLock_('endLive', function() {
    closeActiveOperation_({
      reason: String((data && data.reason) || 'ENCERRADO'),
      screenshotDataUrl: data && data.screenshotDataUrl
    });

    return getOperationData();
  });
}

function closeActiveOperation_(data) {
  const info = getActiveOperationRowInfo_();
  if (!info) throw new Error('Não existe live em andamento.');

  const now = new Date();
  let snapshotUrl = '';

  if (data && data.screenshotDataUrl) {
    snapshotUrl = saveSnapshot_(data.screenshotDataUrl, info.data.id, info.data.host, info.data.opDate);
  }

  info.sheet.getRange(info.rowIndex, 5, 1, 9).setValues([[
    Utilities.formatDate(now, APP.TZ, 'HH:mm'),
    'FECHADO',
    String((data && data.reason) || 'ENCERRADO'),
    snapshotUrl,
    Number(info.data.totalQty || 0),
    Number(info.data.totalValue || 0),
    JSON.stringify(info.data.storeTotals || []),
    info.sheet.getRange(info.rowIndex, 12).getValue(),
    now
  ]]);

  const after = findOperationById_(info.data.id).data;
  registerAudit_('APRESENTADORAS', 'FECHAR_LIVE', info.data.id, info.data, after, getCurrentUserLabel_());
  clearDashCache_();

  return after;
}

function getAdminBundle(pin, filter) {
  assertAdmin_(pin);
  setupSheets();

  const normalizedFilter = normalizeDashboardFilter_(filter);

  return {
    config: {
      stores: getAllStores_(),
      hosts: getAllHosts_()
    },
    dashboard: getDashboard_(normalizedFilter),
    operations: getLatestOperations_(25),
    sales: getSalesHistory_(normalizedFilter, 80),
    generatedAt: formatDateTime_(new Date(), APP.TZ)
  };
}

function getDashboard_(filter) {
  const normalized = normalizeDashboardFilter_(filter);
  const cacheKey = APP.CACHE_DASH_PREFIX + JSON.stringify(normalized);
  const cache = CacheService.getScriptCache();
  const cached = cache.get(cacheKey);

  if (cached) {
    try { return JSON.parse(cached); } catch (e) {}
  }

  const all = getAllOperations_();
  const filtered = all.filter(function(item) {
    return item.opDate >= normalized.dateFrom && item.opDate <= normalized.dateTo;
  });

  const summary = {
    dateFrom: normalized.dateFrom,
    dateTo: normalized.dateTo,
    blocks: filtered.length,
    sales: filtered.reduce(function(sum, x) { return sum + Number(x.totalQty || 0); }, 0),
    revenue: round2_(filtered.reduce(function(sum, x) { return sum + Number(x.totalValue || 0); }, 0)),
    avgTicket: 0
  };

  summary.avgTicket = summary.sales > 0 ? round2_(summary.revenue / summary.sales) : 0;

  const byStoreMap = {};
  const byHostMap = {};
  const dailyMap = {};

  filtered.forEach(function(item) {
    const hostKey = item.host || 'SEM NOME';

    if (!byHostMap[hostKey]) byHostMap[hostKey] = { nome: hostKey, vendas: 0, valor: 0, blocos: 0 };
    byHostMap[hostKey].vendas += Number(item.totalQty || 0);
    byHostMap[hostKey].valor += Number(item.totalValue || 0);
    byHostMap[hostKey].blocos += 1;

    if (!dailyMap[item.opDate]) dailyMap[item.opDate] = { data: item.opDate, blocos: 0, vendas: 0, valor: 0 };
    dailyMap[item.opDate].blocos += 1;
    dailyMap[item.opDate].vendas += Number(item.totalQty || 0);
    dailyMap[item.opDate].valor += Number(item.totalValue || 0);

    (item.storeTotals || []).forEach(function(store) {
      const storeKey = store.nome || 'SEM LOJA';

      if (!byStoreMap[storeKey]) byStoreMap[storeKey] = { nome: storeKey, vendas: 0, valor: 0, blocos: 0 };
      byStoreMap[storeKey].vendas += Number(store.quantidade || 0);
      byStoreMap[storeKey].valor += Number(store.valor || 0);

      if (Number(store.quantidade || 0) > 0 || Number(store.valor || 0) > 0) {
        byStoreMap[storeKey].blocos += 1;
      }
    });
  });

  const dashboard = {
    filter: normalized,
    summary: summary,
    byStore: Object.keys(byStoreMap).map(function(k) { return byStoreMap[k]; }).sort(rankSort_),
    byHost: Object.keys(byHostMap).map(function(k) { return byHostMap[k]; }).sort(rankSort_),
    daily: Object.keys(dailyMap).map(function(k) {
      dailyMap[k].valor = round2_(dailyMap[k].valor);
      return dailyMap[k];
    }).sort(function(a, b) { return a.data < b.data ? 1 : -1; })
  };

  cache.put(cacheKey, JSON.stringify(dashboard), APP.CACHE_DASH_SECONDS);
  return dashboard;
}

function normalizeDashboardFilter_(filter) {
  if (filter && filter.dateFrom && filter.dateTo) {
    return {
      mode: 'custom',
      dateFrom: String(filter.dateFrom),
      dateTo: String(filter.dateTo)
    };
  }

  const today = Utilities.formatDate(new Date(), APP.TZ, 'yyyy-MM-dd');
  const days = Math.max(1, Number(filter && filter.days ? filter.days : 7));

  return {
    mode: 'preset',
    days: days,
    dateFrom: addDaysToDateStr_(today, -(days - 1)),
    dateTo: today
  };
}

function getSalesHistory_(filter, limit) {
  const sheet = getSheet_(APP.SHEETS.SALES);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const normalized = normalizeDashboardFilter_(filter);
  const wanted = Math.max(1, Number(limit || 80));

  // Leitura otimizada: no admin comum, buscamos de baixo para cima e paramos cedo.
  // Isso evita carregar milhares de vendas no HTML toda vez que abre o painel.
  const totalRows = lastRow - 1;
  const maxScan = wanted >= 1000 ? totalRows : Math.min(totalRows, 1500);
  const startRow = Math.max(2, lastRow - maxScan + 1);
  const numRows = lastRow - startRow + 1;

  const values = sheet.getRange(startRow, 1, numRows, 10).getValues();
  const result = [];

  for (let i = values.length - 1; i >= 0; i--) {
    const item = rowToSaleObj_(values[i]);

    if (item.opDate >= normalized.dateFrom && item.opDate <= normalized.dateTo) {
      result.push(item);
      if (result.length >= wanted) break;
    }

    // Se a planilha estiver em ordem cronológica e já passamos muito do início do período,
    // dá para parar antes. Mantemos tolerância para planilhas editadas manualmente.
    if (result.length > 0 && item.opDate && item.opDate < normalized.dateFrom) {
      break;
    }
  }

  return result;
}


function exportCSV(pin, filter, stores, hosts, type) {
  assertAdmin_(pin);

  const normalized = normalizeDashboardFilter_(filter);
  const storeFilter = (stores || []).map(normalizeKey_);
  const hostFilter = (hosts || []).map(normalizeKey_);
  const exportType = String(type || 'sales').toLowerCase();

  function matchStore(name) {
    return !storeFilter.length || storeFilter.indexOf(normalizeKey_(name)) > -1;
  }

  function matchHost(name) {
    return !hostFilter.length || hostFilter.indexOf(normalizeKey_(name)) > -1;
  }

  let rows = [];

  if (exportType === 'sales') {
    rows.push(['Data Operação', 'Hora', 'Loja', 'Apresentadora', 'Valor', 'Operador', 'ID Venda', 'ID Bloco']);

    getSalesHistory_(normalized, 5000)
      .filter(function(sale) {
        return matchStore(sale.store) && matchHost(sale.host);
      })
      .forEach(function(sale) {
        rows.push([
          formatDateBR_(sale.opDate),
          sale.saleTime,
          sale.store,
          sale.host,
          Number(sale.value || 0).toFixed(2).replace('.', ','),
          sale.operator,
          sale.id,
          sale.operationId
        ]);
      });

  } else if (exportType === 'operations') {
    rows.push(['Data', 'Apresentadora', 'Início', 'Término', 'Status', 'Motivo', 'Vendas', 'Valor Total', 'Lojas']);

    getAllOperations_()
      .filter(function(operation) {
        return operation.opDate >= normalized.dateFrom && operation.opDate <= normalized.dateTo && matchHost(operation.host);
      })
      .forEach(function(operation) {
        const lojasResumo = (operation.storeTotals || [])
          .filter(function(store) { return matchStore(store.nome); })
          .map(function(store) {
            return store.nome + ': ' + store.quantidade + ' venda(s) / ' + Number(store.valor || 0).toFixed(2).replace('.', ',');
          })
          .join(' | ');

        rows.push([
          formatDateBR_(operation.opDate),
          operation.host,
          operation.startTime,
          operation.endTime || '',
          operation.status,
          operation.reason || '',
          operation.totalQty,
          Number(operation.totalValue || 0).toFixed(2).replace('.', ','),
          lojasResumo
        ]);
      });

  } else if (exportType === 'summary') {
    const dashboard = getDashboard_(normalized);

    rows.push(['Resumo por loja']);
    rows.push(['Loja', 'Vendas', 'Valor Total', 'Ticket Médio', 'Blocos com Venda']);

    (dashboard.byStore || [])
      .filter(function(store) { return matchStore(store.nome); })
      .forEach(function(store) {
        const ticket = Number(store.vendas || 0) > 0 ? round2_(Number(store.valor || 0) / Number(store.vendas || 0)) : 0;
        rows.push([
          store.nome,
          store.vendas,
          Number(store.valor || 0).toFixed(2).replace('.', ','),
          ticket.toFixed(2).replace('.', ','),
          store.blocos
        ]);
      });

    rows.push([]);
    rows.push(['Resumo por apresentadora']);
    rows.push(['Apresentadora', 'Vendas', 'Valor Total', 'Ticket Médio', 'Blocos']);

    (dashboard.byHost || [])
      .filter(function(host) { return matchHost(host.nome); })
      .forEach(function(host) {
        const ticket = Number(host.vendas || 0) > 0 ? round2_(Number(host.valor || 0) / Number(host.vendas || 0)) : 0;
        rows.push([
          host.nome,
          host.vendas,
          Number(host.valor || 0).toFixed(2).replace('.', ','),
          ticket.toFixed(2).replace('.', ','),
          host.blocos
        ]);
      });

  } else {
    throw new Error('Tipo de exportação inválido: ' + exportType);
  }

  return '\uFEFF' + rows.map(function(row) {
    return row.map(csvCell_).join(',');
  }).join('\r\n');
}

function csvCell_(value) {
  const text = String(value === null || value === undefined ? '' : value);
  if (/[",\n\r]/.test(text)) {
    return '"' + text.replace(/"/g, '""') + '"';
  }
  return text;
}


function saveStoreAdmin(data, pin) {
  return withLock_('saveStoreAdmin', function() {
    assertAdmin_(pin);

    const name = canonicalStoreName_(data && data.name);
    if (!name) throw new Error('Digite o nome da loja.');

    const all = getAllStores_();
    if (all.some(function(item) {
      return normalizeKey_(item.nome) === normalizeKey_(name) && String(item.id) !== String(data.id || '');
    })) {
      throw new Error('Já existe uma loja com esse nome.');
    }

    if (data.id) {
      const info = findStoreById_(data.id);
      if (!info) throw new Error('Loja não encontrada.');

      const before = info.data;
      info.sheet.getRange(info.rowIndex, 2, 1, 5).setValues([[
        name,
        data.active ? 'SIM' : 'NAO',
        Number(data.order || 0) || 0,
        info.sheet.getRange(info.rowIndex, 5).getValue(),
        new Date()
      ]]);

      const after = findStoreById_(data.id).data;
      registerAudit_('CONFIG', 'ATUALIZAR_LOJA', data.id, before, after, getCurrentUserLabel_());
    } else {
      const now = new Date();
      const row = [
        generateId_('STORE'),
        name,
        data.active === false ? 'NAO' : 'SIM',
        Number(data.order || 0) || (all.length + 1),
        now,
        now
      ];

      getSheet_(APP.SHEETS.STORES).appendRow(row);
      registerAudit_('CONFIG', 'CRIAR_LOJA', row[0], null, rowToStoreObj_(row), getCurrentUserLabel_());
    }

    clearDashCache_();
    return getAdminBundle(pin, data.dashboardFilter || {});
  });
}

function deleteStoreAdmin(id, pin, dashboardFilter) {
  return withLock_('deleteStoreAdmin', function() {
    assertAdmin_(pin);

    const info = findStoreById_(id);
    if (!info) throw new Error('Loja não encontrada.');

    const active = getActiveOperation_();
    if (active && (active.storeTotals || []).some(function(store) {
      return normalizeKey_(store.nome) === normalizeKey_(info.data.nome);
    })) {
      throw new Error('Não é possível excluir uma loja que está na live em andamento. Encerre a live ou desative a loja.');
    }

    registerAudit_('CONFIG', 'EXCLUIR_LOJA', id, info.data, null, getCurrentUserLabel_());
    info.sheet.deleteRow(info.rowIndex);
    clearDashCache_();

    return getAdminBundle(pin, dashboardFilter || {});
  });
}

function saveHostAdmin(data, pin) {
  return withLock_('saveHostAdmin', function() {
    assertAdmin_(pin);

    const name = canonicalHostName_(data && data.name);
    if (!name) throw new Error('Digite o nome da apresentadora.');

    const all = getAllHosts_();
    if (all.some(function(item) {
      return normalizeKey_(item.nome) === normalizeKey_(name) && String(item.id) !== String(data.id || '');
    })) {
      throw new Error('Já existe uma apresentadora com esse nome.');
    }

    if (data.id) {
      const info = findHostById_(data.id);
      if (!info) throw new Error('Apresentadora não encontrada.');

      const before = info.data;
      info.sheet.getRange(info.rowIndex, 2, 1, 4).setValues([[
        name,
        data.active ? 'SIM' : 'NAO',
        info.sheet.getRange(info.rowIndex, 4).getValue(),
        new Date()
      ]]);

      const after = findHostById_(data.id).data;
      registerAudit_('CONFIG', 'ATUALIZAR_HOST', data.id, before, after, getCurrentUserLabel_());
    } else {
      const now = new Date();
      const row = [
        generateId_('HOST'),
        name,
        data.active === false ? 'NAO' : 'SIM',
        now,
        now
      ];

      getSheet_(APP.SHEETS.HOSTS).appendRow(row);
      registerAudit_('CONFIG', 'CRIAR_HOST', row[0], null, rowToHostObj_(row), getCurrentUserLabel_());
    }

    clearDashCache_();
    return getAdminBundle(pin, data.dashboardFilter || {});
  });
}

function deleteHostAdmin(id, pin, dashboardFilter) {
  return withLock_('deleteHostAdmin', function() {
    assertAdmin_(pin);

    const info = findHostById_(id);
    if (!info) throw new Error('Apresentadora não encontrada.');

    const active = getActiveOperation_();
    if (active && normalizeKey_(active.host) === normalizeKey_(info.data.nome)) {
      throw new Error('Não é possível excluir a apresentadora da live em andamento. Encerre a live ou troque a apresentadora.');
    }

    registerAudit_('CONFIG', 'EXCLUIR_HOST', id, info.data, null, getCurrentUserLabel_());
    info.sheet.deleteRow(info.rowIndex);
    clearDashCache_();

    return getAdminBundle(pin, dashboardFilter || {});
  });
}

function updateOperationAdmin(data, pin) {
  return withLock_('updateOperationAdmin', function() {
    assertAdmin_(pin);

    const info = findOperationById_(data && data.id);
    if (!info) throw new Error('Bloco não encontrado.');

    const host = canonicalHostName_(data.host);
    const opDate = String(data.opDate || '').trim();
    const status = String(data.status || '').trim().toUpperCase();

    if (!host) throw new Error('Host inválido.');
    if (!opDate) throw new Error('Data inválida.');
    if (['EM_ANDAMENTO', 'FECHADO'].indexOf(status) === -1) throw new Error('Status inválido.');

    const active = getActiveOperation_();
    if (status === 'EM_ANDAMENTO' && active && String(active.id) !== String(data.id)) {
      throw new Error('Já existe outro bloco em andamento. Feche antes de marcar este como em andamento.');
    }

    const cleanedStoreTotals = (Array.isArray(data.storeTotals) ? data.storeTotals : [])
      .map(function(item) {
        const name = canonicalStoreName_(item.nome);
        return {
          id: String(item.id || ''),
          nome: name,
          quantidade: Math.max(0, Number(item.quantidade || 0)),
          valor: round2_(toNumber_(item.valor))
        };
      })
      .filter(function(item) { return !!item.nome; });

    const totalQty = cleanedStoreTotals.reduce(function(sum, item) { return sum + Number(item.quantidade || 0); }, 0);
    const totalValue = round2_(cleanedStoreTotals.reduce(function(sum, item) { return sum + Number(item.valor || 0); }, 0));

    const before = info.data;
    info.sheet.getRange(info.rowIndex, 2, 1, 12).setValues([[
      host,
      opDate,
      String(data.startTime || '').trim(),
      String(data.endTime || '').trim(),
      status,
      String(data.reason || '').trim(),
      String(data.snapshotUrl || ''),
      totalQty,
      totalValue,
      JSON.stringify(cleanedStoreTotals),
      info.sheet.getRange(info.rowIndex, 12).getValue(),
      new Date()
    ]]);

    syncSalesWithOperation_(data.id, opDate, host);

    const after = findOperationById_(data.id).data;
    registerAudit_('APRESENTADORAS', 'ADMIN_ATUALIZAR_BLOCO', data.id, before, after, getCurrentUserLabel_());
    clearDashCache_();

    return getAdminBundle(pin, data.dashboardFilter || {});
  });
}

function deleteOperationAdmin(id, pin, dashboardFilter) {
  return withLock_('deleteOperationAdmin', function() {
    assertAdmin_(pin);

    const info = findOperationById_(id);
    if (!info) throw new Error('Bloco não encontrado.');

    deleteSalesByOperationId_(id);
    registerAudit_('APRESENTADORAS', 'EXCLUIR_BLOCO', id, info.data, null, getCurrentUserLabel_());
    info.sheet.deleteRow(info.rowIndex);
    clearDashCache_();

    return getAdminBundle(pin, dashboardFilter || {});
  });
}

function syncSalesWithOperation_(operationId, opDate, host) {
  const sheet = getSheet_(APP.SHEETS.SALES);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  const values = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][1] || '') === String(operationId)) {
      sheet.getRange(i + 2, 3, 1, 6).setValues([[
        opDate,
        opDate,
        values[i][4],
        values[i][5],
        values[i][6],
        host
      ]]);
    }
  }
}

function deleteSalesByOperationId_(operationId) {
  const sheet = getSheet_(APP.SHEETS.SALES);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  const values = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (String(values[i][1] || '') === String(operationId)) {
      sheet.deleteRow(i + 2);
    }
  }
}

function getRecentClosedOperations_(limit) {
  return getAllOperations_()
    .filter(function(item) { return item.status === 'FECHADO'; })
    .slice(0, limit || 10);
}

function getLatestOperations_(limit) {
  return getAllOperations_().slice(0, limit || 100);
}

function getAllOperations_() {
  const sheet = getSheet_(APP.SHEETS.OPERATIONS);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  return sheet.getRange(2, 1, lastRow - 1, 13).getValues()
    .map(rowToOperationObj_)
    .sort(sortOperationsDesc_);
}

function getActiveOperation_() {
  const info = getActiveOperationRowInfo_();
  return info ? info.data : null;
}

function getActiveOperationRowInfo_() {
  const sheet = getSheet_(APP.SHEETS.OPERATIONS);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return null;

  const values = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (String(values[i][5] || '') === 'EM_ANDAMENTO') {
      return {
        sheet: sheet,
        rowIndex: i + 2,
        data: rowToOperationObj_(values[i])
      };
    }
  }
  return null;
}

function findOperationById_(id) {
  const sheet = getSheet_(APP.SHEETS.OPERATIONS);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return null;

  const values = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '') === String(id || '')) {
      return {
        sheet: sheet,
        rowIndex: i + 2,
        data: rowToOperationObj_(values[i])
      };
    }
  }
  return null;
}

function rowToOperationObj_(row) {
  let stores = [];
  try {
    stores = JSON.parse(String(row[10] || '[]'));
  } catch (e) {
    stores = [];
  }

  stores = stores.map(function(item) {
    return {
      id: String(item.id || ''),
      nome: String(item.nome || ''),
      quantidade: Number(item.quantidade || 0),
      valor: Number(item.valor || 0)
    };
  }).sort(function(a, b) {
    return String(a.nome).localeCompare(String(b.nome), 'pt-BR');
  });

  return {
    id: String(row[0] || ''),
    host: String(row[1] || ''),
    opDate: formatDateToInput_(row[2], APP.TZ),
    startTime: formatTimeValue_(row[3], APP.TZ),
    endTime: formatTimeValue_(row[4], APP.TZ),
    status: String(row[5] || ''),
    reason: String(row[6] || ''),
    snapshotUrl: String(row[7] || ''),
    totalQty: Number(row[8] || 0),
    totalValue: Number(row[9] || 0),
    storeTotals: stores,
    createdAt: formatDateTime_(row[11], APP.TZ),
    updatedAt: formatDateTime_(row[12], APP.TZ)
  };
}

function rowToSaleObj_(row) {
  return {
    id: String(row[0] || ''),
    operationId: String(row[1] || ''),
    opDate: formatDateToInput_(row[2], APP.TZ),
    saleDate: formatDateToInput_(row[3], APP.TZ),
    saleTime: formatTimeValue_(row[4], APP.TZ),
    store: String(row[5] || ''),
    value: Number(row[6] || 0),
    host: String(row[7] || ''),
    operator: String(row[8] || ''),
    createdAt: formatDateTime_(row[9], APP.TZ)
  };
}

function getAllStores_() {
  const sheet = getSheet_(APP.SHEETS.STORES);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  return sheet.getRange(2, 1, lastRow - 1, 6).getValues()
    .map(rowToStoreObj_)
    .sort(function(a, b) {
      if (Number(a.ordem || 0) !== Number(b.ordem || 0)) return Number(a.ordem || 0) - Number(b.ordem || 0);
      return String(a.nome).localeCompare(String(b.nome), 'pt-BR');
    });
}

function getActiveStores_() {
  return getAllStores_().filter(function(item) { return item.ativo; });
}

function findStoreById_(id) {
  const sheet = getSheet_(APP.SHEETS.STORES);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return null;

  const values = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '') === String(id || '')) {
      return { sheet: sheet, rowIndex: i + 2, data: rowToStoreObj_(values[i]) };
    }
  }
  return null;
}

function rowToStoreObj_(row) {
  return {
    id: String(row[0] || ''),
    nome: String(row[1] || ''),
    ativo: String(row[2] || 'SIM') === 'SIM',
    ordem: Number(row[3] || 0),
    createdAt: formatDateTime_(row[4], APP.TZ),
    updatedAt: formatDateTime_(row[5], APP.TZ)
  };
}

function getAllHosts_() {
  const sheet = getSheet_(APP.SHEETS.HOSTS);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  return sheet.getRange(2, 1, lastRow - 1, 5).getValues()
    .map(rowToHostObj_)
    .sort(function(a, b) {
      return String(a.nome).localeCompare(String(b.nome), 'pt-BR');
    });
}

function getActiveHosts_() {
  return getAllHosts_().filter(function(item) { return item.ativo; });
}

function findHostById_(id) {
  const sheet = getSheet_(APP.SHEETS.HOSTS);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return null;

  const values = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '') === String(id || '')) {
      return { sheet: sheet, rowIndex: i + 2, data: rowToHostObj_(values[i]) };
    }
  }
  return null;
}

function rowToHostObj_(row) {
  return {
    id: String(row[0] || ''),
    nome: String(row[1] || ''),
    ativo: String(row[2] || 'SIM') === 'SIM',
    createdAt: formatDateTime_(row[3], APP.TZ),
    updatedAt: formatDateTime_(row[4], APP.TZ)
  };
}

function readRows_(sheet) {
  if (!sheet || sheet.getLastRow() <= 1) return [];
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(String);

  return values.slice(1)
    .filter(function(row) {
      return row.some(function(cell) { return String(cell).trim() !== ''; });
    })
    .map(function(row) {
      const obj = {};
      headers.forEach(function(h, i) { obj[h] = row[i]; });
      return obj;
    });
}

function readRowsWithIndex_(sheet) {
  if (!sheet || sheet.getLastRow() <= 1) return [];
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(String);

  return values.slice(1)
    .map(function(row, idx) {
      const obj = {};
      headers.forEach(function(h, i) { obj[h] = row[i]; });
      return { rowIndex: idx + 2, data: obj };
    })
    .filter(function(item) {
      return Object.keys(item.data).some(function(k) {
        return String(item.data[k]).trim() !== '';
      });
    });
}

function registerAudit_(module, action, recordId, before, after, user) {
  getSheet_(APP.SHEETS.AUDIT).appendRow([
    new Date(),
    module,
    action,
    recordId || '',
    before ? JSON.stringify(before) : '',
    after ? JSON.stringify(after) : '',
    user || ''
  ]);
}

function getCurrentUserLabel_() {
  try {
    return Session.getActiveUser().getEmail() || 'OPERADOR';
  } catch (e) {
    return 'OPERADOR';
  }
}

function withLock_(label, callback) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(25000);
    return callback();
  } catch (err) {
    throw new Error(err && err.message ? err.message : 'Erro em ' + label);
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function generateId_(prefix) {
  return prefix + '-' + Utilities.formatDate(new Date(), APP.TZ, 'yyyyMMdd-HHmmss') + '-' + Math.floor(Math.random() * 9000 + 1000);
}

function clearDashCache_() {}

function clone_(obj) {
  return JSON.parse(JSON.stringify(obj || null));
}

function normalizeKey_(text) {
  return String(text || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, '')
    .trim();
}

function canonicalHostName_(host) {
  const raw = String(host || '').trim();
  if (!raw) return '';

  return raw
    .toLowerCase()
    .split(' ')
    .filter(Boolean)
    .map(function(part) {
      return part.charAt(0).toUpperCase() + part.slice(1);
    })
    .join(' ');
}

function canonicalStoreName_(store) {
  const raw = String(store || '').trim();
  if (!raw) return '';
  return raw.toUpperCase();
}

function toNumber_(value) {
  if (typeof value === 'number') return value;
  if (value === null || value === undefined || value === '') return 0;

  let s = String(value).trim().replace(/\s/g, '');

  if (s.includes(',') && s.includes('.')) {
    if (s.lastIndexOf(',') > s.lastIndexOf('.')) {
      s = s.replace(/\./g, '').replace(',', '.');
    } else {
      s = s.replace(/,/g, '');
    }
  } else if (s.includes(',')) {
    s = s.replace(',', '.');
  }

  s = s.replace(/[^\d.-]/g, '');
  return Number(s) || 0;
}

function round2_(v) {
  return Math.round((Number(v) || 0) * 100) / 100;
}

function addDaysToDateStr_(dateStr, days) {
  const d = parseDateStrSafe_(dateStr);
  d.setDate(d.getDate() + Number(days || 0));
  return Utilities.formatDate(d, APP.TZ, 'yyyy-MM-dd');
}

function parseDateStrSafe_(dateStr) {
  const parts = String(dateStr || '').split('-');
  if (parts.length !== 3) return new Date();
  return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]), 12, 0, 0, 0);
}

function sortOperationsDesc_(a, b) {
  const aKey = String(a.opDate || '') + ' ' + String(a.startTime || '');
  const bKey = String(b.opDate || '') + ' ' + String(b.startTime || '');
  return aKey < bKey ? 1 : -1;
}

function rankSort_(a, b) {
  if (Number(b.vendas || 0) !== Number(a.vendas || 0)) return Number(b.vendas || 0) - Number(a.vendas || 0);
  if (Number(b.valor || 0) !== Number(a.valor || 0)) return Number(b.valor || 0) - Number(a.valor || 0);
  return String(a.nome || '').localeCompare(String(b.nome || ''), 'pt-BR');
}

function saveSnapshot_(dataUrl, id, host, date) {
  const raw = String(dataUrl || '');
  const match = raw.match(/^data:(.*?);base64,(.*)$/);
  if (!match) return '';

  const folder = getOrCreateFolder_(APP.SNAPSHOT_FOLDER_NAME);
  const blob = Utilities.newBlob(Utilities.base64Decode(match[2]), 'image/png', id + '.png');
  const file = folder.createFile(blob);
  file.setName(id + ' - ' + host + ' - ' + date + '.png');

  return file.getUrl();
}

function getOrCreateFolder_(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(folderName);
}


function formatDateBR_(dateStr) {
  if (!dateStr) return '';
  const parts = String(dateStr).split('-');
  return parts.length === 3 ? parts[2] + '/' + parts[1] + '/' + parts[0] : String(dateStr);
}


function formatDateToInput_(value, timezone) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return Utilities.formatDate(value, timezone, 'yyyy-MM-dd');
  }

  const tentative = new Date(value);
  if (!isNaN(tentative)) return Utilities.formatDate(tentative, timezone, 'yyyy-MM-dd');

  return String(value);
}

function formatDateTime_(value, timezone) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return Utilities.formatDate(value, timezone, 'dd/MM/yyyy HH:mm:ss');
  }

  const tentative = new Date(value);
  if (!isNaN(tentative)) return Utilities.formatDate(tentative, timezone, 'dd/MM/yyyy HH:mm:ss');

  return String(value);
}

function formatTimeValue_(value, timezone) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return Utilities.formatDate(value, timezone, 'HH:mm');
  }

  const text = String(value).trim();
  if (/^\d{1,2}:\d{2}$/.test(text)) return text.padStart(5, '0');

  const tentative = new Date(value);
  if (!isNaN(tentative)) return Utilities.formatDate(tentative, timezone, 'HH:mm');

  return text;
}


function pingBackend() {
  return {
    ok: true,
    message: 'Backend SoulBom ativo.',
    time: formatDateTime_(new Date(), APP.TZ)
  };
}


function getSystemVersion() {
  return {
    ok: true,
    name: 'SoulBom',
    version: 'V9 Admin Leve',
    time: formatDateTime_(new Date(), APP.TZ)
  };
}
