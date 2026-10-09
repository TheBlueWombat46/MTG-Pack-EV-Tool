/**
 * MTG Pack ROI Tracker: live price refresh for Google Sheets.
 *
 * Pulls current card prices for one Magic: The Gathering set from the
 * Scryfall API into CARD_DATA. Sheet formulas then combine those prices
 * with booster pull rates to estimate the theoretical and realizable
 * expected value (EV) of opening a pack versus selling it sealed.
 */

const SCRYFALL_HEADERS = {
  'User-Agent': 'MTG-Pack-ROI-Google-Sheets/1.0',
  'Accept': 'application/json;q=0.9,*/*;q=0.8'
};

const FIRST_DATA_ROW = 4;  // Rows 1-3 hold the title and headers
const DATA_COLUMNS = 21;   // CARD_DATA columns A:U
const HEADER_ROW = 3;  // FIRST_DATA_ROW (4) already exists in your file

// POOL_CONFIG columns that describe a pool rather than filter cards
const POOL_META_HEADERS = ['pool id', 'pool type', 'notes', 'card count'];

// CARD_DATA columns copied into BOOSTER_POOL, in output order
const POOL_OUTPUT_HEADERS = [
  'UUID / Scryfall ID', 'Finish', 'Set Code', 'Collector #',
  'Card Name', 'Market Price ($)', 'Realizable Price ($)'
];

// CARD_DATA columns that hold comma-separated lists
const LIST_COLUMNS = ['frame effect', 'promo type', 'set code', 'rarity', 'finish'];

// Price fields Scryfall returns for each finish
const FINISHES = [
  { finish: 'normal', priceKey: 'usd' },
  { finish: 'foil', priceKey: 'usd_foil' },
  { finish: 'etched', priceKey: 'usd_etched' }
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('MTG Tracker')
    .addItem('Refresh Selected Set', 'refreshSelectedSet')
    .addItem('Build Booster Pools', 'buildBoosterPool')
    .addItem('Build Pool Options', 'buildPoolOptionLists')
    .addItem('Create Daily Price Refresh', 'createDailyTrigger')
    .addToUi();
}

/**
 * Fetches every paper printing in the set chosen in SETTINGS!B4 and
 * rewrites CARD_DATA with one row per card per priced finish.
 */
function refreshSelectedSet() {
  const ss = SpreadsheetApp.getActive();
  const settings = ss.getSheetByName('SETTINGS');
  const cardSheet = ss.getSheetByName('CARD_DATA');

  const setCode = String(settings.getRange('B4').getValue()).trim().toLowerCase();
  if (!setCode || setCode === 'demo') {
    throw new Error('Enter a real Scryfall set code in SETTINGS!B4 before refreshing.');
  }

  let url = 'https://api.scryfall.com/cards/search?unique=prints&order=set&q=' +
    encodeURIComponent('set:' + setCode + ' game:paper');

  const rows = [];
  const now = new Date();

  // Scryfall paginates results; follow next_page until has_more is false
  while (url) {
    const payload = JSON.parse(fetchScryfall_(url).getContentText());

    payload.data.forEach(card => {
      FINISHES.forEach(({ finish, priceKey }) => {
        const price = card.prices && card.prices[priceKey];
        if (price) rows.push(makeRow_(card, finish, price, now));
      });
    });

    url = payload.has_more ? payload.next_page : null;
    if (url) Utilities.sleep(500);  // Stay well within Scryfall's rate limits
  }

  if (rows.length === 0) {
    throw new Error('No priced paper cards found for ' + setCode.toUpperCase());
  }

  // Make sure the sheet has enough rows for the new dataset
  const neededRows = rows.length + FIRST_DATA_ROW - 1;
  const existingRows = cardSheet.getMaxRows();
  if (neededRows > existingRows) {
    cardSheet.insertRowsAfter(existingRows, neededRows - existingRows);
  }

  // Clear the previous dataset, then write columns A:U
  cardSheet
    .getRange(FIRST_DATA_ROW, 1, cardSheet.getMaxRows() - FIRST_DATA_ROW + 1, DATA_COLUMNS)
    .clearContent();
  cardSheet.getRange(FIRST_DATA_ROW, 1, rows.length, DATA_COLUMNS).setValues(rows);

  // Realizable price (column Q), calculated from market price (column P):
  // cards below the minimum listing value are worth $0; everything else is
  // reduced by the marketplace fee and sale haircut, minus fulfillment cost.
  const formulas = [];
  for (let r = FIRST_DATA_ROW; r < FIRST_DATA_ROW + rows.length; r++) {
    formulas.push([
      '=IF(P' + r + '<SETTINGS!$B$9,0,' +
      'MAX(0,P' + r + '*(1-SETTINGS!$B$6)*(1-SETTINGS!$B$7)-SETTINGS!$B$8))'
    ]);
  }
  cardSheet.getRange(FIRST_DATA_ROW, 17, rows.length, 1).setFormulas(formulas);

  // Format market and realizable prices (P:Q) and the timestamp (T)
  cardSheet.getRange(FIRST_DATA_ROW, 16, rows.length, 2)
    .setNumberFormat('$#,##0.00;[Red]($#,##0.00);-');
  cardSheet.getRange(FIRST_DATA_ROW, 20, rows.length, 1)
    .setNumberFormat('yyyy-mm-dd hh:mm');

  settings.getRange('B12').setValue(now).setNumberFormat('yyyy-mm-dd hh:mm');
  SpreadsheetApp.flush();

  buildPoolOptionLists();
}

function buildBoosterPool() {
  const ss = SpreadsheetApp.getActive();

  // 1. READ: pull both sheets into memory
  const cards = readTable_(ss.getSheetByName('CARD_DATA'));
  const pools = readTable_(ss.getSheetByName('POOL_CONFIG'));

  // 2. MAP: header name -> column index
  const cardMap = makeHeaderMap_(cards.headers);

  // 3. PARSE: turn POOL_CONFIG rows into pool definition objects
  const poolDefs = parsePoolDefs_(pools.headers, pools.rows, cardMap);

  const testPool = poolDefs.find(p => p.poolId === 'M21_DRAFT_COMMONS');
  testPool.criteria.forEach(c => {
    const passing = cards.rows.filter(card => cardMatches_(card, [c])).length;
    Logger.log(cards.headers[c.col] + ' | ' + c.mode + ' | ' + c.values.join(',') + ' -> ' + passing + ' cards pass');
  });

  // 4. MATCH: find every card that belongs to every pool
  const result = buildPools_(cardMap, cards.rows, poolDefs);

  // 5-6. CHECK and WRITE
  writeBoosterPool_(ss.getSheetByName('BOOSTER_POOL'), result.outputRows);
  writePoolCounts_(ss.getSheetByName('POOL_CONFIG'), pools, result.counts);

  Logger.log(JSON.stringify(poolDefs, null, 2));
}

/**
 * Converts one Scryfall card and finish into a CARD_DATA row (A:U).
 * Column Q is left blank because refreshSelectedSet() writes its formula.
 */
function makeRow_(card, finish, price, now) {
  const typeLine = card.type_line || '';
  const promoTypes = card.promo_types || [];

  return [
    card.id,                                            // A - Scryfall ID
    String(card.set || '').toUpperCase(),               // B - Set Code
    card.collector_number || '',                        // C - Collector #
    card.name || '',                                    // D - Card Name
    card.rarity || '',                                  // E - Rarity
    finish,                                             // F - Finish
    card.border_color || '',                            // G - Border Color
    (card.frame_effects || []).join(', '),              // H - Frame Effects
    promoTypes.join(', '),                              // I - Promo Types
    card.full_art || false,                             // J - Full Art
    card.booster || false,                              // K - Booster Eligible
    card.booster || (promoTypes.includes('boosterfun') &&
      !(card.frame_effects || []).includes('extendedart')),// L - Pullable in Pack
    card.layout || '',                                  // M - Layout
    typeLine,                                           // N - Type Line
    typeLine.includes('Basic Land'),                    // O - Basic Land?
    Number(price),                                      // P - Market Price
    '',                                                 // Q - Realizable Price
    'Scryfall',                                         // R - Price Provider
    'USD',                                              // S - Currency
    now,                                                // T - Last Updated
    card.scryfall_uri || ''                             // U - Source URL
  ];
}

function readTable_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const values = sheet.getRange(HEADER_ROW, 1, lastRow - HEADER_ROW + 1, lastCol).getValues();

  return {
    headers: values[0].map(h => String(h).trim()),
    rows: values.slice(1)
  };
}

function makeHeaderMap_(headers) {
  const map = {};
  headers.forEach((header, index) => {
    if (header) map[normalize_(header)] = index;
  });
  return map;
}

function parseCriterion_(cellValue, cardCol, isList) {
  const crit = { col: cardCol, isList, include: [], exclude: [], includeEmpty: false };

  normalize_(cellValue).split(',').map(v => v.trim()).filter(v => v).forEach(v => {
    if (v === '(none)') crit.includeEmpty = true;
    else if (v.startsWith('not ')) crit.exclude.push(v.slice(4).trim());
    else crit.include.push(v);
  });
  return crit;
}

function parsePoolDefs_(poolHeaders, poolRows, cardMap) {
  const poolIdCol = poolHeaders.findIndex(h => normalize_(h) === 'pool id');
  if (poolIdCol === -1) throw new Error('POOL_CONFIG needs a "Pool ID" column.');

  // Decide once which POOL_CONFIG columns are filters
  const criteriaCols = [];
  poolHeaders.forEach((header, i) => {
    const key = normalize_(header);
    if (!key || POOL_META_HEADERS.includes(key)) return;
    if (!(key in cardMap)) {
      throw new Error('POOL_CONFIG column "' + header + '" does not match any CARD_DATA header.');
    }
    criteriaCols.push({ poolCol: i, cardCol: cardMap[key], isList: LIST_COLUMNS.includes(key) });
  });

  // Turn each row into a pool definition
  const defs = [];
  const seen = new Set();
  poolRows.forEach((row, r) => {
    const poolId = String(row[poolIdCol]).trim();
    if (!poolId) return;  // skip blank rows
    const sheetRow = r + FIRST_DATA_ROW;
    if (seen.has(poolId)) throw new Error('Duplicate Pool ID "' + poolId + '" on POOL_CONFIG row ' + sheetRow);
    seen.add(poolId);

    const criteria = criteriaCols
      .filter(c => normalize_(row[c.poolCol]) !== '')
      .map(c => parseCriterion_(row[c.poolCol], c.cardCol, c.isList));
    if (criteria.length === 0) throw new Error('Pool "' + poolId + '" has no criteria and would match every card.');

    defs.push({ poolId, criteria });
  });
  return defs;
}

function cardMatches_(card, criteria) {
  return criteria.every(c => {
    const cell = normalize_(card[c.col]);
    const tokens = c.isList ? cell.split(',').map(t => t.trim()).filter(t => t) : [cell];

    const hasIncludes = c.include.length > 0 || c.includeEmpty;
    const included = !hasIncludes
      || (c.includeEmpty && cell === '')
      || c.include.some(v => tokens.includes(v));
    const excluded = c.exclude.some(v => tokens.includes(v));

    return included && !excluded;
  });
}

function buildPools_(cardMap, cardRows, poolDefs) {
  const outCols = POOL_OUTPUT_HEADERS.map(h => {
    const i = cardMap[normalize_(h)];
    if (i === undefined) throw new Error('CARD_DATA is missing the "' + h + '" column.');
    return i;
  });

  const outputRows = [];
  const counts = {};

  poolDefs.forEach(pool => {
    counts[pool.poolId] = 0;
    cardRows.forEach(card => {
      if (!cardMatches_(card, pool.criteria)) return;
      outputRows.push([pool.poolId, ...outCols.map(i => card[i]), 1]);  // trailing 1 = Weight
      counts[pool.poolId]++;
    });
  });

  return { outputRows, counts };
}

function writeBoosterPool_(sheet, outputRows) {
  const width = POOL_OUTPUT_HEADERS.length + 2;  // + Pool ID + Weight
  const lastRow = sheet.getLastRow();
  if (lastRow >= FIRST_DATA_ROW) {
    sheet.getRange(FIRST_DATA_ROW, 1, lastRow - FIRST_DATA_ROW + 1, width).clearContent();
  }
  if (outputRows.length === 0) return;
  ensureRows_(sheet, outputRows.length + FIRST_DATA_ROW - 1);
  sheet.getRange(FIRST_DATA_ROW, 1, outputRows.length, width).setValues(outputRows);
}

function writePoolCounts_(sheet, pools, counts) {
  const idCol = pools.headers.findIndex(h => normalize_(h) === 'pool id');
  const countCol = pools.headers.findIndex(h => normalize_(h) === 'card count');
  if (countCol === -1 || pools.rows.length === 0) return;

  const values = pools.rows.map(row => {
    const id = String(row[idCol]).trim();
    return [id ? counts[id] : ''];
  });
  sheet.getRange(FIRST_DATA_ROW, countCol + 1, values.length, 1).setValues(values);
}

function ensureRows_(sheet, neededRows) {
  const existing = sheet.getMaxRows();
  if (neededRows > existing) sheet.insertRowsAfter(existing, neededRows - existing);
}

function buildPoolOptionLists() {
  const ss = SpreadsheetApp.getActive();
  const listSheet = ss.getSheetByName('POOL_LISTS') || ss.insertSheet('POOL_LISTS');
  const cards = readTable_(ss.getSheetByName('CARD_DATA'));
  const pools = readTable_(ss.getSheetByName('POOL_CONFIG'));
  const cardMap = makeHeaderMap_(cards.headers);

  // One list per criteria column, in POOL_CONFIG column order
  const columns = [];
  pools.headers.forEach(header => {
    const key = normalize_(header);
    if (!key || POOL_META_HEADERS.includes(key) || !(key in cardMap)) return;

    const col = cardMap[key];
    const isList = LIST_COLUMNS.includes(key);
    const found = new Set();
    cards.rows.forEach(row => {
      const cell = normalize_(row[col]);
      if (!cell) return;
      const parts = isList ? cell.split(',').map(t => t.trim()) : [cell];
      parts.forEach(v => found.add(v));
    });

    const sorted = [...found].sort();
    columns.push([header, '(none)', ...sorted, ...sorted.map(v => 'not ' + v)]);
  });

  // Flip column lists into rows, since setValues writes row by row
  const height = Math.max(...columns.map(c => c.length));
  const grid = [];
  for (let r = 0; r < height; r++) {
    grid.push(columns.map(c => c[r] ?? ''));
  }

  listSheet.clearContents();
  ensureRows_(listSheet, height);
  listSheet.getRange(1, 1, height, columns.length).setValues(grid);
}


/**
 * Replaces any existing refresh trigger with one that runs daily around 10 AM.
 */
function createDailyTrigger() {
  const fn = 'refreshSelectedSet';

  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === fn)
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger(fn).timeBased().everyDays(1).atHour(10).create();
}

function normalize_(value) {
  return String(value).trim().toLowerCase();
}

/**
 * Fetches a Scryfall URL, retrying with increasing delays when rate limited
 * (HTTP 429). Any other error stops the refresh with a readable message.
 */
function fetchScryfall_(url) {
  const retryDelays = [2000, 5000, 10000, 20000];

  for (let attempt = 0; attempt <= retryDelays.length; attempt++) {
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: SCRYFALL_HEADERS,
      muteHttpExceptions: true
    });
    const code = response.getResponseCode();

    if (code === 200) return response;

    if (code === 429 && attempt < retryDelays.length) {
      Utilities.sleep(retryDelays[attempt]);
      continue;
    }

    throw new Error('Scryfall request failed: HTTP ' + code + ' - ' +
      response.getContentText().slice(0, 300));
  }
}
