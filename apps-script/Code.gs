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
