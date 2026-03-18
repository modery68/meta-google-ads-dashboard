// =============================================================================
// GOOGLE ADS SYNC (PMax Enhanced) → Google Sheets
// =============================================================================
// Pulls 4 data sets into separate tabs:
//   1. google_ads_daily     — Campaign-level daily performance
//   2. google_ads_assets    — Asset group performance (PMax breakdown)
//   3. google_ads_products  — Shopping product performance (what's selling)
//   4. google_ads_search    — Search term insights (what queries trigger your ads)
//
// Setup:
//   1. Google Ads → Tools & Settings → Scripts → New Script
//   2. Paste this code
//   3. Set SPREADSHEET_URL below
//   4. Run once manually to authorize, then schedule daily
// =============================================================================

var CONFIG = {
  SPREADSHEET_URL: 'YOUR_GOOGLE_SHEET_URL_HERE',
  LOOKBACK_DAYS: 30,
};

function main() {
  var ss = SpreadsheetApp.openByUrl(CONFIG.SPREADSHEET_URL);
  var today = new Date();
  var lookback = new Date(today);
  lookback.setDate(today.getDate() - CONFIG.LOOKBACK_DAYS);
  var dateFrom = formatDate(lookback);
  var dateTo = formatDate(today);

  syncCampaignDaily(ss, dateFrom, dateTo);
  syncAssetGroups(ss, dateFrom, dateTo);
  syncProducts(ss, dateFrom, dateTo);
  syncSearchTerms(ss, dateFrom, dateTo);

  var logSheet = ss.getSheetByName('sync_log');
  if (logSheet) {
    logSheet.getRange(logSheet.getLastRow() + 1, 1, 1, 2).setValues([[
      new Date().toISOString(),
      'Google Ads sync complete (campaign + assets + products + search terms)'
    ]]);
  }
}

// =============================================================================
// 1. CAMPAIGN DAILY
// =============================================================================
function syncCampaignDaily(ss, dateFrom, dateTo) {
  var sheet = getOrCreateSheet(ss, 'google_ads_daily');

  var headers = [
    'date', 'campaign', 'campaign_type', 'spend', 'impressions', 'clicks',
    'conversions', 'conversion_value', 'cpc', 'cpm', 'ctr',
    'cost_per_conversion', 'conversion_rate', 'roas'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f3f3f3');

  var query = 'SELECT ' +
    'segments.date, campaign.name, campaign.advertising_channel_type, ' +
    'metrics.cost_micros, metrics.impressions, metrics.clicks, ' +
    'metrics.conversions, metrics.conversions_value, ' +
    'metrics.average_cpc, metrics.average_cpm, metrics.ctr, ' +
    'metrics.cost_per_conversion, metrics.conversions_from_interactions_rate ' +
    'FROM campaign ' +
    'WHERE segments.date BETWEEN "' + dateFrom + '" AND "' + dateTo + '" ' +
    'AND metrics.impressions > 0 ' +
    'ORDER BY segments.date DESC, metrics.cost_micros DESC';

  var rows = [];
  var iter = AdsApp.report(query).rows();
  while (iter.hasNext()) {
    var r = iter.next();
    var spend = r['metrics.cost_micros'] / 1e6;
    var convVal = parseFloat(r['metrics.conversions_value']) || 0;
    rows.push([
      r['segments.date'], r['campaign.name'], r['campaign.advertising_channel_type'],
      spend, parseInt(r['metrics.impressions']) || 0, parseInt(r['metrics.clicks']) || 0,
      parseFloat(r['metrics.conversions']) || 0, convVal,
      r['metrics.average_cpc'] / 1e6, r['metrics.average_cpm'] / 1e6,
      (parseFloat(r['metrics.ctr']) || 0) * 100,
      r['metrics.cost_per_conversion'] / 1e6,
      (parseFloat(r['metrics.conversions_from_interactions_rate']) || 0) * 100,
      spend > 0 ? convVal / spend : 0
    ]);
  }

  writeRows(sheet, headers, rows, [
    [4, '"$"#,##0.00'], [5, '#,##0'], [6, '#,##0'], [7, '#,##0.00'],
    [8, '"$"#,##0.00'], [9, '"$"#,##0.00'], [10, '"$"#,##0.00'],
    [11, '0.00"%"'], [12, '"$"#,##0.00'], [13, '0.00"%"'], [14, '0.00"x"']
  ]);
  Logger.log('Campaign daily: ' + rows.length + ' rows');
}

// =============================================================================
// 2. ASSET GROUP PERFORMANCE — the closest thing to "ad set" in PMax
// =============================================================================
function syncAssetGroups(ss, dateFrom, dateTo) {
  var sheet = getOrCreateSheet(ss, 'google_ads_assets');

  var headers = [
    'date', 'campaign', 'asset_group', 'status',
    'spend', 'impressions', 'clicks', 'conversions', 'conversion_value',
    'ctr', 'conversion_rate', 'cpc', 'roas'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f3f3f3');

  var query = 'SELECT ' +
    'segments.date, campaign.name, asset_group.name, asset_group.status, ' +
    'metrics.cost_micros, metrics.impressions, metrics.clicks, ' +
    'metrics.conversions, metrics.conversions_value, ' +
    'metrics.ctr, metrics.conversions_from_interactions_rate, metrics.average_cpc ' +
    'FROM asset_group ' +
    'WHERE segments.date BETWEEN "' + dateFrom + '" AND "' + dateTo + '" ' +
    'AND metrics.impressions > 0 ' +
    'ORDER BY segments.date DESC, metrics.cost_micros DESC';

  var rows = [];
  try {
    var iter = AdsApp.report(query).rows();
    while (iter.hasNext()) {
      var r = iter.next();
      var spend = r['metrics.cost_micros'] / 1e6;
      var convVal = parseFloat(r['metrics.conversions_value']) || 0;
      rows.push([
        r['segments.date'], r['campaign.name'], r['asset_group.name'], r['asset_group.status'],
        spend, parseInt(r['metrics.impressions']) || 0, parseInt(r['metrics.clicks']) || 0,
        parseFloat(r['metrics.conversions']) || 0, convVal,
        (parseFloat(r['metrics.ctr']) || 0) * 100,
        (parseFloat(r['metrics.conversions_from_interactions_rate']) || 0) * 100,
        r['metrics.average_cpc'] / 1e6,
        spend > 0 ? convVal / spend : 0
      ]);
    }
  } catch (e) {
    Logger.log('Asset group query failed (may not have PMax campaigns): ' + e.message);
  }

  writeRows(sheet, headers, rows, [
    [5, '"$"#,##0.00'], [6, '#,##0'], [7, '#,##0'], [8, '#,##0.00'],
    [9, '"$"#,##0.00'], [10, '0.00"%"'], [11, '0.00"%"'],
    [12, '"$"#,##0.00'], [13, '0.00"x"']
  ]);
  Logger.log('Asset groups: ' + rows.length + ' rows');
}

// =============================================================================
// 3. PRODUCT PERFORMANCE — which products are PMax pushing and converting
// =============================================================================
function syncProducts(ss, dateFrom, dateTo) {
  var sheet = getOrCreateSheet(ss, 'google_ads_products');

  var headers = [
    'date', 'campaign', 'product_title', 'product_type', 'product_id',
    'spend', 'impressions', 'clicks', 'conversions', 'conversion_value',
    'ctr', 'cpc', 'roas'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f3f3f3');

  var query = 'SELECT ' +
    'segments.date, campaign.name, ' +
    'segments.product_title, segments.product_type_l1, segments.product_item_id, ' +
    'metrics.cost_micros, metrics.impressions, metrics.clicks, ' +
    'metrics.conversions, metrics.conversions_value, ' +
    'metrics.ctr, metrics.average_cpc ' +
    'FROM shopping_performance_view ' +
    'WHERE segments.date BETWEEN "' + dateFrom + '" AND "' + dateTo + '" ' +
    'AND metrics.impressions > 0 ' +
    'ORDER BY metrics.cost_micros DESC';

  var rows = [];
  try {
    var iter = AdsApp.report(query).rows();
    while (iter.hasNext()) {
      var r = iter.next();
      var spend = r['metrics.cost_micros'] / 1e6;
      var convVal = parseFloat(r['metrics.conversions_value']) || 0;
      rows.push([
        r['segments.date'], r['campaign.name'],
        r['segments.product_title'] || '', r['segments.product_type_l1'] || '',
        r['segments.product_item_id'] || '',
        spend, parseInt(r['metrics.impressions']) || 0, parseInt(r['metrics.clicks']) || 0,
        parseFloat(r['metrics.conversions']) || 0, convVal,
        (parseFloat(r['metrics.ctr']) || 0) * 100,
        r['metrics.average_cpc'] / 1e6,
        spend > 0 ? convVal / spend : 0
      ]);
    }
  } catch (e) {
    Logger.log('Shopping performance query failed: ' + e.message);
  }

  writeRows(sheet, headers, rows, [
    [6, '"$"#,##0.00'], [7, '#,##0'], [8, '#,##0'], [9, '#,##0.00'],
    [10, '"$"#,##0.00'], [11, '0.00"%"'], [12, '"$"#,##0.00'], [13, '0.00"x"']
  ]);
  Logger.log('Products: ' + rows.length + ' rows');
}

// =============================================================================
// 4. SEARCH TERMS — standard Search/Shopping + PMax search insights combined
// =============================================================================
function syncSearchTerms(ss, dateFrom, dateTo) {
  var sheet = getOrCreateSheet(ss, 'google_ads_search');

  var headers = [
    'date', 'campaign', 'search_term', 'source',
    'spend', 'impressions', 'clicks', 'conversions', 'conversion_value',
    'ctr', 'cpc', 'roas'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f3f3f3');

  var rows = [];

  // ── Part A: Standard Search & Shopping campaigns ──────────────────────────
  // Note: PMax campaigns do NOT appear here — they use a separate API (Part B).
  var searchQuery = 'SELECT ' +
    'segments.date, campaign.name, search_term_view.search_term, ' +
    'metrics.cost_micros, metrics.impressions, metrics.clicks, ' +
    'metrics.conversions, metrics.conversions_value, ' +
    'metrics.ctr, metrics.average_cpc ' +
    'FROM search_term_view ' +
    'WHERE segments.date BETWEEN "' + dateFrom + '" AND "' + dateTo + '" ' +
    'AND metrics.impressions > 0 ' +
    'ORDER BY metrics.cost_micros DESC';

  try {
    var iter = AdsApp.report(searchQuery).rows();
    while (iter.hasNext()) {
      var r = iter.next();
      var spend = r['metrics.cost_micros'] / 1e6;
      var convVal = parseFloat(r['metrics.conversions_value']) || 0;
      rows.push([
        r['segments.date'], r['campaign.name'],
        r['search_term_view.search_term'] || '', 'Search/Shopping',
        spend, parseInt(r['metrics.impressions']) || 0, parseInt(r['metrics.clicks']) || 0,
        parseFloat(r['metrics.conversions']) || 0, convVal,
        (parseFloat(r['metrics.ctr']) || 0) * 100,
        r['metrics.average_cpc'] / 1e6,
        spend > 0 ? convVal / spend : 0
      ]);
    }
    Logger.log('Search/Shopping terms: ' + rows.length + ' rows');
  } catch (e) {
    Logger.log('Search term query failed: ' + e.message);
  }

  // ── Part B: PMax search term insights ─────────────────────────────────────
  // PMax doesn't use search_term_view. Use pmax_search_term_insight instead.
  // Note: Google only reveals terms that meet their privacy threshold (~100+ impressions).
  var pmaxQuery = 'SELECT ' +
    'campaign.name, ' +
    'pmax_search_term_insight.search_term_insight_category, ' +  // category/theme
    'pmax_search_term_insight.search_term, ' +
    'metrics.impressions ' +
    'FROM pmax_search_term_insight ' +
    'WHERE segments.date BETWEEN "' + dateFrom + '" AND "' + dateTo + '" ' +
    'ORDER BY metrics.impressions DESC';

  try {
    var pmaxIter = AdsApp.report(pmaxQuery).rows();
    var pmaxCount = 0;
    while (pmaxIter.hasNext()) {
      var pr = pmaxIter.next();
      var term = pr['pmax_search_term_insight.search_term'] || '';
      var category = pr['pmax_search_term_insight.search_term_insight_category'] || '';
      var label = term ? term : ('[Category: ' + category + ']');
      rows.push([
        dateTo, pr['campaign.name'],
        label, 'PMax',
        '', // PMax doesn't expose spend per search term
        parseInt(pr['metrics.impressions']) || 0,
        '', '', '', '', '', ''
      ]);
      pmaxCount++;
    }
    Logger.log('PMax search insights: ' + pmaxCount + ' rows');
  } catch (e) {
    Logger.log('PMax search insight query failed (may need account-level access): ' + e.message);
  }

  writeRows(sheet, headers, rows, [
    [5, '"$"#,##0.00'], [6, '#,##0'], [7, '#,##0'], [8, '#,##0.00'],
    [9, '"$"#,##0.00'], [10, '0.00"%"'], [11, '"$"#,##0.00'], [12, '0.00"x"']
  ]);
  Logger.log('Search terms total: ' + rows.length + ' rows');
}

// =============================================================================
// HELPERS
// =============================================================================
function getOrCreateSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  sheet.clear();
  return sheet;
}

function formatDate(date) {
  return date.getFullYear() + '-' +
    ('0' + (date.getMonth() + 1)).slice(-2) + '-' +
    ('0' + date.getDate()).slice(-2);
}

function writeRows(sheet, headers, rows, formatMap) {
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    formatMap.forEach(function(pair) {
      sheet.getRange(2, pair[0], rows.length, 1).setNumberFormat(pair[1]);
    });
  }
}
