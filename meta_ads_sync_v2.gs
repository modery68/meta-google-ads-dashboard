// =============================================================================
// META ADS → GOOGLE SHEETS AUTO-SYNC + DECISION ENGINE
// =============================================================================
// v2 — Now with: Dashboard, Action Items, Creative Scoring, Fatigue Detection
//
// Setup: 
//   1. Create a new Google Sheet
//   2. Extensions → Apps Script → paste this entire script
//   3. Fill in CONFIG below with your Meta credentials
//   4. Run setupTriggers() once to schedule daily auto-refresh
// =============================================================================

const CONFIG = {
  // ── Meta API Credentials ──────────────────────────────────────────────
  ACCESS_TOKEN: 'YOUR_ACCESS_TOKEN_HERE',
  AD_ACCOUNT_ID: 'act_YOUR_ACCOUNT_ID',

  // ── Data Settings ─────────────────────────────────────────────────────
  LOOKBACK_DAYS: 30,
  API_VERSION: 'v25.0',
  SCREENSHOT_MODE: false,
  BRAND_KEYWORDS: ['yourbrand', 'your brand'],  // Add your brand name variations for PMax brand vs non-brand split

  // ── Decision Thresholds ───────────────────────────────────────────────
  // These drive every recommendation. Tune once, decisions auto-update.
  THRESHOLDS: {
    TARGET_ROAS: 1.5,               // Your ROAS floor — below this = underperforming
    KILL_ROAS: 0.8,                 // Below this with enough spend = kill it
    SCALE_ROAS: 2.0,               // Above this = candidate for more budget
    MIN_SPEND_FOR_JUDGMENT: 50,     // Don't judge creatives under this spend ($)
    TESTING_SPEND_CAP: 150,         // Under this = still in "testing" phase
    FREQUENCY_WARNING: 2.5,         // Above this = audience seeing it too much
    FREQUENCY_CRITICAL: 4.0,        // Above this = definitely fatiguing
    CTR_FLOOR: 0.8,                 // Below this CTR% = weak creative signal
    THUMBSTOP_FLOOR: 20,            // Below this thumbstop% = not stopping scroll
    HOLDRATE_FLOOR: 20,             // Below this hold% = losing people mid-video
    TOP_N_CREATIVES: 10,            // How many to show in dashboard "Top" lists
    BUDGET_SHIFT_THRESHOLD: 0.3,    // Flag campaigns spending >30% with ROAS < target
  },

  // ── AI Insight (optional) ───────────────────────────────────────────
  // Set API key to enable AI-powered analysis. Leave blank to skip.
  // Supports: 'claude' (Anthropic) or 'gemini' (Google)
  AI: {
    PROVIDER: 'claude',                     // 'claude' or 'gemini'
    API_KEY: '',                            // Your API key (Anthropic or Google AI Studio)
    MODEL: 'claude-sonnet-4-20250514',     // claude: 'claude-sonnet-4-20250514' | gemini: 'gemini-2.0-flash'
  },

  // ── Sheet Tab Names ───────────────────────────────────────────────────
  TABS: {
    DASHBOARD:        '📊 dashboard',
    META_DEEP:        '📈 meta_deep_dives',
    GOOGLE_DEEP:      '📈 google_deep_dives',
    META_CREATIVE:    'meta_creative',
    META_AGE:         'meta_age_gender',
    META_DAILY:       'meta_daily',
    META_ADSET:       'meta_adset_daily',
    GOOGLE_ADS:       'google_ads_daily',
    LOG:              'sync_log'
  }
};

// Data cache — lets the diagnostic engine access parsed data
const DATA_CACHE = {
  ageGender: [],
  creatives: []
};

// =============================================================================
// MAIN SYNC FUNCTIONS
// =============================================================================

function syncAll() {
  const log = [];
  const start = new Date();
  log.push(`Sync started: ${start.toISOString()}`);

  // Clean up legacy tabs from older versions
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const legacyTabs = ['⚡ actions'];
  legacyTabs.forEach(name => {
    const old = ss.getSheetByName(name);
    if (old) {
      ss.deleteSheet(old);
      log.push(`Removed legacy tab: ${name}`);
    }
  });

  try {
    // Phase 1: Pull raw data (same as before)
    syncMetaDaily(log);
    syncMetaAdsetDaily(log);
    syncMetaCreative(log);
    syncMetaAgeGender(log);

    // Phase 2: Dashboard with diagnostics
    buildDashboard(log);

    const end = new Date();
    const duration = ((end - start) / 1000).toFixed(1);
    log.push(`Sync completed in ${duration}s`);
  } catch (e) {
    log.push(`ERROR: ${e.message}`);
  }

  writeLog(log);
}

// =============================================================================
// 📊 DASHBOARD
// =============================================================================

function buildDashboard(log) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const T = CONFIG.THRESHOLDS;

  let sheet = ss.getSheetByName(CONFIG.TABS.DASHBOARD);
  if (!sheet) sheet = ss.insertSheet(CONFIG.TABS.DASHBOARD);
  sheet.clear();
  sheet.clearConditionalFormatRules();

  // Remove old charts
  const existingCharts = sheet.getCharts();
  existingCharts.forEach(c => sheet.removeChart(c));

  // ── Pull ALL daily data with full metrics ─────────────────────────────
  const dailySheet = ss.getSheetByName(CONFIG.TABS.META_DAILY);
  let dailyRows = [];  // [{date, spend, revenue, purchases, impressions, clicks, atc, reach}, ...]

  if (dailySheet && dailySheet.getLastRow() > 1) {
    const data = dailySheet.getRange(2, 1, dailySheet.getLastRow() - 1, dailySheet.getLastColumn()).getValues();
    const headers = dailySheet.getRange(1, 1, 1, dailySheet.getLastColumn()).getValues()[0];
    const di = {};
    headers.forEach((h, i) => di[h] = i);

    data.forEach(row => {
      let date = row[di['date']];
      if (date instanceof Date) {
        date = Utilities.formatDate(date, 'UTC', 'yyyy-MM-dd');
      } else {
        date = String(date).substring(0, 10);
      }

      dailyRows.push({
        date: date,
        campaign: row[di['campaign']] || '',
        spend: parseFloat(row[di['spend']]) || 0,
        revenue: parseFloat(row[di['purchase_value']]) || 0,
        purchases: parseFloat(row[di['purchases']]) || 0,
        impressions: parseInt(row[di['impressions']]) || 0,
        clicks: parseInt(row[di['clicks']]) || 0,
        atc: parseInt(row[di['atc']]) || 0,
        reach: parseInt(row[di['reach']]) || 0,
      });
    });
  }

  // ── Aggregate helper ──────────────────────────────────────────────────
  function aggregate(rows) {
    const a = { spend: 0, revenue: 0, purchases: 0, impressions: 0, clicks: 0, atc: 0, reach: 0 };
    rows.forEach(r => {
      a.spend += r.spend;
      a.revenue += r.revenue;
      a.purchases += r.purchases;
      a.impressions += r.impressions;
      a.clicks += r.clicks;
      a.atc += r.atc;
      a.reach += r.reach;
    });
    a.roas = a.spend > 0 ? a.revenue / a.spend : 0;
    a.cpa = a.purchases > 0 ? a.spend / a.purchases : 0;
    a.cpc = a.clicks > 0 ? a.spend / a.clicks : 0;
    a.cpm = a.impressions > 0 ? (a.spend / a.impressions) * 1000 : 0;
    a.ctr = a.impressions > 0 ? (a.clicks / a.impressions) * 100 : 0;
    a.cvr = a.clicks > 0 ? (a.purchases / a.clicks) * 100 : 0;
    a.aov = a.purchases > 0 ? a.revenue / a.purchases : 0;
    return a;
  }

  // ── Build time windows ────────────────────────────────────────────────
  // Aggregate by date first (campaigns collapse into single day)
  const dailyTrend = {};
  dailyRows.forEach(r => {
    if (!dailyTrend[r.date]) dailyTrend[r.date] = [];
    dailyTrend[r.date].push(r);
  });

  const sortedDates = Object.keys(dailyTrend).sort();

  // Exclude today (incomplete day) — Meta's "last 7 days" = last 7 COMPLETED days
  const todayStr = Utilities.formatDate(new Date(), 'UTC', 'yyyy-MM-dd');
  const completedDates = sortedDates.filter(d => d < todayStr);

  const allAgg = aggregate(dailyRows);

  const last7Dates = completedDates.slice(-7);
  const prior7Dates = completedDates.slice(-14, -7);
  const last7Rows = dailyRows.filter(r => last7Dates.includes(r.date));
  const prior7Rows = dailyRows.filter(r => prior7Dates.includes(r.date));
  const last7 = aggregate(last7Rows);
  const prior7 = aggregate(prior7Rows);

  // Actual date range labels (no year — we know what year it is)
  function shortDate(d) {
    const parts = d.split('-');
    return parts[1] + '/' + parts[2];
  }
  const dateRange = getDateRange();
  const fullRangeLabel = `${shortDate(dateRange.since)} → ${shortDate(dateRange.until)}`;
  const last7Label = last7Dates.length > 0 ? `${shortDate(last7Dates[0])} → ${shortDate(last7Dates[last7Dates.length - 1])}` : 'N/A';
  const prior7Label = prior7Dates.length > 0 ? `${shortDate(prior7Dates[0])} → ${shortDate(prior7Dates[prior7Dates.length - 1])}` : 'N/A';

  // ── FREQUENCY — Pull account-level data (can't compute from campaign×day) ─
  // Frequency = impressions / unique_reach, but reach overlaps across campaigns
  // and days. Only account-level time_increment=all_days gives the real number.
  function fetchAccountFrequency(sinceDate, untilDate) {
    const url = `https://graph.facebook.com/${CONFIG.API_VERSION}/${CONFIG.AD_ACCOUNT_ID}/insights` +
      `?access_token=${CONFIG.ACCESS_TOKEN}` +
      `&time_range=${encodeURIComponent(`{"since":"${sinceDate}","until":"${untilDate}"}`)}` +
      `&level=account&time_increment=all_days&fields=frequency,impressions,reach&limit=1`;
    try {
      const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      if (resp.getResponseCode() === 200) {
        const json = JSON.parse(resp.getContentText());
        if (json.data && json.data.length > 0) {
          return {
            frequency: parseFloat(json.data[0].frequency) || 0,
            reach: parseInt(json.data[0].reach) || 0,
            impressions: parseInt(json.data[0].impressions) || 0,
          };
        }
      }
    } catch (e) { /* fail silently, will use 0 */ }
    return { frequency: 0, reach: 0, impressions: 0 };
  }

  const last7Freq = last7Dates.length > 0
    ? fetchAccountFrequency(last7Dates[0], last7Dates[last7Dates.length - 1])
    : { frequency: 0, reach: 0, impressions: 0 };
  const prior7Freq = prior7Dates.length > 0
    ? fetchAccountFrequency(prior7Dates[0], prior7Dates[prior7Dates.length - 1])
    : { frequency: 0, reach: 0, impressions: 0 };

  // Override aggregate reach with deduplicated account-level reach
  last7.reach = last7Freq.reach || last7.reach;
  prior7.reach = prior7Freq.reach || prior7.reach;
  const avgFreq7 = last7Freq.frequency;
  const avgFreqP7 = prior7Freq.frequency;

  // Trend calculation helper
  const WOW_NOISE_FLOOR = 5;

  function trendPct(current, previous, invertGood) {
    if (!previous || previous === 0) return '';
    const pct = ((current - previous) / previous) * 100;
    const absPct = Math.abs(pct);
    let arrow;
    if (invertGood) {
      arrow = pct <= 0 ? '▲' : '▼';
    } else {
      arrow = pct >= 0 ? '▲' : '▼';
    }
    return `${arrow} ${absPct.toFixed(1)}%`;
  }

  function trendRawPct(current, previous) {
    if (!previous || previous === 0) return 0;
    return ((current - previous) / previous) * 100;
  }

  // ── READ ADSET DATA FOR PER-ADSET FREQUENCY IN DIAGNOSTICS ────────────
  const adsetSheet = ss.getSheetByName(CONFIG.TABS.META_ADSET);
  let adsetRows = [];
  if (adsetSheet && adsetSheet.getLastRow() > 1) {
    const adata = adsetSheet.getRange(2, 1, adsetSheet.getLastRow() - 1, adsetSheet.getLastColumn()).getValues();
    const aheaders = adsetSheet.getRange(1, 1, 1, adsetSheet.getLastColumn()).getValues()[0];
    const adi = {};
    aheaders.forEach((h, i) => adi[h] = i);
    adata.forEach(row => {
      let date = row[adi['date']];
      if (date instanceof Date) date = Utilities.formatDate(date, 'UTC', 'yyyy-MM-dd');
      else date = String(date).substring(0, 10);
      adsetRows.push({
        date, adset: row[adi['adset']], campaign: row[adi['campaign']],
        frequency: parseFloat(row[adi['frequency']]) || 0,
        spend: parseFloat(row[adi['spend']]) || 0,
        impressions: parseInt(row[adi['impressions']]) || 0,
        reach: parseInt(row[adi['reach']]) || 0,
      });
    });
  }

  // Frequency is already computed from account-level API calls above (avgFreq7, avgFreqP7)
  // Adset-level frequency from daily rows is used only for per-adset diagnostics below
  const freq7 = adsetRows.filter(r => last7Dates.includes(r.date));

  // ══════════════════════════════════════════════════════════════════════
  // DIAGNOSTIC ENGINE — trace WHY performance shifted
  // ══════════════════════════════════════════════════════════════════════
  // ROAS = (CVR / CPC) × AOV
  // Decompose the drop: is it creative (CTR), auction (CPM), funnel (CVR), or all?

  // ══════════════════════════════════════════════════════════════════════
  // READ GOOGLE ADS DATA (if google_ads_daily tab exists)
  // ══════════════════════════════════════════════════════════════════════
  let hasGoogleAds = false;
  let gAdsL7 = { spend: 0, revenue: 0, purchases: 0, impressions: 0, clicks: 0, roas: 0, cpa: 0, cpc: 0, cpm: 0, ctr: 0, cvr: 0 };
  let gAdsP7 = { spend: 0, revenue: 0, purchases: 0, impressions: 0, clicks: 0, roas: 0, cpa: 0, cpc: 0, cpm: 0, ctr: 0, cvr: 0 };
  let gAdsAll = { spend: 0, revenue: 0, purchases: 0, impressions: 0, clicks: 0, roas: 0, cpa: 0, cpc: 0, cpm: 0, ctr: 0, cvr: 0 };

  const gAdsSheet = ss.getSheetByName(CONFIG.TABS.GOOGLE_ADS);
  if (gAdsSheet && gAdsSheet.getLastRow() > 1) {
    hasGoogleAds = true;
    const gdata = gAdsSheet.getRange(2, 1, gAdsSheet.getLastRow() - 1, gAdsSheet.getLastColumn()).getValues();
    const gheaders = gAdsSheet.getRange(1, 1, 1, gAdsSheet.getLastColumn()).getValues()[0];
    const gi = {};
    gheaders.forEach((h, idx) => gi[h] = idx);

    const gRows = [];
    gdata.forEach(row => {
      let date = row[gi['date']];
      if (date instanceof Date) date = Utilities.formatDate(date, 'UTC', 'yyyy-MM-dd');
      else date = String(date).substring(0, 10);
      gRows.push({
        date,
        spend: parseFloat(row[gi['spend']]) || 0,
        revenue: parseFloat(row[gi['conversion_value']]) || 0,
        purchases: parseFloat(row[gi['conversions']]) || 0,
        impressions: parseInt(row[gi['impressions']]) || 0,
        clicks: parseInt(row[gi['clicks']]) || 0,
      });
    });

    function gAggregate(rows) {
      const a = { spend: 0, revenue: 0, purchases: 0, impressions: 0, clicks: 0 };
      rows.forEach(r2 => { a.spend += r2.spend; a.revenue += r2.revenue; a.purchases += r2.purchases; a.impressions += r2.impressions; a.clicks += r2.clicks; });
      a.roas = a.spend > 0 ? a.revenue / a.spend : 0;
      a.cpa = a.purchases > 0 ? a.spend / a.purchases : 0;
      a.cpc = a.clicks > 0 ? a.spend / a.clicks : 0;
      a.cpm = a.impressions > 0 ? (a.spend / a.impressions) * 1000 : 0;
      a.ctr = a.impressions > 0 ? (a.clicks / a.impressions) * 100 : 0;
      a.cvr = a.clicks > 0 ? (a.purchases / a.clicks) * 100 : 0;
      return a;
    }

    const gL7Rows = gRows.filter(r2 => last7Dates.includes(r2.date));
    const gP7Rows = gRows.filter(r2 => prior7Dates.includes(r2.date));
    gAdsL7 = gAggregate(gL7Rows);
    gAdsP7 = gAggregate(gP7Rows);
    gAdsAll = gAggregate(gRows);
  }

  // Combined totals (Meta + Google Ads)
  const combinedL7 = {
    spend: last7.spend + gAdsL7.spend,
    revenue: last7.revenue + gAdsL7.revenue,
    purchases: last7.purchases + gAdsL7.purchases,
  };
  combinedL7.roas = combinedL7.spend > 0 ? combinedL7.revenue / combinedL7.spend : 0;
  combinedL7.cpa = combinedL7.purchases > 0 ? combinedL7.spend / combinedL7.purchases : 0;

  const combinedP7 = {
    spend: prior7.spend + gAdsP7.spend,
    revenue: prior7.revenue + gAdsP7.revenue,
    purchases: prior7.purchases + gAdsP7.purchases,
  };
  combinedP7.roas = combinedP7.spend > 0 ? combinedP7.revenue / combinedP7.spend : 0;
  combinedP7.cpa = combinedP7.purchases > 0 ? combinedP7.spend / combinedP7.purchases : 0;

  const findings = []; // { signal, severity, evidence, action }

  // ── Detect creative format: static vs video ─────────────────────────
  // If any active creative has thumbstop or holdRate > 0, we have video ads
  let hasVideo = false;
  if (DATA_CACHE.creatives && DATA_CACHE.creatives.length > 0) {
    hasVideo = DATA_CACHE.creatives.some(c =>
      c.recent.spend > 0 && (c.recent.thumbstop > 0 || c.recent.holdRate > 0)
    );
  }
  // Format-aware language
  const creativeNoun = hasVideo ? 'hooks/creatives' : 'ad images and headlines';
  const fatigueAction = hasVideo
    ? 'Rotate in fresh creatives — new hooks, new opening frames. Check the meta_creative tab for ▼ trends.'
    : 'Rotate in fresh creatives — new images, new headlines, new angles. Check the meta_creative tab for ▼ trends.';
  const ctrAction = hasVideo
    ? 'Test new opening hooks on your top-spend creatives. The body/offer still converts (CVR is holding), people just aren\'t clicking.'
    : 'Test new headlines, primary text, and image variations. CVR is holding so the offer works — the ad just isn\'t grabbing attention in the feed anymore.';
  const batchAction = hasVideo
    ? 'You need a fresh creative batch. Current angles are exhausted. Test new hooks, new formats, or new offers.'
    : 'You need a fresh creative batch. Current angles are exhausted. Test new images, new copy angles, different social proof, or new offers.';
  const declineHint = hasVideo ? 'test new hook or kill' : 'test new image/headline or kill';

  const pctChange = (curr, prev) => prev > 0 ? ((curr - prev) / prev * 100) : 0;

  const roasDelta = pctChange(last7.roas, prior7.roas);
  const ctrDelta = pctChange(last7.ctr, prior7.ctr);
  const cpmDelta = pctChange(last7.cpm, prior7.cpm);
  const cvrDelta = pctChange(last7.cvr, prior7.cvr);
  const cpaDelta = pctChange(last7.cpa, prior7.cpa);
  const cpcDelta = pctChange(last7.cpc, prior7.cpc);
  const spendDelta = pctChange(last7.spend, prior7.spend);
  const reachDelta = pctChange(last7.reach, prior7.reach);
  const freqDelta = pctChange(avgFreq7, avgFreqP7);
  const aovDelta = pctChange(last7.aov, prior7.aov);
  const atcRate7 = last7.clicks > 0 ? (last7.atc / last7.clicks * 100) : 0;
  const atcRateP7 = prior7.clicks > 0 ? (prior7.atc / prior7.clicks * 100) : 0;
  const atcToPurch7 = last7.atc > 0 ? (last7.purchases / last7.atc * 100) : 0;
  const atcToPurchP7 = prior7.atc > 0 ? (prior7.purchases / prior7.atc * 100) : 0;
  const atcToPurchDelta = pctChange(atcToPurch7, atcToPurchP7);
  const atcRateDelta = pctChange(atcRate7, atcRateP7);
  const purchDeltaEarly = pctChange(last7.purchases, prior7.purchases);

  // Only run diagnostics if there's a meaningful ROAS shift
  const roasDropped = roasDelta < -WOW_NOISE_FLOOR;
  const roasImproved = roasDelta > WOW_NOISE_FLOOR;

  // Quality-vs-quantity signals (need outer scope for trend narrative)
  let qualityTradeOff = false;
  let cvrImproved = cvrDelta > WOW_NOISE_FLOOR;
  let cpcUp = cpcDelta > 10;
  let ctrDown = ctrDelta < -5;

  if (roasDropped) {
    const cpaStable = Math.abs(cpaDelta) < WOW_NOISE_FLOOR;
    const cpaWorse = cpaDelta > WOW_NOISE_FLOOR;
    qualityTradeOff = cpcUp && cvrImproved && (cpaStable || (cpaDelta < cpcDelta * 0.5));

    // ── QUALITY vs QUANTITY trade-off (CPC up + CVR up) ─────────────────
    if (qualityTradeOff) {
      findings.push({
        signal: '🟢 Quality Trade-off',
        severity: 'none',
        evidence: `CPC up ${cpcDelta.toFixed(1)}% but CVR also up ${cvrDelta.toFixed(1)}%. ` +
          `CPA only moved ${cpaDelta.toFixed(1)}% (vs ${cpcDelta.toFixed(1)}% CPC increase). Meta is paying more per click but getting higher-intent users.`,
        action: 'Healthy optimization. More expensive clicks that convert better. No action needed.'
      });
    }

    // ── CTR down but CVR up (fewer clicks, better quality) ──────────────
    const audienceRefinement = ctrDown && cvrImproved && !qualityTradeOff && (cpaStable || (cpaDelta < Math.abs(ctrDelta) * 0.5));
    if (audienceRefinement) {
      findings.push({
        signal: '🟢 Audience Refinement',
        severity: 'none',
        evidence: `CTR dropped ${Math.abs(ctrDelta).toFixed(1)}% but CVR improved ${cvrDelta.toFixed(1)}%. ` +
          `Fewer people click, but those who do are more likely to buy. CPA is flat.`,
        action: 'Meta is narrowing delivery to high-intent users. This is fine as long as CPA holds and volume is sufficient.'
      });
    }

    // ── CHECK 1: Creative Fatigue (CTR down + frequency up + CVR NOT compensating) ──
    if (ctrDelta < -5 && freqDelta > 5 && !cvrImproved) {
      findings.push({
        signal: '🔴 Creative Fatigue',
        severity: 'high',
        evidence: `CTR dropped ${Math.abs(ctrDelta).toFixed(1)}% while frequency rose ${freqDelta.toFixed(1)}%. ` +
          `Avg frequency last 7d: ${avgFreq7.toFixed(1)} (was ${avgFreqP7.toFixed(1)}). CVR didn't compensate.`,
        action: fatigueAction
      });
    } else if (ctrDelta < -8 && !cvrImproved) {
      findings.push({
        signal: '🟡 CTR Declining',
        severity: 'medium',
        evidence: `CTR dropped ${Math.abs(ctrDelta).toFixed(1)}% WoW (${prior7.ctr.toFixed(2)}% → ${last7.ctr.toFixed(2)}%). ` +
          `Frequency is ${avgFreq7.toFixed(1)}. CVR didn't improve to offset, so CPA is getting worse.`,
        action: ctrAction
      });
    }

    // ── CHECK 2: Audience Saturation (reach down, CPM up) ───────────────
    if (reachDelta < -10 && cpmDelta > 5) {
      findings.push({
        signal: '🔴 Audience Saturation',
        severity: 'high',
        evidence: `Reach dropped ${Math.abs(reachDelta).toFixed(1)}% while CPM rose ${cpmDelta.toFixed(1)}%. ` +
          `You're paying more to reach fewer new people.`,
        action: 'Expand targeting. Test broader audiences, lookalikes from different seed events (ATC, ViewContent), or new interest stacks.'
      });
    }

    // ── CHECK 3: CPM Inflation (auction pressure, not your ads) ─────────
    if (cpmDelta > 10 && Math.abs(ctrDelta) < 5 && Math.abs(cvrDelta) < 5) {
      findings.push({
        signal: '🟡 Auction Pressure (CPM Spike)',
        severity: 'medium',
        evidence: `CPM jumped ${cpmDelta.toFixed(1)}% ($${prior7.cpm.toFixed(2)} → $${last7.cpm.toFixed(2)}) ` +
          `but CTR and CVR are stable. The creative isn't the problem, the auction is more expensive.`,
        action: 'May be seasonal or competitive. If persistent, improve CTR to offset higher CPMs, or shift budget to lower-CPM placements.'
      });
    }

    // ── CHECK 4: Funnel Breakdown (CVR dropping) ────────────────────────
    if (cvrDelta < -10 && Math.abs(ctrDelta) < 5) {
      findings.push({
        signal: '🔴 Conversion Rate Drop',
        severity: 'high',
        evidence: `CVR dropped ${Math.abs(cvrDelta).toFixed(1)}% (${prior7.cvr.toFixed(2)}% → ${last7.cvr.toFixed(2)}%) ` +
          `while CTR is stable. People click but don't buy.`,
        action: 'Check landing page, cart flow, and offer. Price change? OOS? Slow load times? ' +
          'Also check if new creatives are attracting a less qualified audience.'
      });
    } else if (cvrDelta < -5 && !ctrDown) {
      // Only flag CVR softening if CTR isn't also down (avoid double-counting)
      findings.push({
        signal: '🟡 CVR Softening',
        severity: 'medium',
        evidence: `CVR down ${Math.abs(cvrDelta).toFixed(1)}% (${prior7.cvr.toFixed(2)}% → ${last7.cvr.toFixed(2)}%). ` +
          `Contributing to the ROAS decline.`,
        action: 'Monitor for another week. If it continues, audit the landing page and check ad-to-LP message match.'
      });
    }

    // ── CHECK 4b: AOV Drop ──────────────────────────────────────────────
    if (aovDelta < -8) {
      findings.push({
        signal: '🟡 AOV Declining',
        severity: 'medium',
        evidence: `AOV dropped ${Math.abs(aovDelta).toFixed(1)}% ($${prior7.aov.toFixed(2)} → $${last7.aov.toFixed(2)}). ` +
          `Same conversion rate but less revenue per purchase drags ROAS.`,
        action: 'Check discount codes, bundling changes, or product mix shift. Are creatives driving traffic to lower-priced items?'
      });
    }

    // ── CHECK 4c: Checkout Drop-off ─────────────────────────────────────
    if (atcToPurchDelta < -10 && last7.atc > 50) {
      findings.push({
        signal: '🔴 Checkout Drop-off',
        severity: 'high',
        evidence: `ATC→Purchase rate dropped ${Math.abs(atcToPurchDelta).toFixed(1)}% ` +
          `(${atcToPurchP7.toFixed(1)}% → ${atcToPurch7.toFixed(1)}%). People add to cart but don't complete.`,
        action: 'NOT an ad problem. Check: cart page changes, shipping cost surprises, payment issues, OOS, or broken checkout step.'
      });
    }

    // ── CHECK 5: CPC Spike (only if NOT explained by quality trade-off) ─
    if (cpcDelta > 15 && !cvrImproved) {
      const alreadyExplained = findings.some(f =>
        f.signal.includes('Creative Fatigue') || f.signal.includes('CPM Spike') || f.signal.includes('Quality')
      );
      if (!alreadyExplained) {
        findings.push({
          signal: '🟡 CPC Spike',
          severity: 'medium',
          evidence: `CPC jumped ${cpcDelta.toFixed(1)}% ($${prior7.cpc.toFixed(2)} → $${last7.cpc.toFixed(2)}). ` +
            `CVR didn't improve to offset, so this directly hurts CPA.`,
          action: 'CPC = CPM / (CTR x 10). Fix whichever input is off: lower CPM (expand audience) or raise CTR (better creatives).'
        });
      }
    }

    // ── CHECK 6: Creative Concentration Risk ────────────────────────────
    if (DATA_CACHE.creatives && DATA_CACHE.creatives.length > 0) {
      const activeCreatives = DATA_CACHE.creatives.filter(c => c.phase === 'active' && c.recent.spend > 0);
      const totalCreativeSpend = activeCreatives.reduce((s, c) => s + c.recent.spend, 0);
      const topCreative = activeCreatives.sort((a, b) => b.recent.spend - a.recent.spend)[0];

      if (topCreative && totalCreativeSpend > 0) {
        const topShare = topCreative.recent.spend / totalCreativeSpend;
        if (topShare > 0.40) {
          findings.push({
            signal: '⚠️ Creative Concentration',
            severity: 'medium',
            evidence: `"${topCreative.name}" is eating ${(topShare * 100).toFixed(0)}% of active creative spend. ` +
              `If this one fatigues, account ROAS tanks.`,
            action: 'Diversify — you need 3-4 creatives carrying meaningful spend. Test more variations of your winning angles.'
          });
        }
      }

      // Count declining creatives
      const declining = DATA_CACHE.creatives.filter(c =>
        c.phase === 'active' && c.roasTrend === '▼' && c.recent.spend > T.MIN_SPEND_FOR_JUDGMENT
      );
      if (declining.length >= 3) {
        findings.push({
          signal: '🔴 Multiple Creatives Declining',
          severity: 'high',
          evidence: `${declining.length} active creatives showing ▼ ROAS trend (7d vs all-time). ` +
            `This isn't one bad ad — it's systemic fatigue across the account.`,
          action: batchAction
        });
      }
    }

    // ── CHECK 7: High-frequency ad sets ───────────────────────────────
    const highFreqAdsets = [];
    const adsetFreqMap = {};
    freq7.forEach(r => {
      if (!adsetFreqMap[r.adset]) adsetFreqMap[r.adset] = { imprSum: 0, reachSum: 0, spend: 0 };
      adsetFreqMap[r.adset].imprSum += r.impressions || 0;
      adsetFreqMap[r.adset].reachSum += r.reach || 0;
      adsetFreqMap[r.adset].spend += r.spend;
    });
    Object.entries(adsetFreqMap).forEach(([name, d]) => {
      // impr/reach at adset level across days is a reasonable proxy
      const freq = d.reachSum > 0 ? d.imprSum / d.reachSum : 0;
      if (freq >= T.FREQUENCY_CRITICAL && d.spend > 100) {
        highFreqAdsets.push({ name, freq: freq, spend: d.spend });
      }
    });
    if (highFreqAdsets.length > 0) {
      highFreqAdsets.sort((a, b) => b.freq - a.freq);
      const top3 = highFreqAdsets.slice(0, 3).map(a => `${a.name} (${a.freq.toFixed(1)}x)`).join(', ');
      findings.push({
        signal: '🔴 High Frequency Ad Sets',
        severity: 'high',
        evidence: `${highFreqAdsets.length} ad set(s) above ${T.FREQUENCY_CRITICAL}x frequency: ${top3}`,
        action: 'Swap creatives in these ad sets or expand the audience size. Frequency above 4x means most of your audience has seen the ad multiple times.'
      });
    }

    // ── CHECK 8: Scaling Pressure (spend up, ROAS dipped) ─────────────
    if (spendDelta > 12 && Math.abs(roasDelta) < 15) {
      const alreadyHasScaling = findings.some(f => f.signal.includes('Scaling') || f.signal.includes('Quality'));
      if (!alreadyHasScaling) {
        findings.push({
          signal: '🟢 Scaling Pressure',
          severity: 'none',
          evidence: `Spend increased ${spendDelta.toFixed(1)}% WoW while ROAS dipped ${Math.abs(roasDelta).toFixed(1)}%. ` +
            `Higher spend naturally reduces efficiency as Meta reaches further into the audience.`,
          action: 'Expected behavior when scaling. Monitor CPA, not just ROAS. If CPA is within target, the scaling is working.'
        });
      }
    }

    // ── CHECK 9: New Creative Dilution ────────────────────────────────
    if (DATA_CACHE.creatives && DATA_CACHE.creatives.length > 0) {
      const testingCreatives = DATA_CACHE.creatives.filter(c => c.phase === 'testing');
      const testingSpend = testingCreatives.reduce((s, c) => s + c.recent.spend, 0);
      const totalSpend7 = last7.spend;
      const testPct = totalSpend7 > 0 ? (testingSpend / totalSpend7 * 100) : 0;
      if (testPct > 15 && testingCreatives.length >= 3) {
        const avgTestRoas = testingCreatives.length > 0
          ? testingCreatives.reduce((s, c) => s + c.recent.roas, 0) / testingCreatives.length : 0;
        findings.push({
          signal: '🟡 New Creative Dilution',
          severity: 'medium',
          evidence: `${testingCreatives.length} creatives in testing are eating ${testPct.toFixed(0)}% of spend. ` +
            `Avg test ROAS: ${avgTestRoas.toFixed(2)}x vs account ${last7.roas.toFixed(2)}x. This drags blended numbers.`,
          action: 'Expected cost of testing. Kill underperformers faster (anything below ' +
            `${T.KILL_ROAS}x after $${T.MIN_SPEND_FOR_JUDGMENT}+ spend) to limit the drag.`
        });
      }
    }

    // ── CHECK 10: Click→ATC Drop (landing page / product page issue) ──
    if (atcRateDelta < -10 && Math.abs(atcToPurchDelta) < 5 && Math.abs(ctrDelta) < 5) {
      findings.push({
        signal: '🟡 Landing Page Drop-off',
        severity: 'medium',
        evidence: `Click→ATC rate dropped ${Math.abs(atcRateDelta).toFixed(1)}% ` +
          `(${atcRateP7.toFixed(1)}% → ${atcRate7.toFixed(1)}%) but CTR and checkout completion are stable. ` +
          `People land on the page but don't add to cart.`,
        action: 'Check product page: price changes, OOS, reviews, page load speed, or ad-to-LP mismatch. ' +
          'If you changed creatives, the new angle may promise something the LP doesn\'t deliver.'
      });
    }

    // ── CHECK 11: Spend Reallocation (budget shifted between campaigns) ─
    const campL7 = {};
    const campP7 = {};
    last7Rows.forEach(row => {
      if (!campL7[row.campaign]) campL7[row.campaign] = { spend: 0, revenue: 0 };
      campL7[row.campaign].spend += row.spend;
      campL7[row.campaign].revenue += row.revenue;
    });
    prior7Rows.forEach(row => {
      if (!campP7[row.campaign]) campP7[row.campaign] = { spend: 0, revenue: 0 };
      campP7[row.campaign].spend += row.spend;
      campP7[row.campaign].revenue += row.revenue;
    });

    // Find if a low-ROAS campaign gained share while a high-ROAS campaign lost share
    const allCampNames = [...new Set([...Object.keys(campL7), ...Object.keys(campP7)])];
    let spendShiftEvidence = '';
    allCampNames.forEach(name => {
      const l = campL7[name] || { spend: 0, revenue: 0 };
      const p = campP7[name] || { spend: 0, revenue: 0 };
      const lShare = last7.spend > 0 ? (l.spend / last7.spend * 100) : 0;
      const pShare = prior7.spend > 0 ? (p.spend / prior7.spend * 100) : 0;
      const shareDelta = lShare - pShare;
      const campRoas = l.spend > 0 ? l.revenue / l.spend : 0;

      if (shareDelta > 5 && campRoas < T.TARGET_ROAS) {
        spendShiftEvidence += `${name} gained ${shareDelta.toFixed(0)}pp share (${campRoas.toFixed(2)}x ROAS). `;
      } else if (shareDelta < -5 && campRoas >= T.TARGET_ROAS) {
        spendShiftEvidence += `${name} lost ${Math.abs(shareDelta).toFixed(0)}pp share (${campRoas.toFixed(2)}x ROAS). `;
      }
    });
    if (spendShiftEvidence) {
      findings.push({
        signal: '🟡 Spend Reallocation',
        severity: 'medium',
        evidence: `Budget shifted between campaigns: ${spendShiftEvidence.trim()}`,
        action: 'Blended ROAS can drop purely from mix shift even if no campaign got worse. Check campaign-level ROAS below to confirm.'
      });
    }

    // ── CHECK 12: Weekend/Weekday Mix Shift ───────────────────────────
    function countWeekendDays(dates) {
      return dates.filter(d => {
        const parts = d.split('-');
        const day = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2])).getDay();
        return day === 0 || day === 6;
      }).length;
    }
    const l7Weekends = countWeekendDays(last7Dates);
    const p7Weekends = countWeekendDays(prior7Dates);
    if (Math.abs(l7Weekends - p7Weekends) >= 2) {
      findings.push({
        signal: '🟡 Day Mix Shift',
        severity: 'medium',
        evidence: `Last 7d had ${l7Weekends} weekend days vs ${p7Weekends} in prior 7d. ` +
          `Performance naturally varies by day of week, so WoW comparison is skewed.`,
        action: 'Check meta_deep_dives DOW table for your best/worst days. If weekends perform differently, this explains part of the shift.'
      });
    }

    // If nothing specific found, give a general note
    // (quality trade-off findings count as explanations)
    const hasRealFindings = findings.some(f => f.severity !== 'none');
    const hasQualityFindings = findings.some(f => f.signal.includes('Quality') || f.signal.includes('Refinement'));
    if (findings.length === 0) {
      findings.push({
        signal: '🟡 ROAS Declined — No Clear Single Cause',
        severity: 'medium',
        evidence: `ROAS dropped ${Math.abs(roasDelta).toFixed(1)}% but no single metric shows a dramatic shift. ` +
          `Multiple small declines are compounding.`,
        action: 'Review the creative tab for any ▼ trends. Check if spend shifted to lower-performing campaigns. ' +
          'Sometimes it\'s just weekly variance — compare over 14d if this persists.'
      });
    } else if (!hasRealFindings && hasQualityFindings) {
      // ROAS dropped but it's explained by quality trade-offs — add context
      findings.push({
        signal: '🟡 Volume vs Efficiency Trade-off',
        severity: 'medium',
        evidence: `ROAS dropped ${Math.abs(roasDelta).toFixed(1)}% but the metrics suggest Meta is trading volume for quality. ` +
          `Check if total purchase volume is acceptable. If you need more volume, expand audiences.`,
        action: 'If purchase count is too low, the quality optimization is too aggressive. Try broader targeting or raise budget to give Meta more room.'
      });
    }
  } else if (roasImproved) {
    // Check if this is a budget cut rebound (spend down significantly + ROAS up)
    if (spendDelta < -10) {
      findings.push({
        signal: '🟡 Budget Cut Rebound',
        severity: 'medium',
        evidence: `ROAS improved ${roasDelta.toFixed(1)}% but spend dropped ${Math.abs(spendDelta).toFixed(1)}%. ` +
          `When you cut budget, remaining spend concentrates on the best-performing audiences, inflating ROAS.`,
        action: 'Don\'t mistake this for real improvement. The efficiency gain came from spending less, not performing better. ' +
          'If you scale spend back up, expect ROAS to come back down.'
      });
    } else {
      findings.push({
        signal: '🟢 Performance Improving',
        severity: 'none',
        evidence: `ROAS up ${roasDelta.toFixed(1)}% WoW (${prior7.roas.toFixed(2)}x → ${last7.roas.toFixed(2)}x). `,
        action: cvrDelta > 5
          ? `CVR improved ${cvrDelta.toFixed(1)}%. Funnel is converting better. Test increasing CBO budgets 20-30%.`
          : ctrDelta > 5
          ? `CTR up ${ctrDelta.toFixed(1)}%. Creatives are resonating. Test scaling top campaigns.`
          : `Broad efficiency gains. Test scaling daily budgets 20-30% on top campaigns.`
      });
    }
  } else {
    findings.push({
      signal: '✅ Stable Performance',
      severity: 'none',
      evidence: `ROAS moved ${roasDelta >= 0 ? '+' : ''}${roasDelta.toFixed(1)}% WoW — within normal variance.`,
      action: 'No fires. Focus on testing new creatives and scaling winners.'
    });
  }

  // ── Cross-channel: Google Ads campaign health ────────────────────────
  if (hasGoogleAds && gAdsL7.spend > 0) {
    // Check for wasteful Google campaigns
    const gAdsSheet3 = ss.getSheetByName(CONFIG.TABS.GOOGLE_ADS);
    if (gAdsSheet3 && gAdsSheet3.getLastRow() > 1) {
      const gh3 = gAdsSheet3.getRange(1, 1, 1, gAdsSheet3.getLastColumn()).getValues()[0];
      const gd3 = gAdsSheet3.getRange(2, 1, gAdsSheet3.getLastRow() - 1, gAdsSheet3.getLastColumn()).getValues();
      const gi3 = {}; gh3.forEach((h, i) => gi3[h] = i);

      const gCampL7 = {};
      gd3.forEach(row => {
        let date = row[gi3['date']];
        if (date instanceof Date) date = Utilities.formatDate(date, 'UTC', 'yyyy-MM-dd');
        else date = String(date).substring(0, 10);
        if (!completedDates.includes(date)) return;
        if (!completedDates.slice(-7).includes(date)) return;
        const name = row[gi3['campaign']] || '';
        if (!gCampL7[name]) gCampL7[name] = { spend: 0, conv: 0, revenue: 0 };
        gCampL7[name].spend += parseFloat(row[gi3['spend']]) || 0;
        gCampL7[name].conv += parseFloat(row[gi3['conversions']]) || 0;
        gCampL7[name].revenue += parseFloat(row[gi3['conversion_value']]) || 0;
      });

      const wasteful = Object.entries(gCampL7)
        .filter(([name, d]) => d.spend > 30 && d.conv < 1)
        .map(([name, d]) => `${name} ($${d.spend.toFixed(0)} spent, ${d.conv.toFixed(0)} conversions)`);

      if (wasteful.length > 0) {
        findings.push({
          signal: '🔵 Google Ads: Wasteful Campaign',
          severity: 'medium',
          evidence: `${wasteful.join('; ')}. Spending with no return.`,
          action: 'Pause these campaigns and consolidate budget into your performing PMax campaign.'
        });
      }
    }
  }

  // ── Build trend narrative (synthesizes the whole picture) ────────────
  let trendNarrative = '';
  const spendUp = spendDelta > 5;
  const spendDown = spendDelta < -5;

  if (qualityTradeOff || (cvrImproved && cpcUp)) {
    trendNarrative = `Meta is optimizing for quality over volume. ` +
      `CPC up ${cpcDelta.toFixed(0)}% but CVR up ${cvrDelta.toFixed(0)}%, so CPA only moved ${cpaDelta.toFixed(0)}%. `;
    if (purchDeltaEarly < -5) {
      trendNarrative += `However, purchase volume dropped ${Math.abs(purchDeltaEarly).toFixed(0)}% (${prior7.purchases} → ${last7.purchases}). Watch if this continues.`;
    } else {
      trendNarrative += `Purchase volume is holding (${last7.purchases} vs ${prior7.purchases} prior). The algorithm is working.`;
    }
  } else if (roasDropped && spendDelta > 12) {
    // Scaling narrative takes priority if spend ramped
    trendNarrative = `You scaled spend ${spendDelta.toFixed(0)}% WoW and ROAS dipped ${Math.abs(roasDelta).toFixed(0)}%. ` +
      `This is normal — more spend reaches further into the audience at diminishing returns. ` +
      `If total revenue and purchase volume grew, the scaling is working even if ROAS is lower.`;
  } else if (roasDropped && ctrDown && !cvrImproved) {
    trendNarrative = `Performance is declining across the funnel. ` +
      `CTR down ${Math.abs(ctrDelta).toFixed(0)}%, CPC up ${cpcDelta.toFixed(0)}%, and CVR didn't compensate. `;
    if (avgFreq7 > T.FREQUENCY_WARNING) {
      trendNarrative += `Frequency at ${avgFreq7.toFixed(1)} suggests ad fatigue. Fresh creatives are the priority.`;
    } else {
      trendNarrative += `Frequency is still reasonable (${avgFreq7.toFixed(1)}), so this may be audience-level fatigue rather than ad-level. Test new angles.`;
    }
  } else if (roasDropped && cpmDelta > 8 && Math.abs(ctrDelta) < 5) {
    trendNarrative = `Auction costs are driving the ROAS decline, not your ads. ` +
      `CPM up ${cpmDelta.toFixed(0)}% while CTR and CVR are relatively stable. ` +
      `Likely competitive pressure or seasonal. If it persists beyond 2 weeks, improve CTR to offset.`;
  } else if (roasDropped && aovDelta < -5) {
    trendNarrative = `Revenue per order is shrinking. AOV down ${Math.abs(aovDelta).toFixed(0)}% ($${prior7.aov.toFixed(0)} → $${last7.aov.toFixed(0)}). ` +
      `Your ads are driving the same traffic but people are buying cheaper products or smaller orders.`;
  } else if (roasDropped && atcRateDelta < -10) {
    trendNarrative = `The leak is between click and add-to-cart. People land on the page but aren't adding to cart ` +
      `(Click→ATC rate dropped ${Math.abs(atcRateDelta).toFixed(0)}%). This is a landing page or product page problem, not an ad problem.`;
  } else if (roasImproved && spendDelta < -10) {
    trendNarrative = `ROAS improved ${roasDelta.toFixed(0)}% but spend dropped ${Math.abs(spendDelta).toFixed(0)}%. ` +
      `This is likely a budget cut rebound — remaining spend concentrates on the best audiences. Don't mistake this for real improvement.`;
  } else if (roasImproved) {
    trendNarrative = `Strong week. ROAS improved ${roasDelta.toFixed(0)}% WoW. `;
    if (cvrDelta > 5) trendNarrative += `Funnel efficiency is the driver — CVR up ${cvrDelta.toFixed(0)}%. `;
    if (spendUp) trendNarrative += `Scaling well — spend up and efficiency up simultaneously.`;
    else trendNarrative += `Consider testing 20-30% budget increases on top campaigns while momentum holds.`;
  } else if (!roasDropped) {
    trendNarrative = `Steady state. ROAS within normal variance. Focus on creative testing and incremental scaling.`;
  } else {
    trendNarrative = `ROAS declined ${Math.abs(roasDelta).toFixed(0)}% with multiple small metric shifts. No single clear cause — likely compounding effects. Review creative tab for ▼ trends.`;
  }

  // ── One-line status ───────────────────────────────────────────────────
  let statusLine = `${allAgg.roas.toFixed(2)}x ROAS`;
  let statusColor = '#137333';
  if (allAgg.roas < T.KILL_ROAS) {
    statusLine += ` — Below ${T.KILL_ROAS}x kill floor`;
    statusColor = '#C5221F';
  } else if (allAgg.roas < T.TARGET_ROAS) {
    statusLine += ` — Below ${T.TARGET_ROAS}x target`;
    statusColor = '#E37400';
  } else {
    statusLine += ` — Above target`;
  }
  if (roasDropped) {
    statusLine += ` | ROAS dropped ${Math.abs(roasDelta).toFixed(0)}% WoW — see diagnosis below`;
    if (statusColor === '#137333') statusColor = '#E37400';
  }


  // ══════════════════════════════════════════════════════════════════════
  // WRITE THE DASHBOARD
  // ══════════════════════════════════════════════════════════════════════
  let r = 1;

  // Title
  sheet.getRange(r, 1).setValue(hasGoogleAds ? 'ADS DASHBOARD' : 'META ADS DASHBOARD');
  sheet.getRange(r, 1).setFontSize(14).setFontWeight('bold');
  r += 1;
  sheet.getRange(r, 1).setValue(`Last synced: ${new Date().toLocaleString()}`);
  sheet.getRange(r, 1).setFontColor('#666666').setFontSize(10);
  r += 1;

  // Status line
  const statusRoas = hasGoogleAds ? combinedL7.roas : allAgg.roas;
  const statusPrefix = hasGoogleAds ? 'Combined ' : '';
  sheet.getRange(r, 1, 1, 13).merge();
  let displayStatus = hasGoogleAds
    ? `${statusPrefix}${combinedL7.roas.toFixed(2)}x ROAS (Meta ${last7.roas.toFixed(2)}x + Google ${gAdsL7.roas.toFixed(2)}x)`
    : statusLine;
  if (hasGoogleAds && roasDropped) displayStatus += ` | Meta ROAS dropped ${Math.abs(roasDelta).toFixed(0)}% WoW`;
  sheet.getRange(r, 1).setValue(hasGoogleAds ? displayStatus : statusLine);
  sheet.getRange(r, 1).setFontSize(11).setFontWeight('bold').setFontColor(statusColor).setWrap(true);
  sheet.setRowHeight(r, 30);
  r += 1;

  // ── COMBINED TOTALS (Meta + Google Ads — only if Google Ads data exists) ──
  if (hasGoogleAds) {
    // Header: channel | L7 Spend | L7 Revenue | L7 Purch | L7 ROAS | L7 CPA | WoW ROAS
    const chHeaders = ['Channel', 'L7 Spend', 'L7 Revenue', 'L7 Purch', 'L7 ROAS', 'L7 CPA'];
    sheet.getRange(r, 1, 1, chHeaders.length).setValues([chHeaders]);
    sheet.getRange(r, 1, 1, chHeaders.length)
      .setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#FFFFFF')
      .setHorizontalAlignment('center').setFontSize(8);
    r += 1;

    const channelRows = [
      ['Meta Ads', last7.spend, last7.revenue, last7.purchases, last7.roas, last7.cpa],
      ['Google Ads', gAdsL7.spend, gAdsL7.revenue, gAdsL7.purchases, gAdsL7.roas, gAdsL7.cpa],
      ['TOTAL', combinedL7.spend, combinedL7.revenue, combinedL7.purchases, combinedL7.roas, combinedL7.cpa],
    ];

    channelRows.forEach((cr, i) => {
      sheet.getRange(r, 1, 1, chHeaders.length).setValues([cr]);
      sheet.getRange(r, 2).setNumberFormat('"$"#,##0');
      sheet.getRange(r, 3).setNumberFormat('"$"#,##0');
      sheet.getRange(r, 4).setNumberFormat('#,##0');
      sheet.getRange(r, 5).setNumberFormat('0.00"x"');
      sheet.getRange(r, 6).setNumberFormat('"$"#,##0.00');
      sheet.getRange(r, 1, 1, chHeaders.length).setHorizontalAlignment('center').setVerticalAlignment('middle').setFontSize(9);
      sheet.getRange(r, 1).setHorizontalAlignment('left').setFontWeight('bold');

      // ROAS coloring
      const roasVal = cr[4];
      const roasCell = sheet.getRange(r, 5);
      if (roasVal >= T.SCALE_ROAS) roasCell.setFontColor('#137333').setFontWeight('bold');
      else if (roasVal >= T.TARGET_ROAS) roasCell.setFontColor('#137333');
      else if (roasVal >= T.KILL_ROAS) roasCell.setFontColor('#E37400');
      else roasCell.setFontColor('#C5221F');

      // Total row styling
      if (i === 2) {
        sheet.getRange(r, 1, 1, chHeaders.length).setFontWeight('bold').setBackground('#e8eaf6');
      } else if (i % 2 === 0) {
        sheet.getRange(r, 1, 1, chHeaders.length).setBackground('#fafafa');
      }
      r++;
    });
  }

  // ── META KPI TABLE ────────────────────────────────────────────────────
  sheet.getRange(r, 1).setValue(hasGoogleAds ? 'META ADS DETAIL' : 'KPI OVERVIEW').setFontWeight('bold').setFontSize(11);
  r += 1;

  const kpiStartRow = r;
  const kpiHeaders = ['Metric', 'Last 7 Days', 'Prior 7 Days', 'WoW Change', `Full ${CONFIG.LOOKBACK_DAYS}d`];
  sheet.getRange(r, 1, 1, kpiHeaders.length).setValues([kpiHeaders]);
  sheet.getRange(r, 1, 1, kpiHeaders.length)
    .setFontWeight('bold').setBackground('#000000').setFontColor('#FFFFFF')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  r += 1;

  // Date range sub-header row
  const dateSubRow = ['', last7Label, prior7Label, 'Last 7d vs Prior 7d', fullRangeLabel];
  sheet.getRange(r, 1, 1, dateSubRow.length).setValues([dateSubRow]);
  sheet.getRange(r, 1, 1, dateSubRow.length)
    .setFontSize(9).setFontColor('#666666').setFontStyle('italic')
    .setHorizontalAlignment('center').setBackground('#f3f3f3');
  r += 1;

  // Metrics: only the 8 that drive decisions. Rest lives in deep_dives.
  const kpiConfig = [
    { name: 'Spend',       l: last7.spend,       p: prior7.spend,       inv: false, f: allAgg.spend },
    { name: 'Revenue',     l: last7.revenue,      p: prior7.revenue,     inv: false, f: allAgg.revenue },
    { name: 'Purchases',   l: last7.purchases,    p: prior7.purchases,   inv: false, f: allAgg.purchases },
    { name: 'ROAS',        l: last7.roas,         p: prior7.roas,        inv: false, f: allAgg.roas },
    { name: 'CPA',         l: last7.cpa,          p: prior7.cpa,         inv: true,  f: allAgg.cpa },
    { name: 'AOV',         l: last7.aov,          p: prior7.aov,         inv: false, f: allAgg.aov },
    { name: 'CVR',         l: last7.cvr,          p: prior7.cvr,         inv: false, f: allAgg.cvr },
    { name: 'CPC',         l: last7.cpc,          p: prior7.cpc,         inv: true,  f: allAgg.cpc },
  ];

  const kpiData = kpiConfig.map(m => [
    m.name, m.l, m.p, trendPct(m.l, m.p, m.inv), m.f
  ]);

  sheet.getRange(r, 1, kpiData.length, kpiHeaders.length).setValues(kpiData);
  sheet.getRange(r, 1, kpiData.length, kpiHeaders.length).setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRange(r, 1, kpiData.length, 1).setHorizontalAlignment('left').setFontWeight('bold');

  // Number formats per metric row
  const formats = {
    'Spend': '"$"#,##0.00', 'Revenue': '"$"#,##0.00',
    'Purchases': '#,##0', 'ROAS': '0.00"x"',
    'CPA': '"$"#,##0.00', 'AOV': '"$"#,##0.00',
    'CVR': '0.00"%"', 'CPC': '"$"#,##0.00',
  };

  for (let i = 0; i < kpiData.length; i++) {
    const metric = kpiConfig[i];
    const fmt = formats[metric.name];
    if (fmt) {
      sheet.getRange(r + i, 2).setNumberFormat(fmt);
      sheet.getRange(r + i, 3).setNumberFormat(fmt);
      sheet.getRange(r + i, 5).setNumberFormat(fmt);
    }

    // Smart WoW coloring — only color meaningful changes
    const rawPct = trendRawPct(metric.l, metric.p);
    const absPct = Math.abs(rawPct);
    const trendCell = sheet.getRange(r + i, 4);

    if (absPct >= WOW_NOISE_FLOOR) {
      // Significant change — color based on good/bad direction
      const isGoodDirection = metric.inv ? (rawPct < 0) : (rawPct > 0);
      if (isGoodDirection) {
        trendCell.setFontColor('#137333').setFontWeight('bold');
      } else {
        // Only go red for big swings (>15%), orange for moderate (5-15%)
        if (absPct >= 15) {
          trendCell.setFontColor('#C5221F').setFontWeight('bold');
        } else {
          trendCell.setFontColor('#E37400'); // Orange = "heads up" not "panic"
        }
      }
    } else {
      // Small change — keep it neutral grey
      trendCell.setFontColor('#888888');
    }

    // Alternate row shading
    if (i % 2 === 0) {
      sheet.getRange(r + i, 1, 1, kpiHeaders.length).setBackground('#fafafa');
    }
  }

  // Color-code ROAS cells specifically (Last 7d = col2, Prior 7d = col3, Full Period = col5)
  const roasRowIdx = kpiData.findIndex(d => d[0] === 'ROAS');
  if (roasRowIdx >= 0) {
    const roasColMap = { 2: 1, 3: 2, 5: 4 }; // sheet col → kpiData array index
    [2, 3, 5].forEach(col => {
      const val = parseFloat(kpiData[roasRowIdx][roasColMap[col]]) || 0;
      const cell = sheet.getRange(r + roasRowIdx, col);
      if (val >= T.TARGET_ROAS) cell.setFontColor('#137333').setFontWeight('bold');
      else if (val >= T.KILL_ROAS) cell.setFontColor('#E37400').setFontWeight('bold');
      else cell.setFontColor('#C5221F').setFontWeight('bold');
    });
  }

  r += kpiData.length;


  // ── Funnel rate computation (used by diagnosis, right panel, and deep dives) ──
  const funnelL7 = {
    imprToClick: last7.impressions > 0 ? (last7.clicks / last7.impressions * 100) : 0,
    clickToAtc: last7.clicks > 0 ? (last7.atc / last7.clicks * 100) : 0,
    atcToPurch: last7.atc > 0 ? (last7.purchases / last7.atc * 100) : 0,
    clickToPurch: last7.clicks > 0 ? (last7.purchases / last7.clicks * 100) : 0,
  };
  const funnelP7 = {
    imprToClick: prior7.impressions > 0 ? (prior7.clicks / prior7.impressions * 100) : 0,
    clickToAtc: prior7.clicks > 0 ? (prior7.atc / prior7.clicks * 100) : 0,
    atcToPurch: prior7.atc > 0 ? (prior7.purchases / prior7.atc * 100) : 0,
    clickToPurch: prior7.clicks > 0 ? (prior7.purchases / prior7.clicks * 100) : 0,
  };
  // ── DIAGNOSIS SECTION ─────────────────────────────────────────────────
  const actionableFindings = findings.filter(f => f.severity !== 'none');
  const infoFindings = findings.filter(f => f.severity === 'none');

  if (actionableFindings.length > 0 || infoFindings.length > 0) {
    sheet.getRange(r, 1).setValue('DIAGNOSIS').setFontWeight('bold').setFontSize(11);
    r += 1;

    // Trend narrative — the story, not just the symptoms
    if (trendNarrative) {
      sheet.getRange(r, 1, 1, 6).merge().setValue(trendNarrative);
      sheet.getRange(r, 1).setWrap(true).setFontSize(9).setFontStyle('italic').setFontColor('#333333')
        .setBackground('#f0f4ff').setVerticalAlignment('middle');
      sheet.setRowHeight(r, 35);
      r += 1;
    }

    // Individual findings
    const orderedFindings = [...infoFindings, ...actionableFindings];

    orderedFindings.forEach((f, i) => {
      sheet.getRange(r, 1).setValue(f.signal).setFontWeight('bold').setFontSize(9);
      sheet.getRange(r, 2, 1, 5).merge().setValue(`${f.evidence}  →  ${f.action}`);
      sheet.getRange(r, 2).setWrap(true).setFontSize(9);
      sheet.getRange(r, 1, 1, 6).setVerticalAlignment('middle');

      if (f.severity === 'high') {
        sheet.getRange(r, 1, 1, 6).setBackground('#fde8e8');
      } else if (f.severity === 'medium') {
        sheet.getRange(r, 1, 1, 6).setBackground('#fef7e0');
      } else if (f.severity === 'none') {
        sheet.getRange(r, 1, 1, 6).setBackground('#e6f4ea');
      }

      sheet.setRowHeight(r, 40);
      r++;
    });
  } else if (findings.length > 0) {
    // Stable/improving — trend narrative + summary
    if (trendNarrative) {
      sheet.getRange(r, 1, 1, 6).merge().setValue(trendNarrative);
      sheet.getRange(r, 1).setFontSize(10).setFontColor('#137333').setWrap(true).setFontWeight('bold');
      sheet.setRowHeight(r, 30);
      r += 1;
    } else {
      sheet.getRange(r, 1, 1, 5).merge();
      const f = findings[0];
      sheet.getRange(r, 1).setValue(`${f.signal}  ${f.evidence} ${f.action}`);
      sheet.getRange(r, 1).setFontSize(10).setFontColor('#137333').setWrap(true).setFontWeight('bold');
      sheet.setRowHeight(r, 30);
      r += 1;
    }
  }

  // ── AI INSIGHT (if configured) ──────────────────────────────────────
  if (CONFIG.AI.API_KEY) {
    // Build data packet for the AI
    const aiData = {
      period: { last7: last7Label, prior7: prior7Label, full: fullRangeLabel },
      kpi_last7: {
        spend: last7.spend, revenue: last7.revenue, purchases: last7.purchases,
        roas: last7.roas, cpa: last7.cpa, aov: last7.aov, cvr: last7.cvr,
        cpc: last7.cpc, cpm: last7.cpm, ctr: last7.ctr,
        frequency: avgFreq7, reach: last7.reach, impressions: last7.impressions
      },
      kpi_prior7: {
        spend: prior7.spend, revenue: prior7.revenue, purchases: prior7.purchases,
        roas: prior7.roas, cpa: prior7.cpa, aov: prior7.aov, cvr: prior7.cvr,
        cpc: prior7.cpc, cpm: prior7.cpm, ctr: prior7.ctr,
        frequency: avgFreqP7, reach: prior7.reach, impressions: prior7.impressions
      },
      wow_deltas: {
        roas: roasDelta, ctr: ctrDelta, cpm: cpmDelta, cvr: cvrDelta,
        cpa: cpaDelta, cpc: cpcDelta, aov: aovDelta, spend: spendDelta,
        reach: reachDelta, frequency: freqDelta
      },
      diagnosis_signals: findings.map(f => ({ signal: f.signal, severity: f.severity, evidence: f.evidence })),
      campaigns: activeCamps.map(c => ({
        name: c.name, l7_spend: c.l.spend, l7_roas: c.l.roas, l7_cpa: c.l.cpa,
        p7_spend: c.p.spend, p7_roas: c.p.roas, p7_cpa: c.p.cpa
      })),
      funnel: {
        last7: funnelL7, prior7: funnelP7,
        atc_to_purchase_delta: atcToPurchDelta
      },
      creative_velocity: {
        testing: testing.length,
        test_results: testResults.length,
        active: active.length,
        declining: DATA_CACHE.creatives ? DATA_CACHE.creatives.filter(c => c.roasTrend === '▼' && c.phase === 'active').length : 0,
        top_concentration_pct: totalActiveSpend > 0 ? (topCreativeSpend / totalActiveSpend * 100) : 0,
      },
      thresholds: { target_roas: T.TARGET_ROAS, kill_roas: T.KILL_ROAS, scale_roas: T.SCALE_ROAS }
    };

    const aiInsight = generateAIInsight(aiData);

    if (aiInsight && !aiInsight.startsWith('[AI Error')) {
      sheet.getRange(r, 1).setValue('🤖 AI INSIGHT').setFontWeight('bold').setFontSize(12);
      r += 1;
      sheet.getRange(r, 1, 1, 5).merge();
      sheet.getRange(r, 1).setValue(aiInsight).setWrap(true).setVerticalAlignment('top')
        .setFontSize(10).setBackground('#f0f4ff').setFontColor('#1a1a1a');
      // Auto-height based on text length
      const lineCount = Math.ceil(aiInsight.length / 80);
      sheet.setRowHeight(r, Math.max(80, lineCount * 18));
      r += 1;
    } else if (aiInsight) {
      // Show error for debugging
      sheet.getRange(r, 1).setValue(aiInsight).setFontColor('#999999').setFontSize(9);
      r += 1;
    }
  }

  // ── CAMPAIGNS WoW (right panel, cols G-M, aligned with KPI) ─────────
  const campMap = {};
  dailyRows.forEach(dr2 => {
    if (!campMap[dr2.campaign]) campMap[dr2.campaign] = { last7: [], prior7: [], source: 'Meta' };
  });
  last7Rows.forEach(dr2 => { if (campMap[dr2.campaign]) campMap[dr2.campaign].last7.push(dr2); });
  prior7Rows.forEach(dr2 => { if (campMap[dr2.campaign]) campMap[dr2.campaign].prior7.push(dr2); });

  // Include Google Ads campaigns if data exists
  if (hasGoogleAds) {
    const gAdsSheet2 = ss.getSheetByName(CONFIG.TABS.GOOGLE_ADS);
    if (gAdsSheet2 && gAdsSheet2.getLastRow() > 1) {
      const gh = gAdsSheet2.getRange(1, 1, 1, gAdsSheet2.getLastColumn()).getValues()[0];
      const gd = gAdsSheet2.getRange(2, 1, gAdsSheet2.getLastRow() - 1, gAdsSheet2.getLastColumn()).getValues();
      const gIdx = {}; gh.forEach((h, i) => gIdx[h] = i);
      gd.forEach(row => {
        let date = row[gIdx['date']];
        if (date instanceof Date) date = Utilities.formatDate(date, 'UTC', 'yyyy-MM-dd');
        else date = String(date).substring(0, 10);
        const name = '🔵 ' + (row[gIdx['campaign']] || 'Google Ads');
        if (!campMap[name]) campMap[name] = { last7: [], prior7: [], source: 'Google' };
        const gRow = {
          date, campaign: name,
          spend: parseFloat(row[gIdx['spend']]) || 0,
          revenue: parseFloat(row[gIdx['conversion_value']]) || 0,
          purchases: parseFloat(row[gIdx['conversions']]) || 0,
          impressions: parseInt(row[gIdx['impressions']]) || 0,
          clicks: parseInt(row[gIdx['clicks']]) || 0,
          atc: 0, reach: 0,
        };
        if (last7Dates.includes(date)) campMap[name].last7.push(gRow);
        if (prior7Dates.includes(date)) campMap[name].prior7.push(gRow);
      });
    }
  }

  const activeCamps = Object.entries(campMap)
    .map(([name, data]) => ({ name, l: aggregate(data.last7), p: aggregate(data.prior7), source: data.source }))
    .filter(c => c.l.spend >= 50 || c.p.spend >= 50)  // Skip dead/paused campaigns
    .sort((a, b) => b.l.spend - a.l.spend);

  const RC = 7; // right column start (col G)
  let rr = kpiStartRow;

  sheet.getRange(rr, RC).setValue('ALL CAMPAIGNS (WoW)').setFontWeight('bold').setFontSize(11);
  rr += 1;

  if (activeCamps.length > 0) {
    const campHeaders = ['Campaign', 'L7 Spend', 'L7 ROAS', 'L7 CPA', 'L7 Purch', 'ROAS Δ', 'CPA Δ'];
    sheet.getRange(rr, RC, 1, campHeaders.length).setValues([campHeaders]);
    sheet.getRange(rr, RC, 1, campHeaders.length)
      .setFontWeight('bold').setBackground('#000000').setFontColor('#FFFFFF')
      .setHorizontalAlignment('center').setFontSize(9);
    rr += 1;

    activeCamps.forEach((c, i) => {
      sheet.getRange(rr, RC, 1, campHeaders.length).setValues([[
        c.name, c.l.spend, c.l.roas, c.l.cpa, c.l.purchases,
        trendPct(c.l.roas, c.p.roas, false), trendPct(c.l.cpa, c.p.cpa, true)
      ]]);
      sheet.getRange(rr, RC + 1).setNumberFormat('"$"#,##0');
      sheet.getRange(rr, RC + 2).setNumberFormat('0.00"x"');
      sheet.getRange(rr, RC + 3).setNumberFormat('"$"#,##0.00');
      sheet.getRange(rr, RC + 4).setNumberFormat('#,##0');
      sheet.getRange(rr, RC, 1, campHeaders.length).setHorizontalAlignment('center').setVerticalAlignment('middle').setFontSize(9);
      sheet.getRange(rr, RC).setHorizontalAlignment('left').setFontWeight('bold').setFontSize(8);

      // Color ROAS delta
      const rawRD = trendRawPct(c.l.roas, c.p.roas);
      const rdCell = sheet.getRange(rr, RC + 5);
      if (Math.abs(rawRD) < WOW_NOISE_FLOOR) rdCell.setFontColor('#888888');
      else if (rawRD > 0) rdCell.setFontColor('#137333').setFontWeight('bold');
      else if (rawRD < -15) rdCell.setFontColor('#C5221F').setFontWeight('bold');
      else rdCell.setFontColor('#E37400');

      // Color ROAS value
      const roasCell = sheet.getRange(rr, RC + 2);
      if (c.l.roas >= T.SCALE_ROAS) roasCell.setBackground('#e6f4ea').setFontColor('#137333');
      else if (c.l.roas >= T.TARGET_ROAS) roasCell.setFontColor('#137333');
      else if (c.l.roas >= T.KILL_ROAS) roasCell.setFontColor('#E37400');
      else if (c.l.spend > 0) roasCell.setFontColor('#C5221F');

      // Google campaigns get a subtle blue tint
      if (c.source === 'Google') {
        sheet.getRange(rr, RC, 1, campHeaders.length).setBackground(i % 2 === 0 ? '#e8f0fe' : '#f5f8ff');
      } else {
        if (i % 2 === 0) sheet.getRange(rr, RC, 1, campHeaders.length).setBackground('#fafafa');
      }
      rr++;
    });
  }

  // r = max of left and right panel ends
  r = Math.max(r, rr) + 1;

  // ── Column widths ─────────────────────────────────────────────────────
  // Left: Channel + KPI (A-F)
  sheet.setColumnWidth(1, 130);  // Metric / Channel
  sheet.setColumnWidth(2, 110);  // Last 7d / L7 Spend
  sheet.setColumnWidth(3, 110);  // Prior 7d / L7 Revenue
  sheet.setColumnWidth(4, 95);   // WoW / L7 Purch
  sheet.setColumnWidth(5, 110);  // Full 30d / L7 ROAS
  sheet.setColumnWidth(6, 85);   // L7 CPA (channel table uses this)
  // Right: Campaigns (G-M)
  sheet.setColumnWidth(7, 175);  // Campaign name
  sheet.setColumnWidth(8, 80);   // L7 Spend
  sheet.setColumnWidth(9, 65);   // L7 ROAS
  sheet.setColumnWidth(10, 75);  // L7 CPA
  sheet.setColumnWidth(11, 65);  // L7 Purch
  sheet.setColumnWidth(12, 65);  // ROAS Δ
  sheet.setColumnWidth(13, 65);  // CPA Δ
  // Trim beyond col 13
  try {
    if (sheet.getMaxColumns() > 13) sheet.deleteColumns(14, sheet.getMaxColumns() - 13);
  } catch (e) {}

  // Trim empty rows
  try {
    const lastDataRow = sheet.getLastRow();
    if (sheet.getMaxRows() > lastDataRow + 1) sheet.deleteRows(lastDataRow + 2, sheet.getMaxRows() - lastDataRow - 1);
  } catch (e) {}

  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(1);

  // ══════════════════════════════════════════════════════════════════════
  // 📈 DEEP DIVES TAB — DOW, Funnel Charts, Weekly Charts
  // ══════════════════════════════════════════════════════════════════════
  let ddSheet = ss.getSheetByName(CONFIG.TABS.META_DEEP);
  if (!ddSheet) ddSheet = ss.insertSheet(CONFIG.TABS.META_DEEP);
  ddSheet.clear();
  ddSheet.clearConditionalFormatRules();
  const ddOldCharts = ddSheet.getCharts();
  ddOldCharts.forEach(c => ddSheet.removeChart(c));

  let dr = 1;
  ddSheet.getRange(dr, 1).setValue('DEEP DIVES').setFontWeight('bold').setFontSize(14);
  dr += 1;
  ddSheet.getRange(dr, 1).setValue(`${dateRange.since} → ${dateRange.until}`).setFontSize(9).setFontColor('#666666');
  dr += 2;

  // ── FULL METRICS (everything the dashboard trimmed) ─────────────────
  ddSheet.getRange(dr, 1).setValue('ALL METRICS (WoW)').setFontWeight('bold').setFontSize(12);
  dr += 1;

  const ddKpiHeaders = ['Metric', 'Last 7d', 'Prior 7d', 'WoW Δ', `Full ${CONFIG.LOOKBACK_DAYS}d`];
  ddSheet.getRange(dr, 1, 1, 5).setValues([ddKpiHeaders]);
  ddSheet.getRange(dr, 1, 1, 5)
    .setFontWeight('bold').setBackground('#000000').setFontColor('#FFFFFF').setHorizontalAlignment('center');
  dr += 1;

  const ddMetrics = [
    ['Spend', last7.spend, prior7.spend, trendPct(last7.spend, prior7.spend, false), allAgg.spend, '"$"#,##0'],
    ['Revenue', last7.revenue, prior7.revenue, trendPct(last7.revenue, prior7.revenue, false), allAgg.revenue, '"$"#,##0'],
    ['Purchases', last7.purchases, prior7.purchases, trendPct(last7.purchases, prior7.purchases, false), allAgg.purchases, '#,##0'],
    ['ROAS', last7.roas, prior7.roas, trendPct(last7.roas, prior7.roas, false), allAgg.roas, '0.00"x"'],
    ['CPA', last7.cpa, prior7.cpa, trendPct(last7.cpa, prior7.cpa, true), allAgg.cpa, '"$"#,##0.00'],
    ['AOV', last7.aov, prior7.aov, trendPct(last7.aov, prior7.aov, false), allAgg.aov, '"$"#,##0.00'],
    ['CVR', last7.cvr, prior7.cvr, trendPct(last7.cvr, prior7.cvr, false), allAgg.cvr, '0.00"%"'],
    ['CPC', last7.cpc, prior7.cpc, trendPct(last7.cpc, prior7.cpc, true), allAgg.cpc, '"$"#,##0.00'],
    ['CPM', last7.cpm, prior7.cpm, trendPct(last7.cpm, prior7.cpm, true), allAgg.cpm, '"$"#,##0.00'],
    ['CTR', last7.ctr, prior7.ctr, trendPct(last7.ctr, prior7.ctr, false), allAgg.ctr, '0.00"%"'],
    ['Frequency', avgFreq7, avgFreqP7, trendPct(avgFreq7, avgFreqP7, true), '', '0.00'],
    ['Reach', last7.reach, prior7.reach, trendPct(last7.reach, prior7.reach, false), allAgg.reach, '#,##0'],
    ['Impressions', last7.impressions, prior7.impressions, trendPct(last7.impressions, prior7.impressions, false), allAgg.impressions, '#,##0'],
    ['Clicks', last7.clicks, prior7.clicks, trendPct(last7.clicks, prior7.clicks, false), allAgg.clicks, '#,##0'],
    ['ATC', last7.atc, prior7.atc, trendPct(last7.atc, prior7.atc, false), allAgg.atc, '#,##0'],
  ];

  ddMetrics.forEach((m, i) => {
    ddSheet.getRange(dr, 1, 1, 5).setValues([[m[0], m[1], m[2], m[3], m[4]]]);
    ddSheet.getRange(dr, 1).setFontWeight('bold').setHorizontalAlignment('left');
    ddSheet.getRange(dr, 2).setNumberFormat(m[5]);
    ddSheet.getRange(dr, 3).setNumberFormat(m[5]);
    if (m[4] !== '') ddSheet.getRange(dr, 5).setNumberFormat(m[5]);
    ddSheet.getRange(dr, 1, 1, 5).setHorizontalAlignment('center');
    ddSheet.getRange(dr, 1).setHorizontalAlignment('left');
    if (i % 2 === 0) ddSheet.getRange(dr, 1, 1, 5).setBackground('#fafafa');

    // Color WoW delta (col 4)
    // Note: trendPct already handles inversion (▲ always = good direction, ▼ always = bad)
    const deltaStr = m[3];
    if (typeof deltaStr === 'string') {
      const isUp = deltaStr.includes('▲');
      const pctMatch = deltaStr.match(/([\d.]+)%/);
      const pctVal = pctMatch ? parseFloat(pctMatch[1]) : 0;
      const dCell = ddSheet.getRange(dr, 4);
      if (pctVal < 5) {
        dCell.setFontColor('#888888');
      } else if (isUp) {
        dCell.setFontColor('#137333').setFontWeight('bold');
        if (pctVal >= 15) dCell.setBackground('#e6f4ea');
      } else {
        dCell.setFontColor('#C5221F').setFontWeight('bold');
        if (pctVal >= 15) dCell.setBackground('#fde8e8');
      }
    }

    // Color ROAS row specifically
    if (m[0] === 'ROAS') {
      [2, 3, 5].forEach(col => {
        const val = parseFloat(ddSheet.getRange(dr, col).getValue()) || 0;
        const cell = ddSheet.getRange(dr, col);
        if (val >= T.SCALE_ROAS) cell.setFontColor('#137333').setFontWeight('bold');
        else if (val >= T.TARGET_ROAS) cell.setFontColor('#137333');
        else if (val >= T.KILL_ROAS) cell.setFontColor('#E37400');
        else cell.setFontColor('#C5221F');
      });
    }
    dr++;
  });
  dr += 1;

  // ── FUNNEL CHARTS ─────────────────────────────────────────────────────
  ddSheet.getRange(dr, 1).setValue('FUNNEL VISUALIZATION').setFontWeight('bold').setFontSize(12);
  dr += 1;

  // Raw volume table (cols 1-3)
  const ddVolTable = [
    ['Stage', 'Last 7d', 'Prior 7d'],
    ['Impressions', last7.impressions, prior7.impressions],
    ['Clicks (Sessions)', last7.clicks, prior7.clicks],
    ['Add to Cart', last7.atc, prior7.atc],
    ['Purchase', last7.purchases, prior7.purchases],
  ];
  ddSheet.getRange(dr, 1, ddVolTable.length, 3).setValues(ddVolTable);
  ddSheet.getRange(dr, 1, 1, 3).setFontWeight('bold').setBackground('#f3f3f3').setFontSize(9);
  ddSheet.getRange(dr + 1, 1, ddVolTable.length - 1, 3).setFontSize(9).setFontColor('#666666');
  ddSheet.getRange(dr + 1, 2, ddVolTable.length - 1, 2).setNumberFormat('#,##0');

  // Shopping funnel: normalized to Clicks = 100% (like Shopify's Sessions = 100%)
  // This keeps all bars visible on a linear scale
  const l7ClickBase = last7.clicks || 1;
  const p7ClickBase = prior7.clicks || 1;

  const ddFunnelData = [
    ['Stage', 'Last 7d', 'Prior 7d'],
    ['Sessions (Clicks)', 100, 100],
    ['Added to Cart', parseFloat((last7.atc / l7ClickBase * 100).toFixed(2)), parseFloat((prior7.atc / p7ClickBase * 100).toFixed(2))],
    ['Purchase', parseFloat((last7.purchases / l7ClickBase * 100).toFixed(2)), parseFloat((prior7.purchases / p7ClickBase * 100).toFixed(2))],
  ];
  ddSheet.getRange(dr, 5, ddFunnelData.length, 3).setValues(ddFunnelData);
  ddSheet.getRange(dr, 5, 1, 3).setFontWeight('bold').setBackground('#f3f3f3').setFontSize(9);
  ddSheet.getRange(dr + 1, 5, ddFunnelData.length - 1, 3).setFontSize(9).setFontColor('#666666');
  ddSheet.getRange(dr + 1, 6, ddFunnelData.length - 1, 2).setNumberFormat('0.00"%"');
  const ddFunnelRow = dr;

  // Stage-to-stage rates (cols 9-11)
  const ddRateData = [
    ['Rate', 'Last 7d', 'Prior 7d'],
    ['CTR', parseFloat(funnelL7.imprToClick.toFixed(2)), parseFloat(funnelP7.imprToClick.toFixed(2))],
    ['Click→ATC', parseFloat(funnelL7.clickToAtc.toFixed(2)), parseFloat(funnelP7.clickToAtc.toFixed(2))],
    ['ATC→Purch', parseFloat(funnelL7.atcToPurch.toFixed(2)), parseFloat(funnelP7.atcToPurch.toFixed(2))],
    ['CVR', parseFloat(funnelL7.clickToPurch.toFixed(2)), parseFloat(funnelP7.clickToPurch.toFixed(2))],
  ];
  ddSheet.getRange(dr, 9, ddRateData.length, 3).setValues(ddRateData);
  ddSheet.getRange(dr, 9, 1, 3).setFontWeight('bold').setBackground('#f3f3f3').setFontSize(9);
  ddSheet.getRange(dr + 1, 9, ddRateData.length - 1, 3).setFontSize(9).setFontColor('#666666');
  const ddRateRow = dr;

  dr += Math.max(ddFunnelData.length, ddRateData.length, ddVolTable.length) + 1;

  // Chart 1: Shopping Funnel — Clicks=100%, linear scale, all bars visible
  const ddFunnelChart = ddSheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(ddSheet.getRange(ddFunnelRow, 5, ddFunnelData.length, 3))
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setOption('title', 'Shopping Funnel (% of Sessions)')
    .setOption('titleTextStyle', { fontSize: 10, bold: true })
    .setOption('legend', { position: 'top' })
    .setOption('vAxis', { format: '0"%"', textStyle: { fontSize: 9 }, minValue: 0, maxValue: 105 })
    .setOption('series', { 0: { color: '#4285F4' }, 1: { color: '#B0BEC5' } })
    .setOption('bar', { groupWidth: '55%' })
    .setNumHeaders(1)
    .setPosition(dr, 1, 0, 0)
    .setOption('width', 520).setOption('height', 300)
    .build();
  ddSheet.insertChart(ddFunnelChart);

  // Chart 2: Stage-to-stage conversion rates
  const ddRateChart = ddSheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(ddSheet.getRange(ddRateRow, 9, ddRateData.length, 3))
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setOption('title', 'Stage-to-Stage Conversion Rates %')
    .setOption('titleTextStyle', { fontSize: 10, bold: true })
    .setOption('legend', { position: 'top' })
    .setOption('vAxis', { format: '0.0"%"', textStyle: { fontSize: 9 } })
    .setOption('series', { 0: { color: '#4285F4' }, 1: { color: '#B0BEC5' } })
    .setOption('bar', { groupWidth: '55%' })
    .setNumHeaders(1)
    .setPosition(dr, 7, 0, 0)
    .setOption('width', 500).setOption('height', 300)
    .build();
  ddSheet.insertChart(ddRateChart);

  dr += 16;

  // ── DAY OF WEEK ───────────────────────────────────────────────────────
  ddSheet.getRange(dr, 1).setValue('DAY OF WEEK PERFORMANCE').setFontWeight('bold').setFontSize(12);
  dr += 1;

  const dowNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  const dowBuckets = {};
  dowNames.forEach(d => dowBuckets[d] = []);
  sortedDates.forEach(dateStr => {
    const dayRows2 = dailyRows.filter(r2 => r2.date === dateStr);
    const parts = dateStr.split('-');
    const dateObj = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
    const dowName = dowNames[dateObj.getDay()];
    dayRows2.forEach(row2 => dowBuckets[dowName].push(row2));
  });

  const dowHeaders = ['Day', 'Spend', 'Purchases', 'Revenue', 'Impr', 'Clicks', 'ROAS', 'CPA', 'CVR', 'CPC', 'CPM'];
  ddSheet.getRange(dr, 1, 1, dowHeaders.length).setValues([dowHeaders]);
  ddSheet.getRange(dr, 1, 1, dowHeaders.length)
    .setFontWeight('bold').setBackground('#000000').setFontColor('#FFFFFF').setHorizontalAlignment('center');
  dr += 1;

  dowNames.forEach((dayName, i) => {
    const agg2 = aggregate(dowBuckets[dayName]);
    ddSheet.getRange(dr, 1, 1, dowHeaders.length).setValues([[
      dayName, agg2.spend, agg2.purchases, agg2.revenue,
      agg2.impressions, agg2.clicks, agg2.roas, agg2.cpa,
      agg2.cvr, agg2.cpc, agg2.cpm
    ]]);
    ddSheet.getRange(dr, 2).setNumberFormat('"$"#,##0');
    ddSheet.getRange(dr, 3).setNumberFormat('#,##0');
    ddSheet.getRange(dr, 4).setNumberFormat('"$"#,##0');
    ddSheet.getRange(dr, 5).setNumberFormat('#,##0');
    ddSheet.getRange(dr, 6).setNumberFormat('#,##0');
    ddSheet.getRange(dr, 7).setNumberFormat('0.00');
    ddSheet.getRange(dr, 8).setNumberFormat('"$"#,##0.00');
    ddSheet.getRange(dr, 9).setNumberFormat('0.00"%"');
    ddSheet.getRange(dr, 10).setNumberFormat('"$"#,##0.00');
    ddSheet.getRange(dr, 11).setNumberFormat('"$"#,##0.00');
    const roasCell = ddSheet.getRange(dr, 7);
    if (agg2.roas >= T.SCALE_ROAS) roasCell.setBackground('#e6f4ea').setFontColor('#137333').setFontWeight('bold');
    else if (agg2.roas >= T.TARGET_ROAS) roasCell.setBackground('#e6f4ea').setFontColor('#137333');
    else if (agg2.roas >= T.KILL_ROAS) roasCell.setBackground('#fef7e0').setFontColor('#E37400');
    else roasCell.setBackground('#fde8e8').setFontColor('#C5221F');
    ddSheet.getRange(dr, 1, 1, dowHeaders.length).setHorizontalAlignment('center').setVerticalAlignment('middle');
    ddSheet.getRange(dr, 1).setHorizontalAlignment('left').setFontWeight('bold');
    if (i % 2 === 0) ddSheet.getRange(dr, 1, 1, dowHeaders.length).setBackground('#fafafa');
    dr++;
  });

  const dowTotal = aggregate(dailyRows);
  ddSheet.getRange(dr, 1, 1, dowHeaders.length).setValues([[
    'Total', dowTotal.spend, dowTotal.purchases, dowTotal.revenue,
    dowTotal.impressions, dowTotal.clicks, dowTotal.roas, dowTotal.cpa,
    dowTotal.cvr, dowTotal.cpc, dowTotal.cpm
  ]]);
  ddSheet.getRange(dr, 1, 1, dowHeaders.length).setFontWeight('bold').setBackground('#e8eaf6').setHorizontalAlignment('center');
  ddSheet.getRange(dr, 1).setHorizontalAlignment('left');
  [2,3,4,5,6,7,8,9,10,11].forEach(c => {
    const fmts = {'2':'"$"#,##0','3':'#,##0','4':'"$"#,##0','5':'#,##0','6':'#,##0','7':'0.00','8':'"$"#,##0.00','9':'0.00"%"','10':'"$"#,##0.00','11':'"$"#,##0.00'};
    ddSheet.getRange(dr, c).setNumberFormat(fmts[String(c)]);
  });
  dr += 2;

  // ── WEEKLY PERFORMANCE CHARTS ─────────────────────────────────────────
  ddSheet.getRange(dr, 1).setValue('WEEKLY PERFORMANCE CHARTS').setFontWeight('bold').setFontSize(12);
  dr += 1;

  const weekBuckets = {};
  sortedDates.forEach(dateStr => {
    const parts = dateStr.split('-');
    const dateObj = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
    const day = dateObj.getDay();
    const mondayOffset = day === 0 ? -6 : 1 - day;
    const monday = new Date(dateObj);
    monday.setDate(dateObj.getDate() + mondayOffset);
    const weekKey = Utilities.formatDate(monday, 'UTC', 'MM/dd');
    if (!weekBuckets[weekKey]) weekBuckets[weekKey] = [];
    dailyRows.filter(row2 => row2.date === dateStr).forEach(row2 => weekBuckets[weekKey].push(row2));
  });

  const weekKeys = Object.keys(weekBuckets).sort();
  const weekAggs = weekKeys.map(wk => {
    const a = aggregate(weekBuckets[wk]);
    const weekDates = Object.keys(dailyTrend).filter(d => {
      const p = d.split('-');
      const dObj = new Date(parseInt(p[0]), parseInt(p[1]) - 1, parseInt(p[2]));
      const dy = dObj.getDay();
      const mo = dy === 0 ? -6 : 1 - dy;
      const mon = new Date(dObj);
      mon.setDate(dObj.getDate() + mo);
      return Utilities.formatDate(mon, 'UTC', 'MM/dd') === wk;
    });
    const weekAdsetRows = adsetRows.filter(ar => weekDates.includes(ar.date));
    const weekImpr = weekAdsetRows.reduce((s, ar) => s + (ar.impressions || 0), 0);
    const weekReach = weekAdsetRows.reduce((s, ar) => s + (ar.reach || 0), 0);
    const avgFreq = weekReach > 0 ? weekImpr / weekReach : 0;
    return { week: 'Wk ' + wk, ...a, frequency: avgFreq };
  });

  if (weekAggs.length >= 2) {
    const chartHeaders = [
      'Week', 'Spend', 'Revenue', 'ROAS', 'Target',
      'Reach', 'Frequency', 'Impressions', 'CTR', 'CVR',
      'CPA', 'CPM', 'CPC', 'Purchases'
    ];
    ddSheet.getRange(dr, 1, 1, chartHeaders.length).setValues([chartHeaders]);
    ddSheet.getRange(dr, 1, 1, chartHeaders.length)
      .setFontWeight('bold').setFontSize(8).setFontColor('#999999').setBackground('#f9f9f9');
    const chartHeaderRow = dr;
    dr += 1;

    const chartRows = weekAggs.map(w => [
      w.week, w.spend, w.revenue,
      parseFloat(w.roas.toFixed(2)), T.TARGET_ROAS,
      w.reach, parseFloat(w.frequency.toFixed(2)),
      w.impressions, parseFloat(w.ctr.toFixed(2)), parseFloat(w.cvr.toFixed(2)),
      parseFloat(w.cpa.toFixed(2)), parseFloat(w.cpm.toFixed(2)),
      parseFloat(w.cpc.toFixed(2)), w.purchases
    ]);
    ddSheet.getRange(dr, 1, chartRows.length, chartHeaders.length).setValues(chartRows);
    ddSheet.getRange(dr, 1, chartRows.length, chartHeaders.length).setFontSize(8).setFontColor('#999999');

    const nR = chartRows.length + 1;
    const colIdx = {};
    chartHeaders.forEach((h, i) => colIdx[h] = i + 1);
    function ddColRange(name) { return ddSheet.getRange(chartHeaderRow, colIdx[name], nR, 1); }
    function ddWeekRange() { return ddSheet.getRange(chartHeaderRow, colIdx['Week'], nR, 1); }

    let chartRow = chartHeaderRow;
    let chartCol = 1;
    const CW = 500, CH = 270, CS = 14;

    function ddChart(type, title, ranges, series, vFmt, v2Fmt) {
      const b = ddSheet.newChart().setChartType(type)
        .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS);
      b.addRange(ddWeekRange());
      ranges.forEach(n => b.addRange(ddColRange(n)));
      b.setNumHeaders(1)
        .setOption('title', title)
        .setOption('titleTextStyle', { fontSize: 10, bold: true })
        .setOption('legend', { position: 'top', textStyle: { fontSize: 9 } })
        .setOption('hAxis', { textStyle: { fontSize: 8 } })
        .setOption('bar', { groupWidth: '65%' });
      if (v2Fmt) b.setOption('vAxes', { 0: { format: vFmt, textStyle: { fontSize: 9 } }, 1: { format: v2Fmt, textStyle: { fontSize: 9 } } });
      else b.setOption('vAxis', { format: vFmt, textStyle: { fontSize: 9 } });
      if (series) b.setOption('series', series);
      b.setPosition(chartRow, chartCol, 0, 0).setOption('width', CW).setOption('height', CH);
      ddSheet.insertChart(b.build());
      if (chartCol === 1) { chartCol = 8; } else { chartCol = 1; chartRow += CS; }
    }

    ddChart(Charts.ChartType.COLUMN, 'Weekly Spend vs Revenue', ['Spend', 'Revenue'],
      { 0: { color: '#4285F4' }, 1: { color: '#34A853' } }, '$#,##0');
    ddChart(Charts.ChartType.LINE, 'Weekly ROAS vs ' + T.TARGET_ROAS + 'x Target', ['ROAS', 'Target'],
      { 0: { color: '#4285F4', lineWidth: 2, pointSize: 5 }, 1: { color: '#EA4335', lineWidth: 2, lineDashStyle: [4, 4], pointSize: 0 } }, '0.0');
    ddChart(Charts.ChartType.COMBO, 'Reach & Frequency', ['Reach', 'Frequency'],
      { 0: { type: 'bars', color: '#4285F4', targetAxisIndex: 0 }, 1: { type: 'line', color: '#34A853', lineWidth: 2, pointSize: 5, targetAxisIndex: 1 } }, '#,##0', '0.0');
    ddChart(Charts.ChartType.COMBO, 'Purchases & CPA', ['Purchases', 'CPA'],
      { 0: { type: 'bars', color: '#4285F4', targetAxisIndex: 0 }, 1: { type: 'line', color: '#EA4335', lineWidth: 2, pointSize: 5, targetAxisIndex: 1 } }, '#,##0', '$#,##0');
    ddChart(Charts.ChartType.COMBO, 'CTR & CPM', ['CTR', 'CPM'],
      { 0: { type: 'line', color: '#4285F4', lineWidth: 2, pointSize: 5, targetAxisIndex: 0 }, 1: { type: 'line', color: '#E37400', lineWidth: 2, pointSize: 5, targetAxisIndex: 1 } }, '0.0"%"', '$#,##0');
    ddChart(Charts.ChartType.COMBO, 'CVR & CPC', ['CVR', 'CPC'],
      { 0: { type: 'line', color: '#137333', lineWidth: 2, pointSize: 5, targetAxisIndex: 0 }, 1: { type: 'line', color: '#E37400', lineWidth: 2, pointSize: 5, targetAxisIndex: 1 } }, '0.0"%"', '$0.00');
  }

  for (let c = 1; c <= 14; c++) ddSheet.setColumnWidth(c, 100);
  ddSheet.setColumnWidth(1, 130);

  // ══════════════════════════════════════════════════════════════════════
  // 📈 GOOGLE DEEP DIVES TAB — PMax Asset Groups, Products, Search Terms
  // ══════════════════════════════════════════════════════════════════════
  if (hasGoogleAds) {
    let gSheet = ss.getSheetByName(CONFIG.TABS.GOOGLE_DEEP);
    if (!gSheet) gSheet = ss.insertSheet(CONFIG.TABS.GOOGLE_DEEP);
    gSheet.clear();
    const gOldCharts = gSheet.getCharts();
    gOldCharts.forEach(c => gSheet.removeChart(c));

    let gr = 1;
    gSheet.getRange(gr, 1).setValue('GOOGLE ADS DEEP DIVE').setFontWeight('bold').setFontSize(14);
    gr += 1;
    gSheet.getRange(gr, 1).setValue(`${dateRange.since} → ${dateRange.until}`).setFontSize(9).setFontColor('#666666');
    gr += 2;

    // ── Google Ads KPI Summary ──────────────────────────────────────────
    gSheet.getRange(gr, 1).setValue('GOOGLE ADS METRICS (WoW)').setFontWeight('bold').setFontSize(12);
    gr += 1;

    const gKpiHeaders = ['Metric', 'Last 7d', 'Prior 7d', 'WoW Δ'];
    gSheet.getRange(gr, 1, 1, 4).setValues([gKpiHeaders]);
    gSheet.getRange(gr, 1, 1, 4).setFontWeight('bold').setBackground('#000000').setFontColor('#FFFFFF').setHorizontalAlignment('center');
    gr += 1;

    const gMetrics = [
      ['Spend', gAdsL7.spend, gAdsP7.spend, trendPct(gAdsL7.spend, gAdsP7.spend, false), '"$"#,##0'],
      ['Revenue', gAdsL7.revenue, gAdsP7.revenue, trendPct(gAdsL7.revenue, gAdsP7.revenue, false), '"$"#,##0'],
      ['Conversions', gAdsL7.purchases, gAdsP7.purchases, trendPct(gAdsL7.purchases, gAdsP7.purchases, false), '#,##0.0'],
      ['ROAS', gAdsL7.roas, gAdsP7.roas, trendPct(gAdsL7.roas, gAdsP7.roas, false), '0.00"x"'],
      ['CPA', gAdsL7.cpa, gAdsP7.cpa, trendPct(gAdsL7.cpa, gAdsP7.cpa, true), '"$"#,##0.00'],
      ['CPC', gAdsL7.cpc, gAdsP7.cpc, trendPct(gAdsL7.cpc, gAdsP7.cpc, true), '"$"#,##0.00'],
      ['CTR', gAdsL7.ctr, gAdsP7.ctr, trendPct(gAdsL7.ctr, gAdsP7.ctr, false), '0.00"%"'],
      ['CVR', gAdsL7.cvr, gAdsP7.cvr, trendPct(gAdsL7.cvr, gAdsP7.cvr, false), '0.00"%"'],
    ];

    gMetrics.forEach((m, i) => {
      gSheet.getRange(gr, 1, 1, 4).setValues([[m[0], m[1], m[2], m[3]]]);
      gSheet.getRange(gr, 1).setFontWeight('bold').setHorizontalAlignment('left');
      gSheet.getRange(gr, 2).setNumberFormat(m[4]);
      gSheet.getRange(gr, 3).setNumberFormat(m[4]);
      gSheet.getRange(gr, 1, 1, 4).setHorizontalAlignment('center');
      gSheet.getRange(gr, 1).setHorizontalAlignment('left');
      if (i % 2 === 0) gSheet.getRange(gr, 1, 1, 4).setBackground('#fafafa');

      // Color WoW delta
      const deltaStr = m[3];
      if (typeof deltaStr === 'string') {
        const isUp = deltaStr.includes('▲');
        const pctMatch = deltaStr.match(/([\d.]+)%/);
        const pctVal = pctMatch ? parseFloat(pctMatch[1]) : 0;
        const dCell = gSheet.getRange(gr, 4);
        if (pctVal < 5) {
          dCell.setFontColor('#888888');
        } else if (isUp) {
          dCell.setFontColor('#137333').setFontWeight('bold');
          if (pctVal >= 15) dCell.setBackground('#e6f4ea');
        } else {
          dCell.setFontColor('#C5221F').setFontWeight('bold');
          if (pctVal >= 15) dCell.setBackground('#fde8e8');
        }
      }

      // Color ROAS row
      if (m[0] === 'ROAS') {
        [2, 3].forEach(col => {
          const val = m[col - 1]; // m[1] = L7, m[2] = P7
          const cell = gSheet.getRange(gr, col);
          if (val >= T.SCALE_ROAS) cell.setFontColor('#137333').setFontWeight('bold');
          else if (val >= T.TARGET_ROAS) cell.setFontColor('#137333');
          else if (val >= T.KILL_ROAS) cell.setFontColor('#E37400');
          else cell.setFontColor('#C5221F');
        });
      }
      gr++;
    });
    gr += 1;

    // ── Helper: read tab into rows ────────────────────────────────────
    function readGTab(tabName) {
      const s = ss.getSheetByName(tabName);
      if (!s || s.getLastRow() < 2) return [];
      const hdr = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
      const data = s.getRange(2, 1, s.getLastRow() - 1, s.getLastColumn()).getValues();
      return data.map(row => {
        const obj = {};
        hdr.forEach((h, i) => obj[h] = row[i]);
        if (obj.date instanceof Date) obj.date = Utilities.formatDate(obj.date, 'UTC', 'yyyy-MM-dd');
        else obj.date = String(obj.date).substring(0, 10);
        return obj;
      });
    }

    const assetRows = readGTab('google_ads_assets');
    const productRows = readGTab('google_ads_products');
    const searchRows = readGTab('google_ads_search');

    // Build Google Ads own date windows (independent of Meta dates)
    const allGDates = [...new Set([
      ...assetRows.map(r2 => r2.date),
      ...productRows.map(r2 => r2.date),
      ...searchRows.map(r2 => r2.date),
    ])].sort();
    const gCompletedDates = allGDates.filter(d => d < todayStr);
    const gLast7Dates = gCompletedDates.slice(-7);
    const gPrior7Dates = gCompletedDates.slice(-14, -7);

    // ── ASSET GROUPS ──────────────────────────────────────────────────
    if (assetRows.length > 0) {
      gSheet.getRange(gr, 1).setValue('ASSET GROUPS (WoW)').setFontWeight('bold').setFontSize(12);
      gr += 1;

      const agMap = {};
      assetRows.forEach(r2 => {
        const key = r2.asset_group || r2.campaign;
        if (!agMap[key]) agMap[key] = { l7: [], p7: [] };
        if (gLast7Dates.includes(r2.date)) agMap[key].l7.push(r2);
        if (gPrior7Dates.includes(r2.date)) agMap[key].p7.push(r2);
      });

      function gAgg(rows) {
        const a = { spend: 0, revenue: 0, conversions: 0, impressions: 0, clicks: 0 };
        rows.forEach(r2 => {
          a.spend += parseFloat(r2.spend) || 0; a.revenue += parseFloat(r2.conversion_value) || 0;
          a.conversions += parseFloat(r2.conversions) || 0; a.impressions += parseInt(r2.impressions) || 0;
          a.clicks += parseInt(r2.clicks) || 0;
        });
        a.roas = a.spend > 0 ? a.revenue / a.spend : 0;
        a.cpa = a.conversions > 0 ? a.spend / a.conversions : 0;
        a.ctr = a.impressions > 0 ? (a.clicks / a.impressions * 100) : 0;
        a.cvr = a.clicks > 0 ? (a.conversions / a.clicks * 100) : 0;
        return a;
      }

      const activeAGs = Object.entries(agMap)
        .map(([name, data]) => ({ name, l: gAgg(data.l7), p: gAgg(data.p7) }))
        .filter(a => a.l.spend > 0 || a.p.spend > 0)
        .sort((a, b) => b.l.spend - a.l.spend);

      if (activeAGs.length > 0) {
        const agHeaders = ['Asset Group', 'L7 Spend', 'L7 ROAS', 'L7 CPA', 'L7 Conv', 'L7 CTR', 'L7 CVR', 'ROAS Δ', 'CPA Δ'];
        gSheet.getRange(gr, 1, 1, agHeaders.length).setValues([agHeaders]);
        gSheet.getRange(gr, 1, 1, agHeaders.length)
          .setFontWeight('bold').setBackground('#000000').setFontColor('#FFFFFF').setHorizontalAlignment('center').setFontSize(9);
        gr += 1;

        activeAGs.forEach((ag, i) => {
          gSheet.getRange(gr, 1, 1, agHeaders.length).setValues([[
            ag.name, ag.l.spend, ag.l.roas, ag.l.cpa, ag.l.conversions, ag.l.ctr, ag.l.cvr,
            trendPct(ag.l.roas, ag.p.roas, false), trendPct(ag.l.cpa, ag.p.cpa, true)
          ]]);
          gSheet.getRange(gr, 2).setNumberFormat('"$"#,##0');
          gSheet.getRange(gr, 3).setNumberFormat('0.00"x"');
          gSheet.getRange(gr, 4).setNumberFormat('"$"#,##0.00');
          gSheet.getRange(gr, 5).setNumberFormat('#,##0.0');
          gSheet.getRange(gr, 6).setNumberFormat('0.00"%"');
          gSheet.getRange(gr, 7).setNumberFormat('0.00"%"');
          gSheet.getRange(gr, 1, 1, agHeaders.length).setHorizontalAlignment('center').setFontSize(9);
          gSheet.getRange(gr, 1).setHorizontalAlignment('left').setFontWeight('bold');

          const roasCell = gSheet.getRange(gr, 3);
          if (ag.l.roas >= T.SCALE_ROAS) roasCell.setBackground('#e6f4ea').setFontColor('#137333');
          else if (ag.l.roas >= T.TARGET_ROAS) roasCell.setFontColor('#137333');
          else if (ag.l.roas >= T.KILL_ROAS) roasCell.setFontColor('#E37400');
          else if (ag.l.spend > 0) roasCell.setFontColor('#C5221F');

          if (i % 2 === 0) gSheet.getRange(gr, 1, 1, agHeaders.length).setBackground('#fafafa');
          gr++;
        });
        gr += 1;
      }
    }

    // ── PRODUCTS ─────────────────────────────────────────────────────
    if (productRows.length > 0) {
      const prodDateLabel = gLast7Dates.length > 0
        ? `${gLast7Dates[0]} → ${gLast7Dates[gLast7Dates.length - 1]}`
        : 'Last 7d';

      gSheet.getRange(gr, 1).setValue(`PRODUCT PERFORMANCE (${prodDateLabel})`).setFontWeight('bold').setFontSize(12);
      gr += 1;

      const prodMap = {};
      productRows.forEach(r2 => {
        if (!gLast7Dates.includes(r2.date)) return;
        const key = r2.product_title || r2.product_id || 'Unknown';
        if (!prodMap[key]) prodMap[key] = { spend: 0, revenue: 0, conversions: 0, clicks: 0, impressions: 0 };
        prodMap[key].spend += parseFloat(r2.spend) || 0;
        prodMap[key].revenue += parseFloat(r2.conversion_value) || 0;
        prodMap[key].conversions += parseFloat(r2.conversions) || 0;
        prodMap[key].clicks += parseInt(r2.clicks) || 0;
        prodMap[key].impressions += parseInt(r2.impressions) || 0;
      });

      const products = Object.entries(prodMap)
        .map(([name, d]) => ({ name, ...d, roas: d.spend > 0 ? d.revenue / d.spend : 0, cpc: d.clicks > 0 ? d.spend / d.clicks : 0 }))
        .filter(p => p.spend > 0).sort((a, b) => b.spend - a.spend);

      const topProducts = products.slice(0, 15);
      const totalProdSpend = products.reduce((s, p) => s + p.spend, 0);

      if (topProducts.length > 0) {
        const prodHeaders = ['Product', 'Spend', '% of Total', 'ROAS', 'Conv', 'Revenue', 'CPC'];
        gSheet.getRange(gr, 1, 1, prodHeaders.length).setValues([prodHeaders]);
        gSheet.getRange(gr, 1, 1, prodHeaders.length)
          .setFontWeight('bold').setBackground('#000000').setFontColor('#FFFFFF').setHorizontalAlignment('center').setFontSize(9);
        gr += 1;

        topProducts.forEach((p, i) => {
          gSheet.getRange(gr, 1, 1, prodHeaders.length).setValues([[
            p.name, p.spend, totalProdSpend > 0 ? p.spend/totalProdSpend*100 : 0,
            p.roas, p.conversions, p.revenue, p.cpc
          ]]);
          gSheet.getRange(gr, 2).setNumberFormat('"$"#,##0.00');
          gSheet.getRange(gr, 3).setNumberFormat('0.0"%"');
          gSheet.getRange(gr, 4).setNumberFormat('0.00"x"');
          gSheet.getRange(gr, 5).setNumberFormat('#,##0.0');
          gSheet.getRange(gr, 6).setNumberFormat('"$"#,##0.00');
          gSheet.getRange(gr, 7).setNumberFormat('"$"#,##0.00');
          gSheet.getRange(gr, 1, 1, prodHeaders.length).setHorizontalAlignment('center').setFontSize(9);
          gSheet.getRange(gr, 1).setHorizontalAlignment('left');

          const rCell = gSheet.getRange(gr, 4);
          if (p.roas >= T.SCALE_ROAS) rCell.setFontColor('#137333').setFontWeight('bold');
          else if (p.roas >= T.TARGET_ROAS) rCell.setFontColor('#137333');
          else if (p.roas >= T.KILL_ROAS) rCell.setFontColor('#E37400');
          else rCell.setFontColor('#C5221F');

          if (i % 2 === 0) gSheet.getRange(gr, 1, 1, prodHeaders.length).setBackground('#fafafa');
          gr++;
        });
        gr += 1;
      }
    }

    // ── SEARCH TERMS — Brand vs Non-Brand ───────────────────────────
    if (searchRows.length > 0) {
      const searchDateLabel = gLast7Dates.length > 0
        ? `${gLast7Dates[0]} → ${gLast7Dates[gLast7Dates.length - 1]}`
        : 'Last 7d';

      gSheet.getRange(gr, 1).setValue(`SEARCH TERMS (${searchDateLabel})`).setFontWeight('bold').setFontSize(12);
      gr += 1;

      const brandKeywords = CONFIG.BRAND_KEYWORDS || [];
      const l7Search = searchRows.filter(r2 => gLast7Dates.includes(r2.date));

      let brandSpend = 0, brandRev = 0, brandConv = 0, brandClicks = 0;
      let nbSpend = 0, nbRev = 0, nbConv = 0, nbClicks = 0;
      const termMap = {};

      l7Search.forEach(r2 => {
        const term = (r2.search_term || '').toLowerCase().trim();
        const spend = parseFloat(r2.spend) || 0;
        const rev = parseFloat(r2.conversion_value) || 0;
        const conv = parseFloat(r2.conversions) || 0;
        const clicks = parseInt(r2.clicks) || 0;
        const isBrand = brandKeywords.some(bk => term.includes(bk));
        if (isBrand) { brandSpend += spend; brandRev += rev; brandConv += conv; brandClicks += clicks; }
        else { nbSpend += spend; nbRev += rev; nbConv += conv; nbClicks += clicks; }
        if (!termMap[term]) termMap[term] = { spend: 0, revenue: 0, conversions: 0, clicks: 0, impressions: 0 };
        termMap[term].spend += spend; termMap[term].revenue += rev;
        termMap[term].conversions += conv; termMap[term].clicks += clicks;
        termMap[term].impressions += parseInt(r2.impressions) || 0;
      });

      // Brand vs Non-Brand
      const totalSearchSpend = brandSpend + nbSpend;
      const bHeaders = ['Segment', 'Spend', '% of Spend', 'Revenue', 'ROAS', 'Conv', 'CPC'];
      gSheet.getRange(gr, 1, 1, bHeaders.length).setValues([bHeaders]);
      gSheet.getRange(gr, 1, 1, bHeaders.length)
        .setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#FFFFFF').setHorizontalAlignment('center').setFontSize(9);
      gr += 1;

      const bRows = [
        ['Brand', brandSpend, totalSearchSpend > 0 ? brandSpend/totalSearchSpend*100 : 0, brandRev, brandSpend > 0 ? brandRev/brandSpend : 0, brandConv, brandClicks > 0 ? brandSpend/brandClicks : 0],
        ['Non-Brand', nbSpend, totalSearchSpend > 0 ? nbSpend/totalSearchSpend*100 : 0, nbRev, nbSpend > 0 ? nbRev/nbSpend : 0, nbConv, nbClicks > 0 ? nbSpend/nbClicks : 0],
        ['Total', totalSearchSpend, 100, brandRev+nbRev, totalSearchSpend > 0 ? (brandRev+nbRev)/totalSearchSpend : 0, brandConv+nbConv, (brandClicks+nbClicks) > 0 ? totalSearchSpend/(brandClicks+nbClicks) : 0],
      ];
      bRows.forEach((br, i) => {
        gSheet.getRange(gr, 1, 1, bHeaders.length).setValues([br]);
        gSheet.getRange(gr, 2).setNumberFormat('"$"#,##0.00');
        gSheet.getRange(gr, 3).setNumberFormat('0.0"%"');
        gSheet.getRange(gr, 4).setNumberFormat('"$"#,##0.00');
        gSheet.getRange(gr, 5).setNumberFormat('0.00"x"');
        gSheet.getRange(gr, 6).setNumberFormat('#,##0.0');
        gSheet.getRange(gr, 7).setNumberFormat('"$"#,##0.00');
        gSheet.getRange(gr, 1, 1, bHeaders.length).setHorizontalAlignment('center').setFontSize(9);
        gSheet.getRange(gr, 1).setHorizontalAlignment('left').setFontWeight('bold');
        if (i === 2) gSheet.getRange(gr, 1, 1, bHeaders.length).setFontWeight('bold').setBackground('#e8eaf6');
        else if (i % 2 === 0) gSheet.getRange(gr, 1, 1, bHeaders.length).setBackground('#fafafa');
        gr++;
      });

      const brandPct = totalSearchSpend > 0 ? (brandSpend / totalSearchSpend * 100) : 0;
      if (brandPct > 50) {
        gSheet.getRange(gr, 1, 1, 7).merge();
        gSheet.getRange(gr, 1).setValue(`⚠️ ${brandPct.toFixed(0)}% of PMax search spend is brand terms. PMax is mostly capturing existing demand, not prospecting.`)
          .setFontColor('#E37400').setFontWeight('bold').setFontSize(9).setWrap(true);
        gr += 1;
      }
      gr += 1;

      // Top terms
      gSheet.getRange(gr, 1).setValue(`TOP SEARCH TERMS (${searchDateLabel})`).setFontWeight('bold').setFontSize(11);
      gr += 1;

      const topTerms = Object.entries(termMap)
        .map(([term, d]) => ({
          term, ...d, roas: d.spend > 0 ? d.revenue / d.spend : 0,
          cpc: d.clicks > 0 ? d.spend / d.clicks : 0,
          isBrand: brandKeywords.some(bk => term.includes(bk)),
        }))
        .filter(t => t.spend > 0.5).sort((a, b) => b.spend - a.spend).slice(0, 20);

      if (topTerms.length > 0) {
        const stHeaders = ['Search Term', 'Type', 'Spend', 'Clicks', 'Conv', 'Revenue', 'ROAS', 'CPC'];
        gSheet.getRange(gr, 1, 1, stHeaders.length).setValues([stHeaders]);
        gSheet.getRange(gr, 1, 1, stHeaders.length)
          .setFontWeight('bold').setBackground('#000000').setFontColor('#FFFFFF').setHorizontalAlignment('center').setFontSize(9);
        gr += 1;

        topTerms.forEach((t, i) => {
          gSheet.getRange(gr, 1, 1, stHeaders.length).setValues([[
            t.term, t.isBrand ? 'Brand' : 'Non-Brand', t.spend, t.clicks, t.conversions, t.revenue, t.roas, t.cpc
          ]]);
          gSheet.getRange(gr, 3).setNumberFormat('"$"#,##0.00');
          gSheet.getRange(gr, 4).setNumberFormat('#,##0');
          gSheet.getRange(gr, 5).setNumberFormat('#,##0.0');
          gSheet.getRange(gr, 6).setNumberFormat('"$"#,##0.00');
          gSheet.getRange(gr, 7).setNumberFormat('0.00"x"');
          gSheet.getRange(gr, 8).setNumberFormat('"$"#,##0.00');
          gSheet.getRange(gr, 1, 1, stHeaders.length).setHorizontalAlignment('center').setFontSize(9);
          gSheet.getRange(gr, 1).setHorizontalAlignment('left');
          if (!t.isBrand) gSheet.getRange(gr, 2).setFontColor('#4285F4').setFontWeight('bold');
          else gSheet.getRange(gr, 2).setFontColor('#666666').setFontStyle('italic');

          const rCell = gSheet.getRange(gr, 7);
          if (t.roas >= T.SCALE_ROAS) rCell.setFontColor('#137333').setFontWeight('bold');
          else if (t.roas >= T.TARGET_ROAS) rCell.setFontColor('#137333');
          else if (t.roas >= T.KILL_ROAS) rCell.setFontColor('#E37400');
          else if (t.spend > 1) rCell.setFontColor('#C5221F');

          if (i % 2 === 0) gSheet.getRange(gr, 1, 1, stHeaders.length).setBackground('#fafafa');
          gr++;
        });
      }
    }

    // Column widths
    gSheet.setColumnWidth(1, 200);
    for (let c = 2; c <= 9; c++) gSheet.setColumnWidth(c, 100);
  }

  // ══════════════════════════════════════════════════════════════════════
  // TAB ORDERING + CLEANUP
  // ══════════════════════════════════════════════════════════════════════
  // Delete legacy tabs
  const legacyNames = ['📈 deep_dives', '⚡ actions'];
  legacyNames.forEach(name => {
    const old = ss.getSheetByName(name);
    if (old) { try { ss.deleteSheet(old); } catch(e) {} }
  });

  // Order: dashboard(1), meta_deep(2), google_deep(3), meta_creative(4), meta_age(5), rest
  const tabOrder = [
    CONFIG.TABS.DASHBOARD,
    CONFIG.TABS.META_DEEP,
    CONFIG.TABS.GOOGLE_DEEP,
    CONFIG.TABS.META_CREATIVE,
    CONFIG.TABS.META_AGE,
  ];

  tabOrder.forEach((tabName, idx) => {
    const tab = ss.getSheetByName(tabName);
    if (tab) {
      ss.setActiveSheet(tab);
      ss.moveActiveSheet(idx + 1);
    }
  });

  // Re-activate dashboard
  const dashTab = ss.getSheetByName(CONFIG.TABS.DASHBOARD);
  if (dashTab) ss.setActiveSheet(dashTab);

  log.push('dashboard + meta_deep_dives + google_deep_dives: built');
}


// =============================================================================
// ORIGINAL SYNC FUNCTIONS (with Status column added to Creatives)
// =============================================================================

function syncMetaDaily(log) {
  const fields = [
    'campaign_name', 'spend', 'impressions', 'reach', 'clicks',
    'actions', 'action_values', 'cpm', 'cpc', 'ctr', 'frequency'
  ];

  const params = {
    level: 'campaign',
    time_increment: 1,
    breakdowns: '',
    fields: fields.join(','),
  };

  const data = fetchInsights(params, log);

  const headers = [
    'date', 'campaign', 'spend', 'impressions', 'reach', 'clicks',
    'purchases', 'purchase_value', 'atc', 'cpm', 'cpc', 'ctr', 'frequency',
    'cpa', 'roas'
  ];

  const rows = data.map(row => {
    const purchases = extractAction(row.actions, 'omni_purchase') || 0;
    const purchaseValue = extractAction(row.action_values, 'omni_purchase') || 0;
    const spend = parseFloat(row.spend) || 0;
    const atc = extractAction(row.actions, 'omni_add_to_cart') || 0;

    return [
      row.date_start,
      row.campaign_name,
      spend,
      parseInt(row.impressions) || 0,
      parseInt(row.reach) || 0,
      parseInt(row.clicks) || 0,
      purchases,
      purchaseValue,
      atc,
      parseFloat(row.cpm) || 0,
      parseFloat(row.cpc) || 0,
      parseFloat(row.ctr) || 0,
      parseFloat(row.frequency) || 0,
      purchases > 0 ? (spend / purchases).toFixed(2) : '',
      spend > 0 ? (purchaseValue / spend).toFixed(2) : ''
    ];
  });

  writeToSheet(CONFIG.TABS.META_DAILY, headers, rows);
  applyConditionalFormatting(CONFIG.TABS.META_DAILY, headers, rows.length);
  log.push(`meta_daily: ${rows.length} rows synced`);
}

function syncMetaAdsetDaily(log) {
  const fields = [
    'campaign_name', 'adset_name', 'adset_id', 'spend', 'impressions', 'reach',
    'clicks', 'actions', 'action_values', 'frequency'
  ];

  const params = {
    level: 'adset',
    time_increment: 1,
    breakdowns: '',
    fields: fields.join(','),
  };

  const data = fetchInsights(params, log);

  const headers = [
    'date', 'campaign', 'adset', 'spend', 'impressions', 'reach', 'clicks',
    'purchases', 'purchase_value', 'atc', 'frequency', 'cpa', 'roas'
  ];

  const rows = data.map(row => {
    const purchases = extractAction(row.actions, 'omni_purchase') || 0;
    const purchaseValue = extractAction(row.action_values, 'omni_purchase') || 0;
    const spend = parseFloat(row.spend) || 0;
    const atc = extractAction(row.actions, 'omni_add_to_cart') || 0;

    return [
      row.date_start,
      row.campaign_name || '',
      row.adset_name || row.adset_id || 'unknown_adset',
      spend,
      parseInt(row.impressions) || 0,
      parseInt(row.reach) || 0,
      parseInt(row.clicks) || 0,
      purchases,
      purchaseValue,
      atc,
      parseFloat(row.frequency) || 0,
      purchases > 0 ? (spend / purchases).toFixed(2) : '',
      spend > 0 ? (purchaseValue / spend).toFixed(2) : ''
    ];
  });

  writeToSheet(CONFIG.TABS.META_ADSET, headers, rows);
  applyConditionalFormatting(CONFIG.TABS.META_ADSET, headers, rows.length);
  log.push(`meta_adset_daily: ${rows.length} rows synced`);
}

function syncMetaCreative(log) {
  const T = CONFIG.THRESHOLDS;
  const { creativeMap: thumbnailMap, statusMap } = fetchAdThumbnails(log);

  const fields = [
    'ad_id', 'ad_name', 'spend', 'impressions', 'reach', 'clicks',
    'actions', 'action_values', 'frequency',
    'video_thruplay_watched_actions', 'video_p25_watched_actions',
    'video_p50_watched_actions', 'video_p75_watched_actions', 'video_p100_watched_actions'
  ];

  // ── Two API calls: full period + last 7 days ──────────────────────────
  const fullParams = { level: 'ad', time_increment: 'all_days', breakdowns: '', fields: fields.join(',') };
  const fullData = fetchInsights(fullParams, log);

  // Last 7 days — custom date override
  const now = new Date();
  const since7 = new Date(now);
  since7.setDate(since7.getDate() - 7);
  const since7Str = Utilities.formatDate(since7, 'UTC', 'yyyy-MM-dd');
  const untilStr = Utilities.formatDate(now, 'UTC', 'yyyy-MM-dd');

  const recent7Url = `https://graph.facebook.com/${CONFIG.API_VERSION}/${CONFIG.AD_ACCOUNT_ID}/insights` +
    `?access_token=${CONFIG.ACCESS_TOKEN}` +
    `&time_range=${encodeURIComponent(`{"since":"${since7Str}","until":"${untilStr}"}`)}` +
    `&level=ad&time_increment=all_days&fields=${fields.join(',')}&limit=500`;

  let recent7Data = [];
  let r7Url = recent7Url;
  let r7Pages = 0;
  while (r7Url && r7Pages < 20) {
    const resp = UrlFetchApp.fetch(r7Url, { muteHttpExceptions: true });
    if (resp.getResponseCode() === 200) {
      const json = JSON.parse(resp.getContentText());
      if (json.data) recent7Data = recent7Data.concat(json.data);
      r7Url = json.paging && json.paging.next ? json.paging.next : null;
    } else {
      log.push(`Recent7d creative API error: ${resp.getResponseCode()}`);
      break;
    }
    r7Pages++;
    if (r7Url) Utilities.sleep(500);
  }

  // ── Rollup helper ─────────────────────────────────────────────────────
  function rollupData(rawData) {
    const rollup = {};
    rawData.forEach(row => {
      const adId = row.ad_id;
      const adName = row.ad_name;
      const spend = parseFloat(row.spend) || 0;
      const impressions = parseInt(row.impressions) || 0;
      const clicks = parseInt(row.clicks) || 0;
      const purchases = extractAction(row.actions, 'omni_purchase') || 0;
      const purchaseValue = extractAction(row.action_values, 'omni_purchase') || 0;
      const atc = extractAction(row.actions, 'omni_add_to_cart') || 0;
      const p25 = extractVideoAction(row.video_p25_watched_actions) || 0;
      const p75 = extractVideoAction(row.video_p75_watched_actions) || 0;
      const imgUrl = thumbnailMap[adId] || '';

      if (!rollup[adId]) {
        rollup[adId] = { ad_id: adId, ad_name: adName, img_url: '', spend: 0, impressions: 0, clicks: 0,
          purchases: 0, revenue: 0, atc: 0, p25: 0, p75: 0 };
      }
      const c = rollup[adId];
      c.spend += spend; c.impressions += impressions; c.clicks += clicks;
      c.purchases += purchases; c.revenue += purchaseValue; c.atc += atc;
      c.p25 += p25; c.p75 += p75;
      if (!c.img_url && imgUrl) c.img_url = imgUrl;
    });
    return rollup;
  }

  const fullRollup = rollupData(fullData);
  const recentRollup = rollupData(recent7Data);

  // ── Compute derived metrics for a rollup entry ────────────────────────
  function computeMetrics(d) {
    if (!d) return { spend: 0, roas: 0, cpa: 0, cpc: 0, ctr: 0, thumbstop: 0, holdRate: 0, purchases: 0, revenue: 0, impressions: 0, clicks: 0, atc: 0 };
    return {
      spend: d.spend,
      purchases: d.purchases,
      revenue: d.revenue,
      impressions: d.impressions,
      clicks: d.clicks,
      atc: d.atc,
      roas: d.spend > 0 ? d.revenue / d.spend : 0,
      cpa: d.purchases > 0 ? d.spend / d.purchases : 0,
      cpc: d.clicks > 0 ? d.spend / d.clicks : 0,
      ctr: d.impressions > 0 ? (d.clicks / d.impressions * 100) : 0,
      thumbstop: d.impressions > 0 ? (d.p25 / d.impressions * 100) : 0,
      holdRate: d.p25 > 0 ? (d.p75 / d.p25 * 100) : 0,
    };
  }

  // ── Build combined creative objects (keyed by ad_id) ──────────────────
  const allCreatives = Object.keys(fullRollup).map(adId => {
    const full = computeMetrics(fullRollup[adId]);
    const recent = computeMetrics(recentRollup[adId]);
    const adName = fullRollup[adId].ad_name;
    const imgUrl = fullRollup[adId].img_url || '';

    // How much spend happened BEFORE the last 7 days?
    const priorSpend = Math.max(0, full.spend - recent.spend);
    const isNew = priorSpend < T.MIN_SPEND_FOR_JUDGMENT; // barely any spend before this week

    // ── Phase: where does this creative sit in the workflow? ───────────
    // Check effective_status by ad_id — catches adset-paused, campaign-paused, etc.
    const effectiveStatus = statusMap[adId] || '';
    const isEffectivelyPaused = effectiveStatus && effectiveStatus !== 'ACTIVE';

    // 1. Paused/off or no recent spend → skip
    // 2. New + still ramping (< threshold in 7d) → "testing"
    // 3. New + crossed threshold (>= threshold in 7d) → "test_results" (decision time)
    // 4. Has prior history + recent spend → "active" (proven creative still running)
    let phase = '';
    let status = '';

    if (recent.spend === 0 || isEffectivelyPaused) {
      phase = 'inactive'; // Will be filtered out
      status = '⏸️ Inactive';
    } else if (isNew && recent.spend < T.TESTING_SPEND_CAP) {
      // Check if creative has been running long enough to make a call even at low spend
      // If Meta gave it 7+ days and barely spent on it, that's a signal — Meta doesn't believe in it
      const daysActive = full.spend > 0 && recent.spend > 0 ? 7 : 0; // approximation: if both windows have data, it's been running 7+ days
      const lowSpendLongRun = recent.spend >= T.MIN_SPEND_FOR_JUDGMENT && daysActive >= 7;

      if (lowSpendLongRun) {
        // Enough days + some data — graduate to test_results even under $150
        phase = 'test_results';
        if (recent.roas >= T.SCALE_ROAS) {
          status = '🟢 Winner — Scale';
        } else if (recent.roas >= T.TARGET_ROAS) {
          status = '✅ Passed — Graduate';
        } else if (recent.roas >= T.KILL_ROAS) {
          status = '🟡 Borderline (low spend)';
        } else {
          status = '🔴 Low delivery + poor ROAS — Kill';
        }
      } else {
        phase = 'testing';
        if (recent.spend < T.MIN_SPEND_FOR_JUDGMENT) {
          status = '🧪 Ramping';
        } else if (recent.roas >= T.TARGET_ROAS) {
          status = '🧪 Promising';
        } else if (recent.roas >= T.KILL_ROAS) {
          status = '🧪 Testing';
        } else {
          status = '🧪 Weak Signal';
        }
      }
    } else if (isNew && recent.spend >= T.TESTING_SPEND_CAP) {
      // Crossed the spend threshold in 7d — enough data to make a call
      phase = 'test_results';
      if (recent.roas >= T.SCALE_ROAS) {
        status = '🟢 Winner — Scale';
      } else if (recent.roas >= T.TARGET_ROAS) {
        status = '✅ Passed — Graduate';
      } else if (recent.roas >= T.KILL_ROAS) {
        status = '🟡 Borderline';
      } else {
        status = '🔴 Failed — Kill';
      }
    } else {
      // Has prior history AND recent spend — active proven creative
      phase = 'active';
      if (recent.roas >= T.SCALE_ROAS) {
        status = '🟢 Scale';
      } else if (recent.roas >= T.TARGET_ROAS) {
        status = '✅ Profitable';
      } else if (recent.roas >= T.KILL_ROAS) {
        status = '🟡 Declining';
      } else {
        status = '🔴 Kill';
      }
    }

    // ROAS trend (all-time vs recent 7d) — only meaningful for active creatives
    let roasTrend = '';
    if (full.roas > 0 && recent.roas > 0 && phase === 'active') {
      const delta = ((recent.roas - full.roas) / full.roas * 100);
      if (delta > 5) roasTrend = '▲';
      else if (delta < -5) roasTrend = '▼';
      else roasTrend = '—';
    }

    return {
      name: adName, imgUrl, status, phase, roasTrend,
      full, recent, isNew
    };
  });

  // Cache for action items engine
  DATA_CACHE.creatives = allCreatives;

  // Detect creative format for language
  const hasVideoAds = allCreatives.some(c =>
    c.recent.spend > 0 && (c.recent.thumbstop > 0 || c.recent.holdRate > 0)
  );
  const declineHintCreative = hasVideoAds ? 'test new hook or kill' : 'test new image/headline or kill';

  // ── Split into sections (no inactive/paused) ─────────────────────────
  const testing = allCreatives.filter(c => c.phase === 'testing')
    .sort((a, b) => b.recent.spend - a.recent.spend);

  const testResults = allCreatives.filter(c => c.phase === 'test_results')
    .sort((a, b) => b.recent.roas - a.recent.roas); // Best performers first

  const active = allCreatives.filter(c => c.phase === 'active')
    .sort((a, b) => b.recent.spend - a.recent.spend);

  // ── Write to sheet ────────────────────────────────────────────────────
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.TABS.META_CREATIVE);
  if (!sheet) sheet = ss.insertSheet(CONFIG.TABS.META_CREATIVE);
  sheet.clear();
  sheet.clearConditionalFormatRules();

  let r = 1;

  // Helper: write a section of creatives
  function writeCreativeSection(sectionTitle, sectionSubtitle, creatives, bgColor) {
    sheet.getRange(r, 1).setValue(sectionTitle).setFontWeight('bold').setFontSize(12);
    r += 1;
    if (sectionSubtitle) {
      sheet.getRange(r, 1).setValue(sectionSubtitle).setFontSize(9).setFontColor('#666666').setFontStyle('italic');
      r += 1;
    }

    const headers = hasVideoAds
      ? ['Thumb', 'Status', 'Ad Name',
         '7d Spend', '7d Purch', '7d Rev', '7d ROAS', '7d CPA', '7d CPC', '7d CTR',
         'Thumbstop', 'Hold Rate',
         'All Spend', 'All ROAS', 'Trend', 'Decision']
      : ['Thumb', 'Status', 'Ad Name',
         '7d Spend', '7d Purch', '7d Rev', '7d ROAS', '7d CPA', '7d CPC', '7d CTR',
         'All Spend', 'All ROAS', 'Trend', 'Decision'];

    const recentColCount = hasVideoAds ? 9 : 7; // number of 7d metric columns for header coloring
    const allTimeStartCol = hasVideoAds ? 13 : 11;

    sheet.getRange(r, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(r, 1, 1, headers.length)
      .setFontWeight('bold').setBackground('#000000').setFontColor('#FFFFFF')
      .setHorizontalAlignment('center').setVerticalAlignment('middle').setFontSize(9);
    sheet.getRange(r, 4, 1, recentColCount).setBackground('#1a3a5c');
    sheet.getRange(r, allTimeStartCol, 1, 2).setBackground('#333333');
    r += 1;

    if (creatives.length === 0) {
      sheet.getRange(r, 1).setValue('No creatives in this category');
      sheet.getRange(r, 1).setFontColor('#888888').setFontStyle('italic');
      r += 2;
      return;
    }

    creatives.forEach((c, i) => {
      const imageFormula = c.imgUrl ? `=IMAGE("${c.imgUrl}")` : '';
      const rec = c.recent;
      const all = c.full;

      // Decision guidance — varies by phase
      let decision = '';
      if (c.phase === 'testing') {
        if (rec.spend < T.MIN_SPEND_FOR_JUDGMENT) decision = 'Ramping — not enough data yet';
        else if (rec.roas >= T.SCALE_ROAS) decision = 'Strong early signal — keep running to hit threshold';
        else if (rec.roas >= T.TARGET_ROAS) decision = 'On track — let the 7d test complete';
        else if (rec.roas >= T.KILL_ROAS) decision = 'Marginal — give it a few more days';
        else decision = 'Weak — consider killing early to save budget';
      } else if (c.phase === 'test_results') {
        if (rec.roas >= T.SCALE_ROAS) decision = '✓ PASSED — Move to main scaling campaign';
        else if (rec.roas >= T.TARGET_ROAS) decision = '✓ PASSED — Graduate to main campaigns';
        else if (rec.roas >= T.KILL_ROAS) decision = '~ BORDERLINE — Run 3 more days or kill';
        else decision = '✗ FAILED — Kill and redirect budget';
      } else {
        if (rec.roas >= T.SCALE_ROAS) decision = 'Winner — increase CBO budget or move to scaling campaign';
        else if (rec.roas >= T.TARGET_ROAS && c.roasTrend === '▼') decision = 'Profitable but declining — watch for fatigue';
        else if (rec.roas >= T.TARGET_ROAS) decision = 'Healthy — maintain';
        else if (rec.roas >= T.KILL_ROAS) decision = 'Underperforming recently — ' + declineHintCreative;
        else decision = 'Kill — recent ROAS below floor';
      }

      const rowData = hasVideoAds
        ? [imageFormula, c.status, c.name,
           rec.spend, rec.purchases, rec.revenue, rec.roas, rec.cpa, rec.cpc, rec.ctr,
           rec.thumbstop, rec.holdRate,
           all.spend, all.roas, c.roasTrend, decision]
        : [imageFormula, c.status, c.name,
           rec.spend, rec.purchases, rec.revenue, rec.roas, rec.cpa, rec.cpc, rec.ctr,
           all.spend, all.roas, c.roasTrend, decision];

      sheet.getRange(r, 1, 1, headers.length).setValues([rowData]);

      // Number formats — column indices shift based on format
      sheet.getRange(r, 4).setNumberFormat('"$"#,##0');       // 7d spend
      sheet.getRange(r, 5).setNumberFormat('#,##0');           // 7d purchases
      sheet.getRange(r, 6).setNumberFormat('"$"#,##0');        // 7d revenue
      sheet.getRange(r, 7).setNumberFormat('0.00"x"');         // 7d ROAS
      sheet.getRange(r, 8).setNumberFormat('"$"#,##0.00');     // 7d CPA
      sheet.getRange(r, 9).setNumberFormat('"$"#,##0.00');     // 7d CPC
      sheet.getRange(r, 10).setNumberFormat('0.00"%"');        // 7d CTR
      if (hasVideoAds) {
        sheet.getRange(r, 11).setNumberFormat('0.0"%"');       // Thumbstop
        sheet.getRange(r, 12).setNumberFormat('0.0"%"');       // Hold rate
      }
      sheet.getRange(r, allTimeStartCol).setNumberFormat('"$"#,##0');    // All spend
      sheet.getRange(r, allTimeStartCol + 1).setNumberFormat('0.00"x"'); // All ROAS

      const decisionCol = headers.length;
      const trendCol = headers.length - 1;

      // Center alignment
      sheet.getRange(r, 1, 1, headers.length).setHorizontalAlignment('center').setVerticalAlignment('middle');
      sheet.getRange(r, 3).setHorizontalAlignment('left');     // Name left-aligned
      sheet.getRange(r, decisionCol).setHorizontalAlignment('left');  // Decision left-aligned

      // Color ROAS cells
      const roas7Cell = sheet.getRange(r, 7);
      if (rec.roas >= T.SCALE_ROAS) roas7Cell.setBackground('#e6f4ea').setFontColor('#137333').setFontWeight('bold');
      else if (rec.roas >= T.TARGET_ROAS) roas7Cell.setBackground('#e6f4ea').setFontColor('#137333');
      else if (rec.roas >= T.KILL_ROAS) roas7Cell.setBackground('#fef7e0').setFontColor('#E37400');
      else if (rec.spend > 0) roas7Cell.setBackground('#fde8e8').setFontColor('#C5221F');

      // Color status cell
      const statusCell = sheet.getRange(r, 2);
      if (c.status.includes('🟢')) statusCell.setBackground('#e6f4ea');
      else if (c.status.includes('✅')) statusCell.setBackground('#e6f4ea');
      else if (c.status.includes('🟡')) statusCell.setBackground('#fef7e0');
      else if (c.status.includes('🔴')) statusCell.setBackground('#fde8e8');
      else if (c.status.includes('🧪')) statusCell.setBackground('#e8eaf6');
      else if (c.status.includes('⏸️')) statusCell.setBackground('#f5f5f5');

      // Trend color
      const trendCell = sheet.getRange(r, trendCol);
      if (c.roasTrend === '▲') trendCell.setFontColor('#137333').setFontWeight('bold');
      else if (c.roasTrend === '▼') trendCell.setFontColor('#C5221F').setFontWeight('bold');
      else trendCell.setFontColor('#888888');

      // Light alternating row
      if (i % 2 === 0) {
        sheet.getRange(r, 3, 1, headers.length - 2).setBackground('#fafafa');
      }

      r++;
    });

    // Row heights for thumbnails
    const dataStartRow = r - creatives.length;
    if (creatives.length > 0) {
      sheet.setRowHeights(dataStartRow, creatives.length, 70);
    }
    r += 1; // spacing between sections
  }

  // ── CREATIVE VELOCITY SUMMARY ─────────────────────────────────────────
  sheet.getRange(r, 1).setValue('CREATIVE VELOCITY').setFontWeight('bold').setFontSize(12);
  r += 1;

  const totalTested = testResults.length + testing.length;
  const passed = testResults.filter(c => c.recent.roas >= T.TARGET_ROAS).length;
  const failed = testResults.filter(c => c.recent.roas < T.KILL_ROAS).length;
  const borderline = testResults.length - passed - failed;
  const activeCount = active.length;
  const decliningCount = active.filter(c => c.roasTrend === '▼').length;
  const scalingCount = active.filter(c => c.status.includes('🟢')).length;
  const totalActiveSpend = active.reduce((s, c) => s + c.recent.spend, 0);
  const topCreativeSpend = active.length > 0
    ? Math.max(...active.map(c => c.recent.spend)) : 0;
  const topConcentration = totalActiveSpend > 0
    ? (topCreativeSpend / totalActiveSpend * 100).toFixed(0) : 0;

  // Use merged cells to fit the creative tab's narrow column grid
  // Metric spans cols 1-3, Value spans cols 4-8
  const velHeaderRange = sheet.getRange(r, 1, 1, 8);
  sheet.getRange(r, 1, 1, 3).merge().setValue('Metric');
  sheet.getRange(r, 4, 1, 5).merge().setValue('Value');
  velHeaderRange.setFontWeight('bold').setBackground('#000000').setFontColor('#FFFFFF')
    .setHorizontalAlignment('center');
  r += 1;

  const winRate = totalTested > 0 ? (passed / totalTested * 100).toFixed(0) + '%' : 'N/A';
  const velData = [
    ['In Testing', `${testing.length} creative(s) gathering data`],
    ['Awaiting Decision', `${testResults.length} — ${passed} passed, ${failed} failed, ${borderline} borderline`],
    ['Test Win Rate', winRate],
    ['Active Creatives', `${activeCount} running (${scalingCount} scaling, ${decliningCount} declining)`],
    ['Top Concentration', `${topConcentration}% of active spend in one creative`],
  ];

  velData.forEach((vRow, i) => {
    sheet.getRange(r, 1, 1, 3).merge().setValue(vRow[0]).setFontWeight('bold');
    sheet.getRange(r, 4, 1, 5).merge().setValue(vRow[1]);
    sheet.getRange(r, 1, 1, 8).setVerticalAlignment('middle').setHorizontalAlignment('left');
    if (i % 2 === 0) sheet.getRange(r, 1, 1, 8).setBackground('#fafafa');
    r++;
  });

  // Warn on low test volume or high concentration
  if (totalTested === 0) {
    sheet.getRange(r, 1, 1, 8).merge();
    sheet.getRange(r, 1).setValue('⚠️ No new creatives in testing. You need fresh tests to sustain performance.')
      .setFontColor('#E37400').setFontWeight('bold');
    r += 2;
  } else if (parseInt(topConcentration) > 40) {
    sheet.getRange(r, 1, 1, 8).merge();
    sheet.getRange(r, 1).setValue(`⚠️ Top creative is ${topConcentration}% of spend. Diversify — if it fatigues, account tanks.`)
      .setFontColor('#E37400').setFontWeight('bold');
    r += 2;
  } else {
    r += 1;
  }

  // ── Write sections ────────────────────────────────────────────────────
  writeCreativeSection(
    '📋 TEST RESULTS — Decision Time',
    `New creatives that hit $${T.TESTING_SPEND_CAP}+ spend in 7 days. Enough data to call. Green = scale. Red = kill.`,
    testResults
  );

  writeCreativeSection(
    '🧪 STILL TESTING — Gathering Data',
    `Launched recently, < $${T.TESTING_SPEND_CAP} spend so far. Let them run unless signal is clearly bad.`,
    testing
  );

  writeCreativeSection(
    '📊 ACTIVE PERFORMERS — How Your Proven Creatives Are Doing Now',
    `Creatives with history. "7d" = this week. "Trend" = 7d ROAS vs all-time. Only showing ads with recent spend.`,
    active
  );

  // ── Column widths (format-aware) ────────────────────────────────────
  sheet.setColumnWidth(1, 75);   // Thumb
  sheet.setColumnWidth(2, 110);  // Status
  sheet.setColumnWidth(3, 250);  // Name
  for (let c = 4; c <= 10; c++) sheet.setColumnWidth(c, 85);  // 7d metrics
  if (hasVideoAds) {
    sheet.setColumnWidth(11, 75);  // Thumbstop
    sheet.setColumnWidth(12, 70);  // Hold Rate
    sheet.setColumnWidth(13, 85);  // All Spend
    sheet.setColumnWidth(14, 70);  // All ROAS
    sheet.setColumnWidth(15, 50);  // Trend
    sheet.setColumnWidth(16, 320); // Decision
  } else {
    sheet.setColumnWidth(11, 85);  // All Spend
    sheet.setColumnWidth(12, 70);  // All ROAS
    sheet.setColumnWidth(13, 50);  // Trend
    sheet.setColumnWidth(14, 320); // Decision
  }

  log.push(`meta_creative: ${testResults.length} test results + ${testing.length} testing + ${active.length} active creatives synced`);
}

function syncMetaAgeGender(log) {
  const T = CONFIG.THRESHOLDS;
  const fields = [
    'spend', 'impressions', 'reach', 'clicks', 'actions', 'action_values'
  ];

  const params = {
    level: 'account',
    time_increment: 'all_days',
    breakdowns: 'age,gender',
    fields: fields.join(','),
  };

  const data = fetchInsights(params, log);

  // ── Parse raw data, skip unknowns ─────────────────────────────────────
  const parsed = data
    .filter(row => row.gender !== 'unknown' && row.age !== 'Unknown')
    .map(row => {
      const spend = parseFloat(row.spend) || 0;
      const purchases = extractAction(row.actions, 'omni_purchase') || 0;
      const purchaseValue = extractAction(row.action_values, 'omni_purchase') || 0;
      return {
        age: row.age,
        gender: row.gender,
        spend,
        impressions: parseInt(row.impressions) || 0,
        reach: parseInt(row.reach) || 0,
        clicks: parseInt(row.clicks) || 0,
        purchases,
        revenue: purchaseValue,
      };
    });

  // Cache for action items engine
  DATA_CACHE.ageGender = parsed.map(p => ({
    age: p.age, gender: p.gender, spend: p.spend, revenue: p.revenue,
    purchases: p.purchases, roas: p.spend > 0 ? p.revenue / p.spend : 0
  }));

  const totalSpend = parsed.reduce((s, r) => s + r.spend, 0);
  const totalRevenue = parsed.reduce((s, r) => s + r.revenue, 0);
  const totalPurchases = parsed.reduce((s, r) => s + r.purchases, 0);
  const totalImpressions = parsed.reduce((s, r) => s + r.impressions, 0);
  const totalClicks = parsed.reduce((s, r) => s + r.clicks, 0);

  // ── Aggregate helper ──────────────────────────────────────────────────
  function agg(rows) {
    const a = { spend: 0, revenue: 0, purchases: 0, impressions: 0, clicks: 0 };
    rows.forEach(r => {
      a.spend += r.spend; a.revenue += r.revenue; a.purchases += r.purchases;
      a.impressions += r.impressions; a.clicks += r.clicks;
    });
    a.roas = a.spend > 0 ? a.revenue / a.spend : 0;
    a.cpa = a.purchases > 0 ? a.spend / a.purchases : 0;
    a.cpc = a.clicks > 0 ? a.spend / a.clicks : 0;
    a.cpm = a.impressions > 0 ? (a.spend / a.impressions) * 1000 : 0;
    a.ctr = a.impressions > 0 ? (a.clicks / a.impressions * 100) : 0;
    a.cvr = a.clicks > 0 ? (a.purchases / a.clicks * 100) : 0;
    a.share = totalSpend > 0 ? a.spend / totalSpend : 0;
    return a;
  }

  // ── Build the sheet ───────────────────────────────────────────────────
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.TABS.META_AGE);
  if (!sheet) sheet = ss.insertSheet(CONFIG.TABS.META_AGE);
  sheet.clear();
  sheet.clearConditionalFormatRules();

  const dateRange = getDateRange();
  const metricHeaders = ['Spend', '% Budget', 'Purchases', 'Revenue', 'ROAS', 'CPA', 'CVR', 'CPC', 'CPM', 'CTR', 'Impressions', 'Clicks'];

  // Row-writing helper — keeps formatting consistent
  function writeMetricRow(sheet, r, label, a, headers) {
    const vals = [
      label, a.spend, (a.share * 100).toFixed(1) + '%', a.purchases, a.revenue,
      a.roas, a.cpa, a.cvr, a.cpc, a.cpm, a.ctr, a.impressions, a.clicks
    ];
    sheet.getRange(r, 1, 1, vals.length).setValues([vals]);
    sheet.getRange(r, 1, 1, vals.length).setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.getRange(r, 1).setHorizontalAlignment('left').setFontWeight('bold');

    // Number formats
    sheet.getRange(r, 2).setNumberFormat('"$"#,##0');
    sheet.getRange(r, 4).setNumberFormat('#,##0');
    sheet.getRange(r, 5).setNumberFormat('"$"#,##0');
    sheet.getRange(r, 6).setNumberFormat('0.00"x"');
    sheet.getRange(r, 7).setNumberFormat('"$"#,##0.00');
    sheet.getRange(r, 8).setNumberFormat('0.00"%"');
    sheet.getRange(r, 9).setNumberFormat('"$"#,##0.00');
    sheet.getRange(r, 10).setNumberFormat('"$"#,##0.00');
    sheet.getRange(r, 11).setNumberFormat('0.00"%"');
    sheet.getRange(r, 12).setNumberFormat('#,##0');
    sheet.getRange(r, 13).setNumberFormat('#,##0');

    // Color ROAS
    const roasCell = sheet.getRange(r, 6);
    if (a.roas >= T.SCALE_ROAS) roasCell.setBackground('#e6f4ea').setFontColor('#137333').setFontWeight('bold');
    else if (a.roas >= T.TARGET_ROAS) roasCell.setBackground('#e6f4ea').setFontColor('#137333');
    else if (a.roas >= T.KILL_ROAS) roasCell.setBackground('#fef7e0').setFontColor('#E37400');
    else if (a.spend > 0) roasCell.setBackground('#fde8e8').setFontColor('#C5221F');
  }

  function writeHeaderRow(sheet, r, firstCol) {
    const h = [firstCol, ...metricHeaders];
    sheet.getRange(r, 1, 1, h.length).setValues([h]);
    sheet.getRange(r, 1, 1, h.length)
      .setFontWeight('bold').setBackground('#000000').setFontColor('#FFFFFF')
      .setHorizontalAlignment('center').setVerticalAlignment('middle').setFontSize(9);
  }

  let r = 1;

  // ── Title + Verdict ───────────────────────────────────────────────────
  sheet.getRange(r, 1).setValue('AUDIENCE BREAKDOWN').setFontWeight('bold').setFontSize(14);
  r += 1;
  sheet.getRange(r, 1).setValue(`${dateRange.since} → ${dateRange.until}`).setFontSize(9).setFontColor('#666666');
  r += 1;

  // Quick verdict: find best and worst segments
  const AGE_ORDER = ['18-24', '25-34', '35-44', '45-54', '55-64', '65+'];
  const segmentList = [];
  AGE_ORDER.forEach(age => {
    ['female', 'male'].forEach(gender => {
      const rows = parsed.filter(p => p.age === age && p.gender === gender);
      if (rows.length === 0) return;
      const a = agg(rows);
      if (a.spend > 100) segmentList.push({ label: `${age} ${gender}`, ...a });
    });
  });
  segmentList.sort((a, b) => b.roas - a.roas);

  const best = segmentList[0];
  const worst = segmentList[segmentList.length - 1];
  let verdict = '';
  if (best && worst) {
    verdict = `Best: ${best.label} (${best.roas.toFixed(2)}x, ${(best.share * 100).toFixed(0)}% budget)`;
    verdict += `  ·  Worst: ${worst.label} (${worst.roas.toFixed(2)}x, ${(worst.share * 100).toFixed(0)}% budget)`;
  }
  sheet.getRange(r, 1, 1, 8).merge();
  sheet.getRange(r, 1).setValue(verdict).setFontSize(10).setFontWeight('bold').setWrap(true);
  sheet.setRowHeight(r, 30);
  r += 2;

  // ── SECTION 1: Gender ─────────────────────────────────────────────────
  sheet.getRange(r, 1).setValue('GENDER').setFontWeight('bold').setFontSize(12);
  r += 1;
  writeHeaderRow(sheet, r, 'Gender');
  r += 1;

  ['female', 'male'].forEach((g, i) => {
    const a = agg(parsed.filter(p => p.gender === g));
    writeMetricRow(sheet, r, g.charAt(0).toUpperCase() + g.slice(1), a);
    if (i % 2 === 0) sheet.getRange(r, 1, 1, metricHeaders.length + 1).setBackground('#fafafa');
    r++;
  });

  // Gender total
  const gTotal = agg(parsed);
  gTotal.share = 1;
  writeMetricRow(sheet, r, 'Total', gTotal);
  sheet.getRange(r, 1, 1, metricHeaders.length + 1).setFontWeight('bold').setBackground('#e8eaf6');
  r += 2;

  // ── SECTION 2: Age Group ──────────────────────────────────────────────
  sheet.getRange(r, 1).setValue('AGE GROUP').setFontWeight('bold').setFontSize(12);
  r += 1;
  writeHeaderRow(sheet, r, 'Age');
  r += 1;

  AGE_ORDER.forEach((age, i) => {
    const rows = parsed.filter(p => p.age === age);
    if (rows.length === 0) return;
    const a = agg(rows);
    writeMetricRow(sheet, r, age, a);

    // Bold high-spend rows
    if (a.share >= 0.15) sheet.getRange(r, 1, 1, metricHeaders.length + 1).setFontWeight('bold');
    if (i % 2 === 0) sheet.getRange(r, 1, 1, metricHeaders.length + 1).setBackground('#fafafa');
    r++;
  });

  // Age total
  const aTotal = agg(parsed);
  aTotal.share = 1;
  writeMetricRow(sheet, r, 'Total', aTotal);
  sheet.getRange(r, 1, 1, metricHeaders.length + 1).setFontWeight('bold').setBackground('#e8eaf6');
  r += 2;

  // ── SECTION 3: Age × Gender (compact, sorted by ROAS desc) ───────────
  sheet.getRange(r, 1).setValue('AGE × GENDER DETAIL').setFontWeight('bold').setFontSize(12);
  r += 1;
  sheet.getRange(r, 1).setValue('Sorted by ROAS. Only segments with $50+ spend shown.')
    .setFontSize(9).setFontColor('#666666').setFontStyle('italic');
  r += 1;
  writeHeaderRow(sheet, r, 'Segment');
  r += 1;

  // Build all age×gender combos, filter low spend, sort by ROAS
  const combos = [];
  AGE_ORDER.forEach(age => {
    ['female', 'male'].forEach(gender => {
      const rows = parsed.filter(p => p.age === age && p.gender === gender);
      if (rows.length === 0) return;
      const a = agg(rows);
      if (a.spend < 50) return; // Too little spend to matter
      combos.push({ label: `${age} / ${gender.charAt(0).toUpperCase() + gender.slice(1)}`, ...a });
    });
  });
  combos.sort((a, b) => b.roas - a.roas);

  combos.forEach((seg, i) => {
    writeMetricRow(sheet, r, seg.label, seg);
    if (i % 2 === 0) sheet.getRange(r, 1, 1, metricHeaders.length + 1).setBackground('#fafafa');
    r++;
  });

  // Combo total
  const comboTotal = agg(parsed);
  comboTotal.share = 1;
  writeMetricRow(sheet, r, 'Total', comboTotal);
  sheet.getRange(r, 1, 1, metricHeaders.length + 1).setFontWeight('bold').setBackground('#e8eaf6');

  // ── Column widths ─────────────────────────────────────────────────────
  sheet.setColumnWidth(1, 140);  // Label
  sheet.setColumnWidth(2, 95);   // Spend
  sheet.setColumnWidth(3, 75);   // % Budget
  sheet.setColumnWidth(4, 80);   // Purchases
  sheet.setColumnWidth(5, 95);   // Revenue
  sheet.setColumnWidth(6, 70);   // ROAS
  sheet.setColumnWidth(7, 80);   // CPA
  sheet.setColumnWidth(8, 65);   // CVR
  sheet.setColumnWidth(9, 75);   // CPC
  sheet.setColumnWidth(10, 75);  // CPM
  sheet.setColumnWidth(11, 65);  // CTR
  sheet.setColumnWidth(12, 95);  // Impressions
  sheet.setColumnWidth(13, 80);  // Clicks

  log.push(`meta_age_gender: ${combos.length} segments synced`);
}


// =============================================================================
// CONDITIONAL FORMATTING — auto-color ROAS, CPA, frequency columns
// =============================================================================

function applyConditionalFormatting(tabName, headers, rowCount) {
  if (rowCount === 0) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) return;

  const T = CONFIG.THRESHOLDS;

  for (let i = 0; i < headers.length; i++) {
    const col = headers[i].toLowerCase();
    const range = sheet.getRange(2, i + 1, rowCount, 1);

    if (col === 'roas') {
      // Green if above target, yellow if marginal, red if bad
      sheet.getRange(2, i + 1, rowCount, 1).clearFormat();
      const rules = sheet.getConditionalFormatRules();

      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThanOrEqualTo(T.TARGET_ROAS)
        .setBackground('#e6f4ea').setFontColor('#137333')
        .setRanges([range]).build());

      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenNumberBetween(T.KILL_ROAS, T.TARGET_ROAS)
        .setBackground('#fef7e0').setFontColor('#E37400')
        .setRanges([range]).build());

      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThan(T.KILL_ROAS)
        .setBackground('#fde8e8').setFontColor('#C5221F')
        .setRanges([range]).build());

      sheet.setConditionalFormatRules(rules);
    }

    if (col === 'frequency') {
      const rules = sheet.getConditionalFormatRules();

      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThanOrEqualTo(T.FREQUENCY_CRITICAL)
        .setBackground('#fde8e8').setFontColor('#C5221F')
        .setRanges([range]).build());

      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThanOrEqualTo(T.FREQUENCY_WARNING)
        .setBackground('#fef7e0').setFontColor('#E37400')
        .setRanges([range]).build());

      sheet.setConditionalFormatRules(rules);
    }
  }
}


// =============================================================================
// META API HELPERS (unchanged)
// =============================================================================

function fetchInsights(params, log) {
  const dateRange = getDateRange();
  const baseUrl = `https://graph.facebook.com/${CONFIG.API_VERSION}/${CONFIG.AD_ACCOUNT_ID}/insights`;

  let queryParams = [
    `access_token=${CONFIG.ACCESS_TOKEN}`,
    `time_range=${encodeURIComponent(`{"since":"${dateRange.since}","until":"${dateRange.until}"}`)}`,
    `level=${params.level}`,
    `fields=${params.fields}`,
    `limit=500`
  ];

  if (params.time_increment) queryParams.push(`time_increment=${params.time_increment}`);
  if (params.breakdowns) queryParams.push(`breakdowns=${params.breakdowns}`);

  let url = `${baseUrl}?${queryParams.join('&')}`;
  let allData = [];
  let pageCount = 0;

  while (url && pageCount < 20) {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const code = response.getResponseCode();

    if (code !== 200) {
      const errorBody = response.getContentText();
      log.push(`API Error (${code}): ${errorBody.substring(0, 200)}`);
      throw new Error(`Meta API returned ${code}`);
    }

    const json = JSON.parse(response.getContentText());
    if (json.data) allData = allData.concat(json.data);

    url = json.paging && json.paging.next ? json.paging.next : null;
    pageCount++;
    if (url) Utilities.sleep(500);
  }

  return allData;
}

function extractAction(actions, actionType) {
  if (!actions || !Array.isArray(actions)) return 0;
  const found = actions.find(a => a.action_type === actionType);
  return found ? parseFloat(found.value) : 0;
}

function extractVideoAction(videoActions) {
  if (!videoActions || !Array.isArray(videoActions)) return 0;
  return videoActions.reduce((sum, a) => sum + (parseFloat(a.value) || 0), 0);
}

function getDateRange() {
  const now = new Date();
  const until = Utilities.formatDate(now, 'UTC', 'yyyy-MM-dd');
  const sinceDate = new Date(now);
  sinceDate.setDate(sinceDate.getDate() - CONFIG.LOOKBACK_DAYS);
  const since = Utilities.formatDate(sinceDate, 'UTC', 'yyyy-MM-dd');
  return { since, until };
}

function fetchAdThumbnails(log) {
  // Fetch thumbnails AND effective_status — effective_status accounts for parent campaign/adset being paused
  const fieldsString = encodeURIComponent('id,name,effective_status,creative{thumbnail_url,image_url}');
  let url = `https://graph.facebook.com/${CONFIG.API_VERSION}/${CONFIG.AD_ACCOUNT_ID}/ads?fields=${fieldsString}&limit=500&access_token=${CONFIG.ACCESS_TOKEN}`;

  let creativeMap = {};   // ad_id → thumbnail_url
  let statusMap = {};     // ad_id → effective_status
  let nameStatusMap = {}; // ad_name → effective_status (for matching with insights which group by name)
  let pageCount = 0;

  while (url && pageCount < 20) {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const code = response.getResponseCode();

    if (code === 200) {
      const json = JSON.parse(response.getContentText());
      if (json.data) {
        json.data.forEach(ad => {
          if (ad.creative) {
            const imgUrl = ad.creative.thumbnail_url || ad.creative.image_url;
            if (imgUrl) creativeMap[ad.id] = imgUrl;
          }
          if (ad.effective_status) {
            statusMap[ad.id] = ad.effective_status;
            // Map by name too — if ANY ad with this name is truly ACTIVE, mark it active
            if (ad.name) {
              const existing = nameStatusMap[ad.name];
              if (!existing || ad.effective_status === 'ACTIVE') {
                nameStatusMap[ad.name] = ad.effective_status;
              }
            }
          }
        });
      }
      url = json.paging && json.paging.next ? json.paging.next : null;
    } else {
      log.push(`Creative API Error (${code}): ${response.getContentText()}`);
      break;
    }
    pageCount++;
  }

  return { creativeMap, statusMap, nameStatusMap };
}


// =============================================================================
// GOOGLE SHEETS HELPERS (unchanged)
// =============================================================================

function writeToSheet(tabName, headers, rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(tabName);
  if (!sheet) sheet = ss.insertSheet(tabName);

  sheet.clear();
  sheet.clearConditionalFormatRules();

  if (headers.length > 0) {
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setBackground('#000000').setFontColor('#FFFFFF')
      .setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
  }

  if (rows.length > 0) {
    const dataRange = sheet.getRange(2, 1, rows.length, rows[0].length);
    dataRange.setValues(rows);
    dataRange.setVerticalAlignment('middle').setHorizontalAlignment('center');

    for (let i = 0; i < headers.length; i++) {
      const colName = headers[i].toLowerCase();
      const colRange = sheet.getRange(2, i + 1, rows.length, 1);

      if (['spend', 'purchase_value', 'cpm', 'cpc', 'cpa', 'revenue'].includes(colName)) {
        colRange.setNumberFormat('"$"#,##0.00');
      } else if (['impressions', 'reach', 'clicks', 'purchases', 'atc'].includes(colName)) {
        colRange.setNumberFormat('#,##0');
      } else if (['roas', 'frequency'].includes(colName)) {
        colRange.setNumberFormat('0.00');
      }
    }
  }

  sheet.setFrozenRows(1);
  for (let i = 1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);
    sheet.setColumnWidth(i, sheet.getColumnWidth(i) + 15);
  }
}

function writeLog(logEntries) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.TABS.LOG);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.TABS.LOG);
    sheet.getRange(1, 1, 1, 2).setValues([['timestamp', 'message']]);
    sheet.getRange(1, 1, 1, 2).setFontWeight('bold');
  }

  const lastRow = sheet.getLastRow();
  const timestamp = new Date().toISOString();
  const newRows = logEntries.map(entry => [timestamp, entry]);
  if (newRows.length > 0) {
    sheet.getRange(lastRow + 1, 1, newRows.length, 2).setValues(newRows);
  }
}


// =============================================================================
// SETUP & SCHEDULING
// =============================================================================

function setupTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'syncAll') ScriptApp.deleteTrigger(trigger);
  });

  ScriptApp.newTrigger('syncAll')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();

  Logger.log('Daily sync trigger created for 6 AM');
}

function testConnection() {
  const dateRange = getDateRange();
  const url = `https://graph.facebook.com/${CONFIG.API_VERSION}/${CONFIG.AD_ACCOUNT_ID}/insights?access_token=${CONFIG.ACCESS_TOKEN}&time_range={"since":"${dateRange.since}","until":"${dateRange.until}"}&level=account&fields=spend,impressions&limit=1`;

  try {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const code = response.getResponseCode();
    const body = response.getContentText();

    if (code === 200) {
      const json = JSON.parse(body);
      Logger.log('SUCCESS — Connection works!');
      Logger.log(`Account spend data: ${JSON.stringify(json.data)}`);
    } else {
      Logger.log(`FAILED (${code}): ${body}`);
    }
  } catch (e) {
    Logger.log(`ERROR: ${e.message}`);
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Meta Sync')
    .addItem('🔄 Sync Now', 'syncAll')
    .addItem('📊 Rebuild Dashboard Only', 'rebuildDashboardOnly')
    .addItem('🔌 Test Connection', 'testConnection')
    .addItem('⏰ Setup Daily Trigger', 'setupTriggers')
    .addToUi();
}

/** Rebuild just the dashboard without re-pulling API data */
// =============================================================================
// AI INSIGHT ENGINE
// =============================================================================

/**
 * Generate AI-powered analysis from structured dashboard data.
 * Returns insight text or empty string if AI is not configured.
 */
function generateAIInsight(dataPacket) {
  const ai = CONFIG.AI;
  if (!ai.API_KEY) return '';

  const prompt = `You are a senior Meta ads media buyer analyzing a DTC skincare e-commerce account. 
You're looking at the weekly performance data and need to give the account manager a brief, actionable analysis.

RULES:
- Be direct and specific. No fluff, no "great job" cheerleading.
- Reference actual numbers. Don't just say "CTR dropped" — say "CTR dropped 10.1% WoW from 3.24% to 2.91%"
- Prioritize what matters most for ROAS and profitability
- If multiple issues compound, explain the chain (e.g. "CTR drop → CPC spike → CPA increase → ROAS decline")
- End with 2-3 specific action items for this week, ranked by impact
- Keep it under 200 words
- The account runs static image ads (not video)

DATA:
${JSON.stringify(dataPacket, null, 2)}`;

  try {
    let responseText = '';

    if (ai.PROVIDER === 'claude') {
      const resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': ai.API_KEY,
          'anthropic-version': '2023-06-01'
        },
        payload: JSON.stringify({
          model: ai.MODEL || 'claude-sonnet-4-20250514',
          max_tokens: 500,
          messages: [{ role: 'user', content: prompt }]
        }),
        muteHttpExceptions: true
      });

      const code = resp.getResponseCode();
      if (code === 200) {
        const json = JSON.parse(resp.getContentText());
        responseText = json.content && json.content[0] ? json.content[0].text : '';
      } else {
        return `[AI Error ${code}: ${resp.getContentText().substring(0, 100)}]`;
      }
    } else if (ai.PROVIDER === 'gemini') {
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${ai.MODEL || 'gemini-2.0-flash'}:generateContent?key=${ai.API_KEY}`;
      const resp = UrlFetchApp.fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        payload: JSON.stringify({
          contents: [{ parts: [{ text: prompt }] }],
          generationConfig: { maxOutputTokens: 500 }
        }),
        muteHttpExceptions: true
      });

      const code = resp.getResponseCode();
      if (code === 200) {
        const json = JSON.parse(resp.getContentText());
        responseText = json.candidates && json.candidates[0]
          ? json.candidates[0].content.parts[0].text : '';
      } else {
        return `[AI Error ${code}: ${resp.getContentText().substring(0, 100)}]`;
      }
    }

    return responseText;
  } catch (e) {
    return `[AI Error: ${e.message}]`;
  }
}

function rebuildDashboardOnly() {
  const log = [];
  buildDashboard(log);
  writeLog(log);
}
