# Meta & Google Ads Dashboard Sync

This repository contains two scripts designed to synchronize advertising data from **Meta Ads** and **Google Ads** directly into **Google Sheets**. Together, they form an automated performance dashboard to track KPIs, find creative fatigue, and highlight actionable insights.

![Overview](assets/dashboard_overview.png)
*Automated Dashboard with AI Insights*

![Campaigns](assets/campaigns_wow.png)
*WoW Campaign Performance Tracking*

---

## 1. Meta Ads Sync + Decision Engine (`meta_ads_sync_v2.gs`)
This script pulls data from Meta's API into Google Sheets. It includes a built-in decision engine that generates action items, creative scoring, and fatigue detection.

![Creative Scoring](assets/meta_creative_scoring.png)
*Creative Scoring and Fatigue Detection Tab*

### Features
- **Dashboard Construction**: Aggregates top-level numbers and flags issues (e.g., Roas dropped, CPC spikes).
- **Creative & Ad Set Deep Dives**: Tracks performance per ad set and tracks creative engagement metrics (Thumbstop, Hold Rate).
- **AI Integration**: Optionally plug in a Claude or Gemini API key for advanced AI insights on your ad performance.

![Meta KPIs](assets/meta_deep_dives_kpi.png)
*Detailed Meta Metric Analysis*

![Meta Funnel](assets/meta_deep_dives_funnel.png)
*Visualized Meta Shopping Funnel*

### Setup Instructions
1. Create a **New Google Sheet**.
2. Go to **Extensions > Apps Script** from the top menu.
3. Replace the default `.gs` file content by copy-pasting the full code from `meta_ads_sync_v2.gs`.
4. Update the `CONFIG` object at the top of the file:
   - `ACCESS_TOKEN`: The access token for your Meta Developer app.
   - `AD_ACCOUNT_ID`: Your Meta Ad Account ID (format: `act_XXXXXXXXX`).
   - *Optional:* Tune `THRESHOLDS` (like `TARGET_ROAS`) for your business.
5. Save the file.
6. In the Apps Script editor, select the function `setupTriggers` from the dropdown and hit **Run**.
7. Authorize access when prompted. This will automatically schedule the script to refresh daily.

---

## 2. Google Ads Sync - PMax Enhanced (`google_ads_sync.js`)
This script executes inside Google Ads and pushes 4 specialized data sets to the *same* Google Sheet:
- `google_ads_daily` — Campaign-level daily performance
- `google_ads_assets` — Asset group performance (PMax breakdowns)
- `google_ads_products` — Shopping product performance (what is actually selling)
- `google_ads_search` — Search term insights (queries triggering your PMax ads)

![Google Ads Deep Dives](assets/google_deep_dives.png)
*Cross-channel Google Ads and PMax Insights*

### Setup Instructions
1. Open your **Google Ads Account**.
2. Go to **Tools & Settings > Bulk Actions > Scripts**.
3. Click the **+** button to create a **New Script**.
4. Paste the full code from `google_ads_sync.js`.
5. Update `CONFIG.SPREADSHEET_URL` with the full URL of the Google Sheet you created in the Meta setup step.
6. Click **Authorize** to let the script access your Google Ads account.
7. Click **Run** manually the first time to generate the sheets and data.
8. Go back to the Scripts overview page and schedule the script to run **Daily**.

---

## The Output
Once both scripts are set up and scheduled, you will have a single Google Sheet serving as an end-to-end mission control. It updates continuously, highlighting exactly what to scale, what to kill, and where your creative and funnel opportunities lie.
