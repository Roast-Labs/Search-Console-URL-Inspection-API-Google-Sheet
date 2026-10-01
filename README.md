# GSC Index Status Checker

A Google Sheets + Apps Script template that checks index status for a set of URLs on a schedule using the Search Console **URL Inspection API**, stores every result, and reports on it with pivots and charts.

This is **not** the Google Indexing API. The Indexing API only accepts `JobPosting` and `BroadcastEvent` submissions, and its metadata endpoint reports what you submitted — not whether Google has indexed a page.

The endpoint this tool calls:

```
POST https://searchconsole.googleapis.com/v1/urlInspection/index:inspect
```

One URL per request. No batch endpoint exists. Inspection uses `UrlFetchApp` + `ScriptApp.getOAuthToken()` with scope `https://www.googleapis.com/auth/webmasters.readonly`. Each user authorises with their own Google account and links **their own** Cloud project.

---

## Use it (no Git, no clasp)

GitHub is the source code. Users should never install from here.

Read more about how to use this tool - https://weareroast.com/resources/tools/google-indexing-api-tool/

**[Make a copy of the template](https://docs.google.com/spreadsheets/d/1ZE47UnzW0toJzh8H9iPAiGGjTKuUK1EiszVji48zGBA/edit?gid=0#gid=0)**

Then:

1. Create or pick **your own** Google Cloud project (the template’s Cloud project does not copy). Enable **Google Search Console API**, set the OAuth consent screen, then in this copy: **Extensions → Apps Script → Project Settings → Change project** and paste **your** project number. Full steps are under [Switch off Default GCP](#switch-off-default-gcp).
2. Open **Index Checker → Setup**.
3. Authorise your Google account when asked.
4. Pick the Search Console property.
5. Import a sitemap or paste URLs.
6. Done. The daily trigger is created for you.

That `/copy` URL looks like:

```
https://docs.google.com/spreadsheets/d/FILE_ID/copy
```

Share the template as **Anyone with the link → Viewer**. `/copy` puts a private copy in their Drive, with the script already attached. They never see this repo.

Until that link is published, there is no one-click install — Apps Script cannot be installed from GitHub the way a Chrome extension can.

---

## Quotas (verified 2026-08-21)

Re-check before treating these as gospel. Google changes them without much notice.

**URL Inspection API** — [official limits](https://developers.google.com/webmaster-tools/limits)

| Limit | Value |
|---|---|
| Inspections per property per day | 2,000 |
| Inspections per property per minute | 600 |
| Per-project (usually irrelevant) | 10,000,000 / day, 15,000 / min |
| Quota reset | Pacific time midnight |

`DAILY_CAP` defaults to **1,900** to leave headroom.

**Apps Script** — [official quotas](https://developers.google.com/apps-script/guides/services/quotas)

| Limit | Consumer | Workspace |
|---|---|---|
| Script runtime per execution | 6 min | 6 min |
| Trigger total runtime per day | ~90 min | ~6 hr |
| UrlFetch calls per day | ~20,000 | ~100,000 |
| Emails per day | ~100 | ~1,500 |
| Cells per spreadsheet | 10 million | 10 million |

The trigger runtime cap is the real constraint on consumer accounts: 90 minutes of continuation triggers per day limits practical throughput well below 2,000 URLs. Setup warns if the account looks like gmail.com.

`mobileUsabilityResult` is **deprecated** on the inspect schema. The sheet still captures Mobile Verdict if present; do not use it in client reporting. The Search Console Mobile Usability report was retired on 1 December 2023.

---

## Publish the template (once, as the maintainer)

This is the only “install” work. After this, users get a button.

1. Create a blank Google Sheet.
2. **Extensions → Apps Script**. Copy the **Script ID** from Project Settings.
3. Install [clasp](https://github.com/google/clasp), then:

```bash
cp .clasp.json.example .clasp.json
# paste the Script ID
clasp push
```

4. Switch the script off **Default** GCP (required — see below).
5. Run `bootstrapTemplate`. Authorise when prompted.
6. **Share → Anyone with the link → Viewer.**
7. The user link is `https://docs.google.com/spreadsheets/d/FILE_ID/copy` — put it in the button above.

Do not share the template as Editor. Viewers who hit `/copy` get their own sheet and their own script copy; they authorise as themselves.

A Workspace Marketplace add-on would be even easier for users (“Install” from Google), but it needs OAuth verification and a Marketplace listing. The `/copy` link is the easy path for a public GitHub project.

### Switch off Default GCP

`Default` is a hidden Cloud project. You cannot enable APIs on it from Cloud Console. Each `/copy` gets a **new** Apps Script project on Default GCP — the template owner’s Cloud project does **not** copy. Every user creates or links **their own** standard Cloud project.

1. Open [Google Cloud Console](https://console.cloud.google.com/) with the **same Google account**.
2. Create a project (or pick an existing one). Name it e.g. `gsc-index-checker`.
3. Copy the **Project number** (APIs & Services → Settings, or the project dashboard). That is digits, not the project ID string.
4. Enable **Google Search Console API**:
   [console.cloud.google.com/apis/library/searchconsole.googleapis.com](https://console.cloud.google.com/apis/library/searchconsole.googleapis.com)
5. **APIs & Services → OAuth consent screen**:
   - User type: Internal (Workspace) or External (gmail).
   - App name, user support email.
   - If External: Publishing status **Testing**, add your email under Test users.
6. In the **Apps Script** editor: Project Settings (gear) → **Google Cloud Platform Project → Change project** → paste the **project number**.
7. Re-run **Index Checker → Setup** and authorise again (the Cloud project change invalidates the old token).

Until you do this, `sites.list` and URL Inspection will 403.

---

## Install from this repo e.g. run the code yourselve rather than copy the template

For changing the code, not for handing to clients.

1. Create a blank Google Sheet → **Extensions → Apps Script** → copy the Script ID.
2. `cp .clasp.json.example .clasp.json` and paste that ID.
3. `clasp push`
4. Set timezone, enable the Search Console API, run `bootstrapTemplate`.
5. **Index Checker → Setup** in the sheet.

---

## Tabs

| Tab | Visibility | Purpose |
|---|---|---|
| `Settings` | Visible | Named-range backed config |
| `URLs` | Visible | Input list plus live credit calculator |
| `Data` | Visible | Append-only result store |
| `Run Log` | Visible | One row per run |
| `Pivots` | Visible | Pivot tables over `Data` |
| `Dashboard` | Visible | Scorecards and charts |
| `Glossary` | Visible | Status definitions and API limits |
| `_Queue` | Hidden | Resumable batch state |
| `_Quota` | Hidden | Daily quota ledger (Pacific date) |
| `_Import` | Hidden | Staging area for sitemap wizard |
| `_Deferred` | Hidden | 429 / trim leftovers requeued for tomorrow |

The script reads Settings **by named range**, never by A1.

---

## How a run works

Apps Script kills executions at 6 minutes, so a full daily cap cannot finish in one shot.

1. `buildQueue()` takes `Active = TRUE` URLs whose `Next Due` is today or earlier, sorts by Priority then oldest due first, truncates to remaining quota, writes `_Queue`, and sets `Reserved`.
2. `processQueue()` takes a document lock (30s timeout), reads a cursor from Document Properties, inspects the next `BATCH_SIZE` URLs in `UrlFetchApp.fetchAll()` chunks of 10, sleeps ~1.1s between chunks to stay under 600 QPM, appends every result to `Data`, and bumps `_Quota`.
3. After 4.5 minutes it persists the cursor and creates a one-off trigger for +1 minute.
4. When the queue empties it writes `Run Log`, updates `Last Checked` / `Next Due` / `Last Verdict` on `URLs`, clears `Reserved`, deletes continuation triggers, and emails if configured.

### Retry policy

| Response | Action |
|---|---|
| 429 | Exponential backoff, 3 attempts, then requeue for tomorrow at Priority 1 |
| 5xx | Exponential backoff, 3 attempts, then log as failed |
| 403 | **Abort the entire run** — almost always no property access |
| 404 from the API | Logged as a valid inspection result, not a failure |
| URL outside the property | Logged, run continues (no inspect call) |

### Quota ledger

Keyed on **Pacific date**, not local date. A 03:00 London run still draws on the previous Pacific day's allowance.

```
Remaining = DAILY_CAP − Scheduled Used − Ad-hoc Used − Reserved
```

During the day, `Reserved` is kept equal to the number of scheduled URLs still due, so ad-hoc runs cannot silently spend tonight's budget. Opening the sheet or **View Quota** refreshes that reservation.

---

## Ad-hoc runs

**Index Checker → Run Ad-hoc** opens a sidebar. Paste URLs or use the current `URLs` selection.

If the request fits in `Remaining`, it runs immediately against **Ad-hoc Used**.

If not, a modal shows the explicit numbers and three actions:

| Button | Behaviour |
|---|---|
| **Trim to fit** | Runs what fits, queues the rest for tomorrow at Priority 1 |
| **Borrow from tonight** | Second confirmation. Releases reservation, runs the request, shrinks tonight's scheduled run, flags `Borrowed From Scheduled` on the Run Log. Hidden when `ADHOC_BORROW_ALLOWED` is FALSE |
| **Queue for tomorrow** | No spend now |

---

## Sitemap import

**Index Checker → Import from Sitemap**

- **Search Console** — `GET /webmasters/v3/sites/{siteUrl}/sitemaps`, then fetch each feed.
- **Direct URL** — paste a sitemap (needed for staging / unsubmitted files).

Parser: sitemap index recursion (depth cap 3), `.xml.gz` via `Utilities.ungzip`, `<lastmod>` captured, malformed XML fails with the sitemap URL named (no half-import). Filter with include/exclude regex before commit. `Section` is derived from the first path segment and can be remapped. Orphaned URLs are flagged `Orphaned = TRUE`, never deleted.

Priority from `lastmod`: ≤7 days → 1, ≤30 days → 2, older or absent → 3.

---

## Menu

| Item | What it does |
|---|---|
| Setup | Property picker, consumer-account warning, daily trigger, smoke test |
| Run Now | Build queue + process |
| Run Ad-hoc | Sidebar |
| Import from Sitemap | Wizard |
| Build Queue Only | Write `_Queue` without inspecting |
| View Quota | Today's Pacific ledger |
| Reset Triggers | Delete all project triggers; recreate the daily one if setup is complete |
| Purge Old Data | Delete `Data` rows older than `DATA_RETENTION_DAYS` |
| Rebuild sheet structure | Re-run `bootstrapTemplate` (does not wipe Data / URLs / Run Log / quota history) |

Setup is guarded by a `SETUP_COMPLETE` document property so reopening the sheet does not retrigger the wizard.

---

## Test cases

| Case | Expected |
|---|---|
| Empty `URLs` sheet | Queue builds with 0 rows, run logs “Nothing due.” |
| URL outside the property | Data row with a clear error, run continues |
| No property access (403) | Run aborts immediately with a readable message |
| 2,500 due vs 1,900 cap | Truncates; oldest overdue + highest priority kept |
| Execution hits 4.5 minutes | Continuation resumes at the stored cursor |
| Manual run during a scheduled run | Lock refuses the second caller |
| Ad-hoc over remaining | Trim / Borrow / Queue tomorrow all work |
| Gzipped sitemap index three levels deep | Recurses to depth 3 |
| 50,000-URL sitemap | Filter step before commit; no dump onto `URLs` |
| URL removed from sitemap | `Orphaned = TRUE`, row kept |
| Run spanning Pacific midnight | Later inspections increment the new PT date |
| Retention purge with 200 days of data | Rows older than `DATA_RETENTION_DAYS` removed |

---

## Project layout

```
src/
  appsscript.json
  00_Constants.gs … 15_Bootstrap.gs
  Setup.html, AdhocSidebar.html, QuotaWarning.html,
  BorrowConfirm.html, SitemapWizard.html, QuotaView.html, Styles.html
```

Runtime helpers are named with a trailing `_`. Those names are **not** callable from `google.script.run`; public wrappers live in `13_Menu.gs`.
