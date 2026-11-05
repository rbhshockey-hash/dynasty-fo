# Dynasty Front Office

A data-driven front-office system for a 12-team superflex dynasty fantasy football league on Sleeper.
The goal is a **perennial contender**: consistent playoff appearances and real title equity every year,
not a one-season spike.

This repo holds:

- **`apps-script/`**: the full data pipeline (Google Apps Script). It runs inside the Google Sheet, pulls from
  the Sleeper API, and acceptance-tests every stage before anything downstream may use it.
- **`docs/`**: a read-only website (GitHub Pages) showing the tested tables: pipeline health, rosters,
  pick ledger and contention record.

**Live site:** https://rbhshockey-hash.github.io/dynasty-fo/

---

## How it works

```
Sleeper API ──► Apps Script (ingest + tests) ──► Google Sheet (tested tables)
                                                      │
                                   DA_DataEndpoint.gs (read-only JSON, whitelisted tabs)
                                                      │
                                              GitHub Pages site (docs/)
```

### Principles

1. **Store evidence, derive conclusions.** `RAW_*` tabs hold source records exactly as Sleeper returns
   them. Anything interpretive is derived later, with uncertainty attached.
2. **Every stage has an acceptance gate.** Each step writes a `*_TESTS` tab with pre-registered thresholds
   and a gate: `PASS`, `PASS_RESTRICTED` (usable, with named limits) or `FAIL` (downstream refuses to build).
3. **Two currencies, never fused.** Realized production (league-specific Sleeper data) drives every strength
   measure. Consensus market value (FantasyCalc) is admitted **only** into the decision layer, as one scoped
   exception, and is **never published on this site**.
4. **Thin evidence is flagged, not hidden.** Sample-size limits (e.g. only two completed seasons) are
   stated wherever a conclusion depends on them.
5. **The site computes nothing.** It renders what is in the sheet, so the site and any analysis of the sheet
   always see the same numbers.

---

## Pipeline

| Menu | Script | What it does | Gate tab |
|---|---|---|---|
| 3 / 4 | `G9_Ingest.gs` | SAMPLE shape test (tests only), then FULL ingest of every rostered player-week, **bench included** | `G9_TESTS` |
| 3b | `G9_Ingest.gs` | Weekly UPDATE: new weeks only, plus re-pulls the last 2 weeks for stat corrections. Completed seasons are sealed. New weeks are tested *before* being written | `G9_TESTS` |
| 5 | `G5_PlayerMeta.gs` | Player metadata: name, position, eligibility, NFL team, birth date, experience | `G5_TESTS` |
| 5b | `RS_RosterSnapshot.gs` | Current rosters with IR/taxi status | `RS_TESTS` |
| 5c | `UV_Universe.gs` | Playoff brackets (contention outcomes), draft picks and ownership, free-agent pool | `UV_TESTS` |
| 5d | `MV_MarketValues.gs` | FantasyCalc market values: append-only daily history, **decision layer only** | `MV_TESTS` |
| 6 | `S0_Stats.gs` | Weekly NFL stats, scoring audit (reproduces Sleeper's points), bye-week classification | `S0_TESTS` |
| 7 | `C1_PlayerWeeks.gs` | Canonical player-week dataset (join, no re-derivation) | `C1D_TESTS` |
| n/a | `DA_DataEndpoint.gs` | Read-only JSON for the site (whitelisted tabs only) | n/a |

**Weekly run order:** 3b → 5 → 5b → 5c → 5d → 6 → 7. Each is a separate run because Apps Script limits
one execution to 6 minutes. A run log is kept in `RUN_LOG`.

### System architecture (beyond the data layer)

| Layer | Purpose | Status |
|---|---|---|
| C1 Asset | Per-player forward weekly production distribution, with evidence strength | Data built; method tests pending |
| C2 Lineup | Lineup-strength distributions for your team and all 11 opponents | Pending |
| A1 Strength | Field-relative strength and per-slot weakness | Pending |
| A2 Phase | Standing, trajectory, window | Pending |
| D0–D4 Decisions | Posture, trade evaluation, acquisition search, prioritization | Pending |

---

## Setting up the site's data feed

1. In the Google Sheet: **Extensions → Apps Script**, add `DA_DataEndpoint.gs`, and save.
2. **Deploy → New deployment → Web app**. Set Execute as: **Me**, Who has access: **Anyone**. Copy the `/exec` URL.
3. Put the URL in `docs/config.js` as `DATA_URL` and commit. After any later change to the endpoint,
   use **Manage deployments → Edit → New version**. That keeps the same URL.

The endpoint serves only: gate statuses and run log, current rosters with metadata, pick ownership, and
completed-season playoff results. Market values, raw stats, raw player-weeks and config are never exposed.

## Using the scripts in your own sheet

Create a Google Sheet, open **Extensions → Apps Script**, add every file from `apps-script/` (same names),
save, and reload the sheet. A **Dynasty FO** menu appears. Run **1. Set up tabs**, enter your Sleeper league
IDs in `CONFIG` (as text), then run **3 → 4 → 5 → 5b → 5c → 5d → 6 → 7**.
