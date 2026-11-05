/***** Market values (FantasyCalc) — economic shell input + MV acceptance tests ******************************
 * Add as a SEVENTH file in the same Apps Script project (reuses helpers from G9_Ingest.gs and S0_Stats.gs).
 * Menu item "5d. Pull market values (FantasyCalc) + tests" (G9_Ingest.gs onOpen). Run after 5b (roster snapshot).
 *
 * DECISION (2026-11-05): the league-data-only rule has ONE scoped exception — consensus market value from
 * FantasyCalc, for the ECONOMIC SHELL ONLY (decision layer D0–D4: per-asset value, liquidity, valuation
 * uncertainty). Rationale: market value IS consensus by definition; no league-only source can price players
 * (a handful of trades/season). Boundaries, enforced by convention and stated in every consumer:
 *   - Nothing in this tab may enter C1, C2, A1 or A2 (two-currency separation; Foundation Spec v0.3 withdrew
 *     market economics from C1–A2). Redraft_Value is stored because the API returns it, but C1 does NOT use it
 *     (C1 Test Plan D6).
 *   - Values are a consensus of many leagues, not this league's market, and are NOT adjusted for this league's
 *     TE premium (+0.5/rec). Both limits are carried as named uncertainty, never corrected silently.
 *
 * Request parameters are DERIVED from RAW_LEAGUE (current season), never hand-set:
 *   isDynasty=true · numTeams = total_rosters · numQbs = 2 if SUPER_FLEX or >1 QB slot else 1 · ppr = scoring rec.
 *
 * Writes RAW_MARKET_VALUES as an APPEND-ONLY history: one snapshot per calendar day (a re-run on the same day
 * replaces that day). FantasyCalc exposes current values only, so this tab is the only way a value history
 * (trend, timing, leakage-free tests later) will ever exist. Raw_JSON keeps every field the API returned.
 *
 * Pre-registered thresholds:
 *   MV-1 response is a non-empty array of entries carrying a numeric value → else FAIL (nothing written)
 *   MV-2 coverage of current rostered QB/RB/WR/TE (RAW_ROSTER_SNAPSHOT) by Sleeper ID:
 *        PASS ≥90% · WARN 75–90% · FAIL <75%. Uncovered players are listed (they carry "no market value").
 *   MV-3 duplicate Sleeper IDs within the snapshot → FAIL
 *   MV-4 picks, uncertainty fields, history depth, scope limits → INFO
 *   Gate: FAIL if any FAIL · PASS_RESTRICTED if any WARN · else PASS
 ***************************************************************************************************/
const MV = { VALUES: 'RAW_MARKET_VALUES', TESTS: 'MV_TESTS' };
const MV_API = 'https://api.fantasycalc.com/values/current';
const MV_HEADERS = ['Snapshot_Date', 'Sleeper_ID', 'FC_ID', 'Name', 'Position', 'Is_Pick', 'NFL_Team', 'Value',
  'Overall_Rank', 'Position_Rank', 'Trend_30Day', 'Moving_SD', 'Moving_SD_Pct', 'Trade_Frequency', 'Redraft_Value',
  'Request_Params', 'Raw_JSON', 'Fetched_At'];

function runMarketValues() { run_('runMarketValues', 'MV', () => marketValues_()); }

function marketValues_() {
  const cfg = readConfig_();
  const cur = cfg.leagues[cfg.leagues.length - 1];

  // ---- derive request parameters from RAW_LEAGUE (current season) ----
  const lv = sheet_(T.LEAGUE).getDataRange().getValues(), lh = lv[0].map(String);
  const row = lv.find((r, i) => i > 0 && Number(r[lh.indexOf('Config_Season')]) === cur.season);
  if (!row) throw new Error(`RAW_LEAGUE has no row for ${cur.season}. Run G9 (step 2 or 3b) first.`);
  const rp = JSON.parse(row[lh.indexOf('Roster_Positions_JSON')] || '[]');
  const sc = JSON.parse(row[lh.indexOf('Scoring_Settings_JSON')] || '{}');
  const numTeams = Number(row[lh.indexOf('Total_Rosters')]);
  const numQbs = (rp.indexOf('SUPER_FLEX') >= 0 || rp.filter(p => p === 'QB').length > 1) ? 2 : 1;
  const ppr = Number(sc.rec) || 0;
  const tePrem = Number(sc.bonus_rec_te) || 0;
  const params = `isDynasty=true&numQbs=${numQbs}&numTeams=${numTeams}&ppr=${ppr}`;

  // ---- fetch ----
  const resp = UrlFetchApp.fetch(`${MV_API}?${params}`, { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) throw new Error(`FantasyCalc HTTP ${resp.getResponseCode()} for ${MV_API}?${params}`);
  let data;
  try { data = JSON.parse(resp.getContentText()); } catch (e) { throw new Error('FantasyCalc returned non-JSON.'); }

  const out = [], at = new Date(), now = new Date();
  const add = (id, q, cond, actual, status) => out.push(['MV', cur.season, id, q, cond, actual, status, at]);
  const day = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const num = x => (x == null || x === '' || isNaN(Number(x)) ? '' : Number(x));

  const entries = Array.isArray(data) ? data.filter(e => e && e.player && num(e.value) !== '') : [];
  add('MV-0', 'Request parameters derived from league settings', 'informational',
    `${params} (league: ${numTeams} teams · ${numQbs === 2 ? 'superflex' : '1QB'} · ${ppr} PPR)`, 'INFO');
  add('MV-1', 'FantasyCalc returned usable values', 'non-empty array, entries carry a numeric value (FAIL: nothing written)',
    `${Array.isArray(data) ? data.length : 'not an array'} entries · ${entries.length} usable`, entries.length ? 'PASS' : 'FAIL');
  if (!entries.length) return finishMv_(out, cur, 0, 'MV-1 failed — RAW_MARKET_VALUES unchanged');

  // ---- shape rows ----
  const rows = entries.map(e => {
    const p = e.player || {};
    const pos = String(p.position || '');
    return [day, p.sleeperId != null ? String(p.sleeperId) : '', p.id != null ? String(p.id) : '', p.name || '', pos,
      pos === 'PICK', p.maybeTeam || '', num(e.value), num(e.overallRank), num(e.positionRank), num(e.trend30Day),
      num(e.maybeMovingStandardDeviation), num(e.maybeMovingStandardDeviationPerc), num(e.maybeTradeFrequency),
      num(e.redraftValue), params, JSON.stringify(e).slice(0, 49000), now];
  });

  // ---- tests on the new snapshot ----
  const ids = rows.filter(r => r[1]).map(r => r[1]);
  const dup = ids.length - new Set(ids).size;
  add('MV-3', 'No duplicate Sleeper IDs in the snapshot', '0 (FAIL)', `${dup}`, dup ? 'FAIL' : 'PASS');

  const byId = {};
  rows.forEach(r => { if (r[1]) byId[r[1]] = r; });
  const snap = SpreadsheetApp.getActive().getSheetByName(RS.SNAP);
  const pm = SpreadsheetApp.getActive().getSheetByName(G5.META);
  if (!snap || snap.getLastRow() < 2 || !pm || pm.getLastRow() < 2) {
    add('MV-2', 'Coverage of current rostered QB/RB/WR/TE', 'needs RAW_ROSTER_SNAPSHOT + RAW_PLAYER_META',
      'missing — run 5 (G5) and 5b (snapshot) first', 'FAIL');
  } else {
    const mv = pm.getDataRange().getValues(), mh = mv[0].map(String);
    const name = {}, elig = {};
    for (let r = 1; r < mv.length; r++) {
      const pid = String(mv[r][mh.indexOf('Player_ID')]).trim();
      name[pid] = String(mv[r][mh.indexOf('Full_Name')] || '');
      let fp = [];
      try { fp = JSON.parse(mv[r][mh.indexOf('Fantasy_Positions_JSON')] || '[]'); } catch (e) { fp = []; }
      elig[pid] = fp.some(x => LINEUP_POS.indexOf(x) >= 0);
    }
    const sv = snap.getDataRange().getValues(), sh = sv[0].map(String);
    const rost = [...new Set(sv.slice(1).map(r => String(r[sh.indexOf('Player_ID')]).trim()))].filter(p => elig[p]);
    const missing = rost.filter(p => !byId[p]);
    const cov = rost.length ? (rost.length - missing.length) / rost.length : 0;
    add('MV-2', 'Coverage of current rostered QB/RB/WR/TE by Sleeper ID', 'PASS ≥90% · WARN 75–90% · FAIL <75%',
      `${rost.length - missing.length}/${rost.length} (${(100 * cov).toFixed(1)}%)` +
        (missing.length ? ` · no market value: ${missing.slice(0, 25).map(p => name[p] || p).join(', ')}` : ''),
      cov >= 0.9 ? 'PASS' : cov >= 0.75 ? 'WARN' : 'FAIL');
  }

  const picks = rows.filter(r => r[5]);
  add('MV-4a', 'Draft-pick entries returned (names as FantasyCalc labels them)', 'informational — mapping to PICK_OWNERSHIP is a decision-layer step',
    `${picks.length}` + (picks.length ? `: ${picks.slice(0, 12).map(r => `${r[3]} = ${r[7]}`).join(' | ')}` : ''), 'INFO');
  const pl = rows.filter(r => !r[5]);
  const pct = f => pl.length ? `${(100 * pl.filter(r => r[f] !== '').length / pl.length).toFixed(0)}%` : 'n/a';
  add('MV-4b', 'Economic-shell primitive coverage (players)', 'informational: value · liquidity · valuation uncertainty',
    `value 100% · trade frequency (liquidity) ${pct(13)} · moving SD (valuation uncertainty) ${pct(11)} · 30-day trend ${pct(10)}`, 'INFO');
  add('MV-4c', 'Scope limits of this exception', 'informational — carried as named uncertainty by every consumer',
    `consensus across many leagues, not this league's market · ` +
      (tePrem ? `league has TE premium (+${tePrem}/rec); values are NOT TE-premium-adjusted · ` : '') +
      `decision layer only — never an input to C1/C2/A1/A2; Redraft_Value stored but unused (D6)`, 'INFO');

  // ---- write (append-only history; same-day re-run replaces that day) ----
  const fails = out.filter(t => t[6] === 'FAIL').map(t => t[2]);
  let how = 'not written (gate FAIL)';
  if (!fails.length) how = writeMvSnapshot_(rows, day);
  const sv = sheet_(MV.VALUES, true);
  const days = sv.getLastRow() > 1 ? new Set(sv.getRange(2, 1, sv.getLastRow() - 1, 1).getDisplayValues().map(r => r[0])) : new Set();
  add('MV-4d', 'Market-value history depth', 'informational (FantasyCalc serves current values only; history accrues here)',
    `${days.size} snapshot day(s): ${[...days].sort().slice(-5).join(', ')}`, 'INFO');
  return finishMv_(out, cur, fails.length ? 0 : rows.length, how);
}

function writeMvSnapshot_(rows, day) {
  const sh = sheet_(MV.VALUES, true);
  const n = MV_HEADERS.length;
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, n).setValues([MV_HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  const last = sh.getLastRow();
  let start = last + 1, replaced = 0;
  if (last > 1) {
    const d = sh.getRange(2, 1, last - 1, 1).getDisplayValues().map(r => r[0]);
    const first = d.indexOf(day);
    if (first >= 0) {
      if (d.slice(first).every(x => x === day)) {           // today's rows are the tail: clear them in place
        sh.getRange(2 + first, 1, d.length - first, n).clearContent();
        start = 2 + first; replaced = d.length - first;
      } else {                                               // unexpected order: filter + rewrite
        const keep = sh.getRange(2, 1, last - 1, n).getValues().filter((r, i) => d[i] !== day);
        writeTableChunked_(MV.VALUES, MV_HEADERS, keep, [1, 2, 3, 16, 17]);
        start = keep.length + 2; replaced = d.length - keep.length;
      }
    }
  }
  [1, 2, 3, 16, 17].forEach(c => sh.getRange(start, c, rows.length, 1).setNumberFormat('@')); // date/IDs/JSON text
  for (let i = 0; i < rows.length; i += 2000) {
    const part = rows.slice(i, i + 2000);
    sh.getRange(start + i, 1, part.length, n).setValues(part);
  }
  return `appended ${rows.length} rows for ${day}` + (replaced ? ` (replaced ${replaced} earlier rows from today)` : '');
}

function finishMv_(out, cur, rowsN, how) {
  const at = new Date();
  const fails = out.filter(t => t[6] === 'FAIL').map(t => t[2]);
  const warns = out.filter(t => t[6] === 'WARN').map(t => t[2]);
  const gate = fails.length ? 'FAIL' : warns.length ? 'PASS_RESTRICTED' : 'PASS';
  out.push(['MV', cur.season, 'MV-GATE', 'Market values acceptable → decision layer (D0–D4) may consume them',
    'FAIL if any FAIL (nothing written) · PASS_RESTRICTED if any WARN · else PASS',
    fails.length ? `blocked by ${fails.join(', ')}` : warns.length ? `restricted by ${warns.join(', ')}` : 'all conditions met', gate, at]);
  const th = sheet_(MV.TESTS, true);
  th.clearContents();
  th.getRange(1, 1, 1, 8).setValues([['Run', 'Season', 'Test_ID', 'Question', 'Pass_Condition', 'Actual', 'Status', 'Run_At']]).setFontWeight('bold');
  th.getRange(2, 6, out.length, 1).setNumberFormat('@');
  th.getRange(2, 1, out.length, 8).setValues(out);
  th.setFrozenRows(1);
  return { seasons: String(cur.season), calls: 1, rows: rowsN, status: gate === 'FAIL' ? 'TESTS FAILED' : 'OK',
           notes: `MV-GATE ${gate} | ${how}` };
}
