/***** Dynasty Front Office — G9 ingestion + acceptance tests (Sleeper) ***************************
 * Tabs: README, CONFIG, RAW_LEAGUE, RAW_ROSTERS, RAW_G9, G9_TESTS, RUN_LOG (+ optional REF_RAW_STARTERS)
 * Order: 1 Set up tabs → enter League IDs in CONFIG → 3 Run G9 SAMPLE → review G9_TESTS
 *        → 4 Run G9 FULL (refuses to run unless SAMPLE G9-ACCEPT = PASS for every configured season)
 *        → thereafter 3b Run G9 UPDATE weekly (new weeks only; FULL becomes an occasional rebuild)
 * Store evidence, derive conclusions: RAW_* tabs hold source records only; tests live in G9_TESTS.
 *
 * v2 (2026-11-05) — incremental ingestion:
 *   UPDATE fetches only (a) completed weeks not yet in RAW_G9 and (b) the last Refresh_Weeks completed weeks of the
 *   CURRENT season (stat-correction window; also re-pulls a week ingested before it fully finished). Completed
 *   seasons are sealed: never re-fetched. Fetched weeks are tested BEFORE they are written; if any G9-ACCEPT fails,
 *   RAW_G9 is left untouched. Rows are appended in place (no full-table rewrite) when the new weeks sit at the end
 *   of the sorted table; otherwise a sorted full rewrite is done and logged.
 *   Playoff-only updates are tested with the scope FULL already applies to playoff weeks (Q1b, Q3d, Q4p).
 ***************************************************************************************************/
const API = 'https://api.sleeper.app/v1';
const T = { README: 'README', CONFIG: 'CONFIG', LEAGUE: 'RAW_LEAGUE', ROSTERS: 'RAW_ROSTERS',
            G9: 'RAW_G9', TESTS: 'G9_TESTS', LOG: 'RUN_LOG', REF: 'REF_RAW_STARTERS' };
const TOL = 0.01;
const TEST_HEADERS = ['Mode', 'Season', 'Test_ID', 'Question', 'Pass_Condition', 'Actual', 'Status', 'Run_At'];
const LOG_HEADERS = ['Run_At', 'Function', 'Mode', 'Seasons', 'API_Calls', 'Rows_Written', 'Status', 'Notes'];
const G9_HEADERS = ['Season', 'League_ID', 'Week', 'Is_Regular_Season', 'Roster_ID', 'Matchup_ID', 'Team_Points',
  'Player_ID', 'Points', 'In_Players', 'In_Players_Points', 'Is_Starter', 'Starter_Slot_Index', 'Ingest_Mode', 'Fetched_At'];

/* ============================== menu + public entry points ============================== */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Dynasty FO')
    .addItem('1. Set up tabs', 'setupTabs')
    .addItem('2. Pull league + rosters only', 'pullLeague')
    .addItem('3. Run G9 SAMPLE + tests', 'runG9Sample')
    .addItem('3b. Run G9 UPDATE (new weeks only) + tests', 'runG9Update')
    .addItem('4. Run G9 FULL + tests (gated, full rebuild)', 'runG9Full')
    .addItem('5. Run G5 player meta + tests', 'runG5')
    .addItem('5b. Pull current roster snapshot + tests', 'runRosterSnapshot')
    .addItem('5c. Pull universe: brackets, picks, free agents + tests', 'runUniverse')
    .addItem('5d. Pull market values (FantasyCalc) + tests', 'runMarketValues')
    .addItem('6. Run Step 0 stats + scoring audit', 'runStep0')
    .addItem('7. Build C1_PLAYER_WEEKS (Step 2)', 'runStep2')
    .addToUi();
}

function setupTabs() {
  const ss = SpreadsheetApp.getActive();
  [T.README, T.CONFIG, T.LEAGUE, T.ROSTERS, T.G9, T.TESTS, T.LOG].forEach(n => sheet_(n, true));
  const cfg = ss.getSheetByName(T.CONFIG);
  if (cfg.getLastRow() === 0) {
    cfg.getRange('B:B').setNumberFormat('@');   // League IDs are 18–19 digits: must be stored as TEXT
    cfg.getRange('F:F').setNumberFormat('@');   // keeps "2024,2025" from being parsed as a number
    cfg.getRange(1, 1, 4, 3).setValues([['Season', 'League_ID', 'Include'], [2024, '', true], [2025, '', true], [2026, '', true]]);
    cfg.getRange(1, 5, 6, 2).setValues([
      ['Parameter', 'Value'],
      ['Max_Week', '18'],
      ['Ref_Seasons', '2024,2025'],
      ['Expected_Ref_Starters', '3360'],
      ['Refresh_Weeks', '2'],
      ['Note', 'Column B is pre-formatted as text. Paste IDs as plain text.']]);
    cfg.getRange('A1:F1').setFontWeight('bold');
    cfg.setFrozenRows(1);
  }
  header_(T.TESTS, TEST_HEADERS);
  header_(T.LOG, LOG_HEADERS);
  writeReadme_();
  ss.toast('Tabs ready. Enter League IDs in CONFIG, then run step 3.', 'Dynasty FO', 8);
}

function pullLeague() {
  run_('pullLeague', 'LEAGUE', () => {
    const cfg = readConfig_();
    const lp = pullLeagueCore_(cfg);
    return { seasons: seasonsStr_(cfg), calls: lp.calls, rows: lp.rows,
             notes: `NFL state: season ${lp.state.season}, week ${lp.state.week}` };
  });
}

function runG9Sample() { run_('runG9Sample', 'SAMPLE', () => ingest_('SAMPLE')); }

function runG9Full() {
  run_('runG9Full', 'FULL', () => { assertSampleGate_(readConfig_()); return ingest_('FULL'); });
}

function runG9Update() {
  run_('runG9Update', 'UPDATE', () => { assertSampleGate_(readConfig_()); return ingestUpdate_(); });
}

/* ============================== config ============================== */
function readConfig_() {
  const v = sheet_(T.CONFIG).getDataRange().getValues();
  const leagues = [], p = {};
  for (let r = 1; r < v.length; r++) {
    const season = v[r][0], id = v[r][1], inc = v[r][2];
    if (season !== '' && inc !== false && String(inc).toUpperCase() !== 'FALSE') {
      if (typeof id === 'number') {
        throw new Error(`CONFIG row ${r + 1}: League_ID was stored as a number and has lost precision. Re-enter it as text.`);
      }
      const lid = String(id || '').trim();
      if (lid) leagues.push({ season: Number(season), leagueId: lid });
    }
    if (v[r][4] !== '' && v[r][4] != null) p[String(v[r][4]).trim()] = v[r][5];
  }
  if (!leagues.length) throw new Error('CONFIG has no included seasons with a League_ID.');
  leagues.sort((a, b) => a.season - b.season);
  const rw = p.Refresh_Weeks;
  return {
    leagues,
    maxWeek: Number(p.Max_Week) || 18,
    refSeasons: String(p.Ref_Seasons || '').split(',').map(s => Number(s.trim())).filter(Boolean),
    expectedRef: Number(p.Expected_Ref_Starters) || null,
    refreshWeeks: (rw === undefined || rw === '') ? 2 : Math.max(0, Math.floor(Number(rw) || 0))
  };
}

/* ============================== league + rosters (source records) ============================== */
function pullLeagueCore_(cfg) {
  const urls = [API + '/state/nfl'];
  cfg.leagues.forEach(l => urls.push(`${API}/league/${l.leagueId}`, `${API}/league/${l.leagueId}/rosters`, `${API}/league/${l.leagueId}/users`));
  const res = fetchAllJson_(urls);
  const state = res[0] || {};
  const now = new Date();
  const metas = [], leagueRows = [], rosterRows = [];

  cfg.leagues.forEach((l, i) => {
    const lg = res[1 + i * 3], rosters = res[2 + i * 3] || [], users = res[3 + i * 3] || [];
    if (!lg) throw new Error(`League ${l.leagueId} (CONFIG season ${l.season}) returned null. Check the ID.`);
    const s = lg.settings || {};
    const m = {
      cfgSeason: l.season, leagueId: l.leagueId, apiSeason: Number(lg.season),
      totalRosters: Number(lg.total_rosters),
      prevId: lg.previous_league_id ? String(lg.previous_league_id) : '',
      playoffStart: Number(s.playoff_week_start) || null,
      rosterPositions: lg.roster_positions || [],
      reserveSlots: Number(s.reserve_slots) || 0, taxiSlots: Number(s.taxi_slots) || 0
    };
    metas.push(m);
    leagueRows.push([l.season, l.leagueId, String(lg.season), lg.name, lg.status, lg.total_rosters, m.prevId,
      s.playoff_week_start == null ? '' : s.playoff_week_start, s.last_scored_leg == null ? '' : s.last_scored_leg,
      JSON.stringify(lg.roster_positions || []), JSON.stringify(lg.scoring_settings || {}), JSON.stringify(s), now]);

    const uById = {};
    users.forEach(u => { uById[u.user_id] = u; });
    rosters.forEach(r => {
      const u = uById[r.owner_id] || {};
      rosterRows.push([l.season, l.leagueId, r.roster_id, r.owner_id ? String(r.owner_id) : '', u.display_name || '',
        (u.metadata && u.metadata.team_name) || u.display_name || '', now]);
    });
  });

  writeTable_(T.LEAGUE, ['Config_Season', 'League_ID', 'API_Season', 'League_Name', 'Status', 'Total_Rosters',
    'Previous_League_ID', 'Playoff_Week_Start', 'Last_Scored_Leg', 'Roster_Positions_JSON', 'Scoring_Settings_JSON',
    'Settings_JSON', 'Fetched_At'], leagueRows, [2, 3, 7]);
  writeTable_(T.ROSTERS, ['Season', 'League_ID', 'Roster_ID', 'Owner_ID', 'Display_Name', 'Team_Name', 'Fetched_At'],
    rosterRows, [2, 4]);
  return { metas, state, calls: urls.length, rows: leagueRows.length + rosterRows.length };
}

/* ============================== G9 ingestion ============================== */
function weekPlan_(m, state, maxWeek, mode) {
  let last = maxWeek;
  const stSeason = Number(state.season);
  if (stSeason === m.apiSeason) last = Math.min(maxWeek, Math.max(0, Number(state.week) - 1)); // completed weeks only
  else if (stSeason < m.apiSeason) last = 0;                                                  // season not started
  const lastReg = Math.min(last, (m.playoffStart || maxWeek + 1) - 1);
  if (mode === 'SAMPLE') return [...new Set([1, lastReg])].filter(w => w >= 1 && w <= last);
  const out = [];
  for (let w = 1; w <= last; w++) out.push(w);
  return out;
}

// Fetch matchup weeks and shape them into RAW_G9 rows + the per-roster-week records the tests need.
function fetchMatchups_(reqs, mode) {
  const data = fetchAllJson_(reqs.map(r => r.url));
  const now = new Date();
  const rows = [], rws = [], weeks = [];
  reqs.forEach((rq, i) => {
    const arr = Array.isArray(data[i]) ? data[i] : [];
    const isReg = rq.m.playoffStart ? rq.w < rq.m.playoffStart : true;
    weeks.push({ season: rq.m.cfgSeason, week: rq.w, isReg, n: arr.length, expected: rq.m.totalRosters, entries: arr });
    arr.forEach(e => {
      const players = (e.players || []).map(String);
      const starters = (e.starters || []).map(String);
      const pp = e.players_points || {};
      rws.push({ season: rq.m.cfgSeason, week: rq.w, isReg, rosterId: e.roster_id, matchupId: e.matchup_id,
        teamPoints: e.points, customPoints: e.custom_points, players, starters, sp: e.starters_points || [], pp });
      const ids = [...new Set(players.concat(Object.keys(pp)))];
      ids.forEach(pid => {
        const si = starters.indexOf(pid);
        rows.push([rq.m.cfgSeason, rq.m.leagueId, rq.w, isReg, e.roster_id,
          e.matchup_id == null ? '' : e.matchup_id, e.points == null ? '' : e.points,
          pid, (pid in pp) ? pp[pid] : '', players.indexOf(pid) >= 0, (pid in pp),
          si >= 0, si >= 0 ? si : '', mode, now]);
      });
    });
  });
  return { rows, rws, weeks };
}

function ingest_(mode) {
  const cfg = readConfig_();
  const lp = pullLeagueCore_(cfg);
  const reqs = [];
  lp.metas.forEach(m => weekPlan_(m, lp.state, cfg.maxWeek, mode).forEach(w =>
    reqs.push({ m, w, url: `${API}/league/${m.leagueId}/matchups/${w}` })));
  const b = fetchMatchups_(reqs, mode);

  // SAMPLE is a shape test only: it never touches RAW_G9 (it used to overwrite the full table with 6 weeks).
  if (mode !== 'SAMPLE') writeTable_(T.G9, G9_HEADERS, b.rows, [2, 8]);
  const tests = runTests_(mode, lp.metas, b.weeks, b.rws, b.rows, cfg);
  writeTests_(mode, tests);

  const verdict = tests.filter(t => t[2] === 'G9-ACCEPT').map(t => `${t[1]}:${t[6]}`).join(' ');
  const q5 = tests.filter(t => t[2] === 'Q5').map(t => `${t[1]}:${t[6]}`).join(' ');
  return { seasons: seasonsStr_(cfg), calls: lp.calls + reqs.length, rows: mode === 'SAMPLE' ? 0 : b.rows.length,
           status: /FAIL/.test(verdict) ? 'TESTS FAILED' : 'OK',
           notes: `G9-ACCEPT ${verdict} | Q5 ${q5} | ${reqs.length} matchup weeks` +
             (mode === 'SAMPLE' ? ' | tests only — RAW_G9 not modified' : '') };
}

/* ============================== G9 incremental UPDATE ============================== */
function ingestUpdate_() {
  const cfg = readConfig_();
  const sh = sheet_(T.G9);
  const last = sh.getLastRow();
  if (last < 2) throw new Error('RAW_G9 is empty. Run G9 FULL once before UPDATE.');
  const hdr = sh.getRange(1, 1, 1, G9_HEADERS.length).getValues()[0].map(String);
  if (hdr.join('|') !== G9_HEADERS.join('|')) throw new Error('RAW_G9 columns do not match the current schema. Run G9 FULL.');

  // Only Season + Week are read (cols 1 and 3): cheap regardless of table size.
  const sw = sh.getRange(2, 1, last - 1, 3).getValues().map(r => ({ s: Number(r[0]), w: Number(r[2]) }));
  const have = {};
  sw.forEach(x => { (have[x.s] = have[x.s] || {})[x.w] = true; });

  const lp = pullLeagueCore_(cfg);
  const curApi = Number(lp.state.season);
  const reqs = [], replace = new Set(), plan = [];
  lp.metas.forEach(m => {
    const done = have[m.cfgSeason] || {};
    const weeks = weekPlan_(m, lp.state, cfg.maxWeek, 'FULL');
    const missing = weeks.filter(w => !done[w]);
    const refresh = (m.apiSeason === curApi && cfg.refreshWeeks > 0)
      ? weeks.slice(-cfg.refreshWeeks).filter(w => done[w]) : [];
    refresh.forEach(w => replace.add(`${m.cfgSeason}|${w}`));
    missing.concat(refresh).sort((a, b) => a - b).forEach(w =>
      reqs.push({ m, w, url: `${API}/league/${m.leagueId}/matchups/${w}` }));
    plan.push(`${m.cfgSeason}: ` + (missing.length || refresh.length
      ? `${missing.length} new${missing.length ? ` (wk ${missing.join(',')})` : ''}, ` +
        `${refresh.length} refreshed${refresh.length ? ` (wk ${refresh.join(',')})` : ''}`
      : (m.apiSeason < curApi ? 'sealed' : 'up to date')));
  });

  if (!reqs.length) {
    return { seasons: seasonsStr_(cfg), calls: lp.calls, rows: 0, notes: `nothing to fetch | ${plan.join(' · ')}` };
  }

  const b = fetchMatchups_(reqs, 'UPDATE');
  const tested = lp.metas.filter(m => reqs.some(r => r.m === m));
  const tests = runTests_('UPDATE', tested, b.weeks, b.rws, b.rows, cfg);
  writeTests_('UPDATE', tests);

  const verdict = tests.filter(t => t[2] === 'G9-ACCEPT');
  const vtxt = verdict.map(t => `${t[1]}:${t[6]}`).join(' ');
  if (!verdict.length || verdict.some(t => t[6] !== 'PASS')) {
    return { seasons: seasonsStr_(cfg), calls: lp.calls + reqs.length, rows: 0, status: 'TESTS FAILED',
             notes: `G9-ACCEPT ${vtxt || 'none'} | RAW_G9 NOT changed | ${plan.join(' · ')}` };
  }
  const how = writeG9Increment_(sh, sw, replace, b.rows);
  return { seasons: seasonsStr_(cfg), calls: lp.calls + reqs.length, rows: b.rows.length, status: 'OK',
           notes: `G9-ACCEPT ${vtxt} | ${plan.join(' · ')} | ${how} | next: G5 → roster snapshot → Step 0 → Step 2` };
}

// Append in place when the new weeks sit at the end of the sorted table; else a sorted full rewrite.
function writeG9Increment_(sh, sw, replace, newRows) {
  if (!newRows.length) return 'no rows returned (nothing written)';
  const n = G9_HEADERS.length, key = x => x.s * 100 + x.w;
  let first = sw.length;
  for (let i = 0; i < sw.length; i++) if (replace.has(`${sw[i].s}|${sw[i].w}`)) { first = i; break; }
  let sorted = true;
  for (let i = 1; i < sw.length; i++) if (key(sw[i]) < key(sw[i - 1])) { sorted = false; break; }
  const tailIsReplaced = sw.slice(first).every(x => replace.has(`${x.s}|${x.w}`));
  const lastKept = first > 0 ? key(sw[first - 1]) : -Infinity;
  const minNew = Math.min.apply(null, newRows.map(r => Number(r[0]) * 100 + Number(r[2])));

  if (sorted && tailIsReplaced && minNew > lastKept) {
    if (first < sw.length) sh.getRange(2 + first, 1, sw.length - first, n).clearContent();
    const start = 2 + first;
    [2, 8].forEach(c => sh.getRange(start, c, newRows.length, 1).setNumberFormat('@')); // IDs stay text
    for (let i = 0; i < newRows.length; i += 2000) {
      const part = newRows.slice(i, i + 2000);
      sh.getRange(start + i, 1, part.length, n).setValues(part);
    }
    return `appended ${newRows.length} rows` + (first < sw.length ? ` (replaced ${sw.length - first} refreshed-week rows)` : '');
  }

  const all = sh.getRange(2, 1, sw.length, n).getValues()
    .filter(r => !replace.has(`${Number(r[0])}|${Number(r[2])}`)).concat(newRows);
  all.sort((a, b) => (Number(a[0]) - Number(b[0])) || (Number(a[2]) - Number(b[2]))); // stable: keeps in-week order
  writeTableChunked_(T.G9, G9_HEADERS, all, [2, 8]);
  return `FULL REWRITE (new weeks not at table end or table unsorted) — ${all.length} rows`;
}

/* ============================== acceptance tests (VII.G9) ============================== */
function badIdWeeks_(W, totalRosters) {
  let bad = 0;
  W.forEach(w => {
    const ids = w.entries.map(e => e.roster_id);
    if (new Set(ids).size !== ids.length || ids.some(id => id == null || id < 1 || id > totalRosters)) bad++;
  });
  return bad;
}

function crossRosterDupes_(R) {
  let dupes = 0;
  const seen = {};
  R.forEach(r => {
    seen[r.week] = seen[r.week] || {};
    r.players.forEach(p => {
      const prior = seen[r.week][p];
      if (prior !== undefined && prior !== r.rosterId) dupes++; else seen[r.week][p] = r.rosterId;
    });
  });
  return dupes;
}

function runTests_(mode, metas, weeks, rws, rows, cfg) {
  const out = [], at = new Date();
  const add = (season, id, q, cond, actual, status) => out.push([mode, season, id, q, cond, actual, status, at]);
  const statusOf = (season, id) => { const r = out.find(x => x[1] === season && x[2] === id); return r ? r[6] : 'MISSING'; };

  metas.forEach((m, idx) => {
    const S = m.cfgSeason;
    add(S, 'CFG-1', 'API season matches CONFIG season', 'API_Season == Config_Season',
      `${m.apiSeason} vs ${S}`, m.apiSeason === S ? 'PASS' : 'FAIL');
    if (idx > 0) {
      const prev = metas[idx - 1];
      add(S, 'CFG-2', 'League chain links to prior configured season', 'previous_league_id == prior CONFIG League_ID',
        `${m.prevId || '(none)'} vs ${prev.leagueId}`, m.prevId === prev.leagueId ? 'PASS' : 'WARN');
    }

    const W = weeks.filter(w => w.season === S), Wr = W.filter(w => w.isReg);
    const R = rws.filter(r => r.season === S), Rr = R.filter(r => r.isReg);
    if (!Wr.length) {
      if (mode === 'UPDATE' && W.length) {
        // Playoff-only update: same scope FULL applies to playoff weeks (Q1b, Q3d, Q4p).
        const bi = badIdWeeks_(W, m.totalRosters), du = crossRosterDupes_(R);
        add(S, 'Q1b', 'Q1 — Roster IDs present, unique, in range within each week', '0 bad weeks',
          `${bi} bad weeks of ${W.length}`, bi ? 'FAIL' : 'PASS');
        add(S, 'Q3d', 'Q3 — No player on two rosters in the same week', '0 duplicates', `${du} duplicates`, du ? 'FAIL' : 'PASS');
        add(S, 'Q4p', 'Playoff weeks returned (info)', 'informational',
          `${W.filter(w => w.n > 0).length}/${W.length} non-empty (weeks ${W.map(w => w.week).join(',')})`, 'INFO');
        const ok = !bi && !du;
        add(S, 'G9-ACCEPT', 'G9 ingestion acceptable (playoff weeks only — FULL playoff scope)', 'Q1b and Q3d PASS',
          ok ? 'playoff-scope checks pass' : 'failing: ' + [bi ? 'Q1b' : '', du ? 'Q3d' : ''].filter(Boolean).join(', '),
          ok ? 'PASS' : 'FAIL');
        return;
      }
      add(S, 'Q4', 'All requested regular-season weeks returned data', '≥1 regular-season week requested',
        '0 regular-season weeks requested (season not started?)', 'FAIL');
      add(S, 'G9-ACCEPT', 'G9 ingestion acceptable (Q1–Q4)', 'All Q1–Q4 PASS (Q2b may WARN)', 'no data', 'FAIL');
      return;
    }

    // ---- Q1: each response identifies roster (week is endpoint-scoped) ----
    const incomplete = Wr.filter(w => w.n !== w.expected);
    add(S, 'Q1a', 'Q1 — Full roster set returned per regular-season week (week = endpoint)',
      `every reg week returns ${m.totalRosters} roster entries`,
      `${Wr.length - incomplete.length}/${Wr.length} complete` +
        (incomplete.length ? ` (bad: ${incomplete.map(w => 'wk' + w.week + '=' + w.n).join(', ')})` : ''),
      incomplete.length ? 'FAIL' : 'PASS');
    const badIdWeeks = badIdWeeks_(W, m.totalRosters);
    add(S, 'Q1b', 'Q1 — Roster IDs present, unique, in range within each week', '0 bad weeks',
      `${badIdWeeks} bad weeks of ${W.length}`, badIdWeeks ? 'FAIL' : 'PASS');
    const noMid = Rr.filter(r => r.matchupId == null).length;
    add(S, 'Q1c', 'Q1 — Matchup ID present (regular season)', '0 roster-weeks missing matchup_id',
      `${noMid} missing of ${Rr.length}`, noMid ? 'FAIL' : 'PASS');

    // ---- Q2: bench included in players_points ----
    let missingKey = 0, totalPlayers = 0, extraKeys = 0;
    const benchSizes = [];
    Rr.forEach(r => {
      const st = new Set(r.starters.filter(p => p !== '0'));
      r.players.forEach(p => { totalPlayers++; if (!(p in r.pp)) missingKey++; });
      Object.keys(r.pp).forEach(k => { if (r.players.indexOf(k) < 0) extraKeys++; });
      benchSizes.push(r.players.filter(p => !st.has(p)).length);
    });
    add(S, 'Q2a', 'Q2 — Every rostered player (bench incl.) has a players_points entry', '0 missing',
      `${missingKey} missing of ${totalPlayers} rostered player-weeks`, (missingKey || !totalPlayers) ? 'FAIL' : 'PASS');
    const zeroBench = benchSizes.filter(b => b === 0).length;
    add(S, 'Q2b', 'Q2 — Bench players actually present', 'bench > 0 in every reg roster-week (WARN if some 0)',
      `${benchSizes.length - zeroBench}/${benchSizes.length} roster-weeks have bench; bench size ${stats_(benchSizes)}`,
      !benchSizes.length || zeroBench === benchSizes.length ? 'FAIL' : (zeroBench ? 'WARN' : 'PASS'));
    add(S, 'Q2c', 'Q2 — players_points keys not in players list (info)', 'informational',
      `${extraKeys} extra keys`, 'INFO');

    // ---- Q3: scores reliably tied to the correct roster ----
    let notSubset = 0, spMismatch = 0, totMismatch = 0, customUsed = 0, emptySlots = 0;
    Rr.forEach(r => {
      r.starters.forEach((p, i) => {
        if (p === '0') { emptySlots++; return; }
        if (r.players.indexOf(p) < 0) notSubset++;
        if (Math.abs((Number(r.sp[i]) || 0) - (Number(r.pp[p]) || 0)) > TOL) spMismatch++;
      });
      const sum = r.sp.reduce((a, b) => a + (Number(b) || 0), 0);
      if (r.customPoints != null) customUsed++;
      else if (Math.abs(sum - (Number(r.teamPoints) || 0)) > TOL * 5) totMismatch++;
    });
    add(S, 'Q3a', 'Q3 — Starters are a subset of the roster', '0 violations',
      `${notSubset} violations; ${emptySlots} empty starter slots`, notSubset ? 'FAIL' : 'PASS');
    add(S, 'Q3b', 'Q3 — starters_points agrees with players_points', '0 mismatches (±0.01)',
      `${spMismatch} mismatches`, spMismatch ? 'FAIL' : 'PASS');
    add(S, 'Q3c', 'Q3 — Sum of starters_points = team points', '0 mismatches (±0.05); custom_points excluded',
      `${totMismatch} mismatches; ${customUsed} roster-weeks with custom_points`, totMismatch ? 'FAIL' : 'PASS');
    const dupes = crossRosterDupes_(R);
    add(S, 'Q3d', 'Q3 — No player on two rosters in the same week', '0 duplicates', `${dupes} duplicates`, dupes ? 'FAIL' : 'PASS');

    // ---- Q4: coverage across requested weeks ----
    const empty = Wr.filter(w => w.n === 0).length;
    add(S, 'Q4', 'Q4 — All requested regular-season weeks returned data', 'all non-empty',
      `${Wr.length - empty}/${Wr.length} non-empty (weeks ${Wr.map(w => w.week).join(',')})`, empty ? 'FAIL' : 'PASS');
    const Wp = W.filter(w => !w.isReg);
    if (Wp.length) add(S, 'Q4p', 'Playoff weeks returned (info)', 'informational',
      `${Wp.filter(w => w.n > 0).length}/${Wp.length} non-empty`, 'INFO');

    // ---- Q5: historical ownership reconstructable ----
    const sizes = Rr.map(r => r.players.length);
    const nStart = m.rosterPositions.filter(p => p !== 'BN').length;
    const cap = m.rosterPositions.length + m.reserveSlots + m.taxiSlots;
    const oor = sizes.filter(n => n < nStart || n > cap).length;
    add(S, 'Q5a', 'Q5 — Roster sizes plausible vs league limits', `${nStart} ≤ size ≤ ${cap} (starters..BN+IR+taxi)`,
      `size ${stats_(sizes)}; ${oor} out of range`, oor ? 'WARN' : 'PASS');
    const byR = {};
    Rr.forEach(r => { (byR[r.rosterId] = byR[r.rosterId] || {})[r.week] = new Set(r.players); });
    let pairs = 0, retained = 0;
    Object.values(byR).forEach(wk => Object.keys(wk).map(Number).forEach(w => {
      if (wk[w + 1]) {
        pairs++;
        let k = 0;
        wk[w].forEach(p => { if (wk[w + 1].has(p)) k++; });
        retained += wk[w].size ? k / wk[w].size : 0;
      }
    }));
    add(S, 'Q5b', 'Q5 — Week-over-week roster continuity (info)', 'informational; expect high, not 100%',
      pairs ? `${(100 * retained / pairs).toFixed(1)}% avg retained over ${pairs} consecutive roster-week pairs`
            : 'n/a (no consecutive weeks in this mode)', 'INFO');
    const structural = ['Q1a', 'Q1b', 'Q2a', 'Q3a', 'Q3d'].every(id => statusOf(S, id) === 'PASS');
    add(S, 'Q5', 'Q5 — Historical ownership (player→roster→week) reconstructable',
      'Q1a,Q1b,Q2a,Q3a,Q3d all PASS. Reconstructable ≠ comparable to A1 (that is G10)',
      structural ? 'structural checks pass' : 'one or more structural checks failed', structural ? 'PASS' : 'FAIL');

    const gateIds = ['Q1a', 'Q1b', 'Q1c', 'Q2a', 'Q3a', 'Q3b', 'Q3c', 'Q3d', 'Q4'];
    const failed = gateIds.filter(id => statusOf(S, id) !== 'PASS');
    const q2b = statusOf(S, 'Q2b');
    const ok = !failed.length && q2b !== 'FAIL';
    add(S, 'G9-ACCEPT', 'G9 ingestion acceptable (Q1–Q4)', 'All Q1–Q4 PASS (Q2b may WARN)',
      ok ? `all gate tests pass${q2b === 'WARN' ? ' (Q2b WARN — review)' : ''}`
         : `failing: ${failed.concat(q2b === 'FAIL' ? ['Q2b'] : []).join(', ')}`, ok ? 'PASS' : 'FAIL');
  });

  // ---- Reconciliation vs existing Raw_Starters (FULL only) ----
  if (mode === 'FULL' && cfg.refSeasons.length) {
    const st = rows.filter(r => r[11] === true && r[3] === true && cfg.refSeasons.indexOf(r[0]) >= 0);
    const empties = rws.filter(r => r.isReg && cfg.refSeasons.indexOf(r.season) >= 0)
      .reduce((a, r) => a + r.starters.filter(p => p === '0').length, 0);
    if (cfg.expectedRef) add('ALL', 'REC-1', `Reg-season starter SLOTS (${cfg.refSeasons.join(',')}) match Raw_Starters count`,
      `filled + empty slots = ${cfg.expectedRef}`, `${st.length} filled + ${empties} empty = ${st.length + empties}`,
      st.length + empties === cfg.expectedRef ? 'PASS' : 'FAIL');
    const ref = SpreadsheetApp.getActive().getSheetByName(T.REF);
    if (!ref) {
      add('ALL', 'REC-2', 'Row-level match vs REF_RAW_STARTERS', 'tab present',
        'REF_RAW_STARTERS not found; row-level reconciliation skipped', 'INFO');
    } else {
      const v = ref.getDataRange().getValues();
      const h = v[0].map(x => String(x).trim());
      const c = ['Season', 'Week', 'Roster_ID', 'Player_ID', 'Points'].map(n => h.indexOf(n));
      if (c.some(i => i < 0)) {
        add('ALL', 'REC-2', 'Row-level match vs REF_RAW_STARTERS', 'needs Season, Week, Roster_ID, Player_ID, Points',
          `missing columns in header: ${h.join(', ')}`, 'FAIL');
      } else {
        const g = new Map();
        st.forEach(r => g.set(`${r[0]}|${r[2]}|${r[4]}|${r[7]}`, Number(r[8])));
        let matched = 0, missing = 0, ptsDiff = 0;
        const hit = new Set();
        for (let i = 1; i < v.length; i++) {
          if (v[i][c[0]] === '') continue;
          const k = `${Number(v[i][c[0]])}|${Number(v[i][c[1]])}|${Number(v[i][c[2]])}|${String(v[i][c[3]]).trim()}`;
          if (!g.has(k)) { missing++; continue; }
          hit.add(k);
          if (Math.abs(g.get(k) - Number(v[i][c[4]])) > TOL) ptsDiff++; else matched++;
        }
        const extra = [...g.keys()].filter(k => !hit.has(k)).length;
        add('ALL', 'REC-2', 'Row-level match vs REF_RAW_STARTERS', '0 missing, 0 extra, 0 points mismatches',
          `${matched} matched, ${missing} missing, ${extra} extra, ${ptsDiff} points mismatches`,
          (missing || extra || ptsDiff) ? 'FAIL' : 'PASS');
      }
    }
  }
  return out;
}

function assertSampleGate_(cfg) {
  const sh = sheet_(T.TESTS);
  if (sh.getLastRow() < 2) throw new Error('Gate closed: no SAMPLE tests found. Run step 3 first.');
  const v = sh.getRange(2, 1, sh.getLastRow() - 1, TEST_HEADERS.length).getValues().filter(r => r[0] === 'SAMPLE');
  const bad = cfg.leagues.filter(l => {
    const r = v.find(x => Number(x[1]) === l.season && x[2] === 'G9-ACCEPT');
    return !r || r[6] !== 'PASS';
  });
  if (bad.length) throw new Error(`Gate closed: SAMPLE G9-ACCEPT is not PASS for ${bad.map(b => b.season).join(', ')}. Review G9_TESTS.`);
}

/* ============================== IO helpers ============================== */
function fetchAllJson_(urls) {
  const out = [], CHUNK = 40;
  for (let i = 0; i < urls.length; i += CHUNK) {
    const chunk = urls.slice(i, i + CHUNK);
    const resps = UrlFetchApp.fetchAll(chunk.map(u => ({ url: u, muteHttpExceptions: true })));
    resps.forEach((r, j) => {
      const code = r.getResponseCode();
      if (code !== 200) throw new Error(`HTTP ${code} for ${chunk[j]}`);
      const txt = r.getContentText();
      out.push(txt ? JSON.parse(txt) : null);
    });
  }
  return out;
}

function sheet_(name, create) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    if (!create) throw new Error(`Missing tab ${name}. Run "1. Set up tabs".`);
    sh = ss.insertSheet(name);
  }
  return sh;
}

function header_(name, headers) {
  const sh = sheet_(name, true);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
}

function writeTable_(name, headers, rows, textCols) {
  const sh = sheet_(name, true);
  sh.clearContents();
  const n = Math.max(rows.length + 1, 2);
  (textCols || []).forEach(c => sh.getRange(1, c, n, 1).setNumberFormat('@')); // IDs stay text
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  if (rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sh.setFrozenRows(1);
}

function writeTests_(mode, tests) {
  const sh = sheet_(T.TESTS, true);
  const keep = sh.getLastRow() > 1
    ? sh.getRange(2, 1, sh.getLastRow() - 1, TEST_HEADERS.length).getValues().filter(r => r[0] !== mode) : [];
  const all = keep.concat(tests);
  sh.clearContents();
  sh.getRange(1, 1, 1, TEST_HEADERS.length).setValues([TEST_HEADERS]).setFontWeight('bold');
  if (all.length) sh.getRange(2, 1, all.length, TEST_HEADERS.length).setValues(all);
  sh.setFrozenRows(1);
}

function log_(fn, mode, seasons, calls, rowsN, status, notes) {
  header_(T.LOG, LOG_HEADERS);
  const sh = sheet_(T.LOG);
  const r = sh.getLastRow() + 1;
  sh.getRange(r, 4).setNumberFormat('@'); // appendRow ignores text format; write via setValues instead
  sh.getRange(r, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sh.getRange(r, 1, 1, 8).setValues([[new Date(), fn, mode, seasons, calls, rowsN, status, String(notes || '').slice(0, 45000)]]);
}

function run_(fn, mode, body) {
  const t0 = Date.now();
  try {
    const r = body() || {};
    const secs = ((Date.now() - t0) / 1000).toFixed(1);
    log_(fn, mode, r.seasons || '', r.calls || 0, r.rows || 0, r.status || 'OK', `${r.notes || ''} | ${secs}s`);
    SpreadsheetApp.getActive().toast(`${r.status || 'OK'}: ${r.notes || ''}`, `Dynasty FO — ${fn}`, 12);
  } catch (e) {
    log_(fn, mode, '', 0, 0, 'ERROR', e.message);
    SpreadsheetApp.getActive().toast(e.message, `Dynasty FO — ${fn} ERROR`, 15);
    throw e;
  }
}

function seasonsStr_(cfg) { return cfg.leagues.map(l => l.season).join(','); }

function stats_(a) {
  if (!a.length) return 'n/a';
  const s = a.slice().sort((x, y) => x - y);
  return `min ${s[0]} / median ${s[Math.floor(s.length / 2)]} / max ${s[s.length - 1]}`;
}

/* ============================== README ============================== */
function writeReadme_() {
  const lines = [
    ['DYNASTY FRONT OFFICE — G9 INGESTION LAYER'],
    [''],
    ['PURPOSE'],
    ['Ingest unconditional weekly production (bench included) from Sleeper matchups players_points, and verify via the'],
    ['VII.G9 acceptance test whether historical roster ownership is reconstructable (decides the G7/G10 branch).'],
    [''],
    ['WORKFLOW'],
    ['1. Dynasty FO → Set up tabs.  2. Enter League IDs in CONFIG (as text).'],
    ['3. Run G9 SAMPLE: tests week 1 + last regular-season week per season (Q1–Q5). Tests only; never writes RAW_G9.'],
    ['4. Review G9_TESTS. 5. Run G9 FULL once: refuses to run unless SAMPLE G9-ACCEPT = PASS for every season.'],
    ['WEEKLY: 3b G9 UPDATE → 5 G5 → 5b roster snapshot → 5c universe → 5d market values → 6 Step 0 → 7 Step 2. Each is its own run (6-min limit).'],
    ['UPDATE fetches only new completed weeks + the last Refresh_Weeks weeks of the current season (stat corrections).'],
    ['Completed seasons are sealed. Fetched weeks are tested BEFORE writing; on any G9-ACCEPT failure RAW_G9 is unchanged.'],
    ['Optional: paste the existing Raw_Starters as tab REF_RAW_STARTERS for row-level reconciliation (REC-2, FULL only).'],
    [''],
    ['TABS'],
    ['CONFIG: inputs you set only (league IDs, week scope, Refresh_Weeks — default 2). Never scoring values.'],
    ['RAW_LEAGUE: league settings verbatim from Sleeper per season (scoring, roster_positions, playoff start).'],
    ['RAW_ROSTERS: roster_id → owner/team per season. Reflects the league’s final/current owner mapping, not week-level history.'],
    ['RAW_G9: one row per season × week × roster × player. Source fields only; nothing derived.'],
    ['RAW_ROSTER_SNAPSHOT: current roster membership + IR/taxi status (see RS_TESTS).'],
    ['RAW_BRACKETS / RAW_DRAFTS / RAW_TRADED_PICKS / PICK_OWNERSHIP / RAW_FREE_AGENTS: see UV_TESTS.'],
    ['RAW_MARKET_VALUES: FantasyCalc, append-only daily history. DECISION LAYER ONLY — never an input to C1/C2/A1/A2 (MV_TESTS).'],
    ['G9_TESTS: single source of truth for test status (SAMPLE / FULL / UPDATE rows). RUN_LOG: every run, counts, errors.'],
    [''],
    ['RULES'],
    ['Current season: only completed weeks are ingested (state.week − 1).'],
    ['Q5 PASS means ownership is reconstructable, NOT that historical capacity is comparable to A1 (that is G10).'],
    ['Test status lives in G9_TESTS only.']
  ];
  const sh = sheet_(T.README, true);
  sh.clearContents();
  sh.getRange(1, 1, lines.length, 1).setValues(lines);
  [1, 3, 7, 16, 26].forEach(r => sh.getRange(r, 1).setFontWeight('bold'));
  sh.setColumnWidth(1, 900);
}
