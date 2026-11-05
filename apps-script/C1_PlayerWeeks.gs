/***** Step 2 — Canonical player-week dataset (C1_PLAYER_WEEKS) + C1D acceptance tests ********************
 * Add as a FOURTH file in the same Apps Script project (reuses helpers from the other three files).
 * Needs: RAW_G9, S0_BYE_CLASS + S0_TESTS (from Step 0), RAW_STATS, RAW_PLAYER_META.
 * Writes: C1_PLAYER_WEEKS and C1D_TESTS.
 *
 * What it is: a JOIN, not a re-derivation. One row per RAW_G9 player-week, carrying:
 *   - facts from RAW_G9 (points verbatim, starter flag, roster, regular-season flag)
 *   - Bye_Status / NFL_Team / Team_Source exactly as Step 0 classified and tested them (S0_BYE_CLASS)
 *   - Team_Change_Since_Prior_Obs: TRUE/FALSE only when the current week's team AND the player's most recent
 *     earlier own-stats-row team (any earlier week, any season) both come from own stats rows; else UNKNOWN.
 *     A first-ever observation is UNKNOWN, never FALSE.
 *   - Position and Eligible_Positions (fantasy_positions ∩ QB/RB/WR/TE) from RAW_PLAYER_META.
 * Not stored: averages, ages, projections, evidence-strength counts (cutoff-dependent → derived in Step 3).
 * Role change is deliberately absent (not recoverable from this data — Foundation Spec §II.6).
 *
 * Refuses to build if S0-GATE = FAIL, or if S0_BYE_CLASS does not cover exactly the RAW_G9 player-weeks
 * (means Step 0 is stale relative to the last G9 ingest → rerun Step 0). On refusal C1_PLAYER_WEEKS is cleared,
 * so a stale table can never be consumed.
 * Gate: C1D-GATE = FAIL if any FAIL · PASS_RESTRICTED if S0-GATE was PASS_RESTRICTED · else PASS.
 ***************************************************************************************************/
const C1 = { WEEKS: 'C1_PLAYER_WEEKS', TESTS: 'C1D_TESTS' };
const C1_HEADERS = ['Season', 'Week', 'Is_Regular_Season', 'Player_ID', 'Player_Name', 'Roster_ID', 'Points',
  'Is_Starter', 'Bye_Status', 'NFL_Team', 'Team_Source', 'Team_Change_Since_Prior_Obs', 'Position',
  'Eligible_Positions', 'Built_At'];

function runStep2() { run_('runStep2', 'C1D', () => step2_()); }

function step2_() {
  const cfg = readConfig_();
  const seasons = cfg.leagues.map(l => l.season);
  const truthy = v => v === true || String(v).toUpperCase() === 'TRUE';
  const tab = name => {
    const v = sheet_(name).getDataRange().getValues(), h = v[0].map(String);
    return { v, c: n => { const i = h.indexOf(n); if (i < 0) throw new Error(`${name} is missing column ${n}. Rerun the step that writes it.`); return i; } };
  };

  // ---- inherit the Step 0 gate ----
  const s0 = sheet_(S0.TESTS).getDataRange().getValues().find(r => r[2] === 'S0-GATE');
  if (!s0) throw new Error('No S0-GATE in S0_TESTS. Run step 6 first.');
  if (s0[6] === 'FAIL') throw new Error(`S0-GATE is FAIL (${s0[5]}). Step 2 refuses to build.`);
  const s0Gate = String(s0[6]), s0Detail = String(s0[5]);

  // ---- load sources ----
  const G = tab(T.G9), B = tab(S0.CLASS), RS = tab(S0.STATS), PM = tab(G5.META);
  const gS = G.c('Season'), gW = G.c('Week'), gReg = G.c('Is_Regular_Season'), gR = G.c('Roster_ID'),
        gP = G.c('Player_ID'), gPts = G.c('Points'), gSt = G.c('Is_Starter');

  const bye = {};
  let byeDupes = 0;
  const bS = B.c('Season'), bW = B.c('Week'), bP = B.c('Player_ID'), bSt = B.c('Bye_Status'), bT = B.c('NFL_Team'), bSrc = B.c('Team_Source');
  for (let r = 1; r < B.v.length; r++) {
    const k = `${Number(B.v[r][bS])}|${Number(B.v[r][bW])}|${String(B.v[r][bP]).trim()}`;
    if (bye[k]) byeDupes++;
    bye[k] = { status: String(B.v[r][bSt]), team: String(B.v[r][bT] || ''), src: String(B.v[r][bSrc]) };
  }

  // own-stats-row team observations per player, in time order
  const teamObs = {};
  const rS = RS.c('Season'), rW = RS.c('Week'), rP = RS.c('Player_ID'), rF = RS.c('Stats_Row_Found'), rT = RS.c('NFL_Team');
  for (let r = 1; r < RS.v.length; r++) {
    const team = String(RS.v[r][rT] || '');
    if (!truthy(RS.v[r][rF]) || !team) continue;
    const pid = String(RS.v[r][rP]).trim();
    (teamObs[pid] = teamObs[pid] || []).push({ s: Number(RS.v[r][rS]), w: Number(RS.v[r][rW]), team });
  }
  Object.keys(teamObs).forEach(p => teamObs[p].sort((a, b) => a.s - b.s || a.w - b.w));

  const meta = {};
  const mP = PM.c('Player_ID'), mN = PM.c('Full_Name'), mPos = PM.c('Position'), mFp = PM.c('Fantasy_Positions_JSON');
  for (let r = 1; r < PM.v.length; r++) {
    let elig = [];
    try { elig = JSON.parse(PM.v[r][mFp] || '[]').filter(p => LINEUP_POS.indexOf(p) >= 0); } catch (e) { elig = null; }
    meta[String(PM.v[r][mP]).trim()] = { name: String(PM.v[r][mN] || ''), pos: String(PM.v[r][mPos] || ''), elig };
  }

  // ---- build ----
  const g9 = [];
  for (let r = 1; r < G.v.length; r++) {
    const pid = String(G.v[r][gP]).trim();
    if (!pid) continue;
    g9.push({ s: Number(G.v[r][gS]), w: Number(G.v[r][gW]), reg: truthy(G.v[r][gReg]), roster: G.v[r][gR],
              pid, pts: G.v[r][gPts], starter: truthy(G.v[r][gSt]) });
  }
  g9.sort((a, b) => a.s - b.s || a.w - b.w || (a.pid < b.pid ? -1 : a.pid > b.pid ? 1 : 0));
  const g9Keys = new Set(g9.map(x => `${x.s}|${x.w}|${x.pid}`));
  const missingClass = g9.filter(x => !bye[`${x.s}|${x.w}|${x.pid}`]);
  const extraClass = Object.keys(bye).filter(k => !g9Keys.has(k));

  const now = new Date();
  const out = g9.map(x => {
    const b = bye[`${x.s}|${x.w}|${x.pid}`] || { status: '', team: '', src: '' };
    let change = 'UNKNOWN';
    if (b.src === 'STATS' && b.team) {
      const obs = teamObs[x.pid] || [];
      let prior = null;
      for (let i = obs.length - 1; i >= 0; i--) {
        if (obs[i].s < x.s || (obs[i].s === x.s && obs[i].w < x.w)) { prior = obs[i]; break; }
      }
      if (prior) change = prior.team === b.team ? 'FALSE' : 'TRUE';
    }
    const m = meta[x.pid] || { name: '', pos: '', elig: [] };
    return [x.s, x.w, x.reg, x.pid, m.name, x.roster, x.pts, x.starter, b.status, b.team, b.src, change, m.pos,
      m.elig ? m.elig.join('|') : 'UNPARSEABLE', now];
  });

  // ---- tests ----
  const tests = [], at = new Date();
  const add = (season, id, q, cond, actual, status) => tests.push(['C1D', season, id, q, cond, actual, status, at]);
  const pct = (n, d) => d ? `${n}/${d} (${(100 * n / d).toFixed(2)}%)` : 'n/a (0)';

  add('ALL', 'C1D-0', 'Inherited Step 0 gate', 'not FAIL (checked before build)', `${s0Gate}: ${s0Detail}`,
    s0Gate === 'PASS' ? 'PASS' : 'WARN');
  add('ALL', 'C1D-1', 'Exactly one row per RAW_G9 player-week', `${g9.length} rows`, `${out.length} rows`,
    out.length === g9.length ? 'PASS' : 'FAIL');
  const keys = out.map(r => `${r[0]}|${r[1]}|${r[3]}`);
  const dupes = keys.length - new Set(keys).size;
  add('ALL', 'C1D-2', 'No duplicate player-weeks (season, week, player)', '0', `${dupes}`, dupes ? 'FAIL' : 'PASS');
  const badPts = out.filter(r => r[6] !== '' && typeof r[6] !== 'number' && isNaN(Number(r[6])));
  add('ALL', 'C1D-3', 'Points carried verbatim from RAW_G9 players_points', '0 non-numeric values; blanks kept as blank',
    `${badPts.length} non-numeric · ${out.filter(r => r[6] === '').length} blank`, badPts.length ? 'FAIL' : 'PASS');
  add('ALL', 'C1D-4', 'Step 0 classification covers exactly the RAW_G9 player-weeks (not stale)',
    '0 missing, 0 extra, 0 duplicate classification rows (else rerun Step 0)',
    `missing ${missingClass.length} · extra ${extraClass.length} · duplicate ${byeDupes}` +
      (missingClass.length ? ` · e.g. ${missingClass.slice(0, 5).map(x => `${x.s} wk${x.w} ${x.pid}`).join(', ')}` : ''),
    missingClass.length || extraClass.length || byeDupes ? 'FAIL' : 'PASS');
  const badStatus = out.filter(r => ['BYE', 'NOT_BYE', 'UNKNOWN'].indexOf(r[8]) < 0).length;
  add('ALL', 'C1D-5', 'Every row has a valid Bye_Status', '0 rows outside BYE/NOT_BYE/UNKNOWN', `${badStatus}`, badStatus ? 'FAIL' : 'PASS');
  const byePts = out.filter(r => r[8] === 'BYE' && Number(r[6]));
  add('ALL', 'C1D-6', 'No BYE row with non-zero points', '0',
    `${byePts.length}` + (byePts.length ? `: ${byePts.slice(0, 8).map(r => `${r[0]} wk${r[1]} ${r[4] || r[3]} ${r[6]}`).join(' | ')}` : ''),
    byePts.length ? 'FAIL' : 'PASS');
  add('ALL', 'C1D-7', 'Role-change field absent (not recoverable — §II.6)', 'no role field in schema',
    C1_HEADERS.some(h => /role/i.test(h)) ? 'role field present' : 'absent', C1_HEADERS.some(h => /role/i.test(h)) ? 'FAIL' : 'PASS');

  seasons.forEach(S => {
    const R = out.filter(r => r[0] === S);
    const nb = R.filter(r => r[8] === 'NOT_BYE'), nbReg = nb.filter(r => r[2] === true);
    const zero = nbReg.filter(r => Number(r[6]) === 0).length;
    add(S, 'C1D-8', 'Modelling population (NOT_BYE) and exclusions', 'informational',
      `rows ${R.length} · NOT_BYE ${nb.length} (regular season ${nbReg.length}, of which zero-point ${pct(zero, nbReg.length)}) · ` +
        `BYE excluded ${R.filter(r => r[8] === 'BYE').length} · UNKNOWN excluded ${R.filter(r => r[8] === 'UNKNOWN').length} · ` +
        `distinct players with ≥1 NOT_BYE ${new Set(nb.map(r => r[3])).size}`, 'INFO');
    const tc = { TRUE: 0, FALSE: 0, UNKNOWN: 0 };
    R.forEach(r => { tc[r[11]]++; });
    add(S, 'C1D-9', 'Team_Change_Since_Prior_Obs distribution', 'informational (UNKNOWN never read as "no change")',
      `TRUE ${tc.TRUE} · FALSE ${tc.FALSE} · UNKNOWN ${tc.UNKNOWN}`, 'INFO');
  });
  const changed = out.filter(r => r[11] === 'TRUE');
  add('ALL', 'C1D-10', 'Inspection sample — detected team changes', 'manual check (most recent first)',
    changed.slice(-10).reverse().map(r => `${r[0]} wk${r[1]} ${r[4] || r[3]} → ${r[9]}`).join(' | ') || 'none', 'INFO');
  const noElig = [...new Set(out.filter(r => !r[13] || r[13] === 'UNPARSEABLE').map(r => `${r[4] || r[3]} [${r[12]}]`))];
  add('ALL', 'C1D-11', 'Players with no lineup-eligible position (can never start)', 'informational',
    noElig.length ? noElig.join(', ') : 'none', 'INFO');

  const fails = tests.filter(t => t[6] === 'FAIL').map(t => t[2]);
  const gate = fails.length ? 'FAIL' : s0Gate === 'PASS_RESTRICTED' ? 'PASS_RESTRICTED' : 'PASS';
  add('ALL', 'C1D-GATE', 'C1_PLAYER_WEEKS acceptable → Step 3 may consume it',
    'FAIL if any FAIL (table cleared) · PASS_RESTRICTED if Step 0 was restricted (UNKNOWN excluded + reported) · else PASS',
    fails.length ? `blocked by ${fails.join(', ')}` : gate === 'PASS_RESTRICTED' ? `restricted (inherited): ${s0Detail}` : 'all conditions met', gate);

  // ---- write ----
  if (gate === 'FAIL') {
    const sh = sheet_(C1.WEEKS, true);
    sh.clearContents();
    sh.getRange(1, 1).setValue('Cleared: C1D-GATE FAIL — see C1D_TESTS. Do not consume.');
  } else {
    writeTableChunked_(C1.WEEKS, C1_HEADERS, out, [4, 10, 14]);
  }
  const th = sheet_(C1.TESTS, true);
  th.clearContents();
  th.getRange(1, 1, 1, 8).setValues([['Run', 'Season', 'Test_ID', 'Question', 'Pass_Condition', 'Actual', 'Status', 'Run_At']]).setFontWeight('bold');
  th.getRange(2, 1, tests.length, 8).setValues(tests);
  th.setFrozenRows(1);

  return { seasons: seasonsStr_(cfg), calls: 0, rows: gate === 'FAIL' ? 0 : out.length,
           status: gate === 'FAIL' ? 'TESTS FAILED' : 'OK', notes: `C1D-GATE ${gate}` };
}
