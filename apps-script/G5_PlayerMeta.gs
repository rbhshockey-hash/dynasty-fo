/***** G5 — Player metadata (Sleeper) + typed acceptance test (VII.12b) ***************************
 * Add this as a SECOND file in the same Apps Script project (it reuses helpers from G9_Ingest.gs).
 * Population = every player in RAW_G9 (all seasons, bench incl.) ∪ the current league's live rosters.
 * Stores source facts only: name, position, team, status, birth_date (as text), years_exp. No age is stored;
 * age is derived later, as of an explicit date, so nothing goes stale and no timezone shift can touch DOB.
 * Writes: RAW_PLAYER_META (replaces any legacy Player_Map) and G5_TESTS.
 ***************************************************************************************************/
const G5 = { META: 'RAW_PLAYER_META', TESTS: 'G5_TESTS' };
const G5_TEST_HEADERS = ['Run', 'Type', 'Test_ID', 'Question', 'Pass_Condition', 'Actual', 'Status', 'Run_At'];
const META_HEADERS = ['Player_ID', 'Matched', 'Full_Name', 'Position', 'Fantasy_Positions_JSON', 'NFL_Team', 'Status',
  'Birth_Date', 'Years_Exp', 'On_Current_Roster', 'Current_Roster_ID', 'Rostered_Player_Weeks', 'Started_Weeks_All',
  'Started_Weeks_Current_Season', 'Seasons_Rostered', 'Fetched_At'];
const LINEUP_POS = ['QB', 'RB', 'WR', 'TE'];

function runG5() { run_('runG5', 'G5', () => g5_()); }

function g5_() {
  const cfg = readConfig_();
  const cur = cfg.leagues[cfg.leagues.length - 1];

  // ---- historical population from RAW_G9 ----
  const g9 = sheet_(T.G9).getDataRange().getValues();
  const h = g9[0].map(String);
  const cS = h.indexOf('Season'), cP = h.indexOf('Player_ID'), cSt = h.indexOf('Is_Starter');
  if ([cS, cP, cSt].some(c => c < 0) || g9.length < 2) throw new Error('RAW_G9 is empty or missing columns. Run G9 FULL first.');
  const hist = {};
  for (let r = 1; r < g9.length; r++) {
    const pid = String(g9[r][cP]).trim();
    if (!pid) continue;
    const s = Number(g9[r][cS]);
    const o = hist[pid] || (hist[pid] = { weeks: 0, starts: 0, startsCur: 0, seasons: {} });
    o.weeks++;
    o.seasons[s] = true;
    if (g9[r][cSt] === true || String(g9[r][cSt]).toUpperCase() === 'TRUE') {
      o.starts++;
      if (s === cur.season) o.startsCur++;
    }
  }

  // ---- live rosters + Sleeper player feed ----
  const res = fetchAllJson_([`${API}/league/${cur.leagueId}/rosters`, `${API}/players/nfl`]);
  const rosters = res[0] || [], players = res[1] || {};
  if (!Object.keys(players).length) throw new Error('Sleeper /players/nfl returned no data.');
  const curRoster = {};
  rosters.forEach(r => [].concat(r.players || [], r.reserve || [], r.taxi || [])
    .forEach(p => { curRoster[String(p)] = r.roster_id; }));

  const ids = Object.keys(Object.assign({}, hist, curRoster)).sort();
  const now = new Date();
  const rows = ids.map(pid => {
    const p = players[pid];
    const hh = hist[pid] || { weeks: 0, starts: 0, startsCur: 0, seasons: {} };
    const onCur = curRoster[pid] !== undefined;
    return [pid, !!p,
      p ? (p.full_name || [p.first_name, p.last_name].filter(Boolean).join(' ')) : '',
      p ? (p.position || '') : '', p ? JSON.stringify(p.fantasy_positions || []) : '',
      p ? (p.team || '') : '', p ? (p.status || '') : '',
      p && p.birth_date ? String(p.birth_date) : '', p && p.years_exp != null ? p.years_exp : '',
      onCur, onCur ? curRoster[pid] : '', hh.weeks, hh.starts, hh.startsCur,
      Object.keys(hh.seasons).sort().join('|'), now];
  });
  writeTable_(G5.META, META_HEADERS, rows, [1, 8, 15]); // IDs, DOB, seasons stay text

  // ---- typed acceptance test ----
  const out = [], at = new Date();
  const add = (type, id, q, cond, actual, status) => out.push(['G5', type, id, q, cond, actual, status, at]);
  const R = rows.map(r => ({ pid: r[0], matched: r[1], name: r[2], pos: r[3], dob: r[7], yexp: r[8],
                             fp: r[4], onCur: r[9], startsCur: r[13] }));
  // Eligibility comes from fantasy_positions (e.g. Travis Hunter: Position=DB, fantasy_positions=[DB,WR]).
  const eligibleOf = x => { try { return JSON.parse(x.fp || '[]').filter(p => LINEUP_POS.indexOf(p) >= 0); } catch (e) { return null; } };
  const lineup = x => { const e = eligibleOf(x); return !!(e && e.length); };
  const pct = (n, d) => d ? `${n}/${d} (${(100 * n / d).toFixed(1)}%)` : 'n/a (0)';

  // HARD
  const uniq = new Set(rows.map(r => r[0])).size;
  add('HARD', 'H1', 'Every requested ID has exactly one row', `${ids.length} rows, no duplicates`,
    `${rows.length} rows, ${uniq} unique`, rows.length === ids.length && uniq === ids.length ? 'PASS' : 'FAIL');
  const unmatched = R.filter(x => !x.matched);
  add('HARD', 'H2', 'Unmatched-ID list emitted in full', 'complete list reported (Matched=FALSE in RAW_PLAYER_META)',
    unmatched.length ? `${unmatched.length}: ${unmatched.map(x => x.pid).join(', ')}` : '0 unmatched', 'PASS');
  const curNoPos = R.filter(x => x.onCur && !x.pos);
  add('HARD', 'H3', 'Every current-roster player has a position (C2 legality)', '0 missing',
    curNoPos.length ? `${curNoPos.length} missing: ${curNoPos.map(x => x.pid).join(', ')}` : '0 missing',
    curNoPos.length ? 'FAIL' : 'PASS');
  const badFp = R.filter(x => x.onCur && eligibleOf(x) === null);
  add('HARD', 'H4', 'Lineup eligibility derivable (fantasy_positions parseable) for every current-roster player', '0 unparseable',
    badFp.length ? `${badFp.length}: ${badFp.map(x => x.pid).join(', ')}` : '0 unparseable', badFp.length ? 'FAIL' : 'PASS');

  // OPERATIONAL (provisional contributor definition until A1 exists)
  const contrib = R.filter(x => x.onCur && lineup(x) && x.startsCur > 0);
  const contribDob = contrib.filter(x => x.dob);
  const okOp = contrib.length && contribDob.length / contrib.length >= 0.95;
  add('OPERATIONAL', 'O1', 'DOB coverage — significant contributors (PROVISIONAL: on current roster, QB/RB/WR/TE, ≥1 start this season)',
    '≥95% (operational gate, not proof the window estimate is valid)', pct(contribDob.length, contrib.length) +
      (contrib.length - contribDob.length ? `; missing: ${contrib.filter(x => !x.dob).map(x => x.name || x.pid).join(', ')}` : ''),
    okOp ? 'PASS' : 'FLAG');

  // INFORMATIONAL
  const curL = R.filter(x => x.onCur && lineup(x)), allL = R.filter(x => lineup(x));
  add('INFO', 'I1', 'DOB coverage — full current roster (lineup-eligible)', 'informational',
    pct(curL.filter(x => x.dob).length, curL.length), 'INFO');
  add('INFO', 'I2', 'DOB coverage — full population (QB/RB/WR/TE, all seasons)', 'informational',
    pct(allL.filter(x => x.dob).length, allL.length), 'INFO');
  add('INFO', 'I3', 'Years_Exp coverage — current roster / full population', 'informational',
    `${pct(curL.filter(x => x.yexp !== '').length, curL.length)} / ${pct(allL.filter(x => x.yexp !== '').length, allL.length)}`, 'INFO');
  const curMissing = curL.filter(x => !x.dob);
  add('INFO', 'I4', 'Current-roster players missing DOB (resolution candidates)', 'informational',
    curMissing.length ? curMissing.map(x => `${x.name || '?'} (${x.pid})`).join(', ') : 'none', 'INFO');
  const nonLineup = R.filter(x => x.onCur && !lineup(x));
  add('INFO', 'I5', 'Current-roster players with NO lineup-eligible fantasy position (can never start)', 'informational',
    nonLineup.length ? nonLineup.map(x => `${x.name || x.pid} [${x.fp}]`).join(', ') : 'none', 'INFO');
  const multi = R.filter(x => x.onCur && x.pos && eligibleOf(x) && eligibleOf(x).join() !== x.pos);
  add('INFO', 'I7', 'Current-roster players whose eligibility differs from Position (C2 must use eligibility)', 'informational',
    multi.length ? multi.map(x => `${x.name || x.pid}: Position ${x.pos} → eligible ${eligibleOf(x).join('/')}`).join(', ') : 'none', 'INFO');
  add('INFO', 'I6', 'Population size', 'informational',
    `${ids.length} IDs: ${Object.keys(hist).length} from RAW_G9, ${Object.keys(curRoster).length} on current rosters`, 'INFO');

  const sh = sheet_(G5.TESTS, true);
  sh.clearContents();
  sh.getRange(1, 1, 1, G5_TEST_HEADERS.length).setValues([G5_TEST_HEADERS]).setFontWeight('bold');
  sh.getRange(2, 1, out.length, G5_TEST_HEADERS.length).setValues(out);
  sh.setFrozenRows(1);

  const hardFail = out.some(t => t[1] === 'HARD' && t[6] === 'FAIL');
  return { seasons: seasonsStr_(cfg), calls: 2, rows: rows.length, status: hardFail ? 'TESTS FAILED' : 'OK',
           notes: `G5 HARD ${hardFail ? 'FAIL' : 'PASS'} | O1 ${out.find(t => t[2] === 'O1')[6]} | ${ids.length} players, ${unmatched.length} unmatched` };
}
