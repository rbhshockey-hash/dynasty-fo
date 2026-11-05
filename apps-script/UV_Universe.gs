/***** Asset universe + contention outcomes (Sleeper) + UV acceptance tests *********************************
 * Add as a SIXTH file in the same Apps Script project (reuses helpers from G9_Ingest.gs, G5_PlayerMeta.gs and
 * S0_Stats.gs). Menu item "5c. Pull universe: brackets, picks, free agents + tests" (G9_Ingest.gs onOpen).
 * Run after G5 (free agents are flagged with whether they have league history in RAW_PLAYER_META).
 *
 * Writes (all source facts; nothing interpretive):
 *   RAW_BRACKETS      every configured season's winners + losers bracket, verbatim (A2 standing: who contended).
 *                     Consumers must use COMPLETED seasons only (RAW_LEAGUE Status = complete): Sleeper pre-fills a
 *                     provisional bracket during the season, which is not an outcome.
 *   RAW_DRAFTS        every configured league's drafts: status, rounds, draft order, slot → roster (verbatim)
 *   RAW_TRADED_PICKS  current league's traded-pick records, verbatim (roster_id = ORIGINAL owner, owner_id = CURRENT)
 *   PICK_OWNERSHIP    normalized grid: one row per future season × round × original roster, with its current owner.
 *                     Rule (league mechanics, not judgement): a pick belongs to its original roster unless a
 *                     traded-pick record says otherwise. Draft_Slot is filled only when Sleeper has set the order —
 *                     never projected from standings.
 *   RAW_FREE_AGENTS   QB/RB/WR/TE-eligible players NOT on any current roster who are on an NFL team (incl. practice
 *                     squad / IR). Unsigned players are excluded — D2 must disclose that exclusion. Has_League_History = present in RAW_PLAYER_META (i.e. C1 has own-history for them;
 *                     everyone else can only get the position prior — Test Plan D8).
 *   UV_TESTS
 *
 * CONFIG parameter (optional): Pick_Seasons_Ahead (default 3). Sleeper does not expose how many future seasons
 * are tradeable; any traded pick outside the window is reported by UV-P2, never silently dropped.
 *
 * Pre-registered thresholds:
 *   UV-B1 completed season: winners bracket present AND champion decided → else FAIL; playoff-team count ≠
 *         settings.playoff_teams → WARN.  In-progress season: INFO only.
 *   UV-B2 any bracket roster ID outside 1..N → FAIL
 *   UV-P1 draft_rounds known → else FAIL (grid not built)
 *   UV-P2 live traded pick beyond the season/round window → WARN (listed); records for already-drafted seasons → INFO
 *   UV-P3 two traded-pick records for the same pick with different owners → FAIL
 *   UV-P4 any owner roster ID outside 1..N → FAIL
 *   UV-F1 current rosters fetched = total_rosters → else FAIL;  UV-F2 free-agent list overlaps a roster → FAIL
 *   Gate: FAIL if any FAIL · PASS_RESTRICTED if any WARN · else PASS
 ***************************************************************************************************/
const UV = { BRACKETS: 'RAW_BRACKETS', DRAFTS: 'RAW_DRAFTS', TRADED: 'RAW_TRADED_PICKS', PICKS: 'PICK_OWNERSHIP',
             FA: 'RAW_FREE_AGENTS', TESTS: 'UV_TESTS' };
const UV_BRACKET_HEADERS = ['Season', 'League_ID', 'Bracket', 'Round', 'Match_ID', 'Team1_Roster_ID', 'Team2_Roster_ID',
  'Winner_Roster_ID', 'Loser_Roster_ID', 'Place', 'Team1_From_JSON', 'Team2_From_JSON', 'Fetched_At'];
const UV_DRAFT_HEADERS = ['Config_Season', 'League_ID', 'Draft_ID', 'Draft_Season', 'Status', 'Type', 'Rounds',
  'Draft_Order_JSON', 'Slot_To_Roster_JSON', 'Start_Time', 'Fetched_At'];
const UV_TRADED_HEADERS = ['Season', 'Round', 'Original_Roster_ID', 'Previous_Owner_Roster_ID', 'Owner_Roster_ID', 'Fetched_At'];
const UV_PICK_HEADERS = ['Season', 'Round', 'Original_Roster_ID', 'Original_Team', 'Owner_Roster_ID', 'Owner_Team',
  'Is_Traded', 'Draft_Slot', 'Slot_Known', 'Snapshot_At'];
const UV_FA_HEADERS = ['Player_ID', 'Full_Name', 'Position', 'Fantasy_Positions_JSON', 'NFL_Team', 'Status', 'Active',
  'Birth_Date', 'Years_Exp', 'Has_League_History', 'Snapshot_At'];

function runUniverse() { run_('runUniverse', 'UV', () => universe_()); }

function uvParam_(name, dflt) {
  const v = sheet_(T.CONFIG).getDataRange().getValues();
  for (let r = 1; r < v.length; r++) if (String(v[r][4]).trim() === name && v[r][5] !== '') return v[r][5];
  return dflt;
}

function universe_() {
  const cfg = readConfig_();
  const cur = cfg.leagues[cfg.leagues.length - 1];
  const ahead = Math.max(1, Math.floor(Number(uvParam_('Pick_Seasons_Ahead', 3)) || 3));

  // ---- fetch ----
  const urls = [];
  cfg.leagues.forEach(l => urls.push(`${API}/league/${l.leagueId}`, `${API}/league/${l.leagueId}/winners_bracket`,
    `${API}/league/${l.leagueId}/losers_bracket`, `${API}/league/${l.leagueId}/drafts`));
  const base = `${API}/league/${cur.leagueId}`;
  urls.push(`${base}/rosters`, `${base}/users`, `${base}/traded_picks`);
  const res = fetchAllJson_(urls);
  const k = cfg.leagues.length * 4;
  const rosters = res[k] || [], users = res[k + 1] || [], traded = res[k + 2] || [];
  const players = fetchAllJson_([`${API}/players/nfl`])[0] || {};
  const calls = urls.length + 1;
  const now = new Date();

  // ---- team names: per-season from RAW_ROSTERS (history), current from users ----
  const nameOf = {};
  const rr = SpreadsheetApp.getActive().getSheetByName(T.ROSTERS);
  if (rr && rr.getLastRow() > 1) {
    const v = rr.getDataRange().getValues(), h = v[0].map(String);
    const cS = h.indexOf('Season'), cR = h.indexOf('Roster_ID'), cN = h.indexOf('Team_Name');
    for (let r = 1; r < v.length; r++) nameOf[`${Number(v[r][cS])}|${Number(v[r][cR])}`] = String(v[r][cN] || '');
  }
  const uById = {};
  users.forEach(u => { uById[u.user_id] = u; });
  const curName = {};
  rosters.forEach(r => {
    const u = uById[r.owner_id] || {};
    curName[Number(r.roster_id)] = (u.metadata && u.metadata.team_name) || u.display_name || `Roster ${r.roster_id}`;
  });
  const nm = (s, rid) => nameOf[`${s}|${rid}`] || (s === cur.season ? curName[rid] : '') || `R${rid}`;

  const out = [], at = new Date();
  const add = (season, id, q, cond, actual, status) => out.push(['UV', season, id, q, cond, actual, status, at]);
  const nz = x => (x == null ? '' : x);
  const isId = x => typeof x === 'number';

  // ---- brackets + drafts, every configured season ----
  const bRows = [], dRows = [];
  let curLeague = null, curDrafts = [];
  cfg.leagues.forEach((l, i) => {
    const lg = res[i * 4] || {}, wb = Array.isArray(res[i * 4 + 1]) ? res[i * 4 + 1] : [],
          lb = Array.isArray(res[i * 4 + 2]) ? res[i * 4 + 2] : [], drafts = Array.isArray(res[i * 4 + 3]) ? res[i * 4 + 3] : [];
    if (l === cur) { curLeague = lg; curDrafts = drafts; }
    const S = l.season, N = Number(lg.total_rosters) || 0, st = lg.settings || {};
    [['W', wb], ['L', lb]].forEach(([b, arr]) => arr.forEach(x => bRows.push([S, l.leagueId, b, nz(x.r), nz(x.m),
      nz(x.t1), nz(x.t2), nz(x.w), nz(x.l), nz(x.p), x.t1_from ? JSON.stringify(x.t1_from) : '',
      x.t2_from ? JSON.stringify(x.t2_from) : '', now])));
    drafts.forEach(d => dRows.push([S, l.leagueId, String(d.draft_id), String(d.season || ''), d.status || '', d.type || '',
      (d.settings && d.settings.rounds) || '', d.draft_order ? JSON.stringify(d.draft_order) : '',
      d.slot_to_roster_id ? JSON.stringify(d.slot_to_roster_id) : '', d.start_time || '', now]));

    const teams = [...new Set(wb.reduce((a, x) => a.concat([x.t1, x.t2]), []).filter(isId))].sort((a, b) => a - b);
    const champ = wb.find(x => Number(x.p) === 1);
    const champId = champ && isId(champ.w) ? champ.w : null;
    const nPO = Number(st.playoff_teams) || null;
    if (lg.status === 'complete') {
      const ok = wb.length && champId;
      add(S, 'UV-B1', 'Playoff bracket present and resolved (completed season)',
        `winners bracket non-empty + champion decided (FAIL) · ${nPO || '?'} playoff teams (WARN if not)`,
        `${wb.length} matches · ${teams.length} playoff teams · champion ${champId ? `R${champId} ${nm(S, champId)}` : 'none'}`,
        !ok ? 'FAIL' : (nPO && teams.length !== nPO ? 'WARN' : 'PASS'));
    } else {
      add(S, 'UV-B1', `Playoff bracket PROVISIONAL (season status: ${lg.status || '?'})`,
        'informational — Sleeper pre-fills an in-season bracket; it is NOT a contention outcome and must not be consumed',
        `${wb.length} provisional winners-bracket matches`, 'INFO');
    }
    const bad = wb.concat(lb).reduce((a, x) => a.concat([x.t1, x.t2, x.w, x.l]), []).filter(t => isId(t) && (t < 1 || t > N));
    add(S, 'UV-B2', 'Bracket roster IDs in range', `all in 1..${N}`, `${bad.length} out of range`, bad.length ? 'FAIL' : 'PASS');
    if (teams.length && lg.status === 'complete') add(S, 'UV-B3', 'Playoff teams + champion (manual check against Sleeper)', 'informational',
      `${teams.map(t => `R${t} ${nm(S, t)}`).join(' | ')}` + (champId ? ` · CHAMPION R${champId} ${nm(S, champId)}` : ''), 'INFO');
  });
  writeTableChunked_(UV.BRACKETS, UV_BRACKET_HEADERS, bRows, [2]);
  writeTableChunked_(UV.DRAFTS, UV_DRAFT_HEADERS, dRows, [2, 3, 4, 8, 9]);

  // ---- picks: current league ----
  const N = Number(curLeague.total_rosters) || 0;
  const rounds = Number((curLeague.settings || {}).draft_rounds) || null;
  const curSeason = Number(curLeague.season);
  const thisDraft = curDrafts.find(d => Number(d.season) === curSeason);
  const first = thisDraft && thisDraft.status === 'complete' ? curSeason + 1 : curSeason; // next undrafted season
  const seasons = [];
  for (let y = first; y < first + ahead; y++) seasons.push(y);

  writeTableChunked_(UV.TRADED, UV_TRADED_HEADERS, traded.map(t => [Number(t.season), t.round, t.roster_id,
    nz(t.previous_owner_id), t.owner_id, now]), []);

  add('ALL', 'UV-P1', 'Draft rounds known (league settings)', 'draft_rounds present (else FAIL; grid not built)',
    `${rounds || 'missing'} rounds · window ${seasons[0]}–${seasons[seasons.length - 1]} (Pick_Seasons_Ahead=${ahead})`,
    rounds ? 'PASS' : 'FAIL');

  const pRows = [];
  if (rounds) {
    const owner = {}, applied = {}, outside = [], used = [], conflicts = [];
    seasons.forEach(y => { for (let r = 1; r <= rounds; r++) for (let o = 1; o <= N; o++) owner[`${y}|${r}|${o}`] = o; });
    traded.forEach(t => {
      const key = `${Number(t.season)}|${Number(t.round)}|${Number(t.roster_id)}`;
      const label = `${t.season} R${t.round} orig R${t.roster_id}→R${t.owner_id}`;
      if (Number(t.season) < first) { used.push(label); return; } // already drafted: history, not a live pick
      if (!(key in owner)) { outside.push(label); return; }
      if (applied[key] !== undefined && applied[key] !== Number(t.owner_id)) conflicts.push(key);
      applied[key] = Number(t.owner_id);
      owner[key] = Number(t.owner_id);
    });
    add('ALL', 'UV-P2', 'Every live traded-pick record falls inside the season/round window',
      '0 beyond the window or round count (WARN, listed — not dropped silently)',
      outside.length ? `${outside.length} outside: ${outside.slice(0, 10).join(', ')}` : `0 of ${traded.length - used.length} live records`,
      outside.length ? 'WARN' : 'PASS');
    add('ALL', 'UV-P2h', 'Traded-pick records for seasons already drafted (history; kept in RAW_TRADED_PICKS)', 'informational',
      used.length ? `${used.length}: ${used.join(', ')}` : '0', 'INFO');
    add('ALL', 'UV-P3', 'No pick has conflicting traded-pick records', '0 (FAIL)', conflicts.length ? `${conflicts.length}: ${conflicts.join(', ')}` : '0',
      conflicts.length ? 'FAIL' : 'PASS');

    // draft slots only where Sleeper has set them
    const slotOf = {};
    curDrafts.forEach(d => {
      const m = d.slot_to_roster_id || {};
      Object.keys(m).forEach(slot => { if (m[slot] != null) slotOf[`${Number(d.season)}|${Number(m[slot])}`] = Number(slot); });
    });
    Object.keys(owner).sort((a, b) => {
      const A = a.split('|').map(Number), B = b.split('|').map(Number);
      return A[0] - B[0] || A[1] - B[1] || A[2] - B[2];
    }).forEach(key => {
      const [y, r, o] = key.split('|').map(Number), ow = owner[key];
      const slot = slotOf[`${y}|${o}`];
      pRows.push([y, r, o, curName[o] || `R${o}`, ow, curName[ow] || `R${ow}`, ow !== o, slot == null ? '' : slot, slot != null, now]);
    });
    const badOwner = pRows.filter(p => !(p[4] >= 1 && p[4] <= N)).length;
    add('ALL', 'UV-P4', 'Every pick owner is a valid roster', `all in 1..${N}`, `${badOwner} invalid`, badOwner ? 'FAIL' : 'PASS');
    const byTeam = {};
    pRows.forEach(p => {
      const t = byTeam[p[4]] = byTeam[p[4]] || { all: 0, firsts: 0 };
      t.all++; if (p[1] === 1) t.firsts++;
    });
    add('ALL', 'UV-P5', `Picks owned per team, ${seasons[0]}–${seasons[seasons.length - 1]} (total / 1sts)`, 'informational',
      Object.keys(byTeam).map(Number).sort((a, b) => a - b).map(o => `R${o} ${curName[o] || ''}: ${byTeam[o].all}/${byTeam[o].firsts}`).join(' | '), 'INFO');
    add('ALL', 'UV-P6', 'Draft slots known', 'informational (blank until Sleeper sets the order; never projected)',
      `${pRows.filter(p => p[8]).length} of ${pRows.length} picks have a slot`, 'INFO');
  }
  writeTableChunked_(UV.PICKS, UV_PICK_HEADERS, pRows, []);

  // ---- free agents ----
  add('ALL', 'UV-F1', 'All current rosters fetched', `${N} rosters`, `${rosters.length} rosters`, rosters.length === N ? 'PASS' : 'FAIL');
  const rostered = new Set();
  rosters.forEach(r => [].concat(r.players || [], r.reserve || [], r.taxi || []).forEach(p => rostered.add(String(p))));
  const history = new Set();
  const pm = SpreadsheetApp.getActive().getSheetByName(G5.META);
  if (pm && pm.getLastRow() > 1) pm.getRange(2, 1, pm.getLastRow() - 1, 1).getValues().forEach(r => history.add(String(r[0]).trim()));
  const faRows = [];
  Object.keys(players).forEach(pid => {
    const p = players[pid] || {};
    const elig = (p.fantasy_positions || []).filter(x => LINEUP_POS.indexOf(x) >= 0);
    if (!elig.length || rostered.has(pid)) return;
    if (!p.team) return; // on an NFL roster/practice squad; Sleeper's 'active' flag alone includes ~2,300 unsigned players
    faRows.push([pid, p.full_name || [p.first_name, p.last_name].filter(Boolean).join(' '), p.position || '',
      JSON.stringify(p.fantasy_positions || []), p.team || '', p.status || '', p.active === true,
      p.birth_date ? String(p.birth_date) : '', p.years_exp != null ? p.years_exp : '', history.has(pid), now]);
  });
  faRows.sort((a, b) => (a[2] < b[2] ? -1 : a[2] > b[2] ? 1 : 0) || (a[1] < b[1] ? -1 : 1));
  writeTableChunked_(UV.FA, UV_FA_HEADERS, faRows, [1, 8]);
  const overlap = faRows.filter(r => rostered.has(r[0])).length;
  add('ALL', 'UV-F2', 'No free agent is on a current roster', '0 (FAIL)', `${overlap}`, overlap ? 'FAIL' : 'PASS');
  const byPos = {};
  faRows.forEach(r => { byPos[r[2]] = (byPos[r[2]] || 0) + 1; });
  const bySt = {};
  faRows.forEach(r => { bySt[r[5] || '(none)'] = (bySt[r[5] || '(none)'] || 0) + 1; });
  add('ALL', 'UV-F3', 'Free-agent pool by position and status', 'informational (QB/RB/WR/TE-eligible, on an NFL team; unsigned players excluded)',
    `${faRows.length} players · ${Object.keys(byPos).sort().map(p => `${p} ${byPos[p]}`).join(' · ')} | status: ` +
      Object.keys(bySt).sort().map(s => `${s} ${bySt[s]}`).join(' · '), 'INFO');
  const withHist = faRows.filter(r => r[9]).length;
  add('ALL', 'UV-F4', 'Free agents with league production history (C1 own-data possible)', 'informational (others → position prior only, D8)',
    `${withHist} of ${faRows.length}` + (history.size ? '' : ' — RAW_PLAYER_META missing, run G5'), 'INFO');

  // ---- gate + write tests ----
  const fails = out.filter(t => t[6] === 'FAIL').map(t => `${t[2]}(${t[1]})`);
  const warns = out.filter(t => t[6] === 'WARN').map(t => `${t[2]}(${t[1]})`);
  const gate = fails.length ? 'FAIL' : warns.length ? 'PASS_RESTRICTED' : 'PASS';
  add('ALL', 'UV-GATE', 'Universe tables acceptable → A2 standing reference / decision layer may consume them',
    'FAIL if any FAIL · PASS_RESTRICTED if any WARN · else PASS',
    fails.length ? `blocked by ${fails.join(', ')}` : warns.length ? `restricted by ${warns.join(', ')}` : 'all conditions met', gate);

  const th = sheet_(UV.TESTS, true);
  th.clearContents();
  th.getRange(1, 1, 1, 8).setValues([['Run', 'Season', 'Test_ID', 'Question', 'Pass_Condition', 'Actual', 'Status', 'Run_At']]).setFontWeight('bold');
  th.getRange(2, 6, out.length, 1).setNumberFormat('@'); // Actual stays text
  th.getRange(2, 1, out.length, 8).setValues(out);
  th.setFrozenRows(1);

  return { seasons: seasonsStr_(cfg), calls, rows: bRows.length + dRows.length + traded.length + pRows.length + faRows.length,
           status: gate === 'FAIL' ? 'TESTS FAILED' : 'OK',
           notes: `UV-GATE ${gate} | brackets ${bRows.length} · drafts ${dRows.length} · traded ${traded.length} · picks ${pRows.length} · FA ${faRows.length}` };
}
