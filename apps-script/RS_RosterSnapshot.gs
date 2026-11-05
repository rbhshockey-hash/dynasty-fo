/***** Roster snapshot — current ownership + roster status (Sleeper) + RS acceptance tests ****************
 * Add as a FIFTH file in the same Apps Script project (reuses helpers from G9_Ingest.gs, G5_PlayerMeta.gs and
 * S0_Stats.gs). Menu item "5b. Pull current roster snapshot + tests" (defined in G9_Ingest.gs onOpen).
 *
 * Population: the CURRENT league (last included CONFIG season), as of the moment it runs.
 * Writes: RAW_ROSTER_SNAPSHOT (one row per roster × player, replaced each run) and RS_TESTS.
 *
 * Stores source facts only: membership in Sleeper's players / starters / reserve / taxi lists exactly as returned,
 * plus one normalized field, Roster_Status (IR > TAXI > ACTIVE; CONFLICT if on both IR and taxi).
 *   - In_Starters is the lineup the manager has SET right now. It records manager behaviour; it is never an input
 *     to lineup strength (C2 builds lineups itself).
 *   - Weekly ownership HISTORY is not stored here: it is reconstructable from RAW_G9 (G9 Q5 PASS).
 *   - Draft picks / free agents (the rest of the asset universe) are decision-layer inputs and are NOT pulled here.
 * The snapshot is always written (raw evidence); RS-GATE says whether C2/A1 may consume it.
 * Run after G5 (RS-6/RS-7 check this snapshot against RAW_PLAYER_META).
 *
 * Pre-registered thresholds:
 *   RS-0 league season = CONFIG current season → else FAIL
 *   RS-1 exactly total_rosters rosters, IDs unique and in 1..N → else FAIL;  RS-1b orphan (no owner) roster → WARN
 *   RS-2 player on two rosters → FAIL
 *   RS-3 IR/taxi player not also in players list → WARN (kept via union, In_Players=FALSE)
 *   RS-4 player on both IR and taxi → FAIL (status ambiguous)
 *   RS-5 active > roster_positions length, IR > reserve_slots, or taxi > taxi_slots → WARN
 *   RS-6 rostered player missing from RAW_PLAYER_META or without a position → FAIL (C2 legality; rerun G5)
 *   RS-7 RAW_PLAYER_META's current-roster view disagrees with this snapshot → WARN (G5 stale; rerun G5)
 *   Gate: FAIL if any FAIL · PASS_RESTRICTED if any WARN · else PASS
 ***************************************************************************************************/
const RS = { SNAP: 'RAW_ROSTER_SNAPSHOT', TESTS: 'RS_TESTS' };
const RS_HEADERS = ['Snapshot_At', 'Season', 'League_ID', 'NFL_Season_At_Snapshot', 'NFL_Week_At_Snapshot', 'Roster_ID',
  'Owner_ID', 'Display_Name', 'Team_Name', 'Player_ID', 'In_Players', 'In_Starters', 'In_Reserve', 'In_Taxi', 'Roster_Status'];

function runRosterSnapshot() { run_('runRosterSnapshot', 'RS', () => rosterSnapshot_()); }

function rosterSnapshot_() {
  const cfg = readConfig_();
  const cur = cfg.leagues[cfg.leagues.length - 1];
  const base = `${API}/league/${cur.leagueId}`;
  const res = fetchAllJson_([`${API}/state/nfl`, base, `${base}/rosters`, `${base}/users`]);
  const state = res[0] || {}, lg = res[1], rosters = res[2] || [], users = res[3] || [];
  if (!lg) throw new Error(`League ${cur.leagueId} (CONFIG season ${cur.season}) returned null. Check CONFIG.`);

  const s = lg.settings || {};
  const totalRosters = Number(lg.total_rosters);
  const rp = lg.roster_positions || [];
  const reserveSlots = Number(s.reserve_slots) || 0, taxiSlots = Number(s.taxi_slots) || 0;
  const uById = {};
  users.forEach(u => { uById[u.user_id] = u; });
  const ids = x => (x || []).map(String).filter(p => p && p !== '0');
  const now = new Date();

  // ---- build snapshot rows ----
  const rows = [], perRoster = [], notInPlayers = [], both = [];
  rosters.forEach(r => {
    const P = ids(r.players), Pset = new Set(P);
    const St = new Set(ids(r.starters)), Rs = new Set(ids(r.reserve)), Tx = new Set(ids(r.taxi));
    const all = [...new Set(P.concat([...Rs], [...Tx]))].sort();
    const u = uById[r.owner_id] || {};
    const teamName = (u.metadata && u.metadata.team_name) || u.display_name || `Roster ${r.roster_id}`;
    const c = { ACTIVE: 0, IR: 0, TAXI: 0, CONFLICT: 0 };
    all.forEach(pid => {
      const inR = Rs.has(pid), inT = Tx.has(pid);
      if (!Pset.has(pid)) notInPlayers.push(`R${r.roster_id}:${pid}`);
      if (inR && inT) both.push(`R${r.roster_id}:${pid}`);
      const status = inR && inT ? 'CONFLICT' : inR ? 'IR' : inT ? 'TAXI' : 'ACTIVE';
      c[status]++;
      rows.push([now, cur.season, cur.leagueId, String(state.season || ''), state.week == null ? '' : state.week,
        r.roster_id, r.owner_id ? String(r.owner_id) : '', u.display_name || '', teamName,
        pid, Pset.has(pid), St.has(pid), inR, inT, status]);
    });
    perRoster.push({ id: r.roster_id, owner: r.owner_id, name: teamName, n: all.length, c });
  });
  writeTableChunked_(RS.SNAP, RS_HEADERS, rows, [3, 4, 7, 10]); // League_ID, NFL season, Owner_ID, Player_ID stay text

  // ---- tests ----
  const out = [], at = new Date();
  const add = (id, q, cond, actual, status) => out.push(['RS', cur.season, id, q, cond, actual, status, at]);

  add('RS-0', 'League season matches CONFIG current season', 'API season == CONFIG season',
    `${lg.season} vs ${cur.season}`, Number(lg.season) === cur.season ? 'PASS' : 'FAIL');

  const rids = rosters.map(r => Number(r.roster_id)).sort((a, b) => a - b);
  const idsOk = rosters.length === totalRosters && new Set(rids).size === rids.length &&
    rids.every(id => id >= 1 && id <= totalRosters);
  add('RS-1', 'One snapshot entry per roster', `${totalRosters} rosters, unique IDs 1..${totalRosters}`,
    `${rosters.length} rosters; IDs ${rids.join(',')}`, idsOk ? 'PASS' : 'FAIL');
  const orphan = perRoster.filter(r => !r.owner);
  add('RS-1b', 'Every roster has an owner', '0 orphaned (WARN: an orphan roster still counts as an opponent)',
    orphan.length ? `${orphan.length}: ${orphan.map(r => `R${r.id}`).join(', ')}` : '0', orphan.length ? 'WARN' : 'PASS');

  const seen = {}, dup = [];
  rows.forEach(r => {
    const pid = r[9], rid = Number(r[5]);
    if (seen[pid] !== undefined && seen[pid] !== rid) dup.push(`${pid} (R${seen[pid]} & R${rid})`); else seen[pid] = rid;
  });
  add('RS-2', 'No player on two rosters', '0', dup.length ? `${dup.length}: ${dup.slice(0, 10).join(', ')}` : '0',
    dup.length ? 'FAIL' : 'PASS');

  add('RS-3', 'IR and taxi players also appear in Sleeper’s players list (expected shape)',
    '0 exceptions (WARN: kept via union, flagged In_Players=FALSE)',
    notInPlayers.length ? `${notInPlayers.length}: ${notInPlayers.slice(0, 10).join(', ')}` : '0',
    notInPlayers.length ? 'WARN' : 'PASS');
  add('RS-4', 'No player on both IR and taxi', '0 (FAIL: status ambiguous)',
    both.length ? `${both.length}: ${both.join(', ')}` : '0', both.length ? 'FAIL' : 'PASS');

  const over = perRoster.filter(r => r.c.ACTIVE > rp.length || r.c.IR > reserveSlots || r.c.TAXI > taxiSlots);
  add('RS-5', 'Roster counts within league limits',
    `active ≤ ${rp.length} (starters + bench) · IR ≤ ${reserveSlots} · taxi ≤ ${taxiSlots}; WARN if over`,
    over.length ? over.map(r => `${r.name}: ${r.c.ACTIVE}/${r.c.IR}/${r.c.TAXI}`).join(' | ') : 'all within limits',
    over.length ? 'WARN' : 'PASS');

  const pm = SpreadsheetApp.getActive().getSheetByName(G5.META);
  if (!pm || pm.getLastRow() < 2) {
    add('RS-6', 'Every rostered player has metadata + position (C2 legality)', '0 missing',
      'RAW_PLAYER_META missing or empty — run G5 first', 'FAIL');
  } else {
    const v = pm.getDataRange().getValues(), h = v[0].map(String);
    const col = n => { const i = h.indexOf(n); if (i < 0) throw new Error(`RAW_PLAYER_META is missing column ${n}. Rerun G5.`); return i; };
    const cP = col('Player_ID'), cN = col('Full_Name'), cPos = col('Position'), cOn = col('On_Current_Roster'), cR = col('Current_Roster_ID');
    const meta = {};
    for (let r = 1; r < v.length; r++) {
      meta[String(v[r][cP]).trim()] = {
        name: String(v[r][cN] || ''), pos: String(v[r][cPos] || ''),
        on: v[r][cOn] === true || String(v[r][cOn]).toUpperCase() === 'TRUE',
        roster: v[r][cR] === '' ? null : Number(v[r][cR])
      };
    }
    const missing = rows.filter(r => !meta[r[9]] || !meta[r[9]].pos).map(r => r[9]);
    add('RS-6', 'Every rostered player has metadata + position in RAW_PLAYER_META (C2 legality)',
      '0 missing (FAIL → rerun G5, then this)', missing.length ? `${missing.length} missing — IDs ${missing.slice(0, 15).join(', ')}` : '0 missing',
      missing.length ? 'FAIL' : 'PASS');
    const disagree = rows.filter(r => meta[r[9]] && meta[r[9]].roster !== Number(r[5]))
      .map(r => `${meta[r[9]].name || r[9]}: meta R${meta[r[9]].roster == null ? '-' : meta[r[9]].roster} vs snapshot R${r[5]}`);
    const metaOnly = Object.keys(meta).filter(p => meta[p].on && seen[p] === undefined)
      .map(p => `${meta[p].name || p}: meta R${meta[p].roster}, not in snapshot`);
    const d = disagree.concat(metaOnly);
    add('RS-7', 'RAW_PLAYER_META current-roster view agrees with this snapshot (G5 freshness)',
      '0 disagreements (WARN → rerun G5 so ages/positions cover the live rosters)',
      d.length ? `${d.length}: ${d.slice(0, 10).join(' | ')}` : '0', d.length ? 'WARN' : 'PASS');
  }

  add('RS-8', 'Per-roster counts (active / IR / taxi)', 'informational',
    perRoster.slice().sort((a, b) => a.id - b.id)
      .map(r => `R${r.id} ${r.name}: ${r.c.ACTIVE}/${r.c.IR}/${r.c.TAXI}`).join(' | '), 'INFO');
  add('RS-9', 'Snapshot timing', 'informational (In_Starters reflects the lineup set as of this moment)',
    `NFL ${state.season} week ${state.week} · taken ${Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')}`, 'INFO');

  const fails = out.filter(t => t[6] === 'FAIL').map(t => t[2]);
  const warns = out.filter(t => t[6] === 'WARN').map(t => t[2]);
  const gate = fails.length ? 'FAIL' : warns.length ? 'PASS_RESTRICTED' : 'PASS';
  add('RS-GATE', 'Roster snapshot acceptable → C2/A1 may consume it',
    'FAIL if any FAIL · PASS_RESTRICTED if any WARN · else PASS',
    fails.length ? `blocked by ${fails.join(', ')}` : warns.length ? `restricted by ${warns.join(', ')}` : 'all conditions met', gate);

  const th = sheet_(RS.TESTS, true);
  th.clearContents();
  th.getRange(1, 1, 1, 8).setValues([['Run', 'Season', 'Test_ID', 'Question', 'Pass_Condition', 'Actual', 'Status', 'Run_At']]).setFontWeight('bold');
  th.getRange(2, 6, out.length, 1).setNumberFormat('@'); // Actual stays text: "1: 13350" must not parse as a duration
  th.getRange(2, 1, out.length, 8).setValues(out);
  th.setFrozenRows(1);

  return { seasons: String(cur.season), calls: 4, rows: rows.length, status: gate === 'FAIL' ? 'TESTS FAILED' : 'OK',
           notes: `RS-GATE ${gate} | ${rosters.length} rosters, ${rows.length} players | NFL ${state.season} wk ${state.week}` };
}
