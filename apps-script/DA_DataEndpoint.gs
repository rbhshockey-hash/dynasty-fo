/***** Public read-only data endpoint for the site (GitHub Pages) ********************************************
 * Add as an EIGHTH file in the same Apps Script project, then Deploy → New deployment → type "Web app":
 *   Execute as: Me · Who has access: Anyone. Copy the /exec URL into docs/config.js (DATA_URL).
 * Re-deploy (Manage deployments → edit → new version) after changing this file.
 *
 * WHAT IS PUBLISHED (explicit whitelist — nothing else in the spreadsheet is reachable through this endpoint):
 *   health     gate status + last run per pipeline stage (from the *_TESTS tabs and RUN_LOG)
 *   rosters    current rosters (RAW_ROSTER_SNAPSHOT ⨝ RAW_PLAYER_META): name, eligibility, NFL team, age, roster status
 *   picks      PICK_OWNERSHIP (future picks and current owners)
 *   contention completed seasons only: playoff teams, champion, runner-up (RAW_BRACKETS ⨝ RAW_ROSTERS)
 * WHAT IS NEVER PUBLISHED: RAW_MARKET_VALUES (FantasyCalc — third-party data, decision layer only), raw stats,
 *   raw player-weeks, CONFIG. Only the MV gate STATUS is shown, never a value.
 *
 * This endpoint only selects and joins already-tested tables; it computes no interpretation. The one derived
 * field is Age (from Sleeper birth_date, as of the request date) — a display convenience, not a model input.
 ***************************************************************************************************/
const DA_GATES = [
  { stage: 'G9 ingestion (weekly production)', tab: 'G9_TESTS', id: 'G9-ACCEPT', multi: true },
  { stage: 'G5 player metadata', tab: 'G5_TESTS', id: null },
  { stage: 'Roster snapshot', tab: 'RS_TESTS', id: 'RS-GATE' },
  { stage: 'Universe (brackets, picks, free agents)', tab: 'UV_TESTS', id: 'UV-GATE' },
  { stage: 'Market values (gate status only)', tab: 'MV_TESTS', id: 'MV-GATE' },
  { stage: 'Step 0 (stats audit + byes)', tab: 'S0_TESTS', id: 'S0-GATE' },
  { stage: 'Step 2 (C1 player-weeks)', tab: 'C1D_TESTS', id: 'C1D-GATE' }
];

function doGet() {
  const tz = Session.getScriptTimeZone();
  const body = {
    generatedAt: new Date().toISOString(),
    health: daHealth_(),
    rosters: daRosters_(tz),
    picks: daPicks_(),
    contention: daContention_()
  };
  return ContentService.createTextOutput(JSON.stringify(body)).setMimeType(ContentService.MimeType.JSON);
}

function daTab_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return null;
  const v = sh.getDataRange().getValues(), h = v[0].map(String);
  return { rows: v.slice(1), c: n => h.indexOf(n) };
}

function daHealth_() {
  const gates = [];
  DA_GATES.forEach(g => {
    const t = daTab_(g.tab);
    if (!t) { gates.push({ stage: g.stage, status: 'MISSING', detail: `${g.tab} not found — step not run`, runAt: null }); return; }
    const cId = 2, cAct = 5, cSt = 6, cAt = 7; // all *_TESTS tabs share: Run|Mode, Season|Type, Test_ID, Q, Cond, Actual, Status, Run_At
    if (g.id === null) { // G5: gate = any HARD FAIL
      const hard = t.rows.filter(r => r[1] === 'HARD');
      const fail = hard.some(r => r[cSt] === 'FAIL');
      gates.push({ stage: g.stage, status: fail ? 'FAIL' : 'PASS', detail: `${hard.length} hard checks`, runAt: t.rows[0][cAt] });
      return;
    }
    const hits = t.rows.filter(r => r[cId] === g.id);
    if (!hits.length) { gates.push({ stage: g.stage, status: 'MISSING', detail: `${g.id} not found`, runAt: null }); return; }
    if (g.multi) {
      hits.filter(r => r[0] !== 'SAMPLE').forEach(r => gates.push({ stage: `${g.stage} · ${r[0]} ${r[1]}`, status: String(r[cSt]), detail: String(r[cAct]), runAt: r[cAt] }));
    } else {
      const r = hits[hits.length - 1];
      gates.push({ stage: g.stage, status: String(r[cSt]), detail: String(r[cAct]), runAt: r[cAt] });
    }
  });

  const log = daTab_('RUN_LOG'), last = {};
  if (log) log.rows.forEach(r => {
    last[r[1]] = { fn: String(r[1]), at: r[0], status: String(r[6]), rows: r[5], notes: String(r[7]).slice(0, 300) };
  });
  return { gates, lastRuns: Object.keys(last).map(k => last[k]) };
}

function daRosters_(tz) {
  const snap = daTab_('RAW_ROSTER_SNAPSHOT'), meta = daTab_('RAW_PLAYER_META');
  if (!snap || !meta) return null;
  const m = {};
  meta.rows.forEach(r => {
    let elig = [];
    try { elig = JSON.parse(r[meta.c('Fantasy_Positions_JSON')] || '[]').filter(p => ['QB', 'RB', 'WR', 'TE'].indexOf(p) >= 0); } catch (e) {}
    m[String(r[meta.c('Player_ID')]).trim()] = {
      name: String(r[meta.c('Full_Name')] || ''), pos: String(r[meta.c('Position')] || ''), elig,
      nflTeam: String(r[meta.c('NFL_Team')] || ''), dob: String(r[meta.c('Birth_Date')] || ''),
      yearsExp: r[meta.c('Years_Exp')]
    };
  });
  const today = new Date(Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd') + 'T00:00:00');
  const age = dob => {
    if (!/^\d{4}-\d{2}-\d{2}/.test(dob)) return null;
    const b = new Date(dob.slice(0, 10) + 'T00:00:00');
    let a = today.getFullYear() - b.getFullYear();
    if (today.getMonth() < b.getMonth() || (today.getMonth() === b.getMonth() && today.getDate() < b.getDate())) a--;
    return a;
  };
  const teams = {};
  let snapshotAt = null;
  snap.rows.forEach(r => {
    const rid = Number(r[snap.c('Roster_ID')]);
    const t = teams[rid] = teams[rid] || { rosterId: rid, team: String(r[snap.c('Team_Name')] || ''),
      owner: String(r[snap.c('Display_Name')] || ''), players: [] };
    snapshotAt = snapshotAt || r[snap.c('Snapshot_At')];
    const pid = String(r[snap.c('Player_ID')]).trim(), p = m[pid] || {};
    t.players.push({ id: pid, name: p.name || pid, pos: p.pos || '', elig: p.elig || [], nflTeam: p.nflTeam || '',
      age: age(p.dob || ''), yearsExp: p.yearsExp === '' ? null : p.yearsExp,
      status: String(r[snap.c('Roster_Status')]), setAsStarter: r[snap.c('In_Starters')] === true });
  });
  return { snapshotAt, teams: Object.keys(teams).map(k => teams[k]).sort((a, b) => a.rosterId - b.rosterId) };
}

function daPicks_() {
  const t = daTab_('PICK_OWNERSHIP');
  if (!t) return null;
  return t.rows.map(r => ({
    season: r[t.c('Season')], round: r[t.c('Round')],
    originalRosterId: r[t.c('Original_Roster_ID')], originalTeam: String(r[t.c('Original_Team')]),
    ownerRosterId: r[t.c('Owner_Roster_ID')], ownerTeam: String(r[t.c('Owner_Team')]),
    traded: r[t.c('Is_Traded')] === true, slot: r[t.c('Draft_Slot')] === '' ? null : r[t.c('Draft_Slot')]
  }));
}

function daContention_() {
  const lg = daTab_('RAW_LEAGUE'), br = daTab_('RAW_BRACKETS'), ro = daTab_('RAW_ROSTERS');
  if (!lg || !br) return null;
  const name = {};
  if (ro) ro.rows.forEach(r => { name[`${Number(r[ro.c('Season')])}|${Number(r[ro.c('Roster_ID')])}`] = String(r[ro.c('Team_Name')] || ''); });
  const nm = (s, id) => (id === '' || id == null) ? null : { rosterId: Number(id), team: name[`${s}|${Number(id)}`] || `R${id}` };
  const out = [];
  lg.rows.forEach(r => {
    const S = Number(r[lg.c('Config_Season')]);
    if (String(r[lg.c('Status')]) !== 'complete') { out.push({ season: S, complete: false }); return; }
    const W = br.rows.filter(b => Number(b[br.c('Season')]) === S && b[br.c('Bracket')] === 'W');
    const ids = new Set();
    W.forEach(b => [b[br.c('Team1_Roster_ID')], b[br.c('Team2_Roster_ID')]].forEach(x => { if (x !== '') ids.add(Number(x)); }));
    const final = W.find(b => Number(b[br.c('Place')]) === 1);
    out.push({ season: S, complete: true,
      playoffTeams: [...ids].sort((a, b) => a - b).map(id => nm(S, id)),
      champion: final ? nm(S, final[br.c('Winner_Roster_ID')]) : null,
      runnerUp: final ? nm(S, final[br.c('Loser_Roster_ID')]) : null });
  });
  return out.sort((a, b) => b.season - a.season);
}
