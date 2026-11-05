/***** Step 0 — Raw weekly stats (Sleeper) + scoring audit + bye classification — v3 ********************
 * Add as a THIRD file in the same Apps Script project (reuses helpers from G9_Ingest.gs and G5_PlayerMeta.gs).
 * Needs: RAW_LEAGUE (scoring), RAW_G9 (player-weeks), RAW_PLAYER_META (positions, names, current NFL team).
 * Writes: RAW_STATS (exactly one row per RAW_G9 player-week, stats stored RAW/unfiltered), RAW_TEAM_WEEKS (team-played
 *         evidence per season-week), S0_BYE_CLASS (the classification Step 2 consumes) and S0_TESTS.
 *
 * Endpoints (undocumented feed; both shapes accepted, the one used is recorded per week):
 *   primary  https://api.sleeper.com/stats/nfl/{season}/{week}?season_type=regular   (array; carries team, opponent)
 *   fallback https://api.sleeper.app/v1/stats/nfl/regular/{season}/{week}            (object; no team → byes undeterminable)
 *
 * Classification (Option A, decided 2026-11-02): BYE | NOT_BYE | UNKNOWN, from team + schedule evidence ONLY.
 *   - GP / gms_active are stored raw but NOTHING depends on them (semantics unvalidated).
 *   - "Team played week w" = some feed row for that team lists an opponent or a non-zero scoring stat.
 *   - Team universe: a season's universe is ESTABLISHED only if the observed teams equal the canonical 32 exactly.
 *   - Player's team that week, in order:
 *       STATS                  own stats row carries a team
 *       INFERRED_CORROBORATED  nearest prior and nearest next observation in the same season agree
 *       INFERRED_SNAPSHOT      latest-ingested-week edge of the CURRENT season only: nearest prior observation
 *                              agrees with today's team in RAW_PLAYER_META (today stands in for "next week")
 *   - NO_BYES_THIS_WEEK      no team established, but the universe is established AND all 32 played that week.
 *   - Anything else → UNKNOWN. Disagreeing or one-sided evidence never yields BYE or NOT_BYE.
 *   - Played-0 and did-not-play weeks are both NOT_BYE and are not distinguished (by design).
 *
 * Pre-registered thresholds (fixed before results):
 *   S0-2 non-zero players_points with no stats row: any → WARN (restricted), listed by name
 *   S0-3 audit per position: non-zero weeks PASS ≥99% within ±0.02 · WARN 95–99% · FAIL <95%;
 *        any zero-point week whose stats imply points caps the position at WARN
 *   S0-3TE on TE weeks with ≥1 reception: PASS only if bonus ×1 is best AND ≥99%
 *   S0-4U team universe = canonical 32 (no missing, no unexpected): else WARN (NO_BYES inference disabled)
 *   S0-4 complete-season claim only with weeks exactly 1–18 AND established universe; any team not absent exactly
 *        once → FAIL (blocks Step 2). Partial/unverified seasons are ALWAYS WARN (structure not fully validated)
 *   Known-team paths (STATS / INFERRED_CORROBORATED / INFERRED_SNAPSHOT): team played → NOT_BYE; team absent → BYE
 *        only if the season universe is established, else UNKNOWN (<source>_UNIVERSE_UNVERIFIED)
 *   S0-3 accuracy is scoped to player-weeks with a stats row; rows missing on zero-point weeks are unaudited
 *   Classifier: a week whose team-played set is empty → UNKNOWN (NO_TEAM_GAME_EVIDENCE), never BYE.
 *   INFERRED_SNAPSHOT applies only to the single latest ingested week of the current season.
 *   S0-5 UNKNOWN share: PASS ≤2% · WARN 2–10% (restricted: UNKNOWN excluded, reported) · FAIL >10%
 *   S0-6 BYE with non-zero points: any → FAIL
 *   S0-9 raw stats too large for a cell (>49,000 chars): any → WARN (row flagged, never silently truncated)
 *   Gate: FAIL if any FAIL · PASS_RESTRICTED if any WARN among gated tests · else PASS
 ***************************************************************************************************/
const S0 = { STATS: 'RAW_STATS', TESTS: 'S0_TESTS', CLASS: 'S0_BYE_CLASS', TEAMWEEKS: 'RAW_TEAM_WEEKS' };
const CLASS_HEADERS = ['Season', 'Week', 'Player_ID', 'Bye_Status', 'NFL_Team', 'Team_Source', 'Classified_At'];
const TEAMWEEK_HEADERS = ['Season', 'Week', 'NFL_Team', 'Game_Evidence', 'Universe_Established', 'Fetched_At'];
const STATS_HEADERS = ['Season', 'Week', 'Player_ID', 'Stats_Row_Found', 'Endpoint', 'NFL_Team', 'Opponent', 'GP',
  'Stats_Raw_JSON', 'Fetched_At'];
const AUDIT_TOL = 0.02;
const NFL_REG_WEEKS = 18;
const CELL_MAX = 49000;
const NFL_TEAMS = ['ARI', 'ATL', 'BAL', 'BUF', 'CAR', 'CHI', 'CIN', 'CLE', 'DAL', 'DEN', 'DET', 'GB', 'HOU', 'IND',
  'JAX', 'KC', 'LAC', 'LAR', 'LV', 'MIA', 'MIN', 'NE', 'NO', 'NYG', 'NYJ', 'PHI', 'PIT', 'SEA', 'SF', 'TB', 'TEN', 'WAS'];

function runStep0() { run_('runStep0', 'S0', () => step0_()); }

function step0_() {
  const t0 = Date.now();
  const cfg = readConfig_();
  const seasons = cfg.leagues.map(l => l.season);
  const curSeason = seasons[seasons.length - 1];

  // ---- scoring per season (verbatim from RAW_LEAGUE) ----
  const lg = sheet_(T.LEAGUE).getDataRange().getValues(), lh = lg[0].map(String);
  const scoring = {};
  for (let r = 1; r < lg.length; r++) {
    scoring[Number(lg[r][lh.indexOf('Config_Season')])] = JSON.parse(lg[r][lh.indexOf('Scoring_Settings_JSON')] || '{}');
  }

  // ---- positions, names, current NFL team ----
  const pm = sheet_(G5.META).getDataRange().getValues(), ph = pm[0].map(String);
  const posOf = {}, nameOf = {}, snapTeam = {};
  for (let r = 1; r < pm.length; r++) {
    const pid = String(pm[r][ph.indexOf('Player_ID')]);
    posOf[pid] = String(pm[r][ph.indexOf('Position')]);
    nameOf[pid] = String(pm[r][ph.indexOf('Full_Name')]);
    snapTeam[pid] = String(pm[r][ph.indexOf('NFL_Team')] || '');
  }

  // ---- RAW_G9 player-weeks (the population) ----
  const g9 = sheet_(T.G9).getDataRange().getValues(), gh = g9[0].map(String), gi = n => gh.indexOf(n);
  const cS = gi('Season'), cW = gi('Week'), cP = gi('Player_ID'), cPts = gi('Points');
  const g9pw = [], pidSet = {}, weeksBySeason = {};
  for (let r = 1; r < g9.length; r++) {
    const pid = String(g9[r][cP]).trim();
    if (!pid) continue;
    const s = Number(g9[r][cS]), w = Number(g9[r][cW]);
    pidSet[pid] = true;
    (weeksBySeason[s] = weeksBySeason[s] || {})[w] = true;
    g9pw.push({ s, w, pid, pts: g9[r][cPts] === '' ? null : Number(g9[r][cPts]) });
  }
  const lastWeekCur = Math.max.apply(null, Object.keys(weeksBySeason[curSeason] || { 0: 1 }).map(Number));

  // ---- fetch stats week by week (chunked, memory-safe) ----
  const reqs = [];
  Object.keys(weeksBySeason).forEach(s => Object.keys(weeksBySeason[s]).forEach(w => reqs.push({ s: Number(s), w: Number(w) })));
  reqs.sort((a, b) => a.s - b.s || a.w - b.w);
  const statMap = {}, teamsPlayed = {}, weekInfo = {}, teamsSeen = {};
  const CH = 4;
  for (let i = 0; i < reqs.length; i += CH) {
    if (Date.now() - t0 > 300000) {
      throw new Error(`Stopped before ${reqs[i].s} wk${reqs[i].w} to stay under the 6-minute limit. Ask Claude to split the run by season.`);
    }
    const chunk = reqs.slice(i, i + CH);
    const prim = UrlFetchApp.fetchAll(chunk.map(q => ({
      url: `https://api.sleeper.com/stats/nfl/${q.s}/${q.w}?season_type=regular`, muteHttpExceptions: true })));
    const parsed = chunk.map((q, j) => tryParseStats_(prim[j], 'api.sleeper.com'));
    const need = parsed.map((p, j) => p ? -1 : j).filter(j => j >= 0);
    if (need.length) {
      const fb = UrlFetchApp.fetchAll(need.map(j => ({
        url: `${API}/stats/nfl/regular/${chunk[j].s}/${chunk[j].w}`, muteHttpExceptions: true })));
      need.forEach((j, k) => { parsed[j] = tryParseStats_(fb[k], 'api.sleeper.app/v1'); });
    }
    chunk.forEach((q, j) => {
      const p = parsed[j], key = `${q.s}|${q.w}`;
      if (!p) { weekInfo[key] = { s: q.s, w: q.w, endpoint: 'NONE', n: 0, hasTeam: false }; teamsPlayed[key] = null; return; }
      const tp = {};
      let hasTeam = false;
      teamsSeen[q.s] = teamsSeen[q.s] || {};
      Object.keys(p.rows).forEach(pid => {
        const r = p.rows[pid];
        if (r.team) { hasTeam = true; teamsSeen[q.s][r.team] = true; }
        if (r.team && teamGameEvidence_(r, scoring[q.s] || {})) tp[r.team] = true;
        if (pidSet[pid]) statMap[`${key}|${pid}`] = r;
      });
      teamsPlayed[key] = hasTeam ? tp : null;
      weekInfo[key] = { s: q.s, w: q.w, endpoint: p.endpoint, n: Object.keys(p.rows).length, hasTeam };
    });
  }

  // ---- team universe per season ----
  const universe = {};
  seasons.forEach(S => {
    const seen = Object.keys(teamsSeen[S] || {});
    const missing = NFL_TEAMS.filter(t => seen.indexOf(t) < 0), unexpected = seen.filter(t => NFL_TEAMS.indexOf(t) < 0);
    universe[S] = { ok: !missing.length && !unexpected.length, missing, unexpected, n: seen.length };
  });

  // ---- RAW_STATS: exactly one row per RAW_G9 player-week, raw stats unfiltered ----
  const now = new Date();
  let oversize = 0;
  const oversizeEx = [];
  const statRows = g9pw.map(x => {
    const key = `${x.s}|${x.w}`, r = statMap[`${key}|${x.pid}`], wi = weekInfo[key] || { endpoint: 'NONE' };
    let raw = r ? JSON.stringify(r.stats) : '';
    if (raw.length > CELL_MAX) { oversize++; if (oversizeEx.length < 10) oversizeEx.push(`${x.s} wk${x.w} ${x.pid}`); raw = 'OVERSIZE_NOT_STORED'; }
    return [x.s, x.w, x.pid, !!r, wi.endpoint, r ? r.team : '', r ? r.opp : '',
      r && r.stats.gp != null ? r.stats.gp : '', raw, now];
  });
  writeTableChunked_(S0.STATS, STATS_HEADERS, statRows, [3, 9]);

  // ---- classification of every RAW_G9 player-week ----
  const teamByPW = buildTeamIndex_(statMap);
  const ctx = { statMap, teamsPlayed, teamByPW, universe, curSeason, lastWeekCur, snapTeam };
  const cls = g9pw.map(x => Object.assign({ x }, classifyBye_(x.s, x.w, x.pid, ctx)));

  // ---- persist the classification and the team-week evidence it was built from (Step 2 joins these) ----
  writeTableChunked_(S0.CLASS, CLASS_HEADERS,
    cls.map(c => [c.x.s, c.x.w, c.x.pid, c.status, c.team || '', c.teamSource, now]), [3]);
  const twRows = [];
  Object.keys(teamsPlayed).sort().forEach(k => {
    const tp = teamsPlayed[k];
    if (!tp) return;
    const [ss, ww] = k.split('|').map(Number);
    const teams = [...new Set(NFL_TEAMS.concat(Object.keys(tp)))].sort();
    teams.forEach(t => twRows.push([ss, ww, t, !!tp[t], !!(universe[ss] && universe[ss].ok), now]));
  });
  writeTableChunked_(S0.TEAMWEEKS, TEAMWEEK_HEADERS, twRows, []);

  // ---- tests ----
  const out = [], at = new Date();
  const add = (season, id, q, cond, actual, status) => out.push(['S0', season, id, q, cond, actual, status, at]);
  const pct = (n, d) => d ? `${n}/${d} (${(100 * n / d).toFixed(2)}%)` : 'n/a (0)';
  const row = x => statMap[`${x.s}|${x.w}|${x.pid}`];
  const POS = ['QB', 'RB', 'WR', 'TE'];

  // S0-1 endpoint / shape
  seasons.forEach(S => {
    const W = Object.values(weekInfo).filter(x => x.s === S);
    const none = W.filter(x => x.endpoint === 'NONE'), noTeam = W.filter(x => x.endpoint !== 'NONE' && !x.hasTeam);
    add(S, 'S0-1', 'Stats feed returned data for every ingested week', 'all weeks non-empty (FAIL); team field present (WARN if not)',
      `${W.length - none.length}/${W.length} weeks via ${[...new Set(W.map(x => x.endpoint))].join(', ')}` +
        (none.length ? `; EMPTY: wk ${none.map(x => x.w).join(',')}` : '') +
        (noTeam.length ? `; no team field: wk ${noTeam.map(x => x.w).join(',')}` : ''),
      none.length ? 'FAIL' : (noTeam.length ? 'WARN' : 'PASS'));
  });

  // S0-2 coverage by season × position; non-zero points without a stats row → WARN (gated)
  seasons.forEach(S => {
    const parts = [], miss = [];
    POS.concat(['OTHER']).forEach(pos => {
      const P = g9pw.filter(x => x.s === S && (pos === 'OTHER' ? POS.indexOf(posOf[x.pid]) < 0 : posOf[x.pid] === pos));
      if (!P.length) return;
      const hit = P.filter(row).length;
      const m = P.filter(x => x.pts && !row(x));
      miss.push.apply(miss, m);
      parts.push(`${pos} ${hit}/${P.length}`);
    });
    add(S, 'S0-2', 'Stats-row coverage by position; non-zero players_points must have a stats row',
      'any non-zero-point week without a stats row → WARN (restricted), listed',
      `${parts.join(' · ')} · non-zero without row: ${miss.length}` +
        (miss.length ? ` (${miss.slice(0, 15).map(x => `wk${x.w} ${nameOf[x.pid] || x.pid} ${x.pts}`).join(' | ')})` : ''),
      miss.length ? 'WARN' : 'PASS');
  });

  // S0-3 scoring audit by position — coverage, non-zero accuracy, zero-point contradictions
  const auditStatus = {};
  POS.forEach(pos => {
    const all = g9pw.filter(x => posOf[x.pid] === pos);
    const audited = all.filter(x => row(x) && x.pts != null);
    const calc = x => scorePts_(row(x).stats, scoring[x.s] || {}, pos, 1);
    const nz = audited.filter(x => x.pts !== 0);
    const miss = nz.filter(x => Math.abs(calc(x) - x.pts) > AUDIT_TOL);
    const zMiss = audited.filter(x => x.pts === 0 && Math.abs(calc(x)) > AUDIT_TOL);
    const rate = nz.length ? (nz.length - miss.length) / nz.length : 0;
    let st = !nz.length ? 'FAIL' : rate >= 0.99 ? 'PASS' : rate >= 0.95 ? 'WARN' : 'FAIL';
    if (st === 'PASS' && zMiss.length) st = 'WARN';
    auditStatus[pos] = st;
    add('ALL', `S0-3${pos}`, `Scoring audit — ${pos}: stats × confirmed settings reproduces players_points`,
      'Accuracy on player-weeks WITH a stats row only: non-zero weeks PASS ≥99% (±0.02) · WARN 95–99% · FAIL <95%; any zero-point contradiction caps at WARN. Coverage reported separately; non-zero weeks lacking a row restrict via S0-2; zero-point weeks lacking a row are UNAUDITED, not assumed correct',
      `coverage: eligible ${all.length} · with stats row ${all.filter(row).length} · unaudited (no row) ${all.length - all.filter(row).length} | accuracy: audited ${audited.length} · non-zero ${nz.length} · ` +
        `match ${pct(nz.length - miss.length, nz.length)} · zero-point contradictions ${zMiss.length}` +
        (miss.length ? ` · mismatches: ${miss.slice(0, 8).map(x => `${x.s} wk${x.w} ${nameOf[x.pid] || x.pid} sleeper ${x.pts} vs calc ${calc(x)}`).join(' | ')}` : '') +
        (zMiss.length ? ` · contradictions: ${zMiss.slice(0, 8).map(x => `${x.s} wk${x.w} ${nameOf[x.pid] || x.pid} calc ${calc(x)}`).join(' | ')}` : ''),
      st);
  });
  const other = g9pw.filter(x => POS.indexOf(posOf[x.pid]) < 0);
  add('ALL', 'S0-3OTHER', 'Player-weeks at non-QB/RB/WR/TE positions (not audited)', 'informational',
    `${other.length} player-weeks: ${[...new Set(other.map(x => `${nameOf[x.pid] || x.pid} [${posOf[x.pid] || '?'}]`))].join(', ') || 'none'}`, 'INFO');

  const teRec = g9pw.filter(x => posOf[x.pid] === 'TE' && x.pts != null && row(x) && (Number(row(x).stats.rec) || 0) > 0);
  const rateAt = m => teRec.length ? teRec.filter(x =>
    Math.abs(scorePts_(row(x).stats, scoring[x.s] || {}, 'TE', m) - x.pts) <= AUDIT_TOL).length / teRec.length : 0;
  const r0 = rateAt(0), r1 = rateAt(1), r2 = rateAt(2);
  const teOk = r1 >= 0.99 && r1 > r0 && r1 > r2;
  add('ALL', 'S0-3TE', 'TE premium applied exactly once (competing hypotheses, TE weeks with ≥1 reception)',
    'PASS only if ×1 is best AND ≥99%',
    `n=${teRec.length} · bonus ×0: ${(100 * r0).toFixed(1)}% · ×1: ${(100 * r1).toFixed(1)}% · ×2: ${(100 * r2).toFixed(1)}%`,
    teRec.length && teOk ? 'PASS' : 'FAIL');

  // S0-4U team universe, S0-4 bye structure
  seasons.forEach(S => {
    const u = universe[S];
    add(S, 'S0-4U', 'Team universe = canonical 32 NFL teams', 'no missing, no unexpected (else WARN; NO_BYES inference disabled)',
      `${u.n} observed · missing: ${u.missing.join(',') || 'none'} · unexpected: ${u.unexpected.join(',') || 'none'}`,
      u.ok ? 'PASS' : 'WARN');
    const keys = Object.keys(teamsPlayed).filter(k => Number(k.split('|')[0]) === S && teamsPlayed[k] &&
      Number(k.split('|')[1]) <= NFL_REG_WEEKS);
    if (!keys.length) { add(S, 'S0-4', 'Bye structure recoverable', 'team data present', 'no team data', 'FAIL'); return; }
    const absent = {};
    NFL_TEAMS.forEach(t => { absent[t] = keys.filter(k => !teamsPlayed[k][t]).length; });
    const dist = {};
    Object.values(absent).forEach(n => { dist[n] = (dist[n] || 0) + 1; });
    const observedWeeks = keys.map(k => Number(k.split('|')[1])).sort((x, y) => x - y);
    const complete = observedWeeks.length === NFL_REG_WEEKS && observedWeeks.every((wk, i) => wk === i + 1) && u.ok;
    if (complete) {
      const bad = NFL_TEAMS.filter(t => absent[t] !== 1);
      add(S, 'S0-4', 'Bye structure — COMPLETE season (18 weeks, 32-team universe): each team absent exactly once',
        'all 32 absent exactly 1 week (any deviation → FAIL: schedule evidence inconsistent)',
        `weeks-absent distribution ${JSON.stringify(dist)}` + (bad.length ? `; off: ${bad.map(t => `${t}=${absent[t]}`).join(',')}` : ''),
        bad.length ? 'FAIL' : 'PASS');
    } else {
      const bad = NFL_TEAMS.filter(t => absent[t] > 1);
      add(S, 'S0-4', `Bye structure — PROVISIONAL (${keys.length}/${NFL_REG_WEEKS} weeks${u.ok ? '' : ', universe NOT established'}): no team absent more than once`,
        'always WARN until the season is complete (structure not fully validated); Actual shows whether it looks plausible', `weeks-absent distribution ${JSON.stringify(dist)}` +
          (bad.length ? `; off: ${bad.map(t => `${t}=${absent[t]}`).join(',')}` : '; structure plausible so far'), 'WARN');
    }
  });

  // S0-5 classification · S0-6 impossible byes · S0-7 GP diagnostic
  seasons.forEach(S => {
    const C = cls.filter(c => c.x.s === S);
    const n = { BYE: 0, NOT_BYE: 0, UNKNOWN: 0 }, src = {};
    C.forEach(c => { n[c.status]++; src[c.teamSource] = (src[c.teamSource] || 0) + 1; });
    const unk = C.length ? n.UNKNOWN / C.length : 1;
    add(S, 'S0-5', 'Bye classification of RAW_G9 player-weeks', 'UNKNOWN: PASS ≤2% · WARN 2–10% (restricted) · FAIL >10%',
      `BYE ${n.BYE} · NOT_BYE ${n.NOT_BYE} · UNKNOWN ${n.UNKNOWN} (${(100 * unk).toFixed(2)}%) · team source ${JSON.stringify(src)}`,
      unk <= 0.02 ? 'PASS' : unk <= 0.10 ? 'WARN' : 'FAIL');
    const imp = C.filter(c => c.status === 'BYE' && c.x.pts);
    add(S, 'S0-6', 'No BYE week with non-zero points', '0 (any → FAIL: classifier is wrong)',
      `${imp.length}` + (imp.length ? `: ${imp.slice(0, 10).map(c => `wk${c.x.w} ${nameOf[c.x.pid] || c.x.pid} ${c.x.pts}pts`).join(' | ')}` : ''),
      imp.length ? 'FAIL' : 'PASS');
    const nb0 = C.filter(c => c.status === 'NOT_BYE' && c.x.pts === 0);
    const gpPos = nb0.filter(c => row(c.x) && Number(row(c.x).stats.gp) > 0).length;
    const gpZero = nb0.filter(c => row(c.x) && row(c.x).stats.gp != null && Number(row(c.x).stats.gp) === 0).length;
    const noRow = nb0.filter(c => !row(c.x)).length;
    add(S, 'S0-7', 'NOT_BYE zero-point weeks by raw GP (diagnostic only — nothing depends on GP)', 'informational (input for a possible v2)',
      `${nb0.length} weeks: GP>0 ${gpPos} · GP=0 ${gpZero} · GP absent ${nb0.length - gpPos - gpZero - noRow} · no stats row ${noRow}`, 'INFO');
  });

  // S0-8 manual-inspection samples (most recent first)
  const sample = (label, f) => {
    const s = cls.filter(f).sort((a, b) => b.x.s - a.x.s || b.x.w - a.x.w).slice(0, 8);
    add('ALL', 'S0-8', `Inspection sample — ${label}`, 'manual check',
      s.length ? s.map(c => `${c.x.s} wk${c.x.w} ${nameOf[c.x.pid] || c.x.pid} ${c.team || '?'}(${c.teamSource}) ${c.x.pts}pts`).join(' | ') : 'none', 'INFO');
  };
  sample('BYE', c => c.status === 'BYE');
  sample('BYE via snapshot corroboration', c => c.status === 'BYE' && c.teamSource === 'INFERRED_SNAPSHOT');
  sample('NOT_BYE, zero points', c => c.status === 'NOT_BYE' && c.x.pts === 0);
  sample('NOT_BYE, positive points', c => c.status === 'NOT_BYE' && c.x.pts > 0);
  sample('UNKNOWN', c => c.status === 'UNKNOWN');

  // S0-9 raw storage integrity
  add('ALL', 'S0-9', 'Raw stats stored unfiltered; none too large for a cell', `0 rows over ${CELL_MAX} chars (else WARN, flagged)`,
    `${oversize} oversize` + (oversizeEx.length ? `: ${oversizeEx.join(' | ')}` : ''), oversize ? 'WARN' : 'PASS');

  // gate
  const gatedIds = ['S0-1', 'S0-2', 'S0-3', 'S0-4', 'S0-5', 'S0-6', 'S0-9'];
  const gated = out.filter(t => t[6] !== 'INFO' && gatedIds.some(g => String(t[2]).indexOf(g) === 0));
  const fails = gated.filter(t => t[6] === 'FAIL').map(t => `${t[2]}(${t[1]})`);
  const warns = gated.filter(t => t[6] === 'WARN').map(t => `${t[2]}(${t[1]})`);
  const gate = fails.length ? 'FAIL' : warns.length ? 'PASS_RESTRICTED' : 'PASS';
  add('ALL', 'S0-GATE', 'Step 0 acceptable → Step 2 may build C1_PLAYER_WEEKS',
    'FAIL if any FAIL · PASS_RESTRICTED if any WARN (UNKNOWN excluded + reported; issues named) · else PASS',
    fails.length ? `blocked by ${fails.join(', ')}` : warns.length ? `restricted by ${warns.join(', ')}` : 'all gate conditions met', gate);

  const sh = sheet_(S0.TESTS, true);
  sh.clearContents();
  sh.getRange(1, 1, 1, 8).setValues([['Run', 'Season', 'Test_ID', 'Question', 'Pass_Condition', 'Actual', 'Status', 'Run_At']]).setFontWeight('bold');
  sh.getRange(2, 1, out.length, 8).setValues(out);
  sh.setFrozenRows(1);

  return { seasons: seasonsStr_(cfg), calls: reqs.length, rows: statRows.length, status: gate === 'FAIL' ? 'TESTS FAILED' : 'OK',
           notes: `S0-GATE ${gate} | audit QB:${auditStatus.QB} RB:${auditStatus.RB} WR:${auditStatus.WR} TE:${auditStatus.TE} TEx1:${teOk ? 'PASS' : 'FAIL'}` };
}

/* ---------------- shared helpers (reused by Step 2) ---------------- */
function tryParseStats_(resp, endpoint) {
  if (resp.getResponseCode() !== 200) return null;
  const txt = resp.getContentText();
  if (!txt) return null;
  let j;
  try { j = JSON.parse(txt); } catch (e) { return null; }
  const rows = {};
  if (Array.isArray(j)) {
    j.forEach(x => {
      if (!x || x.player_id == null) return;
      rows[String(x.player_id)] = { team: x.team || '', opp: x.opponent || '', stats: x.stats || {} };
    });
  } else if (j && typeof j === 'object') {
    Object.keys(j).forEach(pid => {
      const st = j[pid];
      if (st && typeof st === 'object') rows[String(pid)] = { team: '', opp: '', stats: st };
    });
  }
  return Object.keys(rows).length ? { endpoint, rows } : null;
}

// Team-played evidence without relying on GP: an opponent is listed, or any scoring stat is non-zero.
function teamGameEvidence_(r, sc) {
  if (r.opp) return true;
  return Object.keys(sc).some(k => Number(r.stats[k]));
}

// TE bonus always computed from receptions × multiplier (never a pre-baked key) so ×0/×1/×2 are clean hypotheses.
function scorePts_(stats, sc, pos, teMult) {
  let t = 0;
  Object.keys(sc).forEach(k => {
    if (k === 'bonus_rec_te') return;
    const v = Number(stats[k]);
    if (v) t += v * Number(sc[k]);
  });
  const b = Number(sc.bonus_rec_te) || 0;
  if (pos === 'TE' && b) t += teMult * b * (Number(stats.rec) || 0);
  return Math.round(t * 100) / 100;
}

function buildTeamIndex_(statMap) {
  const idx = {};
  Object.keys(statMap).forEach(k => {
    const [s, w, pid] = k.split('|');
    const team = statMap[k].team;
    if (!team) return;
    ((idx[pid] = idx[pid] || {})[s] = idx[pid][s] || {})[w] = team;
  });
  return idx;
}

// ctx: { statMap, teamsPlayed, teamByPW, universe, curSeason, lastWeekCur, snapTeam }
// A team present in the week's played set establishes NOT_BYE. An absent team establishes BYE only when the
// season's canonical 32-team universe is established; otherwise absence may be feed coverage → UNKNOWN.
function classifyKnownTeam_(team, tp, universe, source) {
  if (tp[team]) return { status: 'NOT_BYE', team, teamSource: source };
  if (universe && universe.ok) return { status: 'BYE', team, teamSource: source };
  return { status: 'UNKNOWN', team, teamSource: `${source}_UNIVERSE_UNVERIFIED` };
}

function classifyBye_(s, w, pid, ctx) {
  s = Number(s); w = Number(w);
  const tp = ctx.teamsPlayed[`${s}|${w}`];
  if (!tp) return { status: 'UNKNOWN', team: '', teamSource: 'NO_TEAM_DATA' };
  if (!Object.keys(tp).length) return { status: 'UNKNOWN', team: '', teamSource: 'NO_TEAM_GAME_EVIDENCE' };
  const u = ctx.universe[s];
  const r = ctx.statMap[`${s}|${w}|${pid}`];
  if (r && r.team) return classifyKnownTeam_(r.team, tp, u, 'STATS');

  const ws = ctx.teamByPW[pid] && ctx.teamByPW[pid][s];
  const wk = ws ? Object.keys(ws).map(Number) : [];
  const prev = wk.filter(x => x < w), next = wk.filter(x => x > w);
  const a = prev.length ? ws[Math.max.apply(null, prev)] : '';
  if (a && next.length) {
    const b = ws[Math.min.apply(null, next)];
    if (a === b) return classifyKnownTeam_(a, tp, u, 'INFERRED_CORROBORATED');
  } else if (a && !next.length && s === ctx.curSeason && w === ctx.lastWeekCur &&
             ctx.snapTeam[pid] && ctx.snapTeam[pid] === a) {
    // trailing edge of the current season: today's team stands in for "next week"
    return classifyKnownTeam_(a, tp, u, 'INFERRED_SNAPSHOT');
  }

  if (u && u.ok && NFL_TEAMS.every(t => tp[t])) return { status: 'NOT_BYE', team: '', teamSource: 'NO_BYES_THIS_WEEK' };
  return { status: 'UNKNOWN', team: '', teamSource: 'NONE' };
}

function writeTableChunked_(name, headers, rows, textCols) {
  const sh = sheet_(name, true);
  sh.clearContents();
  const n = Math.max(rows.length + 1, 2);
  (textCols || []).forEach(c => sh.getRange(1, c, n, 1).setNumberFormat('@'));
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  const B = 2000;
  for (let i = 0; i < rows.length; i += B) {
    const part = rows.slice(i, i + B);
    sh.getRange(2 + i, 1, part.length, headers.length).setValues(part);
  }
  sh.setFrozenRows(1);
}
