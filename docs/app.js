// Dynasty Front Office — read-only data room.
// Renders the JSON served by apps-script/DA_DataEndpoint.gs. Presentation only: it sorts, groups and formats,
// and never computes a measure. Anything interpretive belongs in the sheet, behind a test gate.
(function () {
  const cfg = window.DFO_CONFIG || {};
  const view = document.getElementById('view');
  const POS_ORDER = { QB: 0, RB: 1, WR: 2, TE: 3 };
  const VIEWS = { health: renderHealth, rosters: renderRosters, picks: renderPicks, contention: renderContention, roadmap: renderRoadmap };
  let data = null, selectedTeam = cfg.MY_ROSTER_ID || null;

  document.getElementById('league-sub').textContent = cfg.LEAGUE_NAME || '';

  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const pill = s => `<span class="pill ${esc(s)}">${esc(String(s).replace('_', ' '))}</span>`;
  const when = d => { if (!d) return '—'; const t = new Date(d); return isNaN(t) ? esc(d) : t.toLocaleString(); };
  const isMine = id => Number(id) === Number(cfg.MY_ROSTER_ID);

  function route() {
    const name = (location.hash || '#health').slice(1);
    const fn = VIEWS[name] || renderHealth;
    document.querySelectorAll('#nav a').forEach(a => a.classList.toggle('active', a.getAttribute('href') === '#' + (VIEWS[name] ? name : 'health')));
    if (name === 'roadmap') { view.innerHTML = fn(); return; }
    if (!data) return;
    view.innerHTML = fn();
    bind();
  }

  function bind() {
    view.querySelectorAll('[data-team]').forEach(b => b.addEventListener('click', () => {
      selectedTeam = Number(b.getAttribute('data-team')); route();
    }));
  }

  function teamChips(teams) {
    return `<div class="chips">${teams.map(t => `<button class="chip" data-team="${t.rosterId}" aria-pressed="${t.rosterId === selectedTeam}">
      ${esc(t.team || 'R' + t.rosterId)}${isMine(t.rosterId) ? ' ★' : ''}</button>`).join('')}</div>`;
  }

  // ---------- views ----------
  function renderHealth() {
    const h = data.health || { gates: [], lastRuns: [] };
    const bad = h.gates.filter(g => g.status === 'FAIL' || g.status === 'MISSING');
    return `<h2>Pipeline health</h2>
      <p class="lede">Every stage is acceptance-tested before anything downstream may use it. PASS_RESTRICTED means usable,
        with named limits (e.g. the 2026 season is still in progress).</p>
      ${bad.length ? `<div class="callout"><strong>${bad.length} stage(s) blocked.</strong> Downstream views may be stale until fixed.</div>` : ''}
      <div class="panel"><table>
        <thead><tr><th>Stage</th><th>Gate</th><th>Detail</th><th>Tested</th></tr></thead>
        <tbody>${h.gates.map(g => `<tr><td>${esc(g.stage)}</td><td>${pill(g.status)}</td>
          <td class="detail">${esc(g.detail)}</td><td class="muted">${when(g.runAt)}</td></tr>`).join('')}</tbody>
      </table></div>
      <h3>Last run of each step</h3>
      <div class="panel"><table>
        <thead><tr><th>Step</th><th>Status</th><th class="num">Rows</th><th>Notes</th><th>When</th></tr></thead>
        <tbody>${h.lastRuns.map(r => `<tr><td><code>${esc(r.fn)}</code></td><td>${pill(r.status)}</td>
          <td class="num">${esc(r.rows)}</td><td class="detail">${esc(r.notes)}</td><td class="muted">${when(r.at)}</td></tr>`).join('')}</tbody>
      </table></div>`;
  }

  function renderRosters() {
    const R = data.rosters;
    if (!R) return missing('RAW_ROSTER_SNAPSHOT / RAW_PLAYER_META');
    if (!R.teams.some(t => t.rosterId === selectedTeam)) selectedTeam = R.teams[0] && R.teams[0].rosterId;
    const t = R.teams.find(x => x.rosterId === selectedTeam);
    const players = t.players.slice().sort((a, b) =>
      (POS_ORDER[a.elig[0]] ?? 9) - (POS_ORDER[b.elig[0]] ?? 9) || (a.status > b.status ? 1 : -1) || a.name.localeCompare(b.name));
    const count = s => t.players.filter(p => p.status === s).length;
    return `<h2>Rosters</h2>
      <p class="lede">Current Sleeper rosters as of the last snapshot (${when(R.snapshotAt)}). Age is from Sleeper birth dates.
        “Set to start” is the lineup the manager has set right now — not a strength judgement.</p>
      ${teamChips(R.teams)}
      <div class="panel">
        <h3 style="margin-top:0">${esc(t.team)} <span class="muted">· ${esc(t.owner)}</span></h3>
        <p class="muted">${count('ACTIVE')} active · ${count('IR')} IR · ${count('TAXI')} taxi</p>
        <table>
          <thead><tr><th>Player</th><th>Eligible</th><th>NFL</th><th class="num">Age</th><th class="num">Yrs exp</th><th>Roster</th><th>Set to start</th></tr></thead>
          <tbody>${players.map(p => `<tr><td>${esc(p.name)}</td><td>${esc(p.elig.join('/') || p.pos)}${p.elig.join('/') !== p.pos && p.pos ? `<span class="tag">${esc(p.pos)}</span>` : ''}</td>
            <td>${esc(p.nflTeam || 'FA')}</td><td class="num">${p.age ?? '—'}</td><td class="num">${p.yearsExp ?? '—'}</td>
            <td>${p.status === 'ACTIVE' ? '' : `<span class="tag">${esc(p.status)}</span>`}</td><td>${p.setAsStarter ? '✓' : ''}</td></tr>`).join('')}</tbody>
        </table>
      </div>`;
  }

  function renderPicks() {
    const P = data.picks;
    if (!P) return missing('PICK_OWNERSHIP');
    const teams = {};
    P.forEach(p => { (teams[p.ownerRosterId] = teams[p.ownerRosterId] || { rosterId: p.ownerRosterId, team: p.ownerTeam, picks: [] }).picks.push(p); });
    const list = Object.values(teams).sort((a, b) => a.rosterId - b.rosterId);
    const seasons = [...new Set(P.map(p => p.season))].sort();
    return `<h2>Pick ledger</h2>
      <p class="lede">Who owns each future rookie pick (${seasons.join('–')}). A pick belongs to its original team unless a Sleeper trade record moved it.
        Draft slots appear only once Sleeper sets the order — never projected.</p>
      <div class="panel"><table>
        <thead><tr><th>Team</th>${seasons.map(s => `<th>${s}</th>`).join('')}<th class="num">Total</th></tr></thead>
        <tbody>${list.map(t => `<tr class="${isMine(t.rosterId) ? 'mine' : ''}"><td>${esc(t.team)}</td>
          ${seasons.map(s => `<td>${t.picks.filter(p => p.season === s).sort((a, b) => a.round - b.round || a.originalRosterId - b.originalRosterId)
            .map(p => `R${p.round}${p.traded ? `<span class="tag" title="originally ${esc(p.originalTeam)}">via ${esc(p.originalTeam)}</span>` : ''}${p.slot ? ` (#${p.slot})` : ''}`).join('<br>')}</td>`).join('')}
          <td class="num">${t.picks.length}</td></tr>`).join('')}</tbody>
      </table></div>`;
  }

  function renderContention() {
    const C = data.contention;
    if (!C) return missing('RAW_BRACKETS / RAW_LEAGUE');
    const done = C.filter(c => c.complete), live = C.filter(c => !c.complete);
    return `<h2>Contention record</h2>
      <p class="lede">Who actually made the playoffs and won, from Sleeper’s final brackets. Completed seasons only —
        Sleeper’s in-season bracket is provisional and is not shown.</p>
      <div class="callout">Sample size: ${done.length} completed season(s) × 12 teams. Too thin to define “what a contender looks like” with confidence; it grows by one season a year.</div>
      <div class="grid">${done.map(c => `<div class="panel">
        <h3 style="margin-top:0">${c.season}</h3>
        <p><span class="muted">Champion</span><br><span class="stat">${esc(c.champion ? c.champion.team : '—')}</span></p>
        <p class="muted">Runner-up: ${esc(c.runnerUp ? c.runnerUp.team : '—')}</p>
        <p class="muted" style="margin-bottom:.2rem">Playoff field</p>
        <ul class="plain">${c.playoffTeams.map(t => `<li${isMine(t.rosterId) ? ' style="font-weight:700"' : ''}>${esc(t.team)}</li>`).join('')}</ul>
      </div>`).join('')}</div>
      ${live.length ? `<p class="muted">In progress: ${live.map(c => c.season).join(', ')}.</p>` : ''}`;
  }

  function renderRoadmap() {
    const rows = [
      ['Data layer — Sleeper ingestion, metadata, rosters, picks, brackets, free agents', 'Built and acceptance-tested', 'PASS'],
      ['C1 — canonical player-week dataset (bench included, byes classified)', 'Built; restricted only by the in-progress season', 'PASS_RESTRICTED'],
      ['C1 — forward production distributions (method tests, Step 3)', 'Pre-registered test plan awaiting sign-off', 'INFO'],
      ['C2 — lineup strength distributions (own + 11 opponents)', 'Pending C1 verdict', 'INFO'],
      ['A1 — field-relative strength and weakness by slot', 'Pending C2', 'INFO'],
      ['A2 — phase: standing, trajectory, window', 'Pending A1; trajectory may be unresolvable from two seasons', 'INFO'],
      ['D0–D4 — decision layer (posture, trade evaluation, acquisition search, prioritisation)', 'Pending A2. Uses consensus market values internally; they are never published here', 'INFO']
    ];
    return `<h2>Roadmap</h2>
      <p class="lede">The system is built bottom-up: no layer is shown here until the tests beneath it pass.
        Two currencies are kept separate throughout — realized production drives strength; market value only enters the decision layer.</p>
      <div class="panel"><table><thead><tr><th>Component</th><th>State</th><th></th></tr></thead>
        <tbody>${rows.map(r => `<tr><td>${esc(r[0])}</td><td class="detail">${esc(r[1])}</td><td>${pill(r[2] === 'INFO' ? 'PENDING' : r[2])}</td></tr>`).join('')}</tbody>
      </table></div>`;
  }

  function missing(what) { return `<div class="callout">Not available yet: ${esc(what)} has not been built. Check Pipeline health.</div>`; }

  // ---------- load ----------
  window.addEventListener('hashchange', route);
  if (!cfg.DATA_URL) {
    view.innerHTML = `<h2>Setup needed</h2><p class="lede">Set <code>DATA_URL</code> in <code>docs/config.js</code> to the deployed
      Apps Script web-app URL (see <code>apps-script/DA_DataEndpoint.gs</code>).</p>${renderRoadmap()}`;
    return;
  }
  fetch(cfg.DATA_URL).then(r => { if (!r.ok) throw new Error('HTTP ' + r.status); return r.json(); })
    .then(j => { data = j; document.getElementById('generated').textContent = 'Data served ' + when(j.generatedAt); route(); })
    .catch(e => { view.innerHTML = `<div class="callout">Could not load data from the sheet endpoint (${esc(e.message)}).</div>${renderRoadmap()}`; });
  route();
})();
