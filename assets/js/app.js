/* USP Brain Research Dashboard
   Theme-first navigation: an Overview tab plus one tab per dashboard theme.
   Each theme tab carries its own KPIs, charts and faceted filters. */
(function () {
  "use strict";

  const CONFIG = {
    dataFile: "data/USP_brain_research_dashboard_data.xlsx",
    sheet: "Dashboard data",
    brainVideo: "brain/brain-rotating.mp4",
    brainPoster: "brain/brain-hero-poster.jpg",
  };

  // Source column headers (must match the Excel header row).
  const COL = {
    title: "Research / project title",
    theme: "General brain research theme",
    institute: "Main institute",
    department: "Involved internal department / research group",
    collaborators: "Collaborating departments / partners",
    website: "Website",
    innovation: "Utrecht Brain innovation theme",
    domain: "Research domain",
    disease: "Disease / condition / application",
    methods: "Methods & technologies",
    population: "Population / model system",
    stage: "Research stage",
    description: "Dashboard description",
    tags: "Tags / subthemes",
    contact: "Main contact / onderzoeker / contactafdeling",
    recordType: "Record type",
  };

  // Facets shown inside a theme view (theme itself is fixed by the active tab).
  const FACETS = [
    { key: "innovation", label: "Innovation theme", multi: true },
    { key: "recordType", label: "Record type", multi: false },
    { key: "stage", label: "Research stage", multi: false },
    { key: "domain", label: "Research domain", multi: true },
    { key: "institute", label: "Institute", multi: false },
    { key: "collab", label: "Collaboration", multi: false },
  ];

  const HOME = "__home__";
  const OVERVIEW = "__overview__";

  const PROJECT_TEAM = {
    name: "Wiring Brain Research Group 2",
    members: ["Emma Durac", "Fiene Agterberg", "Femke Janssen", "Brandon Teerlink", "Myrthe Wiegers"],
  };

  const state = {
    projects: [],
    themes: [],            // [name, count] ordered by count desc
    activeTab: HOME,
    filters: { innovation: "", recordType: "", stage: "", domain: "", institute: "", collab: "", q: "" },
    charts: [],
    rebuildCharts: null,   // set per theme view so pill removal can refresh charts
  };

  const $ = (id) => document.getElementById(id);
  const view = () => $("view");

  // --- Data ---------------------------------------------------------------
  function splitMulti(value) {
    if (!value) return [];
    return String(value).split(/[;\n]/).map((s) => s.trim()).filter(Boolean);
  }

  // Resolve each logical field to a column index by header text, tolerant to
  // whitespace/case. The dashboard-theme header has been blanked before, so it
  // falls back to the column right after the title when its header is missing.
  function resolveColumns(header) {
    const norm = (s) => String(s == null ? "" : s).trim().toLowerCase();
    const idx = {};
    for (const key in COL) idx[key] = header.findIndex((h) => norm(h) === norm(COL[key]));
    if (idx.theme < 0 && idx.title >= 0) {
      idx.theme = idx.title + 1;
      console.warn("Theme column header not found — using the column after the title (col " + (idx.theme + 1) + ").");
    }
    return idx;
  }

  async function loadData() {
    const res = await fetch(CONFIG.dataFile);
    if (!res.ok) throw new Error(`Could not load ${CONFIG.dataFile} (${res.status})`);
    const wb = XLSX.read(await res.arrayBuffer(), { type: "array" });
    const ws = wb.Sheets[CONFIG.sheet] || wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", blankrows: false });
    if (!rows.length) return [];
    const idx = resolveColumns(rows[0]);
    const cell = (r, k) => (idx[k] >= 0 ? String(r[idx[k]] == null ? "" : r[idx[k]]).trim() : "");

    return rows.slice(1)
      .map((r) => {
        const collabRaw = cell(r, "collaborators");
        const hasCollab = collabRaw !== "" && !/^(no|none|n\/a|-|geen)$/i.test(collabRaw);
        return {
          title: cell(r, "title"),
          theme: cell(r, "theme"),
          institute: cell(r, "institute"),
          department: cell(r, "department"),
          collaborators: hasCollab ? collabRaw : "",
          collab: hasCollab ? "Collaboration" : "No Collaboration",
          website: cell(r, "website"),
          innovation: cell(r, "innovation"),
          domain: cell(r, "domain"),
          disease: cell(r, "disease"),
          methods: cell(r, "methods"),
          population: cell(r, "population"),
          stage: cell(r, "stage"),
          description: cell(r, "description"),
          tags: splitMulti(cell(r, "tags")),
          contact: cell(r, "contact"),
          recordType: cell(r, "recordType"),
        };
      })
      .filter((p) => p.title);
  }

  // --- Aggregation --------------------------------------------------------
  function countBy(projects, key, multi) {
    const counts = new Map();
    for (const p of projects) {
      const values = multi ? splitMulti(p[key]) : (p[key] ? [p[key]] : []);
      for (const v of values) counts.set(v, (counts.get(v) || 0) + 1);
    }
    return [...counts.entries()].sort((a, b) => b[1] - a[1]);
  }

  function valuesOf(project, key, multi) {
    return multi ? splitMulti(project[key]) : (project[key] ? [project[key]] : []);
  }

  // Projects belonging to the active tab (whole theme, before facet filters).
  function themeProjects() {
    if (state.activeTab === OVERVIEW) return state.projects;
    // Multi-valued theme: include a project if the active theme is one of its themes.
    return state.projects.filter((p) => valuesOf(p, "theme", true).includes(state.activeTab));
  }

  // Apply the facet + search filters on top of the theme subset. Passing
  // `except` skips one facet — used so a chart can show its own full set of
  // options while still reflecting the filters chosen on the other charts
  // (cross-filtering).
  function applyFacets(projects, except) {
    const needle = state.filters.q.trim().toLowerCase();
    return projects.filter((p) => {
      for (const f of FACETS) {
        if (f.key === except) continue;
        const sel = state.filters[f.key];
        if (sel && !valuesOf(p, f.key, f.multi).includes(sel)) return false;
      }
      if (needle) {
        const hay = [p.title, p.description, p.tags.join(" "), p.methods, p.disease, p.domain]
          .join(" ").toLowerCase();
        if (!hay.includes(needle)) return false;
      }
      return true;
    });
  }

  // --- Small view helpers -------------------------------------------------
  const esc = (s) => String(s).replace(/[&<>"']/g, (c) =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

  // Stable UU-palette colour per theme (consistent across tiles + badges).
  function themeColor(name) {
    const palette = (window.DashCharts && DashCharts.PALETTE) || ["#000000"];
    const i = state.themes.findIndex(([t]) => t === name);
    return palette[(i < 0 ? 0 : i) % palette.length];
  }

  function kpiHtml(items) {
    return `<section class="kpis">${items
      .map((k) => `<div class="kpi"><div class="kpi__value">${k.value}</div><div class="kpi__label">${esc(k.label)}</div></div>`)
      .join("")}</section>`;
  }

  function cardHtml(p) {
    const tags = p.tags.map((t) => `<span class="tag">${esc(t)}</span>`).join("");
    const link = p.website
      ? `<a class="card__link" href="${esc(p.website)}" target="_blank" rel="noopener">Visit project →</a>` : "";
    const hasCollab = p.collab === "Collaboration";
    const collabBadge = hasCollab
      ? `<span class="badge badge--collab"><span class="badge__dot" style="background:var(--green)"></span>Collaboration</span>`
      : `<span class="badge badge--solo"><span class="badge__dot" style="background:#b9b9b9"></span>No Collaboration</span>`;
    const themeBadges = state.activeTab === OVERVIEW
      ? splitMulti(p.theme).map((t) =>
          `<span class="badge"><span class="badge__dot" style="background:${themeColor(t)}"></span>${esc(t)}</span>`).join("")
      : "";
    const badges = [
      themeBadges,
      p.recordType ? `<span class="badge badge--type">${esc(p.recordType)}</span>` : "",
      p.stage ? `<span class="badge badge--stage">${esc(p.stage)}</span>` : "",
      collabBadge,
    ].join("");
    const detail = [
      p.institute ? `<div class="card__detail-row"><span class="card__detail-k">Main institute:</span> ${esc(p.institute)}</div>` : "",
      p.department ? `<div class="card__detail-row"><span class="card__detail-k">Involved department:</span> ${esc(p.department)}</div>` : "",
      hasCollab ? `<div class="card__detail-row"><span class="card__detail-k">Collaborating with:</span> ${esc(p.collaborators)}</div>` : "",
      p.population ? `<div class="card__detail-row"><span class="card__detail-k">Population / model:</span> ${esc(p.population)}</div>` : "",
      p.contact ? `<div class="card__detail-row"><span class="card__detail-k">Main contact:</span> ${esc(p.contact)}</div>` : "",
    ].join("");
    return `<article class="card">
      <div class="card__badges">${badges}</div>
      <h3 class="card__title">${esc(p.title)}</h3>
      <p class="card__desc">${esc(p.description)}</p>
      <div class="card__detail">${detail}</div>
      <div class="card__tags">${tags}</div>
      ${link}
    </article>`;
  }

  function destroyCharts() {
    state.charts.forEach((c) => c && c.destroy());
    state.charts = [];
  }

  // --- Tabs ---------------------------------------------------------------
  function renderTabs() {
    const tabs = $("tabs");
    const btns = [{ id: OVERVIEW, label: "Overview" }]
      .concat(state.themes.map(([name]) => ({ id: name, label: name })));
    tabs.innerHTML = btns
      .map((b) => {
        const active = b.id === state.activeTab;
        return `<button class="tab${active ? " is-active" : ""}" data-tab="${esc(b.id)}"${active ? ' aria-current="page"' : ""}>${esc(b.label)}</button>`;
      })
      .join("");
    tabs.querySelectorAll(".tab").forEach((btn) =>
      btn.addEventListener("click", () => selectTab(btn.dataset.tab)));
    // The animated brain is the primary navigation on the home page.
    tabs.hidden = state.activeTab === HOME;
  }

  function selectTab(tab) {
    state.activeTab = tab;
    state.filters = { innovation: "", recordType: "", stage: "", domain: "", institute: "", collab: "", q: "" };
    const slug = tab === HOME ? "" : tab === OVERVIEW ? "overview" : encodeURIComponent(tab);
    const current = decodeURIComponent((location.hash || "").slice(1));
    const expected = tab === HOME ? "" : tab === OVERVIEW ? "overview" : tab;
    if (current !== expected) {
      history.pushState(null, "", slug ? `#${slug}` : location.pathname);
    }
    renderTabs();
    render();
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function tabFromHash() {
    const raw = decodeURIComponent((location.hash || "").slice(1));
    if (!raw) return HOME;
    if (raw === "overview") return OVERVIEW;
    if (raw && state.themes.some(([name]) => name === raw)) return raw;
    return HOME;
  }

  // --- Render: cards only (filter application) ----------------------------
  function applyFilters() {
    const list = applyFacets(themeProjects());
    const grid = $("projectGrid");
    if (grid) grid.innerHTML = list.map(cardHtml).join("");
    const empty = $("emptyState");
    if (empty) empty.hidden = list.length > 0;
    const count = $("resultCount");
    if (count) count.textContent = `${list.length} of ${themeProjects().length} projects`;
    renderActivePills();
  }

  // Removable chips for every active filter — the clear way to deselect.
  function renderActivePills() {
    const bar = $("activeFilters");
    if (!bar) return;
    const pills = FACETS.filter((f) => state.filters[f.key])
      .map((f) => ({ key: f.key, label: f.label, value: state.filters[f.key] }));
    if (state.filters.q) pills.push({ key: "q", label: "Search", value: state.filters.q });
    if (!pills.length) { bar.hidden = true; bar.innerHTML = ""; return; }
    bar.hidden = false;
    bar.innerHTML =
      `<span class="active-filters__label">Active filters</span>` +
      pills.map((p) => `<button class="pill" data-key="${esc(p.key)}" aria-label="Remove filter ${esc(p.label)}: ${esc(p.value)}">
        <span class="pill__k">${esc(p.label)}:</span> ${esc(p.value)} <span class="pill__x" aria-hidden="true">×</span>
      </button>`).join("") +
      `<button class="pill pill--clear" data-key="__all__">Clear all</button>`;
    bar.querySelectorAll(".pill").forEach((btn) =>
      btn.addEventListener("click", () => removeFilter(btn.dataset.key)));
  }

  function removeFilter(key) {
    if (key === "__all__") state.filters = { innovation: "", recordType: "", stage: "", domain: "", institute: "", collab: "", q: "" };
    else state.filters[key] = "";
    const si = $("searchInput");
    if (si) si.value = state.filters.q;
    syncFacetControls();
    if (state.rebuildCharts) state.rebuildCharts();
    applyFilters();
  }

  // --- Render: full view --------------------------------------------------
  function render() {
    destroyCharts();
    state.rebuildCharts = null;
    if (state.activeTab === HOME) renderHome();
    else if (state.activeTab === OVERVIEW) renderOverview();
    else renderThemeView();
  }

  function renderHome() {
    // state.themes is [[themeName, projectCount], ...] — feed the live counts to the hero.
    const themeCounts = Object.fromEntries(state.themes);
    const brainHero = window.BrainHero
      ? window.BrainHero.render({
          assetPath: CONFIG.brainVideo,
          posterPath: CONFIG.brainPoster,
          counts: themeCounts,
          total: state.projects.length,
        })
      : "";
    view().innerHTML = `
      ${brainHero}
      <div class="brain-note">
        <button class="overview-link" id="overviewLink" type="button">View the complete portfolio overview <span aria-hidden="true">→</span></button>
      </div>
      <section class="project-team" aria-labelledby="projectTeamTitle">
        <div class="project-team__intro">
          <span class="project-team__eyebrow">Behind the dashboard</span>
          <h2 id="projectTeamTitle">${esc(PROJECT_TEAM.name)}</h2>
          <p>This research dashboard was developed by:</p>
        </div>
        <ul class="project-team__members" aria-label="Project members">
          ${PROJECT_TEAM.members.map((name) => `<li>${esc(name)}</li>`).join("")}
        </ul>
      </section>`;

    $("overviewLink").addEventListener("click", () => selectTab(OVERVIEW));
    if (window.BrainHero) {
      window.BrainHero.mount({
        root: $("brainHero"),
        onThemeSelect(theme) { selectTab(theme.label); },
      });
    }
  }

  function renderOverview() {
    const p = state.projects;
    const kpis = [
      { value: p.length, label: "Research projects" },
      { value: state.themes.length, label: "Dashboard themes" },
      { value: new Set(p.flatMap((x) => splitMulti(x.domain))).size, label: "Research domains" },
      { value: new Set(p.map((x) => x.institute).filter(Boolean)).size, label: "Institutes" },
      { value: p.filter((x) => x.collab === "Collaboration").length, label: "Collaborations" },
    ];
    view().innerHTML = `
      <section class="intro">
        <h2>Explore by theme</h2>
        <p>Select a theme to open its dashboard, or scan the big-picture charts below. Each theme page has its own filters and charts.</p>
      </section>
      <div class="theme-grid">${state.themes.map(themeTileHtml).join("")}</div>
      ${kpiHtml(kpis)}
      <section class="charts">
        <figure class="chart-card chart-card--wide"><figcaption>Projects per theme <span class="chart-hint">Click to open</span></figcaption><div class="chart-wrap chart-wrap--tall"><canvas id="cTheme"></canvas></div></figure>
        <figure class="chart-card"><figcaption>Innovation theme</figcaption><div class="chart-wrap chart-wrap--tall"><canvas id="cInnovation"></canvas></div></figure>
        <figure class="chart-card"><figcaption>Record type</figcaption><div class="chart-wrap chart-wrap--tall"><canvas id="cRecordType"></canvas></div></figure>
      </section>`;

    state.charts.push(DashCharts.makeBar($("cTheme"), state.themes, (name) => selectTab(name), { horizontal: true }));
    state.charts.push(DashCharts.makeDoughnut($("cInnovation"), countBy(p, "innovation", true), null));
    state.charts.push(DashCharts.makeBar($("cRecordType"), countBy(p, "recordType", false), null, { horizontal: true }));

    view().querySelectorAll(".theme-tile").forEach((t) =>
      t.addEventListener("click", () => selectTab(t.dataset.theme)));
  }

  function themeTileHtml([name, count]) {
    const color = themeColor(name);
    return `<button class="theme-tile" data-theme="${esc(name)}" style="border-top-color:${color}">
      <span class="theme-tile__count">${count}</span>
      <span class="theme-tile__name">${esc(name)}</span>
      <span class="theme-tile__go">Open dashboard →</span>
    </button>`;
  }

  function renderThemeView() {
    const subset = themeProjects();
    const kpis = [
      { value: subset.length, label: "Projects in theme" },
      { value: new Set(subset.flatMap((x) => splitMulti(x.domain))).size, label: "Research domains" },
      { value: new Set(subset.map((x) => x.stage).filter(Boolean)).size, label: "Research stages" },
      { value: new Set(subset.map((x) => x.institute).filter(Boolean)).size, label: "Institutes" },
      { value: subset.filter((x) => x.collab === "Collaboration").length, label: "Collaborations" },
    ];
    const facetControls = FACETS.map((f) => {
      const opts = countBy(subset, f.key, f.multi);
      return `<select class="facet" data-key="${f.key}" aria-label="${esc(f.label)}">
        <option value="">${esc(f.label)} — all</option>
        ${opts.map(([v, n]) => `<option value="${esc(v)}">${esc(v)} (${n})</option>`).join("")}
      </select>`;
    }).join("");

    view().innerHTML = `
      <section class="theme-head">
        <span class="theme-head__eyebrow">Theme</span>
        <h2>${esc(state.activeTab)}</h2>
        <p>${subset.length} research projects in this theme across the Utrecht Science Park.</p>
      </section>
      ${kpiHtml(kpis)}
      <section class="charts">
        <figure class="chart-card chart-card--wide"><figcaption>Research domain <span class="chart-hint">Click to filter</span></figcaption><div class="chart-wrap chart-wrap--tall"><canvas id="cDomain"></canvas></div></figure>
        <figure class="chart-card"><figcaption>Innovation theme <span class="chart-hint">Click to filter</span></figcaption><div class="chart-wrap chart-wrap--tall"><canvas id="cInnovation"></canvas></div></figure>
        <figure class="chart-card"><figcaption>Research stage <span class="chart-hint">Click to filter</span></figcaption><div class="chart-wrap chart-wrap--tall"><canvas id="cStage"></canvas></div></figure>
      </section>
      <section class="controls">
        <div class="controls__search">
          <svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/></svg>
          <input type="search" id="searchInput" placeholder="Search within this theme…" autocomplete="off" />
        </div>
        <div class="controls__facets">${facetControls}</div>
      </section>
      <section class="results">
        <div class="active-filters" id="activeFilters" hidden></div>
        <div class="results__head"><h2>Projects</h2><span class="results__count" id="resultCount"></span></div>
        <div class="grid" id="projectGrid"></div>
        <p class="empty" id="emptyState" hidden>No projects match the current filters.</p>
      </section>`;

    // Cross-filtering: each chart's bars/slices reflect the projects left after
    // applying every *other* facet (and the search), while keeping its own full
    // set of options so you can still switch selection. The active facet is
    // highlighted. Charts are rebuilt whenever any filter changes, so picking a
    // value on one chart instantly reshapes the others — and the cards.
    function buildCharts() {
      destroyCharts();
      state.charts.push(DashCharts.makeBar($("cDomain"), countBy(applyFacets(subset, "domain"), "domain", true), setFacet("domain"),
        { horizontal: true, topN: 10, active: state.filters.domain }));
      state.charts.push(DashCharts.makeDoughnut($("cInnovation"), countBy(applyFacets(subset, "innovation"), "innovation", true), setFacet("innovation"),
        { active: state.filters.innovation }));
      state.charts.push(DashCharts.makeBar($("cStage"), countBy(applyFacets(subset, "stage"), "stage", false), setFacet("stage"),
        { horizontal: true, topN: 8, active: state.filters.stage }));
    }
    const setFacet = (key) => (value) => {
      state.filters[key] = state.filters[key] === value ? "" : value;
      syncFacetControls();
      buildCharts();
      applyFilters();
    };
    buildCharts();
    state.rebuildCharts = buildCharts; // lets removeFilter() refresh chart highlights

    // Wire controls once; a facet change rebuilds charts + cards, search only cards.
    view().querySelectorAll(".facet").forEach((sel) =>
      sel.addEventListener("change", () => {
        state.filters[sel.dataset.key] = sel.value;
        buildCharts();
        applyFilters();
      }));
    $("searchInput").addEventListener("input", (e) => {
      state.filters.q = e.target.value;
      applyFilters();
    });

    applyFilters();
  }

  function syncFacetControls() {
    view().querySelectorAll(".facet").forEach((sel) => {
      sel.value = state.filters[sel.dataset.key] || "";
    });
  }

  // --- Boot ---------------------------------------------------------------
  async function init() {
    const status = $("loadStatus");
    try {
      state.projects = await loadData();
      // Theme is multi-valued (";"-separated) — a project can belong to several themes.
      state.themes = countBy(state.projects, "theme", true);
      status.textContent = `${state.projects.length} projects loaded`;
      status.dataset.state = "ready";
      $("dataMeta").textContent = `${state.projects.length} projects · ${state.themes.length} themes`;
      state.activeTab = tabFromHash();
      renderTabs();
      render();
      const homeLink = $("homeLink");
      if (homeLink) homeLink.addEventListener("click", (event) => {
        event.preventDefault();
        selectTab(HOME);
      });
      window.addEventListener("hashchange", () => {
        const tab = tabFromHash();
        if (tab !== state.activeTab) selectTab(tab);
      });
      window.addEventListener("popstate", () => {
        const tab = tabFromHash();
        if (tab !== state.activeTab) {
          state.activeTab = tab;
          state.filters = { innovation: "", recordType: "", stage: "", domain: "", institute: "", collab: "", q: "" };
          renderTabs();
          render();
          window.scrollTo({ top: 0, behavior: "smooth" });
        }
      });
    } catch (err) {
      console.error(err);
      status.textContent = "Failed to load data";
      status.dataset.state = "error";
      const isFile = location.protocol === "file:";
      view().innerHTML = `<section class="loaderr">
        <h2>Could not load the data file</h2>
        ${isFile
          ? `<p>You opened this page directly from disk (<code>file://</code>). Browsers block reading the Excel file that way, so no themes can be shown.</p>
             <p>Run a tiny local server from the project folder and open it over <code>http://</code>:</p>
             <pre>python3 -m http.server 8000</pre>
             <p>Then visit <a href="http://localhost:8000">http://localhost:8000</a>.</p>`
          : `<p>${esc(String(err.message || err))}</p>
             <p>Check that <code>${esc(CONFIG.dataFile)}</code> exists and the sheet is named <code>${esc(CONFIG.sheet)}</code>.</p>`}
      </section>`;
    }
  }

  document.addEventListener("DOMContentLoaded", init);
})();
