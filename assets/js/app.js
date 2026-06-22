/* USP Brain Research Dashboard
   Theme-first navigation: an Overview tab plus one tab per dashboard theme.
   Each theme tab carries its own KPIs, charts and faceted filters. */
(function () {
  "use strict";

  const CONFIG = {
    dataFile: "data/2final_USP_brain_research_dashboard_data.xlsx",
    sheet: "Dashboard data",
  };

  // Source column headers (must match the Excel header row).
  const COL = {
    title: "Research / project title",
    theme: "Dashboard theme (hoofdpagina's in het dashboard)",
    institute: "Main institute",
    department: "Involved internal department / research group",
    website: "Website",
    innovation: "Utrecht Brain innovation theme",
    domain: "Research domain",
    disease: "Disease / condition / application",
    methods: "Methods & technologies",
    stage: "Research stage",
    description: "Dashboard description",
    tags: "Tags / subthemes",
  };

  // Facets shown inside a theme view (theme itself is fixed by the active tab).
  const FACETS = [
    { key: "innovation", label: "Innovation theme", multi: true },
    { key: "stage", label: "Research stage", multi: false },
    { key: "domain", label: "Research domain", multi: true },
    { key: "institute", label: "Institute", multi: false },
  ];

  const OVERVIEW = "__overview__";

  const state = {
    projects: [],
    themes: [],            // [name, count] ordered by count desc
    activeTab: OVERVIEW,
    filters: { innovation: "", stage: "", domain: "", institute: "", q: "" },
    charts: [],
  };

  const $ = (id) => document.getElementById(id);
  const view = () => $("view");

  // --- Data ---------------------------------------------------------------
  function splitMulti(value) {
    if (!value) return [];
    return String(value).split(/[;\n]/).map((s) => s.trim()).filter(Boolean);
  }

  async function loadData() {
    const res = await fetch(CONFIG.dataFile);
    if (!res.ok) throw new Error(`Could not load ${CONFIG.dataFile} (${res.status})`);
    const wb = XLSX.read(await res.arrayBuffer(), { type: "array" });
    const ws = wb.Sheets[CONFIG.sheet] || wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
    const get = (r, k) => String(r[COL[k]] || "").trim();

    return rows
      .map((r) => ({
        title: get(r, "title"),
        theme: get(r, "theme"),
        institute: get(r, "institute"),
        department: get(r, "department"),
        website: get(r, "website"),
        innovation: get(r, "innovation"),
        domain: get(r, "domain"),
        disease: get(r, "disease"),
        methods: get(r, "methods"),
        stage: get(r, "stage"),
        description: get(r, "description"),
        tags: splitMulti(get(r, "tags")),
      }))
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
    return state.projects.filter((p) => p.theme === state.activeTab);
  }

  // Apply the facet + search filters on top of the theme subset.
  function applyFacets(projects) {
    const needle = state.filters.q.trim().toLowerCase();
    return projects.filter((p) => {
      for (const f of FACETS) {
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
    const tags = p.tags.slice(0, 5).map((t) => `<span class="tag">${esc(t)}</span>`).join("");
    const link = p.website
      ? `<a class="card__link" href="${esc(p.website)}" target="_blank" rel="noopener">Visit project →</a>` : "";
    const badges = [
      state.activeTab === OVERVIEW
        ? `<span class="badge"><span class="badge__dot" style="background:${themeColor(p.theme)}"></span>${esc(p.theme)}</span>` : "",
      p.stage ? `<span class="badge badge--stage">${esc(p.stage)}</span>` : "",
    ].join("");
    return `<article class="card">
      <div class="card__badges">${badges}</div>
      <h3 class="card__title">${esc(p.title)}</h3>
      <p class="card__desc">${esc(p.description)}</p>
      <div class="card__meta">${esc(p.institute)}${p.department ? " · " + esc(p.department) : ""}</div>
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
  }

  function selectTab(tab) {
    state.activeTab = tab;
    state.filters = { innovation: "", stage: "", domain: "", institute: "", q: "" };
    const slug = tab === OVERVIEW ? "" : encodeURIComponent(tab);
    if (decodeURIComponent((location.hash || "").slice(1)) !== tab) {
      history.replaceState(null, "", slug ? `#${slug}` : location.pathname);
    }
    renderTabs();
    render();
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function tabFromHash() {
    const raw = decodeURIComponent((location.hash || "").slice(1));
    if (raw && state.themes.some(([name]) => name === raw)) return raw;
    return OVERVIEW;
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
    const clear = $("clearFilters");
    if (clear) clear.hidden = !FACETS.some((f) => state.filters[f.key]) && !state.filters.q;
  }

  // --- Render: full view --------------------------------------------------
  function render() {
    destroyCharts();
    if (state.activeTab === OVERVIEW) renderOverview();
    else renderThemeView();
  }

  function renderOverview() {
    const p = state.projects;
    const kpis = [
      { value: p.length, label: "Research projects" },
      { value: state.themes.length, label: "Dashboard themes" },
      { value: new Set(p.flatMap((x) => splitMulti(x.domain))).size, label: "Research domains" },
      { value: new Set(p.map((x) => x.institute).filter(Boolean)).size, label: "Institutes" },
    ];
    view().innerHTML = `
      ${kpiHtml(kpis)}
      <section class="intro">
        <h2>Explore by theme</h2>
        <p>Pick a theme tab above to dive into its projects, or use the charts below for the big picture. Each theme page has its own filters and dashboards.</p>
      </section>
      <section class="charts">
        <figure class="chart-card chart-card--wide"><figcaption>Projects per theme</figcaption><div class="chart-wrap chart-wrap--tall"><canvas id="cTheme"></canvas></div></figure>
        <figure class="chart-card"><figcaption>Innovation theme</figcaption><div class="chart-wrap"><canvas id="cInnovation"></canvas></div></figure>
        <figure class="chart-card"><figcaption>Research stage</figcaption><div class="chart-wrap"><canvas id="cStage"></canvas></div></figure>
      </section>
      <section class="results">
        <div class="results__head"><h2>Themes</h2><span class="results__count">${state.themes.length} themes</span></div>
        <div class="theme-grid">${state.themes.map(themeTileHtml).join("")}</div>
      </section>`;

    state.charts.push(DashCharts.makeBar($("cTheme"), state.themes, (name) => selectTab(name), { horizontal: true }));
    state.charts.push(DashCharts.makeDoughnut($("cInnovation"), countBy(p, "innovation", true), null));
    state.charts.push(DashCharts.makeBar($("cStage"), countBy(p, "stage", false), null));

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
        <figure class="chart-card"><figcaption>Innovation theme</figcaption><div class="chart-wrap"><canvas id="cInnovation"></canvas></div></figure>
        <figure class="chart-card"><figcaption>Research stage</figcaption><div class="chart-wrap"><canvas id="cStage"></canvas></div></figure>
        <figure class="chart-card"><figcaption>Research domain</figcaption><div class="chart-wrap chart-wrap--tall"><canvas id="cDomain"></canvas></div></figure>
      </section>
      <section class="controls">
        <div class="controls__search">
          <svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/></svg>
          <input type="search" id="searchInput" placeholder="Search within this theme…" autocomplete="off" />
        </div>
        <div class="controls__facets">${facetControls}</div>
        <button class="btn-clear" id="clearFilters" type="button" hidden>Clear filters</button>
      </section>
      <section class="results">
        <div class="results__head"><h2>Projects</h2><span class="results__count" id="resultCount"></span></div>
        <div class="grid" id="projectGrid"></div>
        <p class="empty" id="emptyState" hidden>No projects match the current filters.</p>
      </section>`;

    // Charts reflect the full theme subset; clicking sets the matching facet.
    const setFacet = (key) => (value) => {
      state.filters[key] = state.filters[key] === value ? "" : value;
      syncFacetControls();
      applyFilters();
    };
    state.charts.push(DashCharts.makeDoughnut($("cInnovation"), countBy(subset, "innovation", true), setFacet("innovation")));
    state.charts.push(DashCharts.makeBar($("cStage"), countBy(subset, "stage", false), setFacet("stage")));
    state.charts.push(DashCharts.makeBar($("cDomain"), countBy(subset, "domain", true), setFacet("domain"), { horizontal: true }));

    // Wire controls once; filter application only re-renders cards.
    view().querySelectorAll(".facet").forEach((sel) =>
      sel.addEventListener("change", () => {
        state.filters[sel.dataset.key] = sel.value;
        applyFilters();
      }));
    $("searchInput").addEventListener("input", (e) => {
      state.filters.q = e.target.value;
      applyFilters();
    });
    $("clearFilters").addEventListener("click", () => {
      state.filters = { innovation: "", stage: "", domain: "", institute: "", q: "" };
      $("searchInput").value = "";
      syncFacetControls();
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
      state.themes = countBy(state.projects, "theme", false);
      status.textContent = `${state.projects.length} projects loaded`;
      status.dataset.state = "ready";
      $("dataMeta").textContent = `${state.projects.length} projects · ${state.themes.length} themes`;
      state.activeTab = tabFromHash();
      renderTabs();
      render();
      window.addEventListener("hashchange", () => {
        const tab = tabFromHash();
        if (tab !== state.activeTab) selectTab(tab);
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
