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
    instituteFilter: "Main institution filter",
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
  // Model system + research stage use the normalised buckets so they stay clean.
  const FACETS = [
    { key: "innovation", label: "Innovation theme", multi: true },
    { key: "disease", label: "Disease / condition", multi: true },
    { key: "modelNorm", label: "Model system", multi: true },
    { key: "stageNorm", label: "Research stage", multi: false },
    { key: "recordType", label: "Record type", multi: false },
    { key: "instituteFilter", label: "Institute", multi: true },
    { key: "collab", label: "Collaboration", multi: false },
  ];

  // One filter slot per facet, plus a free-text search and a tag chip filter.
  function emptyFilters() {
    const f = { q: "", tag: "" };
    FACETS.forEach((x) => { f[x.key] = ""; });
    return f;
  }

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
    activeProject: null,   // project id when a detail page is open
    returnTab: OVERVIEW,   // tab to go back to from a detail page
    filters: emptyFilters(),
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

  // The raw "Research stage" and "Population / model system" fields are free-text
  // with a long messy tail. We roll them into a small set of meaningful buckets so
  // their charts show every project (no "Other"), keyed on what's actually insightful.
  function normStage(raw) {
    const prim = String(raw || "").split(" / ")[0].toLowerCase();
    if (!prim) return "";
    if (prim.includes("student")) return "Student project";
    if (prim.includes("fundamental")) return "Fundamental";
    if (prim.includes("translational")) return "Translational";
    if (prim.includes("clinical")) return "Clinical";
    if (prim.includes("applied") || prim.includes("impact")) return "Applied / impact";
    return "Infrastructure & methods";
  }
  function normModelToken(token) {
    const s = String(token || "").toLowerCase();
    const has = (...ks) => ks.some((k) => s.includes(k));
    if (has("comput", "silico", "algorithm", "simulation", "machine learning", "mri", "fmri", "eeg", "imaging data", "neural data", "dataset", "registry", "database", "/data", "data-")) return "Computational & data";
    if (has("organoid", "ipsc", "cell", "tissue", "in vitro", "culture", "biopsy", "sample", "line", "assembloid", "slice", "spheroid", "neuron", "axon", "protein", "molecular", "circuit", "-on-chip", "tumour model", "tumor model", "disease model", "embryo")) return "Cell, tissue & organoid";
    if (has("animal", "mouse", "mice", "murine", "rat", "rodent", "zebrafish", "in vivo", "xenopus", "primate", "drosophila", "elegans", "pig", "sheep", "dog", "cat", "comparative model")) return "Animal models";
    return "Humans & participants";
  }
  function modelSystems(population) {
    return [...new Set(splitMulti(population).map(normModelToken))].join("; ");
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
        // No external partners when the cell is a bare "no" OR the revised phrasing
        // "No external collaborators publicly specified…".
        const hasCollab = collabRaw !== ""
          && !/^(no|none|n\/a|-|geen)$/i.test(collabRaw)
          && !/^no external collaborators/i.test(collabRaw);
        const stage = cell(r, "stage");
        const population = cell(r, "population");
        return {
          title: cell(r, "title"),
          theme: cell(r, "theme"),
          institute: cell(r, "institute"),
          instituteFilter: cell(r, "instituteFilter"),
          department: cell(r, "department"),
          collaborators: collabRaw,
          collab: hasCollab ? "Collaboration" : "No Collaboration",
          website: cell(r, "website"),
          innovation: cell(r, "innovation"),
          domain: cell(r, "domain"),
          disease: cell(r, "disease"),
          methods: cell(r, "methods"),
          population,
          modelNorm: modelSystems(population),
          stage,
          stageNorm: normStage(stage),
          description: cell(r, "description"),
          tags: splitMulti(cell(r, "tags")),
          contact: cell(r, "contact"),
          recordType: cell(r, "recordType"),
        };
      })
      .filter((p) => p.title)
      .map((p, i) => ({ ...p, id: i }));
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
    const tag = state.filters.tag;
    return projects.filter((p) => {
      for (const f of FACETS) {
        if (f.key === except) continue;
        const sel = state.filters[f.key];
        if (sel && !valuesOf(p, f.key, f.multi).includes(sel)) return false;
      }
      if (tag && !p.tags.includes(tag)) return false;
      if (needle) {
        const hay = [p.title, p.description, p.tags.join(" "), p.methods, p.disease, p.domain, p.institute, p.department]
          .join(" ").toLowerCase();
        if (!hay.includes(needle)) return false;
      }
      return true;
    });
  }

  // --- Small view helpers -------------------------------------------------
  const esc = (s) => String(s).replace(/[&<>"']/g, (c) =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

  // Capitalise the first letter for display only, leaving acronyms / mixed-case
  // tokens (iPSC, mRNA, AI, MRI, ALS) untouched. Underlying values stay unchanged
  // so filtering and matching keep working.
  function capFirst(s) {
    s = String(s == null ? "" : s);
    if (!s) return s;
    const firstWord = s.split(/\s/, 1)[0];
    if (/[A-Z]/.test(firstWord)) return s;
    return s.charAt(0).toUpperCase() + s.slice(1);
  }

  // Title-case institution / department / partner names: capitalise each word but
  // keep minor words (of, and, the, within…) lower, and leave acronyms / mixed-case
  // tokens (UMC, iPSC, UU, MRI, ALS, Health~Holland) exactly as written.
  const TITLE_MINOR = new Set(["a", "an", "and", "as", "at", "but", "by", "de", "den", "der", "en", "for", "het", "in", "nor", "of", "on", "or", "per", "the", "to", "van", "via", "vs", "with", "within"]);
  function titleCase(s) {
    s = String(s == null ? "" : s);
    if (!s) return s;
    let wordIndex = 0;
    return s.split(/(\s+|\/)/).map((tok) => {
      if (tok === "" || /^(\s+|\/)$/.test(tok)) return tok;
      const isFirst = wordIndex === 0;
      wordIndex++;
      if (/[A-Z]/.test(tok.slice(1))) return tok;            // mixed-case acronym/proper token
      const lower = tok.toLowerCase();
      if (!isFirst && TITLE_MINOR.has(lower)) return lower;  // minor word stays lower
      return tok.charAt(0).toUpperCase() + tok.slice(1).toLowerCase();
    }).join("");
  }

  // Split a department/collaborator field into clean items for a bulleted list.
  // Involved department uses " / " between units; collaborators use ";" (a "/" there
  // sits inside a single partner name, so we keep it).
  function splitListItems(text, mode) {
    if (!text) return [];
    const sep = mode === "collab" ? /\s*;\s*/ : /\s*\/\s*|\s*;\s*/;
    return text.split(sep).map((s) => s.replace(/[.;,\s]+$/, "").trim()).filter(Boolean);
  }
  function bulletList(text, mode) {
    const items = splitListItems(text, mode);
    if (!items.length) return "";
    return `<ul class="bullets">${items.map((i) => `<li>${esc(titleCase(i))}</li>`).join("")}</ul>`;
  }

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

  function tagChipsHtml(tags, activeTag) {
    return tags.map((t) =>
      `<button type="button" class="tag${t === activeTag ? " is-active" : ""}" data-tag="${esc(t)}" aria-pressed="${t === activeTag}">${esc(capFirst(t))}</button>`).join("");
  }

  function cardHtml(p) {
    const tags = tagChipsHtml(p.tags, state.filters.tag);
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
    const broadInstitute = splitMulti(p.instituteFilter).join(", ");
    const detail = [
      broadInstitute ? `<div class="card__detail-row"><span class="card__detail-k">Main institute:</span> ${esc(broadInstitute)}</div>` : "",
      p.department ? `<div class="card__detail-row card__detail-row--list"><span class="card__detail-k">Involved department</span>${bulletList(p.department, "dept")}</div>` : "",
      hasCollab ? `<div class="card__detail-row card__detail-row--list"><span class="card__detail-k">Collaborating with</span>${bulletList(p.collaborators, "collab")}</div>` : "",
    ].join("");
    return `<article class="card card--link" data-open="${p.id}">
      <div class="card__badges">${badges}</div>
      <h3 class="card__title"><button type="button" class="card__open" data-open="${p.id}">${esc(p.title)}</button></h3>
      <p class="card__desc">${esc(p.description)}</p>
      <div class="card__detail">${detail}</div>
      ${tags ? `<div class="card__tags">${tags}</div>` : ""}
      <div class="card__more" data-open="${p.id}" aria-hidden="true">
        <span>View full project</span>
        <span class="card__more-arrow" aria-hidden="true">→</span>
      </div>
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
    state.activeProject = null;
    state.filters = emptyFilters();
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

  // Open a project's full-detail page; remember where to return to.
  function openProject(id) {
    if (state.activeProject == null) state.returnTab = state.activeTab;
    state.activeProject = id;
    history.pushState(null, "", `#project/${id}`);
    renderTabs();
    render();
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  // Toggle a clickable tag as a filter within the current theme.
  function applyTagFilter(tag) {
    state.filters.tag = state.filters.tag === tag ? "" : tag;
    if (state.rebuildCharts) state.rebuildCharts();
    applyFilters();
  }

  // Read the URL hash and render the matching view (theme, overview, project, home).
  function applyRoute() {
    const raw = decodeURIComponent((location.hash || "").slice(1));
    const m = raw.match(/^project\/(\d+)$/);
    if (m && state.projects.some((p) => p.id === Number(m[1]))) {
      state.activeProject = Number(m[1]);
    } else {
      state.activeProject = null;
      state.activeTab = !raw ? HOME
        : raw === "overview" ? OVERVIEW
        : state.themes.some(([name]) => name === raw) ? raw : HOME;
      state.filters = emptyFilters();
    }
    renderTabs();
    render();
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
    if (state.filters.tag) pills.push({ key: "tag", label: "Tag", value: state.filters.tag });
    if (state.filters.q) pills.push({ key: "q", label: "Search", value: state.filters.q });
    if (!pills.length) { bar.hidden = true; bar.innerHTML = ""; return; }
    bar.hidden = false;
    bar.innerHTML =
      `<span class="active-filters__label">Active filters</span>` +
      pills.map((p) => {
        const dv = p.key === "q" ? p.value : capFirst(p.value);
        return `<button class="pill" data-key="${esc(p.key)}" aria-label="Remove filter ${esc(p.label)}: ${esc(dv)}">
        <span class="pill__k">${esc(p.label)}:</span> ${esc(dv)} <span class="pill__x" aria-hidden="true">×</span>
      </button>`;
      }).join("") +
      `<button class="pill pill--clear" data-key="__all__">Clear all</button>`;
    bar.querySelectorAll(".pill").forEach((btn) =>
      btn.addEventListener("click", () => removeFilter(btn.dataset.key)));
  }

  function removeFilter(key) {
    if (key === "__all__") state.filters = emptyFilters();
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
    if (state.activeProject != null) renderProject();
    else if (state.activeTab === HOME) renderHome();
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
      { value: new Set(p.flatMap((x) => splitMulti(x.instituteFilter))).size, label: "Institutes" },
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
      { value: new Set(subset.flatMap((x) => splitMulti(x.instituteFilter))).size, label: "Institutes" },
      { value: subset.filter((x) => x.collab === "Collaboration").length, label: "Collaborations" },
    ];
    const facetControls = FACETS.map((f) => {
      const opts = countBy(subset, f.key, f.multi);
      return `<select class="facet" data-key="${f.key}" aria-label="${esc(f.label)}">
        <option value="">${esc(f.label)} — all</option>
        ${opts.map(([v, n]) => `<option value="${esc(v)}">${esc(capFirst(v))} (${n})</option>`).join("")}
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
        <figure class="chart-card chart-card--wide"><figcaption>Model system <span class="chart-hint">Click to filter</span></figcaption><div class="chart-wrap chart-wrap--tall"><canvas id="cModel"></canvas></div></figure>
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
      state.charts.push(DashCharts.makeBar($("cModel"), countBy(applyFacets(subset, "modelNorm"), "modelNorm", true), setFacet("modelNorm"),
        { horizontal: true, active: state.filters.modelNorm }));
      state.charts.push(DashCharts.makeDoughnut($("cInnovation"), countBy(applyFacets(subset, "innovation"), "innovation", true), setFacet("innovation"),
        { active: state.filters.innovation }));
      state.charts.push(DashCharts.makeBar($("cStage"), countBy(applyFacets(subset, "stageNorm"), "stageNorm", false), setFacet("stageNorm"),
        { horizontal: true, active: state.filters.stageNorm }));
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
        sel.classList.toggle("facet--on", !!sel.value);
        buildCharts();
        applyFilters();
      }));
    $("searchInput").addEventListener("input", (e) => {
      state.filters.q = e.target.value;
      applyFilters();
    });

    // Delegated: a tag toggles its filter; clicking anywhere else on a card opens
    // the project's full page. Tags are checked first since they sit inside the card.
    $("projectGrid").addEventListener("click", (e) => {
      const tagBtn = e.target.closest(".tag[data-tag]");
      if (tagBtn) { applyTagFilter(tagBtn.dataset.tag); return; }
      const opener = e.target.closest("[data-open]");
      if (opener) openProject(Number(opener.dataset.open));
    });

    applyFilters();
  }

  function syncFacetControls() {
    view().querySelectorAll(".facet").forEach((sel) => {
      const v = state.filters[sel.dataset.key] || "";
      sel.value = v;
      sel.classList.toggle("facet--on", !!v);
    });
  }

  // Open a theme and pre-apply a tag filter (used from a project's tags).
  function openThemeWithTag(theme, tag) {
    selectTab(theme);
    state.filters.tag = tag;
    syncFacetControls();
    if (state.rebuildCharts) state.rebuildCharts();
    applyFilters();
  }

  // --- Render: single project (every field in the data file) --------------
  function renderProject() {
    const p = state.projects.find((x) => x.id === state.activeProject);
    const backLabel = state.returnTab === OVERVIEW ? "Overview" : state.returnTab;
    if (!p) {
      view().innerHTML = `<section class="project-detail">
        <button class="project-detail__back" id="pBack">← Back</button>
        <p class="empty">Project not found.</p></section>`;
      $("pBack").addEventListener("click", () => selectTab(state.returnTab));
      return;
    }
    const themesArr = splitMulti(p.theme);
    const primaryTheme = themesArr[0] || OVERVIEW;
    const chips = (arr) => arr.map((v) => `<span class="pchip">${esc(capFirst(v))}</span>`).join("");
    const themeChips = themesArr.map((t) =>
      `<button type="button" class="pchip pchip--theme" data-theme="${esc(t)}"><span class="pchip__dot" style="background:${themeColor(t)}"></span>${esc(t)}</button>`).join("");
    const tagChips = p.tags.length
      ? p.tags.map((t) => `<button type="button" class="tag" data-tag="${esc(t)}">${esc(capFirst(t))}</button>`).join("")
      : "";
    const rows = [
      ["Institution (broad)", chips(splitMulti(p.instituteFilter))],
      ["Main institute", chips(splitMulti(p.institute))],
      ["Involved department / research group", bulletList(p.department, "dept")],
      ["Collaborating departments / partners", bulletList(p.collaborators, "collab")],
      ["Utrecht Brain innovation theme", chips(splitMulti(p.innovation))],
      ["Research domain", chips(splitMulti(p.domain))],
      ["Disease / condition / application", chips(splitMulti(p.disease))],
      ["Methods & technologies", chips(splitMulti(p.methods))],
      ["Population / model system", chips(splitMulti(p.population))],
      ["Research stage", p.stage ? esc(capFirst(p.stage)) : ""],
      ["Record type", p.recordType ? esc(capFirst(p.recordType)) : ""],
      ["Main contact", p.contact ? esc(capFirst(p.contact)) : ""],
      ["Website", p.website ? `<a href="${esc(p.website)}" target="_blank" rel="noopener">${esc(p.website)} ↗</a>` : ""],
    ].filter(([, v]) => v);

    view().innerHTML = `
      <section class="project-detail">
        <button class="project-detail__back" id="pBack">← Back to ${esc(backLabel)}</button>
        <header class="project-detail__head">
          <span class="project-detail__eyebrow">Research project</span>
          <h2>${esc(p.title)}</h2>
          <div class="project-detail__themes">${themeChips}</div>
          <div class="card__badges project-detail__badges">
            ${p.recordType ? `<span class="badge badge--type">${esc(p.recordType)}</span>` : ""}
            ${p.stage ? `<span class="badge badge--stage">${esc(p.stage)}</span>` : ""}
            ${p.collab === "Collaboration"
              ? `<span class="badge badge--collab"><span class="badge__dot" style="background:var(--green)"></span>Collaboration</span>`
              : `<span class="badge badge--solo"><span class="badge__dot" style="background:#b9b9b9"></span>No Collaboration</span>`}
          </div>
        </header>
        ${p.description ? `<p class="project-detail__desc">${esc(p.description)}</p>` : ""}
        <dl class="project-detail__grid">
          ${rows.map(([k, v]) => `<div class="pdl"><dt>${esc(k)}</dt><dd>${v}</dd></div>`).join("")}
        </dl>
        ${tagChips ? `<div class="project-detail__tags"><span class="project-detail__tags-label">Tags / subthemes</span><div class="card__tags">${tagChips}</div></div>` : ""}
      </section>`;

    $("pBack").addEventListener("click", () => selectTab(state.returnTab));
    view().querySelectorAll(".pchip--theme").forEach((b) =>
      b.addEventListener("click", () => selectTab(b.dataset.theme)));
    view().querySelectorAll(".project-detail__tags .tag").forEach((b) =>
      b.addEventListener("click", () => openThemeWithTag(primaryTheme, b.dataset.tag)));
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
      applyRoute();
      const homeLink = $("homeLink");
      if (homeLink) homeLink.addEventListener("click", (event) => {
        event.preventDefault();
        selectTab(HOME);
      });
      // Browser back/forward and manual hash edits re-sync the view from the URL.
      const onNav = () => { applyRoute(); window.scrollTo({ top: 0, behavior: "smooth" }); };
      window.addEventListener("hashchange", onNav);
      window.addEventListener("popstate", onNav);
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
