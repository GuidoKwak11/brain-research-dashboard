/* USP Brain Research Dashboard — loads the Excel portfolio client-side,
   builds facet filters, KPI cards, overview charts and a project grid. */
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
    partners: "Collaborating departments / partners",
    website: "Website",
    innovation: "Utrecht Brain innovation theme",
    domain: "Research domain",
    disease: "Disease / condition / application",
    methods: "Methods & technologies",
    population: "Population / model system",
    stage: "Research stage",
    description: "Dashboard description",
    tags: "Tags / subthemes",
  };

  // Facet definitions: which column drives each dropdown.
  const FACETS = [
    { key: "theme", label: "Theme" },
    { key: "innovation", label: "Innovation theme" },
    { key: "stage", label: "Research stage" },
    { key: "domain", label: "Research domain" },
    { key: "institute", label: "Institute" },
  ];

  const state = {
    projects: [],
    filters: { theme: "", innovation: "", stage: "", domain: "", institute: "", q: "" },
    charts: {},
  };

  const $ = (id) => document.getElementById(id);

  // --- Data loading -------------------------------------------------------
  function splitMulti(value) {
    if (!value) return [];
    return String(value)
      .split(/[;\n]/)
      .map((s) => s.trim())
      .filter(Boolean);
  }

  async function loadData() {
    const res = await fetch(CONFIG.dataFile);
    if (!res.ok) throw new Error(`Could not load ${CONFIG.dataFile} (${res.status})`);
    const buf = await res.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const ws = wb.Sheets[CONFIG.sheet] || wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });

    return rows
      .map((r) => ({
        title: String(r[COL.title] || "").trim(),
        theme: String(r[COL.theme] || "").trim(),
        institute: String(r[COL.institute] || "").trim(),
        department: String(r[COL.department] || "").trim(),
        website: String(r[COL.website] || "").trim(),
        innovation: String(r[COL.innovation] || "").trim(),
        domain: String(r[COL.domain] || "").trim(),
        disease: String(r[COL.disease] || "").trim(),
        methods: String(r[COL.methods] || "").trim(),
        stage: String(r[COL.stage] || "").trim(),
        description: String(r[COL.description] || "").trim(),
        tags: splitMulti(r[COL.tags]),
        _domains: splitMulti(r[COL.domain]),
      }))
      .filter((p) => p.title);
  }

  // --- Filtering ----------------------------------------------------------
  function matchesFacet(project, key, value) {
    if (!value) return true;
    if (key === "domain") return project._domains.includes(value);
    return splitMulti(project[key]).includes(value) || project[key] === value;
  }

  function filtered() {
    const { q } = state.filters;
    const needle = q.trim().toLowerCase();
    return state.projects.filter((p) => {
      for (const { key } of FACETS) {
        if (!matchesFacet(p, key, state.filters[key])) return false;
      }
      if (needle) {
        const hay = [p.title, p.description, p.tags.join(" "), p.methods, p.disease]
          .join(" ")
          .toLowerCase();
        if (!hay.includes(needle)) return false;
      }
      return true;
    });
  }

  function countBy(projects, key, { multi = false } = {}) {
    const counts = new Map();
    for (const p of projects) {
      const values = multi ? splitMulti(p[key]) : [p[key]];
      for (const v of values) {
        if (!v) continue;
        counts.set(v, (counts.get(v) || 0) + 1);
      }
    }
    return [...counts.entries()].sort((a, b) => b[1] - a[1]);
  }

  // --- Rendering ----------------------------------------------------------
  function renderKpis() {
    const p = state.projects;
    const kpis = [
      { value: p.length, label: "Research projects" },
      { value: new Set(p.map((x) => x.theme).filter(Boolean)).size, label: "Dashboard themes" },
      { value: new Set(p.flatMap((x) => x._domains)).size, label: "Research domains" },
      { value: new Set(p.map((x) => x.institute).filter(Boolean)).size, label: "Institutes" },
    ];
    $("kpis").innerHTML = kpis
      .map((k) => `<div class="kpi"><div class="kpi__value">${k.value}</div><div class="kpi__label">${k.label}</div></div>`)
      .join("");
  }

  function renderFacets() {
    const wrap = $("facets");
    wrap.innerHTML = "";
    for (const f of FACETS) {
      const opts = countBy(state.projects, f.key, { multi: f.key === "domain" });
      const sel = document.createElement("select");
      sel.className = "facet";
      sel.dataset.key = f.key;
      sel.setAttribute("aria-label", f.label);
      sel.innerHTML =
        `<option value="">${f.label} — all</option>` +
        opts.map(([v, n]) => `<option value="${escapeAttr(v)}">${escapeHtml(v)} (${n})</option>`).join("");
      sel.value = state.filters[f.key];
      sel.addEventListener("change", () => {
        state.filters[f.key] = sel.value;
        update();
      });
      wrap.appendChild(sel);
    }
  }

  function badge(text, color) {
    if (!text) return "";
    return `<span class="badge" style="color:${color}">${escapeHtml(text)}</span>`;
  }

  function renderCards(projects) {
    const grid = $("projectGrid");
    $("emptyState").hidden = projects.length > 0;
    grid.innerHTML = projects
      .map((p) => {
        const tags = p.tags.slice(0, 5).map((t) => `<span class="tag">${escapeHtml(t)}</span>`).join("");
        const link = p.website
          ? `<a class="card__link" href="${escapeAttr(p.website)}" target="_blank" rel="noopener">Visit project →</a>`
          : "";
        return `<article class="card">
          <div class="card__badges">
            ${badge(p.theme, "var(--violet)")}
            ${badge(p.stage, "var(--cyan)")}
          </div>
          <h3 class="card__title">${escapeHtml(p.title)}</h3>
          <p class="card__desc">${escapeHtml(p.description)}</p>
          <div class="card__meta">${escapeHtml(p.institute)}${p.department ? " · " + escapeHtml(p.department) : ""}</div>
          <div class="card__tags">${tags}</div>
          ${link}
        </article>`;
      })
      .join("");
  }

  function renderCharts() {
    const onPick = (key) => (value) => {
      state.filters[key] = state.filters[key] === value ? "" : value;
      renderFacets();
      update();
    };
    Object.values(state.charts).forEach((c) => c && c.destroy());
    state.charts.theme = DashCharts.makeBar(
      $("chartTheme"), countBy(state.projects, "theme"), onPick("theme"), { horizontal: true });
    state.charts.innovation = DashCharts.makeDoughnut(
      $("chartInnovation"), countBy(state.projects, "innovation"), onPick("innovation"));
    state.charts.stage = DashCharts.makeBar(
      $("chartStage"), countBy(state.projects, "stage"), onPick("stage"));
  }

  function update() {
    const list = filtered();
    renderCards(list);
    const active = Object.entries(state.filters).some(([k, v]) => k !== "q" && v) || state.filters.q;
    $("clearFilters").hidden = !active;
    $("resultCount").textContent =
      `${list.length} of ${state.projects.length} projects`;
  }

  function clearFilters() {
    state.filters = { theme: "", innovation: "", stage: "", domain: "", institute: "", q: "" };
    $("searchInput").value = "";
    renderFacets();
    update();
  }

  // --- Utilities ----------------------------------------------------------
  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, (c) =>
      ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  }
  function escapeAttr(s) { return escapeHtml(s); }

  // --- Boot ---------------------------------------------------------------
  async function init() {
    const status = $("loadStatus");
    try {
      state.projects = await loadData();
      status.textContent = `${state.projects.length} projects loaded`;
      status.dataset.state = "ready";
      $("dataMeta").textContent = `${state.projects.length} projects`;
      renderKpis();
      renderFacets();
      renderCharts();
      update();

      $("searchInput").addEventListener("input", (e) => {
        state.filters.q = e.target.value;
        update();
      });
      $("clearFilters").addEventListener("click", clearFilters);
    } catch (err) {
      console.error(err);
      status.textContent = "Failed to load data — see console";
      status.dataset.state = "error";
    }
  }

  document.addEventListener("DOMContentLoaded", init);
})();
