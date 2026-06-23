/* Interactive homepage hero.
   A rotating brain sits in a central disc; the site draws evenly spaced rays out to
   eleven theme nodes around it. Positions are computed on a circle (no overlap), so
   editing only means reordering brainThemes or tweaking the geometry constants below. */
(function (global) {
  "use strict";

  const brainThemes = [
    { id: 1,  label: "Behaviour",                       href: "/themes/behaviour" },
    { id: 2,  label: "Brain and neuron structure",      href: "/themes/brain-and-neuron-structure" },
    { id: 3,  label: "Brain development",               href: "/themes/brain-development" },
    { id: 4,  label: "Cancer",                          href: "/themes/cancer" },
    { id: 5,  label: "Communication impairments",       href: "/themes/communication-impairments" },
    { id: 6,  label: "Drug delivery",                   href: "/themes/drug-delivery" },
    { id: 7,  label: "Epilepsy",                        href: "/themes/epilepsy" },
    { id: 8,  label: "Ethics, society & care systems",  href: "/themes/ethics-society-care-systems" },
    { id: 9,  label: "Neuroimaging & brain technology", href: "/themes/neuroimaging-brain-technology" },
    { id: 10, label: "Neuromuscular disorders",         href: "/themes/neuromuscular-disorders" },
    { id: 11, label: "Stroke",                          href: "/themes/stroke" },
  ];

  // Geometry, in stage percentages. A = stage width / height (keeps the ring circular).
  const GEO = { cx: 50, cy: 50, A: 1.3, ringRx: 33, discRx: 21, startDeg: -90 };

  const escapeHtml = (value) => String(value).replace(/[&<>"']/g, (character) =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[character]));

  function placed(theme, index, count) {
    const ang = ((GEO.startDeg + (index * 360) / count) * Math.PI) / 180;
    const cos = Math.cos(ang), sin = Math.sin(ang);
    const ringRy = GEO.ringRx * GEO.A;
    const discRy = GEO.discRx * GEO.A;
    const node = { x: GEO.cx + GEO.ringRx * cos, y: GEO.cy + ringRy * sin };
    const inner = { x: GEO.cx + GEO.discRx * cos, y: GEO.cy + discRy * sin };
    let side = "top";
    if (cos > 0.25) side = "right";
    else if (cos < -0.25) side = "left";
    else side = sin < 0 ? "top" : "bottom";
    return { ...theme, node, inner, side };
  }

  function rayHtml(t) {
    return `<line class="brain-hero__ray" data-theme-id="${t.id}" x1="${t.inner.x.toFixed(2)}" y1="${t.inner.y.toFixed(2)}" x2="${t.node.x.toFixed(2)}" y2="${t.node.y.toFixed(2)}" pathLength="1" />
      <circle class="brain-hero__spark" data-theme-id="${t.id}" r="0.7">
        <animateMotion dur="3.6s" begin="${(t.id * -0.3).toFixed(2)}s" repeatCount="indefinite" path="M ${t.inner.x.toFixed(2)} ${t.inner.y.toFixed(2)} L ${t.node.x.toFixed(2)} ${t.node.y.toFixed(2)}" />
      </circle>`;
  }

  function hotspotHtml(t) {
    return `<a class="brain-hero__hotspot side-${t.side}" data-theme-id="${t.id}" href="${escapeHtml(t.href)}" aria-label="${escapeHtml(t.label)}" style="--nx:${t.node.x.toFixed(2)}%;--ny:${t.node.y.toFixed(2)}%">
      <span class="brain-hero__node" aria-hidden="true"></span>
      <span class="brain-hero__label">${escapeHtml(t.label)}</span>
    </a>`;
  }

  function render(options = {}) {
    const mp4 = options.assetPath || "brain/brain-rotating.mp4";
    const posterPath = options.posterPath || "brain/brain-hero-poster.jpg";
    const themes = brainThemes.map((t, i) => placed(t, i, brainThemes.length));

    return `<section class="brain-hero" id="brainHero" aria-label="Utrecht Science Park brain research theme map">
      <header class="brain-hero__copy">
        <span class="brain-hero__eyebrow"><span aria-hidden="true"></span> UTRECHT SCIENCE PARK</span>
        <h2 id="brainHeroTitle">Where brain research connects.</h2>
        <p>Explore eleven connected research themes shaping our understanding of the brain, health and human behaviour.</p>
      </header>
      <div class="brain-hero__stage">
        <div class="brain-hero__scene">
          <div class="brain-hero__disc">
            <video class="brain-hero__video" autoplay muted loop playsinline preload="metadata" aria-hidden="true" poster="${escapeHtml(posterPath)}">
              <source src="${escapeHtml(mp4)}" type="video/mp4" />
            </video>
            <span class="brain-hero__disc-ring" aria-hidden="true"></span>
          </div>
          <svg class="brain-hero__rays" viewBox="0 0 100 100" preserveAspectRatio="none" aria-hidden="true">
            ${themes.map(rayHtml).join("")}
          </svg>
          <nav class="brain-hero__hotspots" aria-label="Brain research themes">
            ${themes.map(hotspotHtml).join("")}
          </nav>
        </div>
      </div>
      <div class="brain-hero__footer">
        <span><i aria-hidden="true"></i> Live research landscape</span>
        <span>11 connected themes</span>
      </div>
    </section>`;
  }

  function mount(options = {}) {
    const root = typeof options.root === "string" ? document.querySelector(options.root) : options.root;
    if (!root) return;

    const video = root.querySelector(".brain-hero__video");
    const links = [...root.querySelectorAll(".brain-hero__hotspot")];
    const rays = [...root.querySelectorAll(".brain-hero__ray, .brain-hero__spark")];
    const reducedMotion = window.matchMedia("(prefers-reduced-motion: reduce)").matches;

    const setActive = (id) => {
      root.dataset.activeTheme = id || "";
      links.forEach((link) => link.classList.toggle("is-active", link.dataset.themeId === String(id)));
      rays.forEach((ray) => ray.classList.toggle("is-active", ray.dataset.themeId === String(id)));
    };

    links.forEach((link) => {
      const id = link.dataset.themeId;
      link.addEventListener("pointerenter", () => setActive(id));
      link.addEventListener("pointerleave", () => {
        if (document.activeElement !== link) setActive("");
      });
      link.addEventListener("focus", () => setActive(id));
      link.addEventListener("blur", () => setActive(""));
      if (typeof options.onThemeSelect === "function") {
        link.addEventListener("click", (event) => {
          if (event.defaultPrevented || event.metaKey || event.ctrlKey || event.shiftKey || event.altKey || event.button !== 0) return;
          event.preventDefault();
          const theme = brainThemes.find((item) => String(item.id) === id);
          if (theme) options.onThemeSelect(theme, event);
        });
      }
    });

    if (reducedMotion && video) {
      const pause = () => video.pause();
      if (video.readyState >= 2) pause();
      else video.addEventListener("loadeddata", pause, { once: true });
    }
  }

  global.brainThemes = brainThemes;
  global.BrainHero = { brainThemes, render, mount };
})(window);
