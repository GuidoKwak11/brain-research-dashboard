/* Interactive homepage hero.
   Edit this single config to change theme labels, links, or positions.
   x/y position the hotspot; anchorX/anchorY position its neural connection. */
(function (global) {
  "use strict";

  const brainThemes = [
    { id: 1, label: "Behaviour", href: "/themes/behaviour", x: 13, y: 31, anchorX: 45, anchorY: 45 },
    { id: 2, label: "Brain and neuron structure", href: "/themes/brain-and-neuron-structure", x: 28, y: 13, anchorX: 48, anchorY: 38 },
    { id: 3, label: "Brain development", href: "/themes/brain-development", x: 59, y: 10, anchorX: 52, anchorY: 35 },
    { id: 4, label: "Cancer", href: "/themes/cancer", x: 73, y: 17, anchorX: 58, anchorY: 38 },
    { id: 5, label: "Communication impairments", href: "/themes/communication-impairments", x: 88, y: 29, anchorX: 62, anchorY: 45 },
    { id: 6, label: "Drug delivery", href: "/themes/drug-delivery", x: 78, y: 47, anchorX: 64, anchorY: 52 },
    { id: 7, label: "Epilepsy", href: "/themes/epilepsy", x: 72, y: 76, anchorX: 60, anchorY: 60 },
    { id: 8, label: "Ethics, society & care systems", href: "/themes/ethics-society-care-systems", x: 55, y: 84, anchorX: 54, anchorY: 64 },
    { id: 9, label: "Neuroimaging & brain technology", href: "/themes/neuroimaging-brain-technology", x: 34, y: 82, anchorX: 48, anchorY: 62 },
    { id: 10, label: "Neuromuscular disorders", href: "/themes/neuromuscular-disorders", x: 13, y: 67, anchorX: 43, anchorY: 55 },
    { id: 11, label: "Stroke", href: "/themes/stroke", x: 30.5, y: 53.5, anchorX: 42, anchorY: 50 },
  ];

  // These five coordinates track the circle endpoints already rendered in the
  // source video. The remaining six themes are completed by our SVG overlay.
  const videoNodeIds = new Set([2, 3, 6, 7, 11]);

  const escapeHtml = (value) => String(value).replace(/[&<>"']/g, (character) =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[character]));

  function placement(theme) {
    if (theme.x <= 38) return "is-left";
    if (theme.x >= 62) return "is-right";
    return theme.y < 50 ? "is-top" : "is-bottom";
  }

  function connectionHtml(theme) {
    const videoNodeClass = videoNodeIds.has(theme.id) ? " is-video-node" : "";
    return `<g class="brain-hero__connection${videoNodeClass}" data-theme-id="${theme.id}">
      <line class="brain-hero__line brain-hero__line--base" x1="${theme.x}" y1="${theme.y}" x2="${theme.anchorX}" y2="${theme.anchorY}" pathLength="1" />
      <line class="brain-hero__line brain-hero__line--pulse" x1="${theme.x}" y1="${theme.y}" x2="${theme.anchorX}" y2="${theme.anchorY}" pathLength="1" style="--pulse-delay:${theme.id * -0.27}s" />
      <circle class="brain-hero__anchor" cx="${theme.anchorX}" cy="${theme.anchorY}" r="0.42" />
    </g>`;
  }

  function hotspotHtml(theme) {
    const videoNodeClass = videoNodeIds.has(theme.id) ? " is-video-node" : "";
    return `<a class="brain-hero__hotspot ${placement(theme)}${videoNodeClass}" data-theme-id="${theme.id}" href="${escapeHtml(theme.href)}" aria-label="${escapeHtml(theme.label)}" style="--x:${theme.x}%;--y:${theme.y}%;--hotspot-delay:${theme.id * -0.19}s">
      <span class="brain-hero__node" aria-hidden="true"><span class="brain-hero__node-core"></span></span>
      <span class="brain-hero__label">${escapeHtml(theme.label)}</span>
    </a>`;
  }

  function render(options = {}) {
    const assetPath = options.assetPath || "brain/openart-02178216296258000000000000000000000ffffc0a86fc0d40593_1782163076164_aef8943f.mp4";
    const posterPath = options.posterPath || "brain/brain-hero-poster.jpg";

    return `<section class="brain-hero" id="brainHero" aria-labelledby="brainHeroTitle">
      <div class="brain-hero__ambient" aria-hidden="true"></div>
      <header class="brain-hero__copy">
        <span class="brain-hero__eyebrow"><span aria-hidden="true"></span> Utrecht Science Park</span>
        <h2 id="brainHeroTitle">Where brain research connects.</h2>
        <p>Explore eleven connected research themes shaping our understanding of the brain, health and human behaviour.</p>
      </header>
      <div class="brain-hero__stage">
        <div class="brain-hero__media" aria-hidden="true">
          <video class="brain-hero__video" autoplay muted loop playsinline preload="metadata" poster="${escapeHtml(posterPath)}">
            <source src="${escapeHtml(assetPath)}" type="video/mp4" />
          </video>
          <div class="brain-hero__media-vignette"></div>
        </div>
        <svg class="brain-hero__connections" viewBox="0 0 100 100" preserveAspectRatio="none" aria-hidden="true">
          ${brainThemes.map(connectionHtml).join("")}
        </svg>
        <nav class="brain-hero__hotspots" aria-label="Brain research themes">
          ${brainThemes.map(hotspotHtml).join("")}
        </nav>
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
    const connections = [...root.querySelectorAll(".brain-hero__connection")];
    const reducedMotion = window.matchMedia("(prefers-reduced-motion: reduce)").matches;
    const finePointer = window.matchMedia("(hover: hover) and (pointer: fine)").matches;
    let frame = 0;

    const setActive = (id) => {
      root.dataset.activeTheme = id || "";
      links.forEach((link) => link.classList.toggle("is-active", link.dataset.themeId === String(id)));
      connections.forEach((line) => line.classList.toggle("is-active", line.dataset.themeId === String(id)));
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
          event.preventDefault();
          const theme = brainThemes.find((item) => String(item.id) === id);
          if (theme) options.onThemeSelect(theme, event);
        });
      }
    });

    if (reducedMotion && video) {
      const pauseAtPosterFrame = () => {
        video.pause();
        if (Number.isFinite(video.duration)) video.currentTime = Math.min(1.5, video.duration);
      };
      if (video.readyState >= 1) pauseAtPosterFrame();
      else video.addEventListener("loadedmetadata", pauseAtPosterFrame, { once: true });
    }

    if (!reducedMotion && finePointer) {
      root.addEventListener("pointermove", (event) => {
        const bounds = root.getBoundingClientRect();
        const x = ((event.clientX - bounds.left) / bounds.width - 0.5) * 2;
        const y = ((event.clientY - bounds.top) / bounds.height - 0.5) * 2;
        cancelAnimationFrame(frame);
        frame = requestAnimationFrame(() => {
          root.style.setProperty("--brain-shift-x", `${x * 8}px`);
          root.style.setProperty("--brain-shift-y", `${y * 5}px`);
        });
      });
      root.addEventListener("pointerleave", () => {
        root.style.setProperty("--brain-shift-x", "0px");
        root.style.setProperty("--brain-shift-y", "0px");
      });
    }
  }

  global.brainThemes = brainThemes;
  global.BrainHero = { brainThemes, render, mount };
})(window);
