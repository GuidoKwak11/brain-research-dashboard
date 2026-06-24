/* Chart factory — themed Chart.js helpers for the dashboard.
   Readable (wrapped labels, direct value labels, dark ticks) and interactive
   (clickable to filter/drill, active selection highlighted). */
(function (global) {
  // Utrecht University secondary palette (yellow is reserved for identity,
  // it is too light to read on a white chart background).
  const PALETTE = [
    "#24A793", // green
    "#5287C6", // blue
    "#5B2182", // purple
    "#F3965E", // orange
    "#C00A35", // red
    "#AA1555", // bordeaux
    "#001240", // dark blue
    "#6E3B23", // brown
  ];
  const GRID = "rgba(0,0,0,0.07)";
  const TICK = "#1a1a1a";

  Chart.defaults.font.family = "'Open Sans', system-ui, sans-serif";
  Chart.defaults.color = TICK;

  function pickColors(n) {
    return Array.from({ length: n }, (_, i) => PALETTE[i % PALETTE.length]);
  }

  // Dim every colour except the active label so the selection stands out.
  function shade(labels, active) {
    const base = pickColors(labels.length);
    if (!active || !labels.includes(active)) return base.map((c) => c + "e6"); // ~90% opaque
    return labels.map((l, i) => (l === active ? base[i] : base[i] + "33"));
  }

  // Wrap a long category label onto at most `maxLines` lines, ellipsising the
  // overflow. The full label always remains available in the tooltip, so we keep
  // axis ticks compact and aligned instead of letting them sprawl and collide.
  // Breaks on spaces and after slashes, and hard-breaks any over-long token.
  function wrapLabel(str, max, maxLines = 2) {
    const tokens = String(str).replace(/\//g, "/ ").split(/\s+/).filter(Boolean);
    const lines = [];
    let cur = "";
    for (let t of tokens) {
      while (t.length > max) {
        if (cur) { lines.push(cur); cur = ""; }
        lines.push(t.slice(0, max));
        t = t.slice(max);
      }
      if ((cur + " " + t).trim().length > max && cur) { lines.push(cur); cur = t; }
      else cur = (cur + " " + t).trim();
    }
    if (cur) lines.push(cur);
    if (lines.length <= maxLines) return lines;
    const kept = lines.slice(0, maxLines);
    const last = kept[maxLines - 1];
    kept[maxLines - 1] = (last.length > max - 1 ? last.slice(0, max - 1).trimEnd() : last) + "…";
    return kept;
  }

  // Collapse a long tail of entries into a single non-clickable "Other" bar.
  function withTopN(entries, topN) {
    if (!topN || entries.length <= topN) return entries;
    const head = entries.slice(0, topN);
    const tail = entries.slice(topN).reduce((s, e) => s + e[1], 0);
    return head.concat([[`Other (${entries.length - topN})`, tail]]);
  }

  function clickHandler(chart, labels, onPick) {
    return (evt) => {
      const hit = chart.getElementsAtEventForMode(evt, "nearest", { intersect: true }, true);
      if (!hit.length) return;
      const label = labels[hit[0].index];
      if (/^Other \(/.test(label)) return; // aggregate bucket isn't a real filter
      if (typeof onPick === "function") onPick(label);
    };
  }

  // Plugin: draw each value at the end of its bar (direct labeling).
  function valueLabelPlugin(horizontal) {
    return {
      id: "valueLabels",
      afterDatasetsDraw(chart) {
        const { ctx } = chart;
        const meta = chart.getDatasetMeta(0);
        ctx.save();
        ctx.font = "700 11px 'Open Sans', sans-serif";
        ctx.fillStyle = "#1a1a1a";
        meta.data.forEach((el, i) => {
          const v = chart.data.datasets[0].data[i];
          if (horizontal) {
            ctx.textAlign = "left";
            ctx.textBaseline = "middle";
            ctx.fillText(v, el.x + 6, el.y);
          } else {
            ctx.textAlign = "center";
            ctx.textBaseline = "bottom";
            ctx.fillText(v, el.x, el.y - 4);
          }
        });
        ctx.restore();
      },
    };
  }

  function tooltipCfg(unit) {
    return {
      backgroundColor: "rgba(0,0,0,0.88)",
      padding: 10,
      titleFont: { size: 12, weight: "600" },
      bodyFont: { size: 12 },
      callbacks: {
        label(ctx) {
          const v = ctx.parsed.x ?? ctx.parsed.y ?? ctx.parsed;
          const total = ctx.dataset.data.reduce((a, b) => a + b, 0);
          const pct = total ? Math.round((v / total) * 100) : 0;
          return ` ${v} ${unit}  ·  ${pct}%`;
        },
      },
    };
  }

  function makeBar(canvas, rawEntries, onPick, opts = {}) {
    const { horizontal = false, active = null, topN = 0, unit = "projects" } = opts;
    const entries = withTopN(rawEntries, topN);
    const labels = entries.map((e) => e[0]);
    const data = entries.map((e) => e[1]);
    const catAxis = horizontal ? "y" : "x";
    const valAxis = horizontal ? "x" : "y";
    const clickable = typeof onPick === "function";

    const scales = {
      [valAxis]: {
        beginAtZero: true,
        grid: { color: GRID, drawBorder: false },
        ticks: { color: TICK, precision: 0, font: { size: 11 } },
      },
      [catAxis]: {
        grid: { display: false },
        ticks: {
          color: TICK,
          autoSkip: false,
          font: { size: 12 },
          callback(value) { return wrapLabel(this.getLabelForValue(value), horizontal ? 18 : 12); },
        },
      },
    };

    const chart = new Chart(canvas, {
      type: "bar",
      data: {
        labels,
        datasets: [{
          data,
          backgroundColor: shade(labels, active),
          hoverBackgroundColor: pickColors(labels.length),
          borderRadius: 6,
          borderWidth: 0,
          maxBarThickness: 30,
        }],
      },
      options: {
        indexAxis: catAxis,
        responsive: true,
        maintainAspectRatio: false,
        layout: { padding: { right: horizontal ? 34 : 8, top: horizontal ? 0 : 18 } },
        animation: { duration: 450 },
        plugins: { legend: { display: false }, tooltip: tooltipCfg(unit) },
        scales,
      },
      plugins: [valueLabelPlugin(horizontal)],
    });

    if (clickable) {
      canvas.onclick = clickHandler(chart, labels, onPick);
      canvas.style.cursor = "pointer";
    } else {
      canvas.onclick = null;
      canvas.style.cursor = "default";
    }
    return chart;
  }

  function makeDoughnut(canvas, entries, onPick, opts = {}) {
    const { active = null } = opts;
    const labels = entries.map((e) => e[0]);
    const data = entries.map((e) => e[1]);
    const total = data.reduce((a, b) => a + b, 0);
    const clickable = typeof onPick === "function";

    const centerText = {
      id: "centerText",
      afterDraw(chart) {
        const { ctx, chartArea } = chart;
        const cx = (chartArea.left + chartArea.right) / 2;
        const cy = (chartArea.top + chartArea.bottom) / 2;
        ctx.save();
        ctx.textAlign = "center";
        const sel = active && labels.includes(active);
        ctx.fillStyle = "#000";
        ctx.font = "900 26px 'Merriweather', serif";
        ctx.fillText(sel ? data[labels.indexOf(active)] : total, cx, cy - 2);
        ctx.fillStyle = "#595959";
        ctx.font = "600 10px 'Open Sans', sans-serif";
        ctx.fillText(sel ? "selected" : "projects", cx, cy + 16);
        ctx.restore();
      },
    };

    const chart = new Chart(canvas, {
      type: "doughnut",
      data: {
        labels,
        datasets: [{
          data,
          backgroundColor: shade(labels, active),
          hoverBackgroundColor: pickColors(labels.length),
          borderColor: "#ffffff",
          borderWidth: 2,
        }],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        cutout: "62%",
        animation: { duration: 450 },
        plugins: {
          legend: {
            position: "bottom",
            labels: { boxWidth: 10, padding: 9, color: "#1a1a1a", font: { size: 11 } },
            onClick: clickable ? (e, item) => onPick(labels[item.index]) : undefined,
          },
          tooltip: tooltipCfg("projects"),
        },
      },
      plugins: [centerText],
    });

    if (clickable) {
      canvas.onclick = clickHandler(chart, labels, onPick);
      canvas.style.cursor = "pointer";
    } else {
      canvas.onclick = null;
      canvas.style.cursor = "default";
    }
    return chart;
  }

  global.DashCharts = { makeBar, makeDoughnut, PALETTE };
})(window);
