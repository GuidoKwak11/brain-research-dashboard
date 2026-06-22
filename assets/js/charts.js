/* Chart factory — themed Chart.js helpers for the dashboard.
   Charts are clickable: clicking a segment fires onPick(label). */
(function (global) {
  const PALETTE = [
    "#7c5cff", "#22d3ee", "#f472b6", "#fbbf24", "#34d399",
    "#60a5fa", "#a78bfa", "#fb7185", "#2dd4bf", "#facc15",
    "#f87171", "#818cf8",
  ];
  const GRID = "rgba(138,144,173,0.15)";
  const TICK = "#8b91ad";

  Chart.defaults.font.family = "Inter, system-ui, sans-serif";
  Chart.defaults.color = TICK;

  function pickColors(n) {
    return Array.from({ length: n }, (_, i) => PALETTE[i % PALETTE.length]);
  }

  function clickHandler(chart, labels, onPick) {
    return (evt) => {
      const hit = chart.getElementsAtEventForMode(evt, "nearest", { intersect: true }, true);
      if (hit.length && typeof onPick === "function") onPick(labels[hit[0].index]);
    };
  }

  function makeBar(canvas, entries, onPick, { horizontal = false } = {}) {
    const labels = entries.map((e) => e[0]);
    const data = entries.map((e) => e[1]);
    const chart = new Chart(canvas, {
      type: "bar",
      data: {
        labels,
        datasets: [{
          data,
          backgroundColor: pickColors(labels.length).map((c) => c + "cc"),
          hoverBackgroundColor: pickColors(labels.length),
          borderRadius: 6,
          borderWidth: 0,
        }],
      },
      options: {
        indexAxis: horizontal ? "y" : "x",
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { display: false } },
        onClick: null,
        scales: {
          x: { grid: { color: GRID }, ticks: { color: TICK, autoSkip: false } },
          y: { grid: { color: GRID }, ticks: { color: TICK, autoSkip: false } },
        },
      },
    });
    canvas.onclick = clickHandler(chart, labels, onPick);
    canvas.style.cursor = "pointer";
    return chart;
  }

  function makeDoughnut(canvas, entries, onPick) {
    const labels = entries.map((e) => e[0]);
    const data = entries.map((e) => e[1]);
    const chart = new Chart(canvas, {
      type: "doughnut",
      data: {
        labels,
        datasets: [{
          data,
          backgroundColor: pickColors(labels.length),
          borderColor: "#141829",
          borderWidth: 2,
        }],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        cutout: "58%",
        plugins: {
          legend: { position: "bottom", labels: { boxWidth: 10, padding: 10, font: { size: 11 } } },
        },
      },
    });
    canvas.onclick = clickHandler(chart, labels, onPick);
    canvas.style.cursor = "pointer";
    return chart;
  }

  global.DashCharts = { makeBar, makeDoughnut, PALETTE };
})(window);
