/* Chart factory — themed Chart.js helpers for the dashboard.
   Charts are clickable: clicking a segment fires onPick(label). */
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
  const GRID = "rgba(0,0,0,0.08)";
  const TICK = "#444444";

  Chart.defaults.font.family = "'Open Sans', system-ui, sans-serif";
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
          borderColor: "#ffffff",
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
