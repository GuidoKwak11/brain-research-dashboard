# USP Brain Research Dashboard

A static, dynamic dashboard that explores the brain-research portfolio of the
Utrecht Science Park (USP). It reads the source **Excel** file directly in the
browser and renders KPIs, overview charts and a filterable project grid.

No backend, no build step — it deploys to **GitHub Pages** as-is.

## How it works

- The Excel file in [`data/`](data/) is fetched and parsed **client-side** with
  [SheetJS](https://sheetjs.com/).
- Charts are drawn with [Chart.js](https://www.chartjs.org/) and are clickable
  (clicking a segment filters the project grid).
- Faceted filters (theme, innovation theme, research stage, domain, institute) and
  a free-text search narrow the results.
- All asset paths are **relative**, so it works at any path — including a project
  page subpath like `/Pages11-/`.

## Updating the data

1. Replace the Excel file in `data/` (keep the same structure / header row).
2. If you rename the file, update `CONFIG.dataFile` in
   [`assets/js/app.js`](assets/js/app.js).
3. The sheet used is **`Dashboard data`** — column headers must match the keys in
   the `COL` map at the top of `app.js`.

## Run locally

Because the page fetches the `.xlsx`, open it via a local server (not `file://`):

```bash
cd brain-research-dashboard
python3 -m http.server 8000
# open http://localhost:8000
```

## Deploy to GitHub Pages

Push this repository, then in **Settings → Pages** set the source to the
`main` branch (root). The `.nojekyll` file ensures the `assets/` folder is served
untouched.

## Structure

```
index.html              # markup + CDN scripts
assets/css/styles.css   # styling
assets/js/charts.js     # Chart.js helpers
assets/js/app.js        # data loading, filtering, rendering
data/*.xlsx             # source data
```
