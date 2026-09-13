/*
 * Dash Architect task pane — the thin adapter wiring the engine/render
 * layers to Excel. All the decision logic (column roles, layout,
 * rendering) already exists in /engine and /render; this file only:
 *   1. reads the selected range (addin/excel-io.js, chunked, format
 *      sampled) and feeds it to engine.analyzeTable exactly like the CSV
 *      path does,
 *   2. shows the resulting dashboard live in this pane
 *      (render/dom.js#mount),
 *   3. on "Place on sheet": rasterizes it (render/share.js), inserts the
 *      PNG via worksheet.shapes.addImage, and saves a self-contained
 *      snapshot (engine/layout.js#buildStorageSnapshot) + the column
 *      mapping into workbook.settings, keyed to the picture's name
 *      (addin/dashboard-io.js),
 *   4. reopens a saved dashboard — via the reliable "Dashboards in this
 *      workbook" list, or a best-effort guess when the selection changes
 *      to something that isn't a normal cell range — as a frozen,
 *      no-recompute view (render/dom.js#mountFrozen).
 *
 * Runs two ways, chosen once at startup by `host` below:
 *   - inside Excel: every step above is real Office.js.
 *   - opened as a plain webpage (no Office host): a `host` that reads a
 *     bundled fixture instead of a live range and keeps placed dashboards
 *     in memory instead of workbook.settings, so the entire flow — every
 *     view, every button — can be exercised and verified without an Excel
 *     session. Nothing in engine/render/addin's Excel-facing modules
 *     changes between the two; only which `host` calls them differently.
 */
(() => {
  'use strict';

  const $ = (id) => document.getElementById(id);
  const wait = (ms) => new Promise((r) => setTimeout(r, ms));
  const plural = (n, one, many) => `${n.toLocaleString('en-US')} ${n === 1 ? one : many}`;

  const els = {};
  const state = {
    excel: false,
    picking: false,
    range: null, // {address, rows, cols}
    beforePick: null,
    busy: false,
    analysis: null,
    layoutSpec: null,
    sourceAddress: null,
    title: null,
    previewController: null,
    previewTheme: 'light',
    result: null,
  };
  let host = null;
  let activeScale = null; // {mountEl, stageEl, canvasSize} — reapplied on resize

  /* ---------- start ---------- */
  let started = false;
  function start(isExcel) {
    if (started) return;
    started = true;
    state.excel = !!isExcel;
    host = state.excel ? excelHost : previewHost;
    init();
  }

  /* ---------- hosts: real Excel vs. in-browser preview ---------- */

  const excelHost = {
    async readSelectionInfo() {
      try {
        return await Excel.run(async (ctx) => {
          const range = ctx.workbook.getSelectedRange();
          range.load('address,rowCount,columnCount');
          await ctx.sync();
          return { address: range.address, rows: range.rowCount, cols: range.columnCount };
        });
      } catch (err) {
        return null; // selection is a shape/chart/multi-area, not a plain range
      }
    },
    async readRangeForEngine() {
      return Excel.run(async (ctx) => {
        const range = ctx.workbook.getSelectedRange();
        return window.DashAddinExcelIo.readRangeForEngine(ctx, range);
      });
    },
    async placeDashboard(args) {
      return Excel.run((ctx) => window.DashAddinDashboardIo.generateAndPlace(ctx, args));
    },
    async listDashboards() {
      return Excel.run((ctx) => window.DashAddinDashboardIo.listDashboards(ctx));
    },
    // Best-effort only — Excel's JS API has no first-class "shape was
    // clicked" event. A failing getSelectedRange() after a selection
    // change is the closest available signal that the selection might now
    // be a picture instead of cells (it could also be a chart, an error,
    // or a transient state) — see the module doc comment and SPEC.md.
    async probeNonRangeSelection() {
      try {
        await Excel.run(async (ctx) => {
          const range = ctx.workbook.getSelectedRange();
          range.load('address');
          await ctx.sync();
        });
        return false;
      } catch (e) {
        return true;
      }
    },
  };

  // Lets every view/button be exercised without an Excel host — see the
  // file header comment. `previewDashboards` stands in for
  // workbook.settings; DashboardIo's naming/payload helpers are pure and
  // used as-is, so this is the same code path minus the Office.js calls.
  const previewDashboards = [];
  const PREVIEW_ADDRESS = "'Store sales'!A1:J901";

  const previewHost = {
    async readSelectionInfo() {
      return { address: PREVIEW_ADDRESS, rows: 901, cols: 10 };
    },
    async readRangeForEngine() {
      const t0 = performance.now();
      const text = await fetch('../fixtures/01_sales_timeseries.csv').then((r) => r.text());
      const { rows } = window.DashEngineCsv.parseCsv(text);
      const t1 = performance.now();
      const [headers, ...dataRows] = rows;
      return { headers, dataRows, totalRows: rows.length, cols: headers.length, timing: { valuesMs: Math.round(t1 - t0), formatMs: 0, totalMs: Math.round(t1 - t0) } };
    },
    async placeDashboard(args) {
      const shapeName = window.DashAddinDashboardIo.makeShapeName(Date.now());
      const payload = window.DashAddinDashboardIo.buildStoragePayload(args);
      previewDashboards.push({ shapeName, payload });
      return { shapeName };
    },
    async listDashboards() {
      return previewDashboards.map((d) => ({ shapeName: d.shapeName, payload: d.payload }));
    },
    async probeNonRangeSelection() {
      return false; // no selection events outside Excel
    },
  };

  // Deferred until here, after excelHost/previewHost are declared above:
  // Office.js's onReady callback can fire synchronously when the page is
  // loaded outside any Office client (there's no real handshake to wait
  // for), which previously ran start() — and its `host = ... : previewHost`
  // — before the `const previewHost` a few lines above it had been
  // evaluated yet.
  if (window.Office && typeof Office.onReady === 'function') {
    Office.onReady((info) => start(info && info.host === Office.HostType.Excel));
  } else {
    document.addEventListener('DOMContentLoaded', () => start(false));
    if (document.readyState !== 'loading') start(false);
  }

  /* ---------- wiring ---------- */

  function init() {
    Object.assign(els, {
      showList: $('show-list'), viewSetup: $('view-setup'), viewPreview: $('view-preview'), viewRestored: $('view-restored'), viewList: $('view-list'),
      refedit: $('refedit'), refBtn: $('ref-btn'), placeholder: $('ref-placeholder'), cells: $('ref-cells'), sheet: $('ref-sheet'),
      hint: $('ref-hint'), meta: $('ref-meta'), size: $('ref-size'), headers: $('headers'),
      summary: $('summary'), build: $('build'), buildIdle: $('build-idle'), buildPreviewActions: $('build-preview-actions'), buildDone: $('build-done'),
      generate: $('generate'), ctaLabel: $('cta-label'), ctaFill: $('cta-fill'), buildError: $('build-error'), buildTiming: $('build-timing'),
      placeOnSheet: $('place-on-sheet'), placeFill: $('place-fill'), placeLabel: $('place-label'), previewStartOver: $('preview-start-over'),
      doneSub: $('done-sub'), doneViewList: $('done-view-list'), doneStartOver: $('done-start-over'),
      previewThemes: $('preview-themes'), previewMount: $('preview-mount'), restoredNote: $('restored-note'), restoredBack: $('restored-back'), restoredMount: $('restored-mount'),
      listBack: $('list-back'), listEmpty: $('list-empty'), dashList: $('dash-list'),
    });

    els.refBtn.addEventListener('click', (e) => { e.stopPropagation(); state.picking ? finishPicking() : startPicking(); });
    els.refedit.addEventListener('click', () => { if (!state.picking && !state.busy) startPicking(); });
    document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && state.picking) cancelPicking(); });

    els.generate.addEventListener('click', generate);
    els.placeOnSheet.addEventListener('click', placeOnSheet);
    els.previewStartOver.addEventListener('click', startOver);
    els.doneStartOver.addEventListener('click', startOver);
    els.doneViewList.addEventListener('click', showList);
    els.showList.addEventListener('click', showList);
    els.listBack.addEventListener('click', () => showView('setup'));
    els.restoredBack.addEventListener('click', () => showView('setup'));

    els.previewThemes.addEventListener('click', (e) => {
      const btn = e.target.closest('button[data-theme]');
      if (!btn) return;
      state.previewTheme = btn.dataset.theme;
      [...els.previewThemes.querySelectorAll('button')].forEach((b) => b.setAttribute('aria-pressed', String(b === btn)));
      if (state.previewController) state.previewController.setTheme(window.DashRenderThemes.THEMES[state.previewTheme]);
    });

    window.addEventListener('resize', () => { if (activeScale) applyScale(activeScale.mountEl, activeScale.stageEl, activeScale.canvasSize); });

    if (state.excel) {
      Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelectionChanged);
    }

    renderProgress();
  }

  /* ---------- range picking ---------- */
  async function startPicking() {
    state.picking = true;
    state.beforePick = state.range;
    els.refedit.dataset.state = 'picking';
    els.refBtn.textContent = 'Done';
    els.hint.hidden = false;
    els.meta.hidden = true;
    const r = await host.readSelectionInfo();
    if (r && state.picking) showAddress(r.address);
  }

  async function onSelectionChanged() {
    if (state.picking) {
      const r = await host.readSelectionInfo();
      if (r && state.picking) showAddress(r.address);
      return;
    }
    await maybeOpenClickedDashboard();
  }

  async function finishPicking() {
    const r = await host.readSelectionInfo();
    state.picking = false;
    els.hint.hidden = true;
    if (r) setRange(r);
    else restoreRange(state.beforePick);
  }

  function cancelPicking() {
    state.picking = false;
    els.hint.hidden = true;
    restoreRange(state.beforePick);
  }

  function restoreRange(r) {
    if (r) { setRange(r); return; }
    state.range = null;
    els.refedit.dataset.state = 'empty';
    els.refBtn.textContent = 'Select range';
    els.placeholder.hidden = false;
    els.cells.hidden = true;
    els.sheet.hidden = true;
    els.meta.hidden = true;
    renderProgress();
  }

  function splitAddress(address) {
    const i = address.lastIndexOf('!');
    if (i < 0) return { sheet: '', cells: address };
    let sheet = address.slice(0, i);
    if (sheet.startsWith("'") && sheet.endsWith("'")) sheet = sheet.slice(1, -1).replace(/''/g, "'");
    return { sheet, cells: address.slice(i + 1) };
  }

  function showAddress(address) {
    const { sheet, cells } = splitAddress(address);
    els.placeholder.hidden = true;
    els.cells.hidden = false;
    els.sheet.hidden = !sheet;
    els.cells.textContent = cells;
    els.sheet.textContent = sheet;
  }

  function setRange(r) {
    state.range = r;
    showAddress(r.address);
    els.refedit.dataset.state = 'set';
    els.refBtn.textContent = 'Change';
    els.size.textContent = `${plural(Math.max(r.rows - 1, 0), 'row', 'rows')}, ${plural(r.cols, 'column', 'columns')}`;
    els.headers.innerHTML = '';
    els.meta.hidden = false;
    renderProgress();
  }

  function renderProgress() {
    document.querySelector('#view-setup .step').classList.toggle('done', !!state.range);
    els.generate.disabled = !state.range;
    els.summary.innerHTML = state.range
      ? `Ready to generate from <strong>${splitAddress(state.range.address).cells}</strong>`
      : 'Choose data above';
  }

  /* ---------- generate ---------- */
  async function generate() {
    if (state.busy || !state.range) return;
    state.busy = true;
    document.body.classList.add('busy');
    els.buildError.hidden = true;
    els.buildTiming.hidden = true;
    els.ctaFill.style.transitionDuration = '400ms';

    try {
      setCta('Reading your data…', 33);
      const { headers, dataRows, timing, totalRows } = await host.readRangeForEngine();

      setCta('Classifying columns…', 66);
      const analysis = window.DashEngine.analyzeTable(headers, dataRows, ',');

      setCta('Building the dashboard…', 100);
      const title = splitAddress(state.range.address).sheet || 'Dashboard';
      const layoutSpec = window.DashEngine.buildLayoutSpec(analysis, { title, subtitle: `${(totalRows - 1).toLocaleString('en-US')} rows` });

      state.analysis = analysis;
      state.layoutSpec = layoutSpec;
      state.sourceAddress = state.range.address;
      state.title = title;

      if (timing) {
        els.buildTiming.hidden = false;
        els.buildTiming.textContent = `Read ${totalRows.toLocaleString('en-US')} rows in ${timing.totalMs}ms (values ${timing.valuesMs}ms, format sample ${timing.formatMs}ms)`;
      }

      showPreview();
      setFooterState('preview');
    } catch (err) {
      els.buildError.textContent = `Could not generate the dashboard: ${err && err.message ? err.message : String(err)}`;
      els.buildError.hidden = false;
    } finally {
      state.busy = false;
      document.body.classList.remove('busy');
      els.ctaFill.style.transitionDuration = '0ms';
      els.ctaFill.style.width = '0';
      els.ctaLabel.textContent = 'Generate dashboard';
    }
  }

  function setCta(label, pct) {
    els.ctaLabel.textContent = label;
    els.ctaFill.style.width = `${pct}%`;
  }

  function showPreview() {
    // showView first: applyScale measures the stage's actual (laid-out)
    // width, which is 0 while its view still has the `hidden` attribute
    // (display:none) — measuring before revealing it silently produced a
    // scale of 1 (no shrink at all) every time.
    showView('preview');
    els.previewMount.innerHTML = '';
    state.previewTheme = 'light';
    [...els.previewThemes.querySelectorAll('button')].forEach((b) => b.setAttribute('aria-pressed', String(b.dataset.theme === 'light')));
    state.previewController = window.DashRenderDom.mount(
      els.previewMount,
      state.analysis,
      window.DashRenderThemes.THEMES.light,
      { title: state.title, subtitle: `${state.analysis.rowCount.toLocaleString('en-US')} rows` }
    );
    activeScale = { mountEl: els.previewMount, stageEl: els.previewMount.parentElement, canvasSize: state.previewController.getLayoutSpec().canvas };
    applyScale(activeScale.mountEl, activeScale.stageEl, activeScale.canvasSize);
  }

  /* ---------- place on sheet ---------- */
  async function placeOnSheet() {
    if (state.busy) return;
    state.busy = true;
    els.buildError.hidden = true;
    els.placeFill.style.transitionDuration = '300ms';

    try {
      setPlaceCta('Rendering image…', 40);
      const theme = window.DashRenderThemes.THEMES[state.previewTheme];
      const pcState = state.previewController.getState();
      const widgets = window.DashEngine.recomputeLayout(state.analysis, state.layoutSpec.widgets, { activeFilters: pcState.activeFilters, sort: pcState.sort });
      const currentSpec = { canvas: state.layoutSpec.canvas, widgets };
      const columnsByName = window.DashEngine.Aggregate.byName(state.analysis.columns);
      const pngBase64 = await window.DashRenderShare.renderPngBase64(currentSpec, theme, columnsByName, { scale: 2 });

      setPlaceCta('Placing on sheet…', 80);
      const snapshot = window.DashEngine.Layout.buildStorageSnapshot(currentSpec, state.analysis.columns);
      const mapping = state.analysis.columns.map((c) => ({ name: c.name, role: c.decision.role, aggregation: c.decision.aggregation }));

      setPlaceCta('Saving…', 100);
      const { shapeName } = await host.placeDashboard({
        pngBase64, canvasSize: currentSpec.canvas, snapshot, mapping, sourceAddress: state.sourceAddress, title: state.title,
      });

      state.result = { shapeName };
      els.doneSub.textContent = state.excel
        ? `It's on the sheet as a picture named "${shapeName}". Anyone opening the file just sees a picture; with this add-in, reopening it — from "Dashboards in this workbook" or by clicking it — restores the interactive version without re-reading the sheet.`
        : `Preview mode (no Excel host here): "${shapeName}" was saved in memory instead of workbook.settings. Open "Dashboards in this workbook" below to see the restored, no-recompute view.`;
      setFooterState('done');
    } catch (err) {
      els.buildError.textContent = `Could not place the dashboard: ${err && err.message ? err.message : String(err)}`;
      els.buildError.hidden = false;
    } finally {
      state.busy = false;
      els.placeFill.style.transitionDuration = '0ms';
      els.placeFill.style.width = '0';
      els.placeLabel.textContent = 'Place on sheet';
    }
  }

  function setPlaceCta(label, pct) {
    els.placeLabel.textContent = label;
    els.placeFill.style.width = `${pct}%`;
  }

  /* ---------- dashboards list / restore ---------- */
  async function showList() {
    showView('list');
    els.dashList.innerHTML = '';
    els.listEmpty.hidden = true;
    let dashboards = [];
    try {
      dashboards = await host.listDashboards();
    } catch (e) {
      els.listEmpty.hidden = false;
      els.listEmpty.textContent = `Could not read saved dashboards: ${e && e.message ? e.message : String(e)}`;
      return;
    }
    if (!dashboards.length) {
      els.listEmpty.hidden = false;
      return;
    }
    for (const d of dashboards) {
      els.dashList.appendChild(renderDashListItem(d));
    }
  }

  function renderDashListItem(d) {
    const li = document.createElement('li');
    const title = document.createElement('p');
    title.className = 'dash-item-title';
    title.textContent = d.payload.title || d.shapeName;
    const meta = document.createElement('p');
    meta.className = 'dash-item-meta';
    const when = d.payload.generatedAt ? new Date(d.payload.generatedAt).toLocaleString() : 'unknown time';
    meta.textContent = `${d.payload.sourceAddress || ''} · generated ${when}`;
    const actions = document.createElement('div');
    actions.className = 'dash-item-actions';
    const openBtn = document.createElement('button');
    openBtn.type = 'button';
    openBtn.className = 'primary';
    openBtn.textContent = 'Open';
    openBtn.addEventListener('click', () => openRestoredDashboard(d));
    actions.appendChild(openBtn);
    li.append(title, meta, actions);
    return li;
  }

  function openRestoredDashboard(d) {
    showView('restored'); // before applyScale — see the comment in showPreview
    els.restoredMount.innerHTML = '';
    const theme = window.DashRenderThemes.THEMES.light;
    window.DashRenderDom.mountFrozen(els.restoredMount, d.payload.layoutSpec, theme);
    const when = d.payload.generatedAt ? new Date(d.payload.generatedAt).toLocaleString() : 'unknown time';
    els.restoredNote.textContent = `${d.payload.title || d.shapeName} · generated ${when} · restored without re-reading the sheet`;
    activeScale = { mountEl: els.restoredMount, stageEl: els.restoredMount.parentElement, canvasSize: d.payload.layoutSpec.canvas };
    applyScale(activeScale.mountEl, activeScale.stageEl, activeScale.canvasSize);
  }

  async function maybeOpenClickedDashboard() {
    if (state.busy || state.picking) return;
    const maybeShape = await host.probeNonRangeSelection();
    if (!maybeShape) return;
    try {
      const dashboards = await host.listDashboards();
      if (dashboards.length === 1) openRestoredDashboard(dashboards[0]);
      else if (dashboards.length > 1) showList();
    } catch (e) { /* leave the pane as it is */ }
  }

  /* ---------- view/footer state ---------- */
  function showView(name) {
    els.viewSetup.hidden = name !== 'setup';
    els.viewPreview.hidden = name !== 'preview';
    els.viewRestored.hidden = name !== 'restored';
    els.viewList.hidden = name !== 'list';
    els.build.hidden = name === 'restored' || name === 'list';
  }

  function setFooterState(s) {
    els.buildIdle.hidden = s !== 'idle';
    els.buildPreviewActions.hidden = s !== 'preview';
    els.buildDone.hidden = s !== 'done';
  }

  function applyScale(mountEl, stageEl, canvasSize) {
    const available = stageEl.clientWidth || mountEl.parentElement.clientWidth || canvasSize.width;
    const scale = Math.min(1, available / canvasSize.width);
    mountEl.style.transformOrigin = 'top left';
    mountEl.style.transform = `scale(${scale})`;
    stageEl.style.height = `${Math.round(canvasSize.height * scale)}px`;
    stageEl.style.overflowX = 'hidden';
  }

  function startOver() {
    state.picking = false;
    state.range = null;
    state.beforePick = null;
    state.analysis = null;
    state.layoutSpec = null;
    state.previewController = null;
    state.result = null;
    activeScale = null;
    els.previewMount.innerHTML = '';
    els.restoredMount.innerHTML = '';
    els.buildError.hidden = true;
    els.buildTiming.hidden = true;
    els.hint.hidden = true;
    restoreRange(null);
    showView('setup');
    setFooterState('idle');
    window.scrollTo(0, 0);
  }
})();
