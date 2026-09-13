/*
 * Dash Architect task pane — the control panel: range picking, reading the
 * selected range, and reacting to what the dialog reports back. The actual
 * interactive dashboard (render/dom.js#mount / #mountFrozen) lives in its
 * own window, addin/dashboard-dialog.html, opened via
 * Office.context.ui.displayDialogAsync — this pane and that dialog talk
 * over addin/dialog-messaging.js's chunked JSON channel (see the "dashboard
 * dialog" section below). This file:
 *   1. reads the selected range (addin/excel-io.js, chunked, format
 *      sampled) and feeds it to engine.analyzeTable exactly like the CSV
 *      path does,
 *   2. opens the dialog and streams it {analysis, meta} so it can mount the
 *      live dashboard itself,
 *   3. on the dialog's "Place image on sheet" message: rasterizes the
 *      dialog's current (already filtered/sorted) spec (render/share.js),
 *      inserts the PNG via worksheet.shapes.addImage, and saves a
 *      self-contained snapshot (engine/layout.js#buildStorageSnapshot) +
 *      the column mapping into workbook.settings, keyed to the picture's
 *      name (addin/dashboard-io.js) — then messages the result back,
 *   4. reopens a saved dashboard — via the reliable "Dashboards in this
 *      workbook" list, or a best-effort guess when the selection changes
 *      to something that isn't a normal cell range — in the same dialog,
 *      frozen/no-recompute (render/dom.js#mountFrozen).
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
    sourceAddress: null,
    title: null,
    result: null,
  };
  let host = null;
  let currentDialog = null; // the dashboard-dialog.html window currently open, if any — see openDialog()

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
      showList: $('show-list'), viewSetup: $('view-setup'), viewList: $('view-list'),
      refedit: $('refedit'), refBtn: $('ref-btn'), placeholder: $('ref-placeholder'), cells: $('ref-cells'), sheet: $('ref-sheet'),
      hint: $('ref-hint'), meta: $('ref-meta'), size: $('ref-size'), headers: $('headers'),
      summary: $('summary'), build: $('build'), buildIdle: $('build-idle'), buildOpen: $('build-open'), buildDone: $('build-done'),
      generate: $('generate'), ctaLabel: $('cta-label'), ctaFill: $('cta-fill'), buildError: $('build-error'), buildTiming: $('build-timing'),
      openStartOver: $('open-start-over'),
      doneSub: $('done-sub'), doneViewList: $('done-view-list'), doneStartOver: $('done-start-over'),
      listBack: $('list-back'), listEmpty: $('list-empty'), dashList: $('dash-list'),
      debugPanel: $('debug-panel'), debugLog: $('debug-log'),
    });

    els.refBtn.addEventListener('click', (e) => { e.stopPropagation(); state.picking ? finishPicking() : startPicking(); });
    els.refedit.addEventListener('click', () => { if (!state.picking && !state.busy) startPicking(); });
    document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && state.picking) cancelPicking(); });

    els.generate.addEventListener('click', generate);
    els.openStartOver.addEventListener('click', startOver);
    els.doneStartOver.addEventListener('click', startOver);
    els.doneViewList.addEventListener('click', showList);
    els.showList.addEventListener('click', showList);
    els.listBack.addEventListener('click', () => showView('setup'));

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

  /* ---------- debug (temporary) — surfaces dialog/messaging lifecycle
     events with no console available in the add-in host; remove once the
     dialog architecture has proven itself in real Excel. ---------- */
  function logDebug(text) {
    els.debugPanel.hidden = false;
    const line = `[${new Date().toLocaleTimeString()}] ${text}`;
    els.debugLog.textContent = els.debugLog.textContent ? `${els.debugLog.textContent}\n${line}` : line;
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

      setCta('Opening dashboard window…', 100);
      const title = splitAddress(state.range.address).sheet || 'Dashboard';

      state.analysis = analysis;
      state.sourceAddress = state.range.address;
      state.title = title;

      if (timing) {
        els.buildTiming.hidden = false;
        els.buildTiming.textContent = `Read ${totalRows.toLocaleString('en-US')} rows in ${timing.totalMs}ms (values ${timing.valuesMs}ms, format sample ${timing.formatMs}ms)`;
      }

      await openLiveDashboard();
      setFooterState('open');
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

  /* ---------- dashboard dialog: opens the interactive dashboard in its own
     window (Office.context.ui.displayDialogAsync) instead of the cramped
     task pane. The dialog can't reach Excel itself, so all it gets is JSON
     over addin/dialog-messaging.js's chunked channel, and the one thing it
     needs Excel for — placing the image — it asks this pane to do. See
     SPEC.md for the full protocol writeup. ---------- */

  function dialogUrl() {
    return new URL('dashboard-dialog.html', location.href).href;
  }

  function openDialogExcel(url) {
    return new Promise((resolve, reject) => {
      Office.context.ui.displayDialogAsync(url, { width: 85, height: 85, promptBeforeOpen: false }, (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) reject(result.error);
        else resolve(result.value);
      });
    });
  }

  // Mimics the native dialog object's shape (messageChild/addEventHandler/
  // close) over window.open + postMessage, so previewHost can exercise the
  // exact same chunking/handshake/place-on-sheet code paths outside Excel.
  // The dialog page's own "Place image on sheet" button always goes through
  // this message round trip too (never a direct local call) — otherwise a
  // browser-only shortcut would stop proving anything about the real path.
  function openDialogPreview(url) {
    const win = window.open(url, 'dash-dialog', 'width=1200,height=850');
    const handlers = {};
    const onMessage = (e) => {
      if (e.source !== win || !e.data || e.data.__dashDialog !== true) return;
      (handlers[Office.EventType.DialogMessageReceived] || []).forEach((cb) => cb({ message: e.data.message }));
    };
    window.addEventListener('message', onMessage);
    const closedTimer = setInterval(() => {
      if (win.closed) {
        clearInterval(closedTimer);
        window.removeEventListener('message', onMessage);
        (handlers[Office.EventType.DialogEventReceived] || []).forEach((cb) => cb({ error: 12006 }));
      }
    }, 400);
    return {
      messageChild(str) { if (!win.closed) win.postMessage({ __dashDialog: true, message: str }, location.origin); },
      addEventHandler(type, cb) { (handlers[type] || (handlers[type] = [])).push(cb); },
      close() { clearInterval(closedTimer); window.removeEventListener('message', onMessage); if (!win.closed) win.close(); },
    };
  }

  // Only one dialog at a time — close whatever's open before starting a new one.
  async function openDialog(url) {
    if (currentDialog) {
      try { currentDialog.close(); } catch (e) { /* already gone */ }
      currentDialog = null;
    }
    const dlg = state.excel ? await openDialogExcel(url) : openDialogPreview(url);
    currentDialog = dlg;
    dlg.addEventHandler(Office.EventType.DialogEventReceived, () => { if (currentDialog === dlg) currentDialog = null; });
    return dlg;
  }

  function sendDataToDialog(dlg, data) {
    const Msg = window.DashDialogMessaging;
    const requestId = Msg.sendChunked((str) => dlg.messageChild(str), Msg.KIND.DATA, data);
    logDebug(`dialog: sent ${data.mode} data (request ${requestId}).`);
  }

  async function openLiveDashboard() {
    const dlg = await openDialog(dialogUrl());
    logDebug('dialog: opened, waiting for ready…');
    const Msg = window.DashDialogMessaging;

    let resolveReady;
    const readyPromise = new Promise((resolve) => { resolveReady = resolve; });
    const readyReceiver = Msg.createChunkReceiver(Msg.KIND.READY, () => resolveReady());
    const placeReceiver = Msg.createChunkReceiver(Msg.KIND.PLACE_ON_SHEET, (spec, requestId) => handlePlaceOnSheet(dlg, spec, requestId));
    dlg.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
      readyReceiver(arg.message) || placeReceiver(arg.message);
    });

    await readyPromise;
    logDebug('dialog: ready, sending data.');
    sendDataToDialog(dlg, {
      mode: 'live',
      analysis: state.analysis,
      meta: { title: state.title, subtitle: `${state.analysis.rowCount.toLocaleString('en-US')} rows` },
    });
  }

  async function openRestoredDashboard(d) {
    showView('setup');
    try {
      const dlg = await openDialog(dialogUrl());
      logDebug('dialog: opened (restored), waiting for ready…');
      const Msg = window.DashDialogMessaging;

      let resolveReady;
      const readyPromise = new Promise((resolve) => { resolveReady = resolve; });
      const readyReceiver = Msg.createChunkReceiver(Msg.KIND.READY, () => resolveReady());
      dlg.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => readyReceiver(arg.message));

      await readyPromise;
      sendDataToDialog(dlg, { mode: 'frozen', layoutSpec: d.payload.layoutSpec });
    } catch (err) {
      logDebug(`openRestoredDashboard: could not open the dialog — ${err && err.message ? err.message : String(err)}`);
    }
  }

  // Runs the actual Excel work for "Place image on sheet" — unchanged from
  // before the dialog existed, apart from taking the already-resolved
  // {canvas, widgets, theme} the dialog sends instead of reading a local
  // mount controller (that controller now lives in the dialog's own window,
  // a separate JS realm this pane has no access to).
  async function placeOnSheet(spec) {
    try {
      const theme = window.DashRenderThemes.THEMES[spec.theme || 'light'];
      const currentSpec = { canvas: spec.canvas, widgets: spec.widgets };
      const columnsByName = window.DashEngine.Aggregate.byName(state.analysis.columns);
      const pngBase64 = await window.DashRenderShare.renderPngBase64(currentSpec, theme, columnsByName, { scale: 2 });

      const snapshot = window.DashEngine.Layout.buildStorageSnapshot(currentSpec, state.analysis.columns);
      const mapping = state.analysis.columns.map((c) => ({ name: c.name, role: c.decision.role, aggregation: c.decision.aggregation }));

      const { shapeName } = await host.placeDashboard({
        pngBase64, canvasSize: currentSpec.canvas, snapshot, mapping, sourceAddress: state.sourceAddress, title: state.title,
      });
      return { ok: true, shapeName };
    } catch (err) {
      return { ok: false, error: err && err.message ? err.message : String(err) };
    }
  }

  async function handlePlaceOnSheet(dlg, spec, requestId) {
    logDebug(`place-on-sheet: request received (${requestId}).`);
    const result = await placeOnSheet(spec);
    const Msg = window.DashDialogMessaging;
    Msg.sendChunked((str) => dlg.messageChild(str), Msg.KIND.PLACE_ON_SHEET_RESULT, result, null, requestId);
    if (result.ok) {
      state.result = { shapeName: result.shapeName };
      els.doneSub.textContent = state.excel
        ? `It's on the sheet as a picture named "${result.shapeName}". Anyone opening the file just sees a picture; with this add-in, reopening it — from "Dashboards in this workbook" or by clicking it — opens the interactive version again in a dialog, without re-reading the sheet.`
        : `Preview mode (no Excel host here): "${result.shapeName}" was saved in memory instead of workbook.settings. Open "Dashboards in this workbook" below to see the restored, no-recompute view.`;
      setFooterState('done');
      logDebug(`place-on-sheet: placed as "${result.shapeName}".`);
    } else {
      logDebug(`place-on-sheet: failed — ${result.error}`);
    }
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

  async function maybeOpenClickedDashboard() {
    if (state.busy || state.picking) return;
    const maybeShape = await host.probeNonRangeSelection();
    if (!maybeShape) return;
    try {
      const dashboards = await host.listDashboards();
      if (dashboards.length === 1) await openRestoredDashboard(dashboards[0]);
      else if (dashboards.length > 1) showList();
    } catch (e) { /* leave the pane as it is */ }
  }

  /* ---------- view/footer state ---------- */
  function showView(name) {
    els.viewSetup.hidden = name !== 'setup';
    els.viewList.hidden = name !== 'list';
    els.build.hidden = name === 'list';
  }

  function setFooterState(s) {
    els.buildIdle.hidden = s !== 'idle';
    els.buildOpen.hidden = s !== 'open';
    els.buildDone.hidden = s !== 'done';
  }

  function startOver() {
    state.picking = false;
    state.range = null;
    state.beforePick = null;
    state.analysis = null;
    state.sourceAddress = null;
    state.title = null;
    state.result = null;
    if (currentDialog) {
      try { currentDialog.close(); } catch (e) { /* already gone */ }
      currentDialog = null;
    }
    els.buildError.hidden = true;
    els.buildTiming.hidden = true;
    els.hint.hidden = true;
    restoreRange(null);
    showView('setup');
    setFooterState('idle');
    window.scrollTo(0, 0);
  }
})();
