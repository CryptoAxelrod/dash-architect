/*
 * Dash Architect task pane — the control panel: range/table picking,
 * reading the selected source, and reacting to what the dialog reports
 * back. The actual interactive dashboard (render/dom.js#mount / #mountFrozen)
 * lives in its own window, addin/dashboard-dialog.html, opened via
 * Office.context.ui.displayDialogAsync — this pane and that dialog talk
 * over addin/dialog-messaging.js's chunked JSON channel. This file:
 *   1. reads the selected source (addin/excel-io.js, chunked, format
 *      sampled) — a plain range by address, or an Excel Table by id if the
 *      selection falls entirely inside one — and feeds it to
 *      engine.analyzeTable exactly like the CSV path does,
 *   2. if there are more than 6 columns, shows a mapping screen so the user
 *      can review/correct roles before anything is built (CLAUDE.md §7:
 *      those picks persist across every later Refresh/Change data range,
 *      layered back onto fresh auto-classification, never lost to it),
 *   3. opens the dialog and streams it {analysis, meta, seedState?} so it
 *      can mount the live dashboard itself,
 *   4. on the dialog's REFRESH_REQUEST / CHANGE_RANGE_REQUEST / PLACE_ON_SHEET
 *      messages: re-reads the source / lets the user pick a different one /
 *      rasterizes+places the image, respectively — see the handlers below,
 *   5. remembers the dialog's live state (STATE_UPDATE) so "Open dashboard"
 *      can reopen it without losing filters/sort/theme/settings, and
 *   6. reopens a saved (placed) dashboard — via the reliable "Dashboards in
 *      this workbook" list, or a best-effort guess when the selection
 *      changes to something that isn't a normal cell range — in the same
 *      dialog, frozen/no-recompute (render/dom.js#mountFrozen).
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
  const plural = (n, one, many) => `${n.toLocaleString('en-US')} ${n === 1 ? one : many}`;

  const els = {};
  const state = {
    excel: false,
    picking: false,
    pickingFor: 'generate', // 'generate' | 'change-range' | 'list-change-range' — which flow finishPicking()/the CTA click should feed into
    range: null, // live {address, rows, cols} while picking, for the address/size display
    beforePick: null,
    busy: false,
    source: null, // {kind:'range', address, headers} | {kind:'table', id, name}
    analysis: null, // effective (override-applied) analysis of the currently open/last-open dashboard
    roleOverrides: null, // {colName: {role, chartEligible}} | null — from the mapping screen, reapplied on every Refresh
    title: null,
    theme: 'light', // chosen on the pre-generate theme/palette step; sticky across dashboards as a "last picked" default
    palette: 'ocean',
    lastDialogState: null, // {activeFilters, sort, theme, widgetConfig} from the dialog's STATE_UPDATE — survives the dialog closing
    result: null,
    currentShapeName: null, // shapeName of the dashboard currently open in the dialog, once it has one (placed, or reopened from the list)
    pendingChangeRange: null, // {dlg, raw, requestId} while pickingFor === 'change-range'
    sourceBeforeChangeRange: null,
    rangeBeforeChangeRange: null,
    pendingListChangeRange: null, // {dashboard, rangeBeforePick, sourceBeforePick} while pickingFor === 'list-change-range' — a headless change-range for a dashboard that isn't open in any dialog
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

  // The only Office.js-touching helper not tied to a single host method:
  // resolves a stored `source` descriptor back into a live Excel.Range,
  // re-reading it fresh every time (never the address/table captured once
  // at pick time) so Refresh naturally picks up a grown Excel Table and a
  // renamed table/sheet doesn't just silently break.
  async function resolveSourceRange(ctx, source) {
    if (source.kind === 'table') {
      const t = ctx.workbook.tables.getItemOrNullObject(source.id);
      t.load('name');
      await ctx.sync();
      if (t.isNullObject) throw new Error(`Table "${source.name}" no longer exists — it may have been deleted.`);
      source.name = t.name; // table.id survives a rename; refresh the cached display name
      return t.getRange();
    }
    const { sheet: sheetName, cells } = splitAddress(source.address);
    const sheet = ctx.workbook.worksheets.getItemOrNullObject(sheetName);
    await ctx.sync();
    if (sheet.isNullObject) throw new Error(`Sheet "${sheetName}" no longer exists — it may have been deleted or renamed.`);
    return sheet.getRange(cells);
  }

  // Deliberately no "did the header row change since this was picked?"
  // guard here (there used to be one) — renaming a column is a normal edit,
  // not a sign the range shifted, and it was blocking exactly that: reopen
  // just re-reads and reclassifies against whatever headers are there now.
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
    // Only called once, when the user confirms a pick ("Done") — not on
    // every DocumentSelectionChanged tick, which only needs the cheap
    // readSelectionInfo above. Detects whether the confirmed selection
    // falls entirely inside one Excel Table (geometry, not by any
    // dedicated "table at range" API — see SPEC.md).
    async detectSource() {
      return Excel.run(async (ctx) => {
        const range = ctx.workbook.getSelectedRange();
        range.load('address,rowIndex,columnIndex,rowCount,columnCount');
        const tables = range.worksheet.tables;
        tables.load('items/name,items/id');
        await ctx.sync();

        const tableGeoms = tables.items.map((t) => {
          const r = t.getRange();
          r.load('rowIndex,columnIndex,rowCount,columnCount');
          return { table: t, r };
        });
        await ctx.sync();

        const hit = tableGeoms.find(({ r }) =>
          range.rowIndex >= r.rowIndex && range.columnIndex >= r.columnIndex &&
          range.rowIndex + range.rowCount <= r.rowIndex + r.rowCount &&
          range.columnIndex + range.columnCount <= r.columnIndex + r.columnCount
        );
        if (hit) return { kind: 'table', id: hit.table.id, name: hit.table.name };

        const headerRow = range.getCell(0, 0).getResizedRange(0, range.columnCount - 1);
        headerRow.load('values');
        await ctx.sync();
        return { kind: 'range', address: range.address, headers: headerRow.values[0].map((v) => (v == null ? '' : String(v))) };
      });
    },
    async readRangeForEngine(source) {
      return Excel.run(async (ctx) => {
        const range = await resolveSourceRange(ctx, source);
        return window.DashAddinExcelIo.readRangeForEngine(ctx, range);
      });
    },
    async placeDashboard(args) {
      return Excel.run((ctx) => window.DashAddinDashboardIo.generateAndPlace(ctx, args));
    },
    // Headless update of an already-placed dashboard — list-triggered
    // rename/change-range, neither of which opens the dialog. Reuses the
    // existing shape/settings identity instead of minting a new one.
    async regenerateDashboard(shapeName, args) {
      return Excel.run((ctx) => window.DashAddinDashboardIo.generateAndPlace(ctx, Object.assign({}, args, { shapeName })));
    },
    // Settings-only save at "Generate dashboard" time — no picture yet, see
    // dashboard-io.js#saveGeneratedDraft. `shapeName` is minted client-side
    // (pure helper, no Excel context needed) so the same id can be reused
    // once/if the user actually places it later.
    async saveDraft(shapeName, args) {
      return Excel.run((ctx) => window.DashAddinDashboardIo.saveGeneratedDraft(ctx, shapeName, args));
    },
    async renameDraft(shapeName, newTitle) {
      return Excel.run((ctx) => window.DashAddinDashboardIo.renameDraft(ctx, shapeName, newTitle));
    },
    async listDashboards() {
      return Excel.run((ctx) => window.DashAddinDashboardIo.listDashboards(ctx));
    },
    async deleteDashboard(shapeName) {
      return Excel.run((ctx) => window.DashAddinDashboardIo.deleteDashboard(ctx, shapeName));
    },
    async countOrphanedDashboardShapes() {
      return Excel.run((ctx) => window.DashAddinDashboardIo.countOrphanedDashboardShapes(ctx));
    },
    // Last fallback in the dashboard-title chain (table name, then sheet
    // name, then this) — `workbook.name` is the file name including its
    // extension (e.g. "Sales.xlsx"), stripped here since a title doesn't
    // want the extension.
    async getWorkbookName() {
      return Excel.run(async (ctx) => {
        const wb = ctx.workbook;
        wb.load('name');
        await ctx.sync();
        return wb.name ? wb.name.replace(/\.[^./\\]+$/, '') : null;
      });
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
  // No table simulation here — the bundled fixture is always a plain range.
  const previewDashboards = [];
  const PREVIEW_ADDRESS = "'Store sales'!A1:J901";

  async function readPreviewFixture() {
    const text = await fetch('../fixtures/01_sales_timeseries.csv').then((r) => r.text());
    return window.DashEngineCsv.parseCsv(text).rows;
  }

  const previewHost = {
    async readSelectionInfo() {
      return { address: PREVIEW_ADDRESS, rows: 901, cols: 10 };
    },
    async detectSource() {
      const rows = await readPreviewFixture();
      return { kind: 'range', address: PREVIEW_ADDRESS, headers: rows[0] };
    },
    async readRangeForEngine(source) {
      const t0 = performance.now();
      const rows = await readPreviewFixture();
      const t1 = performance.now();
      const [headers, ...dataRows] = rows;
      return { headers, dataRows, totalRows: rows.length, cols: headers.length, timing: { valuesMs: Math.round(t1 - t0), formatMs: 0, totalMs: Math.round(t1 - t0) } };
    },
    async placeDashboard(args) {
      const shapeName = args.shapeName || window.DashAddinDashboardIo.makeShapeName(Date.now());
      const payload = window.DashAddinDashboardIo.buildStoragePayload(args);
      const idx = previewDashboards.findIndex((d) => d.shapeName === shapeName);
      if (idx === -1) previewDashboards.push({ shapeName, payload });
      else previewDashboards[idx] = { shapeName, payload };
      return { shapeName };
    },
    async regenerateDashboard(shapeName, args) {
      const idx = previewDashboards.findIndex((d) => d.shapeName === shapeName);
      if (idx === -1) throw new Error(`Preview dashboard "${shapeName}" not found.`);
      const payload = window.DashAddinDashboardIo.buildStoragePayload(args);
      previewDashboards[idx] = { shapeName, payload };
      return { shapeName };
    },
    async saveDraft(shapeName, args) {
      const payload = window.DashAddinDashboardIo.buildStoragePayload(Object.assign({}, args, { placed: false }));
      const idx = previewDashboards.findIndex((d) => d.shapeName === shapeName);
      if (idx === -1) previewDashboards.push({ shapeName, payload });
      else previewDashboards[idx] = { shapeName, payload };
    },
    async renameDraft(shapeName, newTitle) {
      const idx = previewDashboards.findIndex((d) => d.shapeName === shapeName);
      if (idx === -1) throw new Error(`Preview dashboard "${shapeName}" not found.`);
      previewDashboards[idx] = { shapeName, payload: Object.assign({}, previewDashboards[idx].payload, { title: newTitle }) };
    },
    async listDashboards() {
      return previewDashboards.map((d) => ({ shapeName: d.shapeName, payload: d.payload }));
    },
    async deleteDashboard(shapeName) {
      const idx = previewDashboards.findIndex((d) => d.shapeName === shapeName);
      if (idx !== -1) previewDashboards.splice(idx, 1);
    },
    // No real sheet/shapes outside Excel — placeDashboard above never
    // creates one that could go orphaned.
    async countOrphanedDashboardShapes() {
      return 0;
    },
    async getWorkbookName() {
      return 'Preview Workbook';
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
      showList: $('show-list'), versionBadge: $('version-badge'), viewSetup: $('view-setup'), viewList: $('view-list'), viewMapping: $('view-mapping'),
      refedit: $('refedit'), refBtn: $('ref-btn'), placeholder: $('ref-placeholder'), cells: $('ref-cells'), sheet: $('ref-sheet'),
      tableBadge: $('ref-table-badge'),
      hint: $('ref-hint'), meta: $('ref-meta'), size: $('ref-size'), headers: $('headers'),
      summary: $('summary'), build: $('build'), buildIdle: $('build-idle'), buildOpen: $('build-open'), buildDone: $('build-done'),
      generate: $('generate'), ctaLabel: $('cta-label'), ctaFill: $('cta-fill'), buildError: $('build-error'), buildTiming: $('build-timing'),
      reopenRow: $('reopen-row'), reopenDashboard: $('reopen-dashboard'), newDashboard: $('new-dashboard'), cancelChangeRange: $('cancel-change-range'),
      openStartOver: $('open-start-over'),
      doneSub: $('done-sub'), doneSaveHint: $('done-save-hint'), doneViewList: $('done-view-list'), doneStartOver: $('done-start-over'),
      listBack: $('list-back'), listEmpty: $('list-empty'), dashList: $('dash-list'),
      startupList: $('startup-list'), startupDashItems: $('startup-dash-items'), startupOrphanHint: $('startup-orphan-hint'),
      mappingList: $('mapping-list'), mappingGenerate: $('mapping-generate'), mappingCancel: $('mapping-cancel'),
      viewTheme: $('view-theme'), themePicker: $('theme-picker'), themeGenerate: $('theme-generate'), themeCancel: $('theme-cancel'),
      themeTitleInput: $('theme-title-input'),
    });

    els.refBtn.addEventListener('click', (e) => { e.stopPropagation(); state.picking ? finishPicking() : startPicking(); });
    els.refedit.addEventListener('click', () => { if (!state.picking && !state.busy) startPicking(); });
    document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && state.picking) cancelPicking(); });

    els.generate.addEventListener('click', () => {
      if (state.pickingFor === 'change-range') completeChangeRangePicking();
      else if (state.pickingFor === 'list-change-range') completeListChangeRangePicking();
      else generate();
    });
    els.themeGenerate.addEventListener('click', finalizeGenerate);
    els.themeCancel.addEventListener('click', () => showView('setup'));
    // finalizeGenerate() trims this and falls back to a bare "Dashboard" if
    // the user clears the field entirely and generates anyway.
    els.themeTitleInput.addEventListener('input', () => { state.title = els.themeTitleInput.value; });
    els.reopenDashboard.addEventListener('click', reopenDashboard);
    els.newDashboard.addEventListener('click', startNewDashboard);
    els.cancelChangeRange.addEventListener('click', () => {
      if (state.pickingFor === 'list-change-range') cancelListChangeRangePicking();
      else cancelChangeRangePicking();
    });
    els.openStartOver.addEventListener('click', startOver);
    els.doneStartOver.addEventListener('click', startOver);
    els.doneViewList.addEventListener('click', showList);
    els.showList.addEventListener('click', showList);
    els.listBack.addEventListener('click', () => showView('setup'));

    if (state.excel) {
      Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelectionChanged);
    }

    renderProgress();
    refreshStartupList();

    // TEMPORARY — see addin/version.js; remove with it before release.
    if (window.DASH_BUILD_VERSION != null) els.versionBadge.textContent = `v${window.DASH_BUILD_VERSION}`;
  }

  // Shows saved dashboards directly on the startup screen, above "Choose
  // your data" — not gated behind the hamburger menu's full list view,
  // which stays around for later (mid-flow) access. An empty workbook
  // shows the normal creation screen exactly as before; nothing else on
  // view-setup moves or hides to make room for this. Re-run after any
  // moment the set of saved dashboards could have changed (startOver, after
  // a fresh placement) so it never shows stale contents.
  async function refreshStartupList() {
    let dashboards = [];
    try {
      dashboards = await host.listDashboards();
    } catch (e) {
      return; // leave the startup list hidden — the hamburger menu's list view still reports read failures on its own
    }
    els.startupList.hidden = !dashboards.length;
    els.startupDashItems.innerHTML = '';
    for (const d of dashboards) {
      els.startupDashItems.appendChild(renderDashListItem(d));
    }
    // Same "book was closed without saving" signal as showList() — worth
    // surfacing here too since this is the first screen the pane shows.
    els.startupOrphanHint.hidden = true;
    if (!dashboards.length) {
      let orphanCount = 0;
      try {
        orphanCount = await host.countOrphanedDashboardShapes();
      } catch (e) { /* leave the hint hidden */ }
      if (orphanCount > 0) {
        els.startupOrphanHint.hidden = false;
        els.startupOrphanHint.textContent = `There ${orphanCount === 1 ? 'is' : 'are'} ${orphanCount} dashboard picture${orphanCount === 1 ? '' : 's'} on the sheets, but no saved settings to match — the workbook was likely closed without saving after ${orphanCount === 1 ? 'it was' : 'they were'} placed. Generate again to get the interactive version back.`;
      }
    }
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
    if (!r) { restoreRange(state.beforePick); return; }
    setRange(r);
    try {
      state.source = await host.detectSource();
      updateSourceBadge();
      renderProgress(); // sourceLabel() only has something to show once state.source resolves, above
    } catch (err) {
      els.buildError.textContent = `Could not read the selection: ${err && err.message ? err.message : String(err)}`;
      els.buildError.hidden = false;
    }
  }

  function cancelPicking() {
    state.picking = false;
    els.hint.hidden = true;
    restoreRange(state.beforePick);
  }

  function restoreRange(r) {
    if (r) { setRange(r); return; }
    state.range = null;
    state.source = null;
    els.refedit.dataset.state = 'empty';
    els.refBtn.textContent = 'Select range';
    els.placeholder.hidden = false;
    els.cells.hidden = true;
    els.sheet.hidden = true;
    els.meta.hidden = true;
    updateSourceBadge();
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

  // Shown only after "Done" resolves detectSource() — during the drag/pick
  // itself the badge doesn't know yet whether this is a table (see
  // excelHost.detectSource's doc comment on why that check isn't live).
  function updateSourceBadge() {
    const isTable = state.source && state.source.kind === 'table';
    els.tableBadge.hidden = !isTable;
    // Only force the address text hidden when switching *to* a table — for
    // every other case (range source, or no source at all) leave it as
    // showAddress()/restoreRange() already set it; unconditionally writing
    // `els.cells.hidden = isTable` here clobbered restoreRange's own
    // `hidden = true` back to visible whenever isTable was false.
    if (isTable) {
      els.tableBadge.textContent = `Table: ${state.source.name}`;
      els.cells.hidden = true;
    }
  }

  function sourceLabel(source) {
    const s = source || state.source;
    if (!s) return '';
    return s.kind === 'table' ? `Table: ${s.name}` : s.address;
  }

  // Default dashboard name, in order: the Excel Table's own name, else the
  // sheet name, else the workbook's file name, else a bare fallback — then
  // de-duplicated against titles already saved in this workbook (Excel
  // Tables are auto-named "Table1"/"Table2"/... so two dashboards from two
  // different tables on the same sheet would otherwise get an identical
  // suggested name). Only ever produces the *suggested* value — the caller
  // shows it in an editable field (addin/taskpane.js's view-theme title
  // input), so a collision the user doesn't fix is their own choice, not a
  // bug here.
  async function computeDefaultTitle() {
    let base;
    if (state.source && state.source.kind === 'table') {
      base = state.source.name;
    } else if (state.range) {
      const sheet = splitAddress(state.range.address).sheet;
      base = sheet || null;
    }
    if (!base) {
      try {
        base = await host.getWorkbookName();
      } catch (e) {
        base = null;
      }
    }
    if (!base) base = 'Dashboard';

    let existingTitles;
    try {
      existingTitles = new Set((await host.listDashboards()).map((d) => d.payload.title).filter(Boolean));
    } catch (e) {
      return base; // can't check for collisions right now — the field is editable anyway
    }
    if (!existingTitles.has(base)) return base;
    let n = 2;
    while (existingTitles.has(`${base} (${n})`)) n++;
    return `${base} (${n})`;
  }

  // Shared by findSavedMappingForSource below and the list-triggered
  // headless "Change data range" (which already has the exact payload in
  // hand and doesn't need to search for one by source label).
  function mappingToOverrides(mapping) {
    const overrides = {};
    for (const m of mapping) {
      overrides[m.name] = { role: m.role, chartEligible: m.chartEligible !== false };
    }
    return overrides;
  }

  // CLAUDE.md §7: a manual role mapping must survive a panel reload — it
  // was never actually wired up on the read side (payload.mapping was
  // write-only). Looked up fresh every time, not cached at startup, so a
  // dashboard saved in another window/session for this same source is
  // still found — matches purely by the saved sourceAddress label (exact
  // string), same value `placeOnSheet` writes, most recent one wins.
  async function findSavedMappingForSource(label) {
    if (!label) return null;
    let dashboards;
    try {
      dashboards = await host.listDashboards();
    } catch (e) {
      return null; // can't read right now — proceed with plain auto-classification, don't block on it
    }
    const candidates = dashboards.filter((d) => d.payload.sourceAddress === label && Array.isArray(d.payload.mapping));
    if (!candidates.length) return null;
    candidates.sort((a, b) => new Date(b.payload.generatedAt || 0) - new Date(a.payload.generatedAt || 0));
    return mappingToOverrides(candidates[0].payload.mapping);
  }

  function renderProgress() {
    document.querySelector('#view-setup .step').classList.toggle('done', !!state.range);
    els.summary.innerHTML = state.range
      ? `Ready to generate from <strong>${sourceLabel()}</strong>`
      : 'Choose data above';
    syncIdleFooter();
  }

  /* ---------- footer: idle sub-states (fresh / has-dashboard / picking a
     change-range source) ---------- */
  function syncIdleFooter() {
    const changeRangeMode = state.pickingFor === 'change-range' || state.pickingFor === 'list-change-range';
    const hasDashboard = !!state.analysis;
    els.generate.hidden = hasDashboard && !changeRangeMode;
    els.reopenRow.hidden = !hasDashboard || changeRangeMode;
    els.cancelChangeRange.hidden = !changeRangeMode;
    els.ctaLabel.textContent = changeRangeMode ? 'Use this range' : 'Generate dashboard';
    els.generate.disabled = !state.range;
  }

  // No-op: the temporary debug panel this fed is gone (it got in the way
  // day to day) — kept as a function so every call site below stays valid
  // without editing each one individually.
  function logDebug() {}

  /* ---------- mapping screen (>6 columns) ---------- */
  const ROLE_OPTIONS = [
    ['measure', 'Measure'],
    ['dimension', 'Category'],
    ['dimension_filter', 'Filter only'],
    ['time', 'Date'],
    ['text', 'Text (table only)'],
    ['excluded', "Don't use"],
  ];

  function roleOptionFor(decision) {
    if (decision.role === 'dimension') return decision.chartEligible === false ? 'dimension_filter' : 'dimension';
    return decision.role;
  }

  // Shows every column with a role <select> pre-filled from auto-
  // classification; `onConfirm(finalAnalysis, overrides)` fires once the
  // user reviews and clicks Generate. `overrides` covers every column (not
  // just changed ones) — engine/index.js#applyRoleOverrides no-ops for any
  // that already match what auto-classification decided.
  function showMappingScreen(analysis, onConfirm, onCancel) {
    showView('mapping');
    els.mappingList.innerHTML = '';
    const rows = analysis.columns.map((col) => {
      const row = document.createElement('div');
      row.className = 'mapping-row';
      const name = document.createElement('span');
      name.className = 'mapping-name';
      name.textContent = col.name;
      const select = document.createElement('select');
      select.className = 'mapping-select';
      for (const [value, label] of ROLE_OPTIONS) {
        const opt = document.createElement('option');
        opt.value = value;
        opt.textContent = label;
        select.appendChild(opt);
      }
      select.value = roleOptionFor(col.decision);
      row.append(name, select);
      els.mappingList.appendChild(row);
      return { name: col.name, select };
    });

    function onGenerateClick() {
      const overrides = {};
      for (const { name, select } of rows) {
        overrides[name] = select.value === 'dimension_filter'
          ? { role: 'dimension', chartEligible: false }
          : select.value === 'dimension'
            ? { role: 'dimension', chartEligible: true }
            : { role: select.value };
      }
      cleanup();
      onConfirm(window.DashEngine.applyRoleOverrides(analysis, overrides), overrides);
    }
    function onCancelClick() {
      cleanup();
      if (onCancel) onCancel();
      else showView('setup');
    }
    function cleanup() {
      els.mappingGenerate.removeEventListener('click', onGenerateClick);
      els.mappingCancel.removeEventListener('click', onCancelClick);
    }
    els.mappingGenerate.addEventListener('click', onGenerateClick);
    els.mappingCancel.addEventListener('click', onCancelClick);
  }

  /* ---------- generate ---------- */
  async function generate() {
    if (state.busy || !state.range || state.pickingFor === 'change-range') return;
    state.busy = true;
    document.body.classList.add('busy');
    els.buildError.hidden = true;
    els.buildTiming.hidden = true;
    els.ctaFill.style.transitionDuration = '400ms';

    try {
      setCta('Reading your data…', 33);
      const { headers, dataRows, timing, totalRows } = await host.readRangeForEngine(state.source);

      setCta('Classifying columns…', 66);
      const rawAnalysis = window.DashEngine.analyzeTable(headers, dataRows, ',');
      state.title = await computeDefaultTitle();

      if (timing) {
        els.buildTiming.hidden = false;
        els.buildTiming.textContent = `Read ${totalRows.toLocaleString('en-US')} rows in ${timing.totalMs}ms (values ${timing.valuesMs}ms, format sample ${timing.formatMs}ms)`;
      }

      // Restore a mapping saved for this exact source in an earlier session
      // (or another window) before deciding anything else — the mapping
      // screen below must show the restored roles, not bare auto-detection.
      if (state.roleOverrides == null) {
        state.roleOverrides = await findSavedMappingForSource(sourceLabel());
      }
      const effectiveAnalysis = window.DashEngine.applyRoleOverrides(rawAnalysis, state.roleOverrides);

      // Called either directly below (≤6 columns) or later from the mapping
      // screen's own Generate click — either way it just stages the
      // analysis and hands off to the theme/palette step; opening the
      // dialog itself (the part that can fail) happens in finalizeGenerate,
      // triggered by that step's own Generate button.
      const proceed = (analysis, overrides) => {
        state.roleOverrides = overrides !== undefined ? overrides : state.roleOverrides;
        state.analysis = analysis;
        state.lastDialogState = null;
        showThemeScreen();
      };

      if (effectiveAnalysis.columns.length > 6) {
        showMappingScreen(effectiveAnalysis, proceed);
      } else {
        proceed(effectiveAnalysis, state.roleOverrides);
      }
    } catch (err) {
      els.buildError.textContent = `Could not generate the dashboard: ${err && err.message ? err.message : String(err)}`;
      els.buildError.hidden = false;
    } finally {
      state.busy = false;
      document.body.classList.remove('busy');
      els.ctaFill.style.transitionDuration = '0ms';
      els.ctaFill.style.width = '0';
      syncIdleFooter();
    }
  }

  function setCta(label, pct) {
    els.ctaLabel.textContent = label;
    els.ctaFill.style.width = `${pct}%`;
  }

  /* ---------- theme + palette step (between mapping and opening the
     dialog) — single-click swatches + a live preview, see
     addin/theme-picker.js. Picked once here; changed later from the same
     picker inside the dialog's settings panel (addin/dashboard-dialog.js). ---------- */
  function showThemeScreen() {
    showView('theme');
    els.themeTitleInput.value = state.title || '';
    renderThemePickerView();
  }

  function renderThemePickerView() {
    window.DashThemePicker.render(
      els.themePicker,
      { THEMES: window.DashRenderThemes.THEMES, Palettes: window.DashRenderPalettes },
      { theme: state.theme, palette: state.palette },
      {
        onThemeChange: (id) => { state.theme = id; renderThemePickerView(); },
        onPaletteChange: (id) => { state.palette = id; renderThemePickerView(); },
      }
    );
  }

  // The actual dialog-opening step, deferred until the theme/palette step's
  // own Generate click — see the `proceed` comment in generate() above for
  // why this used to run inline there.
  async function finalizeGenerate() {
    if (state.busy) return;
    state.busy = true;
    els.buildError.hidden = true;
    state.title = (state.title || '').trim() || 'Dashboard';
    try {
      // Generate itself now creates a listed entry right away — no picture
      // on the sheet yet (see dashboard-io.js#saveGeneratedDraft), only the
      // mapping/title/theme choice already made, so it isn't lost if the
      // pane closes before an explicit "Place image on sheet" inside the
      // dialog. That stays a separate, deliberate action — Generate does
      // NOT place a picture on its own. A failure here doesn't block
      // opening the live dialog (Place still works from inside it either
      // way) but is surfaced rather than silently dropped — the whole
      // point of doing this is so the list is never the thing lying about
      // what happened.
      state.currentShapeName = window.DashAddinDashboardIo.makeShapeName(Date.now());
      const mapping = state.analysis.columns.map((c) => ({ name: c.name, role: c.decision.role, aggregation: c.decision.aggregation, chartEligible: c.decision.chartEligible !== false }));
      try {
        await host.saveDraft(state.currentShapeName, {
          mapping, sourceAddress: sourceLabel(), source: state.source, title: state.title, theme: state.theme, palette: state.palette,
        });
        await refreshDashboardLists();
      } catch (draftErr) {
        logDebug(`generate: draft save failed — ${draftErr && draftErr.message ? draftErr.message : draftErr}`);
        els.buildError.textContent = `Generated, but could not save it to your dashboard list yet: ${draftErr && draftErr.message ? draftErr.message : draftErr}. You can still place it on the sheet from the dialog.`;
        els.buildError.hidden = false;
      }
      await openLiveDashboard();
      showView('setup');
      setFooterState('open');
    } catch (err) {
      showView('setup');
      setFooterState('idle');
      els.buildError.textContent = `Could not generate the dashboard: ${err && err.message ? err.message : String(err)}`;
      els.buildError.hidden = false;
    } finally {
      state.busy = false;
    }
  }

  async function reopenDashboard() {
    if (state.busy || !state.analysis) return;
    state.busy = true;
    els.buildError.hidden = true;
    try {
      await openLiveDashboard();
      setFooterState('open');
    } catch (err) {
      els.buildError.textContent = `Could not open the dashboard: ${err && err.message ? err.message : String(err)}`;
      els.buildError.hidden = false;
    } finally {
      state.busy = false;
    }
  }

  function startNewDashboard() {
    state.analysis = null;
    state.roleOverrides = null;
    state.lastDialogState = null;
    state.result = null;
    state.currentShapeName = null;
    els.buildError.hidden = true;
    setFooterState('idle');
    syncIdleFooter();
  }

  /* ---------- dashboard dialog: opens the interactive dashboard in its own
     window (Office.context.ui.displayDialogAsync) instead of the cramped
     task pane. The dialog can't reach Excel itself, so all it gets is JSON
     over addin/dialog-messaging.js's chunked channel, and the things it
     needs Excel for — placing the image, refreshing, changing the source —
     it asks this pane to do. See SPEC.md for the full protocol writeup. ---------- */

  // TEMPORARY cache-busting query string — see addin/version.js. Without
  // this, Excel's WebView can serve a cached dashboard-dialog.html (and,
  // transitively, whatever it links to — see the same fix applied to every
  // <script>/<link> tag in addin/taskpane.html and addin/dashboard-dialog.html,
  // and to the two taskpane.html URLs in addin/manifest.xml) even after the
  // task pane itself reloaded fresh. Remove the `?v=...` here along with
  // the rest of version.js's plumbing before release.
  function dialogUrl() {
    const url = new URL('dashboard-dialog.html', location.href);
    if (window.DASH_BUILD_VERSION != null) url.searchParams.set('v', window.DASH_BUILD_VERSION);
    return url.href;
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
  // The dialog page's own buttons always go through this message round
  // trip too (never a direct local call) — otherwise a browser-only
  // shortcut would stop proving anything about the real path.
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

  // Opens a fresh dialog for state.analysis, wires every message kind the
  // dialog can send throughout its lifetime (one combined handler — see the
  // comment on why this isn't several separate addEventHandler calls), and
  // streams the initial hydration once the dialog signals it's ready.
  // `seedState` (state.lastDialogState, from a prior close) is optional —
  // omit it for a genuinely fresh dashboard.
  async function openLiveDashboard() {
    const dlg = await openDialog(dialogUrl());
    logDebug('dialog: opened, waiting for ready…');
    const Msg = window.DashDialogMessaging;

    let resolveReady;
    const readyPromise = new Promise((resolve) => { resolveReady = resolve; });
    const readyReceiver = Msg.createChunkReceiver(Msg.KIND.READY, () => resolveReady());
    const placeReceiver = Msg.createChunkReceiver(Msg.KIND.PLACE_ON_SHEET, (spec, id) => handlePlaceOnSheet(dlg, spec, id));
    const stateReceiver = Msg.createChunkReceiver(Msg.KIND.STATE_UPDATE, (raw) => {
      state.lastDialogState = raw;
      // Renamed via the pencil next to the dialog's title (render/dom.js) —
      // state.title is what placeOnSheet()/handleRefreshRequest actually
      // read, so this has to land here to survive a later Place/Refresh.
      if (raw && typeof raw.title === 'string' && raw.title.trim()) state.title = raw.title.trim();
    });
    const refreshReceiver = Msg.createChunkReceiver(Msg.KIND.REFRESH_REQUEST, (raw, id) => handleRefreshRequest(dlg, raw, id));
    const changeRangeReceiver = Msg.createChunkReceiver(Msg.KIND.CHANGE_RANGE_REQUEST, (raw, id) => handleChangeRangeRequest(dlg, raw, id));
    const openMappingReceiver = Msg.createChunkReceiver(Msg.KIND.OPEN_MAPPING_REQUEST, (raw, id) => handleOpenMappingRequest(dlg, raw, id));
    dlg.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
      readyReceiver(arg.message) || placeReceiver(arg.message) || stateReceiver(arg.message) || refreshReceiver(arg.message) || changeRangeReceiver(arg.message) || openMappingReceiver(arg.message);
    });

    await readyPromise;
    logDebug('dialog: ready, sending data.');
    sendDataToDialog(dlg, {
      mode: 'live',
      analysis: state.analysis,
      meta: { title: state.title, subtitle: `${state.analysis.rowCount.toLocaleString('en-US')} rows` },
      seedState: state.lastDialogState || null,
      source: sourceLabel(),
      // Only consulted by the dialog when there's no seedState to restore
      // theme/palette from (a genuinely fresh dashboard) — see the
      // pre-generate theme/palette step above and dashboard-dialog.js#mountLive.
      initialTheme: state.theme,
      initialPalette: state.palette,
    });
  }

  async function openRestoredDashboard(d) {
    state.currentShapeName = d.shapeName;
    await showList();
    try {
      const dlg = await openDialog(dialogUrl());
      logDebug('dialog: opened (restored), waiting for ready…');
      const Msg = window.DashDialogMessaging;

      let resolveReady;
      const readyPromise = new Promise((resolve) => { resolveReady = resolve; });
      const readyReceiver = Msg.createChunkReceiver(Msg.KIND.READY, () => resolveReady());
      dlg.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => readyReceiver(arg.message));

      await readyPromise;
      sendDataToDialog(dlg, { mode: 'frozen', layoutSpec: d.payload.layoutSpec, initialTheme: d.payload.theme, initialPalette: d.payload.palette });
    } catch (err) {
      logDebug(`openRestoredDashboard: could not open the dialog — ${err && err.message ? err.message : String(err)}`);
    }
  }

  // Runs the actual Excel work for "Place image on sheet" — takes the
  // already-resolved {canvas, widgets, theme} the dialog sends; the mount
  // controller that produced it lives in the dialog's own window, a
  // separate JS realm this pane has no access to.
  async function placeOnSheet(spec) {
    try {
      const baseTheme = window.DashRenderThemes.THEMES[spec.theme || 'light'];
      const theme = window.DashRenderPalettes.withPalette(baseTheme, spec.palette);
      const currentSpec = { canvas: spec.canvas, widgets: spec.widgets };
      const columnsByName = window.DashEngine.Aggregate.byName(state.analysis.columns);
      logDebug('place-on-sheet: rasterizing…');
      const pngBase64 = await window.DashRenderShare.renderPngBase64(currentSpec, theme, columnsByName, { scale: 2 });

      const snapshot = window.DashEngine.Layout.buildStorageSnapshot(currentSpec, state.analysis.columns);
      // chartEligible travels too — dropping it here is exactly how a
      // "Filter only" mapping choice used to go silent on the next restore.
      const mapping = state.analysis.columns.map((c) => ({ name: c.name, role: c.decision.role, aggregation: c.decision.aggregation, chartEligible: c.decision.chartEligible !== false }));

      // Two distinct Office.js writes happen inside host.placeDashboard
      // (addin/dashboard-io.js#generateAndPlace: place the image, then save
      // its settings) sharing one Excel.run — a failure in either one
      // rejects this whole call and lands in the catch below, never
      // silently. Logged as one step here (not two) because from this
      // side of the call they're not individually observable; if a report
      // ever again says "list is empty, no error shown," these debug lines
      // plus the try/catch below are the two facts to check first: did we
      // get this far, and did the catch actually fire.
      logDebug('place-on-sheet: rasterized, placing image + saving settings…');
      // Reuses the id minted at Generate time (or from reopening a saved
      // dashboard) if there is one — otherwise this dashboard was somehow
      // opened without ever going through finalizeGenerate/openUnplacedDraft,
      // and host.placeDashboard mints a fresh one exactly as it always did.
      const { shapeName } = await host.placeDashboard({
        pngBase64, canvasSize: currentSpec.canvas, snapshot, mapping, sourceAddress: sourceLabel(), title: state.title,
        theme: spec.theme, palette: spec.palette, shapeName: state.currentShapeName || undefined,
      });
      logDebug(`place-on-sheet: host.placeDashboard resolved — shape "${shapeName}".`);
      return { ok: true, shapeName };
    } catch (err) {
      logDebug(`place-on-sheet: threw — ${err && err.message ? err.message : String(err)}`);
      return { ok: false, error: err && err.message ? err.message : String(err) };
    }
  }

  async function handlePlaceOnSheet(dlg, spec, requestId) {
    logDebug(`place-on-sheet: request received (${requestId}).`);
    els.buildError.hidden = true; // clear any stale error from a previous attempt before this one runs
    const result = await placeOnSheet(spec);
    const Msg = window.DashDialogMessaging;
    Msg.sendChunked((str) => dlg.messageChild(str), Msg.KIND.PLACE_ON_SHEET_RESULT, result, null, requestId);
    if (result.ok) {
      state.result = { shapeName: result.shapeName };
      state.currentShapeName = result.shapeName;
      refreshStartupList();
      // Read the list back through the exact same path the UI uses,
      // immediately after a successful placement — if this doesn't show the
      // shape we just placed, that's the concrete fact to report back,
      // instead of "the list is empty" with no way to tell write-side from
      // read-side.
      host.listDashboards().then((list) => {
        const found = list.some((d) => d.shapeName === result.shapeName);
        logDebug(`place-on-sheet: post-save listDashboards() has ${list.length} entr${list.length === 1 ? 'y' : 'ies'}, ${found ? 'including' : 'NOT including'} "${result.shapeName}".`);
      }).catch((e) => logDebug(`place-on-sheet: post-save listDashboards() failed — ${e && e.message ? e.message : e}`));
      els.doneSub.textContent = state.excel
        ? `It's on the sheet as a picture named "${result.shapeName}". Anyone opening the file just sees a picture; with this add-in, reopening it — from "Dashboards in this workbook" or by clicking it — opens the interactive version again in a dialog, without re-reading the sheet.`
        : `Preview mode (no Excel host here): "${result.shapeName}" was saved in memory instead of workbook.settings. Open "Dashboards in this workbook" below to see the restored, no-recompute view.`;
      // workbook.settings (where the snapshot lives) only survives a
      // close/reopen once the *file* itself is saved — context.sync() alone
      // only commits to the live session. The picture looks placed either
      // way, so this is the one place that actually tells the user.
      els.doneSaveHint.hidden = !state.excel;
      setFooterState('done');
      logDebug(`place-on-sheet: placed as "${result.shapeName}".`);
    } else {
      // The dialog's own "Place image on sheet" button already shows this
      // (render/dom.js#renderPlaceButton), but that text is transient and
      // in a separate window — a failure here must also be visible in the
      // pane itself, not only in the temporary debug panel below.
      els.buildError.textContent = `Could not place the dashboard on the sheet: ${result.error}`;
      els.buildError.hidden = false;
      logDebug(`place-on-sheet: failed — ${result.error}`);
    }
  }

  // Re-reads state.source fresh and reapplies the saved role mapping
  // (state.roleOverrides) — never re-prompts the mapping screen; that's
  // reserved for a genuinely new source (Change data range). The dialog
  // reconciles/diffs on its side once it gets this.
  async function handleRefreshRequest(dlg, raw, requestId) {
    logDebug(`refresh: request received (${requestId}).`);
    const Msg = window.DashDialogMessaging;
    try {
      const { headers, dataRows } = await host.readRangeForEngine(state.source);
      const rawAnalysis = window.DashEngine.analyzeTable(headers, dataRows, ',');
      const analysis = window.DashEngine.applyRoleOverrides(rawAnalysis, state.roleOverrides);
      state.analysis = analysis;
      const meta = { title: state.title, subtitle: `${analysis.rowCount.toLocaleString('en-US')} rows` };
      Msg.sendChunked((str) => dlg.messageChild(str), Msg.KIND.REFRESH_RESULT, { ok: true, analysis, meta, seedState: raw, source: sourceLabel() }, null, requestId);
      logDebug('refresh: sent fresh data.');
    } catch (err) {
      Msg.sendChunked((str) => dlg.messageChild(str), Msg.KIND.REFRESH_RESULT, { ok: false, error: err && err.message ? err.message : String(err) }, null, requestId);
      logDebug(`refresh: failed — ${err && err.message ? err.message : err}`);
    }
  }

  // The empty-state widget (render/dom.js#renderEmptyState) asked to reopen
  // the mapping screen for the dashboard already showing in the dialog —
  // classification produced zero measures, almost always one column's role
  // guessed wrong rather than a real absence of numbers (see engine/roles.js's
  // identifier rule and CLAUDE.md §7). Reuses state.analysis — the currently
  // effective, override-applied analysis — so the screen shows exactly the
  // roles the dialog has right now, not bare auto-detection from scratch.
  function handleOpenMappingRequest(dlg, raw, requestId) {
    logDebug(`open-mapping: request received (${requestId}).`);
    const Msg = window.DashDialogMessaging;
    const respond = (payload) => Msg.sendChunked((str) => dlg.messageChild(str), Msg.KIND.OPEN_MAPPING_RESULT, payload, null, requestId);
    showMappingScreen(
      state.analysis,
      (analysis, overrides) => {
        state.roleOverrides = overrides !== undefined ? overrides : state.roleOverrides;
        state.analysis = analysis;
        const meta = { title: state.title, subtitle: `${analysis.rowCount.toLocaleString('en-US')} rows` };
        sendDataToDialog(dlg, { mode: 'live', analysis, meta, seedState: raw, source: sourceLabel() });
        respond({ ok: true });
        showView('setup');
        setFooterState('open');
        logDebug('open-mapping: applied new roles, sent fresh data.');
      },
      () => {
        respond({ ok: false, error: 'cancelled' });
        showView('setup');
        setFooterState('open');
        logDebug('open-mapping: cancelled.');
      }
    );
  }

  // The dialog asked to pick a different source. We can't do that
  // synchronously — hand control to the normal setup view and wait for the
  // user to finish picking and click "Use this range" (or cancel).
  function handleChangeRangeRequest(dlg, raw, requestId) {
    logDebug(`change-range: request received (${requestId}).`);
    state.pendingChangeRange = { dlg, raw, requestId };
    state.sourceBeforeChangeRange = state.source;
    state.rangeBeforeChangeRange = state.range;
    state.pickingFor = 'change-range';
    showView('setup');
    setFooterState('idle'); // #build-idle (holding the relabeled Generate/Cancel buttons) must be the visible footer section, not whatever it was mid-dialog (e.g. 'open')
  }

  function cancelChangeRangePicking() {
    const pending = state.pendingChangeRange;
    state.pickingFor = 'generate';
    state.pendingChangeRange = null;
    restoreRange(state.rangeBeforeChangeRange);
    state.source = state.sourceBeforeChangeRange;
    updateSourceBadge();
    setFooterState('open'); // dialog stays open and untouched — back to the normal "dashboard open" footer
    if (pending) {
      window.DashDialogMessaging.sendChunked((str) => pending.dlg.messageChild(str), window.DashDialogMessaging.KIND.CHANGE_RANGE_RESULT, { ok: false, error: 'cancelled' }, null, pending.requestId);
    }
  }

  // Structure match (same column names as the dashboard currently open in
  // the dialog) decides two things at once: whether the saved role mapping
  // carries over, and whether the dialog's filters/sort/settings can be
  // reconciled instead of reset — see engine/reconcile.js and SPEC.md.
  async function completeChangeRangePicking() {
    const pending = state.pendingChangeRange;
    state.pickingFor = 'generate';
    state.pendingChangeRange = null;
    syncIdleFooter();
    if (!pending) return;

    const Msg = window.DashDialogMessaging;
    const respond = (payload) => Msg.sendChunked((str) => pending.dlg.messageChild(str), Msg.KIND.CHANGE_RANGE_RESULT, payload, null, pending.requestId);

    try {
      const { headers, dataRows } = await host.readRangeForEngine(state.source);
      const rawAnalysis = window.DashEngine.analyzeTable(headers, dataRows, ',');
      const oldNames = state.analysis.columns.map((c) => c.name);
      const newNames = rawAnalysis.columns.map((c) => c.name);
      const matches = window.DashEngine.Reconcile.columnsStructureMatches(oldNames, newNames);
      // Same report, different period -> keep carrying today's session
      // mapping. Genuinely different source -> don't blindly wipe it, this
      // new source may have its own saved mapping from an earlier session.
      const overridesToCarry = matches ? state.roleOverrides : await findSavedMappingForSource(sourceLabel());
      if (!matches) state.roleOverrides = overridesToCarry;

      const finish = async (finalAnalysis, overrides) => {
        state.analysis = finalAnalysis;
        state.roleOverrides = overrides !== undefined ? overrides : overridesToCarry;
        // Keep whatever name this dashboard already has (the user may have
        // typed their own) — a changed data range doesn't need a changed
        // title. Only compute a fresh default if it somehow has none yet.
        if (!state.title) state.title = await computeDefaultTitle();
        const meta = { title: state.title, subtitle: `${finalAnalysis.rowCount.toLocaleString('en-US')} rows` };
        respond({ ok: true, analysis: finalAnalysis, meta, seedState: matches ? pending.raw : null, structureReset: !matches, source: sourceLabel() });
        showView('setup');
        setFooterState('open'); // the dialog is still open with the (now updated) dashboard
      };

      const withOverrides = window.DashEngine.applyRoleOverrides(rawAnalysis, overridesToCarry);
      if (!matches && withOverrides.columns.length > 6) {
        showMappingScreen(withOverrides, finish, () => {
          // the dialog is still waiting on this request — a plain
          // showView('setup') here would strand it forever
          respond({ ok: false, error: 'cancelled' });
          showView('setup');
          setFooterState('open');
        });
      } else {
        finish(withOverrides, overridesToCarry);
      }
    } catch (err) {
      logDebug(`change-range: failed — ${err && err.message ? err.message : err}`);
      respond({ ok: false, error: err && err.message ? err.message : String(err) });
      setFooterState('open'); // the dialog is untouched and still open — its own banner shows the failure, no need to strand the pane on the idle/setup footer
    }
  }

  /* ---------- dashboards list / restore ---------- */
  async function showList() {
    showView('list');
    await refreshDashList();
  }

  // Split out from showList() so a list-triggered rename/change-range can
  // refresh the DOM in place after it finishes — those run from either list
  // (startup or hamburger-menu) and must not force-navigate the pane to
  // 'list' if the user was looking at something else when they clicked.
  async function refreshDashList() {
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
      // Empty settings could mean "nothing generated here yet" — or it
      // could mean a dashboard picture is sitting right there on a sheet
      // and its settings just never survived a close without saving (see
      // dashboard-io.js#countOrphanedDashboardShapes). Best-effort only:
      // a failure here just falls back to the plain empty message rather
      // than blocking the view on a second read.
      let orphanCount = 0;
      try {
        orphanCount = await host.countOrphanedDashboardShapes();
      } catch (e) { /* fall back to the plain empty message below */ }
      els.listEmpty.hidden = false;
      els.listEmpty.textContent = orphanCount > 0
        ? `There ${orphanCount === 1 ? 'is' : 'are'} ${orphanCount} dashboard picture${orphanCount === 1 ? '' : 's'} on the sheets, but no saved settings to match — the workbook was likely closed without saving after ${orphanCount === 1 ? 'it was' : 'they were'} placed. ${orphanCount === 1 ? 'It is' : 'They are'} still there; generate the dashboard again to get the interactive version back.`
        : 'No dashboards yet. Select a range and generate one.';
      return;
    }
    for (const d of dashboards) {
      els.dashList.appendChild(renderDashListItem(d));
    }
  }

  // After a list-triggered rename/change-range completes, both lists need
  // fresh data (the item could be showing in either or both), but neither
  // view should be force-switched — the user might be looking at either
  // one, or neither, when the background update finishes.
  async function refreshDashboardLists() {
    await refreshStartupList();
    if (!els.viewList.hidden) await refreshDashList();
  }

  function renderDashListItem(d) {
    const isDraft = d.payload.placed === false;
    const li = document.createElement('li');
    const title = document.createElement('p');
    title.className = 'dash-item-title';
    title.textContent = d.payload.title || d.shapeName;
    const meta = document.createElement('p');
    meta.className = 'dash-item-meta';
    const when = d.payload.generatedAt ? new Date(d.payload.generatedAt).toLocaleString() : 'unknown time';
    // A draft has no picture on the sheet yet — say so plainly rather than
    // showing a generated-time line that implies it's already there.
    meta.textContent = isDraft
      ? `${d.payload.sourceAddress || ''} · not placed on a sheet yet`
      : `${d.payload.sourceAddress || ''} · generated ${when}`;
    const actions = document.createElement('div');
    actions.className = 'dash-item-actions';
    const openBtn = document.createElement('button');
    openBtn.type = 'button';
    openBtn.className = 'primary';
    openBtn.textContent = 'Open';
    openBtn.addEventListener('click', () => (isDraft ? openUnplacedDraft(d) : openRestoredDashboard(d)));
    const renameBtn = document.createElement('button');
    renameBtn.type = 'button';
    renameBtn.textContent = 'Rename';
    renameBtn.addEventListener('click', () => startListRename(d, title));
    actions.append(openBtn, renameBtn);
    // Change data range needs a placed picture to update in place — for a
    // draft, opening it and using the dialog's own Change data range does
    // the same thing (there's nothing on the sheet yet to leave stale).
    if (!isDraft) {
      const rangeBtn = document.createElement('button');
      rangeBtn.type = 'button';
      rangeBtn.textContent = 'Change data range';
      rangeBtn.addEventListener('click', () => startListChangeRange(d));
      actions.appendChild(rangeBtn);
    }
    const deleteBtn = document.createElement('button');
    deleteBtn.type = 'button';
    deleteBtn.className = 'dash-item-delete';
    deleteBtn.title = 'Delete dashboard';
    deleteBtn.setAttribute('aria-label', 'Delete dashboard');
    deleteBtn.textContent = '×';
    deleteBtn.addEventListener('click', () => showListDeleteConfirm(li, d));
    actions.appendChild(deleteBtn);
    li.append(title, meta, actions);
    return li;
  }

  // Inline confirm instead of window.confirm() — consistent with the rest
  // of this UI's own controls rather than a native browser dialog.
  function showListDeleteConfirm(li, d) {
    if (li.querySelector('.dash-item-confirm')) return;
    if (state.currentShapeName === d.shapeName) {
      els.buildError.textContent = `"${d.payload.title || d.shapeName}" is currently open — close it first, then delete it from the list.`;
      els.buildError.hidden = false;
      return;
    }
    const row = document.createElement('div');
    row.className = 'dash-item-confirm';
    const msg = document.createElement('span');
    msg.textContent = `Delete "${d.payload.title || d.shapeName}"?`;
    const yesBtn = document.createElement('button');
    yesBtn.type = 'button';
    yesBtn.className = 'danger';
    yesBtn.textContent = 'Delete';
    yesBtn.addEventListener('click', () => commitListDelete(d));
    const noBtn = document.createElement('button');
    noBtn.type = 'button';
    noBtn.textContent = 'Cancel';
    noBtn.addEventListener('click', () => row.remove());
    row.append(msg, yesBtn, noBtn);
    li.appendChild(row);
  }

  async function commitListDelete(d) {
    if (state.busy) return;
    state.busy = true;
    els.buildError.hidden = true;
    try {
      await host.deleteDashboard(d.shapeName);
      await refreshDashboardLists();
    } catch (err) {
      logDebug(`list-delete: failed — ${err && err.message ? err.message : err}`);
      els.buildError.textContent = `Could not delete "${d.payload.title || d.shapeName}": ${err && err.message ? err.message : err}`;
      els.buildError.hidden = false;
    } finally {
      state.busy = false;
    }
  }

  // Reopens a Generate-but-not-yet-placed dashboard (see
  // dashboard-io.js#saveGeneratedDraft) — no frozen picture to restore
  // from, so this re-reads the saved source live and continues the exact
  // same in-progress session finalizeGenerate() would have started,
  // reusing the same shapeName so a later "Place image on sheet" updates
  // this same list entry instead of creating a second one.
  async function openUnplacedDraft(d) {
    if (state.busy || state.picking) return;
    state.busy = true;
    els.buildError.hidden = true;
    try {
      const source = d.payload.source;
      if (!source) throw new Error('This draft has no saved source to re-read.');
      const { headers, dataRows } = await host.readRangeForEngine(source);
      const rawAnalysis = window.DashEngine.analyzeTable(headers, dataRows, ',');
      const overrides = Array.isArray(d.payload.mapping) ? mappingToOverrides(d.payload.mapping) : null;
      state.source = source;
      state.analysis = window.DashEngine.applyRoleOverrides(rawAnalysis, overrides);
      state.roleOverrides = overrides;
      state.title = d.payload.title;
      state.theme = d.payload.theme || 'light';
      state.palette = d.payload.palette || 'ocean';
      state.currentShapeName = d.shapeName;
      state.lastDialogState = null;
      updateSourceBadge();
      await openLiveDashboard();
      showView('setup');
      setFooterState('open');
    } catch (err) {
      logDebug(`open-draft: failed — ${err && err.message ? err.message : err}`);
      els.buildError.textContent = `Could not reopen "${d.payload.title || d.shapeName}": ${err && err.message ? err.message : err}`;
      els.buildError.hidden = false;
    } finally {
      state.busy = false;
    }
  }

  // ---------- list-triggered rename/change-range: both act on a dashboard
  // that is NOT open in any dialog — headless, no live preview, patches the
  // already-saved settings and the on-sheet picture directly (see
  // addin/dashboard-io.js#regenerateAndReplace). Distinct from the dialog's
  // own rename (render/dom.js's pencil)/Change data range (dialog settings
  // panel), which both go through a live session instead. ----------

  function resolveDashboardTheme(payload) {
    const base = (window.DashRenderThemes.THEMES[payload.theme] || window.DashRenderThemes.THEMES.light);
    return window.DashRenderPalettes.withPalette(base, payload.palette || window.DashRenderPalettes.DEFAULT_PALETTE);
  }

  function patchLayoutSpecTitle(layoutSpec, newTitle) {
    return Object.assign({}, layoutSpec, {
      widgets: layoutSpec.widgets.map((w) => (w.type === 'header' ? Object.assign({}, w, { title: newTitle }) : w)),
    });
  }

  function startListRename(d, titleEl) {
    if (state.busy || state.picking) return;
    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'title-input';
    input.value = d.payload.title || '';
    let settled = false;
    function finish(shouldCommit) {
      if (settled) return;
      settled = true;
      input.remove();
      titleEl.hidden = false;
      if (shouldCommit) commitListRename(d, input.value);
    }
    input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') { e.preventDefault(); finish(true); }
      else if (e.key === 'Escape') { e.preventDefault(); finish(false); }
    });
    input.addEventListener('blur', () => finish(true));
    titleEl.hidden = true;
    titleEl.insertAdjacentElement('afterend', input);
    input.focus();
    input.select();
  }

  async function commitListRename(d, newTitleRaw) {
    const newTitle = (newTitleRaw || '').trim();
    if (!newTitle || newTitle === d.payload.title) return;
    state.busy = true;
    els.buildError.hidden = true;
    try {
      if (d.payload.placed === false) {
        // No picture on the sheet yet — nothing to rasterize/replace.
        await host.renameDraft(d.shapeName, newTitle);
      } else {
        const layoutSpec = patchLayoutSpecTitle(d.payload.layoutSpec, newTitle);
        const theme = resolveDashboardTheme(d.payload);
        const pngBase64 = await window.DashRenderShare.renderPngBase64(layoutSpec, theme, null);
        await host.regenerateDashboard(d.shapeName, {
          pngBase64, canvasSize: layoutSpec.canvas, snapshot: layoutSpec, mapping: d.payload.mapping,
          sourceAddress: d.payload.sourceAddress, title: newTitle, theme: d.payload.theme, palette: d.payload.palette,
        });
      }
      // This dashboard might also be the one currently open in a live
      // dialog (or mid-creation in this very pane) — keep that in sync too,
      // or a later Place/Refresh would resave the OLD title over this rename.
      if (state.currentShapeName === d.shapeName) state.title = newTitle;
      await refreshDashboardLists();
    } catch (err) {
      logDebug(`list-rename: failed — ${err && err.message ? err.message : err}`);
      els.buildError.textContent = `Could not rename "${d.payload.title || d.shapeName}": ${err && err.message ? err.message : err}`;
      els.buildError.hidden = false;
    } finally {
      state.busy = false;
    }
  }

  function startListChangeRange(d) {
    if (state.busy || state.picking) return;
    // This exact dashboard is the one currently open live (in a dialog, or
    // mid-creation in this pane) — that live session has its own source in
    // memory already, and this headless path writing over it underneath
    // would let a later Refresh/Place from the stale live session clobber
    // the change right back. Use the dialog's own Change data range there
    // instead — see addin/dashboard-dialog.js's Data source setting.
    if (state.currentShapeName === d.shapeName) {
      els.buildError.textContent = `"${d.payload.title || d.shapeName}" is currently open — use "Change data range" inside its own window instead.`;
      els.buildError.hidden = false;
      return;
    }
    state.pendingListChangeRange = { dashboard: d, rangeBeforePick: state.range, sourceBeforePick: state.source };
    state.pickingFor = 'list-change-range';
    showView('setup');
    setFooterState('idle');
  }

  function cancelListChangeRangePicking() {
    const pending = state.pendingListChangeRange;
    state.pickingFor = 'generate';
    state.pendingListChangeRange = null;
    if (pending) {
      restoreRange(pending.rangeBeforePick);
      state.source = pending.sourceBeforePick;
      updateSourceBadge();
    }
    showView('list');
  }

  // No mapping-review screen here even for a >6-column result — this flow
  // never opens anything for the user to review before it writes, by
  // design (a fully headless "Change data range" was the explicit choice
  // over "open the dialog and jump into its own Change data range flow").
  // A saved role mapping from the dashboard's PREVIOUS source still applies
  // wherever a column name matches; anything new just falls back to plain
  // auto-classification, same as applyRoleOverrides always does for a name
  // it doesn't recognize.
  async function completeListChangeRangePicking() {
    const pending = state.pendingListChangeRange;
    state.pickingFor = 'generate';
    state.pendingListChangeRange = null;
    if (!pending) return;

    const newSource = state.source;
    const newSourceLabel = sourceLabel(newSource);
    restoreRange(pending.rangeBeforePick);
    state.source = pending.sourceBeforePick;
    updateSourceBadge();
    showView('list');
    els.buildError.hidden = true;

    state.busy = true;
    try {
      const { headers, dataRows } = await host.readRangeForEngine(newSource);
      const rawAnalysis = window.DashEngine.analyzeTable(headers, dataRows, ',');
      const oldMapping = pending.dashboard.payload.mapping;
      const overrides = Array.isArray(oldMapping) ? mappingToOverrides(oldMapping) : null;
      const analysis = window.DashEngine.applyRoleOverrides(rawAnalysis, overrides);
      const meta = { title: pending.dashboard.payload.title, subtitle: `${analysis.rowCount.toLocaleString('en-US')} rows` };
      const layoutSpec = window.DashEngine.buildLayoutSpec(analysis, meta);
      const theme = resolveDashboardTheme(pending.dashboard.payload);
      const columnsByName = window.DashEngine.Aggregate.byName(analysis.columns);
      const pngBase64 = await window.DashRenderShare.renderPngBase64(layoutSpec, theme, columnsByName);
      const snapshot = window.DashEngine.Layout.buildStorageSnapshot(layoutSpec, analysis.columns);
      const mapping = analysis.columns.map((c) => ({ name: c.name, role: c.decision.role, aggregation: c.decision.aggregation, chartEligible: c.decision.chartEligible !== false }));
      await host.regenerateDashboard(pending.dashboard.shapeName, {
        pngBase64, canvasSize: layoutSpec.canvas, snapshot, mapping, sourceAddress: newSourceLabel,
        title: pending.dashboard.payload.title, theme: pending.dashboard.payload.theme, palette: pending.dashboard.payload.palette,
      });
      await refreshDashboardLists();
    } catch (err) {
      logDebug(`list-change-range: failed — ${err && err.message ? err.message : err}`);
      els.buildError.textContent = `Could not change the data range for "${pending.dashboard.payload.title || pending.dashboard.shapeName}": ${err && err.message ? err.message : err}`;
      els.buildError.hidden = false;
    } finally {
      state.busy = false;
    }
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
    els.viewMapping.hidden = name !== 'mapping';
    els.viewTheme.hidden = name !== 'theme';
    els.build.hidden = name === 'list' || name === 'mapping' || name === 'theme';
  }

  function setFooterState(s) {
    els.buildIdle.hidden = s !== 'idle';
    els.buildOpen.hidden = s !== 'open';
    els.buildDone.hidden = s !== 'done';
    syncIdleFooter();
  }

  function startOver() {
    state.picking = false;
    state.pickingFor = 'generate';
    state.range = null;
    state.source = null;
    state.beforePick = null;
    state.analysis = null;
    state.roleOverrides = null;
    state.lastDialogState = null;
    state.result = null;
    state.currentShapeName = null;
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
    refreshStartupList();
  }
})();
