/*
 * The dashboard dialog — a separate window (Office.context.ui.displayDialogAsync)
 * that hosts the live, fully-interactive dashboard so it isn't squeezed into
 * the narrow task pane. It never touches Excel directly (dialogs can't reach
 * the Office document object model): everything it needs arrives from the
 * task pane as JSON over dialog-messaging.js's chunked protocol, and the
 * things it can't do itself — placing the image, re-reading the sheet for
 * Refresh, picking a new source for Change data range — it asks the task
 * pane to do, over the same channel.
 *
 * Two hydration modes, chosen by the task pane at data-send time:
 *   - 'live'   -> Dom.mount(): full interactivity (filters, chart clicks,
 *                 sort, pagination) plus Refresh / Change data range / a
 *                 settings panel (Theme, Palette, KPIs, Charts, Sections) /
 *                 "Place image on sheet".
 *   - 'frozen' -> Dom.mountFrozen(): restoring a previously placed
 *                 dashboard that has no source rows left — same window,
 *                 same theme switcher, filters read-only (mountFrozen's own
 *                 contract), none of the live-only toolbar actions (nothing
 *                 to refresh, no source to change, no image left to place).
 *
 * Every state-changing action in live mode (filter click, sort, theme,
 * settings panel) pushes a STATE_UPDATE to the task pane so it can restore
 * the same state next time this dashboard is (re)opened — see
 * addin/dialog-messaging.js and CLAUDE.md-adjacent SPEC.md notes on this.
 */
(() => {
  'use strict';

  const Msg = window.DashDialogMessaging;
  const Engine = window.DashEngine;
  const THEMES = window.DashRenderThemes.THEMES;
  const Palettes = window.DashRenderPalettes;

  const els = {
    themes: document.getElementById('dlg-themes'),
    status: document.getElementById('dlg-status'),
    banner: document.getElementById('dlg-banner'),
    mount: document.getElementById('dlg-mount'),
    liveActions: document.getElementById('dlg-live-actions'),
    source: document.getElementById('dlg-source'),
    refreshedAt: document.getElementById('dlg-refreshed-at'),
    refreshBtn: document.getElementById('dlg-refresh'),
    changeRangeBtn: document.getElementById('dlg-change-range'),
    settingsToggle: document.getElementById('dlg-settings-toggle'),
    settingsPanel: document.getElementById('dlg-settings'),
    settingsClose: document.getElementById('dlg-settings-close'),
    settingsThemePicker: document.getElementById('dlg-theme-picker'),
    settingsFilters: document.getElementById('dlg-settings-filters'),
    settingsKpis: document.getElementById('dlg-settings-kpis'),
    settingsCharts: document.getElementById('dlg-settings-charts'),
    settingsTableColumns: document.getElementById('dlg-settings-table-columns'),
    cfgTable: document.getElementById('dlg-cfg-table'),
    cfgFilters: document.getElementById('dlg-cfg-filters'),
    versionBadge: document.getElementById('dlg-version-badge'),
  };

  // TEMPORARY — see addin/version.js; remove with it before release.
  if (window.DASH_BUILD_VERSION != null) els.versionBadge.textContent = `v${window.DASH_BUILD_VERSION}`;

  let controller = null;
  let mode = null; // 'live' | 'frozen'
  let theme = 'light';
  let palette = Palettes.DEFAULT_PALETTE;
  let widgetConfig = { enabledKpis: null, enabledCharts: null, chartOverrides: null, showCharts: true, showTable: true, showFilters: true };
  let currentAnalysis = null; // the analysis behind the live-mounted dashboard, if any — used to diff valueTypes and to reconcile a Refresh/Change-range result
  let currentMeta = null;
  let currentSourceLabel = '';
  let canvasSize = null; // {width,height} — captured at mount time; mountFrozen's controller has no getLayoutSpec() to ask later
  let parentSend = null; // (rawString) => void, set once the transport is ready

  function setStatus(text, isError) {
    els.status.hidden = !text;
    els.status.textContent = text || '';
    els.status.dataset.error = isError ? '1' : '0';
  }

  function showBanner(text, isError) {
    els.banner.hidden = !text;
    els.banner.textContent = text || '';
    els.banner.dataset.error = isError ? '1' : '0';
  }

  // ---- transport: real Office dialog, or a same-origin window.opener +
  // postMessage stand-in for testing outside Excel (addin/taskpane.js's
  // previewHost opens this same page via window.open, not displayDialogAsync
  // — see its openDialogPreview). Only the transport differs; every message
  // above this layer is identical either way. ----
  function initTransport(onRawMessage, onReady) {
    if (window.Office && Office.context && Office.context.ui && typeof Office.context.ui.addHandlerAsync === 'function') {
      Office.context.ui.addHandlerAsync(
        Office.EventType.DialogParentMessageReceived,
        (arg) => onRawMessage(arg.message),
        (result) => {
          parentSend = (str) => Office.context.ui.messageParent(str);
          if (result.status === Office.AsyncResultStatus.Failed) setStatus(`Could not connect to the task pane: ${result.error.message}`, true);
          else onReady();
        }
      );
      return;
    }
    if (window.opener) {
      window.addEventListener('message', (e) => {
        if (e.source !== window.opener) return;
        if (!e.data || e.data.__dashDialog !== true) return;
        onRawMessage(e.data.message);
      });
      parentSend = (str) => window.opener.postMessage({ __dashDialog: true, message: str }, location.origin);
      onReady();
      return;
    }
    setStatus('No parent window available — this page must be opened as a dialog from the task pane.', true);
  }

  // theme (a THEMES key) merged with the currently chosen series palette —
  // every mount/setTheme call in this file goes through this, never THEMES[theme]
  // directly, so a palette pick can never silently apply to only some of them.
  function activeTheme() {
    return Palettes.withPalette(THEMES[theme], palette);
  }

  // ---- current state, as sent to the task pane (STATE_UPDATE) or included
  // in a request the task pane will echo back once it has fresh data ----
  function currentRawState() {
    const s = controller && mode === 'live' ? controller.getState() : {};
    return {
      activeFilters: Object.fromEntries(Object.entries(s.activeFilters || {}).map(([k, v]) => [k, [...v]])),
      sort: s.sort || null,
      theme,
      palette,
      widgetConfig,
      title: currentMeta && currentMeta.title,
    };
  }

  function pushState() {
    if (!parentSend || mode !== 'live') return;
    Msg.sendChunked(parentSend, Msg.KIND.STATE_UPDATE, currentRawState());
  }

  // ---- hydration: receive {mode, analysis|layoutSpec, meta, seedState?, source?} and mount ----
  const dataReceiver = Msg.createChunkReceiver(Msg.KIND.DATA, (data) => {
    try {
      mountFromData(data);
    } catch (err) {
      setStatus(`Could not render the dashboard: ${err && err.message ? err.message : String(err)}`, true);
    }
  });

  function mountFromData(data) {
    setStatus('');
    mode = data.mode;
    if (data.mode === 'frozen') {
      els.liveActions.hidden = true;
      // Live mode hides this row (see below) — the same theme switch lives
      // in the settings panel there. Frozen mode has no settings panel at
      // all (nothing to refresh/change-range/configure), so this stays the
      // only way to switch its theme.
      els.themes.hidden = false;
      currentAnalysis = null;
      // A restored dashboard has no live widgetConfig to derive paint from
      // — composition is already baked into layoutSpec.widgets, but color
      // isn't, so the theme/palette it was placed with (addin/dashboard-io.js's
      // saved payload) has to be applied explicitly here.
      if (data.initialTheme) theme = data.initialTheme;
      if (data.initialPalette) palette = data.initialPalette;
      syncThemeButtons();
      canvasSize = data.layoutSpec.canvas;
      controller = window.DashRenderDom.mountFrozen(els.mount, data.layoutSpec, activeTheme());
      applyScale();
      return;
    }

    els.liveActions.hidden = false;
    // Redundant with the settings panel's own Theme & palette section
    // (which also carries the palette, unlike this quick row) — see the
    // frozen branch above for why frozen mode keeps it instead.
    els.themes.hidden = true;
    if (data.source) { currentSourceLabel = data.source; els.source.textContent = data.source; }

    const seedReconciled = data.seedState ? Engine.Reconcile.reconcileDashboardState(data.analysis, data.seedState) : null;
    if (currentAnalysis) {
      const changed = Engine.Reconcile.diffValueTypes(currentAnalysis, data.analysis);
      if (changed.length) showBanner(`Column type changed after refresh: ${changed.join(', ')}. Data was recalculated, your filters and settings were kept.`, false);
    }

    currentAnalysis = data.analysis;
    currentMeta = data.meta;
    mountLive(seedReconciled, data.initialTheme, data.initialPalette);
  }

  // `initialTheme`/`initialPalette` (from addin/taskpane.js's pre-generate
  // theme/palette step) only matter when there's no seedState to restore
  // from — a genuinely fresh dashboard. A Refresh/Change-range always has
  // seedReconciled (carrying forward whatever this dialog session already
  // had), which wins.
  function mountLive(seedReconciled, initialTheme, initialPalette) {
    if (seedReconciled) {
      theme = seedReconciled.theme || theme;
      palette = seedReconciled.palette || palette;
      widgetConfig = seedReconciled.widgetConfig || widgetConfig;
    } else {
      if (initialTheme) theme = initialTheme;
      if (initialPalette) palette = initialPalette;
    }
    syncThemeButtons();
    syncSettingsPanel();

    els.mount.innerHTML = '';
    const initialState = seedReconciled ? { activeFilters: seedReconciled.activeFilters, sort: seedReconciled.sort } : undefined;
    const metaWithConfig = Object.assign({}, currentMeta, { widgetConfig });
    controller = window.DashRenderDom.mount(els.mount, currentAnalysis, activeTheme(), metaWithConfig, {
      onPlaceOnSheet,
      onOpenMapping,
      initialState,
      onStateChange: pushState,
      onTitleChange: onRenameTitle,
    });
    canvasSize = controller.getLayoutSpec().canvas;
    syncCanvasWidthToStage();
  }

  // The pencil next to the title (render/dom.js#renderHeader) commits here.
  // Nothing is written into the workbook by this alone — same rule as every
  // other live-mode setting (theme, KPIs, chart picks): it only becomes part
  // of the saved dashboard the next time "Place image on sheet" runs, via
  // currentRawState()'s `title` feeding the task pane's own state.title.
  function onRenameTitle(newTitle) {
    currentMeta = Object.assign({}, currentMeta, { title: newTitle });
    pushState();
  }

  // ---- place-on-sheet round trip ----
  const pendingPlacements = new Map();
  const placeResultReceiver = Msg.createChunkReceiver(Msg.KIND.PLACE_ON_SHEET_RESULT, (result, requestId) => {
    const pending = pendingPlacements.get(requestId);
    if (!pending) return;
    pendingPlacements.delete(requestId);
    pending.resolve(result);
  });

  // `spec` ({canvas, widgets}) already reflects the current filters/sort
  // (see render/dom.js's getSpecForShare) but not theme, which this module
  // tracks separately (theme is display-only, not part of layoutSpec — see
  // engine/layout.js) — the task pane needs it to rasterize with the same
  // paint the user is looking at.
  function onPlaceOnSheet(spec) {
    return new Promise((resolve) => {
      const requestId = Msg.sendChunked(parentSend, Msg.KIND.PLACE_ON_SHEET, Object.assign({}, spec, { theme, palette }));
      pendingPlacements.set(requestId, { resolve });
    });
  }

  // ---- open mapping (from the empty-state widget, render/dom.js#renderEmptyState) ----
  const pendingOpenMapping = new Map();
  const openMappingResultReceiver = Msg.createChunkReceiver(Msg.KIND.OPEN_MAPPING_RESULT, (result, requestId) => {
    const pending = pendingOpenMapping.get(requestId);
    if (!pending) return;
    pendingOpenMapping.delete(requestId);
    pending.resolve(result);
  });

  function onOpenMapping() {
    return new Promise((resolve) => {
      const requestId = Msg.sendChunked(parentSend, Msg.KIND.OPEN_MAPPING_REQUEST, currentRawState());
      pendingOpenMapping.set(requestId, { resolve });
    });
  }

  // ---- refresh ----
  const pendingRefresh = new Map();
  const refreshResultReceiver = Msg.createChunkReceiver(Msg.KIND.REFRESH_RESULT, (result, requestId) => {
    const pending = pendingRefresh.get(requestId);
    if (!pending) return;
    pendingRefresh.delete(requestId);
    pending.resolve(result);
  });

  els.refreshBtn.addEventListener('click', async () => {
    if (els.refreshBtn.dataset.busy) return;
    els.refreshBtn.dataset.busy = '1';
    const original = els.refreshBtn.textContent;
    els.refreshBtn.textContent = 'Refreshing…';
    try {
      const requestId = Msg.sendChunked(parentSend, Msg.KIND.REFRESH_REQUEST, currentRawState());
      const result = await new Promise((resolve) => pendingRefresh.set(requestId, { resolve }));
      if (result.ok) {
        mountFromData({ mode: 'live', analysis: result.analysis, meta: result.meta, seedState: result.seedState, source: result.source });
        els.refreshBtn.textContent = 'Refreshed';
        els.refreshedAt.textContent = `Updated ${new Date().toLocaleTimeString()}`;
      } else {
        els.refreshBtn.textContent = `Could not refresh`;
        showBanner(`Refresh failed: ${result.error}`, true);
      }
    } catch (err) {
      els.refreshBtn.textContent = 'Could not refresh';
      showBanner(`Refresh failed: ${err && err.message ? err.message : String(err)}`, true);
    }
    setTimeout(() => { els.refreshBtn.textContent = original; delete els.refreshBtn.dataset.busy; }, 2200);
  });

  // ---- change data range ----
  const pendingChangeRange = new Map();
  const changeRangeResultReceiver = Msg.createChunkReceiver(Msg.KIND.CHANGE_RANGE_RESULT, (result, requestId) => {
    const pending = pendingChangeRange.get(requestId);
    if (!pending) return;
    pendingChangeRange.delete(requestId);
    pending.resolve(result);
  });

  els.changeRangeBtn.addEventListener('click', async () => {
    if (els.changeRangeBtn.dataset.busy) return;
    els.changeRangeBtn.dataset.busy = '1';
    const original = els.changeRangeBtn.textContent;
    els.changeRangeBtn.textContent = 'Waiting…';
    setStatus('Pick a new range or table in the task pane, then confirm it there.');
    try {
      const requestId = Msg.sendChunked(parentSend, Msg.KIND.CHANGE_RANGE_REQUEST, currentRawState());
      const result = await new Promise((resolve) => pendingChangeRange.set(requestId, { resolve }));
      setStatus('');
      if (result.ok) {
        if (result.structureReset) showBanner('The new source has different columns — filters and settings were reset.', false);
        mountFromData({ mode: 'live', analysis: result.analysis, meta: result.meta, seedState: result.seedState, source: result.source });
      } else {
        showBanner(`Could not change the data source: ${result.error}`, true);
      }
    } catch (err) {
      setStatus('');
      showBanner(`Could not change the data source: ${err && err.message ? err.message : String(err)}`, true);
    }
    els.changeRangeBtn.textContent = original;
    delete els.changeRangeBtn.dataset.busy;
  });

  // ---- theme ----
  function syncThemeButtons() {
    [...els.themes.querySelectorAll('button')].forEach((b) => b.setAttribute('aria-pressed', String(b.dataset.theme === theme)));
  }
  els.themes.addEventListener('click', (e) => {
    const btn = e.target.closest('button[data-theme]');
    if (!btn || !controller) return;
    theme = btn.dataset.theme;
    syncThemeButtons();
    controller.setTheme(activeTheme()); // re-paints in place — activeFilters/sort/page untouched, see render/dom.js's setTheme
    if (!els.settingsPanel.hidden) renderDialogThemePicker(); // keep the settings-panel picker in sync if it's open too
    pushState();
  });

  // ---- settings panel: Theme, Palette, KPIs, Charts, Sections — a single
  // panel, rebuilt from analysis already in memory (no round trip to the
  // task pane) whenever it's (re)opened or a change here forces a remount. ----
  function syncSettingsPanel() {
    renderDialogThemePicker();
    syncFiltersSection();
    syncKpiSection();
    syncChartsSection();
    syncTableColumnsSection();
    els.cfgTable.checked = widgetConfig.showTable !== false;
    els.cfgFilters.checked = widgetConfig.showFilters !== false;
  }

  // Same picker as the pre-generate step (addin/theme-picker.js) — one
  // widget, two call sites, so they can't visually drift apart. Unlike the
  // toolbar's quick theme buttons, a change here also carries the palette,
  // and every change remounts immediately per the "applies right away, no
  // regenerate" requirement — no separate "Apply" step.
  function renderDialogThemePicker() {
    window.DashThemePicker.render(
      els.settingsThemePicker,
      { THEMES, Palettes },
      { theme, palette },
      {
        onThemeChange: (id) => { theme = id; syncThemeButtons(); controller.setTheme(activeTheme()); renderDialogThemePicker(); pushState(); },
        onPaletteChange: (id) => { palette = id; controller.setTheme(activeTheme()); renderDialogThemePicker(); pushState(); },
      }
    );
  }

  // Shown under a settings checklist (KPI/Filters/Charts) once more boxes
  // are checked than the dashboard has room for — LAYOUT.kpi.maxCards /
  // LAYOUT.filterBar.maxVisible / LAYOUT.chart.maxCount are real, hard caps
  // now (see engine/layout.js), so a checked box past the limit otherwise
  // looks broken ("I checked it and nothing happened") instead of explained.
  function appendLimitHint(container, enabledCount, limit) {
    if (enabledCount <= limit) return;
    const p = document.createElement('p');
    p.className = 'mapping-hint';
    p.textContent = `Showing the first ${limit} of ${enabledCount} checked — uncheck one to add another.`;
    container.appendChild(p);
  }

  function syncKpiSection() {
    els.settingsKpis.innerHTML = '';
    if (!currentAnalysis) return;
    const measures = currentAnalysis.columns.filter((c) => c.decision.role === 'measure');
    const enabled = widgetConfig.enabledKpis ? new Set(widgetConfig.enabledKpis) : null;
    for (const col of measures) {
      const label = document.createElement('label');
      label.className = 'settings-row';
      const input = document.createElement('input');
      input.type = 'checkbox';
      input.checked = enabled ? enabled.has(col.name) : true;
      input.addEventListener('change', () => onKpiToggle(col.name, input.checked, measures));
      label.appendChild(input);
      label.appendChild(document.createTextNode(' ' + col.name));
      els.settingsKpis.appendChild(label);
    }
    appendLimitHint(els.settingsKpis, enabled ? enabled.size : measures.length, Engine.Layout.LAYOUT.kpi.maxCards);
  }

  function onKpiToggle(name, checked, measures) {
    const current = widgetConfig.enabledKpis ? new Set(widgetConfig.enabledKpis) : new Set(measures.map((m) => m.name));
    if (checked) current.add(name); else current.delete(name);
    applyWidgetConfig({ enabledKpis: [...current] });
  }

  // Mirrors syncKpiSection exactly, for the settings panel's new "Filters"
  // section — LAYOUT.filterBar.maxVisible (engine/layout.js) is applied
  // AFTER this filter, so picking specific dimensions here is what actually
  // decides which 4 show, not just "the first 4 in column order."
  function syncFiltersSection() {
    els.settingsFilters.innerHTML = '';
    if (!currentAnalysis) return;
    const dimensions = currentAnalysis.columns.filter((c) => c.decision.role === 'dimension');
    if (!dimensions.length) {
      const p = document.createElement('p');
      p.className = 'mapping-hint';
      p.textContent = 'No filterable columns for this data.';
      els.settingsFilters.appendChild(p);
      return;
    }
    const enabled = widgetConfig.enabledDimensions ? new Set(widgetConfig.enabledDimensions) : null;
    for (const col of dimensions) {
      const label = document.createElement('label');
      label.className = 'settings-row';
      const input = document.createElement('input');
      input.type = 'checkbox';
      input.checked = enabled ? enabled.has(col.name) : true;
      input.addEventListener('change', () => onFilterToggle(col.name, input.checked, dimensions));
      label.appendChild(input);
      label.appendChild(document.createTextNode(' ' + col.name));
      els.settingsFilters.appendChild(label);
    }
    appendLimitHint(els.settingsFilters, enabled ? enabled.size : dimensions.length, Engine.Layout.LAYOUT.filterBar.maxVisible);
  }

  function onFilterToggle(name, checked, dimensions) {
    const current = widgetConfig.enabledDimensions ? new Set(widgetConfig.enabledDimensions) : new Set(dimensions.map((d) => d.name));
    if (checked) current.add(name); else current.delete(name);
    applyWidgetConfig({ enabledDimensions: [...current] });
  }

  // Every non-excluded column is table-eligible (matches
  // engine/layout.js#selectColumns' tableColumns) — a wide source can have
  // more of those than are worth showing in the reference table, so this
  // lets the user drop the ones they don't need, same "enabled list"
  // pattern as KPI cards above.
  function syncTableColumnsSection() {
    els.settingsTableColumns.innerHTML = '';
    if (!currentAnalysis) return;
    const tableColumns = currentAnalysis.columns.filter((c) => c.decision.role !== 'excluded');
    if (!tableColumns.length) {
      const p = document.createElement('p');
      p.className = 'mapping-hint';
      p.textContent = 'No columns to show.';
      els.settingsTableColumns.appendChild(p);
      return;
    }
    const enabled = widgetConfig.enabledTableColumns ? new Set(widgetConfig.enabledTableColumns) : null;
    for (const col of tableColumns) {
      const label = document.createElement('label');
      label.className = 'settings-row';
      const input = document.createElement('input');
      input.type = 'checkbox';
      input.checked = enabled ? enabled.has(col.name) : true;
      input.addEventListener('change', () => onTableColumnToggle(col.name, input.checked, tableColumns));
      label.appendChild(input);
      label.appendChild(document.createTextNode(' ' + col.name));
      els.settingsTableColumns.appendChild(label);
    }
  }

  function onTableColumnToggle(name, checked, tableColumns) {
    const current = widgetConfig.enabledTableColumns ? new Set(widgetConfig.enabledTableColumns) : new Set(tableColumns.map((c) => c.name));
    if (checked) current.add(name); else current.delete(name);
    applyWidgetConfig({ enabledTableColumns: [...current] });
  }

  // Per-chart enable + (within the bounds of that chart's underlying data
  // shape) type/dimension/measure — deliberately not a chart builder: the
  // type list is capped to engine/layout.js#CHART_TYPE_OPTIONS and axes
  // stay single dimension/measure pairs, same restriction the auto-planner
  // itself works under (see engine/layout.js#applyChartOverride).
  // `Engine.Layout.planCharts` re-derives the *full*, unfiltered candidate
  // list every time (not the currently-mounted, already-filtered chartPlan)
  // so a chart the user disabled earlier still has a row here to re-enable.
  function syncChartsSection() {
    els.settingsCharts.innerHTML = '';
    if (!currentAnalysis) return;
    const Layout = Engine.Layout;
    const sel = Layout.selectColumns(currentAnalysis.columns);
    const allCharts = Layout.planCharts(sel);
    if (!allCharts.length) {
      const p = document.createElement('p');
      p.className = 'mapping-hint';
      p.textContent = 'No eligible charts for this data.';
      els.settingsCharts.appendChild(p);
      return;
    }
    const enabledSet = widgetConfig.enabledCharts ? new Set(widgetConfig.enabledCharts) : null;
    const overridesByChart = widgetConfig.chartOverrides || {};

    for (const plan of allCharts) {
      const effective = Layout.applyChartOverride(plan, sel, overridesByChart[plan.id]);
      const enabled = enabledSet ? enabledSet.has(plan.id) : true;
      const measureName = (effective.measure || sel.primaryMeasure).name;

      const row = document.createElement('div');
      row.className = 'chart-config-row';

      const head = document.createElement('label');
      head.className = 'settings-row';
      const cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.checked = enabled;
      cb.addEventListener('change', () => onChartEnabledToggle(plan.id, cb.checked, allCharts));
      head.appendChild(cb);
      const label = plan.kind === 'line'
        ? `${measureName} over time`
        : `${measureName} ${effective.overrideType === 'donut' ? 'share of' : 'by'} ${effective.dimension.name}`;
      head.appendChild(document.createTextNode(' ' + label));
      row.appendChild(head);

      if (enabled) {
        const controls = document.createElement('div');
        controls.className = 'chart-config-controls';

        if (plan.kind === 'bar') {
          const typeSelect = document.createElement('select');
          for (const t of Layout.CHART_TYPE_OPTIONS.bar) {
            const opt = document.createElement('option');
            opt.value = t;
            opt.textContent = t === 'horizontalBar' ? 'Horizontal bar' : t === 'donut' ? 'Donut' : 'Bar';
            typeSelect.appendChild(opt);
          }
          typeSelect.value = effective.overrideType || 'bar';
          typeSelect.addEventListener('change', () => onChartOverrideChange(plan.id, { type: typeSelect.value }));
          controls.appendChild(typeSelect);

          const dimSelect = document.createElement('select');
          const eligibleDims = sel.dimensions.filter((d) => d.profile.uniqueCount <= Layout.LAYOUT.chart.maxDimensionCardinality);
          for (const d of eligibleDims) {
            const opt = document.createElement('option');
            opt.value = d.name;
            opt.textContent = d.name;
            dimSelect.appendChild(opt);
          }
          dimSelect.value = effective.dimension.name;
          dimSelect.addEventListener('change', () => onChartOverrideChange(plan.id, { dimensionColumn: dimSelect.value }));
          controls.appendChild(dimSelect);
        }

        const measureSelect = document.createElement('select');
        for (const m of sel.measures) {
          const opt = document.createElement('option');
          opt.value = m.name;
          opt.textContent = m.name;
          measureSelect.appendChild(opt);
        }
        measureSelect.value = measureName;
        measureSelect.addEventListener('change', () => onChartOverrideChange(plan.id, { measureColumn: measureSelect.value }));
        controls.appendChild(measureSelect);

        row.appendChild(controls);
      }

      els.settingsCharts.appendChild(row);
    }
    appendLimitHint(els.settingsCharts, enabledSet ? enabledSet.size : allCharts.length, Layout.LAYOUT.chart.maxCount);
  }

  function onChartEnabledToggle(chartId, checked, allCharts) {
    const current = widgetConfig.enabledCharts ? new Set(widgetConfig.enabledCharts) : new Set(allCharts.map((c) => c.id));
    if (checked) current.add(chartId); else current.delete(chartId);
    applyWidgetConfig({ enabledCharts: [...current] });
  }

  function onChartOverrideChange(chartId, patch) {
    const overrides = Object.assign({}, widgetConfig.chartOverrides);
    overrides[chartId] = Object.assign({}, overrides[chartId], patch);
    applyWidgetConfig({ chartOverrides: overrides });
  }

  function applyWidgetConfig(patch) {
    widgetConfig = Object.assign({}, widgetConfig, patch);
    const reconciled = { activeFilters: currentRawState().activeFilters, sort: currentRawState().sort, theme, palette, widgetConfig };
    mountLive(reconciled);
    pushState();
  }

  els.cfgTable.addEventListener('change', () => applyWidgetConfig({ showTable: els.cfgTable.checked }));
  els.cfgFilters.addEventListener('change', () => applyWidgetConfig({ showFilters: els.cfgFilters.checked }));

  els.settingsToggle.addEventListener('click', () => { syncSettingsPanel(); els.settingsPanel.hidden = false; });
  els.settingsClose.addEventListener('click', () => { els.settingsPanel.hidden = true; });

  // ---- scale-to-fit (frozen mode only) ----
  // A restored dashboard has no live source to recompute a layout from — the
  // picture's composition is baked into layoutSpec.widgets already, so the
  // only way to make it fit a different window size is to zoom the whole
  // thing as one image. Live mode reflows for real instead — see
  // syncCanvasWidthToStage below.
  function applyScale() {
    if (mode !== 'frozen' || !canvasSize) return;
    const stage = els.mount.parentElement;
    const available = stage.clientWidth || canvasSize.width;
    const scale = Math.min(1, available / canvasSize.width);
    els.mount.style.transformOrigin = 'top left';
    els.mount.style.transform = `scale(${scale})`;
    stage.style.overflowX = 'hidden';
  }

  // ---- reflow-to-fit (live mode only) ----
  // Below this, text/charts get too cramped to be worth reflowing further —
  // the stage scrolls horizontally instead of squeezing the layout past
  // this point.
  const MIN_CANVAS_WIDTH = 640;

  // Rebuilds the skeleton at a new canvasWidth (widgetConfig.canvasWidth —
  // engine/layout.js#buildSkeleton) via the exact same "remount with
  // preserved state" path mountLive already uses for Refresh/theme changes,
  // so activeFilters/sort survive a resize instead of resetting.
  function remountForResize(newWidth) {
    const liveState = controller.getState();
    const seedReconciled = {
      theme, palette,
      widgetConfig: Object.assign({}, widgetConfig, { canvasWidth: newWidth }),
      activeFilters: Object.fromEntries(Object.entries(liveState.activeFilters || {}).map(([k, v]) => [k, [...v]])),
      sort: liveState.sort || null,
    };
    mountLive(seedReconciled, null, null);
  }

  function syncCanvasWidthToStage() {
    if (mode !== 'live' || !controller) return;
    const stage = els.mount.parentElement;
    const available = stage.clientWidth || Engine.Layout.LAYOUT.width;
    stage.style.overflowX = available < MIN_CANVAS_WIDTH ? 'auto' : 'hidden';
    const newWidth = Math.max(MIN_CANVAS_WIDTH, Math.round(available));
    const currentWidth = (widgetConfig && widgetConfig.canvasWidth) || Engine.Layout.LAYOUT.width;
    // A few px of jitter (scrollbar appearing/disappearing, etc.) shouldn't
    // trigger a full remount — only a real, visible width change should.
    if (Math.abs(newWidth - currentWidth) < 8) return;
    remountForResize(newWidth);
  }

  let resizeTimer = null;
  window.addEventListener('resize', () => {
    applyScale();
    clearTimeout(resizeTimer);
    resizeTimer = setTimeout(syncCanvasWidthToStage, 150);
  });

  function start() {
    setStatus('Loading dashboard data…');
    initTransport(
      (raw) => { dataReceiver(raw) || placeResultReceiver(raw) || openMappingResultReceiver(raw) || refreshResultReceiver(raw) || changeRangeResultReceiver(raw); },
      () => Msg.sendChunked(parentSend, Msg.KIND.READY, {})
    );
  }

  if (window.Office && typeof Office.onReady === 'function') {
    Office.onReady(start);
  } else {
    document.addEventListener('DOMContentLoaded', start);
    if (document.readyState !== 'loading') start();
  }
})();
