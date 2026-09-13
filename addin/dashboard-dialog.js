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
 *                 settings panel (KPI + section visibility) / "Place image
 *                 on sheet".
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
    settingsKpis: document.getElementById('dlg-settings-kpis'),
    cfgCharts: document.getElementById('dlg-cfg-charts'),
    cfgTable: document.getElementById('dlg-cfg-table'),
    cfgFilters: document.getElementById('dlg-cfg-filters'),
  };

  let controller = null;
  let mode = null; // 'live' | 'frozen'
  let theme = 'light';
  let widgetConfig = { enabledKpis: null, showCharts: true, showTable: true, showFilters: true };
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

  // ---- current state, as sent to the task pane (STATE_UPDATE) or included
  // in a request the task pane will echo back once it has fresh data ----
  function currentRawState() {
    const s = controller && mode === 'live' ? controller.getState() : {};
    return {
      activeFilters: Object.fromEntries(Object.entries(s.activeFilters || {}).map(([k, v]) => [k, [...v]])),
      sort: s.sort || null,
      theme,
      widgetConfig,
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
      currentAnalysis = null;
      canvasSize = data.layoutSpec.canvas;
      controller = window.DashRenderDom.mountFrozen(els.mount, data.layoutSpec, THEMES[theme]);
      applyScale();
      return;
    }

    els.liveActions.hidden = false;
    if (data.source) { currentSourceLabel = data.source; els.source.textContent = data.source; }

    const seedReconciled = data.seedState ? Engine.Reconcile.reconcileDashboardState(data.analysis, data.seedState) : null;
    if (currentAnalysis) {
      const changed = Engine.Reconcile.diffValueTypes(currentAnalysis, data.analysis);
      if (changed.length) showBanner(`Column type changed after refresh: ${changed.join(', ')}. Data was recalculated, your filters and settings were kept.`, false);
    }

    currentAnalysis = data.analysis;
    currentMeta = data.meta;
    mountLive(seedReconciled);
  }

  function mountLive(seedReconciled) {
    if (seedReconciled) {
      theme = seedReconciled.theme || theme;
      widgetConfig = seedReconciled.widgetConfig || widgetConfig;
    }
    syncThemeButtons();
    syncSettingsPanel();

    els.mount.innerHTML = '';
    const initialState = seedReconciled ? { activeFilters: seedReconciled.activeFilters, sort: seedReconciled.sort } : undefined;
    const metaWithConfig = Object.assign({}, currentMeta, { widgetConfig });
    controller = window.DashRenderDom.mount(els.mount, currentAnalysis, THEMES[theme], metaWithConfig, {
      onPlaceOnSheet,
      initialState,
      onStateChange: pushState,
    });
    canvasSize = controller.getLayoutSpec().canvas;
    applyScale();
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
      const requestId = Msg.sendChunked(parentSend, Msg.KIND.PLACE_ON_SHEET, Object.assign({}, spec, { theme }));
      pendingPlacements.set(requestId, { resolve });
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
    controller.setTheme(THEMES[theme]); // re-paints in place — activeFilters/sort/page untouched, see render/dom.js's setTheme
    pushState();
  });

  // ---- settings panel: KPI + section visibility, rebuilt from analysis
  // already in memory — no round trip to the task pane at all ----
  function syncSettingsPanel() {
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
    els.cfgCharts.checked = widgetConfig.showCharts !== false;
    els.cfgTable.checked = widgetConfig.showTable !== false;
    els.cfgFilters.checked = widgetConfig.showFilters !== false;
  }

  function onKpiToggle(name, checked, measures) {
    const current = widgetConfig.enabledKpis ? new Set(widgetConfig.enabledKpis) : new Set(measures.map((m) => m.name));
    if (checked) current.add(name); else current.delete(name);
    applyWidgetConfig({ enabledKpis: [...current] });
  }

  function applyWidgetConfig(patch) {
    widgetConfig = Object.assign({}, widgetConfig, patch);
    const reconciled = { activeFilters: currentRawState().activeFilters, sort: currentRawState().sort, theme, widgetConfig };
    mountLive(reconciled);
    pushState();
  }

  els.cfgCharts.addEventListener('change', () => applyWidgetConfig({ showCharts: els.cfgCharts.checked }));
  els.cfgTable.addEventListener('change', () => applyWidgetConfig({ showTable: els.cfgTable.checked }));
  els.cfgFilters.addEventListener('change', () => applyWidgetConfig({ showFilters: els.cfgFilters.checked }));

  els.settingsToggle.addEventListener('click', () => { syncSettingsPanel(); els.settingsPanel.hidden = false; });
  els.settingsClose.addEventListener('click', () => { els.settingsPanel.hidden = true; });

  // ---- scale-to-fit ----
  function applyScale() {
    if (!canvasSize) return;
    const stage = els.mount.parentElement;
    const available = stage.clientWidth || canvasSize.width;
    const scale = Math.min(1, available / canvasSize.width);
    els.mount.style.transformOrigin = 'top left';
    els.mount.style.transform = `scale(${scale})`;
    stage.style.overflowX = 'hidden';
  }
  window.addEventListener('resize', applyScale);

  function start() {
    setStatus('Loading dashboard data…');
    initTransport(
      (raw) => { dataReceiver(raw) || placeResultReceiver(raw) || refreshResultReceiver(raw) || changeRangeResultReceiver(raw); },
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
