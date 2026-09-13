/*
 * The dashboard dialog — a separate window (Office.context.ui.displayDialogAsync)
 * that hosts the live, fully-interactive dashboard so it isn't squeezed into
 * the narrow task pane. It never touches Excel directly (dialogs can't reach
 * the Office document object model): everything it needs arrives from the
 * task pane as JSON over dialog-messaging.js's chunked protocol, and the one
 * thing it can't do itself — placing the dashboard image on the sheet — it
 * asks the task pane to do, over the same channel.
 *
 * Two hydration modes, chosen by the task pane at data-send time:
 *   - 'live'   -> Dom.mount(): full interactivity (filters, chart clicks,
 *                 sort, pagination) plus a "Place image on sheet" button.
 *   - 'frozen' -> Dom.mountFrozen(): restoring a previously placed
 *                 dashboard that has no source rows left — same window,
 *                 same theme switcher, filters read-only (mountFrozen's own
 *                 contract), no place button (nothing new to place).
 */
(() => {
  'use strict';

  const Msg = window.DashDialogMessaging;
  const els = {
    themes: document.getElementById('dlg-themes'),
    status: document.getElementById('dlg-status'),
    mount: document.getElementById('dlg-mount'),
  };

  let controller = null;
  let theme = 'light';
  let canvasSize = null; // {width,height} — captured at hydration time; mountFrozen's controller has no getLayoutSpec() to ask later
  let parentSend = null; // (rawString) => void, set once the transport is ready

  function setStatus(text, isError) {
    els.status.hidden = !text;
    els.status.textContent = text || '';
    els.status.dataset.error = isError ? '1' : '0';
  }

  // ---- transport: real Office dialog, or a same-origin window.opener +
  // postMessage stand-in for testing outside Excel (addin/taskpane.js's
  // previewHost opens this same page via window.open, not displayDialogAsync
  // — see its dialogTransport). Only the transport differs; every message
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

  // ---- hydration: receive {mode, analysis|layoutSpec, meta} and mount ----
  const dataReceiver = Msg.createChunkReceiver(Msg.KIND.DATA, (data) => {
    setStatus('');
    try {
      mountDashboard(data);
    } catch (err) {
      setStatus(`Could not render the dashboard: ${err && err.message ? err.message : String(err)}`, true);
    }
  });

  // ---- place-on-sheet round trip: one pending request at a time (the
  // button in render/dom.js disables itself while busy) ----
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

  function mountDashboard(data) {
    els.mount.innerHTML = '';
    const THEMES = window.DashRenderThemes.THEMES;
    if (data.mode === 'frozen') {
      canvasSize = data.layoutSpec.canvas;
      controller = window.DashRenderDom.mountFrozen(els.mount, data.layoutSpec, THEMES[theme]);
    } else {
      controller = window.DashRenderDom.mount(els.mount, data.analysis, THEMES[theme], data.meta, { onPlaceOnSheet });
      canvasSize = controller.getLayoutSpec().canvas;
    }
    applyScale();
  }

  // Dashboards are laid out at a fixed design width (engine/layout.js's
  // LAYOUT.width) — shrink to fit the dialog the same way the task pane
  // used to shrink the preview to fit its narrow column.
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

  els.themes.addEventListener('click', (e) => {
    const btn = e.target.closest('button[data-theme]');
    if (!btn || !controller) return;
    theme = btn.dataset.theme;
    [...els.themes.querySelectorAll('button')].forEach((b) => b.setAttribute('aria-pressed', String(b === btn)));
    controller.setTheme(window.DashRenderThemes.THEMES[theme]); // re-paints in place — activeFilters/sort/page untouched, see render/dom.js's setTheme
  });

  function start() {
    setStatus('Loading dashboard data…');
    initTransport(
      (raw) => { dataReceiver(raw) || placeResultReceiver(raw); },
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
