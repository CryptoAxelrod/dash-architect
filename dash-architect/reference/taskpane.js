/*
 * Dash Architect (demo)
 *
 * The pane looks like a real builder, but every choice leads to the same result:
 * the prepared "Dashboard" sheet is revealed and activated.
 *
 * Why a prepared sheet: the resizable HTML object inside a worksheet is a separate
 * content add-in (manifest-viewer.xml). Office.js has no call that inserts a content
 * add-in, so the demo workbook keeps one ready on a hidden sheet. See README.md.
 *
 * Hidden reset for repeated takes: double-click the logo. It hides the Dashboard
 * sheet again, returns to your data sheet and clears the pane.
 */
(() => {
  'use strict';

  const CONFIG = {
    dashboardSheet: 'Dashboard', // name of the hidden sheet that holds the viewer
    stepMs: 560,                 // duration of each fake build step
  };

  const WORK = {
    kpi: 'Calculating KPI cards',
    chart: 'Drawing the chart',
    table: 'Building the table',
    toggles: 'Wiring up toggles',
  };
  const ORDER = ['kpi', 'chart', 'table', 'toggles'];

  const PREVIEW_RANGE = {
    address: "'Store sales'!A1:E301",
    rows: 301,
    cols: 5,
    headers: ['Date', 'Product category', 'Sales amount', 'Net profit', 'Margin'],
  };

  const $ = (id) => document.getElementById(id);
  const wait = (ms) => new Promise((r) => setTimeout(r, ms));
  const plural = (n, one, many) => `${n.toLocaleString('en-US')} ${n === 1 ? one : many}`;

  const state = {
    excel: false,
    picking: false,
    range: null,        // confirmed range
    beforePick: null,   // range to restore on Escape
    selected: new Set(),
    busy: false,
    result: null,       // { created, source }
  };

  let els = {};

  /* ---------- start ---------- */
  let started = false;
  function start(isExcel) {
    if (started) return;
    started = true;
    state.excel = !!isExcel;
    init();
  }

  if (window.Office && typeof Office.onReady === 'function') {
    Office.onReady((info) => start(info && info.host === Office.HostType.Excel));
  } else {
    document.addEventListener('DOMContentLoaded', () => start(false));
    if (document.readyState !== 'loading') start(false);
  }

  function init() {
    els = {
      refedit: $('refedit'), refBtn: $('ref-btn'), placeholder: $('ref-placeholder'),
      cells: $('ref-cells'), sheet: $('ref-sheet'), hint: $('ref-hint'),
      meta: $('ref-meta'), size: $('ref-size'), headers: $('headers'),
      stepRange: $('step-range'), stepComp: $('step-components'), count: $('comp-count'),
      summary: $('summary'), build: $('build'), generate: $('generate'),
      ctaLabel: $('cta-label'), ctaFill: $('cta-fill'), error: $('build-error'),
      doneSub: $('done-sub'), open: $('open-dashboard'), again: $('start-over'), mark: $('mark'),
    };

    els.refBtn.addEventListener('click', (e) => { e.stopPropagation(); state.picking ? finishPicking() : startPicking(); });
    els.refedit.addEventListener('click', () => { if (!state.picking && !state.busy) startPicking(); });
    document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && state.picking) cancelPicking(); });

    document.querySelectorAll('.card input').forEach((input) => {
      input.addEventListener('change', () => {
        const card = input.closest('.card');
        card.classList.toggle('is-on', input.checked);
        input.checked ? state.selected.add(input.value) : state.selected.delete(input.value);
        renderProgress();
      });
    });

    els.generate.addEventListener('click', generate);
    els.open.addEventListener('click', openDashboard);
    els.again.addEventListener('click', startOver);
    els.mark.addEventListener('dblclick', resetDemo);

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
    const r = await readSelection(false);
    if (r && state.picking) showAddress(r.address);
  }

  async function onSelectionChanged() {
    if (!state.picking) return;
    const r = await readSelection(false);
    if (r && state.picking) showAddress(r.address);
  }

  async function finishPicking() {
    const r = await readSelection(true);
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

  async function readSelection(withHeaders) {
    if (!state.excel) return PREVIEW_RANGE;
    try {
      return await Excel.run(async (ctx) => {
        const range = ctx.workbook.getSelectedRange();
        range.load('address,rowCount,columnCount');
        await ctx.sync();
        const out = { address: range.address, rows: range.rowCount, cols: range.columnCount, headers: [] };
        if (withHeaders) {
          const head = range.getCell(0, 0).getResizedRange(0, Math.min(out.cols, 12) - 1);
          head.load('values');
          await ctx.sync();
          out.headers = head.values[0];
        }
        return out;
      });
    } catch (err) {
      // selection is a shape, chart or several areas
      return null;
    }
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

    const headers = (r.headers || []).filter((v) => v !== '' && v !== null);
    const hasHeaders = headers.length > 0 && headers.every((v) => typeof v === 'string');
    const dataRows = hasHeaders ? Math.max(r.rows - 1, 0) : r.rows;
    els.size.textContent = `${plural(dataRows, 'row', 'rows')}, ${plural(r.cols, 'column', 'columns')}`;

    els.headers.innerHTML = '';
    els.headers.classList.toggle('none', !hasHeaders);
    if (hasHeaders) {
      headers.forEach((h) => {
        const cell = document.createElement('span');
        cell.textContent = h;
        els.headers.appendChild(cell);
      });
    } else {
      const cell = document.createElement('span');
      cell.textContent = 'No header row found. The first row will be treated as data.';
      els.headers.appendChild(cell);
    }
    els.meta.hidden = false;
    renderProgress();
  }

  /* ---------- progress + summary ---------- */
  function renderProgress() {
    const n = state.selected.size;
    els.stepRange.classList.toggle('done', !!state.range);
    els.stepComp.classList.toggle('done', n > 0);
    els.count.textContent = `${n} of 4`;

    const cells = state.range ? splitAddress(state.range.address).cells : null;
    if (!cells && !n) {
      els.summary.textContent = 'Choose data and components above';
    } else {
      const parts = [];
      parts.push(n ? `<strong>${plural(n, 'component', 'components')}</strong>` : 'No components yet');
      parts.push(cells ? `from <strong>${cells}</strong>` : 'and no range yet');
      els.summary.innerHTML = parts.join(' ');
    }
  }

  /* ---------- generate ---------- */
  async function generate() {
    if (state.busy) return;
    els.error.hidden = true;

    if (state.picking) await finishPicking();
    if (!state.range) {
      const r = await readSelection(true);
      if (r) setRange(r);
    }

    state.busy = true;
    document.body.classList.add('busy');
    els.build.dataset.state = 'running';

    const rows = state.range ? Math.max(state.range.rows - 1, 1) : 0;
    const steps = [
      rows ? `Reading ${plural(rows, 'row', 'rows')}` : 'Reading your data',
      ...ORDER.filter((k) => state.selected.has(k)).map((k) => WORK[k]),
      'Placing it on a new sheet',
    ];

    els.ctaFill.style.transitionDuration = `${CONFIG.stepMs}ms`;
    for (let i = 0; i < steps.length; i++) {
      els.ctaLabel.textContent = `${steps[i]}…`;
      els.ctaFill.style.width = `${((i + 1) / steps.length) * 100}%`;
      await wait(CONFIG.stepMs);
    }

    try {
      state.result = await placeDashboard();
      showDone();
    } catch (err) {
      showError(err);
    } finally {
      state.busy = false;
      document.body.classList.remove('busy');
    }
  }

  async function placeDashboard() {
    if (!state.excel) return { preview: true };

    return Excel.run(async (ctx) => {
      const sheets = ctx.workbook.worksheets;
      const active = sheets.getActiveWorksheet();
      active.load('name,position');
      let dash = sheets.getItemOrNullObject(CONFIG.dashboardSheet);
      dash.load('name,visibility');
      await ctx.sync();

      let created = false;
      if (dash.isNullObject) {
        dash = sheets.add(CONFIG.dashboardSheet); // fallback: no prepared sheet in this workbook
        created = true;
      } else if (dash.visibility !== Excel.SheetVisibility.visible) {
        dash.visibility = Excel.SheetVisibility.visible;
      }
      if (Office.context.requirements.isSetSupported('ExcelApi', '1.8')) dash.showGridlines = false;
      dash.activate();
      dash.getRange('A1').select();
      await ctx.sync();

      // place the tab right after the data sheet, like a freshly added sheet
      if (active.name !== CONFIG.dashboardSheet) {
        try {
          const count = sheets.getCount();
          dash.load('position');
          await ctx.sync();
          const target = Math.min(active.position + (dash.position > active.position ? 1 : 0), count.value - 1);
          if (target !== dash.position) { dash.position = target; await ctx.sync(); }
        } catch (e) { /* tab order is cosmetic */ }
      }

      return { created, source: active.name };
    });
  }

  function showDone() {
    const r = state.result || {};
    if (r.preview) {
      els.doneSub.textContent = 'Preview mode. Open this pane inside Excel to place the dashboard on a sheet.';
    } else if (r.created) {
      els.doneSub.textContent = `Added the ${CONFIG.dashboardSheet} sheet. Insert Dash Architect Viewer on it to show the dashboard.`;
    } else {
      els.doneSub.textContent = `It's on the ${CONFIG.dashboardSheet} sheet. Drag its edges to resize.`;
    }
    els.build.dataset.state = 'done';
    resetButton();
  }

  function showError(err) {
    els.build.dataset.state = 'idle';
    resetButton();
    const msg = err && err.message ? err.message : String(err);
    els.error.textContent = `The dashboard wasn't placed: ${msg}`;
    els.error.hidden = false;
  }

  function resetButton() {
    els.ctaFill.style.transitionDuration = '0ms';
    els.ctaFill.style.width = '0';
    els.ctaLabel.textContent = 'Generate dashboard';
  }

  /* ---------- after generation ---------- */
  async function openDashboard() {
    if (!state.excel) { window.open('dashboard.html', '_blank'); return; }
    try {
      await Excel.run(async (ctx) => {
        const dash = ctx.workbook.worksheets.getItemOrNullObject(CONFIG.dashboardSheet);
        await ctx.sync();
        if (!dash.isNullObject) { dash.visibility = Excel.SheetVisibility.visible; dash.activate(); await ctx.sync(); }
      });
    } catch (e) { /* nothing to open */ }
  }

  function startOver() {
    state.picking = false;
    state.range = null;
    state.beforePick = null;
    state.selected.clear();
    document.querySelectorAll('.card input').forEach((i) => { i.checked = false; i.closest('.card').classList.remove('is-on'); });
    els.hint.hidden = true;
    els.error.hidden = true;
    restoreRange(null);
    els.build.dataset.state = 'idle';
    resetButton();
    window.scrollTo(0, 0);
  }

  // double-click the logo: hide the Dashboard sheet again and clear the pane
  async function resetDemo() {
    if (state.busy) return;
    if (state.excel) {
      try {
        await Excel.run(async (ctx) => {
          const sheets = ctx.workbook.worksheets;
          sheets.load('items/name,items/visibility');
          await ctx.sync();
          const dash = sheets.items.find((s) => s.name === CONFIG.dashboardSheet);
          if (!dash) return;
          const source = state.result && state.result.source;
          const back = sheets.items.find((s) => s.name === source && s.name !== dash.name)
            || sheets.items.find((s) => s.name !== dash.name && s.visibility === Excel.SheetVisibility.visible);
          if (!back) return;
          back.activate();
          back.getRange('A1').select();
          if (state.result && state.result.created) dash.delete();
          else dash.visibility = Excel.SheetVisibility.hidden;
          await ctx.sync();
        });
      } catch (e) { /* leave sheets as they are */ }
    }
    state.result = null;
    startOver();
  }
})();
