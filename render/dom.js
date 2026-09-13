/*
 * Interactive DOM renderer: layoutSpec + theme -> a live, absolutely
 * positioned HTML tree (filters, KPI text, sortable/paginated table are
 * real DOM; charts are individually-scoped inline <svg> fragments reusing
 * render/svg.js's per-chart-type drawing code, so a chart looks pixel
 * identical whether it's on screen or in the exported static snapshot).
 *
 * Every widget keeps the exact rect engine/layout.js gave it — filtering
 * or sorting only asks engine/layout.js#recompute for fresh widget *data*
 * and re-paints in place; it never moves or adds/removes a widget.
 *
 * Two entry points:
 *   - mount(container, analysis, theme, meta) — the live, fully-interactive
 *     dashboard, backed by the real column data (filters recompute
 *     aggregates on the fly).
 *   - mountFrozen(container, layoutSpec, theme) — a dashboard restored from
 *     a serialized snapshot (engine/layout.js#buildStorageSnapshot) with no
 *     underlying data at all — e.g. addin/dashboard-io.js reopening a
 *     dashboard someone placed earlier. Filters render read-only (there is
 *     no row data left to recompute against); the table's one embedded
 *     page of rows still sorts client-side; the share button still works.
 * Header/KPI/chart widgets carry all their own numbers either way, so they
 * share one set of stateless renderers between both entry points.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('./format'), require('./svg'), require('./share'));
  } else {
    root.DashRenderDom = factory(root.DashRenderFormat, root.DashRenderSvg, root.DashRenderShare);
  }
})(typeof self !== 'undefined' ? self : this, function (Format, Svg, Share) {
  'use strict';

  const px = (n) => n + 'px';

  function el(tag, attrs, children) {
    const node = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs || {})) {
      if (k === 'style') Object.assign(node.style, v);
      else if (k.startsWith('on') && typeof v === 'function') node.addEventListener(k.slice(2), v);
      else if (v != null) node.setAttribute(k, v);
    }
    for (const child of [].concat(children || [])) {
      if (child == null) continue;
      node.appendChild(typeof child === 'string' ? document.createTextNode(child) : child);
    }
    return node;
  }

  function absRect(r, extra) {
    return Object.assign({ position: 'absolute', left: px(r.x), top: px(r.y), width: px(r.w), height: px(r.h) }, extra);
  }

  function navBtnStyle(c) {
    return { font: 'inherit', color: 'inherit', border: `1px solid ${c.rule}`, background: 'transparent', borderRadius: '6px', padding: '3px 9px', cursor: 'pointer' };
  }

  function formatCell(col, value, theme) {
    if (value == null) return '—';
    if (col.role === 'time' || col.cellFormat === 'date') {
      const d = new Date(value);
      const MONTHS_ = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
      return `${MONTHS_[d.getUTCMonth()]} ${d.getUTCDate()}, ${d.getUTCFullYear()}`;
    }
    if (col.role === 'measure') return Format.formatMeasureValue(col, value, { negativeStyle: theme.negativeStyle });
    return String(value);
  }

  // --- stateless widget renderers: header/KPI/chart carry every number
  // they need inline, live or frozen, so one implementation covers both. --

  function renderHeader(w, theme, shareButton) {
    const c = theme.color;
    const wrap = el('div', { style: absRect(w.rect) });
    const titleCol = el('div', { style: { position: 'absolute', left: 0, top: 0, right: '140px' } });
    titleCol.appendChild(el('h1', { style: { margin: 0, font: `700 ${theme.type.h1}px ${theme.font.family}`, color: c.ink, letterSpacing: '-0.01em' } }, w.title || ''));
    if (w.subtitle) {
      titleCol.appendChild(el('p', { style: { margin: '4px 0 0', font: `${theme.type.subtitle}px ${theme.font.family}`, color: c.muted } }, w.subtitle));
    }
    wrap.appendChild(titleCol);
    if (shareButton) wrap.appendChild(shareButton);
    return wrap;
  }

  function renderShareButton(theme, getSpecForShare, columnsByName) {
    const c = theme.color;
    const btn = el('button', {
      type: 'button',
      style: {
        position: 'absolute', right: 0, top: 0, font: `13px ${theme.font.family}`, fontWeight: 600,
        padding: '8px 14px', borderRadius: px(theme.radius.md), border: `1px solid ${c.accent}`,
        background: c.accent, color: c.panel, cursor: 'pointer',
      },
    }, 'Copy image');
    btn.addEventListener('click', async () => {
      if (btn.dataset.busy) return;
      btn.dataset.busy = '1';
      const original = btn.textContent;
      try {
        const result = await Share.copyDashboardImage(getSpecForShare(), theme, columnsByName);
        btn.textContent = result.method === 'clipboard' ? 'Copied' : 'Downloaded';
      } catch (e) {
        btn.textContent = 'Could not copy';
      }
      setTimeout(() => { btn.textContent = original; delete btn.dataset.busy; }, 1600);
    });
    return btn;
  }

  function renderKpi(w, theme) {
    const c = theme.color;
    const isHero = w.variant === 'hero';
    const wrap = el('div', { style: absRect(w.rect, isHero ? {} : { borderLeft: `1px solid ${c.rule}`, paddingLeft: px(theme.spacing.md) }) });
    wrap.appendChild(el('p', { style: { margin: 0, font: `${theme.type.kpiLabel}px ${theme.font.family}`, color: c.muted } }, w.label));
    const valueStr = Format.formatMeasureValue(w, w.value, { negativeStyle: theme.negativeStyle });
    const negative = w.value != null && w.value < 0;
    const baseSize = isHero ? theme.type.kpiHero : theme.type.kpiValue;
    const availableWidth = isHero ? w.rect.w : w.rect.w - theme.spacing.md;
    const size = Format.fitFontSize(valueStr, availableWidth, baseSize, isHero ? 24 : 14);
    wrap.appendChild(el('p', {
      style: {
        margin: '6px 0 4px', font: `750 ${size}px ${theme.font.family}`, letterSpacing: '-0.01em',
        color: negative && theme.negativeStyle === 'color' ? c.negative : c.ink,
        whiteSpace: 'nowrap',
      },
    }, valueStr));
    const caption = w.approximate ? 'Unweighted average — no verified weight base' : `${w.count.toLocaleString('en-US')} row${w.count === 1 ? '' : 's'}`;
    wrap.appendChild(el('p', { style: { margin: 0, font: `${theme.type.kpiLabel - 1.5}px ${theme.font.family}`, color: c.faint } }, caption));
    return wrap;
  }

  function renderChartWidget(w, theme) {
    const svgMarkup = Svg.renderChart(w, theme);
    const holder = el('div', { style: absRect(w.rect) });
    holder.innerHTML = `<svg viewBox="${w.rect.x} ${w.rect.y} ${w.rect.w} ${w.rect.h}" width="100%" height="100%">${svgMarkup}</svg>`;
    return holder;
  }

  function tableChrome(w, theme) {
    const c = theme.color;
    const wrap = el('div', { style: Object.assign(absRect(w.rect), theme.card.style === 'panel'
      ? { background: c.panel, border: `${theme.card.borderWidth}px solid ${c.panelBorder}`, borderRadius: px(theme.radius.lg), boxSizing: 'border-box', overflow: 'hidden' }
      : { borderBottom: `${theme.card.borderWidth}px solid ${c.rule}`, boxSizing: 'border-box' }) });
    const padPx = theme.card.style === 'panel' ? theme.spacing.lg : theme.spacing.sm;
    const inner = el('div', { style: { padding: px(padPx), display: 'flex', flexDirection: 'column', height: '100%', boxSizing: 'border-box' } });
    wrap.appendChild(inner);
    return { wrap, inner, padPx };
  }

  function tableVisibleRows(w, theme, padPx) {
    const headerH = theme.table.headerSize + 16;
    const footerH = 30;
    return Math.max(1, Math.floor((w.rect.h - padPx * 2 - headerH - footerH) / theme.table.rowHeight));
  }

  // =====================================================================
  // Live mode
  // =====================================================================

  /**
   * @param {HTMLElement} container mounted with position:relative and an explicit size (see updateCanvasSize)
   * @param {{rowCount:number, columns:Array}} analysis engine/index.js#analyzeTable result
   * @param {object} theme render/themes.js theme
   * @param {{title?:string, subtitle?:string}} [meta]
   * @returns {{setTheme(theme):void, getState():object, getLayoutSpec():object}}
   */
  function mount(container, analysis, theme, meta) {
    const Engine = window.DashEngine; // browser-global; Node callers pass their own via a future param if ever needed
    const layoutSpec = Engine.buildLayoutSpec(analysis, meta);
    const columnsByName = Engine.Aggregate.byName(analysis.columns);

    const state = {
      theme,
      activeFilters: {}, // column -> Set(values)
      sort: null, // {key, dir}
      openFilter: null, // column name whose popover is open
      page: 0,
    };

    function currentWidgets() {
      return Engine.recomputeLayout(analysis, layoutSpec.widgets, { activeFilters: state.activeFilters, sort: state.sort });
    }

    function updateCanvasSize() {
      Object.assign(container.style, { position: 'relative', width: px(layoutSpec.canvas.width), height: px(layoutSpec.canvas.height), background: state.theme.color.paper });
    }

    function renderAll() {
      updateCanvasSize();
      container.innerHTML = '';
      for (const w of currentWidgets()) container.appendChild(renderWidget(w));
    }

    function setState(patch) {
      Object.assign(state, patch);
      renderAll();
    }

    function renderWidget(w) {
      if (w.type === 'header') return renderHeader(w, state.theme, renderShareButton(state.theme, () => ({ canvas: layoutSpec.canvas, widgets: currentWidgets() }), columnsByName));
      if (w.type === 'filterBar') return renderFilterBar(w);
      if (w.type === 'kpi') return renderKpi(w, state.theme);
      if (w.type === 'table') return renderTable(w);
      return renderChartWidget(w, state.theme);
    }

    function renderFilterBar(w) {
      const c = state.theme.color;
      const wrap = el('div', { style: Object.assign(absRect(w.rect), { display: 'flex', alignItems: 'center', gap: px(state.theme.spacing.sm), flexWrap: 'wrap' }) });

      for (const f of w.filters) {
        const active = state.activeFilters[f.column];
        const activeArr = active ? [...active] : [];
        const filled = activeArr.length > 0;
        const label = !filled ? `${f.column}: All` : activeArr.length === 1 ? `${f.column}: ${activeArr[0]}` : `${f.column}: ${activeArr.length} selected`;

        const pillWrap = el('div', { class: 'dash-filter-pill', style: { position: 'relative' } });
        const pill = el('button', {
          type: 'button',
          style: {
            font: `12.5px ${state.theme.font.family}`, padding: '6px 13px', borderRadius: px(state.theme.radius.pill),
            border: `1px solid ${filled ? c.accent : c.rule}`, background: filled ? c.accent : 'transparent',
            color: filled ? c.panel : c.ink, cursor: 'pointer',
          },
          onclick: () => setState({ openFilter: state.openFilter === f.column ? null : f.column }),
        }, label);
        pillWrap.appendChild(pill);

        if (state.openFilter === f.column) pillWrap.appendChild(renderFilterPopover(f, active));
        wrap.appendChild(pillWrap);
      }

      if (w.overflow && w.overflow.length) {
        wrap.appendChild(el('span', { style: { font: `12.5px ${state.theme.font.family}`, color: c.muted } }, `+${w.overflow.length} more`));
      }
      return wrap;
    }

    function renderFilterPopover(f, active) {
      const c = state.theme.color;
      const pop = el('div', {
        style: {
          position: 'absolute', top: '36px', left: 0, zIndex: 10, background: c.panel, border: `1px solid ${c.rule}`,
          borderRadius: px(state.theme.radius.md), padding: px(state.theme.spacing.sm), minWidth: '160px',
          boxShadow: '0 6px 20px rgba(0,0,0,.12)', font: `13px ${state.theme.font.family}`,
        },
      });
      const allRow = el('label', { style: { display: 'flex', gap: '6px', padding: '3px 0', cursor: 'pointer', color: c.ink } }, [
        el('input', { type: 'checkbox', checked: !active || active.size === 0 ? true : null, onchange: () => setState({ activeFilters: Object.assign({}, state.activeFilters, { [f.column]: new Set() }) }) }),
        'All',
      ]);
      pop.appendChild(allRow);
      for (const v of f.values) {
        const checked = active && active.has(v.value);
        pop.appendChild(el('label', { style: { display: 'flex', gap: '6px', padding: '3px 0', cursor: 'pointer', color: c.ink } }, [
          el('input', {
            type: 'checkbox', checked: checked ? true : null,
            onchange: (e) => {
              const next = new Set(active || []);
              if (e.target.checked) next.add(v.value);
              else next.delete(v.value);
              setState({ activeFilters: Object.assign({}, state.activeFilters, { [f.column]: next }) });
            },
          }),
          `${v.value} (${v.count})`,
        ]));
      }
      return pop;
    }

    function renderTable(w) {
      const c = state.theme.color;
      const t = state.theme.table;
      const { wrap, inner, padPx } = tableChrome(w, state.theme);
      const visibleRows = tableVisibleRows(w, state.theme, padPx);
      const totalPages = Math.max(1, Math.ceil(w.rowIndices.length / visibleRows));
      state.page = Math.min(state.page, totalPages - 1);
      const from = state.page * visibleRows;
      const pageIdx = w.rowIndices.slice(from, from + visibleRows);

      const table = el('table', { style: { borderCollapse: 'collapse', width: '100%', font: `${t.cellSize}px ${state.theme.font.family}`, flex: '1 1 auto' } });
      const thead = el('thead');
      const headRow = el('tr');
      for (const col of w.columns) {
        const isNum = col.role === 'measure';
        const sorted = w.sort.key === col.name;
        headRow.appendChild(el('th', {
          style: {
            textAlign: isNum ? 'right' : 'left', padding: '4px 10px 8px', fontSize: px(t.headerSize), color: c.muted,
            borderBottom: `1px solid ${c.rule}`, cursor: 'pointer', userSelect: 'none', fontWeight: 650, whiteSpace: 'nowrap',
          },
          onclick: () => setState({ sort: { key: col.name, dir: sorted && w.sort.dir === 'asc' ? 'desc' : 'asc' }, page: 0 }),
        }, col.name + (sorted ? (w.sort.dir === 'asc' ? ' ↑' : ' ↓') : '')));
      }
      thead.appendChild(headRow);
      table.appendChild(thead);

      const tbody = el('tbody');
      for (const rowIndex of pageIdx) {
        const tr = el('tr');
        for (const col of w.columns) {
          const isNum = col.role === 'measure';
          const raw = columnsByName.get(col.name).values[rowIndex];
          const str = formatCell(col, raw, state.theme);
          const isNegative = isNum && raw != null && raw < 0 && state.theme.negativeStyle === 'color';
          tr.appendChild(el('td', { style: { textAlign: isNum ? 'right' : 'left', padding: '6px 10px', borderBottom: `1px solid ${c.ruleSoft}`, color: isNegative ? c.negative : c.ink, whiteSpace: 'nowrap' } }, str));
        }
        tbody.appendChild(tr);
      }
      table.appendChild(tbody);
      inner.appendChild(table);

      const footer = el('div', { style: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', paddingTop: '8px', font: `${t.headerSize}px ${state.theme.font.family}`, color: c.faint } });
      footer.appendChild(el('span', {}, `${from + 1}–${Math.min(from + visibleRows, w.rowIndices.length)} of ${w.totalRows.toLocaleString('en-US')} rows`));
      const nav = el('div', { style: { display: 'flex', gap: '8px' } });
      nav.appendChild(el('button', { type: 'button', disabled: state.page <= 0 ? true : null, style: navBtnStyle(c), onclick: () => setState({ page: state.page - 1 }) }, 'Previous'));
      nav.appendChild(el('button', { type: 'button', disabled: state.page >= totalPages - 1 ? true : null, style: navBtnStyle(c), onclick: () => setState({ page: state.page + 1 }) }, 'Next'));
      footer.appendChild(nav);
      inner.appendChild(footer);

      return wrap;
    }

    renderAll();

    document.addEventListener('click', (e) => {
      if (state.openFilter && !e.target.closest('.dash-filter-pill')) setState({ openFilter: null });
    });

    return {
      setTheme(nextTheme) { state.theme = nextTheme; renderAll(); },
      getState() { return state; },
      getLayoutSpec() { return layoutSpec; },
      getColumnsByName() { return columnsByName; },
    };
  }

  // =====================================================================
  // Frozen mode — restored from a serialized snapshot, no source data.
  // Filters are read-only; the table sorts only the rows already embedded
  // in it (engine/layout.js#buildStorageSnapshot caps that at the same
  // count the static PNG shows) — see this file's header comment.
  // =====================================================================

  /**
   * @param {HTMLElement} container
   * @param {{canvas:object, widgets:Array}} layoutSpec a snapshot from engine/layout.js#buildStorageSnapshot
   * @param {object} theme
   */
  function mountFrozen(container, layoutSpec, theme) {
    const state = { theme, sort: null, page: 0 };

    function currentWidgets() {
      return layoutSpec.widgets.map((w) => {
        if (w.type !== 'table' || !state.sort) return w;
        return Object.assign({}, w, { sort: state.sort, rows: sortFrozenRows(w, state.sort) });
      });
    }

    function sortFrozenRows(w, sort) {
      const colIdx = w.columns.findIndex((c) => c.name === sort.key);
      if (colIdx < 0) return w.rows;
      const mul = sort.dir === 'desc' ? -1 : 1;
      return w.rows.slice().sort((a, b) => {
        const va = a[colIdx];
        const vb = b[colIdx];
        if (va == null && vb == null) return 0;
        if (va == null) return 1;
        if (vb == null) return -1;
        const c = typeof va === 'number' && typeof vb === 'number' ? va - vb : String(va).localeCompare(String(vb));
        return c * mul;
      });
    }

    function updateCanvasSize() {
      Object.assign(container.style, { position: 'relative', width: px(layoutSpec.canvas.width), height: px(layoutSpec.canvas.height), background: state.theme.color.paper });
    }

    function renderAll() {
      updateCanvasSize();
      container.innerHTML = '';
      for (const w of currentWidgets()) container.appendChild(renderWidget(w));
    }

    function setState(patch) {
      Object.assign(state, patch);
      renderAll();
    }

    function renderWidget(w) {
      if (w.type === 'header') return renderHeader(w, state.theme, renderShareButton(state.theme, () => ({ canvas: layoutSpec.canvas, widgets: currentWidgets() }), null));
      if (w.type === 'filterBar') return renderReadOnlyFilterBar(w);
      if (w.type === 'kpi') return renderKpi(w, state.theme);
      if (w.type === 'table') return renderFrozenTable(w);
      return renderChartWidget(w, state.theme);
    }

    // No source rows survive a restore, so a filter selection can't be
    // recomputed — shown as inert labels of the state the dashboard was
    // generated with, not as clickable controls (a control that visibly
    // does nothing on click is worse than no control).
    function renderReadOnlyFilterBar(w) {
      const c = state.theme.color;
      const wrap = el('div', { style: Object.assign(absRect(w.rect), { display: 'flex', alignItems: 'center', gap: px(state.theme.spacing.sm), flexWrap: 'wrap' }) });
      for (const f of w.filters) {
        const active = f.active || [];
        const label = active.length === 0 ? `${f.column}: All` : active.length === 1 ? `${f.column}: ${active[0]}` : `${f.column}: ${active.length} selected`;
        wrap.appendChild(el('span', {
          title: 'Restored from a saved snapshot — open the source range and regenerate to change filters',
          style: {
            font: `12.5px ${state.theme.font.family}`, padding: '6px 13px', borderRadius: px(state.theme.radius.pill),
            border: `1px solid ${c.rule}`, color: c.faint,
          },
        }, label));
      }
      if (w.overflow && w.overflow.length) {
        wrap.appendChild(el('span', { style: { font: `12.5px ${state.theme.font.family}`, color: c.muted } }, `+${w.overflow.length} more`));
      }
      return wrap;
    }

    function renderFrozenTable(w) {
      const c = state.theme.color;
      const t = state.theme.table;
      const { wrap, inner, padPx } = tableChrome(w, state.theme);
      const visibleRows = tableVisibleRows(w, state.theme, padPx);
      const totalPages = Math.max(1, Math.ceil(w.rows.length / visibleRows));
      state.page = Math.min(state.page, totalPages - 1);
      const from = state.page * visibleRows;
      const pageRows = w.rows.slice(from, from + visibleRows);

      const table = el('table', { style: { borderCollapse: 'collapse', width: '100%', font: `${t.cellSize}px ${state.theme.font.family}`, flex: '1 1 auto' } });
      const thead = el('thead');
      const headRow = el('tr');
      for (const col of w.columns) {
        const isNum = col.role === 'measure';
        const sorted = state.sort && state.sort.key === col.name;
        headRow.appendChild(el('th', {
          style: {
            textAlign: isNum ? 'right' : 'left', padding: '4px 10px 8px', fontSize: px(t.headerSize), color: c.muted,
            borderBottom: `1px solid ${c.rule}`, cursor: 'pointer', userSelect: 'none', fontWeight: 650, whiteSpace: 'nowrap',
          },
          onclick: () => setState({ sort: { key: col.name, dir: sorted && state.sort.dir === 'asc' ? 'desc' : 'asc' }, page: 0 }),
        }, col.name + (sorted ? (state.sort.dir === 'asc' ? ' ↑' : ' ↓') : '')));
      }
      thead.appendChild(headRow);
      table.appendChild(thead);

      const tbody = el('tbody');
      for (const rowValues of pageRows) {
        const tr = el('tr');
        w.columns.forEach((col, ci) => {
          const isNum = col.role === 'measure';
          const raw = rowValues[ci];
          const str = formatCell(col, raw, state.theme);
          const isNegative = isNum && raw != null && raw < 0 && state.theme.negativeStyle === 'color';
          tr.appendChild(el('td', { style: { textAlign: isNum ? 'right' : 'left', padding: '6px 10px', borderBottom: `1px solid ${c.ruleSoft}`, color: isNegative ? c.negative : c.ink, whiteSpace: 'nowrap' } }, str));
        });
        tbody.appendChild(tr);
      }
      table.appendChild(tbody);
      inner.appendChild(table);

      const footer = el('div', { style: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', paddingTop: '8px', font: `${t.headerSize}px ${state.theme.font.family}`, color: c.faint } });
      footer.appendChild(el('span', {}, `${from + 1}–${Math.min(from + visibleRows, w.rows.length)} of ${w.shownRows} rows (snapshot)`));
      const nav = el('div', { style: { display: 'flex', gap: '8px' } });
      nav.appendChild(el('button', { type: 'button', disabled: state.page <= 0 ? true : null, style: navBtnStyle(c), onclick: () => setState({ page: state.page - 1 }) }, 'Previous'));
      nav.appendChild(el('button', { type: 'button', disabled: state.page >= totalPages - 1 ? true : null, style: navBtnStyle(c), onclick: () => setState({ page: state.page + 1 }) }, 'Next'));
      footer.appendChild(nav);
      inner.appendChild(footer);

      return wrap;
    }

    renderAll();

    return { setTheme(nextTheme) { state.theme = nextTheme; renderAll(); }, getState() { return state; } };
  }

  return { mount, mountFrozen };
});
