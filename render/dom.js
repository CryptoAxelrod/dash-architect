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
    // A year_like time column (engine/roles.js rule 3) stores the bare year
    // itself as its value (e.g. 2025), not an epoch — same reason
    // engine/aggregate.js#bucketByTime special-cases it before treating the
    // value as milliseconds. Same distinction as `format_date` is what sets
    // it apart from an ordinary date/time column, so no separate flag.
    if (col.granularity === 'year') return String(value);
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

  // `onRename` is only passed by mount() (live mode) — a frozen/restored
  // dashboard (mountFrozen) has no save path for a new title (no source, no
  // "Place image on sheet"), so it never gets the pencil at all rather than
  // offering an edit box that can't actually persist.
  function renderHeader(w, theme, shareButton, placeButton, onRename) {
    const c = theme.color;
    const wrap = el('div', { style: absRect(w.rect) });
    const rightWidth = placeButton ? '300px' : '140px';
    const titleCol = el('div', { style: { position: 'absolute', left: 0, top: 0, right: rightWidth } });

    function subtitleEl() {
      return w.subtitle ? el('p', { style: { margin: '4px 0 0', font: `${theme.type.subtitle}px ${theme.font.family}`, color: c.muted } }, w.subtitle) : null;
    }

    function showDisplay() {
      titleCol.innerHTML = '';
      const row = el('div', { style: { display: 'flex', alignItems: 'baseline', gap: '8px' } });
      row.appendChild(el('h1', { style: { margin: 0, font: `700 ${theme.type.h1}px ${theme.font.family}`, color: c.ink, letterSpacing: '-0.01em' } }, w.title || ''));
      if (onRename) {
        row.appendChild(el('button', {
          type: 'button', title: 'Rename dashboard', 'aria-label': 'Rename dashboard',
          style: {
            border: 'none', background: 'transparent', cursor: 'pointer', padding: '2px 4px', lineHeight: 1,
            color: c.muted, font: `${Math.round(theme.type.h1 * 0.5)}px ${theme.font.family}`, flex: 'none',
          },
          onclick: showEdit,
        }, '✎'));
      }
      titleCol.appendChild(row);
      const sub = subtitleEl();
      if (sub) titleCol.appendChild(sub);
    }

    function showEdit() {
      titleCol.innerHTML = '';
      let settled = false;
      const input = el('input', {
        type: 'text', value: w.title || '', 'aria-label': 'Dashboard name',
        style: {
          margin: 0, font: `700 ${theme.type.h1}px ${theme.font.family}`, color: c.ink, letterSpacing: '-0.01em',
          border: `1px solid ${c.accent}`, borderRadius: '6px', padding: '1px 6px', width: '100%', boxSizing: 'border-box',
          background: c.panel,
        },
      });
      function finish(shouldCommit) {
        if (settled) return;
        settled = true;
        if (shouldCommit) onRename(input.value);
        else showDisplay();
      }
      input.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') { e.preventDefault(); finish(true); }
        else if (e.key === 'Escape') { e.preventDefault(); finish(false); }
      });
      input.addEventListener('blur', () => finish(true));
      titleCol.appendChild(input);
      const sub = subtitleEl();
      if (sub) titleCol.appendChild(sub);
      input.focus();
      input.select();
    }

    showDisplay();
    wrap.appendChild(titleCol);
    if (placeButton) { placeButton.style.right = '108px'; wrap.appendChild(placeButton); }
    if (shareButton) wrap.appendChild(shareButton);
    return wrap;
  }

  function renderShareButton(theme, getSpecForShare, columnsByName) {
    const c = theme.color;
    const btn = el('button', {
      type: 'button',
      style: {
        position: 'absolute', right: 0, top: 0, font: `13px ${theme.font.family}`, fontWeight: 600,
        padding: '8px 14px', borderRadius: px(theme.radius.md), border: `1px solid ${c.rule}`,
        background: 'transparent', color: c.ink, cursor: 'pointer',
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

  // Lives only in the dialog's live mode (see mount()'s `onPlaceOnSheet`
  // param) — the dialog has no Office.js access of its own (it's a separate
  // browser context), so this never touches Excel directly. It hands the
  // current, already-filtered/sorted spec to `onPlaceOnSheet`, which the
  // add-in layer wires to message the task pane and actually place it; this
  // module stays host-agnostic like the rest of render/dom.js.
  function renderPlaceButton(theme, getSpecForShare, onPlaceOnSheet) {
    const c = theme.color;
    // Styled as the primary action here (accent-filled), not "Copy image" —
    // this is the one thing that actually saves anything into the workbook;
    // everything else about a live dashboard is ephemeral until this is
    // clicked. See addin/taskpane.html's build-open-hint for the other half
    // of making that obvious (the task pane's own reminder while this
    // window is open).
    const btn = el('button', {
      type: 'button',
      style: {
        position: 'absolute', right: 0, top: 0, font: `13px ${theme.font.family}`, fontWeight: 600,
        padding: '8px 14px', borderRadius: px(theme.radius.md), border: `1px solid ${c.accent}`,
        background: c.accent, color: c.panel, cursor: 'pointer',
      },
    }, 'Place image on sheet');
    btn.addEventListener('click', async () => {
      if (btn.dataset.busy) return;
      btn.dataset.busy = '1';
      const original = btn.textContent;
      btn.textContent = 'Placing…';
      try {
        const result = await onPlaceOnSheet(getSpecForShare());
        btn.textContent = result && result.ok ? 'Placed' : `Could not place${result && result.error ? `: ${result.error}` : ''}`;
      } catch (e) {
        btn.textContent = `Could not place${e && e.message ? `: ${e.message}` : ''}`;
      }
      setTimeout(() => { btn.textContent = original; delete btn.dataset.busy; }, 2200);
    });
    return btn;
  }

  function renderKpi(w, theme) {
    const c = theme.color;
    const isHero = w.variant === 'hero';
    // overflow:hidden here is the actual fix for a long measure name
    // bleeding into the next card — everything below is already sized to
    // this rect, but nothing previously stopped text from spilling past it.
    const wrap = el('div', { style: absRect(w.rect, Object.assign({ overflow: 'hidden' }, isHero ? {} : { borderLeft: `1px solid ${c.rule}`, paddingLeft: px(theme.spacing.md) })) });
    wrap.appendChild(el('p', {
      title: w.label,
      style: {
        margin: 0, font: `${theme.type.kpiLabel}px ${theme.font.family}`, color: c.muted,
        whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis',
      },
    }, w.label));
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
    const AGGREGATION_LABELS = { sum: 'Sum', avg: 'Average', min: 'Min', max: 'Max', count: 'Count' };
    const rowsText = `${w.count.toLocaleString('en-US')} row${w.count === 1 ? '' : 's'}`;
    // A Min/Avg/Max pick from the settings panel looks identical to a plain
    // Sum otherwise — same number formatting, no other visual cue — so a
    // deliberate override gets spelled out; the classified default doesn't
    // need to (nobody picked it, it's just what the measure naturally is).
    const caption = w.approximate
      ? 'Unweighted average — no verified weight base'
      : w.aggregationOverridden && AGGREGATION_LABELS[w.aggregation]
        ? `${AGGREGATION_LABELS[w.aggregation]} · ${rowsText}`
        : rowsText;
    wrap.appendChild(el('p', {
      title: caption,
      style: { margin: 0, font: `${theme.type.kpiLabel - 1.5}px ${theme.font.family}`, color: c.faint, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' },
    }, caption));
    return wrap;
  }

  // Stands in for the whole KPI/chart block when classification found no
  // measure column at all (engine/layout.js#buildSkeleton) — the old
  // behavior was to render nothing there, a dead end that looked like a
  // bug with no way out even though the fix (correcting one column's role)
  // takes seconds. `onOpenMapping`, when given, sends the dialog's current
  // state to the task pane and asks it to reopen the mapping screen
  // (addin/dashboard-dialog.js's onOpenMapping / addin/taskpane.js's
  // handleOpenMappingRequest); omitted in frozen mode and anywhere else
  // reclassifying isn't possible, in which case only the explanation shows.
  // Free tier only — see engine/layout.js's `showWatermark` widgetConfig
  // flag; the widget only exists at all when that's true, so there's no
  // separate check here. pointerEvents:'none' since its rect overlaps the
  // filter bar's own (engine/layout.js reserves the rightmost slice of
  // that same row for it) — it must never swallow a click meant for a
  // filter chip underneath.
  function renderWatermark(w, theme) {
    // 1.5x theme.type.axis — see render/svg.js's renderWatermark comment;
    // engine/layout.js's reserved watermark width is sized for this.
    return el('div', {
      style: absRect(w.rect, {
        display: 'flex', alignItems: 'center', justifyContent: 'flex-end',
        font: `italic ${theme.type.axis * 1.5}px ${theme.font.family}`, color: theme.color.faint,
        pointerEvents: 'none', whiteSpace: 'nowrap', paddingRight: '6px', boxSizing: 'border-box',
      }),
    }, Format.WATERMARK_TEXT);
  }

  function renderEmptyState(w, theme, onOpenMapping) {
    const c = theme.color;
    const wrap = el('div', { style: Object.assign(absRect(w.rect), theme.card.style === 'panel'
      ? { background: c.panel, border: `${theme.card.borderWidth}px solid ${c.panelBorder}`, borderRadius: px(theme.radius.lg), boxSizing: 'border-box' }
      : { border: `${theme.card.borderWidth}px dashed ${c.rule}`, boxSizing: 'border-box' }) });
    const inner = el('div', {
      style: {
        height: '100%', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center',
        textAlign: 'center', padding: px(theme.spacing.lg), gap: px(theme.spacing.sm), boxSizing: 'border-box',
      },
    });
    inner.appendChild(el('p', {
      style: { margin: 0, font: `600 ${theme.type.panelTitle}px ${theme.font.family}`, color: c.ink },
    }, 'No measures to show'));
    inner.appendChild(el('p', {
      style: { margin: 0, font: `${theme.type.subtitle}px ${theme.font.family}`, color: c.muted, maxWidth: '480px' },
    }, "None of this data's columns were classified as a measure, so there's nothing to chart or total. This is usually one column's role guessed wrong — fix it on the mapping screen."));
    if (onOpenMapping) {
      const btn = el('button', {
        type: 'button',
        style: {
          marginTop: px(theme.spacing.xs), font: `13px ${theme.font.family}`, fontWeight: 600,
          padding: '8px 14px', borderRadius: px(theme.radius.md), border: `1px solid ${c.accent}`,
          background: c.accent, color: c.panel, cursor: 'pointer',
        },
      }, 'Review column roles');
      btn.addEventListener('click', async () => {
        if (btn.dataset.busy) return;
        btn.dataset.busy = '1';
        const original = btn.textContent;
        btn.textContent = 'Waiting for role review…';
        btn.disabled = true;
        try {
          const result = await onOpenMapping();
          // On success a fresh DATA message already replaced this whole
          // mount — nothing here still exists to reset. Only undo the busy
          // state on a cancel/failure, where this widget is still live.
          if (!result || !result.ok) {
            btn.textContent = original;
            btn.disabled = false;
            delete btn.dataset.busy;
          }
        } catch (e) {
          btn.textContent = original;
          btn.disabled = false;
          delete btn.dataset.busy;
        }
      });
      inner.appendChild(btn);
    }
    wrap.appendChild(inner);
    return wrap;
  }

  function regionTooltipText(widget, region, theme) {
    const valueStr = Format.formatMeasureValue(widget, region.value, { negativeStyle: theme.negativeStyle });
    const label = region.series ? `${region.series} — ${region.category}` : region.category;
    return `${label}: ${valueStr}`;
  }

  /**
   * Draws the chart exactly as render/svg.js produced it (same markup, same
   * geometry — see Svg.renderChart's doc comment) and lays a transparent
   * interactive layer over it built from the `regions` that call now also
   * returns: one div per bar/point/sector, sized and positioned from
   * `region.rect` (already in the same absolute coordinate space as
   * `w.rect`, so converting to a position *within* this holder is just
   * subtracting the holder's own origin).
   *
   * @param {Function} [onRegionClick] `(dimensionColumnName, categoryValue) => void` — omit in frozen/read-only mode (mountFrozen) so hover/tooltip still work but nothing is clickable, matching that mode's "no source data to refilter" contract.
   */
  function renderChartWidget(w, theme, onRegionClick) {
    const { markup, regions } = Svg.renderChart(w, theme);
    const holder = el('div', { style: absRect(w.rect) });
    holder.innerHTML = `<svg viewBox="${w.rect.x} ${w.rect.y} ${w.rect.w} ${w.rect.h}" width="100%" height="100%">${markup}</svg>`;
    if (!regions || !regions.length) return holder;

    const tooltip = el('div', {
      style: {
        position: 'absolute', display: 'none', pointerEvents: 'none', zIndex: 20, maxWidth: '220px', whiteSpace: 'nowrap',
        background: theme.color.ink, color: theme.color.panel, font: `12px ${theme.font.family}`,
        padding: '5px 9px', borderRadius: px(theme.radius.sm),
      },
    });

    for (const region of regions) {
      const rx = region.rect.x - w.rect.x;
      const ry = region.rect.y - w.rect.y;
      const clickable = !!(region.dimension && onRegionClick);
      const overlay = el('div', {
        'data-hit-region': '1', // marker only, for taskpane.js's debug panel to count — no behavioral effect
        style: {
          position: 'absolute', left: px(rx), top: px(ry), width: px(region.rect.w), height: px(region.rect.h),
          background: theme.color.accent, opacity: 0, boxSizing: 'border-box', border: '1px solid transparent',
          cursor: clickable ? 'pointer' : 'default', transition: 'opacity .1s',
        },
      });
      overlay.addEventListener('mouseenter', () => {
        overlay.style.opacity = '0.16';
        overlay.style.borderColor = theme.color.accent;
        tooltip.textContent = regionTooltipText(w, region, theme);
        tooltip.style.display = 'block';
        const tw = tooltip.offsetWidth || 80;
        const th = tooltip.offsetHeight || 24;
        tooltip.style.left = px(Math.max(0, Math.min(rx + region.rect.w / 2 - tw / 2, w.rect.w - tw)));
        tooltip.style.top = px(Math.max(0, ry - th - 6));
      });
      overlay.addEventListener('mouseleave', () => {
        overlay.style.opacity = '0';
        overlay.style.borderColor = 'transparent';
        tooltip.style.display = 'none';
      });
      if (clickable) overlay.addEventListener('click', () => onRegionClick(region.dimension, region.category));
      holder.appendChild(overlay);
    }
    holder.appendChild(tooltip);
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
   * @param {{title?:string, subtitle?:string, widgetConfig?:object}} [meta] `widgetConfig` — see engine/layout.js#buildSkeleton
   * @param {object} [opts]
   * @param {(spec:{canvas,widgets}) => Promise<{ok:boolean, shapeName?:string, error?:string}>} [opts.onPlaceOnSheet]
   *   omitted entirely in every host except the dialog (addin/dashboard-dialog.js) — see renderPlaceButton's doc comment.
   * @param {() => Promise<{ok:boolean, error?:string}>} [opts.onOpenMapping]
   *   only present when a widgetConfig with zero measures is mounted — see renderEmptyState's doc comment.
   * @param {{activeFilters?:Object<string,string[]>, sort?:object}} [opts.initialState]
   *   seeds `state` instead of the empty defaults — used when remounting to
   *   apply a widgetConfig change, a Refresh, or a Change data range without
   *   losing the filters/sort the user already had (see addin/dashboard-dialog.js).
   *   `activeFilters` here is plain arrays (JSON-shaped), converted to Sets internally.
   * @param {(state:object) => void} [opts.onStateChange] called after every `setState` with the live internal state object — the dialog forwards the relevant bits to the task pane so they survive a dialog close (see addin/dialog-messaging.js's STATE_UPDATE).
   * @returns {{setTheme(theme):void, getState():object, getLayoutSpec():object}}
   */
  function mount(container, analysis, theme, meta, opts) {
    const onPlaceOnSheet = opts && opts.onPlaceOnSheet;
    const onOpenMapping = opts && opts.onOpenMapping;
    const onStateChange = opts && opts.onStateChange;
    const onTitleChange = opts && opts.onTitleChange;
    const initial = (opts && opts.initialState) || {};
    const Engine = window.DashEngine; // browser-global; Node callers pass their own via a future param if ever needed
    const layoutSpec = Engine.buildLayoutSpec(analysis, meta);
    const columnsByName = Engine.Aggregate.byName(analysis.columns);

    const state = {
      theme,
      activeFilters: Object.fromEntries(Object.entries(initial.activeFilters || {}).map(([k, v]) => [k, new Set(v)])), // column -> Set(values)
      sort: initial.sort || null, // {key, dir}
      expandedFilters: new Set(), // column names currently showing all their value chips, not just the first MAX_CHIPS_SHOWN — not persisted, resets on remount
      page: 0,
    };

    function currentWidgets() {
      return Engine.recomputeLayout(analysis, layoutSpec.widgets, { title: meta.title, subtitle: meta.subtitle, activeFilters: state.activeFilters, sort: state.sort });
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
      if (onStateChange) onStateChange(state);
    }

    // A chart region's click hands back the dimension column driving that
    // bar/point/sector and the category it represents — add it to that
    // dimension's active filter set exactly like checking its box in the
    // filter popover would (additive, not a replace/toggle: clicking twice
    // on two different bars filters to both, matching the popover's own
    // multi-select semantics).
    function addToFilter(dimensionColumn, value) {
      const next = new Set(state.activeFilters[dimensionColumn] || []);
      next.add(value);
      setState({ activeFilters: Object.assign({}, state.activeFilters, { [dimensionColumn]: next }) });
    }

    // Edits `meta.title` in place (not just the one rendered widget) so a
    // later Refresh/theme switch — which rebuilds widgets from `meta` via
    // currentWidgets() — doesn't revert the rename. Empty/unchanged input is
    // a silent no-op: just re-renders back to display mode.
    function commitTitle(newTitle) {
      const trimmed = (newTitle || '').trim();
      if (trimmed && trimmed !== meta.title) {
        meta = Object.assign({}, meta, { title: trimmed });
        renderAll();
        onTitleChange(trimmed);
      } else {
        renderAll();
      }
    }

    function renderWidget(w) {
      if (w.type === 'header') {
        const getSpecForShare = () => ({ canvas: layoutSpec.canvas, widgets: currentWidgets() });
        const shareButton = renderShareButton(state.theme, getSpecForShare, columnsByName);
        const placeButton = onPlaceOnSheet ? renderPlaceButton(state.theme, getSpecForShare, onPlaceOnSheet) : null;
        return renderHeader(w, state.theme, shareButton, placeButton, onTitleChange ? commitTitle : null);
      }
      if (w.type === 'filterBar') return renderFilterBar(w);
      if (w.type === 'kpi') return renderKpi(w, state.theme);
      if (w.type === 'table') return renderTable(w);
      if (w.type === 'emptyState') return renderEmptyState(w, state.theme, onOpenMapping);
      if (w.type === 'watermark') return renderWatermark(w, state.theme);
      return renderChartWidget(w, state.theme, addToFilter);
    }

    const MAX_CHIPS_SHOWN = 12;

    // One row per filterable dimension: an "All" chip (active whenever
    // nothing specific is picked; clears the whole set) followed by a chip
    // per distinct value, matching reference/dashboard.html's chip
    // behavior exactly — click toggles membership, multiple values can be
    // active at once, same Set semantics chart-click filtering already
    // uses via addToFilter. `f.values` always covers every value the
    // dimension has (engine/layout.js fills it from the unfiltered row
    // universe on purpose) so picking one never collapses the row down to
    // just itself. Beyond MAX_CHIPS_SHOWN — already ranked by count,
    // descending, see engine/aggregate.js#distinctValues — the rest
    // collapse behind a "+N" chip that expands them in place (mountFrozen
    // keeps its own separate, inert renderReadOnlyFilterBar — this
    // function only ever runs in live mode).
    function renderFilterBar(w) {
      const c = state.theme.color;
      // Palette color, not theme.color.accent — see render/svg.js's
      // renderFilterBar comment; filters should match the chosen palette
      // like chart series do, not stay a fixed brand green. Always index 0
      // (filters keep that slot even under Spectrum's multiHuePerWidget).
      const filterColor = state.theme.chart.seriesColors[0];
      const wrap = el('div', { style: Object.assign(absRect(w.rect), { display: 'flex', flexDirection: 'column', justifyContent: 'center', gap: '6px', overflow: 'hidden' }) });

      function chip(label, isActive, onclick) {
        return el('button', {
          type: 'button',
          class: 'dash-filter-chip',
          style: {
            font: `12px ${state.theme.font.family}`, padding: '6px 13px', borderRadius: px(state.theme.radius.pill),
            border: `1px solid ${isActive ? filterColor : c.rule}`, background: isActive ? filterColor : 'transparent',
            color: isActive ? c.panel : c.ink, cursor: 'pointer',
          },
          onclick,
        }, label);
      }

      for (const f of w.filters) {
        const active = state.activeFilters[f.column] || new Set();
        const expanded = state.expandedFilters.has(f.column);
        const shown = expanded ? f.values : f.values.slice(0, MAX_CHIPS_SHOWN);
        const hiddenCount = f.values.length - shown.length;

        const row = el('div', { style: { display: 'flex', alignItems: 'center', gap: '6px', flexWrap: 'wrap' } });
        row.appendChild(el('span', { style: { font: `11.5px ${state.theme.font.family}`, color: c.muted, flex: 'none' } }, `${f.column}:`));

        row.appendChild(chip('All', active.size === 0, () => {
          setState({ activeFilters: Object.assign({}, state.activeFilters, { [f.column]: new Set() }) });
        }));

        for (const v of shown) {
          const isActive = active.has(v.value);
          row.appendChild(chip(`${v.value} (${v.count})`, isActive, () => {
            const next = new Set(active);
            if (next.has(v.value)) next.delete(v.value);
            else next.add(v.value);
            setState({ activeFilters: Object.assign({}, state.activeFilters, { [f.column]: next }) });
          }));
        }

        if (hiddenCount > 0) {
          row.appendChild(el('button', {
            type: 'button',
            style: { font: `12px ${state.theme.font.family}`, padding: '4px 8px', border: 'none', background: 'transparent', color: c.muted, cursor: 'pointer', textDecoration: 'underline' },
            onclick: () => setState({ expandedFilters: new Set(state.expandedFilters).add(f.column) }),
          }, `+${hiddenCount}`));
        } else if (expanded && f.values.length > MAX_CHIPS_SHOWN) {
          row.appendChild(el('button', {
            type: 'button',
            style: { font: `12px ${state.theme.font.family}`, padding: '4px 8px', border: 'none', background: 'transparent', color: c.muted, cursor: 'pointer', textDecoration: 'underline' },
            onclick: () => { const next = new Set(state.expandedFilters); next.delete(f.column); setState({ expandedFilters: next }); },
          }, 'less'));
        }

        wrap.appendChild(row);
      }

      if (w.overflow && w.overflow.length) {
        wrap.appendChild(el('span', { style: { font: `12.5px ${state.theme.font.family}`, color: c.muted } }, `+${w.overflow.length} more filter${w.overflow.length === 1 ? '' : 's'} not shown`));
      }
      return wrap;
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

      // Cell strings computed once, up front — used both to size columns
      // (the current page is a representative-enough sample of each
      // column's typical length; scanning the full, possibly huge,
      // rowIndices set on every render would cost far more for no real
      // benefit) and to actually fill the body below, so formatCell only
      // runs once per visible cell.
      const cellStrings = pageIdx.map((rowIndex) => w.columns.map((col) => formatCell(col, columnsByName.get(col.name).values[rowIndex], state.theme)));
      const availableWidth = w.rect.w - padPx * 2;
      const colWidths = Format.computeColumnWidths(
        w.columns.map((col, ci) => ({ header: col.name, samples: cellStrings.map((row) => row[ci]) })),
        availableWidth, { fontSize: t.cellSize }
      );
      const tableWidth = colWidths.reduce((a, b) => a + b, 0); // never wider than needed — a table-layout:fixed table declared wider than its own colgroup sum lets the browser redistribute the gap onto columns itself, past our own maxWidth cap

      // A column set whose combined minimum widths exceed the panel scrolls
      // horizontally instead of clipping the rightmost columns outright —
      // table-layout:fixed only respects colWidths if the table itself
      // isn't forced back to 100% of a narrower parent.
      const tableScroll = el('div', { style: { flex: '1 1 auto', overflowX: 'auto', overflowY: 'hidden' } });
      const table = el('table', { style: { borderCollapse: 'collapse', tableLayout: 'fixed', width: px(tableWidth), font: `${t.cellSize}px ${state.theme.font.family}` } });
      const colgroup = el('colgroup');
      colWidths.forEach((cw) => colgroup.appendChild(el('col', { style: { width: px(cw) } })));
      table.appendChild(colgroup);

      const thead = el('thead');
      const headRow = el('tr');
      w.columns.forEach((col, ci) => {
        const isNum = col.role === 'measure';
        const sorted = w.sort.key === col.name;
        const label = col.name + (sorted ? (w.sort.dir === 'asc' ? ' ↑' : ' ↓') : '');
        headRow.appendChild(el('th', {
          title: col.name,
          style: {
            textAlign: isNum ? 'right' : 'left', padding: '4px 10px 8px', fontSize: px(t.headerSize), color: c.muted,
            borderBottom: `1px solid ${c.rule}`, cursor: 'pointer', userSelect: 'none', fontWeight: 650,
            whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis',
          },
          onclick: () => setState({ sort: { key: col.name, dir: sorted && w.sort.dir === 'asc' ? 'desc' : 'asc' }, page: 0 }),
        }, label));
      });
      thead.appendChild(headRow);
      table.appendChild(thead);

      const tbody = el('tbody');
      pageIdx.forEach((rowIndex, ri) => {
        const tr = el('tr');
        w.columns.forEach((col, ci) => {
          const isNum = col.role === 'measure';
          const raw = columnsByName.get(col.name).values[rowIndex];
          const str = cellStrings[ri][ci];
          const isNegative = isNum && raw != null && raw < 0 && state.theme.negativeStyle === 'color';
          tr.appendChild(el('td', {
            title: str,
            style: {
              textAlign: isNum ? 'right' : 'left', padding: '6px 10px', borderBottom: `1px solid ${c.ruleSoft}`,
              color: isNegative ? c.negative : c.ink, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis',
            },
          }, str));
        });
        tbody.appendChild(tr);
      });
      table.appendChild(tbody);
      tableScroll.appendChild(table);
      inner.appendChild(tableScroll);

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
      // No onOpenMapping here — a frozen snapshot has no source data left to
      // reclassify (see this file's header comment); the message alone still
      // explains the blank space honestly instead of just leaving it empty.
      if (w.type === 'emptyState') return renderEmptyState(w, state.theme, null);
      if (w.type === 'watermark') return renderWatermark(w, state.theme);
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

      // See render/dom.js's live renderTable for why widths/strings are
      // computed together, up front, from just the current page.
      const cellStrings = pageRows.map((rowValues) => w.columns.map((col, ci) => formatCell(col, rowValues[ci], state.theme)));
      const availableWidth = w.rect.w - padPx * 2;
      const colWidths = Format.computeColumnWidths(
        w.columns.map((col, ci) => ({ header: col.name, samples: cellStrings.map((row) => row[ci]) })),
        availableWidth, { fontSize: t.cellSize }
      );
      const tableWidth = colWidths.reduce((a, b) => a + b, 0); // never wider than needed — a table-layout:fixed table declared wider than its own colgroup sum lets the browser redistribute the gap onto columns itself, past our own maxWidth cap

      const tableScroll = el('div', { style: { flex: '1 1 auto', overflowX: 'auto', overflowY: 'hidden' } });
      const table = el('table', { style: { borderCollapse: 'collapse', tableLayout: 'fixed', width: px(tableWidth), font: `${t.cellSize}px ${state.theme.font.family}` } });
      const colgroup = el('colgroup');
      colWidths.forEach((cw) => colgroup.appendChild(el('col', { style: { width: px(cw) } })));
      table.appendChild(colgroup);

      const thead = el('thead');
      const headRow = el('tr');
      for (const col of w.columns) {
        const isNum = col.role === 'measure';
        const sorted = state.sort && state.sort.key === col.name;
        headRow.appendChild(el('th', {
          title: col.name,
          style: {
            textAlign: isNum ? 'right' : 'left', padding: '4px 10px 8px', fontSize: px(t.headerSize), color: c.muted,
            borderBottom: `1px solid ${c.rule}`, cursor: 'pointer', userSelect: 'none', fontWeight: 650,
            whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis',
          },
          onclick: () => setState({ sort: { key: col.name, dir: sorted && state.sort.dir === 'asc' ? 'desc' : 'asc' }, page: 0 }),
        }, col.name + (sorted ? (state.sort.dir === 'asc' ? ' ↑' : ' ↓') : '')));
      }
      thead.appendChild(headRow);
      table.appendChild(thead);

      const tbody = el('tbody');
      pageRows.forEach((rowValues, ri) => {
        const tr = el('tr');
        w.columns.forEach((col, ci) => {
          const isNum = col.role === 'measure';
          const raw = rowValues[ci];
          const str = cellStrings[ri][ci];
          const isNegative = isNum && raw != null && raw < 0 && state.theme.negativeStyle === 'color';
          tr.appendChild(el('td', {
            title: str,
            style: {
              textAlign: isNum ? 'right' : 'left', padding: '6px 10px', borderBottom: `1px solid ${c.ruleSoft}`,
              color: isNegative ? c.negative : c.ink, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis',
            },
          }, str));
        });
        tbody.appendChild(tr);
      });
      table.appendChild(tbody);
      tableScroll.appendChild(table);
      inner.appendChild(tableScroll);

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
