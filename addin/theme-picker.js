/*
 * Shared {theme, palette} picker UI — used both by the task pane's
 * pre-generate step (addin/taskpane.js) and the dashboard dialog's settings
 * panel (addin/dashboard-dialog.js), so the two pickers can't drift apart.
 *
 * The live preview is built with render/svg.js#renderDashboardSvg against a
 * small fixed sample layoutSpec — the exact renderer the live dashboard and
 * PNG export use, not a hand-drawn approximation, so what's picked here is
 * what the dashboard will actually look like (CLAUDE.md §5: one render path
 * for interactive and static output — a preview is no exception).
 *
 * No state of its own: render(container, deps, state, handlers) always
 * rebuilds `container` from scratch and reports changes back through
 * handlers — the caller owns {theme, palette} and re-renders after a
 * change, the same "no framework, re-render on setState" pattern every
 * other view in this add-in already uses.
 */
(function (root) {
  'use strict';

  const THEME_ORDER = ['light', 'dark', 'report', 'print'];
  const PALETTE_LABELS = { ocean: 'Ocean', meadow: 'Meadow', ochre: 'Ochre', spectrum: 'Spectrum' };

  function sampleLayoutSpec() {
    return {
      canvas: { width: 300, height: 300 },
      widgets: [
        { id: 'header', type: 'header', rect: { x: 16, y: 8, w: 268, h: 50 }, title: 'Revenue', subtitle: '128 rows' },
        {
          id: 'kpi-hero', type: 'kpi', variant: 'hero', rect: { x: 16, y: 64, w: 120, h: 92 },
          label: 'Revenue', value: 128400, count: 128, aggregation: 'sum', cellFormat: 'currency', valueScale: null, approximate: false,
        },
        {
          id: 'kpi-secondary', type: 'kpi', variant: 'secondary', rect: { x: 152, y: 64, w: 132, h: 64 },
          label: 'Orders', value: 842, count: 842, aggregation: 'sum', cellFormat: null, valueScale: null, approximate: false,
        },
        {
          id: 'chart-0', type: 'bar', rect: { x: 16, y: 168, w: 268, h: 122 }, title: 'By region',
          dimensionColumn: 'Region', measureColumn: 'Revenue', cellFormat: 'currency', aggregation: 'sum', valueScale: null,
          bars: [{ category: 'North', value: 52000 }, { category: 'South', value: 38000 }, { category: 'East', value: 29000 }, { category: 'West', value: 19000 }],
        },
      ],
    };
  }

  function buildPreviewSvg(theme) {
    return window.DashRenderSvg.renderDashboardSvg(sampleLayoutSpec(), theme, null);
  }

  function el(tag, attrs, children) {
    const node = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs || {})) {
      if (k === 'class') node.className = v;
      else if (k === 'onclick') node.addEventListener('click', v);
      else node.setAttribute(k, v);
    }
    for (const child of [].concat(children || [])) {
      if (child == null) continue;
      node.appendChild(typeof child === 'string' ? document.createTextNode(child) : child);
    }
    return node;
  }

  function dots(colors) {
    return el('span', { class: 'tp-dots' }, colors.slice(0, 4).map((color) => el('span', { class: 'tp-dot', style: `background:${color}` })));
  }

  /**
   * @param {HTMLElement} container cleared and rebuilt every call
   * @param {{THEMES:object, Palettes:object}} deps render/themes.js's THEMES map, render/palettes.js's module
   * @param {{theme:string, palette:string}} state current selection
   * @param {{onThemeChange(id):void, onPaletteChange(id):void}} handlers
   */
  function render(container, deps, state, handlers) {
    const { THEMES, Palettes } = deps;
    const activeTheme = Palettes.withPalette(THEMES[state.theme], state.palette);
    container.innerHTML = '';

    const themeRow = el('div', { class: 'tp-row' });
    for (const id of THEME_ORDER) {
      const t = THEMES[id];
      const isActive = id === state.theme;
      themeRow.appendChild(el('button', {
        type: 'button', class: 'tp-swatch',
        'aria-pressed': String(isActive),
        style: `background:${t.color.paper}; color:${t.color.ink}; border-color:${isActive ? t.color.accent : 'transparent'}`,
        onclick: () => handlers.onThemeChange(id),
      }, [el('span', { class: 'tp-swatch-bar', style: `background:${t.color.accent}` }), t.name]));
    }

    const paletteRow = el('div', { class: 'tp-row' });
    for (const id of Object.keys(Palettes.PALETTES)) {
      const isActive = id === state.palette;
      paletteRow.appendChild(el('button', {
        type: 'button', class: 'tp-swatch',
        'aria-pressed': String(isActive),
        style: `border-color:${isActive ? activeTheme.color.accent : 'transparent'}`,
        onclick: () => handlers.onPaletteChange(id),
      }, [dots(Palettes.seriesColorsFor(id, state.theme)), PALETTE_LABELS[id] || id]));
    }

    const preview = el('div', { class: 'tp-preview' });
    preview.innerHTML = buildPreviewSvg(activeTheme);

    container.appendChild(el('div', { class: 'tp-section' }, [el('h3', {}, 'Theme'), themeRow]));
    container.appendChild(el('div', { class: 'tp-section' }, [el('h3', {}, 'Palette'), paletteRow]));
    container.appendChild(preview);
  }

  root.DashThemePicker = { render, buildPreviewSvg, THEME_ORDER };
})(typeof self !== 'undefined' ? self : this);
