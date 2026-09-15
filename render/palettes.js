/*
 * Chart series palettes — deliberately separate from render/themes.js.
 * Before this file, every theme's `chart.seriesColors` was a shade of its
 * own accent green, so every chart was green regardless of which theme was
 * picked. A palette is a set of *data-encoding* colors (line/bar fills,
 * donut slices, grouped-bar series); `theme.color.accent` stays what it
 * always was — brand/UI-state color (buttons, active filter chips, hover
 * highlight) — untouched by palette choice.
 *
 * Each palette carries one 6-color list per theme (`light`/`dark`/`report`/
 * `print`) rather than one list recolored on the fly, for the same reason
 * render/themes.js hand-tunes contrast per theme instead of inverting
 * `light`: a color that reads on a near-white ground and one that reads on
 * a near-black one are rarely the same value at a different lightness.
 *
 * Four palettes, four genuinely different dominant hues — not four
 * saturation levels of the same one (an earlier version of this file
 * picked a single amber/azure/teal hue triple and only varied
 * saturation/lightness across "palettes," which meant every chart still
 * read as "that one hue family" regardless of which was picked):
 *   - Ocean:    blue-dominant (hues ~186-237°)
 *   - Meadow:   green-dominant (hues ~117-176°)
 *   - Ochre:    warm amber/rust/terracotta-dominant (hues ~16-53°)
 *   - Spectrum: deliberately multi-hue, for charts with many series where
 *               a single dominant hue would leave categories looking alike
 * None of the four use a pure, fully-saturated spectral hue (Ocean/Meadow/
 * Ochre keep saturation ~0.36-0.60, Spectrum ~0.37-0.52) — a pure "screen
 * green" or "traffic orange" reads as cheap against a light background;
 * every color here is deepened or muted instead.
 *
 * Every palette/theme's first three colors are verified against a standard
 * deuteranopia simulation (Brettel/Viénot-style LMS confusion matrix) to
 * stay distinguishable under red-green color blindness, the most common
 * type — Ocean/Meadow/Ochre lean on this simulation's one genuinely
 * reliable lever, lightness separation (deuteranopia leaves luminance
 * perception intact; it collapses hue judgment on the red-green axis, which
 * is exactly the axis a tight same-family hue triple like "blue vs.
 * blue-violet" mostly differs along and so cannot safely rely on — see the
 * dated comment in the project history for the failed first attempt), with
 * a supporting hue drift of up to ±35° (±22° for Ochre, to stay warm rather
 * than drifting into yellow-green) so the six colors within one palette
 * aren't a flat lightness ramp either. Spectrum verifies the same first-
 * three deuteranopia requirement across its widely-spread hues, plus every
 * *adjacent* pair in all six (the order charts actually render them in) for
 * plain color distance, so two neighbors in a legend never blend together.
 *
 * `print`'s variant is intentionally near-grayscale for all four palettes
 * (a faint hue tint per identity, not a real color) — render/svg.js's print
 * theme already tells series apart by hatch pattern + line dash + marker
 * shape (see render/themes.js's comment on `print`), with color as a
 * secondary, half-tone-safe signal; a palette swap there only nudges
 * overall contrast/darkness, never hue, so it can't undermine the
 * black-and-white-safe design.
 *
 * Pure JS, no dependencies.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.DashRenderPalettes = factory();
  }
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const PALETTES = {
    ocean: {
      id: 'ocean', name: 'Ocean',
      light: ['#1C356D', '#3D52A3', '#329AC7', '#203A6B', '#427EBB', '#264B83'],
      dark: ['#585FD0', '#BDD3E5', '#7194C6', '#B0E0E5', '#577FCB', '#B7C3E8'],
      report: ['#182C5B', '#35488E', '#2C87AF', '#1B305A', '#3A6FA5', '#204070'],
      print: ['#111213', '#3A3D41', '#656A72', '#888E96', '#484C51', '#292B2E'],
    },
    meadow: {
      id: 'meadow', name: 'Meadow',
      light: ['#26513F', '#3C8654', '#469E92', '#1F774F', '#2CA384', '#225E45'],
      dark: ['#56D4C9', '#BEEABC', '#59C375', '#B6DDC4', '#68BB82', '#B0EAD8'],
      report: ['#1B4427', '#22803D', '#389A93', '#1E5435', '#308F64', '#1D7029'],
      print: ['#111312', '#3A413D', '#65726B', '#88968E', '#48514C', '#292E2B'],
    },
    ochre: {
      id: 'ochre', name: 'Ochre',
      light: ['#84552A', '#5B2C1A', '#B98131', '#815126', '#B48F38', '#75552A'],
      dark: ['#D8AB84', '#CD7450', '#EAD2B0', '#D8A77A', '#E8D9B6', '#D0AB7B'],
      report: ['#894227', '#9B8F3C', '#592C17', '#9A683F', '#46371D', '#9A6B34'],
      print: ['#131211', '#413E3A', '#726C65', '#969088', '#514D48', '#2E2C29'],
    },
    spectrum: {
      id: 'spectrum', name: 'Spectrum',
      light: ['#456DB7', '#B66843', '#1E5641', '#A539B4', '#592527', '#4298B3'],
      dark: ['#C1D3EB', '#C87252', '#59C18C', '#E7C8E7', '#CA767B', '#82C3CB'],
      report: ['#3D60A1', '#A05B3B', '#194735', '#91329E', '#4A1E20', '#3A859D'],
      print: ['#121212', '#3D3D3D', '#6B6B6B', '#8F8F8F', '#4C4C4C', '#2B2B2B'],
    },
  };

  const DEFAULT_PALETTE = 'ocean';

  function seriesColorsFor(paletteId, themeId) {
    const palette = PALETTES[paletteId] || PALETTES[DEFAULT_PALETTE];
    return palette[themeId] || palette.light;
  }

  // The one place a theme and a palette actually combine: a shallow copy of
  // `theme` with `chart.seriesColors` swapped for the chosen palette's
  // colors for that theme's id. Every renderer keeps reading
  // `theme.chart.seriesColors` exactly as before — callers that pick a
  // theme (addin/dashboard-dialog.js, render/dom.js callers, the PNG export
  // path) route it through here first instead of teaching render/svg.js or
  // render/dom.js a second "palette" parameter.
  function withPalette(theme, paletteId) {
    if (!paletteId) return theme;
    return Object.assign({}, theme, {
      chart: Object.assign({}, theme.chart, {
        seriesColors: seriesColorsFor(paletteId, theme.id),
        // Only Spectrum is meant to read as "many different hues" ACROSS a
        // dashboard, not just within one multi-series chart (see this
        // file's doc comment on why it exists at all). Ocean/Meadow/Ochre
        // are one cohesive hue family on purpose, so every single-series
        // widget (a plain bar/line chart, and the filter chips) shares
        // seriesColors[0] there — render/svg.js#widgetSeriesColor is the
        // one place this flag is read.
        multiHuePerWidget: paletteId === 'spectrum',
      }),
    });
  }

  return { PALETTES, DEFAULT_PALETTE, seriesColorsFor, withPalette };
});
