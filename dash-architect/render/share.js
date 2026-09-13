/*
 * "Copy image" — renders the current (filtered) dashboard state through
 * render/svg.js to PNG and puts it on the clipboard via
 * navigator.clipboard.write/ClipboardItem, falling back to a file download
 * when that API isn't available. Browser-only (canvas, clipboard, Image).
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory(root.DashRenderSvg);
  } else {
    root.DashRenderShare = factory(root.DashRenderSvg);
  }
})(typeof self !== 'undefined' ? self : this, function (Svg) {
  'use strict';

  function svgToPngBlob(svgMarkup, width, height, scale) {
    return new Promise((resolve, reject) => {
      const blob = new Blob([svgMarkup], { type: 'image/svg+xml;charset=utf-8' });
      const url = URL.createObjectURL(blob);
      const img = new Image();
      img.onload = () => {
        const canvas = document.createElement('canvas');
        canvas.width = Math.round(width * scale);
        canvas.height = Math.round(height * scale);
        const ctx = canvas.getContext('2d');
        ctx.scale(scale, scale);
        ctx.drawImage(img, 0, 0, width, height);
        URL.revokeObjectURL(url);
        canvas.toBlob((pngBlob) => (pngBlob ? resolve(pngBlob) : reject(new Error('canvas.toBlob returned null'))), 'image/png');
      };
      img.onerror = () => {
        URL.revokeObjectURL(url);
        reject(new Error('the SVG snapshot failed to rasterize'));
      };
      img.src = url;
    });
  }

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 2000);
  }

  function blobToBase64(blob) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(String(reader.result).replace(/^data:[^,]*,/, ''));
      reader.onerror = () => reject(new Error('could not read the rendered PNG'));
      reader.readAsDataURL(blob);
    });
  }

  /**
   * Renders the given (already filtered/frozen) layoutSpec to a PNG Blob.
   * The one non-Office.js-specific step addin/dashboard-io.js needs too
   * (placing the image on the sheet wants base64 — see renderPngBase64) —
   * exported so that caller and `copyDashboardImage` below share one
   * rasterization path rather than two slightly different ones.
   *
   * @param {{canvas:{width,height}, widgets:Array}} layoutSpec
   * @param {object} theme
   * @param {Map} columnsByName for the table widget's raw cell values (omit/null for a frozen spec, which needs no lookup)
   * @param {{scale?:number}} [opts]
   * @returns {Promise<Blob>}
   */
  function renderPngBlob(layoutSpec, theme, columnsByName, opts) {
    opts = opts || {};
    const svgMarkup = Svg.renderDashboardSvg(layoutSpec, theme, columnsByName);
    return svgToPngBlob(svgMarkup, layoutSpec.canvas.width, layoutSpec.canvas.height, opts.scale || 2);
  }

  /** Same as renderPngBlob, but returns bare base64 (no `data:` prefix) — what `worksheet.shapes.addImage` expects. */
  async function renderPngBase64(layoutSpec, theme, columnsByName, opts) {
    const blob = await renderPngBlob(layoutSpec, theme, columnsByName, opts);
    return blobToBase64(blob);
  }

  /**
   * @param {{canvas:{width,height}, widgets:Array}} layoutSpec the CURRENT (already filtered) spec
   * @param {object} theme
   * @param {Map} columnsByName for the table widget's raw cell values
   * @param {{scale?:number, filename?:string}} [opts]
   * @returns {Promise<{method:'clipboard'|'download'}>}
   */
  async function copyDashboardImage(layoutSpec, theme, columnsByName, opts) {
    opts = opts || {};
    const png = await renderPngBlob(layoutSpec, theme, columnsByName, opts);

    if (navigator.clipboard && typeof window.ClipboardItem === 'function') {
      try {
        await navigator.clipboard.write([new window.ClipboardItem({ 'image/png': png })]);
        return { method: 'clipboard' };
      } catch (e) {
        // permissions denied, insecure context, etc. — fall through to download
      }
    }
    downloadBlob(png, opts.filename || 'dashboard.png');
    return { method: 'download' };
  }

  return { copyDashboardImage, renderPngBlob, renderPngBase64 };
});
