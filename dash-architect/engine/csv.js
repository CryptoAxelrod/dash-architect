/*
 * Minimal CSV parser. Pure JS, no dependencies, runs in Node or a browser.
 * No Office.js/Excel/DOM references — see CLAUDE.md.
 *
 * Handles quoted fields (with escaped "" inside quotes), \r\n and \n line
 * endings, a leading UTF-8 BOM, and auto-detects the field delimiter
 * (comma vs semicolon — semicolon shows up in locales that use comma as the
 * decimal separator, e.g. fixtures/07_ru_locale_text_numbers.csv).
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.DashEngineCsv = factory();
  }
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  function stripBom(text) {
    return text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
  }

  function detectDelimiter(text) {
    const firstLine = text.split(/\r\n|\r|\n/, 1)[0] || '';
    const counts = { ',': 0, ';': 0 };
    let inQuotes = false;
    for (let i = 0; i < firstLine.length; i++) {
      const ch = firstLine[i];
      if (ch === '"') inQuotes = !inQuotes;
      else if (!inQuotes && (ch === ',' || ch === ';')) counts[ch]++;
    }
    return counts[';'] > counts[','] ? ';' : ',';
  }

  /** @param {string} text raw file contents @returns {{delimiter:string, rows:string[][]}} */
  function parseCsv(text) {
    const clean = stripBom(String(text));
    const delimiter = detectDelimiter(clean);

    const rows = [];
    let row = [];
    let field = '';
    let inQuotes = false;
    let i = 0;
    const n = clean.length;

    function endField() {
      row.push(field);
      field = '';
    }
    function endRow() {
      endField();
      rows.push(row);
      row = [];
    }

    while (i < n) {
      const ch = clean[i];
      if (inQuotes) {
        if (ch === '"') {
          if (clean[i + 1] === '"') {
            field += '"';
            i += 2;
            continue;
          }
          inQuotes = false;
          i++;
          continue;
        }
        field += ch;
        i++;
        continue;
      }
      if (ch === '"') {
        inQuotes = true;
        i++;
        continue;
      }
      if (ch === delimiter) {
        endField();
        i++;
        continue;
      }
      if (ch === '\r') {
        if (clean[i + 1] === '\n') i++;
        endRow();
        i++;
        continue;
      }
      if (ch === '\n') {
        endRow();
        i++;
        continue;
      }
      field += ch;
      i++;
    }
    // last field/row, if the file didn't end with a newline
    if (field.length > 0 || row.length > 0) endRow();

    // drop fully blank trailing rows (trailing newlines in the file)
    while (rows.length && rows[rows.length - 1].every((c) => c === '')) rows.pop();

    return { delimiter, rows };
  }

  return { parseCsv };
});
