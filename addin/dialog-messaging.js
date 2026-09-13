/*
 * Wire protocol between the task pane and the dashboard dialog window.
 * Pure JS, no Office.js/DOM — shared by both sides (addin/taskpane.js,
 * addin/dashboard-dialog.js) so the chunking/reassembly logic exists in
 * exactly one place and can never drift between the two ends.
 *
 * Every message — regardless of direction or size — is one or more
 * envelopes of the same shape:
 *   { kind, requestId, seq, total, payload }
 * `payload` is a slice of the JSON-stringified value being sent; `seq`/
 * `total` let the receiving side reassemble slices in any arrival order
 * and multiplex several message kinds over one raw transport callback
 * (each `createChunkReceiver` ignores envelopes whose `kind` isn't its own).
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.DashDialogMessaging = factory();
  }
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const KIND = {
    // dialog -> parent, sent once as soon as the dialog has registered its
    // own message handler — the parent must wait for this before calling
    // messageChild, or an early message is silently dropped.
    READY: 'ready',
    // parent -> dialog: { mode: 'live', analysis, meta } or
    // { mode: 'frozen', layoutSpec, meta }
    DATA: 'data',
    // dialog -> parent: the resolved {canvas, widgets, theme} to place on
    // the sheet (already filtered/sorted — see render/dom.js's getSpecForShare)
    PLACE_ON_SHEET: 'place-on-sheet',
    // parent -> dialog: { ok: true, shapeName } or { ok: false, error }
    PLACE_ON_SHEET_RESULT: 'place-on-sheet-result',
    // dialog -> parent, sent on every filter/sort/theme/widgetConfig change —
    // the parent just remembers the latest one (addin/taskpane.js's
    // state.lastDialogState) so it survives the dialog closing, however it
    // closes. Shape: { activeFilters:Object<string,string[]>, sort, theme, widgetConfig }
    STATE_UPDATE: 'state-update',
    // dialog -> parent: same shape as STATE_UPDATE — "re-read my current
    // source and send me fresh data, preserving (a reconciled version of)
    // the state I'm handing you right now."
    REFRESH_REQUEST: 'refresh-request',
    // parent -> dialog: { ok:true, analysis, meta, seedState } (seedState is
    // just the request's payload echoed back) or { ok:false, error }
    REFRESH_RESULT: 'refresh-result',
    // dialog -> parent: same shape as STATE_UPDATE — "let the user pick a
    // different range/table in the task pane; when they're done, send me
    // fresh data the same way Refresh does."
    CHANGE_RANGE_REQUEST: 'change-range-request',
    // parent -> dialog: { ok:true, analysis, meta, seedState, source } (seedState
    // is the request's payload, only meaningful if the new source's column
    // names are an exact match for the old one — see engine/reconcile.js)
    // or { ok:false, error } (including a plain cancel)
    CHANGE_RANGE_RESULT: 'change-range-result',
  };

  // Conservative default, not a measured failure point of the real Office
  // dialog transport — see SPEC.md's "Dialog messaging" section for why.
  const DEFAULT_CHUNK_SIZE = 700000;

  function chunkString(str, size) {
    const chunks = [];
    for (let i = 0; i < str.length; i += size) chunks.push(str.slice(i, i + size));
    return chunks.length ? chunks : [''];
  }

  function makeRequestId() {
    return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
  }

  // Serializes `value`, splits it into chunkSize-character pieces, and hands
  // each envelope (already JSON-stringified, ready to send as-is) to `send`
  // in order. Returns the requestId used, so a caller that needs to
  // correlate a later reply can reuse it.
  function sendChunked(send, kind, value, chunkSize, requestId) {
    const id = requestId || makeRequestId();
    const json = JSON.stringify(value);
    const chunks = chunkString(json, chunkSize || DEFAULT_CHUNK_SIZE);
    chunks.forEach((payload, seq) => {
      send(JSON.stringify({ kind, requestId: id, seq, total: chunks.length, payload }));
    });
    return id;
  }

  // Returns a `receive(raw)` function: feed it every raw message string this
  // side gets from the transport. It reassembles envelopes matching `kind`
  // (ignoring any other kind multiplexed over the same channel) and calls
  // `onComplete(value, requestId)` once a full message has arrived.
  function createChunkReceiver(kind, onComplete) {
    let requestId = null;
    let parts = null;
    let total = 0;
    return function receive(raw) {
      let envelope;
      try { envelope = JSON.parse(raw); } catch (e) { return false; }
      if (!envelope || envelope.kind !== kind) return false;
      if (requestId !== envelope.requestId) {
        requestId = envelope.requestId;
        parts = [];
        total = envelope.total;
      }
      parts[envelope.seq] = envelope.payload;
      if (parts.filter((p) => p != null).length !== total) return true; // consumed, still waiting on more chunks
      const json = parts.join('');
      const doneRequestId = requestId;
      requestId = null;
      parts = null;
      total = 0;
      onComplete(JSON.parse(json), doneRequestId);
      return true;
    };
  }

  return { KIND, DEFAULT_CHUNK_SIZE, chunkString, makeRequestId, sendChunked, createChunkReceiver };
});
