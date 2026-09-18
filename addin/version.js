/*
 * TEMPORARY — for verifying a fresh deploy actually reached Excel's WebView
 * (see the caching issues discussed while building this add-in). Bump this
 * number on every deploy you want to be able to confirm landed; the visible
 * badges are in addin/taskpane.html's header and addin/dashboard-dialog.html's
 * toolbar. Delete this file and both badges before release — nothing else
 * in the add-in reads DASH_BUILD_VERSION.
 */
(function (root) {
  'use strict';
  root.DASH_BUILD_VERSION = '1.32';
})(typeof self !== 'undefined' ? self : this);
