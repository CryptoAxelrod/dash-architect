/*
 * Placing a generated dashboard on the sheet as a picture, and getting it
 * back later. Implements the workflow from the brief:
 *
 *   1. render/svg.js -> PNG (via render/share.js#renderPngBase64)
 *   2. worksheet.shapes.addImage(png) — a recipient with no add-in just
 *      sees a normal picture; nothing to install, nothing broken.
 *   3. layoutSpec snapshot + column mapping saved into
 *      context.workbook.settings under a key tied to the picture's name.
 *
 * A recipient *with* the add-in reopens that saved state — via
 * addin/taskpane.js's dashboard list or the best-effort click detection —
 * and gets render/dom.js#mountFrozen fed straight from
 * loadDashboardSettings, with no re-read of the sheet and no
 * re-classification ("без пересчёта").
 *
 * The naming/key helpers below are pure and tested without a host
 * (tests/dashboard-io.test.js); every function that touches
 * `context.workbook` is Office.js and can only be verified inside a real
 * Excel add-in — this environment has no Excel host to run one in. See
 * SPEC.md's addin section for exactly what that leaves unverified.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.DashAddinDashboardIo = factory();
  }
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const CONFIG = {
    settingsPrefix: 'dashArchitect:',
    shapeNamePrefix: 'DashArchitectDashboard_',
    // The picture is inserted at this fraction of the layoutSpec's own
    // pixel canvas (which runs ~1180px wide, ~1600-2400px tall depending on
    // content) so it lands at a reasonable on-sheet size; the user can
    // still drag-resize it same as any picture.
    imageScale: 0.5,
  };

  function makeShapeName(seed) {
    return `${CONFIG.shapeNamePrefix}${seed}`;
  }
  function isDashboardShapeName(name) {
    return typeof name === 'string' && name.indexOf(CONFIG.shapeNamePrefix) === 0;
  }
  function settingsKeyForShape(shapeName) {
    return CONFIG.settingsPrefix + shapeName;
  }
  function isDashboardSettingsKey(key) {
    return typeof key === 'string' && key.indexOf(CONFIG.settingsPrefix) === 0;
  }
  function shapeNameFromSettingsKey(key) {
    return key.slice(CONFIG.settingsPrefix.length);
  }

  /**
   * The payload saved per dashboard: enough to restore it without ever
   * re-reading the sheet. `mapping` is the column -> role decision list
   * (engine/roles.js output) — kept distinct from `layoutSpec` so a future
   * "Generate" on the same range can honor a manually-overridden mapping
   * per CLAUDE.md's "user mapping edits win" rule, independent of whatever
   * this particular generation's frozen layout looks like.
   *
   * `theme`/`palette` ride along for one reason: which chart/KPI exists is
   * already baked into `layoutSpec.widgets` by the snapshot itself, but
   * *color* is resolved from theme+palette at render time
   * (render/dom.js#mountFrozen) — without these, a restored dashboard would
   * silently render in light/ocean regardless of what it looked like
   * when placed.
   */
  /**
   * `source` is the structured {kind, address|id/name, headers?} object
   * (addin/taskpane.js's state.source shape) — only present for a DRAFT
   * (`placed: false`, see saveGeneratedDraft): a placed dashboard restores
   * from its frozen `layoutSpec` and never needs to re-read the sheet, but
   * an unplaced one has no picture to restore FROM, so reopening it means
   * re-reading this exact source live instead (addin/taskpane.js's
   * openUnplacedDraft). `placed` defaults true so every payload saved
   * before this field existed still reads as a normal, already-placed
   * dashboard.
   */
  function buildStoragePayload({ snapshot, mapping, sourceAddress, source, title, theme, palette, placed }) {
    return {
      version: 1, generatedAt: new Date().toISOString(), sourceAddress, source: source || null, title, mapping,
      theme: theme || 'light', palette: palette || 'ocean', layoutSpec: snapshot || null, placed: placed !== false,
    };
  }

  // --- Office.js: everything below this line touches context.workbook ---

  /**
   * Finds a shape named `shapeName` across every sheet and deletes it.
   * @returns {Promise<Excel.Worksheet|null>} the sheet it was on, or null if no such shape exists
   */
  async function deleteExistingShape(context, shapeName) {
    const sheets = context.workbook.worksheets;
    sheets.load('items');
    await context.sync();
    for (const sheet of sheets.items) sheet.shapes.load('items/name');
    await context.sync();
    for (const sheet of sheets.items) {
      const match = sheet.shapes.items.find((s) => s.name === shapeName);
      if (match) {
        match.delete();
        return sheet;
      }
    }
    return null;
  }

  function addDashboardImage(sheet, pngBase64, canvasSize, shapeName) {
    const shape = sheet.shapes.addImage(pngBase64);
    shape.name = shapeName;
    shape.left = 0;
    shape.top = 0;
    shape.width = canvasSize.width * CONFIG.imageScale;
    shape.height = canvasSize.height * CONFIG.imageScale;
    return shape;
  }

  /**
   * Places pngBase64 as a picture named `shapeName`. If a shape with that
   * name already exists (on any sheet), it's deleted first and the
   * replacement goes on that SAME sheet — a re-generate (Refresh + Place
   * again, or a headless list-triggered change-range) updates the existing
   * picture in place instead of leaving a stale duplicate behind. A
   * brand-new name (first placement, including placing a draft for the
   * first time) goes on the active worksheet, same as always. Creates
   * unconditionally — see replaceDashboardImageIfExists for the "only if
   * one's already there" variant rename needs.
   * @param {Excel.RequestContext} context
   * @param {string} pngBase64 bare base64 (no `data:` prefix) — see render/share.js#renderPngBase64
   * @param {{width:number, height:number}} canvasSize the layoutSpec's own canvas size, for on-sheet scaling
   * @param {string} shapeName
   * @returns {Promise<Excel.Shape>}
   */
  async function placeDashboardImage(context, pngBase64, canvasSize, shapeName) {
    const existingSheet = await deleteExistingShape(context, shapeName);
    const targetSheet = existingSheet || context.workbook.worksheets.getActiveWorksheet();
    const shape = addDashboardImage(targetSheet, pngBase64, canvasSize, shapeName);
    await context.sync();
    return shape;
  }

  /**
   * Same replace-in-place as placeDashboardImage, but NEVER creates a new
   * picture — if no shape named `shapeName` exists (deleted, or never
   * placed — a draft), this is a no-op. Renaming a dashboard must not
   * resurrect a picture that isn't there: it previously reused
   * placeDashboardImage's create-or-replace behavior, which meant renaming
   * something whose picture had been deleted quietly put a brand new one
   * back on the sheet.
   * @returns {Promise<boolean>} whether a picture was found and replaced
   */
  async function replaceDashboardImageIfExists(context, pngBase64, canvasSize, shapeName) {
    const existingSheet = await deleteExistingShape(context, shapeName);
    if (!existingSheet) return false;
    addDashboardImage(existingSheet, pngBase64, canvasSize, shapeName);
    await context.sync();
    return true;
  }

  async function saveDashboardSettings(context, shapeName, payload) {
    context.workbook.settings.add(settingsKeyForShape(shapeName), payload);
    await context.sync();
  }

  /** @returns {Promise<Array<{shapeName:string, payload:object}>>} every dashboard saved in this workbook */
  async function listDashboards(context) {
    const settings = context.workbook.settings;
    settings.load('items/key, items/value');
    await context.sync();
    return settings.items
      .filter((s) => isDashboardSettingsKey(s.key))
      .map((s) => ({ shapeName: shapeNameFromSettingsKey(s.key), payload: s.value }));
  }

  /** @returns {Promise<object|null>} the saved payload, or null if this shape has none (e.g. deleted) */
  async function loadDashboardSettings(context, shapeName) {
    const setting = context.workbook.settings.getItemOrNullObject(settingsKeyForShape(shapeName));
    setting.load('value');
    await context.sync();
    return setting.isNullObject ? null : setting.value;
  }

  async function deleteDashboardSettings(context, shapeName) {
    context.workbook.settings.getItemOrNullObject(settingsKeyForShape(shapeName)).delete();
    await context.sync();
  }

  /**
   * Deletes a dashboard entirely: its on-sheet picture, if it has one (a
   * draft — see saveGeneratedDraft — has none, and that's fine), and its
   * settings entry. Used by the sidebar list's delete (×) action.
   */
  async function deleteDashboard(context, shapeName) {
    await deleteExistingShape(context, shapeName);
    context.workbook.settings.getItemOrNullObject(settingsKeyForShape(shapeName)).delete();
    await context.sync();
  }

  /**
   * Counts dashboard-named shapes (any sheet) that have no matching
   * workbook.settings entry — the signature left behind when a workbook was
   * closed without saving after "Place on sheet": the picture itself is a
   * normal shape, already part of the file the instant it was inserted, but
   * the settings write that went with it in the same generateAndPlace call
   * only ever lived in the in-memory session and never reached the file.
   * Used to tell that apart from "nothing was ever generated here" — see
   * addin/taskpane.js#showList.
   * @param {Excel.RequestContext} context
   * @returns {Promise<number>}
   */
  async function countOrphanedDashboardShapes(context) {
    const settings = context.workbook.settings;
    settings.load('items/key');
    const worksheets = context.workbook.worksheets;
    worksheets.load('items');
    await context.sync();

    const shapeCollections = worksheets.items.map((sheet) => {
      const shapes = sheet.shapes;
      shapes.load('items/name');
      return shapes;
    });
    await context.sync();

    const savedShapeNames = new Set(
      settings.items.filter((s) => isDashboardSettingsKey(s.key)).map((s) => shapeNameFromSettingsKey(s.key))
    );

    let orphanCount = 0;
    for (const shapes of shapeCollections) {
      for (const shape of shapes.items) {
        if (isDashboardShapeName(shape.name) && !savedShapeNames.has(shape.name)) orphanCount++;
      }
    }
    return orphanCount;
  }

  /**
   * Full "generate" step: place the image and save its settings in one
   * call. `render` is injected (render/share.js#renderPngBase64) rather
   * than required directly, so this file stays free of a hard dependency
   * on the render layer's browser-only APIs (canvas/Image) — keeps the
   * Office.js orchestration and the rasterization concern separately
   * testable.
   *
   * `args.shapeName`, when given, reuses that identity instead of minting a
   * new one — placeDashboardImage then updates the existing picture in
   * place rather than leaving a duplicate. Used for: placing a dashboard
   * that already exists as an unplaced draft (see saveGeneratedDraft), a
   * plain re-place of an already-placed one (Refresh, then Place again),
   * and the task pane's list-triggered rename/change-range.
   *
   * @param {Excel.RequestContext} context
   * @param {{pngBase64:string, canvasSize:object, snapshot:object, mapping:Array, sourceAddress:string, title:string, theme?:string, palette?:string, shapeName?:string}} args
   * @returns {Promise<{shapeName:string}>}
   */
  async function generateAndPlace(context, args) {
    const shapeName = args.shapeName || makeShapeName(Date.now());
    await placeDashboardImage(context, args.pngBase64, args.canvasSize, shapeName);
    // `source` (the structured range/table reference, not just its display
    // string) is what lets this dashboard be reopened LIVE later instead of
    // only as a frozen picture — see saveGeneratedDraft's doc comment and
    // addin/taskpane.js#openDashboardLive.
    const payload = buildStoragePayload({ snapshot: args.snapshot, mapping: args.mapping, sourceAddress: args.sourceAddress, source: args.source, title: args.title, theme: args.theme, palette: args.palette });
    await saveDashboardSettings(context, shapeName, payload);
    // `saveDashboardSettings` resolving without throwing only means
    // Excel.run's batch didn't reject — it's not proof the item is actually
    // there to read back. Read it back in the same call rather than trust
    // that: a "the picture is on the sheet but the list stays empty, no
    // error anywhere" report is exactly what a write that silently didn't
    // stick would look like, and this turns that into a visible, specific
    // failure right here instead of a mystery days later.
    const verify = await loadDashboardSettings(context, shapeName);
    if (!verify) {
      throw new Error(`Settings for "${shapeName}" did not persist — the picture is on the sheet, but nothing was saved to reopen it from the list.`);
    }
    return { shapeName };
  }

  /**
   * Saves just enough to reopen a dashboard live and show it in the
   * sidebar's list — no picture on the sheet yet, `placed: false`. Called
   * right when "Generate dashboard" is clicked, so a mapping/title/theme
   * choice already made isn't lost if the pane closes before "Place image
   * on sheet" — but nothing is written to the sheet itself until that
   * separate, deliberate action happens (addin/taskpane.js#finalizeGenerate
   * vs #placeOnSheet — Generate and Place stay two different actions).
   * @param {string} shapeName minted client-side (makeShapeName) — this same id is reused when it's actually placed later
   * @param {{mapping:Array, sourceAddress:string, source:object, title:string, theme?:string, palette?:string}} args
   */
  async function saveGeneratedDraft(context, shapeName, args) {
    const payload = buildStoragePayload({
      mapping: args.mapping, sourceAddress: args.sourceAddress, source: args.source,
      title: args.title, theme: args.theme, palette: args.palette, placed: false,
    });
    await saveDashboardSettings(context, shapeName, payload);
  }

  /**
   * Renames a dashboard: always updates the stored title AND (when given)
   * the stored frozen layoutSpec's own title, so a later frozen reopen
   * (addin/taskpane.js#openRestoredDashboard — used for a dashboard placed
   * before source-saving existed) shows the new name too, not just the
   * list. Only touches the actual on-sheet PICTURE (if `snapshot.pngBase64`/
   * `canvasSize` are given) when one already exists under this name — see
   * replaceDashboardImageIfExists. A draft (see saveGeneratedDraft) has no
   * picture or snapshot yet, so its rename passes no `snapshot` at all; a
   * placed dashboard whose picture was deleted separately just gets its
   * title corrected, same as a draft, rather than getting a picture put
   * back that the user removed.
   * @param {{layoutSpec?:object, pngBase64?:string, canvasSize?:{width:number,height:number}}} [snapshot]
   */
  async function renameDashboard(context, shapeName, newTitle, snapshot) {
    const existing = await loadDashboardSettings(context, shapeName);
    if (!existing) throw new Error(`Settings for "${shapeName}" not found — it may have been deleted.`);
    const updated = Object.assign({}, existing, { title: newTitle });
    if (snapshot && snapshot.layoutSpec) updated.layoutSpec = snapshot.layoutSpec;
    await saveDashboardSettings(context, shapeName, updated);
    if (snapshot && snapshot.pngBase64 && snapshot.canvasSize) {
      await replaceDashboardImageIfExists(context, snapshot.pngBase64, snapshot.canvasSize, shapeName);
    }
  }

  return {
    CONFIG,
    makeShapeName,
    isDashboardShapeName,
    settingsKeyForShape,
    isDashboardSettingsKey,
    shapeNameFromSettingsKey,
    buildStoragePayload,
    placeDashboardImage,
    replaceDashboardImageIfExists,
    saveDashboardSettings,
    listDashboards,
    loadDashboardSettings,
    deleteDashboardSettings,
    deleteDashboard,
    countOrphanedDashboardShapes,
    generateAndPlace,
    saveGeneratedDraft,
    renameDashboard,
  };
});
