# Renderer V3 Oracle Production Audit

Source: https://chatgpt.com/c/69f97112-7940-8323-a35c-529fe2a1967d

Capture date: 2026-05-04

Source zip: /tmp/bilig2-oracle/bilig2-renderer-v3-oracle-9dfa52ae.zip

Audited commit: 9dfa52ae3355b83af9b2c8dfa1d3feb1af5d0a4b

Status: captured from Oracle; local validation and implementation follow-up required before treating any finding as complete.

## Local Validation Summary

The Oracle response agrees that Renderer V3 is the active product renderer path and that legacy V2 product paths appear removed. It also says V3 is not yet production-complete for world-class spreadsheet UX/performance. The actionable blocker classes are:

1. Sheet lifecycle/reorder identity correctness.
2. Cross-tile text overflow correctness.
3. React commits and subscriptions on hot overlay/tile-delta paths.
4. Always-mounted Canvas2D fallback beside TypeGPU.
5. Main-thread local tile materialization in worker-backed remote mode.
6. Full-tile upload bias when dirty spans are narrower.
7. Freeze/far-scroll tile-interest explosion and freeze revision plumbing.
8. Silent worker render-tile failures and weak tile-batch ordering diagnostics.

## Raw Oracle Response

I’ll audit the zip as authoritative, map live Renderer V3/grid/TypeGPU paths, then trace ownership, dirty propagation, tile identity, uploads, React boundaries, and coverage into file-specific findings with proof tests.

---

Initial tracing shows the product grid surface mounts WorkbookPaneRendererV3 directly; I’m now checking whether “resident” code is only V3 caching/windowing or a hidden drawing path.

---

Old V2/resident-scene symbols are not live product paths, but V3 still isn’t a single GPU-owned runtime: Canvas2D remains mounted beside TypeGPU, and viewport/tile bridge updates still involve React commits.

---

Verdict: Renderer V3 is the active product renderer path, and I did not find hidden product imports of the old V2/resident-scene stack. But Renderer V3 is not yet production-complete for world-class spreadsheet UX/performance. The remaining issues are not “rewrite the renderer” issues; they are identity hardening, hot-path ownership, fallback behavior, text overflow correctness, dirty-span precision, and perf gates.
I did not rerun pnpm run ci; this is a source-level production audit of the attached zip.
1. Renderer V3 completeness / hidden V2 paths
I agree with the prior audit on V2 removal.
The mounted product path is V3:
apps/web/src/WorkerWorkbookApp.tsx:231-323 passes app.workerHandle.viewportStore as both engine and renderTileSource.
packages/grid/src/WorkbookView.tsx:339-389 wraps WorkbookGridSurface in the "workbook-grid" React profiler.
packages/grid/src/WorkbookGridSurface.tsx:7 imports WorkbookPaneRendererV3, and WorkbookGridSurface.tsx:274-284 mounts it.
packages/grid/src/__tests__/renderer-v3-import-boundary.test.ts:17-63 asserts old renderer-v2 files are absent, and renderer-v3-import-boundary.test.ts:66-73 scans product grid paths for renderer-v2 imports.
The old files packages/grid/src/renderer-v2, packages/grid/src/gridResidentDataLayer.ts, and packages/grid/src/gridTileResidencyV2.ts are absent from the unpacked tree.
The important caveat: V3 does not mean “the grid hot path is fully runtime-owned.” React and DOM still own several latency-sensitive paths:
WorkbookGridSurface.tsx:23-81 builds the render/interactions pipeline through React hooks.
WorkbookGridSurface.tsx:266-273 keeps a DOM scroll viewport.
WorkbookGridSurface.tsx:285-306 renders a DOM select-all button.
WorkbookGridSurface.tsx:307-324 renders a DOM fill-handle overlay.
WorkbookGridSurface.tsx:327-358 renders context menu and cell editor overlays.
useWorkbookInteractionOverlayState.ts:52-56 subscribes overlay state through useSyncExternalStore, so hover/selection/fill/range-move/header-drag overlay updates can still cause React commits.
useWorkbookRenderTilePanes.ts:72-76 subscribes render-tile bridge state through useSyncExternalStore, so render-tile deltas can still cause React commits.
So: no hidden V2 renderer path found. But V3 is still partly React-owned and still has main-thread/local fallback behavior.
2. Correctness blockers
A. Tile identity is not robust enough
packages/grid/src/renderer-v3/tile-key.ts:5-33 packs only sheetOrdinal, rowTile, colTile, and dprBucket. sheetId is not part of the tile key. That can be acceptable only if every cache boundary rigorously validates sheet identity and ordinals are stable. The current code does not meet that bar.
Specific identity hazards:
packages/grid/src/runtime/gridRenderTilePaneRuntime.ts:325 filters workbook deltas with:
TypeScriptif (batch.sheetId !== input.sheetId && batch.sheetOrdinal !== sheetOrdinal) return
That condition accepts the delta when either sheetId matches or sheetOrdinal matches. For production sheet reorder/rename/drop cases, this is too permissive. It should require coherent identity: when both fields exist, both should match.
gridRenderTilePaneRuntime.ts:382 keys workbook delta sequence state as ${batch.sheetOrdinal}:${source}. It omits sheetId, so stale-drop behavior can leak across sheets if ordinals are reused or temporarily inconsistent.
apps/web/src/projected-tile-scene-store.ts:83-91 deletes cached tiles on structural axis/freeze deltas when tile.coord.sheetId === batch.sheetId || tile.coord.sheetOrdinal === batch.sheetOrdinal. The || has the same overbroad drop risk.
projected-tile-scene-store.ts:173-193 drops sheets using sheetIds.has(...) || sheetOrdinals.has(...), again allowing unrelated tiles sharing an ordinal to be dropped.
apps/web/src/projected-viewport-store.ts:412-414 falls back to { sheetId: 0, sheetOrdinal: 0 } when emitting local optimistic workbook deltas before sheet identity is known. That can mark the wrong sheet dirty or miss the active sheet entirely.
packages/grid/src/runtime/gridRenderTilePaneRuntime.ts:170-174 retains fixed remote panes only by sheetId. Reuse at gridRenderTilePaneRuntime.ts:197-203 does not validate viewport, DPR bucket, freeze state, host size, axis versions, or sheet ordinal. Existing tests cover sheet switching, but not stale retained panes after viewport/DPR/freeze/axis changes.
packages/grid/src/renderer-v3/tile-residency.ts:292-303 computes stale compatibility from sheetOrdinal, dprBucket, axisSeqX, axisSeqY, and freezeSeq, but not rowTile or colTile. tile-residency-v3.test.ts:26-37 explicitly expects a stale-compatible tile at a different row/col to match. gridTileCoordinator.test.ts:93-99 expects missing/dirty visible tiles to become staleHits, not misses. Today this appears to be used for readiness diagnostics rather than draw substitution, but it is a correctness trap and makes diagnostics misleading.
B. Text overflow across tile boundaries is incomplete
packages/grid/src/gridTextScene.ts:119 sets visibleColumnEnd from the current visibleItems. In V3 tile materialization, visibleItems are the cells in one render tile.
gridTextScene.ts:352-395 only lets left-aligned string/error text spill until visibleColumnEnd. That means text near the last visible column of a tile cannot spill into the next tile, even if the next tile’s cells are empty.
gridTextScene.ts:104-117 also clips data-mode text to the tile-local host bounds, and renderer-v3/text-run-buffer.ts:43-46 turns that into tile-local clip rects. Result: long spreadsheet text can disappear or be clipped at tile boundaries.
packages/grid/src/renderer-v3/text-overflow-index.ts helps dirty propagation for overflow sources, but it does not solve cross-tile drawing. It records dependencies after runs exist; it does not synthesize the missing inbound run in the neighboring tile.
This is a production correctness blocker. Spreadsheet text overflow is visually core behavior.
C. Worker/projected tile partial mutations are unsafe
apps/web/src/projected-tile-scene-store.ts:121-135 handles cellRuns mutations by updating version, lastCellRuns, and dirty flags, but it does not update rectInstances, textMetrics, textRuns, or textCount. If the worker ever emits cellRuns, peekTile() can return stale visual content.
This is currently masked if the worker only emits tileReplace, but the code path exists and should either be completed or rejected at runtime.
D. Freeze versioning is not real
apps/web/src/worker-runtime-render-tile-delta.ts:73 hardcodes:
TypeScriptconst freezeVersion = buildFreezeVersion(0, 0)
renderer-v3/local-render-tile-materializer.ts:36 also uses freezeSeq: 0.
But render-tile subscriptions carry freezeSeq through gridRenderTilePaneRuntime.ts:356-360, and TypeGPU resource keys include freezeSeq in typegpu-tile-resource-revisions.ts:20-23 and :31-34.
The versioning model expects freeze-aware tiles; the materializers are not providing that. Structural invalidation may hide this most of the time, but the version contract is inconsistent.
E. Remote tile miss behavior can flicker or draw stale retained panes
packages/grid/src/runtime/gridRenderTilePaneRuntime.ts:639-663 resolves remote tiles all-or-nothing. If one visible tile is missing, it returns retained panes for the entire sheet if available, otherwise it builds local tiles.
That has three failure modes:
One missing remote tile can suppress all available fresh remote tiles.
Retained panes can stay visible even when viewport/DPR/freeze/axis changed, because retention is keyed only by sheetId.
If no retained pane exists, the renderer falls back to synchronous local tile materialization on the main thread.
apps/web/src/projected-viewport-store.ts:274-280 lazily imports the tile scene store asynchronously; during that gap, peekRenderTile() at projected-viewport-store.ts:302-304 can return null, triggering retained/local fallback.
apps/web/src/worker-runtime.ts:703-704 and worker-runtime.ts:728-729 swallow render-tile subscription/materialization errors silently. A broken render-tile path can become flicker or blank/stale output without useful diagnostics.
3. Performance blockers
A. Canvas2D fallback is always mounted beside TypeGPU
packages/grid/src/renderer-v3/WorkbookPaneRendererV3.tsx:87-114 renders both:
WorkbookPaneCanvasFallbackV3 at :89-99
the TypeGPU canvas at :100-113
The fallback is not conditional on WebGPU failure. It subscribes to camera/scroll stores and schedules RAF draws in WorkbookPaneCanvasFallbackV3.tsx:232-259. That means the WebGPU path is doing extra Canvas2D work during scroll/input, and the fallback can mask TypeGPU correctness bugs.
This is one of the highest-value easy fixes.
B. React can still commit during hot interactions
useWorkbookInteractionOverlayState.ts:52-56 uses useSyncExternalStore for overlay runtime state. The runtime emits on hover, selection, fill-preview, range-move, and header-drag updates through runtime/gridInteractionOverlayRuntime.ts:54-76 and :125-148.
WorkbookGridSurface.tsx:180-230 builds dynamicOverlayBuilder from React-visible state, so hot overlay updates can become prop changes and commits.
useWorkbookRenderTilePanes.ts:72-76 subscribes render-tile bridge state through React. Render-tile deltas and fallback invalidations can cause commits before they reach the renderer runtime.
WorkbookView.tsx:339 already profiles "workbook-grid", but there is no hard commit budget test.
C. Scroll handling is partly runtime-owned, but still too synchronous
packages/grid/src/workbookViewportScrollRuntime.ts:201-204 calls syncVisibleRegion() directly on every scroll event. Resize is RAF-throttled at :192-200, but scroll is not.
The normal bucketed path avoids many React commits through shouldCommitWorkbookVisibleRegion() at workbookViewportScrollRuntime.ts:60-76. However, when requiresLiveViewportState is true, it commits on visible-window changes at :66-68. requiresLiveViewportState is set for editing, fill, resize, header drag, etc. via gridInteractionOverlayRuntime.ts:111-120.
So the worst time to commit—during interaction—is exactly when live viewport mode is more likely.
D. Local fallback can synchronously materialize full tiles on the main thread
gridRenderTilePaneRuntime.ts:240-249 sets forceLocalTiles: true for workbook delta damage.
gridRenderTilePaneRuntime.ts:665-683 builds local tiles with buildLocalFixedRenderTiles().
renderer-v3/local-render-tile-materializer.ts:22-46 maps every tile key in the viewport to materializeGridRenderTileV3() synchronously.
renderer-v3/grid-tile-materializer.ts:53-180 materializes a tile by collecting all tile cells, building GPU scene, building text scene, packing rects, and packing text.
That is too much main-thread work for edit/paste/resize paths in a worker-backed renderer.
E. Dirty spans are too conservative; full uploads are common
packages/grid/src/renderer-v3/render-tile-dirty-spans.ts:69-74 falls back to full rect upload when tile.rectCount !== rowCount * colCount.
But packages/grid/src/gridGpuScene.ts:195-227 can add fills, grid lines, booleans, and explicit borders. Therefore rectCount will often not equal one rect per cell. Full rect uploads are common by design.
render-tile-dirty-spans.ts:106-123 falls back to full text span when no text run matches a dirty range. That means changing an empty or non-text cell can still upload all tile text.
typegpu-tile-buffer-pool.ts:562-571 forces full rect buffer writes when dirty spans are empty, full, count changed, or decoration rects exist. Underline/strike decoration creates another common full-write case.
Buffers are reused, but upload volume is still often full-tile.
F. Frozen panes can inflate tile interest massively
runtime/gridViewportResidencyRuntime.ts:172-177 sets renderTileViewport.rowStart = 0 when freezeRows > 0 and colStart = 0 when freezeCols > 0.
With frozen rows/cols and a far-scrolled body, this creates a bounding rectangle from origin to the body viewport. The pane builder later only needs body tiles plus frozen strips, but the upstream tile range can include many unnecessary intermediate tiles.
This is a serious hidden perf cliff for frozen panes.
G. TypeGPU cache eviction is count-based, not byte/pressure-aware
renderer-v3/typegpu-workbook-backend-v3.ts:159-160 calls residency.evictToSize(256). Tile memory varies by DPR, text density, styling, and decorations. The underlying TileResidencyV3 supports byte budgets, but this backend path uses a fixed entry count.
4. What is implemented correctly
Pane placement is mostly correct. renderer-v3/render-tile-pane-builder.ts:14-137 composes body/top/left/corner placements. render-tile-pane-builder.ts:149-166 computes content offsets from tile origin and body reference. The tests at render-tile-pane-builder.test.ts:79-106 verify reused content tiles for frozen placements, and :108-141 verifies frozen placements stay anchored when body reference is scrolled down/right.
Header/data/overlay separation is structurally correct. renderer-v3/header-pane-builder.ts:50-66 builds dedicated header panes. header-pane-builder.ts:68-94 deliberately builds static header GPU state without dynamic hover/selection state. renderer-v3/typegpu-tile-render-pass.ts:57-92 draws data panes first, :94-101 draws headers next, and :103-109 draws overlay last.
Atlas page uploads are implemented correctly in shape. renderer-v3/typegpu-atlas-manager.ts:318-359 tracks dirty atlas pages, and renderer-v3/typegpu-primitives.ts:538-549 uploads dirty pages when possible. Atlas resize still uploads the full atlas at typegpu-primitives.ts:526-535, which is expected.
TypeGPU buffer reuse exists. renderer-v3/typegpu-tile-buffer-pool.ts:584-610 reuses rect/text buffers when existing capacity is sufficient. renderer-v3/gpu-buffer-arena.ts:43-63 reuses free-list buffers by layout/capacity class. Content resources are keyed by tile id via typegpu-tile-buffer-pool.ts:361-363; placement resources are pane-specific via :365-367, which is right for frozen pane reuse.
Runtime draw ownership is partially good. renderer-v3/workbook-pane-renderer-runtime.ts:209-228 subscribes directly to camera and scroll stores and schedules draws outside React. workbook-pane-renderer-runtime.ts:164-201 resolves latest geometry and draws a TypeGPU frame without needing a React render.
5. Missing tests and benchmarks
The repo has useful V3 boundary and unit tests, but it is missing the tests that would prove production UX/perf.
Missing correctness tests:
Cross-tile text overflow: long left-aligned text starting in the last columns of one tile must render into the next tile, update when a blocker cell is edited, and survive horizontal scroll.
Sheet identity fuzzing: sheet reorder, delete, rename, duplicate ordinal reuse, and local optimistic edits before render-tile subscription identity is available.
Remote tile partial miss: one missing tile should not cause all visible remote tiles to be discarded or trigger local full materialization.
Freeze versioning: freeze changes must change tile freeze versions and prevent stale reuse.
Projected cellRuns: either prove it updates tile content correctly, or test that it is rejected/ignored loudly.
Worker render-tile failure: materialization exceptions must report diagnostics and must not silently leave stale/blank output.
Missing perf tests:
React commit budget under scroll, pointer move, selection drag, fill drag, resize drag, edit, and paste. The "workbook-grid" profiler hook exists at WorkbookView.tsx:339, but there is no hard gate.
GPU upload budget for one-cell edit, empty-cell edit, style edit, row/column resize, range paste, and scroll.
Main-thread local materialization budget: worker-backed render path should have localTileBuildCount = 0 during normal edit/scroll.
Canvas fallback budget: when WebGPU is active, Canvas2D fallback draw count should be zero.
Frozen-pane far-scroll tile interest count: tile materialization should scale with visible body tiles plus frozen strips, not with origin-to-body rectangle area.
Buffer churn budget: buffer creates/destroys/reuses from gpu-buffer-arena.ts should be asserted under scroll/edit.
Flicker/screenshot tests: remote tile async startup and rapid scroll should not show blank/stale tiles beyond an allowed placeholder policy.
6. Optional polish
WorkbookGridSurface.tsx:148 names the current geometry value v2Geometry. It is not a V2 renderer path, but it is misleading and should be renamed to gridGeometry or rendererGeometry.
Several diagnostics already exist, but the production gates should expose counters for local fallback builds, full-tile uploads, partial uploads, canvas fallback draws, render-tile misses, stale retained-pane draws, and worker render-tile errors.
Surgical implementation plan, ordered by risk/value
1. Harden sheet/tile identity first
Files:
packages/grid/src/runtime/gridRenderTilePaneRuntime.ts
apps/web/src/projected-tile-scene-store.ts
apps/web/src/projected-viewport-store.ts
packages/grid/src/renderer-v3/tile-residency.ts
Change the identity logic to require coherent identity. If both sheetId and sheetOrdinal are present, both must match. If only one is present, only that one may be used. Do not use sheetId OR sheetOrdinal matching for mutation/drop.
Specific changes:
Replace gridRenderTilePaneRuntime.ts:325 with a shared matchesSheetIdentity() helper.
Change the sequence key at gridRenderTilePaneRuntime.ts:382 to include both sheet id and ordinal, for example ${batch.sheetId ?? 'x'}:${batch.sheetOrdinal ?? 'x'}:${source}.
Change projected-tile-scene-store.ts:87 and :189 from || identity matching to coherent matching.
Change projected-viewport-store.ts:412-414 so optimistic deltas are not emitted with sheet 0 fallback. Either populate sheetIdentitiesByName from workbook metadata earlier, or skip tile-damage emission until identity is known and report a diagnostic counter.
Expand retained-pane identity in gridRenderTilePaneRuntime.ts:170-203 to include sheetId, sheetOrdinal, dprBucket, renderTileViewport, residentViewport, freeze rows/cols, axis versions/freeze seq, and host size. Retention should only be used when all compatibility fields match.
Fix stale-compatible readiness in tile-residency.ts:292-303. Either include rowTile and colTile in compatibility, or rename the behavior so it cannot be interpreted as a drawable stale hit. Prefer classifying dirty exact tiles separately from truly missing tiles.
Proof tests:
packages/grid/src/__tests__/gridRenderTilePaneRuntime.test.ts
Add tests for wrong sheetId/same ordinal ignored, same sheetId/wrong ordinal ignored when both fields exist, sequence keys isolated by sheet id, retained panes invalidated on DPR/freeze/viewport/axis change.
apps/web/src/__tests__/projected-tile-scene-store.test.ts
Add structural/drop tests proving unrelated tiles sharing an ordinal are not deleted.
apps/web/src/__tests__/projected-viewport-store.test.ts
Add optimistic edit-before-render-subscription test proving no { sheetId: 0, sheetOrdinal: 0 } tile damage is emitted.
packages/grid/src/__tests__/tile-residency-v3.test.ts
Change stale-compatible expectations so different row/col is not treated as compatible for a requested tile.
2. Fix cross-tile text overflow
Files:
packages/grid/src/gridTextScene.ts
packages/grid/src/renderer-v3/grid-tile-materializer.ts
packages/grid/src/renderer-v3/text-overflow-index.ts
Surgical approach:
Keep each tile independently drawable, but add inbound spill runs to the target tile. For a tile starting at column C, scan left per row for the nearest left-aligned string/error source whose spill crosses into C, with empty cells between source and target. Emit a cloned text run into the target tile with negative local x, tile-local clipping, and source row/col metadata. The source tile draws its visible portion; the target tile draws the continuation clipped to the target tile. This avoids requiring a larger source tile surface.
Use clip insets so the cloned run clips to [0, tileSurfaceWidth]:
x = sourceWorldX - tileWorldX
clipInsetLeft = max(0, -x)
clipInsetRight = max(0, x + width - tileSurfaceWidth)
Update TextOverflowIndexV3 so dependencies include inbound spill runs and blocker cells in target tiles.
Proof tests:
packages/grid/src/__tests__/grid-tile-materializer.test.ts
Add a tile-boundary case: long text at col 127 spills into empty col 128; tile 1 contains an inbound spill run with source col 127 and tile-local clipping. Add blocker edit at col 128 proving the inbound spill disappears.
e2e/tests/...
Add screenshot test: long A-style text at the last column of a tile remains visible after horizontal scroll and disappears correctly when a blocker cell is edited.
Perf gate: materializing a tile should scan only a bounded left window per visible row, with a test cap.
3. Disable always-on Canvas2D fallback when TypeGPU is active
File:
packages/grid/src/renderer-v3/WorkbookPaneRendererV3.tsx
Change WorkbookPaneCanvasFallbackV3 to mount only when WebGPU is unavailable, or behind an explicit diagnostic/dev flag. The current unconditional mount at WorkbookPaneRendererV3.tsx:87-99 should not run in normal TypeGPU mode.
Proof tests:
packages/grid/src/__tests__/WorkbookPaneRendererV3.test.tsx
Assert fallback canvas is not mounted when TypeGPU surface is ready. Assert it mounts only in forced fallback/no-WebGPU mode.
Perf benchmark:
During WebGPU scroll, canvasFallbackDraws = 0.
4. Remove React commits from hot overlay updates
Files:
packages/grid/src/useWorkbookInteractionOverlayState.ts
packages/grid/src/runtime/gridInteractionOverlayRuntime.ts
packages/grid/src/renderer-v3/workbook-pane-renderer-runtime.ts
packages/grid/src/WorkbookGridSurface.tsx
Keep GridInteractionOverlayRuntime, but stop subscribing hot overlay state into React through useSyncExternalStore for hover/drag/fill/selection rendering. The renderer runtime should subscribe to overlay runtime directly and request draws, the same way it already subscribes to camera/scroll stores at workbook-pane-renderer-runtime.ts:209-228.
React should keep only semantic state that must affect DOM/editor/menu. Visual overlay state should be read at draw time.
Proof tests:
React profiler test using the existing "workbook-grid" profiler: 100 pointer moves over the grid after initial mount should produce 0 grid commits.
Selection/fill/header drag test: grid commits should be bounded to start/end semantic events, not per pointer move.
Renderer runtime unit test: calling GridInteractionOverlayRuntime.setHoverState() schedules a draw without changing React props.
5. Replace all-or-nothing remote miss fallback with per-tile policy
File:
packages/grid/src/runtime/gridRenderTilePaneRuntime.ts
Change resolveTiles() at :639-663 so one missing remote tile does not discard all available remote tiles.
Policy:
For each requested visible tile key, use exact remote tile if available and identity-valid.
If missing and retained compatible tile exists for the same key and compatibility tuple, use retained only for that key.
If missing and no compatible retained tile exists, emit an explicit blank/loading placeholder pane or omit only that tile.
Do not build local tiles in the worker-backed remote path unless an explicit local-render mode is active.
Change noteWorkbookDeltaDamage() at gridRenderTilePaneRuntime.ts:240-249 so remote-source workbook deltas do not set forceLocalTiles: true. They should mark dirty/request remote tile deltas, not force main-thread materialization.
Proof tests:
gridRenderTilePaneRuntime.test.ts
One of four remote tiles missing returns three fresh remote panes and one retained/blank pane, not local full rebuild.
Workbook delta in remote mode does not set forceLocalTiles.
Startup with async tile-scene-store import does not build local tiles if a remote source is configured.
Perf gate:
Worker-backed edit/scroll has localTileBuildCount = 0.
6. Improve dirty spans and partial uploads
Files:
packages/grid/src/renderer-v3/render-tile-dirty-spans.ts
packages/grid/src/gridGpuScene.ts
packages/grid/src/renderer-v3/typegpu-tile-buffer-pool.ts
Short-term surgical fixes:
Change text dirty behavior at render-tile-dirty-spans.ts:123: if dirty metadata exists and no text run matches, return empty text spans, not full text span. Only do this after overflow dependency tests are in place.
Separate stable base gridline geometry from mutable cell decoration/fill/border/boolean geometry, or add per-cell rect span metadata during materialization. The current tile.rectCount !== cellCount fallback at render-tile-dirty-spans.ts:72-74 guarantees too many full rect uploads.
Avoid forcing full rect uploads for text decoration when dirty spans are precise. typegpu-tile-buffer-pool.ts:562-571 currently full-writes if decoration rects exist.
Proof tests:
render-tile-dirty-spans.test.ts
Add empty-cell value edit yields empty text spans.
Styled/boolean/bordered tile yields bounded rect spans, not full tile.
Underline/strike text update writes only affected decoration spans.
TypeGPU buffer test:
One-cell edit writes bounded bytes, for example less than one row’s worth of tile data, not full tile data.
7. Fix freeze versioning and frozen-pane tile interest bloat
Files:
apps/web/src/worker-runtime-render-tile-delta.ts
packages/grid/src/renderer-v3/local-render-tile-materializer.ts
packages/grid/src/runtime/gridViewportResidencyRuntime.ts
Pass real freeze version/seq into materialization instead of buildFreezeVersion(0, 0) at worker-runtime-render-tile-delta.ts:73 and freezeSeq: 0 at local-render-tile-materializer.ts:36.
Change gridViewportResidencyRuntime.ts:172-177 so frozen panes add origin strips, not a full origin-to-body bounding rectangle. The tile interest should be a union:
visible body tile rectangle
top frozen row strip over visible body columns
left frozen column strip over visible body rows
corner frozen tile rectangle
Do not encode that union as one giant viewport.
Proof tests:
worker-runtime-render-tile-delta.test.ts
Freeze change changes tile version.freeze.
gridViewportResidencyRuntime.test.ts
With freezeRows = 1, freezeCols = 1, and body at row 50,000 / col 500, requested tile count is O(body visible tiles + frozen strips), not O(origin-to-body rectangle).
Perf benchmark:
Frozen far-scroll materialized tile count stays bounded.
8. Add worker diagnostics and ordering guards
Files:
apps/web/src/worker-runtime.ts
apps/web/src/projected-tile-scene-store.ts
Do not silently swallow render-tile publish errors at worker-runtime.ts:703-704 and async subscription errors at :728-729. Emit a diagnostic event/counter and force a controlled fallback state.
Improve batch ordering in projected-tile-scene-store.ts:67-79. Dropping only batch.batchId < this.lastBatchId is weak. Include camera seq and subscription/tile-interest seq for scroll-tile batches so older tiles cannot replace newer interest.
Proof tests:
worker-runtime.test.ts
Materialization exception increments/report diagnostic and does not leave a permanent stale frame.
projected-tile-scene-store.test.ts
Older camera seq for same tile cannot replace newer tile for current interest. Equal batch id with older camera seq is rejected or ignored according to explicit policy.
Final production-readiness assessment
Renderer V3 is real and active. V2 product renderer paths appear removed. Pane placement, header/overlay separation, atlas dirty pages, and TypeGPU buffer reuse have a solid foundation.
The remaining blockers to world-class spreadsheet UX/perf are:
identity correctness under sheet lifecycle/reorder
cross-tile text overflow
React commits on hot overlay/tile-delta paths
always-on Canvas2D fallback
main-thread local tile materialization in remote mode
full-tile upload bias from dirty-span limitations
freeze/far-scroll tile interest explosion
silent worker render-tile failures
Fix those surgically with the tests above. After that, the architecture can plausibly support a production-grade spreadsheet renderer; before that, CI passing does not prove the UX/perf target.
