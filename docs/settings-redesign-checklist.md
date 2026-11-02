# Settings redesign implementation

The redesign uses the existing native Windows controls and configuration data. Existing profiles, shortcuts and per-character settings must retain their meaning.

- [x] Draft settings transaction: edits remain isolated until Apply & restart; Discard and close protection work; saves preserve unrelated runtime updates.
- [x] Shared validation: accepted values match the UI, units and ranges are visible, invalid inputs block Apply with an actionable message.
- [x] Seven task-based navigation areas, subsections, search, persistent selection and scroll positions, a resizable window bounded by the monitor work area.
- [x] Explicit per-section inheritance and profile management, preserving existing effective values and avoiding accidental profile activation.
- [x] Row-based editors for character shortcuts, custom names/colors, application matching and ordered group membership; duplicate protection and undo for removals.
- [x] Readable keyboard shortcuts, recording/clearing/advanced syntax, meaningful checkbox labels, keyboard-accessible help and dependency explanations.
- [x] Coherent thumbnail appearance/layout/visibility/mouse actions, visual border explanation, automatic placement and size reset actions.
- [x] Monitoring sources/status, grouped event alerts and event-specific conditions, DPS dependencies and independent focused-client sound/visual/DPS preferences.
- [x] Simple per-event sound selection/preview/volume plus explicit shared-library and batch workflows with predictable commit behavior.
- [x] Application tray commands, advanced performance/compatibility, diagnostics, clearly labelled updates/downgrades and switching counter.
- [x] Updated user documentation and changelog; meaningful regression tests; full AutoHotkey suite, script validation and build verification.

Commit policy: existing tooltips committed on `dev` as `a00d206`; redesign work on `codex/settings-ux-redesign` without per-phase commits or pushes.

Verification completed:

- AutoHotkey suite: 262 passed, 0 failed.
- Main script `/Validate`: exit 0.
- Ahk2Exe x64 compilation: passed.
- Native fixture rendering: main forms at 1000×740 and 800×500, event and shortcut dialogs, and a 500×400 event dialog with scrolling.
- Local documentation links and Git whitespace checks: passed.
- Migration checks cover legacy effective inheritance, profile-local character assignments/groups, focused-client policies, profile renaming and runtime position/ID preservation.

Live EVE switching and rendering, physical audio output, real display-scale changes and installer handoff were not exercised by these fixtures.

## Follow-up review

- [x] Resize redraw, persistent navigation expansion, consistent source placement.
- [x] Compact separated setting rows and measured wrapping, including DPS help.
- [x] Shared editable grids with trailing entry rows, cell/range selection, keyboard editing, clipboard and bulk changes.
- [x] Close-protection character exceptions and a shared Quick Group gesture control.
- [x] Slider scale ticks and five-unit dragging with exact numeric entry.
- [x] Version 1.7.0, revised guides/screenshots/changelog, regression and build checks.

Follow-up verification: 270/270 tests passed; Main script validation and Ahk2Exe compilation passed. The built file reports version 1.7.0. Native snapshots cover inline editing/range selection, compact DPS wrapping, shared Quick Group controls, event dialogs, and repeated footer resizing. Test logging retries transient file-sharing failures and headless errors exit instead of opening an unseen dialog.

Final visual follow-up: the 21 workbench/grid tests passed after the last layout adjustments. Thin slider ticks retain their width in child dialogs. The resize renderer confirms a dirty, valid draft keeps Apply enabled after repeated resizing.

## Interaction follow-up (November 1)

- [x] Table clicks, including held clicks, preserve pane scroll; keyboard navigation can reveal its focused row.
- [x] Blank target rows are ignored without warnings; nonblank invalid targets still validate.
- [x] Real mouse drag/range selection works in all grids, including inheritance.
- [x] Gesture/choice cells open a usable dropdown immediately.
- [x] Advanced unlocks the shortcut field inline and preserves invalid drafts without popups, including navigation and Discard.
- [x] Full-height sidebar and consistently disabled inherited labels, inputs, recorders and tables.
- [x] Supported direct/group shortcut overlap is excluded from warnings and respects runtime priority.
- [x] Split inheritance by visible subsection, migrate mixed sources without changing effective values, preserve dormant overrides in a backup, and document the mapping.
- [x] Keep the caret after the first character when starting a cell edit, including reused editors and the entry row.
- [x] Keep table scroll positions and click targets stable when committing changed or unchanged cells in a long list.

Final verification: 280/280 tests passed, including ten interaction regressions. Main script validation and x64 compilation passed; the preview executable reports 1.7.0. Native fixture captures include full-height/compact sidebars, inherited controls, subsection source selection, inline Advanced editing, DPS wrapping and repeated footer resizing. Live EVE shortcut dispatch remains outside fixture coverage.

## Appearance and shared Default editing

- [x] Swatches beside editable table color values, preserving the configured RGB in both themes.
- [x] Shared System/Light/Dark theme selection, immediate preview, native editor styling, theme rollback on Discard and resource cleanup.
- [x] Automatic asynchronous update checks on opening Settings and a sidebar notice for newer stable versions.
- [x] Edit inherited Default values in place, with selected profile retained and prominent DEFAULT notices; return without applying or restarting.
- [x] Cover Default field/table/group/event routing, staged activation, invalid drafts and Discard.

Verification: 288/288 tests passed, including eight appearance/update/Default-editing regressions. The complete suite also checks font reuse across repeated settings windows. Windows system pickers retain their native appearance. Update tests use simulated responses; installer handoff and live EVE remain separate checks.

## Repaint, recording and search follow-up

- [x] Repaint text and buttons in both palettes, including enabled-state changes and dark/light launch sequences, without needing a scroll.
- [x] Highlight literal search words in labels, explanations, dropdowns and table values, including multiline text, while preserving edit values.
- [x] Resolve a recorder's top-level activation window for both child panes and owned dialogs.
- [x] Toggle overlapping drag ranges against their starting selection; keep Shift range extension available.

Six focused regressions exercise native drawing, recording lifecycle and mouse messages. Render fixtures include repeated dark/light switches and highlighted search text.

Verification: 294/294 tests passed. Script validation, compilation and native fixture rendering passed. Recording tests exercise the focus timer with a controlled active-root predicate for main panes and owned dialogs; physical keyboard recording remains a manual check.

## Native dark controls follow-up

- [x] Theme tooltips without taking focus, including palette changes and hiding when Settings loses focus.
- [x] Flatten and recolor input/table borders, dropdown arrows and popups, native scrollbars and their corner, spinner buttons, slider tracks and navigation arrows.
- [x] Preserve native selection, keyboard input, numeric changes, scrolling and accessibility; restore light-mode border styles without changing raw shortcut editability.
- [x] Add pixel and input regression coverage and isolated native captures for dropdowns and tooltips.

Verification: 299/299 tests passed, including five native-control regressions. Main script validation, x64 compilation and native rendering passed. Captures include dark dropdown popups and tooltips. Windows system file/color pickers retain their system appearance; real display-scale changes remain a manual check.

## Scrollbar tracking and dropdown bounds

- [x] Handle dark scrollbar thumb dragging, arrow/page clicks and repetition before Windows enters its light tracking loop.
- [x] Keep pane/table positions at 32-bit precision, finish table cell edits before scrolling, release capture safely and preserve an open dropdown's existing capture.
- [x] Constrain dropdowns to their owning settings/editor window and monitor work area, including compact layouts and opening above a low field.
- [x] Cover held dragging and capture changes, both table axes, dropdown scrolling, and popup bounds in both palettes.

## Resize and button paint lifecycle

- [x] Batch layout/page replacement without global child compositing, revealing hidden windows or leaving redraw disabled.
- [x] Wait for the end of interactive resizing before rebuilding a pane; preserve inline edits during height-only resizes.
- [x] Use one custom button paint path across default, focus, press, hover, disabled and reactivation states; keep native click/dialog behavior.
- [x] Buffer scrollbar parts together and repaint after native size/style changes.
- [x] Add tray-profile activation, repeated resize, palette/state and nested update regressions.

Verification: 312/312 tests passed, including eleven native-control and seven paint-lifecycle regressions. Main script validation, x64 compilation (version 1.7.0) and native fixture rendering passed. Tests cover native drag/capture messages and button/profile state changes; real display-scale and compositor timing remain manual checks. Desktop automation was not used for this follow-up.

## Table and canvas scrolling stability

- [x] Remove whole-window compositing; use native table buffering and a separately buffered header.
- [x] Redraw selection in the table body without clearing its background or repainting the header.
- [x] Avoid extra client-DC/descendant redraws on native table paint and mouse messages; guard nonclient painting against reentrancy.
- [x] Repaint visible canvas children after moving, without copying stale clipped pixels; skip unchanged offsets.
- [x] Observe native paint messages for header isolation, idle rendering, separators and selected-cell scrolling.

Verification: 316/316 tests passed. Four new regressions observe WM_PAINT/WM_ERASEBKGND: selection does not repaint the header or erase the table, painting becomes idle after input, separators repaint after canvas movement, and selection survives held thumb dragging. Validation, x64 compilation, native fixture rendering, documentation links and whitespace checks passed. These replace the previous reliance on composited child painting; live frame timing still needs confirmation on the user's Windows session.

## Header coverage during cell clicks

- [x] Exclude the mapped header rectangle from cell invalidation and clip native table drawing around child windows.
- [x] Record native update bounds: the previous full-client selection invalidation fails this regression, while the corrected body rectangle passes.
- [x] Exercise repeated mouse clicks and inline editor focus in dark/light/dark transitions, preserving headings, data and header layout.

Verification: 317/317 tests passed. The strengthened selection test fails before the fix because the update rectangle includes the header; it passes after clipping the invalidation to the table body. Validation, x64 compilation, native rendering, documentation links and whitespace checks passed.

## First scroll, live field updates and continuous input

- [x] Replace the pane's nonclient scrollbar with a separate control; own its first dark paint, capture, keyboard input and accessible geometry.
- [x] Repaint scalar color swatches after valid typing and picker selection; retain the last valid color during incomplete input.
- [x] Keep the native slider thumb untouched during tracking; update five-step values, snap on release, and preserve precise numeric input.
- [x] Guard unchanged label, hint, button and status assignments instead of repainting them on every input event.
- [x] Ignore stale editor focus-loss callbacks while opening/reusing a cell editor, without losing genuine focus-out commits.
- [x] Calculate the final XOR drag selection before painting and skip unchanged endpoints.
- [x] Disable thousands separators for all workbench UpDown controls, including negative coordinates.
- [x] Add seven input regressions covering the reported sequences.

Verification: 324/324 tests passed. Seven new regressions cover typed/picked swatch pixels, positive/negative four-digit spinner edits, continuous main/event slider tracking, unchanged-help paint counts, rapid editor reuse, stable drag endpoints, and first pane scrolling with synchronized range/accessible geometry. Script validation, x64 compilation (1.7.0), native rendering, documentation links and whitespace checks passed. Desktop automation was not used.

## Slider marks, validation recovery and profile creation

- [x] Draw scale marks inside the slider instead of overlapping text controls.
- [x] Request immediate native paint for scalar swatches and changed hint colors; observe WM_PAINT before any print/scroll/reset.
- [x] Recompute relation errors before refreshing hints. Mark both default/minimum size fields and repaint corrected messages with their current color.
- [x] Avoid rebuilding the current table when locating an invalid row. Guard late focus events, reentrant commits and saves which dispose their own grid.
- [x] Center owned editors, color pickers, changelog and log viewers over their settings owner within its monitor work area.
- [x] Add a saved new-profile inheritance template under Default, three creation presets and individually editable subsection flags. Copy effective values for custom subsections.
- [x] Align installed/stable/pre-release version rows and use consistent widths for stacked actions.
- [x] Add eight regressions, including native paint requests, validation recovery, tick pixels, invalid row details, disposal during commit, template round trips, creation presets and popup geometry.

Verification: 332/332 tests passed; the strengthened eight-test review group also passed after the final layout changes. Main validation, x64 compilation (1.7.0), native fixture rendering and whitespace checks passed. Captures include aligned About rows, the profile creation dialog and Default's inheritance template. Desktop automation was not used; physical monitor/DPI and compositor timing remain manual checks.

## Modifier warnings and remaining native scrollbar flashes

- [x] Distinguish a more-specific modifier chord from its plain-key wildcard fallback without changing registration. Keep duplicate/alias and ambiguous custom-combination warnings.
- [x] Exercise Alt+backtick for login cycling alongside backtick for group switching, with automatic wildcards both on and off.
- [x] Suppress intermediate native drawing locally during dark table/tree scrolling, layout and expansion; include actual sidebar expander-button hits.
- [x] Respect external WM_SETREDRAW locks and hidden windows. Keep normal table-cell selection and light-mode drawing on their existing paths.
- [x] Retain header isolation and selected cells, keep table background clearing inside its native memory buffer, and buffer dark navigation background painting.
- [x] Cover event tables, an owned profile-creation inheritance grid, sidebar wheel/expansion/ensure-visible operations, redraw restoration and horizontal scrolling.
- [x] Restore the editing DEFAULT heading suffix, caught by the full inheritance regression.

Verification: 337/337 tests passed, including five new focused regressions; existing native chrome and scrolling tests also passed. Main validation, x64 compilation (1.7.0), native fixture rendering, local documentation links and whitespace checks passed. The observer verifies redraw suppression below the chrome subclass before the native procedure receives an update. No desktop automation was used. Offscreen fixtures cannot establish compositor timing on every real monitor/DPI setup.

## Pre-release packaging

- [x] Review the latest five GitHub releases and retain the two portable Windows x64 package names.
- [x] Build app v1.7.0 and standalone updater v1.6.0 from this branch; give the updater matching display/file versions and current application/suspend resources.
- [x] Share the window-sizing helper between app and updater so standalone compilation resolves the diagnostic viewer dependency.
- [x] Verify x64 identity, version metadata, embedded changelog and every normal/suspend icon frame against the source files.
- [x] Document the repeatable two-package build and make the stable-only behavior of a standalone updater clear in the release downloads.

Verification: 338/338 tests passed, including updater packaging metadata. Main and updater validation and compilation passed. The app/updater icon groups contain the source's 13 normal frames and 12 suspend frames, including suspend resource IDs 206 and 208. Published-package integrity is checked separately after upload.
