# Companion settings verification — 2026-11-04

All application/test execution took place in offline Windows Sandbox using disposable configurations. The host installation was not used for feature verification. Source inspection, dependency downloads, managed compilation and Git operations used workspace-local source/cache paths. Application versions remain unchanged.

## Automated checks

| Check | Result |
| --- | --- |
| Upstream `dev` baseline before feature | 345 passed, 3 failed, 348 total |
| Full AHK suite including new contract/session/bridge/editor/theme tests | 364 passed, 3 failed, 367 total |
| C# protocol/state/pending-edit/clipboard/dialog/restart/lost-acknowledgement assertions | 49 passed |
| Real source offline bridge integration before custom palettes | 14 assertions passed |
| Real source runtime Apply/restart/reconnect before custom palettes | 16 assertions passed |
| Final portable compiled offline bridge, including all four custom choices and saved custom theme | 18 assertions passed |
| Final portable compiled runtime Apply/restart/reconnect and custom theme persistence | 20 assertions passed |
| Offline-to-runtime handoff retaining accepted draft and unfinished text | 5 assertions passed |
| Response-file-write failure after save, exact-ID retry, no repeated canonical save | 4 assertions passed |
| Existing output-directory guard | Rejected nonempty output; existing caller file retained |
| Modern duplicate launch and clean shutdown | Second process exited 0; main editor exited 0; no package helper remained |

The three AHK failures are the same baseline legacy-panel layout assertions:

- `Attack option controls fit the settings panel`
- `System tracking controls fit in Game Logs Monitoring`
- `Logged-in shifting defaults on for older profiles, persists opt-out and sits beside login shifting`

These were observed before the feature and remain unchanged. No newly introduced AHK test fails. Tests include generated-contract completeness, canonical profile/inheritance and rename origins, unknown nested data retention, unrelated runtime edits, conflicts, failed-save rollback, internal/bridge event parity, sender identity, malformed/UTF-8 messages, response replay, retained orphan drafts, editor preference/fallback logic, and palette switching. Preference/failure cases include injected fixtures; real transport, save/restart, offline/compiled launch, response-write failure and process shutdown were exercised separately.

## Visual checks

Generated 240 clean captures across light, dark, The Matrix, Kitsune, Synthwave and Cyberpunk; ten representative sections; 1100×800 and 850×650 windows; 100% and simulated150% raster scales. The capture process asserted that rendering did not stage settings, then released its lease and exited 0. Inspected representative ordinary controls, settings selection, shortcut rows, sounds/events, compact layouts and all four custom palettes. Sample images are in `docs/images/settings-themes/`.

The 150% captures are renderer simulations, not physical monitor-DPI tests. Native AHK palette changes were exercised against real guest controls and checked for immediate repaint, shared colors and unchanged thumbnail data.

## Portable package and measurements

Built using AutoHotkey2.0.19, stable Avalonia12.1.4 and self-contained .NET10.0.14 for Windows x64. The guest had no AutoHotkey or .NET installation under Program Files. Copied build tools remained in the read-only input mapping. Both compiled applications and the temporary canonical host ran from the generated portable package.

One final guest run measured a 1,688ms first-window startup, 153,157,632-byte working set sampled two seconds later, and a 118,307,335-byte uncompressed distribution across224 files. These are single Sandbox observations, not performance guarantees. Debug symbols were excluded; the executable is distributed with its native/runtime dependencies.

## Remaining manual verification

Live EVE thumbnails, actual game-log events/audio hardware, physical multi-monitor layouts, real150%/mixed monitor scaling and live updater handoff were not exercised. Existing thumbnail placement/monitor code was preserved. No release was published or production installation upgraded.


## Local settings UX revision — 2026-11-05

The user authorized tests on the normal desktop using disposable installation/configuration fixtures under `.vs/ui-checks`; Windows Sandbox was unavailable. The working installation configuration was not used.

- Complete AHK fixture suite: 376 passed, zero failed. Includes the corrected runtime-adapter property/method collision, visible offline preview following draft changes, category undo, new-profile baselines, and opaque draft restoration.
- C# state/protocol/clipboard checks: 49 assertions passed.
- Real offline bridge integration: 19 assertions passed, including Graphite and canonical Apply.
- Real Avalonia UI checks: 19 assertions passed. Covers automatic validation, range and related-dimension inline errors, Apply blocking, updates with another category's invalid input, category/last undo, preview commands, search navigation/highlights, insertion rows, automatic table validation, multi-cell selection and bulk deletion.
- Generated 280 fixture captures across seven themes, ten sections, two sizes and 100%/simulated 150% rendering. Inspected representative Graphite, light/dark, shortcut-grid and wrapping-event layouts. Rendering leaves the fixture draft clean and closes the helper.
- Self-contained Windows x64 build: application/editor/host version 1.7.0.3. Changes remain uncommitted; no release was published.

The render-DPI captures do not replace physical mixed-DPI/multi-monitor or live EVE testing.


## Title bar, table navigation and preview follow-up — 2026-11-05

Kept application/editor/host version 1.7.0.3 and all changes uncommitted. No changelog fix entries were added.

- Complete disposable AHK suite: 377 passed, zero failed. The new gesture test moves all preview layers with a 250 ms inherited window delay, checks that movement avoids the cumulative delay and restores the caller's delay setting.
- C# protocol/state/clipboard/color tests: 56 assertions passed. New checks cover single-column newline lists in wider tables, partial-width pastes, excess-column rejection, true black, dark shades and hue restoration.
- Portable compiled-host UI checks: 32 passed. Checks include the compact themed client title bar, Enter navigation, whole-row Delete, numbered gutter dragging, newline-name paste, responsive character appearance at 850 px, integrated color swatches and selecting a dark shade with the real brightness slider.
- Generated and inspected targeted light/Graphite table captures at 850×650 and 1100×800, including 100% and simulated 150% render DPI. The actual OS title caption is replaced by a themed client caption while retaining the OS resize border; this also applies to owned settings dialogs.

Tests used disposable configuration folders under `.vs/table-followup` and the generated package. Physical mixed-DPI drag behavior and live EVE remain separate manual checks.

## Native frame and sound tables follow-up — 2026-11-05

Replaced the custom caption with the native Windows caption/frame, tightened color and numbered-table cells, consolidated sound assignments and playback under Sound library, and added the footer Problems list. A native dark caption is available on Windows 10; exact palette caption colors depend on Windows 11 support.

C# protocol/state/clipboard/color/problem-list checks completed with 61 assertions before the user stopped testing. The full native suite was interrupted; no complete result is claimed for this revision. UI verification and screenshot capture were not run for this revision. The user requested that all testing stop, and disposable test processes were stopped. Testing was subsequently reauthorized explicitly by the user; the following entry records verification for the later layout changes.

## Startup and control layout follow-up — 2026-11-05

The user reauthorized tests with disposable settings. Final checks: 61 C# assertions passed; 15 targeted AHK tests passed (8 canonical, 5 bridge, 1 category undo, 1 metadata); 62 compiled-editor UI checks passed. The final UI run saved an installed font in its disposable configuration and confirmed an empty issue list after sound assignment and volume edits. No production configuration was edited.

UI checks cover the dark native caption before Show, a theme-cache round trip, loading-view handoff, unavailable-font placeholders, selecting/saving the first installed font, Browse/input spacing, centered row headers and shortcut buttons, theme-colored swatches, readable three-digit volume fields at the standard control height, aligned sound settings, and centered protection toggles. Earlier editing, validation, undo, row selection/paste, event/sound and Problems checks also passed.

Inspected light/Graphite captures at 850×650 and 1100×800 (100% and simulated 150% render DPI), plus populated sound tables, protection toggles and the loading view. About follows the native editor’s vertical version/credits/status/action layout. Verification artifacts are under `.vs/layout-polish`; final UI results are in `verified-final/interaction-checks.log`. The full native suite and live EVE/mixed-DPI monitor behavior were not rerun for this follow-up.

The portable package was rebuilt as v1.7.0.3. No commits, pushes, tags, releases or per-fix changelog entries were created.

## Single-file packaging and startup updates — 2026-11-05

The companion is now one self-contained EXE (~102 MiB) with native runtime libraries and the canonical AHK host embedded. Tested a fresh portable folder without source, DLLs, an adjacent host or configuration, then tested the final companion with only its single EXE present. Offline configuration creation/save remained in the explicit installation directory. The tray compatibility probe passed for the bundled EXE.

Checks passed: 63 managed assertions (including readiness retries before lease acquisition), 65 final compiled single-file UI checks, 19 real source-host integration assertions, and targeted native runs of 20 update tests, 8 canonical tests and 3 editor tests. Update selection tests use mocked responses and cover draft/foreign-URL rejection and the internal About download action. UI checks cover the quiet inline stable-release notice, About navigation, both download metadata and larger event color text. No automated check clicked a real browser download or installed a release. Network-failure handling remains quiet and reports status in About.

Final outputs: `.vs/build/v1.7.0.3-single-file-final/EVE-X-Preview.exe`, `EVE-X-Preview-Settings.exe`, `EVE-X-Preview-Updater.exe`, and `EVE-X-Preview-v1.7.0.3-win-x64.zip`. Main/settings versions remain 1.7.0.3; updater retains its independent 1.6.0 version. The ZIP contains exactly the three executables. The intermediate AHK host stays in the build cache and is not a separate distribution dependency. Commit text is saved under `.vs/single-file-followup/COMMIT_MESSAGE.txt`; nothing was committed, pushed or released.


## v1.7.0.4 settings review — 2026-11-06

Implemented direct, staged companion downloads with a first-launch suggestion and launch handoff; shared main/editor/host/updater versions; and transactional main-plus-settings updates that preserve installations without the optional editor. About places the download below pre-release installation and disables it when the latest editor is installed.

The companion now supports category source selectors and scoped Default editing, retains the edited profile through Apply/reconnect, links shortcut warnings to their actual category, counts distinct unapplied values, and explains disconnection with a persistent banner. Reconnection refreshes validation and retries unfinished input. Event options appear below the table in categories and follow event-enabled state. Search has a clear button; numeric fields share 128 px widths (156 px for long ranges), and sliders sit beside their percentage fields. Disabled spinner outlines are removed. Both editors expose the default-hidden Pause Live Previews tray option; recovery remains in Layout and is removed from the tray.

Checks used disposable installation/configuration folders under `.vs/v1704` and ordinary same-user Windows IPC access. No production configuration was edited.

- Full native AHK suite: **382 passed, zero failed**. Includes download placement/version/launch fixtures, optional updater components, failed download and replacement rollback, change counts, inheritance, and native control behavior.
- Managed protocol/state/clipboard/problem navigation checks: **66 assertions passed**.
- Source-host UI checks: **78 passed**. Generated 70 light/Graphite captures across representative pages, 850×650 and 1100×800 sizes, and 100%/simulated 150% rendering. Inspected compact numeric/disabled controls, event options and search clearing.
- Final self-contained settings EXE alone in a fresh folder: **78 UI checks passed**, including embedded-host creation, Apply/profile retention, deliberate host termination, draft recovery and refreshed Problems state. No source, adjacent host or DLLs were supplied to that folder.
- All three release executables report **1.7.0.4**. The portable ZIP contains exactly the main, settings and updater EXEs.

The native rapid-click check now waits up to 200 ms for posted input callbacks to finish rather than assuming a fixed 35 ms completion time. Pointer setup is atomic, and the change counter caches unchanged state to avoid repeated serialization during input.

Final output: `.vs/build/EVE-X-Preview-v1.7.0.4-win-x64.zip`. Live GitHub installation/download, live EVE behavior, and physical mixed-DPI/multi-monitor behavior were not exercised. Nothing was pushed or published.
