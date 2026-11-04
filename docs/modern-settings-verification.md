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
