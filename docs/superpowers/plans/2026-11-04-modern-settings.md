# Companion Settings Implementation Plan

> **For agentic workers:** Use superpowers:subagent-driven-development or superpowers:executing-plans. Track each dependency-ordered phase below. Independent UI/backend implementation may proceed against the fixed bridge contract, with separate file ownership and final integration review.

**Goal:** Deliver both supported editors over canonical AHK behavior, including a complete portable Avalonia companion.

**Architecture:** AHK exports metadata and owns draft operations, validation, merge, conflict detection, persistence and runtime actions. A bounded installation-specific local bridge connects Avalonia and the ephemeral offline settings host. Both editors share a draft lease and Apply implementation.

**Tech Stack:** AutoHotkey v2, .NET 10 LTS, stable Avalonia Fluent, Windows Sandbox, self-contained win-x64 publishing.

**Spec:** `docs/superpowers/specs/2026-11-04-modern-settings-design.md` and the human's supplied khivus brief.

## Global constraints

- Preserve local work, existing file/configuration format, unknown fields and all internal-editor workflows.
- The user later authorized final reviewed commits/pushes to S0CT/EVE-X-Preview-but-gayer and a PR against upstream dev; do not publish releases or change application versions.
- Execute project apps/tests only in Windows Sandbox; keep build/download caches within the workspace.
- Both editors support all settings/workflows; ordinary metadata comes from AHK.
- Main executable must remain usable without companion files.

## Review focus

- Two installations or different working directories must never exchange drafts or save to the wrong file.
- Runtime startup during an offline edit must hand over safely, not create competing writers.
- A successful save followed by a lost response must not replay a destructive profile operation.
- Partial/invalid edits, failed disk writes and disconnects must leave an editable recoverable draft.
- Non-ASCII profiles/paths and multi-monitor placement must preserve existing data and geometry.

## Phase 1: Shared contract and settings transaction

Files: `src/SettingsContract.ahk`, `src/SettingsSession.ahk`, `src/SettingsModel.ahk`, `src/SettingsSchema.ahk`, `src/SettingsWorkbench.ahk`, `tests/Settings_Contract_Test.ahk`, `docs/settings-bridge.md`.

- [x] Inventory scalar fields, tables, groups, event/sound controls and immediate actions against all workbench sections.
- [x] Write tests for metadata completeness, defaults, inheritance/profile rename/undo round trips, unknown fields, conflicting versus unrelated runtime edits, and failed-save rollback. Run in Sandbox to establish failures.
- [x] Export contract and implement shared validation, specialized model operations and shared Apply transaction. Wire built-in Apply through it.
- [x] Run focused and existing Sandbox suites; compare with the three baseline failures.

## Phase 2: Communication, lease and standalone host

Files: `src/SettingsBridge.ahk`, `SettingsHost.ahk`, `Main.ahk`, `src/Main_Class.ahk`, `tests/Settings_Bridge_Test.ahk`.

- [x] Freeze JSON request/response, metadata and snapshot interfaces in `docs/settings-bridge.md` before dependent work.
- [x] Implement bounded, identity-checked local RPC; incompatible protocol, oversized/truncated input, stale session and malformed JSON return structured errors.
- [x] Add lease acquisition/release, draft recovery, canonical offline host, runtime-start handoff and load/save serialization.
- [x] Verify cross-installation isolation, availability/compatibility, concurrent writers, disconnect/reconnect, Apply/restart acknowledgement and clean host shutdown in Sandbox.

## Phase 3: Editor selection and tray routing

Files: `src/SettingsEditor.ahk`, `src/TrayMenu.ahk`, `src/Settings_Gui.ahk`, `Lib/DefaultJSON.ahk`, relevant schema/model files and `tests/Settings_Editor_Test.ahk`.

- [x] Add global selector and tests for all three preferences, bounded probing, missing/incompatible/failing executables, existing editor activation, duplicate prevention and fallback.
- [x] Route Open and double-click; add unconditional direct internal entry. Apply the shared lease protection to profile changes/reload/exit.
- [x] Verify source/compiled relative paths and discoverable fallback diagnostics in Sandbox.

## Phase 4: Avalonia editor

Files: `settings/EVE-X-Preview.Settings/` (focused transport/client/UI modules), `settings/EVE-X-Preview.Settings.Tests/` and `settings/Directory.Build.props`.

- [x] Add .NET 10 project with pinned stable Avalonia packages. Keep ordinary schema out of C#.
- [x] Implement probe, installation discovery, Win32 RPC, single instance/activation, offline bootstrap, reconnect and error recovery. Add protocol/client tests using fixtures and real guest integration.
- [x] Implement metadata-generated ordinary controls, grouped navigation/search, profile selector, all themes, pending-edit validation, keyboard access and Apply/Discard/close protection.
- [x] Implement complete profiles/inheritance, row/group/event/sound editors, hotkey recording/advanced text/conflicts, undo and immediate runtime/diagnostic/update/preview actions.
- [x] Verify every section/workflow against the inventory; inspect both themes at representative sizes/scales in Sandbox.

## Phase 5: Build/package and end-to-end verification

Files: `tools/Build-Release.ps1`, `tools/Test-SettingsIntegration.ps1`, existing development/settings/tray docs and `docs/modern-settings.md`.

- [x] Publish self-contained win-x64 companion and canonical host; include required native libraries and runtime assets. Preserve optional AHK-only build.
- [x] Verify a clean guest without .NET/AutoHotkey installed can launch the package, use both editors, apply/restart, recover and exit all helper processes.
- [x] Run AHK/C#/real bridge integration tests, source and compiled launch checks; record baseline/regressions separately.
- [x] Measure actual package bytes, startup and process memory; perform visual checks and review all changed files.
- [x] Update documentation and record implemented scope, evidence and limitations.
- [x] Publish reviewed source to the authorized fork and open the upstream dev PR.

Verification evidence and explicit limits are recorded in docs/modern-settings-verification.md. Published from S0CT/feature/avalonia-settings in draft PR https://github.com/khivus/EVE-X-Preview/pull/15 targeting upstream dev.
