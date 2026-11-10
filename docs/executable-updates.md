# Executable updates and recovery

Release publishing keeps the existing standalone assets:

- `EVE-X-Preview.exe`
- `EVE-X-Preview-Updater.exe`
- `EVE-X-Preview-Settings.exe` for users who choose the optional editor

No ZIP, release manifest or checksum sidecar is required. `tools/Build-Release.ps1 -OutputDirectory <empty-directory>` emits the release EXEs. `-BuiltInOnly` builds the main application and updater. The settings EXE contains its .NET/Avalonia dependencies and canonical settings host.

## Download selection and validation

The updater downloads the selected version of the main EXE from GitHub Releases. It downloads the matching settings EXE only if settings is installed after recovery. An absent editor stays absent, and the built-in-only path succeeds when no settings asset exists. If an installed editor's matching download is unavailable, neither installed EXE is replaced.

All requested files are staged before replacement. The installer checks their Windows x64 PE structure and normalized file versions. The staged settings executable must also complete its existing `--probe` protocol, checking that its bundled runtime and application can actually load. This catches a truncated .NET bundle even if its PE sections and version resources survived. Validation uses the existing EXEs; release publishers need no additional integrity files.

The current compatible recovery updater is copied into a unique temporary directory for handoff. A downgrade never selects an older installer from the target release. The adjacent recovery updater remains available for future updates and offline recovery. A version-resource marker distinguishes this EXE-based recovery protocol from earlier installers without executing an unknown binary.

## Window handoff and installation

Apply or discard settings edits before updating. The shared handoff requests normal closure of the modern editor and waits for that exact process to exit, including after it releases its settings lease. An unfinished edit or modal dialog prevents the handoff. The old native settings window is hidden before launching the updater and restored if launch fails. Legacy About actions use this same handoff. The offline canonical host remains alive until the detached updater has started.

The installer waits for the installation's main app and editor to exit, then backs up both selected targets. Replacements use an atomically published local journal under `.EVE-X-Preview-update`. Sharing violations restore the previous EXEs. Configuration, sounds, backups, unknown legacy files and the recovery updater are untouched. Renamed main executables are supported.

## Interrupted or offline updates

**Recovery runs before any release download.** The no-argument latest-release path also recovers before querying GitHub. A failed download therefore cannot strand an earlier installation behind its startup guard.

To recover without any network connection, close both applications and run:

```powershell
.\EVE-X-Preview-Updater.exe --recover 'C:\Preview\EVE-X-Preview.exe'
```

The installer validates recovery targets and backup hashes before restoring anything. Missing or damaged backups remain available for inspection, with an error identifying their location. Both apps refuse startup while recovery is incomplete. Do not delete the recovery directory to bypass that check. The zero-byte `.EVE-X-Preview-update.lock` file is retained to serialize updater instances.

To install already downloaded release EXEs offline, supply their directory:

```powershell
.\EVE-X-Preview-Updater.exe 'C:\Preview\EVE-X-Preview.exe' 1.7.0.4 'C:\Downloads\release-exes'
```

The directory needs the main EXE and, when already installed, the matching settings EXE. The same version, startup-probe and recovery checks apply. Running the updater without arguments selects the latest stable version; About also supports an explicit pre-release or downgrade.

## Settings layout

The three earlier legacy layout regressions remain fixed. Descriptions wrap within their panels and subsequent rows follow their measured height; the unsaved logged-in thumbnail label clears its checkbox. Character Group switching hotkeys are presented as **Next → Previous → First**. The modern editor receives this order from the shared AHK service, keeping each binding attached to its original key.

## Sandbox verification

Run application and updater tests in the configured offline Windows Sandbox, using the guest's local filesystem for fixtures:

```powershell
.\tests\Exe_Update_Test.ps1 -OutputDirectory C:\Checks\exe-fixtures
.\tests\Exe_Update_Integration.ps1 -ReleaseDirectory <built-exes> -LegacyArchive <older-portable-build> -ManagedTests <compiled-tests> -OutputDirectory C:\Output\compiled-update
```

The older archive is a disposable integration input, not a new release requirement. The exact GitHub asset selection, network failure after pending recovery, window handoff and shortcut order are covered by the AHK suite. Additional filesystem cases cover optional settings, missing or damaged downloads, rollback, renamed targets, unsafe journals, backup integrity and offline recovery. The compiled integration uses separate local EXEs through the production offline-install path and inspects the real editor's embedded host. Live GitHub downloads and interactive UAC require a separate live check.

Development verification does not publish releases or alter the installed host application. Application and updater versions remain unchanged.

Verified in offline Windows Sandbox on 2026-11-10:

| Check | Result |
| --- | --- |
| Standalone release build | Exactly the main, recovery-updater and optional settings EXEs; no extra release artifacts |
| Filesystem/recovery cases | 14 passed, including the genuine .NET bundle truncated after its PE sections; both installed originals remained intact |
| Recovery before a failed download | AHK regression restored a pending journal, then deliberately failed the next download; old main and cleared startup guard were verified |
| Optional settings | Compiled built-in-only upgrade succeeded with only the main EXE in the source directory and kept the editor absent |
| Compiled upgrade | Actual older portable build updated from 1.7.0.2 to 1.7.0.4 through separate EXEs; personal settings and unrelated file survived; main relaunched |
| Compiled offline recovery | `--recover` restored an interrupted main and relaunched it without network access or release files |
| Managed suite | 73 assertions passed |
| Updated runtime bridge | 22 assertions passed, including Next → Previous → First order |
| Real settings executable and embedded host | 2 inspection assertions passed; editor closed cleanly |
| Window handoff | A delayed editor process released its lease before exiting; handoff waited for its exit, and refused closure while its draft was dirty |
| Startup guards | Both compiled apps refused an incomplete transaction |
| Full AHK suite | 405 passed, 1 failed in each of two connected-desktop runs; the unchanged rapid-cell-switch input test passed when run separately |

The reported recovery-before-download scenario is covered by the pending-journal/network-failure regression. The truncated-bundle case was reproduced against the previous validation before its fix. The settings-order regression also failed before replacing sorted-map enumeration with explicit ordered pairs. The three earlier layout fixes and their common-font checks remain in the AHK suite.

The remaining full-suite failure is `Settings input pulse rapid cell switches always open the clicked editor and ignore stale focus loss`, an intermittent case also observed on the earlier upstream baseline. Its posted-input path sometimes retains a drag state. Neither `SettingsGrid.ahk` nor that test was changed, skipped or weakened. The initial disconnected guest run also lacked its virtual audio endpoint; the results above are from the reconnected desktop, where audio checks passed.
