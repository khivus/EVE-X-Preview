# Portable package updates

The updater installs a versioned Windows x64 ZIP as one transaction. It validates the archive checksum, release manifest, file sizes and SHA-256 hashes before changing the installation. The manifest records package-owned files; later updates remove obsolete owned dependencies while preserving configuration, sounds, backups and unrelated files. An installation using only the built-in editor keeps that choice.

## Release assets

`tools/Build-Release.ps1 -OutputDirectory <empty-directory>` now produces the portable files and `EVE-X-Preview.package.json` inside that directory, plus these files beside it:

- `EVE-X-Preview-win-x64.zip`
- `EVE-X-Preview-win-x64.zip.sha256`

Upload both files to the matching numbered GitHub release, alongside the existing standalone application, updater and optional settings executable assets. The full archive is required for in-app updates, including installations that use only the built-in editor. Such installations download the same verified archive but install only the main application and updater. `-BuiltInOnly` creates a separate `EVE-X-Preview-built-in-win-x64.zip` for manual distribution. `-ArchivePath` overrides the output archive location; it must be outside the portable directory.

The current settings executable embeds its runtime, native dependencies and canonical settings host. Those are updated with that executable. The package format also supports loose DLLs, nested native libraries, an adjacent settings host and the companion's dependency/runtime JSON files. Configuration JSON, sounds, logs, backup archives and arbitrary user files cannot be declared as package payloads.

To package an existing build directory again, run:

```powershell
.\tools\New-ReleasePackage.ps1 -Directory <portable-directory> -Version 1.7.0.4 -ArchivePath <outside-directory>\EVE-X-Preview-win-x64.zip
```

The manifest has schema version 1, application `EVE-X-Preview`, architecture `win-x64`, the numeric release version, and a `files` array containing each relative `path`, byte `size` and SHA-256 `sha256`. The archive contains exactly those payload files and the manifest, with no enclosing directory.

## Installation and recovery

The application stages its compatible installed updater in a unique temporary directory. It inspects the updater's version-resource protocol marker without executing an unknown executable. If that updater is missing, recovery downloads the updater for the **installed** version. Selecting a downgrade does not select an older installer.

The standalone updater accepts `<main-executable> <numeric-version>`, or no arguments to select the latest stable release. An optional third argument is a local package ZIP for an offline installation; its `.sha256` sidecar is still required. For example:

```powershell
.\EVE-X-Preview-Updater.exe 'C:\Preview\EVE-X-Preview.exe' 1.7.0.4 'C:\Downloads\EVE-X-Preview-win-x64.zip'
```

The updater waits for the main app, settings editor and adjacent host to exit normally. It stages files under the installation, rejects unsafe or duplicate paths and reparse points, checks the version and every file hash, and backs up all files before replacement. The manifest is installed last. Renamed main executables are supported. The existing configuration and unowned legacy files survive; only obsolete files recorded by the previous manifest are eligible for removal. Modified obsolete files cause a refusal requiring the user to preserve or remove them explicitly.

Replacement failures restore the previous files and remove newly introduced files. The reserved `.EVE-X-Preview-update` directory holds an atomically published recovery journal and backups. After process interruption, rerunning the updater recovers that transaction before trying another package. If restoration cannot finish, backups remain there and the error identifies their location. The main app and settings app refuse to start while that directory exists. Do not delete it to bypass recovery. The zero-byte `.EVE-X-Preview-update.lock` file is retained to serialize updater instances.

Old releases without the ZIP/checksum/manifest format cannot be installed through this updater. The attempt fails without falling back to an EXE-only replacement. Install an older full portable package manually if needed.

## Verification

Run these tests inside the configured offline Windows Sandbox, using its local filesystem for fixtures:

```powershell
.\tests\Package_Update_Test.ps1 -OutputDirectory C:\Checks\package-fixtures
.\tests\Package_Update_Integration.ps1 -PackageDirectory <built-package> -ArchivePath <release-zip> -LegacyArchive <older-portable-zip> -ManagedTests <compiled-managed-tests> -OutputDirectory C:\Output\compiled-update
```

The integration test requires an actual older portable package and uses the compiled production updater's offline-install option. Source and managed suites retain their existing entry points. See [development](development.md) for the Sandbox runner.

Development verification does not publish releases or modify the installed host application. The application and updater versions are unchanged by this change.

Verified on 2026-11-10 in an offline Windows Sandbox with read-only source/build inputs and dedicated writable output:

| Check | Result |
| --- | --- |
| Real filesystem package tests | 17 passed, including dependencies, nested libraries, config/user-file preservation, obsolete owned-file cleanup, hashes, unsafe/duplicate paths, junctions, locked-file rollback, renamed main, staging/interruption recovery and repackaging |
| Installer termination | Stopped a real PowerShell installer after observing main-file replacement; next run restored the previous main and removed newly introduced files |
| AHK suite on final source | 400 passed, 3 pre-existing layout failures, 403 total |
| Managed protocol/state suite | 73 assertions passed |
| Compiled updater | Upgraded the actual older portable build from 1.7.0.2 to 1.7.0.4, relocated the installed updater, replaced both applications and relaunched main; personal settings marker and unrelated file survived |
| Updated runtime | 21 real bridge assertions passed, including Apply/restart and reconnect |
| Updated production settings executable | Started the matching embedded host; 2 read-only hello/metadata assertions passed without taking the editor lease; editor closed with exit code 0 |
| Incomplete package startup | Compiled main displayed its recovery notice and exited 1; compiled settings exited 1 |
| Offline editor closure during update | Actual host lifetime check failed before the handoff fix and passed afterward with a signalled parent handle |
| Updater selection | Inspected the compiled protocol marker and staged the compatible installed updater without contacting a deliberately nonexistent target release |

The three unchanged AHK failures are **Attack option controls fit the settings panel**, **System tracking controls fit in Game Logs Monitoring**, and **Logged-in shifting defaults on for older profiles, persists opt-out and sits beside login shifting**. The first upstream baseline run also failed **Settings input pulse rapid cell switches always open the clicked editor and ignore stale focus loss**; it passed in both subsequent full runs. Baseline was 399 passed/4 failed; both updated full runs were 400 passed/3 failed.

The compiled updater was exercised with a locally staged release ZIP through its production offline-install option. Live GitHub downloads, release publication and the interactive install-button/UAC flow were not exercised. All application, updater, host and test processes ran in Sandbox.
