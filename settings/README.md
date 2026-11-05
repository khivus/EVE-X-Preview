# Modern settings companion

The Windows editor uses .NET 10 LTS and Avalonia Fluent 12.1.4. It consumes the AHK settings contract at runtime. It contains no ordinary-settings schema and does not read or overwrite the application's configuration file.

Projects:

- `EVE-X-Preview.Settings/EVE-X-Preview.Settings.csproj`: desktop executable, `EVE-X-Preview-Settings.exe`.
- `EVE-X-Preview.Settings.Tests/EVE-X-Preview.Settings.Tests.csproj`: dependency-free console checks linking the production transport/state code.

Build with the .NET 10 SDK. In this repository, keep all SDK/package/temp caches under `.vs/build`, set `DOTNET_GENERATE_ASPNET_CERTIFICATE=false`, `DOTNET_ADD_GLOBAL_TOOLS_TO_PATH=0`, `DOTNET_CLI_TELEMETRY_OPTOUT=1`, and execute applications/tests only in an isolated Windows Sandbox. `NuGet.Config` affects only these projects. Publishing is handled by the repository release tool.

```powershell
dotnet restore settings/EVE-X-Preview.Settings/EVE-X-Preview.Settings.csproj -r win-x64
dotnet build settings/EVE-X-Preview.Settings/EVE-X-Preview.Settings.csproj -c Release -r win-x64 --no-restore
dotnet publish settings/EVE-X-Preview.Settings.Tests/EVE-X-Preview.Settings.Tests.csproj -c Release -r win-x64 --self-contained true -o .vs/build/settings-tests
```

Launch the packaged editor beside EVE-X-Preview, or pass `--installation <directory>`. A direct launch discovers that installation's runtime. If absent, it starts the canonical packaged settings host or `SettingsHost.ahk` with an AutoHotkey v2 interpreter beside the script or in a standard per-machine/per-user location. Published builds embed that canonical host and extract a verified copy into a per-user cache, passing the installation folder explicitly. It never starts the normal thumbnail/hotkey engine. The host watches the editor PID and exits when the editor closes.

`--probe <response-file>` writes only the small compatibility response and exits without Avalonia, IPC, or an editing lease. A named per-user/per-installation gate activates an already-open editor. The bridge uses bounded WM_COPYDATA requests from a real hidden HWND, installation-specific endpoints and matching response files. Detailed protocol: `../docs/settings-bridge.md`.

Ordinary controls are generated from field metadata, including choices, toggles, sliders, numeric text, colors, fonts, folders, multiline text and recorded/advanced shortcuts. Dedicated controls cover profiles, inheritance templates, ordered row tables and clipboard exchange, groups and shortcuts, event conditions, sound library/gain/assignments, running-window selection, diagnostics, updates and runtime previews. Validation, undo, save, conflicts and runtime behavior remain authoritative in AHK.

Typing validates automatically against the canonical settings model. Valid edits update the draft without a separate acceptance step. Invalid input remains editable with feedback beside its field; related dimension checks update immediately and prevent Apply. The bottom action bar offers Apply & restart, Discard all, Undo this category, and Undo last edit. Navigation preserves the draft. Read-only and preview commands remain available while another field needs correction. Problems stays visible in the footer, highlights errors in red, and opens links to the affected category. Shortcut overlaps appear last as saveable warnings.

Tables use a spreadsheet grid with an insertion row, Shift/drag ranges, Ctrl-click selection, Ctrl+A/C/V, single-value bulk fill, clearing, multi-row deletion, reordering and running-window selection. Event alerts use a compact table with enabled state, color, assigned sound name and additional conditions. Sound library contains reusable files, assignments, volume sliders and playback limits. Search highlights matching headings and descriptions. Graphite adds a neutral gray dark palette shared with the internal editor.

Guest verification commands:

```powershell
# Unit checks; exit code 0 means all assertions passed.
.\EVE-X-Preview.Settings.Tests.exe
# Mutates and restores SettingsTheme in an ISOLATED installation fixture;
# also exercises real profiles/inheritance/undo/validation and Apply/restart.
.\EVE-X-Preview.Settings.Tests.exe --integration C:\Checks\Fixture
# Clean fixture lease required; renders then releases lease and exits.
.\EVE-X-Preview-Settings.exe --installation C:\Checks\Fixture --capture-fixtures C:\Checks\Screenshots
```

Captures cover representative sections in light/dark at 1100×800 and 850×650. Files labelled `render150` simulate 150% rendering DPI; they are separate evidence from a real display-scale test. These checks and synthetic fixtures do not establish live EVE behavior. The parent verification log records actual guest results and visual inspection.


The native Windows title bar follows the selected theme, with the profile label and selector sharing one compact client header row. Table columns share available width instead of requiring a 1600-pixel window. Row numbers belong to the table: dragging the gutter selects whole rows; Delete removes complete selected rows, while a partial selection clears cells. Enter moves down one cell in the same column. Clipboard paste accepts newline-separated character names in the first column of wider tables, preserving other columns. Color swatches contain their chooser button and the picker includes a brightness/value slider down to black. Preview movement bypasses inherited per-window delays and avoids repainting unchanged content while translating its layers.

The editor and owned dialogs retain the native Windows caption and resize frame. Dark themes request the native dark caption; Windows 11 also supports exact palette caption, border and text colors.

The companion displays a loading view while the AHK bridge connects. A presentation-only cache remembers the last connected theme for the next startup; settings remain owned by AHK. Unavailable saved fonts appear as placeholders, never as installed list entries. About presents installed/stable/pre-release versions and seven aligned actions, including Download latest settings editor.

The user explicitly reauthorized tests on 2026-11-05; use disposable configurations for verification.

The release publisher creates a single self-contained `EVE-X-Preview-Settings.exe` with native libraries and the canonical host embedded. `SettingsHostPath` supplies the freshly compiled host resource during publishing; ordinary source builds can use the source/adjacent host fallback. The main/updater remain separate executables. Startup waits for a ready host before metadata/lease acquisition; cold probe and launch handshakes allow native extraction. Update checks run in the background and show a quiet stable-release notice linking to About.
