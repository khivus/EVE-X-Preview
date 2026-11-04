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

Launch the packaged editor beside EVE-X-Preview, or pass `--installation <directory>`. A direct launch discovers that installation's runtime. If absent, it starts the canonical packaged settings host or `SettingsHost.ahk` with an AutoHotkey v2 interpreter beside the script or in a standard per-machine/per-user location. It never starts the normal thumbnail/hotkey engine. The host watches the editor PID and exits when the editor closes.

`--probe <response-file>` writes only the small compatibility response and exits without Avalonia, IPC, or an editing lease. A named per-user/per-installation gate activates an already-open editor. The bridge uses bounded WM_COPYDATA requests from a real hidden HWND, installation-specific endpoints and matching response files. Detailed protocol: `../docs/settings-bridge.md`.

Ordinary controls are generated from field metadata, including choices, toggles, sliders, numeric text, colors, fonts, folders, multiline text and recorded/advanced shortcuts. Dedicated controls cover profiles, inheritance templates, ordered row tables and clipboard exchange, groups and shortcuts, event conditions, sound library/gain/assignments, running-window selection, diagnostics, updates and runtime previews. Validation, undo, save, conflicts and runtime behavior remain authoritative in AHK.

Typing creates a retained unfinished edit. **Accept edits** submits the pending values to canonical validation; invalid text stays visible. Unfinished edits stay attached to their original profile across navigation. Apply is disabled until all unfinished edits are accepted. Profile/group rename and removal protect pending edits from becoming orphaned. Apply validates the draft and requests the shared transaction; it does not write configuration directly. Discard and close prompts explicitly protect both accepted and unfinished changes. Reconnection retains the opaque canonical draft and local pending inputs. After confirmed Apply, the same window reconnects without replaying its previous unsaved recovery payload.

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
