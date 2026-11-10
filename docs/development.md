# Development

Release builds emit the existing standalone EXEs, with the settings editor optional. See [executable updates](executable-updates.md) for download selection, recovery and Sandbox integration tests.

[Back to README](../README.MD)

## Run from source

1. Download and install [AutoHotkey v2](https://www.autohotkey.com/).
2. Download the ZIP archive from the [releases page](https://github.com/khivus/EVE-X-Preview/releases).
3. Extract the archive and run `Main.ahk`.

## Contributing

Contributions are welcome. All pull requests must target the [`dev` branch](https://github.com/khivus/EVE-X-Preview/tree/dev).

1. Fork the repository and use the latest upstream `dev` as the starting point whenever possible. If you already have a fork, sync its `dev` branch with this repository before starting.
2. Create a separate branch for your changes from the updated `dev` branch.
3. Keep changes focused, test the affected behavior, and describe what changed and how you tested it in the pull request.
4. Before submitting, update your branch with the latest upstream `dev` whenever possible and resolve any conflicts.
5. Open the pull request against `khivus/EVE-X-Preview` with **base: `dev`**.

### EVE-X-Preview

* Created by [gonzo83](https://github.com/g0nzo83/EVE-X-Preview)
* Forked and maintained by [khivus](https://github.com/khivus)
* Credits to [mrmjstc](https://github.com/mrmjstc/eve-maj-preview) and [CJKondur](https://github.com/CJKondur/EVE-MultiPreview)

## Source layout and tests

`Main.ahk` starts the app, `src/` holds application code, `Lib/` holds libraries and default settings, and `res/` holds icons used by the script and Ahk2Exe. `Changelog.md` is embedded in compiled releases for the offline changelog viewer.

Build the established pair of portable Windows x64 release packages with `tools/Build-Release.ps1 -OutputDirectory <folder>`. It builds `EVE-X-Preview.exe` and `EVE-X-Preview-Updater.exe` using AutoHotkey v2 and verifies their Windows version metadata against the source directives. Both executables embed the application and suspend icons from `res/`. The main application, updater and optional settings editor share the release version. The updater replaces the settings editor only when it is already present.

Run the Windows AutoHotkey v2 suite from the repository root:

```powershell
& 'C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe' /ErrorStdOut tests/TestRunner.ahk --headless
```

Tests use temporary windows and fixture settings. Live EVE checks and rebuilding a release are separate verification steps.

## Avalonia companion and isolated development

The optional companion uses .NET 10 LTS and stable Avalonia dependencies pinned in `settings/`. Its ordinary controls consume metadata from `SettingsSchema`, `SettingsTables` and canonical defaults. `SettingsSession` owns external draft operations and the shared Apply transaction. See [settings editors](modern-settings.md) and [bridge v1](settings-bridge.md).

`tools/Build-Release.ps1 -OutputDirectory <folder>` requires a fresh or empty output folder, compiles and embeds the canonical host and publishes a single self-contained Windows x64 companion, including native/runtime dependencies. The output contains the main, updater and settings executables; the intermediate host is kept under `.vs/build/hosts` instead of shipped separately. It requires .NET 10 SDK and Ahk2Exe. Use `-BuiltInOnly` to produce the established main/updater pair. `-DotNetPath` selects a portable SDK; `-NoRestore` uses restored dependencies. Caches default to `.vs/build`, and SDK certificate/tool-path first-use changes are disabled. Debug symbols are omitted from the portable distribution.

Build independently with:

```powershell
dotnet publish settings/EVE-X-Preview.Settings/EVE-X-Preview.Settings.csproj -c Release -r win-x64 --self-contained true -o <package-folder>
```

For source-mode execution pass `--installation <repository-folder>` to the companion. It discovers the canonical host there and uses a standard AutoHotkey v2 installation or portable `AutoHotkey64.exe` beside the source.

The AHK suite creates real Windows controls. To isolate it, use Windows Sandbox with a read-only source/input snapshot and a dedicated writable output folder. This workspace's `.vs/isolated/README.md` describes its launcher. `tools/Invoke-SandboxChecks.ps1 -SandboxId <id> -Mode Ahk` snapshots tracked and nonignored new files. `-Filter` selects tests; `Managed` runs C# assertions; `Integration` exercises a fresh offline host; `RuntimeIntegration` exercises actual AHK restart/reconnect. These helpers expect guest `C:\Input`, `C:\Output`, and a copied runtime. Never point integration checks at real user settings.

`tests/Exe_Update_Integration.ps1` exercises compiled release updates and runs `tests/Update_Startup_Guard_Test.ps1` first. The latter compiles a disposable notice fixture and checks delayed or absent button readiness, rejection of an incorrect exit code, and cleanup after a missing notice or a process that stays alive after acknowledgement. It can also run independently with `-OutputDirectory <guest-folder>` and `-AutoHotkeyDirectory <portable-AHK-folder>`. The startup guard helper waits for the button to be ready, sends its native acknowledgement command and requires exit code 1. Failed checks terminate only the process they launched; the integration wrapper also bounds its wait and cleans up its exact fixture path.

Publish `settings/EVE-X-Preview.Settings.Tests` self-contained to the guest input's `settings-tests` folder. Run it with no arguments for client tests or `--integration <isolated installation>` for real message-bridge checks. Capture visuals with the companion's `--installation <fixture> --capture-fixtures <output-folder>` mode. Captures at150% are simulated rendering checks; physical monitor scaling needs separate verification.

## Settings architecture and focused checks

The production entry point opens SettingsWorkbench. SettingsModel owns an isolated draft and merges changed leaves into current runtime JSON on Apply. SettingsSchema supplies labels, storage routes, validation and search metadata. SettingsTables preserves related values as rows while retaining legacy JSON formats. Workbench pages, row/event editors and scrollable child dialogs use native Windows controls.

Settings_Review_Followup_Test.ahk covers immediate scalar-swatch WM_PAINT requests, invalid/valid hint colors and size relations, slider-owned ticks, invalid-shortcut details and disposal during commit, new-profile inheritance templates/effective copies, creation presets, popup geometry and About alignment. Run it with `--filter "Settings review"`. New-profile inheritance defaults live in the application-wide `NewProfileInheritance` map; changing the template does not alter existing profiles.

Settings_Redesign_Model_Test.ahk covers isolation, merge behavior, migration, validation and row serialization. Settings_Workbench_Test.ahk exercises real native controls, navigation, dialogs, inheritance, Apply/Discard, lifecycle protection and sound workflows. The legacy page builders remain covered as compatibility helpers; the production settings entry point uses the workbench.

Run a focused group with:

```powershell
& 'C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe' /ErrorStdOut tests/TestRunner.ahk --headless --filter Workbench
```

The filter must match at least one test. Render isolated fixture snapshots with tests/Render_Settings_Workbench.ahk and an optional output-directory argument; it never loads user settings. Verify font/text fit, child-dialog scrolling and monitor-work-area bounds at relevant display scales. Live EVE rendering, audio-device behavior and updater handoff remain separate manual checks.

SettingsGrid supplies native cell overlays, selection painting, clipboard and bulk operations to all settings tables. Settings_Grid_Test.ahk covers row creation, range fills, atomic validation, draft/profile lifecycle, picker/inheritance interaction, slider precision and resize/navigation behavior. Run it with `--filter Grid`.

Settings_Interaction_Followup_Test.ahk sends native mouse and character messages to check pane/table scroll stability, initial typing/caret placement, drag selection and dropdown editing. It also covers inline raw shortcut drafts, blank targets, subsection migration/runtime sources, inherited control states and direct/group shortcut priority. Run it with `--filter "Settings interaction"`.

SettingsTheme owns the settings palette and native text/button/header drawing, with cleanup when windows are destroyed. Color swatches retain their configuration RGB values. Default editing uses a scoped editing target while preserving the sidebar selection and activation draft. Settings_Appearance_Test.ahk covers swatch pixels, dark-theme navigation and child editors, automatic update checks using simulated responses, and Default editing across scalar/structured settings and validation. Run it with `--filter "Settings appearance"`.

Settings_Polish_Test.ahk covers palette-switch pixels, immediate enabled-state repaints, literal/multiline search highlighting, top-level recording ownership for panes and dialogs, and overlapping drag selection. Run it with `--filter "Settings polish"`.

SettingsNativeChrome subclasses controls to paint dark borders, dropdown arrows, scrollbars and spinners. SettingsScrollbars handles dark scrollbar capture, dragging and repeated clicks without entering the system light tracking loop; native ranges and keyboard/accessibility scroll commands remain available. Pane and table thumb positions retain 32-bit precision. Light mode restores only the original border bits, preserving editable/read-only state. Dropdown popup bounds use the actual editor window and monitor work area. SettingsTooltip supplies a native, nonactivating tracked tooltip in either palette. Settings_Chrome_Test.ahk checks native pixels, light/dark restoration, dropdown bounds/selection, numeric changes, held dragging on both axes, capture lifecycle and tooltip focus. Run it with `--filter "Settings chrome"`.

SettingsTheme batches layout and page replacement with nested redraw scopes. Whole-window WS_EX_COMPOSITED is deliberately disabled because it conflicts with local native/custom control painting. Tables use LVS_EX_DOUBLEBUFFER and WS_CLIPCHILDREN; headers buffer their complete paint and suppress background erasure. Cell selection invalidates only the rectangle below the mapped header bounds. RDW_NOCHILDREN alone does not prevent the parent's buffer from covering the header. Moving a settings canvas discards stale copied pixels and repaints its visible children after positioning. Hidden windows stay hidden when an update scope ends. Buttons use the same custom paint handler for normal, pressed, hover, disabled, focused and default states, retaining native input and dialog-key behavior. Workbench sizing defers page replacement until WM_EXITSIZEMOVE; height-only changes retain controls and inline edits. Scrollbar parts are buffered together, with a reentrancy guard and no extra table client-DC pass after WM_PAINT. Settings_Paint_Lifecycle_Test.ahk checks tray activation and resizing (`--filter "Settings painting"`). Settings_Scroll_Rendering_Test.ahk observes paint/erase messages and update bounds to check header isolation, repeated click/editor focus across palettes, idle painting, separators and selection through held scrolling (`--filter "Settings scroll rendering"`).

Top-level Settings and owned editor windows retain WS_VISIBLE throughout nested update scopes. WM_SETREDRAW is reserved for visible child windows because its default handling would briefly remove the Settings taskbar button and disturb focus/window order. The existing control buffers and final redraw present the completed layout. Settings_Paint_Lifecycle_Test.ahk checks visibility during updates and native button dispatch, along with focus and window order across page, profile and palette changes.

SettingsScrollPane now uses a separate native ScrollBar child instead of WS_VSCROLL on its viewport. SettingsNativeChrome owns its dark client painting, hit geometry and keyboard path; SB_CTL retains the native 32-bit range and SBM_GETSCROLLBARINFO exposes matching accessible geometry. First input never falls back to native modal light tracking. Slider callbacks use TB_THUMBTRACK to avoid writing TBM_SETPOS during capture, then snap at release. Unchanged UI values use guarded assignments. Preserved-color swatches have their own paint path; UpDown controls use UDS_NOTHOUSANDS. Grid editor focus loss ignores stale events during opening and focus within the grid; drag selection computes XOR before its single redraw and skips an unchanged endpoint. Settings_Input_Pulse_Test.ahk covers these input sequences (`--filter "Settings input pulse"`).

SettingsNativeChrome also brackets dark ListView/TreeView native scroll, layout and expansion updates with a local WM_SETREDRAW scope, then presents the final client/nonclient paint. Ordinary table selection retains its existing paint path. Existing caller redraw locks and hidden controls are left alone; the subclass tracks WM_SETREDRAW because common controls may consume it without setting SysSetRedraw. Vertical table repaints exclude the header, and dark navigation uses TVS_EX_DOUBLEBUFFER for clean background clearing after collapse. Settings_Release_Safety_Test.ahk observes below the chrome subclass to verify the redraw state received by the native procedure, checks buffered table erasure, exercises real sidebar expander hits and owned profile grids, and confirms modified-chord warnings. Run it with `--filter "Settings release safety"`. Offscreen fixtures cannot verify compositor timing on every display or DPI configuration.

## Audio output lifetime

`EventAudioOutput` owns one XAudio engine and mastering/output voice for the process. Main and the offline settings host prepare it during startup, before advertising readiness for playback; standalone players prepare it when constructed. Preparation is best effort: missing output does not prevent startup, and later playback can retry initialization. Device preparation can still take time, so measure it separately from alert playback. The default virtual endpoint remains enabled for Windows-managed output switching.

Each `EventAudioStream` owns its decoder, source voice, bounded sample queue and gain, and holds a reference to the shared output. Stop destroys that source voice before releasing buffers and its output reference. Other sounds keep playing; completed files remain deletable. The shared output stays available between alerts and is released at process teardown after its remaining stream references are gone. No extra release dependency is required.

The existing native asynchronous-playback test keeps its 450 ms limit. `Six native alerts start within one input-latency budget after audio preparation` also applies that budget to a six-channel burst after preparing/stopping one sound, checks independent stopping, and deletes the original file after disposal. Run these executable/UI checks only inside Windows Sandbox. Fresh-process repetitions distinguish cold preparation cost from the near-zero preparation cost of an already initialized test process.
