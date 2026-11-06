# Settings editors

[Back to README](../README.MD)

EVE-X-Preview supports its internal AutoHotkey settings workbench and the separate Avalonia application `EVE-X-Preview-Settings.exe`. The companion runs on demand and exits when closed. Thumbnails, shortcuts, window handling, game logs and sound playback continue to belong to AutoHotkey.

## Choosing an editor

Under **Application → Settings appearance**, **Settings editor** is shared by every profile:

- **Automatic** (default): open a compatible companion beside the main application, otherwise use internal settings.
- **Modern (C#)**: prefer the companion and use internal settings if it is missing, incompatible or cannot finish startup. The application log records the reason.
- **Built-in (AHK)**: open the internal workbench.

Tray **Open** and default double-click follow this selection. **Open internal settings** remains available separately. Discovery is relative to the application directory. A compatibility probe checks bridge/schema versions; the filename alone is insufficient.

Opening the companion again activates its existing window for that installation. Separate installations have separate sessions. Complete or discard a draft before moving to the other editor. Pending invalid text and editable dialogs also protect the draft. Runtime reload, exit and profile commands cannot silently discard it.

## Editing and applying

Choose the profile to edit, then use navigation or search. Selecting a profile does not activate it; activation is staged for Apply. **Default**, section inheritance, new-profile inheritance templates, groups, rows, events, sound assignments and applicable undo remain supported. Ordinary controls come from the AHK schema. The companion sends mutations to the AHK draft model and displays canonical errors; it does not independently rewrite JSON.

**Apply & restart** validates the draft, checks concurrent edits, merges changed values into the latest runtime state, and uses canonical saving. Unrelated runtime changes such as thumbnail positions and discovered character information survive. Conflicts are reported instead of overwritten. Failed saves leave the draft open. The companion keeps window/page/scroll state while AHK restarts and reconnects to the new runtime.

Both editors support light, dark and system themes. In the companion, the action bar offers Apply & restart, Discard all, Undo this category and Undo last edit. Valid input updates the draft automatically; invalid input stays beside its error. The always-visible Problems button becomes active when issues exist and links to the affected category. Invalid settings turn the footer text red. Shortcut conflicts appear last as warnings: saving remains allowed, but overlapping shortcuts may behave unpredictably. Advanced shortcut syntax is validated by AHK. Previews and immediate runtime actions use the AHK bridge.

**Thumbnails → Layout → Recover off-screen thumbnails** is available in both editors while the main application is running. It immediately restores partially or fully off-screen thumbnails, preserving their sizes and placements fully visible across any connected monitors. Recovery saves only corrected runtime positions and preserves unrelated settings drafts.

The shared selector also includes **The Matrix** (terminal green), **Kitsune** (rose and wine), **Synthwave** (neon pink and violet), and **Cyberpunk** (magenta, midnight blue and teal). Both editors consume the same canonical palettes. Interface themes do not change thumbnail or event alert colors. Selecting a theme previews it immediately; Apply saves the application-wide choice.

Theme previews: [The Matrix](images/settings-themes/matrix.png), [Kitsune](images/settings-themes/kitsune.png), [Synthwave](images/settings-themes/synthwave.png), [Cyberpunk](images/settings-themes/cyberpunk.png).

## Opening directly while the main application is stopped

Opening the companion starts a temporary canonical AHK settings host. It edits the same installation's file using the same migrations and validation and exits when the editor does. It does not start thumbnails or register runtime hotkeys. Runtime-only controls explain that the main application must be running.

If the main application starts during an offline edit, reconnect with the retained draft. Save serialization, installation-specific endpoints and conflicts prevent independent overwrites of the running application's configuration.

## Portable package

`EVE-X-Preview-Settings.exe` is a single self-contained Windows x64 executable. Place it beside `EVE-X-Preview.exe`; no adjacent DLLs, separate settings host or .NET installation are needed. Native .NET/Avalonia libraries unpack into the normal per-user .NET cache on first launch. When editing offline, the embedded canonical AHK host is verified and extracted into a per-user content-addressed cache; the configuration stays in the selected application folder.

The main AHK application remains usable without companion files and uses internal settings. Back up `EVE-X-Preview.json` before replacing an installation.

See [development](development.md), [bridge v1](settings-bridge.md) and [C# project notes](../settings/README.md). Synthetic checks do not establish live EVE rendering or physical multi-monitor/DPI behavior.

Event alerts use a compact table: enabled state, border color, assigned sound filename without its extension, and additional conditions. Sound library groups reusable sound files, batch and individual assignments, file gain, event volume, master volume and playback limits. Both the main editor and its dialogs use the native Windows dark caption and standard resize frame.

Startup shows a loading view while connecting to AHK and applies the native caption theme before showing the window. A small per-installation presentation cache retains the last connected theme without reading or writing the canonical settings file. Unavailable saved fonts are shown as placeholders so they cannot be mistaken for an installed list entry. About displays installed/stable/pre-release versions and seven vertically aligned actions; update results remain on the page.

Both editors check for updates when Settings opens. The companion checks asynchronously without disabling editing; a newer stable release adds a quiet footer link to About. Failures remain in the About status without popups. Both About pages offer **Download latest settings editor** immediately below pre-release installation. It downloads `EVE-X-Preview-Settings.exe` directly beside the application, verifies the release version, and opens it. The action is disabled when the latest editor is already installed. The first launch of this update suggests the optional download if the editor is missing; declining leaves it available in About. Pre-release program installation remains an explicit About action.

The updater uses the main release version for all executables. Every main application update also replaces the settings editor when it is present in the installation folder. Installations without it stay without it. All required downloads are verified before replacement; a failed replacement restores the previous files.
