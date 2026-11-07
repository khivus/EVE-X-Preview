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

Choose the profile to edit, then use navigation or search. The search field has a compact clear button. Inherited categories offer **Edit Default here** while keeping the selected profile and its inheritance sources. Leaving the category returns to that profile. Section inheritance uses one-click checkboxes beneath category headings shown once; Default edits the new-profile inheritance template. Create profile uses the same grouped checkboxes, with Select all, Select none, and Use current profile inheritance buttons. Selecting a profile does not activate it; activation is staged for Apply. **Default**, section inheritance, new-profile inheritance templates, groups, rows, events, sound assignments and applicable undo remain supported. Ordinary controls come from the AHK schema. The companion sends mutations to the AHK draft model and displays canonical errors; it does not independently rewrite JSON.

Create profile keeps the name, copy/preset controls and Create/Cancel actions visible while its category choices scroll. Category navigation keeps the search, profile selector and action bar available. Rebuilding the current category after an inheritance change retains the current scroll offset. Runtime actions such as Stop sound preview do not rebuild the page. The header, sidebar, content and footer have visible separators; selecting a theme immediately updates existing controls and these regions.

Inherited categories explain their disabled state once at the top. Other disabled scalar settings name the controlling toggle in a category summary. Color swatches retain their actual color while disabled. Sound library explains that event assignments follow **Event alerts** inheritance and offers **Edit Default sound assignments here** without leaving Sound library or changing the selected profile. Pending edits retain their original profile while navigation completes, including when Default is edited in place.

In Mouse actions, changing a gesture dropdown edits only that action, even when other cells are selected. Opening a dropdown does not start drag selection across gesture rows. Shift/Ctrl selection and explicit clipboard paste remain available for intentional table operations.

**Monitored characters** owns both **Monitor only selected characters** and its character list, with one inheritance source. **Read game logs for alerts and DPS** remains in **Log sources**. Older profiles with different sources for the filter and list retain both effective values as local Monitored characters settings; their dormant local values are kept in the migration backup. **Advanced** contains the former Performance and compatibility settings.

**Apply & restart** validates the draft, checks concurrent edits, merges changed values into the latest runtime state, and uses canonical saving. Unrelated runtime changes such as thumbnail positions and discovered character information survive. Conflicts are reported instead of overwritten. Failed saves leave the draft open. The companion keeps the edited profile and window/page/scroll state while AHK restarts and reconnects to the new runtime. The footer counts settings changed since the last Apply, including retained unfinished input. Connection loss displays a persistent explanation above the editor, pauses editing and offers Reconnect; the draft is retained. Reconnection refreshes validation so the Problems list reflects the current draft.

Both editors support light, dark and system themes. In the companion, the action bar offers Apply & restart, Discard all, Undo this category and Undo last edit. Valid input updates the draft automatically; invalid input stays beside its error. The always-visible Problems button becomes active when issues exist and links to the affected category. Invalid settings turn the footer text red. Shortcut conflicts appear last as warnings and link to the category containing their bindings: saving remains allowed, but overlapping shortcuts may behave unpredictably. Advanced shortcut syntax is validated by AHK. Previews and immediate runtime actions use the AHK bridge.

**Thumbnails → Layout → Recover off-screen thumbnails** is available in both editors while the main application is running. It immediately restores partially or fully off-screen thumbnails, preserving their sizes and placements fully visible across any connected monitors. Recovery saves only corrected runtime positions and preserves unrelated settings drafts.

The shared selector also includes **The Matrix** (terminal green), **Kitsune** (rose and wine), **Synthwave** (neon pink and violet), and **Cyberpunk** (magenta, midnight blue and teal). Both editors consume the same canonical palettes. Interface themes do not change thumbnail or event alert colors. Selecting a theme previews it immediately; Apply saves the application-wide choice.

Theme previews: [The Matrix](images/settings-themes/matrix.png), [Kitsune](images/settings-themes/kitsune.png), [Synthwave](images/settings-themes/synthwave.png), [Cyberpunk](images/settings-themes/cyberpunk.png).

## Opening directly while the main application is stopped

Opening the companion starts a temporary canonical AHK settings host. It edits the same installation's file using the same migrations and validation and exits when the editor does. It does not start thumbnails or register runtime hotkeys. Runtime-only controls explain that the main application must be running.

If the main application starts during an offline edit, reconnect with the retained draft. Save serialization, installation-specific endpoints and conflicts prevent independent overwrites of the running application's configuration.

## Portable package

`EVE-X-Preview-Settings.exe` is a single self-contained Windows x64 executable. Place it beside `EVE-X-Preview.exe`; no adjacent DLLs, separate settings host or .NET installation are needed. Native .NET/Avalonia libraries unpack into the normal per-user .NET cache on first launch. When editing offline, the embedded canonical AHK host is verified and extracted into a per-user content-addressed cache; the configuration stays in the selected application folder.

The main AHK application remains usable without companion files and uses internal settings. Back up `EVE-X-Preview.json` before replacing an installation.

Opening the main application or the offline settings editor automatically snapshots the installation's settings before startup migration/normalization. The latest five copies are kept under `%LOCALAPPDATA%\EVE-X-Preview\Backups\<installation-id>`, separately for each installation. A backup error is logged and does not prevent startup. **About → Open settings backups** opens this folder in either editor, including offline mode. Migration/recovery backups go into the same folder and are retained separately from the five rotating startup snapshots. Existing `EVE-X-Preview.*.backup.json` files beside the application are moved there on startup without overwriting existing archives.

Under **Thumbnails → Appearance**, **Delay before clearing the border when EVE loses focus** defaults to 1000 ms, accepts 0–60000 ms, and follows Appearance inheritance. Returning to EVE resets the delay; switching directly between tracked clients updates the active highlight immediately. **Switching → Switching behavior** accepts a group hold interval down to 30 ms.

**Switching → Keyboard shortcuts** separates **Pause EVE switching hotkeys** from **Pause all hotkeys**. The existing pause binding retains its assignment with the narrower EVE-only scope. The full-pause binding is new and remains available to resume itself. Both commands have independent visibility controls under **Application → Tray menu commands**. See [hotkey priority and pause behavior](hotkeys.md).

See [development](development.md), [bridge v1](settings-bridge.md) and [C# project notes](../settings/README.md). Synthetic checks do not establish live EVE rendering or physical multi-monitor/DPI behavior.

Event alerts use a compact table for enabled state, border color and assigned sound filename without its extension. Additional conditions appear below the table, grouped by category and event, and become inactive while their event is disabled. Numeric fields share a compact width; sliders and percentage values occupy a single row. Sound library groups reusable sound files, batch and individual assignments, file gain, event volume, master volume and playback limits. Both the main editor and its dialogs use the native Windows dark caption and standard resize frame.

Startup shows a loading view while connecting to AHK and applies the native caption theme before showing the window. A small per-installation presentation cache retains the last connected theme without reading or writing the canonical settings file. New configurations use Segoe UI for thumbnail labels. An old factory Gill Sans MT value is replaced with an installed fallback only when Gill Sans MT is missing; installed and custom font choices are retained. Other unavailable saved fonts are shown as placeholders so they cannot be mistaken for an installed list entry. About displays installed/stable/pre-release versions and seven vertically aligned actions; update results remain on the page.

Both editors check for updates when Settings opens. The companion checks asynchronously without disabling editing; a newer stable release adds a quiet footer link to About. Failures remain in the About status without popups. Both About pages offer **Download latest settings editor** immediately below pre-release installation. It downloads `EVE-X-Preview-Settings.exe` directly beside the application, verifies the release version, and opens it. The action is disabled when the latest editor is already installed. The first launch of this update suggests the optional download if the editor is missing; declining leaves it available in About. Pre-release program installation remains an explicit About action.

The updater uses the main release version for all executables. Every main application update also replaces the settings editor when it is present in the installation folder. Installations without it stay without it. All required downloads are verified before replacement; a failed replacement restores the previous files.

Tray recovery is available from Thumbnails → Layout. **Pause Live Previews** appears in the tray only when enabled under Application → Tray menu commands; it is hidden by default in both settings editors.

In the companion's Diagnostics category, **Pause live previews** changes to **Resume live previews** when previews are paused. The action follows the running application's current state and refreshes while Diagnostics is open. Runtime actions remain unavailable when editing offline.

For the focused manual checks after the latest settings review, see [settings review recheck](settings-review-recheck.md).
