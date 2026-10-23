## If you found a bug or have suggestions, you can file a [GitHub issue](https://github.com/khivus/EVE-X-Preview/issues/new) or [join the Discord server](https://discord.gg/d9EzgUXnT6).

<p align="center">
<img width="702" height="738" alt="Image" src="https://github.com/user-attachments/assets/40e1f9ea-7555-4d7d-b91d-7c42c66e791b" />
</p>

## Before running the program

Before first use, I suggest creating a backup of your settings. The program creates a backup when settings are migrated, but it is still a good idea to make one manually.

Use **Test Thumbnail** at the bottom of the Settings sidebar to preview the current default position, size and appearance. It refreshes as you edit settings or change profiles and displays sample system, DPS and event overlays only when their features are enabled. Active/inactive borders, assigned group colors and enabled events cycle every second. Character names rotate through configured custom-color characters with their matching aliases and colors, then custom-name characters when no custom-color characters are available; otherwise varied test names demonstrate different text lengths. The test remains visible while Settings is open, even when normal previews are hidden or paused. Click **Close Test Thumbnail** or close Settings to remove it. It does not capture an EVE client or save a thumbnail position.

Each sample change advances the system through **Jita → Dodixie → Amarr → Hek → Tama → Turnur → Ahbazon**, then repeats. Each system has a distinct demo color whenever system text is enabled, including when the live system-color option is off. Incoming/outgoing DPS and damage-type percentages also vary, including mixed and single-type damage. Samples respect enabled directions and percentage display, and remain above configured DPS thresholds. These sample values and system colors belong only to the test thumbnail; live color assignments and saved settings are preserved.

The test thumbnail follows **Thumbnails Interactions**: use your configured move and resize gestures (by default right-drag to move, Ctrl + both mouse buttons to resize, or both buttons to resize all). Move All/Resize All affect the test thumbnail, and resizing respects the configured minimum size. Activate-click still advances the sample; middle-click also cycles without taking over a configured gesture. Group actions preview their border state, minimize shows the inactive state, hide lasts until the next sample, and close removes the test thumbnail. Moving and resizing are temporary and survive appearance edits; changing the default geometry or selected profile returns the test to that default. Cycling pauses during a gesture.

The program automatically adds `*` when registering window-switching hotkeys so they also work with extra modifiers held. Bindings already containing `*` are used as entered. Custom `key & key` combinations already have wildcard behavior and are preserved as entered; saved settings and displayed text are preserved.

I care a lot about performance, so if you disable some features, the program will stop checking for unnecessary things.

# EVE-X-Preview

EVE-X-Preview is a Windows tool for EVE Online that helps manage multiple clients.
It shows thumbnails of running EVE clients and lets you bring a client to the foreground by clicking its thumbnail or using a hotkey.
Mouse and keyboard input are supported for interaction.

The program supports multiple profiles, so you can quickly switch hotkey assignments, thumbnail positions, and color settings.

## What it does

* Displays thumbnails for active EVE clients.
* Brings an EVE client to the foreground with a hotkey, mouse click, or window action.
* Supports full keyboard and mouse interaction.
* Lets you create multiple profiles.
* Lets you customize text, borders, font, opacity, colors, and placement.
* Supports different thumbnail sizes.
* Supports hotkey groups.
* Can save and restore client positions.
* Includes a graphical interface for all settings.

## What it does not do

* It does not broadcast keyboard or mouse input to multiple clients.
* It does not send individual input to clients in the way input broadcasting tools do.
* It does not display cropped client content as thumbnails.
* It does not manipulate the game client itself.

EVE-X-Preview does not interact directly with the game client. It only changes the foreground window using hotkeys and thumbnails.
No features will be added now or in the future that intentionally break the EVE Online EULA or Terms of Service.

## How to use

1. Download the `EVE-X-Preview.exe` from the [releases page](https://github.com/khivus/EVE-X-Preview/releases).
2. Create a folder anywhere except Program Files or other protected program folders, then move `EVE-X-Preview.exe` into that folder.
3. Start `EVE-X-Preview.exe`. It will create a JSON file in the same folder. This file must stay together with the executable because it stores all settings.
4. A tray icon will appear in the bottom-right corner of the taskbar.
5. Right-click the tray icon and choose **Open**, or double-click it, then go to **Profile Settings** and create a new profile or edit default one.
6. Configure the settings as needed.

`Important Note:`
`Some settings require a restart of the tool. The settings window must be closed before those changes are applied.`

## Run from source

1. Download and install [AutoHotkey v2](https://www.autohotkey.com/).
2. Download the ZIP archive from the [releases page](https://github.com/khivus/EVE-X-Preview/releases).
3. Extract the archive and run `Main.ahk`.

## Will I get banned if I use it?

As I understand the EVE Online EULA and Terms of Service, this tool does not break any rules.
Since EVE-X-Preview uses the same method as [EVE-O Preview](https://github.com/EveOPlus/eve-o-preview), a similar tool that is permitted, I assume the same applies here.
Nevertheless, using EVE-X-Preview or any other third-party tool is always at your own risk.

**I assume no liability whatsoever.**

## Mouse gestures and actions

Mouse gestures apply to the thumbnail window that is currently hovered.

| Action                                        | Gesture                                                       |
| --------------------------------------------- | ------------------------------------------------------------- |
| Activate the EVE client and bring it to front | Left click the thumbnail                                      |
| Minimize the EVE client                       | Hold `Ctrl` and click the thumbnail                           |
| Move a thumbnail                              | Right-click and drag                                          |
| Move all thumbnails                           | Hold `Ctrl` and right-click drag                              |
| Resize all thumbnails                         | Press both left and right mouse buttons, then drag up or down |
| Resize a single thumbnail                     | Hold `Ctrl` and press both mouse buttons, then drag           |
| Disable thumbnail from hotkey group           | Hold `Shift` and click the thumbnail                          |

## Tray menu

| Option                        | Description                                                          |
| ----------------------------- | -------------------------------------------------------------------- |
| Open                          | Opens the settings window                                            |
| Select profile                | Selects settings profile and reloads program                         |
| Suspend Hotkeys               | Disables all hotkeys                                                 |
| Hide Thumbnails               | Hides all thumbnails                                                 |
| Show Thumbnails Always on Top | Keeps thumbnails above other windows                                 |
| Click Through Thumbnails      | Lets mouse clicks pass through thumbnails                            |
| Minimize Inactive Clients     | Minimizes clients that are not active                                |
| Don't Close Active Client     | Prevents the active client from being closed by the close-all action |
| Close all EVE Clients         | Closes all open EVE windows                                          |
| Restore Client Positions      | Restores previously saved client positions                           |
| Save Client Positions         | Saves the current positions of all EVE clients                       |
| Auto Save Thumbnail Positions | Saves thumbnail positions automatically                              |
| Save Thumbnail Positions      | Saves the current thumbnail positions                                |
| Reload                        | Restarts the program                                                 |
| Exit                          | Closes the program                                                   |

## Hotkey and key syntax

Hotkeys use AutoHotkey v2 syntax.
Click the small keyboard button beside a hotkey field to record a shortcut inline. The field turns gray and shows a waiting prompt; press and release the shortcut to fill it. For mouse buttons or the wheel, use the gray field itself. In a multiline hotkey list, click the target line before recording. Esc, clicking the recorder button again, or leaving the field for another setting cancels and keeps the existing binding. Manual editing remains available.

An amber **!** indicates another assignment with the same or overlapping binding. Hover over it or click it to see the other assignments. Checks include the selected profile's effective bindings and all profile activation shortcuts. Repeated bindings can cause unpredictable or inconsistent behavior; warnings never block saving or importing settings.

For a detailed reference, see [Hotkey Modifier Symbols](https://www.autohotkey.com/docs/v2/Hotkeys.htm#Symbols) and [List of Keys](https://www.autohotkey.com/docs/v2/KeyList.htm).

A few useful rules:

* Simple hotkeys can be a key name or a key code, for example `1`, `Tab`, `LButton`, `XButton1`, or `Ctrl`.
* Modifier prefixes can be combined, for example `^1`, `+1`, `#1`, or `!1`.
* Prefix form can stack, for example `+^1` means Shift+Ctrl+1.
* The `*` prefix allows a hotkey to trigger even when another key is held.
* Using a mouse button as a simple hotkey is not recommended because it may block normal mouse use.

## Settings overview

### Profile selector and group navigation

The left panel contains the profile selector and profile management buttons:

* Add profile
* Rename profile
* Delete profile

Select a profile and enter its **Profile hotkey** using AutoHotkey syntax (for example, `^!F1` for Ctrl+Alt+F1). Close Settings to apply. Each profile has its own shortcut, independent of settings-group overrides, and all configured profile shortcuts work even when EVE is closed. Pressing one activates that profile through the same save-and-reload action as the tray Profiles menu. Profile shortcuts pause while Settings is focused and follow Suspend Hotkeys. Leave the field empty to remove a shortcut, and use a different key combination for each profile and other action. Copying a profile's settings leaves the new profile's shortcut empty; renaming preserves it.

It also contains buttons for settings groups. When you select a group, the main window shows the settings for that group.

### Hotkey Groups

Use hotkey groups to organize characters into ordered groups.
The order in the list matters: top to bottom.

Settings in this group:

* Preserve Hotkeys on Logout: keeps hotkeys assigned after character logout.
* Keep Groups Positions: keeps group positions when switching from one group to another.
* Disable Char in Group (Shift+Click): lets you disable a character in a group with Shift+Click.
* Disabled Thumbnail Color (Hex/RGB): color used for disabled thumbnails.
* Groups Hold Delay (ms): delay used when holding group actions. Minimum is 75 ms.
* Max Active Window Retries / Active Window Retry Interval (ms): bound focus checks only when a recent activation request targets a member of the requested group. Switching from another group or application normally skips these waits. Requests expire after 250 ms and are cleared when focus is observed, moves elsewhere, or the retry budget runs out. Set retries to 1 to disable waiting.
* Add/Delete Groups: create or remove groups.
* Select Group: choose the group to edit.
* Forwards Hotkey: hotkey for moving forward through the group.
* Backwards Hotkey: hotkey for moving backward through the group.
* Characters List: list of characters in the group.

### Hotkeys Settings

This group controls global hotkeys and client actions.

* Suspend All Hotkeys - Hotkey: pauses all hotkeys.
* Hide Thumbnails Hotkey: toggles thumbnail visibility.
* Pause/Resume Live Previews - Hotkey: releases live thumbnails and restores them on the next press, as with the About button.
* Click Through Thumbnails - Hotkey: toggles click-through mode.
* Hotkey Activation Scope: chooses between global hotkeys or hotkeys active only when an EVE window is focused.
* Cycle Login Screens - Hotkey: cycles through login-screen windows.
* Login Screen Cycle Direction: chooses old-to-new or new-to-old order.
* Close Active EVE Window - Hotkey: closes the currently active EVE window.
* Close All EVE Windows - Hotkey: closes all EVE windows.
* Reload EVE-X-Preview - Hotkey: reloads the program.
* Exit EVE-X-Preview - Hotkey: closes the program.
* Character Name: character bound to the hotkey entry.
* Hotkeys: hotkeys assigned to the character.

### Thumbnails Interactions

Quick Group Sort Order chooses how the Quick Group hotkey cycles through selected windows. Added order (the default) follows the order windows were added to Quick Group; Name cycles alphabetically by character name. Removing and adding a window again places it at the end of added order.

**Keep Quick Group on Logout** and **Keep Disabled Status on Logout** are enabled by default. Membership follows the same client window through character selection and relogin, including switching characters. Turn either option off in Thumbnails Interactions to clear that status on a character change or logout. Closing the client clears its membership; restarting the client or EVE-X-Preview starts fresh.

**Don't Close Quick Group Clients** and **Don't Close Disabled Clients** protect those windows from **Close all EVE Clients**, including while on the login screen, independently of **Don't Close Clients on Login Screen**. Directly closing an individual client still works.

### Thumbnails Behavior

This group controls thumbnail placement, movement, and visibility behavior.

* Auto Save Thumbnail Positions: saves thumbnail positions automatically on thumbnail move or resize.
* Hide Thumbnails: hides all thumbnails.
* Hide Thumbnails on Lost Focus: hides thumbnails when no EVE client is active.
* Show Thumbnails Always on Top: keeps thumbnails above other windows.
* Click Through Thumbnails: lets clicks pass through thumbnails.
* Default Thumbnail Position (px): default X/Y position.
* Default Thumbnail Size (px): default width and height.
* Enable Thumbnail Snap: enables thumbnail snapping.
* Thumbnail Snap Distance (px): snap threshold in pixels.
* Hide Thumbnail for Active Window: hides the thumbnail of the active client.
* Shift Thumbnails on Login Screen: offsets new thumbnails from the default position.
* Shift Logged-In Thumbnails (Unsaved): enabled by default; offsets new thumbnails for already logged-in characters without a saved position, using the same shift direction, steps, collision avoidance and monitor bounds as login-screen thumbnails. Saved positions take priority; this option works independently of login-screen shifting.
* Enable Thumbnail Collision Avoidance: prevents thumbnails from overlapping while shifting.
* Thumbnail Shift Direction: stacking direction for new thumbnails.
* Thumbnail Shift Horizontal Step (px): horizontal shift step; `0` uses thumbnail width.
* Thumbnail Shift Vertical Step (px): vertical shift step; `0` uses thumbnail height.
* Preserve Character Name on Logout: keeps the name visible after logout.
* Preserve Thumbnail Position on Logout: keeps the thumbnail position after logout.

### Thumbnails Visuals

This group controls how thumbnails look.

* Show Thumbnail Text Overlay: shows or hides the character name on the thumbnail.
* Thumbnail Text Color (Hex/RGB): text color.
* Thumbnail Text Size: text size.
* Thumbnail Text Font: font used for thumbnail text.
* Thumbnail Text Margins (px): text padding inside the thumbnail. When double borders are enabled, the inner ring's thickness is added automatically on all sides for every text overlay; configured values remain unchanged. Combined and single-border layouts use the configured margins directly.
* Client Highlight Color (Hex/RGB): border color for the active client.
* Client Highlight Border Thickness (px): thickness of the active-client border.
* Show Client Highlight Border: shows or hides the active-client border.
* Event / Active Border Layout: **Event in, active out** (default) draws independent, adjacent event and active rings at the thumbnail edge. **Active in, event out** swaps their positions. **Combined (existing)** uses the previous single shared border. Event flashes use the client highlight thickness; inactive and group borders retain their configured styling. Existing event enable, focus-suppression and stop-on-switch settings still apply. The test thumbnail previews both rings together and updates immediately; live thumbnails apply layout changes when Settings closes.
* Thumbnail Opacity (%): thumbnail transparency.
* Show All Borders: shows borders around every thumbnail.
* Inactive Client Border Thickness (px): thickness of inactive-client borders.
* Inactive Client Border Color (Hex/RGB): border color for inactive clients.
* Thumbnail Background Color (Hex/RGB): background color used when thumbnail content is not available.
* Minimum Thumbnail Size (px): minimum allowed thumbnail size.
* Custom names: enter character names in the left column and displayed names on matching lines in the right column. A blank replacement keeps the original name. Import from Launched appends logged-in characters without duplicating existing entries. Changes apply when settings close and follow the selected Thumbnails Visuals profile; hotkeys and character matching still use the original names.

### Thumbnail Visibility

This group controls whether selected client thumbnails are hidden.

* Select Any Client to Hide The Thumbnail: choose a client to hide.

### Client Settings

This group controls EVE client behavior.

* Minimize Inactive Clients: minimizes clients that are not active.
* Always Maximize Clients: keeps clients maximized.
* EVE Window Minimize Delay (ms): delay before minimizing an inactive client.
* Don't Close Clients on Login Screen: excludes login-screen windows from close actions.
* Don't Close Active Client: excludes the active client from close actions.
* Don't Minimize Clients: list of clients that should never be minimized.
* Don't Close Clients: list of clients that should never be closed.

### Custom Colors

Every color field supports the Windows color picker. Click the colored swatch beside a single color value. For a per-character color list, select the target line and click its picker button. The picker starts with the current color, and Cancel leaves it unchanged. Hex/RGB values can still be entered manually.

Use this group to assign per-character colors.

* Custom Colors Active: enables or disables custom colors.
* Character Name: character this color profile applies to.
* Active Border Color: custom border color for the active client.
* Text Color: custom text color.
* Inactive Border Color: custom border color for inactive clients.

### Game Logs Monitoring

Archive discovery runs in hidden background processes, so Settings and thumbnail controls stay available while large folders are scanned. Recent Local sessions are checked first and published as they are found while older headers continue to be checked. Discovery uses filename notifications and periodic fallback checks, keeps quiet sessions attached, and caches Local headers between scans. Long sessions use a background startup lookup and bounded live reads. Standalone builds embed the worker and must use an AutoHotkey v2 `.exe` compiler base (the usual build), which supports alternate embedded scripts.

This feature does not violate the EULA or Terms of Service, but it is neither endorsed nor supported by CCP. Use it at your own risk.
Performance impact increases with the number of characters tracked at the same time.

* Game Logs Monitoring Enabled: turns combat/event log monitoring on or off.
* Track Current System (Local Chat): follows Local chat system-change announcements and keeps the current system at the bottom of the thumbnail, with event text above it. Defaults to off and works independently of combat monitoring and focused-window event suppression. It shares the monitoring interval and selected-character filter; the thumbnail text overlay must be enabled to display it.
* Color System Names by System: gives characters in the same system the same system-name color, with distinct colors for different occupied systems. Requires current-system tracking and defaults to off. Changes apply immediately. Colors remain stable while a system is occupied; unused colors can be recycled during travel, and assignments are remembered for the current application session.
* Chat Logs (blank = auto): optionally select your Chatlogs folder. Automatic detection checks beside the configured Gamelogs folder, then Documents/logs/Chatlogs and Documents/EVE/logs/Chatlogs. The newest Local session is matched by its Listener header, including when tracking starts after login. Changes apply when settings close.
* Game Logs (blank = auto): optionally select your Gamelogs folder. Like Chat Logs, a blank field detects the sibling of the other configured log folder, then Documents/logs/Gamelogs and Documents/EVE/logs/Gamelogs. Automatic paths are resolved without filling in either settings field.
* Monitoring Interval (ms): polling interval.
* Event Priority if Several Incoming: chooses whether the first or last event wins.
* Supress for Focused Window: suppresses events for the active window.
* Show Event Text: shows a separate event message at the bottom inside the thumbnail (above the system name when tracking is enabled), using the thumbnail text font, color, and margins. The character name stays at the top.
* Flash Border: flashes the border on an event.
* Stop Displaying Event on Switch: stops showing the event after switching.
* Event Display Duration (ms): how long the event remains visible.
* Flash Border Interval (ms): flash timing (if `Flash Border Interval` = `Event Display Duration` border will not flash, but just stay with set color for specified time).
* Monitor Only Selected Characters: limits monitoring to selected characters.
* Characters to Monitor: list of monitored characters.

### Monitored events

Monitored events for Game Logs Monitoring feature.
Enabling more events reduces performance. The effect becomes larger as more characters are tracked.

Click the **♪** button beside any event's color to open **Event Sounds**. **Browse / Add...** adds a user-provided WAV or MP3 file to the shared **Sound library**. Select a library sound, check its events on the right, then click **Apply assignments**. **Select all** prepares to assign one sound to every event in one operation; **Clear checks** prepares to remove this sound from its events. Existing assignments appear in the Current sound column. Before Apply, an orange notice identifies replacements and the button shows how many other sounds will be replaced. Replacements are allowed. Unchecked events using a different sound are preserved. Assigning sounds does not enable disabled events.

The event opened with **♪** is prechecked when choosing a sound, so a single-event assignment also needs only Apply. Checkmarks show existing assignments to the selected file; changes are saved only on Apply. **Clear opened event** immediately removes that event's assignment. **Test selected** previews the library sound; **Test opened event** previews the opened event's actual assignment. The button changes to **♫** when a sound is assigned, with its file path in the tooltip. Playback depends on Windows media support. The application includes no sound files and starts with an empty library.

The dialog has three **0–300% volume** controls. **Master volume** affects every application sound across profiles. **Sound volume** affects every event using the selected library file across profiles. Select an event row to edit **Event volume**, which affects only that event in the effective Monitored Events profile. Sliders snap to 5% steps; the adjacent editable percentage fields accept exact whole values, with or without a percent sign. Slider and text values stay synchronized. Empty or invalid edits preserve the previous setting and normalize on leaving the field; values outside 0-300 are clamped. The gains multiply: 50% master × 50% sound × 50% event = 12.5% final volume. Defaults are all 100%; setting any level to 0% mutes it. Values above 100% amplify the audio: 300% is three times the original signal level. The three gains still multiply, so setting all three to 300% gives 2700% combined gain. Strong boosts can clip already-loud audio. Audio is decoded in bounded chunks and mixed with native gain; original sound files are not modified. Volume edits save immediately and adjust clips already playing without changing Windows or other applications. Test selected uses master × sound; Test opened event also includes its event volume.

The library is shared across profiles; sound assignments and event volumes follow the effective **Monitored Events** profile override. **Remove saved** removes only the library entry, preserving existing event assignments, the file's volume and the original file. The library stores file paths rather than copies, so keep selected files in a stable location. Missing or unsupported audio is skipped during monitoring and recorded in the program log; use Browse to select a moved file again. Sound assignments, all volume levels and the library are included in the settings JSON; moving that JSON to another PC does not transfer audio files.

**Simultaneous sounds** sets the maximum number of event sounds playing together, from 1 to 16 (default 1). A new alert replaces the oldest playing alert when the limit is reached; alerts are not queued for later. Lowering the limit stops the oldest excess clips immediately. **Repeat timeout (ms)** sets the minimum interval between starts of the same sound file across all events and characters, from 0 to 3600000 ms (default 2000). Set it to 0 to allow immediate repeats, including overlapping instances when there is room. Both controls are shared across profiles, save immediately, and are independent of visual alert duration. Muted, missing or failed sounds do not consume the repeat timeout. Manual Test buttons use a separate preview player and ignore these event limits and timeouts.

Enable **Game Logs Monitoring** and the event's checkbox for audio alerts. Sounds work with **Show Event Text** and **Flash Border** disabled, and while thumbnails are hidden. They respect event priority and focused-window suppression (including hiding the active client's thumbnail). The Settings test thumbnail's automatic samples remain silent; use the sound dialog's Test buttons to preview audio.

General NPC recognition matches full English names from a bundled EVE Ref dataset, ignoring case. Player names containing NPC name fragments no longer match. The faction, officer and capital events retain their separate lists and priority. See [NPC data and refresh instructions](src/data/README.md) for coverage and maintenance.

* Stopped Shooting
* Under Attack By Player: the Ignore player smartbomb damage option excludes Smartbomb hits by default; turn it off to include them. The adjacent Neuts checkbox also includes incoming energy neutralization and Nosferatu drains from players, identified by the red GJ amount in the log. Nosferatu logs use `energy drained to` for capacitor lost and `energy drained from` for capacitor gained; only capacitor loss triggers an alert. Outgoing neutralization is ignored.
* Under Attack By NPC: the adjacent Neuts checkbox also includes incoming energy neutralization and Nosferatu drains from NPCs. Both Neuts options default to off and require their parent attack event to be enabled.
* The Misses checkboxes beside Neuts on the Player attack and NPC attack rows independently control whether shots that miss you trigger the corresponding Under Attack event. Both default to on, preserving existing behavior. Disabling them still allows damaging hits and enabled Neuts alerts; faction, officer and capital engagement alerts keep their existing behavior.
* Engaged With Faction BS NPC
* Engaged With Officer NPC
* Engaged With Capital NPC
* Warp Disrupted
* Decloaked
* Gate Jumped
* Convo Request
* Fleet Invited
* Fleet Warped
* Fleet Regrouped
* Conduited
* Crystal Broke
* Mining Stopped
* Mining Bay Is Full

### DPS Monitoring

Enable **Game Logs Monitoring**, then enable incoming DPS, outgoing DPS, or both in **DPS Monitoring**. DPS shares the game-log directory, monitoring interval, and selected-character list. It works with event text and flashing borders disabled. The thumbnail text overlay must be enabled to see the values.

Set **Average Window (seconds)** to choose the rolling averaging period for both directions (1-3600 seconds, default **10**). Each value is the damage recorded during that window divided by its duration, preserving the spacing of the log's timestamps. For artillery firing every 17 seconds, a 20-second or longer window keeps the previous volley included between shots. The live window aligns to newly appended combat records and advances using elapsed time, so a difference between Windows UTC and EVE's clock does not discard hits. Thresholds are the minimum DPS to display for each direction; 0 means no minimum. A value below its threshold or with no recent damage is hidden. Old log contents are skipped when monitoring starts. Neutralization, repairs, misses, and other notifications do not count as damage. Event-specific NPC/player and smartbomb filters do not filter DPS.

**Incoming Damage Types (%)** estimates the EM, thermal, kinetic, and explosive mix only for drones, missiles (including rockets and torpedoes), fighters, and smartbombs. Missiles use their weapon-family names: Mjolnir = EM, Inferno = thermal, Scourge = kinetic, and Nova = explosive. Smartbombs use EMP = EM, Plasma = thermal, Graviton = kinetic, Proton = explosive, and Multispectrum = equal parts of all four. Drones, fighters, and smartbomb naming exceptions use exact names in the log. The bundled [EVE Ref data](https://docs.everef.net/datasets/reference-data) includes mixed-damage profiles. The display uses `?` for the percentage of damage with an unknown type. Other weapons still count toward total incoming DPS but remain unknown in the damage mix. Percentages include unknown damage in the denominator. They weight each known weapon's base damage mix by its recorded hit damage, so they are estimates, not measured post-resistance percentages or the ship's resistances. The existing saved setting key `showIncomingDPSResistances` is retained for compatibility.

DPS and the optional damage mix each occupy one line above event/system text: `DPS: In: 123 | Out: 456` and `% EM 25 Th 75`. Displayed DPS is rounded to whole numbers; calculations and thresholds retain full precision. Damage types that round to 0% are omitted, including unknown damage (`?`). The character name keeps priority and the system stays at the bottom. Active DPS takes priority over event text on compact thumbnails; the percentage breakdown appears when there is enough room. Empty event text reserves no line. Focus suppression hides DPS while the client is active but continues counting its damage, so switching focus does not create gaps. Event expiry and event switching do not clear DPS.

Monitoring reads only appended, complete lines through the existing reader and timer. DPS is updated before event visuals, and alert updates are batched once per character per poll while preserving first/last event priority. Per-second totals keep memory bounded; damage profiles are shared and loaded only when the optional damage mix is enabled. No network access is used while monitoring, and unchanged display text is not rewritten.

### Non-EVE Applications

This section lets you add non-EVE apps by exact executable name, and optionally by window title for better matching.
The thumbnail appears for the app, and the assigned hotkey can activate it.

* Add/Delete Groups: create or remove app groups.
* Select Group: choose the app group to edit.
* Forwards Hotkey: hotkey for forward group cycling.
* Backwards Hotkey: hotkey for backward group cycling.
* Process Names (.exe): executable names to match.
* Titles (optional): optional window titles to match.
* Hotkeys: hotkeys assigned to the app group.

### Tray Menu Settings

Choose which entries appear in the tray menu.

* Suspend Hotkeys.
* Hide Thumbnails.
* Show Thumbnails Always on Top.
* Click Through Thumbnails.
* Minimize Inactive Clients.
* Don't Close Active Client.
* Close all EVE Clients.
* Restore Client Positions.
* Save Client Positions.
* Auto Save Thumbnail Positions.
* Save Thumbnail Positions.

### Other

* Auto-Pause on DWM Failure: enabled by default, including existing profiles. Automatically pauses live previews on DWM failures or Windows composition changes. Applies immediately; turn it off to keep automatic retries. An existing pause still requires manual Resume.
* Slow Thumbnail Creation: off by default, including existing profiles. Off creates ready previews immediately without startup or spacing delays. On waits three seconds, requires stable source dimensions for 500 ms, and spaces queued previews by 150 ms to reduce startup GPU pressure. Applies immediately to queued previews and controls startup, new clients, and resume after pausing.
* Switch Language to English on Error: switches the Windows language to English if hotkey assigning error occurs.
* Update all thumbnails size to default size: use this button to update every thumbnail size to the current default size.

### Profile override behavior

Some settings groups can be overridden from the default profile to every other profile.
Each profile can also disable that override, which makes that group unique for that profile.

* Hotkey Groups: include the hotkey groups section in overrides.
* Hotkeys Settings: include hotkeys settings in overrides.
* Thumbnails Behavior: include thumbnail behavior settings in overrides.
* Thumbnails Visuals: include thumbnail visuals settings in overrides.
* Thumbnail Visibility: include thumbnail visibility settings in overrides.
* Client Settings: include client settings in overrides.
* Custom Colors: include custom colors in overrides.
* Game Logs Monitoring: include game log monitoring in overrides.
* Non-EVE Applications: include non-EVE applications in overrides.
* Monitored Events: include monitored event settings in overrides.
* Tray Menu Settings: include tray menu settings in overrides.
* Other: include other settings in overrides.

### About

Use **Show Changelog** in About to read the full changelog in a resizable window. Markdown headings appear larger and bold, with formatted bullets, inline bold and code. Text wraps and can be selected or copied. It also opens automatically after an update, showing the installed version. Monitoring, hotkeys and Settings remain available while it is open. The changelog is bundled with compiled releases and works offline.

Buttons:
* Check Updates: click to checks aviable updates.
* Update to Release: after update checking can use for update/downgrade to latest release version.
* Update to Pre-Release: after update checking can use for update/downgrade to latest pre-release version.
* Help: opens this README page.
* Report Bug: opens [GitHub issue](https://github.com/khivus/EVE-X-Preview/issues/new) page.
* Show Log: shows a snapshot of the current session's program events, even with Debug Mode off.
* Export Log: saves that event history to a chosen `.log` file.
* Pause Live Previews / Resume Live Previews: temporarily unregisters live DWM thumbnails and hides their overlays without removing existing hotkeys or group membership. With Auto-Pause on DWM Failure enabled in Other (the default), DWM failures or a Windows composition change automatically pause previews; About shows the reason. Resume manually after the display is stable, using this button or the pause/resume hotkey. Resume uses the Slow Thumbnail Creation option in Other. Pause is session-only and resets on restart.

Program diagnostics are kept in `ProgramLog.Text` (up to 256 Ki characters) and written every two seconds to `%LOCALAPPDATA%\EVE-X-Preview\Logs\current.log`. The disk log rotates at approximately 1 MiB into `previous.log`. Startup registration attempts are flushed before touching DWM. Errors reaching the script error handler, DWM failures, and logged runtime errors automatically export an `error-*.log` snapshot; snapshots are limited to one per ten seconds and the newest ten files. Logs can contain character/window names and local paths. Nothing is uploaded automatically.

Slow Thumbnail Creation is optional and disabled by default. When enabled, saved client geometry is restored before the stability wait. Both modes require nonzero source dimensions and retry startup failures independently, with increasing delays capped at 30 seconds. Auto-Pause on DWM Failure is enabled by default and applies immediately when changed in Other. When enabled, DWM failures stop the entire queue and require manual Resume; when disabled, failures remain logged and failed registrations retry. Disabling this setting does not resume already paused previews. Pause cancels interrupted startup attempts, waits for active DWM calls to return before cleanup, and prevents Resume during those calls. Unchanged geometry and visibility do not cause redundant DWM updates.

Diagnostics also record settings load/save/recovery, activation requests and focus results, Windows display-change notifications, and ten-second summaries of DWM geometry/visibility updates and window show commands, with GDI/USER handle counts. These summaries describe the program's operations, not the compositor's frame rate or GPU utilization; display notifications alone do not confirm a driver reset.

Update logs include automatic/manual checks, cache/duplicate skips, HTTP results, selected versions, timeouts, download results, cancellation, and handoff to the updater. The standalone updater writes replacement and relaunch steps to `%LOCALAPPDATA%\EVE-X-Preview\Logs\Updater`. This requires rebuilding and shipping the updated updater executable with the release; older published updater binaries will not contain the new logging.

These measures reduce startup bursts, including restarting Preview while clients are launching. They cannot guarantee recovery from a graphics-driver or DWM crash: those may never raise a script error. Preserve `current.log`/`previous.log` and collect Windows Event Viewer or Reliability Monitor details for the same time when reporting a blackout.

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
