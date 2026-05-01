## If you found a bug or have suggestions, you can file a [GitHub issue](https://github.com/khivus/EVE-X-Preview/issues/new) or [join the Discord server](https://discord.gg/d9EzgUXnT6).

<p align="center">
<img width="702" height="738" alt="Image" src="https://github.com/user-attachments/assets/40e1f9ea-7555-4d7d-b91d-7c42c66e791b" />
</p>

## Before running the program

Before first use, I suggest creating a backup of your settings. The program creates a backup when settings are migrated, but it is still a good idea to make one manually.

I suggest adding `*` at the beginning of every hotkey used for window activation.

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
* Add/Delete Groups: create or remove groups.
* Select Group: choose the group to edit.
* Forwards Hotkey: hotkey for moving forward through the group.
* Backwards Hotkey: hotkey for moving backward through the group.
* Characters List: list of characters in the group.

### Hotkeys Settings

This group controls global hotkeys and client actions.

* Suspend All Hotkeys - Hotkey: pauses all hotkeys.
* Hide Thumbnails Hotkey: toggles thumbnail visibility.
* Click Through Thumbnails - Hotkey: toggles click-through mode.
* Hotkey Activation Scope: chooses between global hotkeys or hotkeys active only when an EVE window is focused.
* Cycle Login Screens - Hotkey: cycles through login-screen windows.
* Login Screen Cycle Direction: chooses old-to-new or new-to-old order.
* Close Active EVE Window - Hotkey: closes the currently active EVE window.
* Close All EVE Windows - Hotkey: closes all EVE windows.
* Reload EVE-X-Preview - Hotkey: reloads the program.
* Character Name: character bound to the hotkey entry.
* Hotkeys: hotkeys assigned to the character.

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
* Thumbnail Text Margins (px): text padding inside the thumbnail.
* Client Highlight Color (Hex/RGB): border color for the active client.
* Client Highlight Border Thickness (px): thickness of the active-client border.
* Show Client Highlight Border: shows or hides the active-client border.
* Thumbnail Opacity (%): thumbnail transparency.
* Show All Borders: shows borders around every thumbnail.
* Inactive Client Border Thickness (px): thickness of inactive-client borders.
* Inactive Client Border Color (Hex/RGB): border color for inactive clients.
* Thumbnail Background Color (Hex/RGB): background color used when thumbnail content is not available.
* Minimum Thumbnail Size (px): minimum allowed thumbnail size.

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

Use this group to assign per-character colors.

* Custom Colors Active: enables or disables custom colors.
* Character Name: character this color profile applies to.
* Active Border Color: custom border color for the active client.
* Text Color: custom text color.
* Inactive Border Color: custom border color for inactive clients.

### Game Logs Monitoring

This feature does not violate the EULA or Terms of Service, but it is neither endorsed nor supported by CCP. Use it at your own risk.
Performance impact increases with the number of characters tracked at the same time.

* Game Logs Monitoring Enabled: turns log monitoring on or off.
* Game Logs Directory: path to the EVE log folder.
* Monitoring Interval (ms): polling interval.
* Event Priority if Several Incoming: chooses whether the first or last event wins.
* Supress for Focused Window: suppresses events for the active window.
* Show Event Text: shows event text in the UI.
* Flash Border: flashes the border on an event.
* Stop Displaying Event on Switch: stops showing the event after switching.
* Event Display Duration (ms): how long the event remains visible.
* Flash Border Interval (ms): flash timing (if `Flash Border Interval` = `Event Display Duration` border will not flash, but just stay with set color for specified time).
* Monitor Only Selected Characters: limits monitoring to selected characters.
* Characters to Monitor: list of monitored characters.

### Monitored events

Monitored events for Game Logs Monitoring feature.
Enabling more events reduces performance. The effect becomes larger as more characters are tracked.

* Stopped Shooting
* Under Attack By Player
* Under Attack By NPC
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

Buttons:
* Check Updates: click to checks aviable updates.
* Update to Release: after update checking can use for update/downgrade to latest release version.
* Update to Pre-Release: after update checking can use for update/downgrade to latest pre-release version.
* Help: opens this README page.
* Report Bug: opens [GitHub issue](https://github.com/khivus/EVE-X-Preview/issues/new) page.

### EVE-X-Preview

* Created by [gonzo83](https://github.com/g0nzo83/EVE-X-Preview)
* Forked and maintained by [khivus](https://github.com/khivus)
* Credits to [mrmjstc](https://github.com/mrmjstc/eve-apm-preview) and [CJKondur](https://github.com/CJKondur/EVE-MultiPreview)
