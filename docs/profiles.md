# Profiles

[Back to README](../README.MD)

## Profile selector and group navigation

The left panel contains the profile selector and profile management buttons:

* Add profile
* Rename profile
* Delete profile

Select a profile and enter its **Profile hotkey** using AutoHotkey syntax (for example, `^!F1` for Ctrl+Alt+F1). Close Settings to apply. Each profile has its own shortcut, independent of settings-group overrides, and all configured profile shortcuts work even when EVE is closed. Pressing one activates that profile through the same save-and-reload action as the tray Profiles menu. Profile shortcuts pause while Settings is focused and follow Suspend Hotkeys. Leave the field empty to remove a shortcut, and use a different key combination for each profile and other action. Copying a profile's settings leaves the new profile's shortcut empty; renaming preserves it.

It also contains buttons for settings groups. When you select a group, the main window shows the settings for that group.

## Profile override behavior

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
