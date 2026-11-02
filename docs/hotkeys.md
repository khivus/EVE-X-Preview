# Hotkeys

[Back to README](../README.MD)

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

## Hotkey Groups

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

## Hotkeys Settings

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

## Non-EVE Applications

This section lets you add non-EVE apps by exact executable name, and optionally by window title for better matching.
The thumbnail appears for the app, and the assigned hotkey can activate it.

* Add/Delete Groups: create or remove app groups.
* Select Group: choose the app group to edit.
* Forwards Hotkey: hotkey for forward group cycling.
* Backwards Hotkey: hotkey for backward group cycling.
* Process Names (.exe): executable names to match.
* Titles (optional): optional window titles to match.
* Hotkeys: hotkeys assigned to the app group.

## Automatic wildcard setting

Open **Hotkey Groups** and turn off **Add * to Switching Hotkeys**, at the top of the page, for exact modifier matching. It is on by default for compatibility. This covers character, group, Quick Group, previous-window, login-screen and other-app switching shortcuts. Explicit `*` prefixes remain effective. Custom `key & key` combinations keep AutoHotkey's wildcard behavior. Saved bindings are never rewritten; duplicate warnings follow the selected policy. Close Settings to re-register shortcuts.
