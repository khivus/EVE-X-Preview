# Thumbnail behavior

[Back to README](../README.MD)

## Thumbnails Interactions

Quick Group Sort Order chooses how the Quick Group hotkey cycles through selected windows. Added order (the default) follows the order windows were added to Quick Group; Name cycles alphabetically by character name. Removing and adding a window again places it at the end of added order.

**Keep Quick Group on Logout** and **Keep Disabled Status on Logout** are enabled by default. Membership follows the same client window through character selection and relogin, including switching characters. Turn either option off in Thumbnails Interactions to clear that status on a character change or logout. Closing the client clears its membership; restarting the client or EVE-X-Preview starts fresh.

**Don't Close Quick Group Clients** and **Don't Close Disabled Clients** protect those windows from **Close all EVE Clients**, including while on the login screen, independently of **Don't Close Clients on Login Screen**. Directly closing an individual client still works.

## Thumbnails Behavior

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

## Thumbnail Visibility

This group controls whether selected client thumbnails are hidden.

* Select Any Client to Hide The Thumbnail: choose a client to hide.

## Lock positions

Open **Thumbnails Behavior** and enable **Lock thumbnail positions and sizes**, beside the default-size and snap settings. This blocks move, move-all, resize, resize-all and snap gestures, including the test thumbnail. Activation and other click actions keep working. Saved positions still restore and new clients still receive their configured placement. Add **Lock Thumbnail Positions** to the tray through **Tray Menu Settings** for a quick toggle.

Unsaved logged-in thumbnails are checked again when a login window becomes a character. Collision avoidance reserves saved positions for other clients awaiting preview startup. If a later login restores a saved rectangle over an unsaved thumbnail, the unsaved thumbnail moves to an available grid position. Saved positions remain authoritative; intentionally overlapping saved layouts are preserved. A full configured grid is reported in the program log and does not disable future placement attempts.
