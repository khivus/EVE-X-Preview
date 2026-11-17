# Thumbnail layout and visibility

[Back to README](../README.MD) · [Settings overview](settings.md)

## Layout

**Thumbnails → Layout** contains automatic position/size saving, layout locking, default position/size, minimum size, snapping and preserving placement after logout. Negative screen coordinates support monitors to the left or above. Related dimensions are validated together.

**Reset this profile's thumbnail sizes on Apply** stages the configured default dimensions for this profile's saved thumbnails and current windows. Positions are retained. Apply & restart activates the reset.

**Recover off-screen thumbnails** immediately brings partially or fully off-screen thumbnails back onto the nearest connected display and saves their corrected positions. It checks the combined area of all connected monitors, so thumbnails on a second or third monitor, or spanning adjacent monitors, stay in place when fully visible. Recovery also works while positions are locked and moves hidden thumbnails without showing them. Sizes are preserved; a thumbnail larger than every display is placed with its top-left corner on screen.

Locking blocks move, move-all, resize, resize-all and snapping gestures, including in the test thumbnail. Activation and other click actions continue to work. Saved placements still restore. The optional tray lock command is configured under **Application → Tray menu commands**.

## Automatic placement

**Automatic placement** controls login-screen thumbnails and logged-in characters without saved positions independently. Choose a direction, collision avoidance and horizontal/vertical steps; zero uses the thumbnail's width/height. Saved positions take priority.

A login screen becoming a character is checked again for placement. Saved positions of clients waiting for preview startup are reserved. A later saved rectangle can reflow an unsaved thumbnail to a free grid position. Intentionally overlapping saved layouts are kept. A full grid is logged and does not prevent future placement attempts.

## Visibility and mouse actions

**Visibility** contains hiding all thumbnails, hiding while other apps are focused, hiding the active client's thumbnail, keeping thumbnails above other windows, capture exclusion and retaining the last character label after logout.

**Hidden characters** is an explicit list: listed characters have no visible thumbnail. Add/remove complete rows, or choose running EVE clients.

**Mouse actions** shows each action and its complete gesture instead of a modifier checkbox matrix. Edit a row to choose a gesture or Disabled. Duplicate gestures are rejected. Click-through and layout-lock notices explain when actions cannot run.

Focused-client event sounds and DPS have separate controls under **Monitoring & alerts**. Hiding the active thumbnail no longer implicitly changes the explicit sound preference after settings migration.

Workbench setting changes remain staged until Apply & restart. Thumbnail recovery and tray commands are immediate runtime actions.
