# Thumbnail appearance

[Back to README](../README.MD)

## Thumbnails Visuals

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

## Custom Colors

Every color field supports the Windows color picker. Click the colored swatch beside a single color value. For a per-character color list, select the target line and click its picker button. The picker starts with the current color, and Cancel leaves it unchanged. Hex/RGB values can still be entered manually.

Use this group to assign per-character colors.

* Custom Colors Active: enables or disables custom colors.
* Character Name: character this color profile applies to.
* Active Border Color: custom border color for the active client.
* Text Color: custom text color.
* Inactive Border Color: custom border color for inactive clients.
