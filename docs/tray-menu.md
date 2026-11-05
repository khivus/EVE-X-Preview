# Tray menu

[Back to README](../README.MD) · [Settings overview](settings.md)

**Open** and tray-icon double-click follow the global **Settings editor** preference. **Open internal settings** provides direct access to the AHK workbench. Complete or discard any protected draft/dialog before changing editors. See [settings editors](modern-settings.md) for companion compatibility and fallback.

Open Settings from the tray's **Open** command. Configure optional commands under **Application → Tray menu commands**. Checked rows control whether commands appear in the menu; they do not activate the features themselves.

| Command | Effect |
| --- | --- |
| Select profile | Activates a profile and restarts, or stages activation when Settings is open |
| Suspend Hotkeys | Toggles app shortcuts |
| Hide Thumbnails | Toggles thumbnail visibility |
| Lock Thumbnail Positions | Toggles protection against moving or resizing thumbnails |
| Anonymize Thumbnails | Toggles anonymous thumbnail labels |
| Show Thumbnails Always on Top | Changes whether thumbnails stay above other windows; requires restart |
| Click Through Thumbnails | Toggles passing mouse input to the windows underneath |
| Minimize Inactive Clients | Toggles minimizing EVE clients after switching away |
| Don't Close Active Client | Toggles focused-client protection from Close all |
| Close all EVE Clients | Closes EVE clients except those protected by the configured exceptions |
| Restore Client Positions | Restores saved EVE window positions and sizes |
| Save Client Positions | Saves current EVE window positions and sizes |
| Auto Save Thumbnail Positions | Toggles automatic saving of thumbnail positions and sizes |
| Save Thumbnail Positions | Saves current thumbnail positions and sizes |
| Recover off-screen thumbnails | Brings partially or fully off-screen thumbnails onto the nearest connected display, preserving visible placements across all monitors |
| Pause/Resume Live Previews | Immediately pauses or recreates live previews; switching hotkeys remain available |
| Reload | Restarts EVE-X-Preview |
| Exit | Closes EVE-X-Preview |

Recover off-screen thumbnails and Pause/Resume Live Previews are always available. The pause/resume caption follows the current preview state. Locking and anonymization entries are hidden by default.

An open draft is protected from Reload and Exit: apply or discard it first. Selecting a tray profile while Settings is open stages activation in Profiles. Finish any open row or event editor before changing profiles. Other immediate runtime changes, including saved positions, are preserved when an unrelated draft is applied.
