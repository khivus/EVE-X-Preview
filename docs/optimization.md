# Thumbnail optimization

[Back to README](../README.MD)

Open **Application → Advanced** for **Time between window and focus checks**. Changes activate with Apply & restart.

| Setting | Default | Effect |
| --- | --- | --- |
| Window polling interval | 50 ms | Increase up to 1000 ms to reduce application window checks. Focus, title and client discovery updates may arrive later. This does not cap live FPS or change log-monitoring intervals. |
| Hide Thumbnail for Active Window | Off | Available in Thumbnails → Visibility. Avoid showing a duplicate of the client you are using. |
| Hide Thumbnails on Lost Focus | Off | Available in Thumbnails → Visibility. Makes DWM thumbnails invisible while working in other apps. |

A 100–250 ms polling interval reduces window checks, with correspondingly slower focus/title updates. These are adjustable tradeoffs, not guaranteed performance gains. FPS remains controlled by Windows DWM and the source client.

## Why there is no live FPS slider

The [DWM thumbnail properties API](https://learn.microsoft.com/en-us/windows/win32/api/dwmapi/ns-dwmapi-dwm_thumbnail_properties) exposes source/destination rectangles, opacity, visibility and client-area selection. It has no per-thumbnail frame-rate or capture-resolution field. Calling the update method less often changes property updates, not the compositor's live refresh rate. [DwmSetPresentParameters is unsupported on Windows 8.1 and later](https://learn.microsoft.com/en-us/windows/win32/api/dwmapi/).

A true FPS or internal-resolution selector would require another capture/rendering backend, such as Windows Graphics Capture with a frame pool and a renderer. That would add GPU resources, copying, device-loss handling and compatibility work; it is not exposed as a pretend optimization for the existing DWM backend.

## Other available measures

Disable unused system/event/DPS monitoring to stop the corresponding work, increase its own log interval if acceptable, or use **Pause/Resume Live Previews** from the tray, Application → Diagnostics or a configured shortcut to release preview resources until resumed. **Minimize Inactive Clients** may reduce client rendering workload, depending on EVE settings, and can affect source updates. **Stagger preview creation at startup** staggers startup DWM calls; it addresses startup bursts rather than steady-state FPS. Hiding previews changes visibility but retains their registrations; pausing releases them.
