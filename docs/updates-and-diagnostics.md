# Updates and diagnostics

[Back to README](../README.MD)

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
