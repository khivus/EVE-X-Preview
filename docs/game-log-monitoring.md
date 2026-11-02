# Game logs, events and DPS

[Back to README](../README.MD)

## Game Logs Monitoring

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

## Monitored events

Monitored events for Game Logs Monitoring feature.
Enabling more events reduces performance. The effect becomes larger as more characters are tracked.

Click the **♪** button beside any event's color to open **Event Sounds**. **Browse / Add...** adds a user-provided WAV or MP3 file to the shared **Sound library**. Select a library sound, check its events on the right, then click **Apply assignments**. **Select all** prepares to assign one sound to every event in one operation; **Clear checks** prepares to remove this sound from its events. Existing assignments appear in the Current sound column. Before Apply, an orange notice identifies replacements and the button shows how many other sounds will be replaced. Replacements are allowed. Unchecked events using a different sound are preserved. Assigning sounds does not enable disabled events.

The event opened with **♪** is prechecked when choosing a sound, so a single-event assignment also needs only Apply. Checkmarks show existing assignments to the selected file; changes are saved only on Apply. **Clear opened event** immediately removes that event's assignment. **Test selected** previews the library sound; **Test opened event** previews the opened event's actual assignment. The button changes to **♫** when a sound is assigned, with its file path in the tooltip. Playback depends on Windows media support. The application includes no sound files and starts with an empty library.

The dialog has three **0–300% volume** controls. **Master volume** affects every application sound across profiles. **Sound volume** affects every event using the selected library file across profiles. Select an event row to edit **Event volume**, which affects only that event in the effective Monitored Events profile. Sliders snap to 5% steps; the adjacent editable percentage fields accept exact whole values, with or without a percent sign. Slider and text values stay synchronized. Empty or invalid edits preserve the previous setting and normalize on leaving the field; values outside 0-300 are clamped. The gains multiply: 50% master × 50% sound × 50% event = 12.5% final volume. Defaults are all 100%; setting any level to 0% mutes it. Values above 100% amplify the audio: 300% is three times the original signal level. The three gains still multiply, so setting all three to 300% gives 2700% combined gain. Strong boosts can clip already-loud audio. Audio is decoded in bounded chunks and mixed with native gain; original sound files are not modified. Volume edits save immediately and adjust clips already playing without changing Windows or other applications. Test selected uses master × sound; Test opened event also includes its event volume.

The library is shared across profiles; sound assignments and event volumes follow the effective **Monitored Events** profile override. **Remove saved** removes only the library entry, preserving existing event assignments, the file's volume and the original file. The library stores file paths rather than copies, so keep selected files in a stable location. Missing or unsupported audio is skipped during monitoring and recorded in the program log; use Browse to select a moved file again. Sound assignments, all volume levels and the library are included in the settings JSON; moving that JSON to another PC does not transfer audio files.

**Simultaneous sounds** sets the maximum number of event sounds playing together, from 1 to 16 (default 1). A new alert replaces the oldest playing alert when the limit is reached; alerts are not queued for later. Lowering the limit stops the oldest excess clips immediately. **Repeat timeout (ms)** sets the minimum interval between starts of the same sound file across all events and characters, from 0 to 3600000 ms (default 2000). Set it to 0 to allow immediate repeats, including overlapping instances when there is room. Both controls are shared across profiles, save immediately, and are independent of visual alert duration. Muted, missing or failed sounds do not consume the repeat timeout. Manual Test buttons use a separate preview player and ignore these event limits and timeouts.

Enable **Game Logs Monitoring** and the event's checkbox for audio alerts. Sounds work with **Show Event Text** and **Flash Border** disabled, and while thumbnails are hidden. They respect event priority and focused-window suppression (including hiding the active client's thumbnail). The Settings test thumbnail's automatic samples remain silent; use the sound dialog's Test buttons to preview audio.

General NPC recognition matches full English names from a bundled EVE Ref dataset, ignoring case. Player names containing NPC name fragments no longer match. The faction, officer and capital events retain their separate lists and priority. See [NPC data and refresh instructions](../src/data/README.md) for coverage and maintenance.

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

## DPS Monitoring

Enable **Game Logs Monitoring**, then enable incoming DPS, outgoing DPS, or both in **DPS Monitoring**. DPS shares the game-log directory, monitoring interval, and selected-character list. It works with event text and flashing borders disabled. The thumbnail text overlay must be enabled to see the values.

Set **Average Window (seconds)** to choose the rolling averaging period for both directions (1-3600 seconds, default **10**). Each value is the damage recorded during that window divided by its duration, preserving the spacing of the log's timestamps. For artillery firing every 17 seconds, a 20-second or longer window keeps the previous volley included between shots. The live window aligns to newly appended combat records and advances using elapsed time, so a difference between Windows UTC and EVE's clock does not discard hits. Thresholds are the minimum DPS to display for each direction; 0 means no minimum. A value below its threshold or with no recent damage is hidden. Old log contents are skipped when monitoring starts. Neutralization, repairs, misses, and other notifications do not count as damage. Event-specific NPC/player and smartbomb filters do not filter DPS.

**Incoming Damage Types (%)** estimates the EM, thermal, kinetic, and explosive mix only for drones, missiles (including rockets and torpedoes), fighters, and smartbombs. Missiles use their weapon-family names: Mjolnir = EM, Inferno = thermal, Scourge = kinetic, and Nova = explosive. Smartbombs use EMP = EM, Plasma = thermal, Graviton = kinetic, Proton = explosive, and Multispectrum = equal parts of all four. Drones, fighters, and smartbomb naming exceptions use exact names in the log. The bundled [EVE Ref data](https://docs.everef.net/datasets/reference-data) includes mixed-damage profiles. The display uses `?` for the percentage of damage with an unknown type. Other weapons still count toward total incoming DPS but remain unknown in the damage mix. Percentages include unknown damage in the denominator. They weight each known weapon's base damage mix by its recorded hit damage, so they are estimates, not measured post-resistance percentages or the ship's resistances. The existing saved setting key `showIncomingDPSResistances` is retained for compatibility.

DPS and the optional damage mix each occupy one line above event/system text: `DPS: In: 123 | Out: 456` and `% EM 25 Th 75`. Displayed DPS is rounded to whole numbers; calculations and thresholds retain full precision. Damage types that round to 0% are omitted, including unknown damage (`?`). The character name keeps priority and the system stays at the bottom. Active DPS takes priority over event text on compact thumbnails; the percentage breakdown appears when there is enough room. Empty event text reserves no line. Focus suppression hides DPS while the client is active but continues counting its damage, so switching focus does not create gaps. Event expiry and event switching do not clear DPS.

Monitoring reads only appended, complete lines through the existing reader and timer. DPS is updated before event visuals, and alert updates are batched once per character per poll while preserving first/last event priority. Per-second totals keep memory bounded; damage profiles are shared and loaded only when the optional damage mix is enabled. No network access is used while monitoring, and unchanged display text is not rewritten.
