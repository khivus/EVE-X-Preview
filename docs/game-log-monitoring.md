# Monitoring, alerts, sounds and DPS

[Back to README](../README.MD) · [Settings overview](settings.md)

## Sources and systems

Open **Monitoring & alerts → Log sources** to enable game-log reading, choose game/chat folders, set the scanning interval or enable the character filter. Blank folders use automatic detection; the detected paths and folder availability are shown below. **Monitored characters** uses a row list instead of a text block.

**Solar systems** separately enables current-system tracking through Local chat logs and consistent system colors. Discovery runs in hidden workers with bounded reads, filename notifications and fallback scans. Large log archives do not block the settings interface. Monitoring changes remain staged until Apply & restart.

## Alerts

**Alert behavior** controls first/latest event priority, visual focused-client suppression, sound focused-client muting, event text, border flashing, clearing on switch and durations. Sound muting and **DPS → Hide DPS for the focused client** are independent preferences. Existing effective suppression choices are preserved during migration.

**Event alerts** lists combat, travel, fleet/communication and mining events with labelled color, sound and volume columns. Click its Enabled cell to toggle it, or fill selected Enabled cells with 1/0. Color, sound path and volume also edit inline. Event details opens one event editor with its detection rule, color picker, sound selection, preview and event volume. Player/NPC neutralization and misses, the player smartbomb exclusion and the shooting-idle interval are edited with the event they affect.

Save to draft commits the complete event edit to the workbench; Cancel leaves it unchanged. Apply & restart activates it. Source/prerequisite messages explain when game-log reading is disabled or event settings are inherited.

## Sounds

**Sounds** contains shared master volume and the shared library. Add a WAV or MP3 file, then select it in an event editor. Preview file uses master volume × file gain; event preview also includes the event's volume. Sound files are referenced in their original locations, not copied or modified. Missing files are reported rather than replaced with a beep.

Master volume, advanced file gain and event volume each accept exact whole percentages from 0 to 300. Their gains multiply: 50% × 50% × 50% = 12.5%. Zero mutes playback; strong boosts can clip already-loud audio. Manual previews use the draft and update with draft master/file gain. They do not save preferences or change other applications' volume.

Advanced file gain affects every event using that file across profiles. **Batch assignments** uses a separate dialog: checked events use the selected file; unchecking removes that file's assignments while leaving other files alone. Save to draft is explicit and Cancel discards the batch. Removing a library reference preserves assignments, gain and the original file, with Undo available.

**Advanced playback** contains the shared simultaneous-sound limit (1–16) and repeat timeout (0–3600000 ms). At the concurrency limit, the oldest sound is replaced. Repeat timeout is shared by the same file across characters/events; zero allows repeats. Manual previews ignore that timeout. These preferences become active with Apply & restart.

## DPS

Enable game-log reading in **Monitoring & alerts → Log sources**, then enable incoming DPS, outgoing DPS or both under **DPS**. DPS shares the game-log directory, monitoring interval, and selected-character list. It works with event text and flashing borders disabled. The thumbnail text overlay must be enabled to see the values.

Set **DPS averaging period** to choose the rolling averaging period for both directions (1-3600 seconds, default **10**). Each value is the damage recorded during that window divided by its duration, preserving the spacing of the log's timestamps. For artillery firing every 17 seconds, a 20-second or longer window keeps the previous volley included between shots. The live window aligns to newly appended combat records and advances using elapsed time, so a difference between Windows UTC and EVE's clock does not discard hits. Thresholds are the minimum DPS to display for each direction; 0 means no minimum. A value below its threshold or with no recent damage is hidden. Old log contents are skipped when monitoring starts. Neutralization, repairs, misses, and other notifications do not count as damage. Event-specific NPC/player and smartbomb filters do not filter DPS.

**Show estimated incoming damage mix** estimates the EM, thermal, kinetic, and explosive mix only for drones, missiles (including rockets and torpedoes), fighters, and smartbombs. Missiles use their weapon-family names: Mjolnir = EM, Inferno = thermal, Scourge = kinetic, and Nova = explosive. Smartbombs use EMP = EM, Plasma = thermal, Graviton = kinetic, Proton = explosive, and Multispectrum = equal parts of all four. Drones, fighters, and smartbomb naming exceptions use exact names in the log. The bundled [EVE Ref data](https://docs.everef.net/datasets/reference-data) includes mixed-damage profiles. The display uses `?` for the percentage of damage with an unknown type. Other weapons still count toward total incoming DPS but remain unknown in the damage mix. Percentages include unknown damage in the denominator. They weight each known weapon's base damage mix by its recorded hit damage, so they are estimates, not measured post-resistance percentages or the ship's resistances. The existing saved setting key `showIncomingDPSResistances` is retained for compatibility.

DPS and the optional damage mix each occupy one line above event/system text: `DPS: In: 123 | Out: 456` and `% EM 25 Th 75`. Displayed DPS is rounded to whole numbers; calculations and thresholds retain full precision. Damage types that round to 0% are omitted, including unknown damage (`?`). The character name keeps priority and the system stays at the bottom. Active DPS takes priority over event text on compact thumbnails; the percentage breakdown appears when there is enough room. Empty event text reserves no line. **Hide DPS for the focused client** independently hides its DPS while continuing to count damage, so switching focus does not create gaps. Event expiry and event switching do not clear DPS.

Monitoring reads only appended, complete lines through the existing reader and timer. DPS is updated before event visuals, and alert updates are batched once per character per poll while preserving first/last event priority. Per-second totals keep memory bounded; damage profiles are shared and loaded only when the optional damage mix is enabled. No network access is used while monitoring, and unchanged display text is not rewritten.

Volume sliders show ten scale divisions and snap to steps of five while dragging. Numeric fields retain exact whole percentages such as 127%. File paths and gains also support inline editing in Sounds.
