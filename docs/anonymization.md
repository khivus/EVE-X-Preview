# Anonymous thumbnail labels

[Back to README](../README.MD)

Open **Thumbnails Visuals** and enable **Anonymize thumbnail labels**, beside the other text-overlay settings. You can also add **Anonymize Thumbnails** to the tray in **Tray Menu Settings**.

Edit the **Random names** box on that same page: one possible name per line. Names are selected randomly without reuse, ignoring blank and duplicate entries. If the list runs out, unique `Pilot 1234` labels are generated. The selected label stays stable for the session; **Reshuffle names** or a restart creates a new assignment. Turning anonymization off restores your original/custom labels. Newly opened thumbnails are covered automatically, including separate labels for login windows.

The list and enabled state are saved per effective **Thumbnails Visuals** profile in `EVE-X-Preview.json`, under `_Profiles → <profile> → Thumbnails Visuals → AnonymousNames` and `AnonymizeThumbnails`. Settings from the former Thumbnails Behavior location are migrated automatically. To change the shipped default list for new settings, edit `Lib/DefaultJSON.ahk` (`AnonymousNames`). Existing settings keep their own list. Scroll the right pane to reach the random-name list.

This feature changes the character label only. It does not rename EVE windows, overwrite custom aliases or redact game images, system/event text, settings, tray profile names or logs. Use **Exclude Thumbnails from Capture** if thumbnails must be omitted from compatible capture software.
