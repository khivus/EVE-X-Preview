# Anonymous thumbnail labels

[Back to README](../README.MD) · [Settings overview](settings.md)

Open **Thumbnails → Labels** and enable **Use anonymous thumbnail labels**. The random-label pool becomes editable; enter one name per line. An empty or exhausted pool uses Pilot plus a number. Labels stay stable during the runtime session.

**Reshuffle preview labels** opens a fresh test thumbnail using the draft. Applying restarts EVE-X-Preview and assigns a fresh runtime set. Turning anonymization off restores the normal/custom labels. Matching clients, shortcuts and game logs continues to use original character names.

The pool and enabled state follow the appearance/labels inheritance source. Existing `AnonymousNames` and `AnonymizeThumbnails` values under the legacy Thumbnails Visuals JSON group are preserved. The former Thumbnails Behavior location is still migrated.

The optional immediate **Anonymize Thumbnails** tray command can be shown through **Application → Tray menu commands**.
