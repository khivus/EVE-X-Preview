# Character list quick checks

- Settings > Thumbnails > Character list: settings run from visibility through layout, ordering and window behavior. Grouping options are gone.
- Open the three-dot menu: no error; topmost, lock, autosave, opacity, direction and automatic filling still work.
- Change fill direction with overflowing entries: the visible themed scroll thumb changes from vertical to horizontal. Drag it and use the wheel to reach all entries.
- Disable automatic filling and choose two or more columns/rows: new entries stay in the first. Drag a character into an empty column/row; its assignment should survive client updates and restarting.
- Drag above/below a row or left/right of an entry in horizontal mode: the visible line marks the exact insertion gap. Release: entries insert without replacing another character. Use the wheel while dragging to reach an offscreen gap.
- Switch focus, resize, collapse and change themes: no white strip, no recreated-window flashing, and only the selected row background/bold text highlight appears beside character names.
- Trigger configured monitored events with thumbnails visible and hidden: event text appears only when Show event text is enabled; row borders flash with the configured event color/interval only when flashing is enabled. Duration, focused-client suppression and clearing on switch follow event settings.
- Verify profile inheritance, list position/size persistence, independent thumbnail visibility and existing switching shortcuts still work.

Live EVE events, pointer behavior, focus and visual appearance require manual verification.
