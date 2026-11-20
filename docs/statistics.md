# Statistics

Open **Statistics â†’ Usage statistics** to view saved totals, filter them, or scan archives manually. Those controls work with live collection off. To track new activity while EVE-X-Preview runs, turn on **Collect live statistics**, then apply and restart. Lifetime counters are saved under `%LOCALAPPDATA%\EVE-X-Preview\Statistics` about once per minute when they change and on normal exit.

The switching combo uses confirmed changes of focus. A gap greater than five times the current profileâ€™s group shortcut hold delay starts a new combo. Character launches count thumbnail creations, including thumbnails created again after a restart. The saved character count deduplicates names across profiles, thumbnail and client positions, character shortcuts, groups, colors, custom names, protection lists, visibility, and monitoring settings.

## Log collection

With live collection on, the app records switching and launch activity and uses the enabled alert, DPS, and Local monitoring features to count observed records. **Collect all supported log events live** also works with alerts or log monitoring off. In that mode a background worker checks new and changed game and chat logs about once a minute and indexes supported records from the time the app started. Manual archive scans fill in earlier records without double-counting the live results. The **Live monitoring** card shows the appâ€™s observed totals separately from the indexed game and chat records.

Choose **Stop archive scan** to save progress and pause scheduled scans for the current run. **Scan or resume log archives** continues from saved file positions and checks for new files or appended records. The dashboard shows when it loads checkpoints, finds files, scans records, and prepares totals. Once the file count is known, it shows the percentage, unreadable files, and estimated time remaining. Filters use the latest completed dashboard snapshot and remain available during later scans.

The worker writes its phases, counts, and errors to `%LOCALAPPDATA%\EVE-X-Preview\Statistics\scan.log`, with one rotated copy. The scan log contains no chat text or game records.

## Read the dashboard

Select a character and a date range. Dates are in EVE's UTC log time. **Last session** selects the latest indexed game log for each selected character and excludes chat. The 7-, 30-, and 365-day views include today. App counters such as switching and launches are lifetime totals.

Monitored files and archive scans share one index, so rereading a file through a different collection mode does not add its records again. A copy in a separate file or the same chat message in multiple clients' logs can still appear more than once. The first scan on this build prepares the compact dashboard snapshot from saved checkpoints. Older files with totals only are read once to populate date and character filters.

Peak incoming and outgoing DPS use each character sessionâ€™s highest rolling 10-second damage total divided by 10. Peak alpha strikes track the largest single logged hit in each direction. The all-characters view shows the highest individual peaks. A window that spans midnight belongs to the UTC day of its last hit. Run an archive scan once to calculate the new combat peaks for previously indexed logs.

Mining totals come from explicit quantities in the logs and show units of each material, including harvested gas and named residue. Separate unattributed residue messages are shown on their own because they do not name a material. The index stores file paths, character names, offsets, and aggregate counts. It does not store chat text.

## Logistics, fleet activity and locations

Logistics shows armor repaired, shields boosted and hull repaired (including records named structure repairs), separately for amounts given and received, plus the number of logistics records (including zero-amount records). Fleet activity counts “Following … in warp”, “Regrouping” / “Regrouping to …”, and exact “Please wait...” notifications. These are logged records rather than deduplicated fleet commands.

The most appearing module counts named weapons/drones in combat hit records, remote repair/boost modules, and modules named in “External factors are preventing…” notifications. This measures appearances in supported log messages, not activation cycles.

The top 10 systems counts “Channel changed to Local” entries in chat logs, so it measures recorded visits, not time spent in a system. Repeated visits and the same visit recorded by several characters each count. Last session excludes chat, so system rankings are empty for that filter.

With All characters selected, each combat peak shows its winning character independently; the module and system rankings show the character with the most records for that entry. Tied characters use the first indexed character. Character and UTC date filters apply to all these statistics. Run Scan or resume log archives once to backfill older checkpoints; the scan replaces their aggregates without adding duplicates. Only aggregate names and counts are saved, not chat messages.
