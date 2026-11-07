# Settings review: focused recheck

Build: **v1.7.0.4**, November 7, 2026. The output folder contains only the main, Settings and updater executables. No ZIP was created.

The shared styling and navigation changes affect the companion throughout. Functional changes are concentrated in profile drafting, mouse gestures, inheritance, monitoring category ownership, sound assignments and preview diagnostics. The internal editor shares the monitoring schema/migration, shorter Advanced label and canonical command state.

These checks cover the changed paths; repeating every setting is unnecessary:

1. **Themes and controls:** switch between Light, Graphite and a colored theme. Header, footer, sidebar, inputs and table separators should update immediately. Hover an enabled Graphite toggle: its knob should remain distinct. Check the centered search X, even disabled numeric backgrounds and visible read-only color swatches.
2. **Create profile:** scroll its grouped inheritance choices. The typed profile name, presets and Create/Cancel buttons should stay visible.
3. **Navigation and scroll:** switch several categories through the sidebar/search. Search, profile selector and footer should stay enabled without the former disabled-page flash. In Section inheritance, scroll midway and toggle a checkbox; it should retain the current position. The shorter category name is Advanced.
4. **Mouse gestures and Default editing:** from an inheriting profile, choose Edit Default here. Change several actions independently, including entries that were all Disabled, and change them back. Each dropdown should affect only its own action even if a range was selected. Create a duplicate, then restore the original gesture and leave the category. The duplicate problem should clear and the selected profile should remain selected. Explicit clipboard paste can still fill a selected range.
5. **Monitoring:** Monitored characters should contain the selection toggle and list under one inheritance source. Log sources should retain the game-log reading toggle and left-aligned game/chat folder fields. If a profile previously inherited the filter and list differently, check that its effective values remain the same; the migration makes this category local to preserve them.
6. **Disabled settings:** an inherited category should explain its source once at the top. A setting disabled by another toggle should have a category-level explanation naming that toggle. Disabled color swatches should show their actual color and reject editing.
7. **Sound library:** inherited event assignments should explain why editing is unavailable, with Edit Default sound assignments here available on the same page. Check master volume, file gain and event-volume slider/value alignment. Scroll down and click Stop sound preview; the page should stay in place.
8. **Diagnostics with the main app running:** click Pause live previews, then Resume live previews. The button should name the next action and thumbnails should follow it. A pause changed elsewhere should be reflected while Diagnostics is open.

Automated checks and their limits are recorded in [modern settings verification](modern-settings-verification.md). Real EVE playback/monitoring and physical mixed-DPI behavior still need the normal live-session spot check.
