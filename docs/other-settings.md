# Other settings

[Back to README](../README.MD)

* Auto-Pause on DWM Failure: enabled by default, including existing profiles. Automatically pauses live previews on DWM failures or Windows composition changes. Applies immediately; turn it off to keep automatic retries. An existing pause still requires manual Resume.
* Slow Thumbnail Creation: off by default, including existing profiles. Off creates ready previews immediately without startup or spacing delays. On waits three seconds, requires stable source dimensions for 500 ms, and spaces queued previews by 150 ms to reduce startup GPU pressure. Applies immediately to queued previews and controls startup, new clients, and resume after pausing.
* Switch Language to English on Error: switches the Windows language to English if hotkey assigning error occurs.
* Update all thumbnails size to default size: use this button to update every thumbnail size to the current default size.
