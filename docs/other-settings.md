# Application and advanced settings

[Back to README](../README.MD) · [Settings overview](settings.md)

**Application → Advanced** contains performance and compatibility options:

- **Time between window and focus checks:** 50–1000 ms. Higher values reduce window polling and delay discovery/focus updates.
- **Stagger preview creation at startup:** waits three seconds, requires stable source dimensions and spaces preview creation to reduce startup GPU bursts. It also affects new clients and resume operations.
- **Pause previews if Windows composition fails:** releases preview resources after rendering/composition failures. Resume explicitly from the tray, a configured shortcut or Diagnostics. Disabling the preference does not resume an existing pause.
- **Retry failed shortcuts using an English keyboard layout:** a registration compatibility fallback, not an interface-language setting.

These preferences are draft settings and activate with **Apply & restart**. Profile inheritance is managed in Profiles. Thumbnail size resets are in Thumbnails → Layout.

**Diagnostics** contains staged debugging/counter preferences and explicitly immediate log/export/preview actions. **About and updates** contains version information, offline release notes, documentation, issue reporting and installers. Older installers are labelled as downgrades and require confirmation; apply/discard the draft first.
