# Profiles

[Back to README](../README.MD) · [Settings overview](settings.md)

## Editing and activating

The sidebar selector chooses the **profile to edit**. It does not activate that profile. Open **Profiles → Profile management** to create, rename or remove profiles and record an activation shortcut.

**Activate this profile on Apply** stages the running-profile change. The page shows both the current running profile and the profile to activate. **Apply & restart** commits the draft and activates the staged profile. Selecting another editing profile preserves the existing draft.

Creating a profile can copy the current profile's effective settings or start from application defaults. Before creating it, choose **Use inheritance defaults**, **Disable all** or **Enable all**, then adjust individual subsection rows. The preset buttons change which sections inherit; the starting-values choice supplies the custom sections. Copying effective settings retains what you currently see even when you disable inheritance. Its activation shortcut starts empty. Renaming preserves assignments. Default cannot be renamed or removed. Removed profiles can be restored with Undo before Apply.

Activation shortcuts are independent of section inheritance and use the same recorder/clear/Advanced workflow as other shortcuts. They are available without EVE, follow Suspend Hotkeys and pause while Settings is focused. If a profile is activated from the tray or a shortcut while the workbench is open, activation is staged in the workbench so the existing draft remains intact.

## Inheritance

Relevant subsections display **Settings source**:

- **Use Default profile settings** shows the effective Default values and explains why inherited controls are read-only.
- **Customize for this profile** uses that profile's stored values.

**Edit Default here** unlocks the inherited subsection's Default values in the same view. The sidebar keeps your selected profile, while the pinned heading says **editing DEFAULT** and a warning explains that other inheriting profiles are affected. This works for fields, shortcuts, groups, tables and event editors. **Done editing Default**, selecting another subsection, or selecting another profile leaves this mode; edits stay in the shared draft. Apply once after preparing all profiles. Editing Default does not activate it or change your staged activation profile.

Each visible subsection has its own source. For example, Appearance, Labels, Layout, Automatic placement and Visibility can inherit independently. Character exceptions follow Close protection; application groups and individual application shortcuts are separate. **Profiles → Section inheritance** lists every inheritable subsection and supports selecting and changing several source cells together.

With **Default** selected, this view edits the inheritance template for **new profiles**. A Yes row means the subsection inherits when the creation dialog uses **Use inheritance defaults**. Existing profiles keep their source choices. The template starts with inheritance disabled until you choose otherwise and is saved with the draft on Apply. Other profiles offer the same three bulk presets for their source rows.

Changing a source retains that profile's stored custom values. Default edits affect only the subsections that inherit them. Shared sound playback settings and application-wide diagnostics have no profile source. Profile activation shortcuts always belong to their own profile.

Migration retains the previous effective values. If all settings in a new subsection previously inherited, it keeps inheriting. A subsection that previously combined Default and local settings becomes custom with the same effective combination. Any dormant local values replaced during this conversion are backed up under `SettingsMigrationBackup` in the profile configuration. The original JSON setting paths are retained, and the migration is saved only when you Apply.

Character shortcuts and character-group definitions remain local for legacy profiles that previously kept those assignments local. Focused-client sound muting follows Alert behavior; focused-client DPS visibility follows DPS. Their initial values preserve the old behavior. Saved thumbnail and EVE-window positions remain specific to the editing profile.

Quick Group's duplicated gesture control edits the same assignment as Mouse actions. Its source is explicitly labelled **Mouse actions** on the Quick Group page, independently of the Quick Group behavior settings.
