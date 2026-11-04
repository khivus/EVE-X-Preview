# Keyboard shortcuts and groups

[Back to README](../README.MD) · [Settings overview](settings.md)

## Assign shortcuts

Open **Switching → Keyboard shortcuts** for application commands and window switching, or **Character shortcuts** for individual characters. Shortcuts display readable key names.

- **Record** captures a keyboard combination. Mouse buttons and the wheel can be recorded over the gray recording field.
- **Clear** removes the assignment.
- **Advanced** unlocks the existing field for AutoHotkey syntax, such as `^!F1` for Ctrl+Alt+F1. **Done** returns to readable key names. Invalid text stays inline and blocks Apply until fixed or cleared.
- Escape, pressing Record again, or leaving recording cancels it.

Conflicts include effective bindings in the editing profile and every profile activation shortcut. **Review shortcut conflicts** lists overlapping assignments. Intentional overlaps remain allowed. Invalid key names are rejected before a row or setting can be saved to the draft.

A character shortcut may share a key with a character-group or Quick Group shortcut. The direct shortcut takes priority while its character window is available; otherwise the group shortcut can run. These supported pairs do not produce conflict warnings.

A plain-key switching shortcut can also coexist with a more-specific modifier combination, such as backtick for a group and Alt+backtick for cycling login screens. Automatic wildcard matching does not make this pair a conflict. The checker still warns for duplicate modifier chords, key aliases and ambiguous overlapping combinations; it does not change shortcut registration.

AutoHotkey syntax supports `^` (Ctrl), `!` (Alt), `+` (Shift), `#` (Windows), `*` (extra modifiers) and custom `key & key` combinations. See the [modifier reference](https://www.autohotkey.com/docs/v2/Hotkeys.htm#Symbols) and [key list](https://www.autohotkey.com/docs/v2/KeyList.htm).

## Character groups

Open **Switching → Character groups**. Create or select a group, add members and record its next, previous and first-available-character shortcuts. Members are stored as rows in switching order. Move up/down keeps each member intact. Duplicate group names are rejected; removal offers Undo.

**Switching behavior** contains keeping shortcuts linked after logout, remembering the switching position in each group, and the minimum held-shortcut interval. **Advanced switching** contains bounded focus retries and their interval; one check skips waiting.

## Quick Group

Configure Quick Group membership using its gesture in **Thumbnails → Mouse actions**. **Switching → Quick Group** contains its shortcut, order, border color, interaction with other groups and logout persistence. Membership follows the same client window through logout/relogin when preservation is enabled; closing the client or restarting EVE-X-Preview starts fresh.

Close-all protection for Quick Group and excluded characters is in **EVE windows → Close protection**.

## Scope and modifier matching

**Keyboard shortcuts → EVE switching shortcuts are available** selects global switching while EVE is open or switching only while EVE is focused. **Allow extra modifier keys** controls automatic `*` prefixes for character, group, Quick Group, previous-window, login-screen and other-app switching shortcuts. Explicit `*` and custom combinations retain their own behavior. Stored shortcut text is not rewritten.

Other-application shortcuts remain available whenever their matching window is open. Runtime shortcuts pause while Settings or one of its dialogs has focus.

All assignments and scope changes remain in the draft until **Apply & restart**.

Quick Group's thumbnail gesture is also editable directly in **Switching → Quick Group**. It shares the same stored assignment as **Thumbnails → Mouse actions**. A link opens the full mouse-action table, and a warning explains when click-through prevents the gesture.
