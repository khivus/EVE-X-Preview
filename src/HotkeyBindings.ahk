; Read-only conflict checks: never register, disable or rewrite the user's binds.
class HotkeyBindings {
    static PriorityHelp := "Shared shortcut priority: Direct character > Group Next/Previous > Group First available. A higher-priority action yields when none of its characters are running. Quick Group takes priority over all three while enabled with a running member."
    static SectionFor(id) {
        if InStr(id, "character:") = 1
            return "character-shortcuts"
        if InStr(id, "group:") = 1
            return "character-groups"
        if InStr(id, "nonGroup:") = 1
            return "app-groups"
        if InStr(id, "nonApp:") = 1
            return "app-shortcuts"
        if InStr(id, "profile:") = 1
            return "profiles"
        return id = "quick" ? "quick-group" : "shortcuts"
    }
    static ForWindowSwitch(key, automatic := true) {
        ; Custom combinations already have wildcard behavior; an explicit *
        ; before their prefix key is rejected by AHK. Preserve existing text.
        return !automatic || key = "" || InStr(key, "*") || RegExMatch(key, "\s&\s") ? key : "*" key
    }

    static Settings := Map(
        "Suspend_Hotkeys_Hotkey", "Pause EVE switching hotkeys", "Suspend_All_Hotkeys_Hotkey", "Pause all hotkeys", "HideThumbnailsHotkey", "Hide thumbnails",
        "ToggleLivePreviewsHotkey", "Pause/resume live previews", "ClickThroughHotkey", "Click through thumbnails",
        "SwitchToPreviousWindow_Hotkey", "Previous window", "CycleEveryLoggedIn_Hotkey", "Cycle logged-in windows",
        "Login_Screen_Cycle_Hotkey", "Cycle login screens", "Close_Active_EVE_Win_Hotkey", "Close active EVE window",
        "Close_All_EVE_Win_Hotkey", "Close all EVE windows", "Reload_Program_Hotkey", "Reload program",
        "Exit_Program_Hotkey", "Exit program")

    static TrayLabel(name) => name = "Suspend Hotkeys" ? "Pause EVE switching hotkeys" : name = "Suspend All Hotkeys" ? "Pause all hotkeys" : name

    static Collect(app) {
        entries := []
        for name, profile in app.Profiles
            this.Add(entries, "profile:" name, "Profile: " name, profile.Get("ActivationHotkey", ""))
        automatic := app.AutoWildcardWindowHotkeys
        settings := app.Profiles[app.ProfileHotkeysSettings].Get("Hotkeys Settings", Map())
        for key, label in this.Settings {
            switchesWindow := key = "SwitchToPreviousWindow_Hotkey" || key = "CycleEveryLoggedIn_Hotkey" || key = "Login_Screen_Cycle_Hotkey"
            this.Add(entries, "setting:" key, label, settings.Get(key, ""), switchesWindow && automatic)
        }
        quickProfile := app.FieldProfile("QuickGroupHotkey", app.ProfileThumbnailsInteractions)
        quick := app.Profiles[quickProfile].Get("Hotkeys Settings", Map())
        this.Add(entries, "quick", "Quick Group", quick.Get("QuickGroupHotkey", ""), automatic)
        for name, group in app.Hotkey_Groups {
            for key in ["ForwardsHotkey", "BackwardsHotkey", "FirstCharHotkey"]
                this.Add(entries, "group:" name ":" key, "Group " name ": " key, group.Get(key, ""), automatic)
        }
        for row, item in app._Hotkeys {
            for name, key in item
                this.Add(entries, "character:" row, "Character: " app.CleanTitle(name), key, automatic)
        }
        for name, group in app.NonEVEGroups {
            for key in ["fkey", "bkey"]
                this.Add(entries, "nonGroup:" name ":" key, "App group " name ": " key, group.Get(key, ""), automatic)
        }
        for row, key in app.NonEVEHotkeys["hotkey"]
            this.Add(entries, "nonApp:" row, "App hotkey row " row, key, automatic)
        return entries
    }

    static Add(entries, id, label, key, switchesWindow := false) {
        key := Trim(key, " `t`r`n")
        if key != ""
            entries.Push({id: id, label: label, key: key, parsed: this.Parse(switchesWindow ? this.ForWindowSwitch(key) : key)})
    }

    static Parse(key) {
        key := Trim(key, " `t`r`n")
        up := RegExMatch(key, "i)\s+up$")
        if up
            key := RegExReplace(key, "i)\s+up$")
        mods := Map("^", 0, "!", 0, "+", 0, "#", 0)
        wildcard := false
        while StrLen(key) > 1 && RegExMatch(key, "^(?:([<>]?)([\^!+#])|([~*$]))", &m) {
            if m[3] != ""
                wildcard := wildcard || m[3] = "*"
            else
                mods[m[2]] := m[1] = "<" ? 1 : m[1] = ">" ? 2 : 3
            key := SubStr(key, StrLen(m[0]) + 1)
        }
        prefix := "", custom := false
        if RegExMatch(key, "^(.+?)\s+&\s+(.+)$", &m) {
            custom := true
            ; AHK custom combinations allow additional modifiers by default.
            wildcard := true
            prefix := StrLower(Trim(m[1]))
            key := Trim(m[2])
            modifier := Map("ctrl", "^", "control", "^", "alt", "!", "shift", "+", "lwin", "#", "rwin", "#",
                "lctrl", "^", "lcontrol", "^", "rctrl", "^", "rcontrol", "^", "lalt", "!", "ralt", "!", "lshift", "+", "rshift", "+")
            if modifier.Has(prefix) {
                mods[modifier[prefix]] := SubStr(prefix, 1, 1) = "l" ? 1 : SubStr(prefix, 1, 1) = "r" ? 2 : 3
                prefix := ""
            } else
                prefix := this.Key(prefix)
        }
        return {key: this.Key(key), mods: mods, prefix: prefix, wildcard: wildcard, up: !!up, custom: custom}
    }

    static Key(name) {
        vk := 0, sc := 0
        try vk := GetKeyVK(name)
        try sc := GetKeySC(name)
        return {vk: vk, sc: sc, name: StrLower(name)}
    }

    static SameKey(a, b) {
        for generic, sides in Map(0x10, [0xA0, 0xA1], 0x11, [0xA2, 0xA3], 0x12, [0xA4, 0xA5]) {
            if (a.vk = generic && (b.vk = sides[1] || b.vk = sides[2]))
                || (b.vk = generic && (a.vk = sides[1] || a.vk = sides[2]))
                return true
        }
        if a.vk && b.vk
            return a.vk = b.vk && (!a.sc || !b.sc || a.sc = b.sc)
        if a.sc && b.sc
            return a.sc = b.sc
        return a.name = b.name
    }

    static Overlap(a, b) {
        if a.up != b.up || !this.SameKey(a.key, b.key)
            return false
        if IsObject(a.prefix) != IsObject(b.prefix)
            return false
        if IsObject(a.prefix) && !this.SameKey(a.prefix, b.prefix)
            return false
        for mod, side in a.mods {
            other := b.mods[mod]
            if side && other {
                if !(side & other)
                    return false
            } else if (side && !b.wildcard) || (other && !a.wildcard)
                return false
        }
        return true
    }

    static IntentionalOverlap(first, second) {
        ; First-available group shortcuts yield to available next/previous groups.
        if InStr(first.id, "group:") = 1 && InStr(second.id, "group:") = 1
            && (RegExMatch(first.id, ":FirstCharHotkey$") && RegExMatch(second.id, ":(?:ForwardsHotkey|BackwardsHotkey)$")
                || RegExMatch(second.id, ":FirstCharHotkey$") && RegExMatch(first.id, ":(?:ForwardsHotkey|BackwardsHotkey)$"))
            return true
        if (first.id = "quick" && (InStr(second.id, "group:") = 1 || InStr(second.id, "character:") = 1))
            || (second.id = "quick" && (InStr(first.id, "group:") = 1 || InStr(first.id, "character:") = 1))
            return true
        direct := InStr(first.id, "character:") = 1 ? first : InStr(second.id, "character:") = 1 ? second : 0
        group := InStr(first.id, "group:") = 1 || first.id = "quick" ? first
            : InStr(second.id, "group:") = 1 || second.id = "quick" ? second : 0
        return IsObject(direct) && IsObject(group)
    }

    static ModifierFallback(a, b) {
        ; A more-specific modifier chord can coexist with a wildcard fallback.
        ; Keep duplicate chords and ambiguous/custom combinations visible.
        if a.custom || b.custom
            return false
        moreA := false, moreB := false
        for mod, side in a.mods {
            other := b.mods[mod]
            if side && other && side != other
                return false
            moreA := moreA || (side && !other)
            moreB := moreB || (other && !side)
        }
        return moreA != moreB
    }

    static Conflicts(entries) {
        result := Map()
        for i, entry in entries {
            loop entries.Length - i {
                other := entries[i + A_Index]
                if !this.IsConflict(entry, other)
                    continue
                if !result.Has(entry.id)
                    result[entry.id] := []
                if !result.Has(other.id)
                    result[other.id] := []
                result[entry.id].Push(other.label " (" other.key ")")
                result[other.id].Push(entry.label " (" entry.key ")")
            }
        }
        return result
    }

    static IsConflict(first, second) => !this.IntentionalOverlap(first, second)
        && this.Overlap(first.parsed, second.parsed) && !this.ModifierFallback(first.parsed, second.parsed)

    static ConflictGroups(entries) {
        ; Connected conflicting chords form one group, including key aliases.
        parents := [], involved := Map()
        for index, entry in entries
            parents.Push(index)
        for index, entry in entries {
            loop entries.Length - index {
                other := index + A_Index
                if !this.IsConflict(entry, entries[other])
                    continue
                firstRoot := index, secondRoot := other
                while parents[firstRoot] != firstRoot
                    firstRoot := parents[firstRoot]
                while parents[secondRoot] != secondRoot
                    secondRoot := parents[secondRoot]
                parents[secondRoot] := firstRoot
                involved[index] := true, involved[other] := true
            }
        }
        groups := Map()
        for index, entry in entries {
            if !involved.Has(index)
                continue
            root := index
            while parents[root] != root
                root := parents[root]
            if !groups.Has(root)
                groups[root] := {keys: [], entries: [], seen: Map()}
            group := groups[root], key := SettingsValue.HotkeyText(entry.key)
            if !group.seen.Has(key) {
                group.seen[key] := true
                group.keys.Push(key)
            }
            group.entries.Push(entry)
        }
        result := []
        for root, group in groups
            result.Push(group)
        return result
    }

    static ReplaceLine(text, row, key) {
        lines := StrSplit(StrReplace(text, "`r"), "`n")
        while lines.Length < row
            lines.Push("")
        lines[row] := key
        result := ""
        for i, line in lines
            result .= (i = 1 ? "" : "`n") line
        return result
    }
}
