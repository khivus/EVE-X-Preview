; The editor owns a draft. Runtime state is touched only after a successful Apply.
class SettingsModel {
    static LegacyGroups := ["Thumbnails Behavior", "Thumbnails Visuals", "Thumbnails Interactions", "Thumbnail Visibility",
        "Custom Colors", "Hotkeys Settings", "Hotkey Groups", "Character shortcuts", "Character groups", "Client Settings", "Game Logs Monitoring",
        "Monitored Events", "DPS Monitoring", "Focused-client sounds", "Focused-client DPS", "Non-EVE Applications", "Tray Menu Settings", "Other"]

    static Groups := []
    static __New() {
        for section in SettingsSchema.Sections
            if section.group != ""
                this.Groups.Push(section.id)
    }
    static Canonical(group) {
        aliases := Map("Thumbnails Behavior", "layout", "Thumbnails Visuals", "appearance", "Thumbnails Interactions", "mouse-actions",
            "Thumbnail Visibility", "hidden-characters", "Custom Colors", "character-appearance", "Hotkeys Settings", "shortcuts",
            "Hotkey Groups", "switching-behavior", "Character shortcuts", "character-shortcuts", "Character groups", "character-groups",
            "Client Settings", "window-behavior", "Game Logs Monitoring", "sources", "Monitored Events", "events", "DPS Monitoring", "dps",
            "Focused-client sounds", "alert-behavior", "Focused-client DPS", "dps", "Non-EVE Applications", "app-groups", "Tray Menu Settings", "tray", "Other", "performance")
        return aliases.Get(group, group)
    }
    static Units() {
        units := Map()
        for section in SettingsSchema.Sections
            if section.group != ""
                units[section.id] := []
        for id, field in SettingsSchema.Fields
            if !field.global && field.group != ""
                units[field.group].Push({path: field.path, legacy: field.legacyGroup})
        extras := [
            ["labels", "Thumbnails Visuals", ["Thumbnails Visuals", "CustomThumbnailNames"]],
            ["character-appearance", "Custom Colors", ["Custom Colors", "cColors"]],
            ["character-shortcuts", "Character shortcuts", ["Hotkeys Settings", "CharacterHotkeys"]],
            ["character-groups", "Character groups", ["Hotkey Groups"]],
            ["mouse-actions", "Thumbnails Interactions", ["Thumbnails Interactions"]],
            ["close-protection", "Client Settings", ["Client Settings", "Dont_Minimize_Clients"]],
            ["close-protection", "Client Settings", ["Client Settings", "DontCloseClients"]],
            ["hidden-characters", "Thumbnail Visibility", ["Thumbnail Visibility"]],
            ["monitored-characters", "Game Logs Monitoring", ["Game Logs Monitoring", "charsToMonitor"]],
            ["events", "Monitored Events", ["Game Logs Monitoring", "monitoredEvents"]],
            ["app-groups", "Non-EVE Applications", ["Non-EVE Applications", "NonEVEGroups"]],
            ["app-shortcuts", "Non-EVE Applications", ["Non-EVE Applications", "NonEVEHotkeys"]],
            ["tray", "Tray Menu Settings", ["Tray Menu Settings", "TrayMenuShortcuts"]]]
        for extra in extras
            units[extra[1]].Push({legacy: extra[2], path: extra[3]})
        return units
    }

    __New(config) {
        this.original := this.Clone(config)
        this.config := this.Clone(config)
        this.defaults := JSON.Load(default_JSON)
        this.profile := this.config["LastUsedProfile"]
        this.profileOrigins := Map()
        for name in this.config["_Profiles"]
            this.profileOrigins[name] := name
        this.undo := []
        this.MigrateInheritance()
        this.saved := JSON.Dump(this.config)
    }

    static Clone(value) => IsObject(value) ? JSON.Load(JSON.Dump(value)) : value
    static Equal(first, second) => JSON.Dump([first]) = JSON.Dump([second])
    Clone(value) => SettingsModel.Clone(value)
    Dirty => JSON.Dump(this.config) != this.saved
    Profiles => this.config["_Profiles"]

    LegacySource(group, name) {
        if name = "Default"
            return name
        flags := this.Profiles[name]["SettingsInheritance"]
        return flags.Get(group, 0) ? "Default" : name
    }
    MigrateInheritance() {
        shared := this.Profiles["Default"]["Other"]["Global_Groups"]
        for name, profile in this.Profiles {
            if !profile.Has("SettingsInheritance") {
                flags := Map()
                localFlags := profile["Other"]["Global_Groups"]
                for group in SettingsModel.LegacyGroups
                    flags[group] := name != "Default" && shared.Get(group, 0) && !localFlags.Get(group, 0) ? 1 : 0
                profile["SettingsInheritance"] := flags
            }
            logs := profile["Game Logs Monitoring"]
            effectiveLogs := this.Profiles[this.LegacySource("Game Logs Monitoring", name)]["Game Logs Monitoring"]
            effectiveThumbnails := this.Profiles[this.LegacySource("Thumbnails Behavior", name)]["Thumbnails Behavior"]
            if !logs.Has("suppressFocusedEventSounds")
                logs["suppressFocusedEventSounds"] := !!(effectiveLogs["supressForFocused"] || effectiveThumbnails["HideThumbForActiveWin"])
            if !profile["DPS Monitoring"].Has("hideDPSForFocused")
                profile["DPS Monitoring"]["hideDPSForFocused"] := !!effectiveLogs["supressForFocused"]
        }
        snapshots := this.Clone(this.Profiles)
        for name, profile in this.Profiles {
            if profile.Get("SettingsSectionsVersion", 0) >= 1
                continue
            backup := profile.Get("SettingsMigrationBackup", Map())
            for section, items in SettingsModel.Units() {
                allDefault := name != "Default"
                for item in items
                    if this.LegacySource(item.legacy, name) != "Default"
                        allDefault := false
                profile["SettingsInheritance"][section] := allDefault ? 1 : 0
                ; Mixed legacy sources become a customized subsection with the
                ; same effective values. Preserve dormant overrides in a backup.
                if !allDefault && name != "Default"
                    for item in items
                        if this.LegacySource(item.legacy, name) = "Default" {
                            original := SettingsModel.At(snapshots[name], item.path, "")
                            key := JSON.Dump(item.path)
                            if !backup.Has(key)
                                backup[key] := this.Clone(original)
                            value := SettingsModel.At(snapshots["Default"], item.path,
                                SettingsModel.At(this.defaults["_Profiles"]["Default"], item.path, ""))
                            SettingsModel.Put(profile, item.path, this.Clone(value))
                        }
            }
            if backup.Count
                profile["SettingsMigrationBackup"] := backup
            profile["SettingsSectionsVersion"] := 1
        }
    }

    Source(group, profile?) {
        if !IsSet(profile)
            profile := this.profile
        if group = "" || profile = "Default"
            return profile
        group := SettingsModel.Canonical(group)
        return this.Profiles[profile]["SettingsInheritance"].Get(group, 0) ? "Default" : profile
    }

    Inherited(group) => group != "" && this.Source(group) != this.profile

    SetInheritance(group, inherit) {
        group := SettingsModel.Canonical(group)
        if this.profile = "Default"
            throw ValueError("The Default profile cannot inherit from itself.")
        if !this.Profiles[this.profile]["SettingsInheritance"].Has(group)
            throw ValueError("Unknown settings section.")
        this.Profiles[this.profile]["SettingsInheritance"][group] := !!inherit
    }

    Get(id) => this.GetFor(id, this.profile)

    GetFor(id, profile) {
        field := SettingsSchema.Fields[id]
        root := field.global ? this.config : this.Profiles[this.Source(field.group, profile)]
        fallback := field.global ? this.defaults : this.defaults["_Profiles"]["Default"]
        return this.At(root, field.path, this.At(fallback, field.path, ""))
    }

    Relations() {
        issues := []
        for profile in this.Profiles {
            for dimension in ["width", "height"] {
                initial := this.GetFor("ThumbnailStartLocation" dimension, profile)
                minimum := this.GetFor("ThumbnailMinimumSize" dimension, profile)
                if IsNumber(initial) && IsNumber(minimum) && minimum > initial
                    issues.Push({profile: profile, id: "ThumbnailStartLocation" dimension,
                        relatedId: "ThumbnailMinimumSize" dimension,
                        message: "Default " dimension " must be at least the minimum " dimension " (" minimum " px)."})
            }
        }
        return issues
    }

    Set(id, input) {
        field := SettingsSchema.Fields[id]
        if !field.global && this.Inherited(field.group)
            throw ValueError("Use Customize for this profile before editing inherited settings.")
        value := SettingsValue.Parse(field, input)
        root := field.global ? this.config : this.Profiles[this.profile]
        this.Put(root, field.path, value)
        return value
    }

    Group(group, storage := "", own := false) {
        if storage = ""
            storage := group
        return this.Profiles[own ? this.profile : this.Source(group)][storage]
    }

    static At(root, path, fallback := "") {
        node := root
        for key in path {
            if !(node is Map) || !node.Has(key)
                return fallback
            node := node[key]
        }
        return node
    }

    At(root, path, fallback := "") => SettingsModel.At(root, path, fallback)

    static Put(root, path, value) {
        node := root
        loop path.Length - 1 {
            key := path[A_Index]
            if !node.Has(key) || !(node[key] is Map)
                node[key] := Map()
            node := node[key]
        }
        node[path[-1]] := value
    }

    Put(root, path, value) => SettingsModel.Put(root, path, value)

    SelectProfile(name) {
        if !this.Profiles.Has(name)
            throw ValueError("This profile no longer exists.")
        this.profile := name
    }

    CheckName(name, collection, previous := "") {
        name := Trim(name, " `t`r`n")
        if name = "" || StrLen(name) > 80 || RegExMatch(name, "[\r\n\t]")
            throw ValueError("Enter a name from 1 to 80 characters, without tabs or line breaks.")
        for existing in collection
            if StrLower(existing) = StrLower(name) && existing != previous
                throw ValueError("A name matching '" name "' already exists.")
        return name
    }

    CreationInheritance() {
        stored := this.config.Get("NewProfileInheritance", Map()), flags := Map()
        for group in SettingsModel.Groups
            flags[group] := !!stored.Get(group, 0)
        return flags
    }
    SetCreationInheritance(flags) {
        normalized := Map()
        for group in SettingsModel.Groups
            normalized[group] := !!flags.Get(group, 0)
        this.config["NewProfileInheritance"] := normalized
    }
    AddProfile(name, copyCurrent := true, inheritance?) {
        name := this.CheckName(name, this.Profiles)
        if this.profileOrigins.Has(name)
            this.profileOrigins.Delete(name)
        source := copyCurrent ? this.Profiles[this.profile] : this.defaults["_Profiles"]["Default"]
        created := this.Clone(source)
        if IsSet(inheritance) && copyCurrent {
            for group, items in SettingsModel.Units()
                for item in items
                    this.Put(created, item.path, this.Clone(this.At(this.Profiles[this.Source(group)], item.path,
                        this.At(this.defaults["_Profiles"]["Default"], item.path, ""))))
            ; Some legacy table containers also hold scalar settings belonging
            ; to other subsections. Restore their effective sources last.
            for id, field in SettingsSchema.Fields
                if !field.global
                    this.Put(created, field.path, this.Clone(this.Get(id)))
        }
        this.Profiles[name] := created
        this.Profiles[name]["ActivationHotkey"] := ""
        if !copyCurrent || this.profile = "Default" {
            flags := Map()
            for group in SettingsModel.Groups
                flags[group] := 0
            this.Profiles[name]["SettingsInheritance"] := flags
            this.Profiles[name]["SettingsSectionsVersion"] := 1
        }
        if IsSet(inheritance) {
            flags := Map()
            for group in SettingsModel.Groups
                flags[group] := !!inheritance.Get(group, 0)
            this.Profiles[name]["SettingsInheritance"] := flags
            this.Profiles[name]["SettingsSectionsVersion"] := 1
        }
        this.MigrateInheritance()
        this.profile := name
        return name
    }

    RenameProfile(name) {
        old := this.profile
        if old = "Default"
            throw ValueError("The Default profile has a fixed name.")
        name := this.CheckName(name, this.Profiles, old)
        if name = old
            return old
        this.Profiles[name] := this.Profiles.Delete(old)
        if this.profileOrigins.Has(old)
            this.profileOrigins[name] := this.profileOrigins.Delete(old)
        if this.config["LastUsedProfile"] = old
            this.config["LastUsedProfile"] := name
        this.profile := name
        return name
    }

    DeleteProfile() {
        name := this.profile
        if name = "Default"
            throw ValueError("The Default profile cannot be removed.")
        this.undo.Push({kind: "profile", name: name, value: this.Clone(this.Profiles[name]), active: this.config["LastUsedProfile"]})
        this.Profiles.Delete(name)
        if this.config["LastUsedProfile"] = name
            this.config["LastUsedProfile"] := "Default"
        this.profile := "Default"
    }

    ActivateProfile() => this.config["LastUsedProfile"] := this.profile

    SaveRows(kind, rows, groupName := "") {
        group := SettingsTables.Group(kind)
        if this.Inherited(group)
            throw ValueError("Customize this section before editing its rows.")
        SettingsTables.Validate(kind, rows)
        profile := this.Profiles[this.profile]
        SettingsTables.Write(profile, kind, rows, groupName)
    }

    Rows(kind, groupName := "") => SettingsTables.Read(this.Profiles[this.Source(SettingsTables.Group(kind))], kind, groupName)

    ResetThumbnailSizes(runtime := 0) {
        width := this.Get("ThumbnailStartLocationwidth"), height := this.Get("ThumbnailStartLocationheight")
        positions := this.Profiles[this.profile]["Thumbnail Positions"]
        for name, position in positions {
            if position is Map
                position["width"] := width, position["height"] := height
            else if position is Array && position.Length >= 4
                positions[name] := Map("x", position[1], "y", position[2], "width", width, "height", height)
        }
        if IsObject(runtime) && runtime.HasOwnProp("ThumbWindows")
            for hwnd, thumb in runtime.ThumbWindows.OwnProps() {
                try {
                    window := thumb["Window"]
                    WinGetPos(&x, &y,,, window.Hwnd)
                    if window.Title != ""
                        positions[window.Title] := Map("x", x, "y", y, "width", width, "height", height)
                }
            }
    }

    Groups(kind) {
        profile := this.Profiles[this.Source(kind = "characters" ? "Character groups" : "Non-EVE Applications")]
        return kind = "characters" ? profile["Hotkey Groups"] : profile["Non-EVE Applications"]["NonEVEGroups"]
    }

    AddGroup(kind, name) {
        group := kind = "characters" ? "Character groups" : "Non-EVE Applications"
        if this.Inherited(group)
            throw ValueError("Customize this section before adding a group.")
        groups := this.Groups(kind)
        name := this.CheckName(name, groups)
        groups[name] := kind = "characters"
            ? Map("Characters", [], "ForwardsHotkey", "", "BackwardsHotkey", "", "FirstCharHotkey", "")
            : Map("exe", [], "title", [], "fkey", "", "bkey", "")
        return name
    }

    RenameGroup(kind, old, name) {
        group := kind = "characters" ? "Character groups" : "Non-EVE Applications"
        if this.Inherited(group)
            throw ValueError("Customize this section before renaming a group.")
        groups := this.Groups(kind)
        if !groups.Has(old)
            throw ValueError("Select a group first.")
        name := this.CheckName(name, groups, old)
        if old != name
            groups[name] := groups.Delete(old)
        return name
    }

    DeleteGroup(kind, name) {
        group := kind = "characters" ? "Character groups" : "Non-EVE Applications"
        if this.Inherited(group)
            throw ValueError("Customize this section before removing a group.")
        groups := this.Groups(kind)
        if !groups.Has(name)
            throw ValueError("Select a group first.")
        this.undo.Push({kind: "group", profile: this.profile, groupKind: kind, name: name, value: this.Clone(groups[name])})
        groups.Delete(name)
    }

    SetGroupShortcut(kind, name, key, value) {
        if kind != "characters" && kind != "apps"
            throw ValueError("Choose character groups or application groups.")
        if this.Inherited(kind = "characters" ? "character-groups" : "app-groups")
            throw ValueError("Customize this section before editing its shortcuts.")
        groups := this.Groups(kind)
        allowed := kind = "characters" ? ["ForwardsHotkey", "BackwardsHotkey", "FirstCharHotkey"] : ["fkey", "bkey"]
        valid := false
        for candidate in allowed
            valid := valid || key = candidate
        if !valid || !groups.Has(name)
            throw ValueError("Choose a valid group and shortcut.")
        SettingsValue.Hotkey(value)
        groups[name][key] := value
    }

    UndoLast() {
        if !this.undo.Length
            return false
        entry := this.undo[-1]
        if entry.kind = "profile" {
            this.CheckName(entry.name, this.Profiles)
            this.Profiles[entry.name] := entry.value
            if entry.active = entry.name && this.config["LastUsedProfile"] = "Default"
                this.config["LastUsedProfile"] := entry.active
            this.profile := entry.name
        } else if entry.kind = "group" {
            if !this.Profiles.Has(entry.profile)
                throw ValueError("Restore the removed profile before restoring its group.")
            this.profile := entry.profile
            if this.Inherited(entry.groupKind = "characters" ? "Character groups" : "Non-EVE Applications")
                throw ValueError("Customize this profile's group section before restoring the removed group.")
            groups := this.Groups(entry.groupKind)
            this.CheckName(entry.name, groups)
            groups[entry.name] := entry.value
        } else if entry.kind = "library" {
            restored := this.Clone(entry.paths)
            for path in EventSounds.Library(this.config) {
                found := false
                for previous in restored
                    if EventSounds.SameFile(previous, path)
                        found := true
                if !found
                    restored.Push(path)
            }
            this.config["SoundLibrary"] := restored
        } else if entry.kind = "table-state" {
            if !this.Profiles.Has(entry.profile)
                throw ValueError("Restore the profile before restoring its table rows.")
            this.profile := entry.profile
            this.SaveRows(entry.tableKind, this.Clone(entry.rows), entry.groupName)
        } else if entry.kind = "row-batch" {
            this.profile := entry.profile
            rows := this.Rows(entry.tableKind, entry.groupName)
            for removed in entry.entries
                rows.InsertAt(Min(removed.index, rows.Length + 1), this.Clone(removed.row))
            this.SaveRows(entry.tableKind, rows, entry.groupName)
        } else if entry.kind = "row" {
            this.profile := entry.profile
            rows := this.Rows(entry.tableKind, entry.groupName)
            rows.InsertAt(Min(entry.index, rows.Length + 1), this.Clone(entry.row))
            this.SaveRows(entry.tableKind, rows, entry.groupName)
        } else {
            this.profile := entry.profile
            this.SaveRows(entry.tableKind, entry.rows, entry.groupName)
        }
        this.undo.Pop()
        return true
    }

    ; Apply only locally changed leaves to the latest runtime JSON. Window
    ; positions, discovered character IDs and unrelated tray changes survive.
    Merge(current) {
        result := SettingsModel.MergeChanged(this.original, this.config, current)
        ; A renamed profile is still the same running profile. Carry across
        ; positions and discovered IDs updated under its old name while editing.
        for name, origin in this.profileOrigins {
            if name = origin || !this.Profiles.Has(name) || !current["_Profiles"].Has(origin)
                continue
            result["_Profiles"][name] := SettingsModel.MergeChanged(this.original["_Profiles"][origin],
                this.Profiles[name], current["_Profiles"][origin])
        }
        return result
    }

    ; Check edited leaves against the same baseline used by Merge. Renames keep
    ; their old profile identity so live discovery/geometry is compared correctly.
    Conflicts(current) {
        before := this.Clone(this.original), after := this.Clone(this.config), live := this.Clone(current), conflicts := []
        for name, origin in this.profileOrigins {
            if name = origin || !this.Profiles.Has(name)
                continue
            if live["_Profiles"].Has(name)
                conflicts.Push("_Profiles/" name)
            if !live["_Profiles"].Has(origin)
                conflicts.Push("_Profiles/" origin)
            else if before["_Profiles"].Has(origin)
                SettingsModel.FindConflicts(before["_Profiles"][origin], this.Profiles[name], live["_Profiles"][origin], "_Profiles/" name, conflicts)
            for collection in [before["_Profiles"], after["_Profiles"], live["_Profiles"]] {
                if collection.Has(origin)
                    collection.Delete(origin)
                if collection.Has(name)
                    collection.Delete(name)
            }
        }
        SettingsModel.FindConflicts(before, after, live, "", conflicts)
        return conflicts
    }

    static FindConflicts(before, after, current, path, result) {
        if this.Equal(before, after) || this.Equal(after, current)
            return
        if before is Map && after is Map && current is Map {
            for key, previous in before {
                child := path = "" ? key : path "/" key
                if !after.Has(key) {
                    if current.Has(key) && !this.Equal(previous, current[key])
                        result.Push(child)
                } else if !this.Equal(previous, after[key]) {
                    if !current.Has(key)
                        result.Push(child)
                    else
                        this.FindConflicts(previous, after[key], current[key], child, result)
                }
            }
            for key, value in after
                if !before.Has(key) && current.Has(key) && !this.Equal(value, current[key])
                    result.Push(path = "" ? key : path "/" key)
        } else if !this.Equal(before, current)
            result.Push(path)
    }

    static MergeChanged(before, after, current) {
        if this.Equal(before, after)
            return this.Clone(current)
        if !(before is Map) || !(after is Map)
            return this.Clone(after)
        result := current is Map ? this.Clone(current) : Map()
        for key, value in before
            if !after.Has(key) && result.Has(key)
                result.Delete(key)
        for key, value in after {
            if !before.Has(key)
                result[key] := this.Clone(value)
            else if !this.Equal(before[key], value)
                result[key] := result.Has(key) ? this.MergeChanged(before[key], value, result[key]) : this.Clone(value)
        }
        return result
    }
}

class SettingsValue {
    static Parse(field, input) {
        if field.kind = "bool" {
            if !IsObject(input) && (input = 0 || input = "false")
                return false
            if !IsObject(input) && (input = 1 || input = "true")
                return true
            throw ValueError("Choose enabled or disabled.")
        }
        value := Trim(input, " `t`r`n")
        if field.kind = "integer" || field.kind = "number" || field.kind = "slider" {
            pattern := field.kind = "number" ? "^[+-]?(?:\d+(?:\.\d*)?|\.\d+)$" : "^[+-]?\d+$"
            if !RegExMatch(value, pattern)
                throw ValueError("Enter " (field.kind = "number" ? "a number" : "a whole number") ".")
            value += 0
            if value < field.min || value > field.max
                throw ValueError("Enter a value from " field.min " to " field.max (field.unit != "" ? " " field.unit : "") ".")
            return value
        }
        if field.kind = "color"
            return this.Color(value)
        if field.kind = "hotkey" {
            this.Hotkey(value)
            return value
        }
        if field.kind = "choice" {
            for option in field.choices
                if option.value = input
                    return option.value
            throw ValueError("Select one of the available options.")
        }
        if field.kind = "font" {
            if value = "" || StrLen(value) > 31
                throw ValueError("Choose an installed font name, up to 31 characters.")
            if !SettingsFonts.Available(value)
                throw ValueError("Choose an installed font from the list.")
            return value
        }
        if field.kind = "directory" {
            if value != "" && (!DirExist(value) || InStr(value, '"'))
                throw ValueError("Choose an existing folder, or leave it blank for automatic detection.")
            return value
        }
        if StrLen(value) > 65535
            throw ValueError("This text is too long (maximum 65535 characters).")
        return field.kind = "multiline" ? StrReplace(input, "`r") : value
    }

    static Color(value) {
        value := Trim(value)
        named := Map("black", "000000", "silver", "C0C0C0", "gray", "808080", "grey", "808080", "white", "FFFFFF",
            "maroon", "800000", "red", "FF0000", "purple", "800080", "fuchsia", "FF00FF", "green", "008000",
            "lime", "00FF00", "olive", "808000", "yellow", "FFFF00", "navy", "000080", "blue", "0000FF",
            "teal", "008080", "aqua", "00FFFF")
        if named.Has(StrLower(value))
            return named[StrLower(value)]
        if RegExMatch(value, "i)^rgb\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)$", &m)
            || RegExMatch(value, "^(\d+)\s*,\s*(\d+)\s*,\s*(\d+)$", &m) {
            for index in [1, 2, 3]
                if m[index] > 255
                    throw ValueError("RGB components must be between 0 and 255.")
            return Format("{:02X}{:02X}{:02X}", m[1], m[2], m[3])
        }
        if RegExMatch(value, "i)^(?:#|0x)?([0-9a-f]{1,6})$", &m)
            return Format("{:06X}", Integer("0x" m[1]))
        throw ValueError("Use an RGB hex color such as FF8800, or RGB values from 0 to 255.")
    }

    static Hotkey(value) {
        if value = ""
            return true
        if StrLen(value) > 100 || RegExMatch(value, "[\r\n\t]")
            throw ValueError("Enter one shortcut, without line breaks.")
        parsed := HotkeyBindings.Parse(value)
        if (!parsed.key.vk && !parsed.key.sc) || (IsObject(parsed.prefix) && !parsed.prefix.vk && !parsed.prefix.sc)
            throw ValueError("Unknown key name. Use Record or valid AutoHotkey shortcut syntax.")
        return true
    }

    static HotkeyText(value) {
        if value = ""
            return "Not assigned"
        parsed := HotkeyBindings.Parse(value)
        text := ""
        for pair in [["^", "Ctrl"], ["!", "Alt"], ["+", "Shift"], ["#", "Win"]] {
            symbol := pair[1], label := pair[2]
            side := parsed.mods[symbol]
            if side
                text .= (side = 1 ? "Left " : side = 2 ? "Right " : "") label "+"
        }
        if IsObject(parsed.prefix)
            text .= this.KeyName(parsed.prefix.name) "+"
        text .= this.KeyName(parsed.key.name)
        if parsed.up
            text .= " (release)"
        if parsed.wildcard
            text .= " (extra modifiers allowed)"
        return text
    }

    static KeyName(value) {
        names := Map("lbutton", "Left click", "rbutton", "Right click", "mbutton", "Middle click", "xbutton1", "Mouse 4",
            "xbutton2", "Mouse 5", "wheelup", "Wheel up", "wheeldown", "Wheel down", "escape", "Esc", "space", "Space")
        if names.Has(value)
            return names[value]
        name := ""
        try name := GetKeyName(value)
        return name != "" ? name : value
    }
}
