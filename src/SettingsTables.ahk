; Adapt row editors to the existing JSON schema. Related values are always
; inserted, removed and reordered as one row, never as independent text lists.
class SettingsTables {
    static GestureOptions := [
        {label: "Disabled", value: 0}, {label: "Left click", value: 1}, {label: "Right click / drag", value: 2},
        {label: "Left + right drag", value: 3}, {label: "Shift + left click", value: 5},
        {label: "Shift + right click / drag", value: 6}, {label: "Shift + left + right drag", value: 7},
        {label: "Ctrl + left click", value: 9}, {label: "Ctrl + right click / drag", value: 10},
        {label: "Ctrl + left + right drag", value: 11}, {label: "Ctrl + Shift + left click", value: 13},
        {label: "Ctrl + Shift + right click / drag", value: 14}, {label: "Ctrl + Shift + left + right drag", value: 15}]

    static Group(kind) {
        switch kind {
            case "character-shortcuts": return "character-shortcuts"
            case "character-colors": return "character-appearance"
            case "custom-names": return "labels"
            case "exceptions": return "close-protection"
            case "hidden-characters": return "hidden-characters"
            case "monitored-characters": return "monitored-characters"
            case "group-members": return "character-groups"
            case "app-group-members": return "app-groups"
            case "app-shortcuts": return "app-shortcuts"
            case "gestures": return "mouse-actions"
        }
        throw ValueError("Unknown row editor: " kind)
    }

    static Columns(kind) {
        switch kind {
            case "character-shortcuts": return [this.Column("Character", "character"), this.Column("Shortcut", "hotkey")]
            case "character-colors": return [this.Column("Character", "character"), this.Column("Active border", "color"),
                this.Column("Inactive border", "color"), this.Column("Label color", "color")]
            case "custom-names": return [this.Column("Character", "character"), this.Column("Displayed name", "text")]
            case "exceptions": return [this.Column("Character", "character"), this.Column("Never auto-minimize", "bool"), this.Column("Protect from Close all", "bool")]
            case "hidden-characters", "monitored-characters", "group-members": return [this.Column("Character", "character")]
            case "app-group-members": return [this.Column("Executable (.exe)", "exe"), this.Column("Window title (optional)", "text")]
            case "app-shortcuts": return [this.Column("Executable (.exe)", "exe"), this.Column("Window title (optional)", "text"), this.Column("Shortcut", "hotkey")]
            case "gestures": return [this.Column("Action", "readonly"), this.Column("Gesture", "gesture")]
        }
        throw ValueError("Unknown row editor: " kind)
    }

    static Column(label, kind) => {label: label, kind: kind}

    static Read(profile, kind, groupName := "") {
        rows := []
        switch kind {
            case "character-shortcuts":
                for item in profile["Hotkeys Settings"]["CharacterHotkeys"]
                    for name, key in item
                        rows.Push([this.CleanName(name), key])
            case "character-colors":
                colors := profile["Custom Colors"]["cColors"]
                arrays := [colors["CharNames"], colors["Bordercolor"], colors.Get("IABordercolor", []), colors["TextColor"]]
                maximum := 0
                for values in arrays
                    maximum := Max(maximum, values.Length)
                loop maximum {
                    index := A_Index
                    row := []
                    for values in arrays
                        row.Push("")
                    for column, values in arrays
                        row[column] := index <= values.Length ? values[index] : ""
                    row[1] := this.CleanName(row[1])
                    rows.Push(row)
                }
            case "custom-names":
                names := profile["Thumbnails Visuals"].Get("CustomThumbnailNames", Map("Characters", "", "Names", ""))
                characters := StrSplit(StrReplace(names["Characters"], "`r"), "`n")
                labels := StrSplit(StrReplace(names["Names"], "`r"), "`n")
                loop Max(characters.Length, labels.Length) {
                    name := A_Index <= characters.Length ? characters[A_Index] : ""
                    label := A_Index <= labels.Length ? labels[A_Index] : ""
                    if name != "" || label != ""
                        rows.Push([this.CleanName(name), label])
                }
            case "exceptions":
                combined := Map()
                combined.CaseSense := "Off"
                for index, key in ["Dont_Minimize_Clients", "DontCloseClients"] {
                    column := index + 1
                    for name in profile["Client Settings"][key] {
                        name := this.CleanName(name)
                        if !combined.Has(name)
                            combined[name] := [name, 0, 0]
                        combined[name][column] := 1
                    }
                }
                for name, row in combined
                    rows.Push(row)
            case "hidden-characters":
                for name, hidden in profile["Thumbnail Visibility"]
                    if hidden
                        rows.Push([this.CleanName(name)])
            case "monitored-characters":
                for name in profile["Game Logs Monitoring"]["charsToMonitor"]
                    rows.Push([this.CleanName(name)])
            case "group-members":
                groups := profile["Hotkey Groups"]
                if groups.Has(groupName) && groups[groupName] is Map
                    for name in groups[groupName]["Characters"]
                        rows.Push([this.CleanName(name)])
            case "app-group-members":
                groups := profile["Non-EVE Applications"]["NonEVEGroups"]
                if groups.Has(groupName) {
                    group := groups[groupName]
                    loop Max(group["exe"].Length, group["title"].Length)
                        rows.Push([A_Index <= group["exe"].Length ? group["exe"][A_Index] : "",
                            A_Index <= group["title"].Length ? group["title"][A_Index] : ""])
                }
            case "app-shortcuts":
                keys := profile["Non-EVE Applications"]["NonEVEHotkeys"]
                loop Max(keys["exe"].Length, keys["title"].Length, keys["hotkey"].Length)
                    rows.Push([A_Index <= keys["exe"].Length ? keys["exe"][A_Index] : "",
                        A_Index <= keys["title"].Length ? keys["title"][A_Index] : "",
                        A_Index <= keys["hotkey"].Length ? keys["hotkey"][A_Index] : ""])
            case "gestures":
                for action in ["ActivateThumbnail", "MoveThumbnail", "MoveAllThumbnails", "ResizeThumbnail", "ResizeAllThumbnails",
                    "HideThumbnail", "MinimizeClient", "CloseClient", "DisableFromGroups", "QuickGroup"] {
                    flags := profile["Thumbnails Interactions"][action]
                    rows.Push([action, flags["lmb"] + flags["rmb"] * 2 + flags["shift"] * 4 + flags["ctrl"] * 8])
                }
        }
        return rows
    }

    static Validate(kind, rows) {
        columns := this.Columns(kind)
        used := Map()
        used.CaseSense := "Off"
        gestures := Map()
        for index, row in rows {
            if row.Length != columns.Length
                throw ValueError("Row " index " has missing columns.")
            for column, spec in columns {
                value := Trim(row[column], " `t`r`n")
                try {
                    switch spec.kind {
                        case "character":
                            if this.CleanName(value) = "" || StrLen(value) > 160 || RegExMatch(value, "[\r\n\t]")
                                throw ValueError("Enter a character name without line breaks.")
                        case "exe":
                            if !RegExMatch(value, 'i)^[^<>:"/\\|?*\r\n]+\.exe$')
                                throw ValueError("Enter an executable name such as notepad.exe, without a folder path.")
                        case "color": SettingsValue.Color(value)
                        case "hotkey": SettingsValue.Hotkey(value)
                        case "gesture":
                            valid := false
                            for option in this.GestureOptions
                                if option.value = value
                                    valid := true
                            if !valid
                                throw ValueError("Choose a gesture containing a mouse button, or Disabled.")
                    }
                } catch as err
                    throw ValueError("Row " index ", " spec.label ": " err.Message)
            }
            if kind = "gestures" {
                mask := row[2] + 0
                if mask && gestures.Has(mask)
                    throw ValueError("The same gesture is assigned to " this.ActionName(gestures[mask]) " and " this.ActionName(row[1]) ".")
                gestures[mask] := row[1]
            } else {
                key := this.CleanName(row[1])
                if kind = "app-shortcuts" || kind = "app-group-members"
                    key := row[1] "`n" row[2]
                if used.Has(key)
                    throw ValueError("Row " index " duplicates row " used[key] ". Each assignment must have a unique target.")
                used[key] := index
            }
        }
    }

    static Write(profile, kind, rows, groupName := "") {
        switch kind {
            case "character-shortcuts":
                entries := []
                for row in rows
                    entries.Push(Map(this.FullName(row[1]), Trim(row[2])))
                profile["Hotkeys Settings"]["CharacterHotkeys"] := entries
            case "character-colors":
                colors := Map("CharNames", [], "Bordercolor", [], "IABordercolor", [], "TextColor", [])
                for row in rows {
                    colors["CharNames"].Push(this.FullName(row[1]))
                    colors["Bordercolor"].Push(SettingsValue.Color(row[2]))
                    colors["IABordercolor"].Push(SettingsValue.Color(row[3]))
                    colors["TextColor"].Push(SettingsValue.Color(row[4]))
                }
                profile["Custom Colors"]["cColors"] := colors
            case "custom-names":
                names := "", labels := ""
                for index, row in rows {
                    names .= (index = 1 ? "" : "`n") this.CleanName(row[1])
                    labels .= (index = 1 ? "" : "`n") Trim(row[2])
                }
                profile["Thumbnails Visuals"]["CustomThumbnailNames"] := Map("Characters", names, "Names", labels)
            case "exceptions":
                minimize := [], close := []
                for row in rows {
                    if row[2]
                        minimize.Push(this.FullName(row[1]))
                    if row[3]
                        close.Push(this.FullName(row[1]))
                }
                profile["Client Settings"]["Dont_Minimize_Clients"] := minimize
                profile["Client Settings"]["DontCloseClients"] := close
            case "hidden-characters":
                hidden := Map()
                for row in rows
                    hidden[this.CleanName(row[1])] := 1
                profile["Thumbnail Visibility"] := hidden
            case "monitored-characters", "group-members":
                names := []
                for row in rows
                    names.Push(this.FullName(row[1]))
                if kind = "group-members"
                    profile["Hotkey Groups"][groupName]["Characters"] := names
                else
                    profile["Game Logs Monitoring"]["charsToMonitor"] := names
            case "app-group-members", "app-shortcuts":
                apps := [], titles := [], hotkeys := []
                for row in rows {
                    apps.Push(Trim(row[1]))
                    titles.Push(Trim(row[2]))
                    if kind = "app-shortcuts"
                        hotkeys.Push(Trim(row[3]))
                }
                if kind = "app-group-members" {
                    profile["Non-EVE Applications"]["NonEVEGroups"][groupName]["exe"] := apps
                    profile["Non-EVE Applications"]["NonEVEGroups"][groupName]["title"] := titles
                } else
                    profile["Non-EVE Applications"]["NonEVEHotkeys"] := Map("exe", apps, "title", titles, "hotkey", hotkeys)
            case "gestures":
                for row in rows {
                    mask := row[2] + 0
                    profile["Thumbnails Interactions"][row[1]] := Map("lmb", !!(mask & 1), "rmb", !!(mask & 2), "shift", !!(mask & 4), "ctrl", !!(mask & 8))
                }
        }
    }

    static CleanName(name) {
        name := Trim(name, " `t`r`n")
        return SubStr(name, 1, 6) = "EVE - " ? SubStr(name, 7) : name = "EVE" ? "" : name
    }

    static FullName(name) => "EVE - " this.CleanName(name)

    static ActionName(action) {
        return Map("ActivateThumbnail", "Activate client", "MoveThumbnail", "Move thumbnail", "MoveAllThumbnails", "Move all thumbnails",
            "ResizeThumbnail", "Resize thumbnail", "ResizeAllThumbnails", "Resize all thumbnails", "HideThumbnail", "Hide thumbnail",
            "MinimizeClient", "Minimize client", "CloseClient", "Close client", "DisableFromGroups", "Exclude from groups", "QuickGroup", "Toggle Quick Group membership")[action]
    }

    static Display(kind, row) {
        result := row.Clone()
        columns := this.Columns(kind)
        for index, column in columns {
            if column.kind = "bool"
                result[index] := row[index] ? "Yes" : "No"
            else if column.kind = "hotkey"
                result[index] := SettingsValue.HotkeyText(row[index])
            else if column.kind = "gesture" {
                result[index] := "Invalid gesture"
                for option in this.GestureOptions
                    if option.value = row[index]
                        result[index] := option.label
            } else if column.kind = "readonly" && kind = "gestures"
                result[index] := this.ActionName(row[index])
        }
        return result
    }
}
