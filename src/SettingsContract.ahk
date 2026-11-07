; Versioned presentation metadata. Storage, defaults and ordinary field lists
; remain authoritative in SettingsSchema and SettingsTables.
class SettingsContract {
    static BridgeVersion := 1
    static SchemaVersion := 1
    static TableKinds := ["character-shortcuts", "character-colors", "custom-names", "exceptions", "hidden-characters",
        "monitored-characters", "group-members", "app-group-members", "app-shortcuts", "gestures"]

    static Bool(value) => !!value

    ; Keep AHK-facing results as native booleans. Convert presentation flags at
    ; the transport boundary; opaque draft/configuration values stay lossless.
    static JsonFlags(value, context := "") {
        if context = "draft" || context = "original" || context = "config"
            return value
        if value is Array {
            result := []
            for item in value
                result.Push(this.JsonFlags(item, context))
            return result
        }
        if !(value is Map)
            return value
        result := Map()
        for key, item in value {
            if context = "inheritance" || context = "creationInheritance"
                result[key] := item ? JSON.true : JSON.false
            else if RegExMatch(key, "^(ready|dirty|clientDirty|inherited|enabled|canUndo|runtime|canApply|applied|restartRequired|closed|active|global|runtimeOnly|destructive|busy)$") && !IsObject(item)
                result[key] := item ? JSON.true : JSON.false
            else
                result[key] := this.JsonFlags(item, key)
        }
        return result
    }

    static Wire(value) {
        if value is Array {
            result := []
            for item in value
                result.Push(this.Wire(item))
            return result
        }
        if value is Map {
            result := Map()
            for key, item in value
                result[key] := this.Wire(item)
            return result
        }
        if Type(value) = "Object" {
            result := Map()
            for key, item in value.OwnProps()
                result[key] := this.Wire(item)
            return result
        }
        return value
    }

    static Commands() {
        result := []
        for item in [
            ["test-thumbnail", "Toggle test thumbnail", "appearance", false, false],
            ["stop-test-thumbnail", "Close test thumbnail", "appearance", false, false],
            ["reshuffle-preview", "Reshuffle preview labels", "labels", true, false],
            ["reset-thumbnail-sizes", "Reset thumbnail sizes on Apply", "layout", false, true],
            ["recover-offscreen-thumbnails", "Recover off-screen thumbnails", "layout", true, false],
            ["capture-client-placement", "Capture open EVE window placement to draft", "window-behavior", true, false],
            ["detected-sources", "Refresh detected log sources", "sources", true, false],
            ["running-windows", "Pick running application windows", "app-groups", true, false],
            ["shortcut-conflicts", "Review shortcut conflicts", "shortcuts", false, false],
            ["show-log", "Show application log", "diagnostics", false, false],
            ["export-diagnostics", "Export diagnostic snapshot", "diagnostics", true, false],
            ["toggle-live-previews", "Pause/resume live previews", "diagnostics", true, false],
            ["live-preview-status", "Live preview status", "", true, false],
            ["preview-sound", "Preview sound", "sounds", false, false],
            ["stop-sound", "Stop sound preview", "sounds", false, false],
            ["check-updates", "Check for updates", "about", false, false],
            ["download-settings", "Download latest settings editor", "about", false, false],
            ["update-status", "Refresh update status", "about", false, false],
            ["install-stable", "Install stable version", "about", true, true],
            ["install-prerelease", "Install pre-release version", "about", true, true],
            ["changelog", "Read included changelog", "about", false, false],
            ["documentation", "Open documentation", "about", false, false],
            ["report-problem", "Report a problem on GitHub", "about", false, false]]
            result.Push(Map("id", item[1], "label", item[2], "section", item[3], "runtimeOnly", this.Bool(item[4]), "destructive", this.Bool(item[5])))
        return result
    }

    static Export() {
        defaults := JSON.Load(default_JSON), sections := [], fields := [], tables := [], events := [], commands := this.Commands()
        for section in SettingsSchema.Sections {
            entry := Map("id", section.id, "page", section.page, "title", section.title, "help", section.help,
                "group", section.group, "fields", section.fields.Clone(), "tables", [], "commands", [])
            for kind in this.TableKinds
                if SettingsTables.Group(kind) = section.id
                    entry["tables"].Push(kind)
            for command in commands
                if command["section"] = section.id
                    entry["commands"].Push(command["id"])
            sections.Push(entry)
        }
        for id, field in SettingsSchema.Fields {
            entry := this.Wire(field)
            entry["global"] := this.Bool(field.global)
            entry["default"] := SettingsModel.At(field.global ? defaults : defaults["_Profiles"]["Default"], field.path, "")
            fields.Push(entry)
        }
        for kind in this.TableKinds
            tables.Push(Map("kind", kind, "section", SettingsTables.Group(kind), "columns", this.Wire(SettingsTables.Columns(kind))))
        for category, ids in SettingsEventTable.Categories
            for id in ids
                events.Push(Map("id", id, "label", SettingsEventTable.Names[id], "category", category))
        return Map("bridgeVersion", this.BridgeVersion, "schemaVersion", this.SchemaVersion, "pages", SettingsSchema.Pages.Clone(),
            "sections", sections, "fields", fields, "tables", tables, "events", events, "commands", commands, "fonts", SettingsFonts.Names().Clone(),
            "gestures", this.Wire(SettingsTables.GestureOptions), "themes", SettingsPalettes.Export())
    }
}

class SettingsServiceError extends Error {
    __New(code, message, details := 0) {
        super.__New(message)
        this.code := code
        this.details := IsObject(details) ? details : []
    }
}

; One installation-specific mutex serializes canonical load/save in runtime and
; offline host. It is held only around file operations, never for an edit lease.
class SettingsFileLock {
    __New(installation := A_ScriptDir) {
        this.handle := DllCall("CreateMutexW", "Ptr", 0, "Int", false, "Str", "Local\EVE-X-Preview.SettingsFile/" SettingsBridge.InstallationId(installation), "Ptr")
        if !this.handle
            throw OSError()
        status := DllCall("WaitForSingleObject", "Ptr", this.handle, "UInt", 2500, "UInt")
        if status != 0 && status != 0x80 {
            DllCall("CloseHandle", "Ptr", this.handle)
            this.handle := 0
            throw SettingsServiceError("busy", "Another process is saving settings. Please try again.")
        }
    }
    Dispose() {
        if !this.HasOwnProp("handle") || !this.handle
            return
        DllCall("ReleaseMutex", "Ptr", this.handle)
        DllCall("CloseHandle", "Ptr", this.handle)
        this.handle := 0
    }
    __Delete() => this.Dispose()
}
