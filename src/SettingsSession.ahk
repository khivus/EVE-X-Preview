; Canonical draft operations shared by the external bridge and built-in Apply.
class SettingsSession {
    clientDirty := false
    __New(owner, recovery := 0, runtime := true) {
        this.owner := owner
        this.runtime := runtime
        this.id := SettingsBridge.NewId()
        this.model := SettingsModel(owner._JSON)
        if IsObject(recovery)
            this.Restore(recovery)
    }
    static ForModel(owner, model) {
        session := SettingsSession(owner)
        session.model := model
        return session
    }
    Dirty => this.model.Dirty || this.clientDirty

    Recovery() => Map("version", 1, "installation", this.InstallationId(), "original", this.model.original, "config", this.model.config,
        "profile", this.model.profile, "profileOrigins", this.model.profileOrigins, "saved", this.model.saved, "baseline", this.model.savedConfig,
        "undo", SettingsContract.Wire(this.model.undo), "clientDirty", SettingsContract.Bool(this.clientDirty))

    Restore(state) {
        if !(state is Map) || state.Get("version", 0) != 1
            throw SettingsServiceError("incompatible", "This saved settings draft uses an incompatible format.")
        if state.Get("installation", "") != this.InstallationId()
            throw SettingsServiceError("invalid-request", "This recovery draft belongs to another EVE-X-Preview installation.")
        for key in ["original", "config", "profileOrigins", "profile", "saved", "undo"]
            if !state.Has(key)
                throw SettingsServiceError("invalid-request", "The recovery draft is incomplete: " key)
        if !(state["original"] is Map) || !(state["config"] is Map) || !(state["profileOrigins"] is Map) || !(state["undo"] is Array)
            throw SettingsServiceError("invalid-request", "The recovery draft has invalid data types.")
        if !state["config"].Has("_Profiles") || !state["config"]["_Profiles"].Has(state["profile"])
            throw SettingsServiceError("invalid-request", "The recovery draft has no selected profile.")
        this.model := SettingsModel(state["config"])
        this.model.original := SettingsModel.Clone(state["original"])
        this.model.config := SettingsModel.Clone(state["config"])
        this.model.profile := state["profile"]
        this.model.profileOrigins := SettingsModel.Clone(state["profileOrigins"])
        this.model.savedConfig := state.Has("baseline") ? SettingsModel.Clone(state["baseline"]) : JSON.Load(state["saved"])
        this.model.saved := JSON.Dump(this.model.savedConfig)
        this.model.countSnapshot := ""
        this.model.undo := []
        for entry in state["undo"] {
            hydrated := {}
            for key, value in entry
                hydrated.%key% := SettingsModel.Clone(value)
            if hydrated.HasOwnProp("entries") {
                entries := []
                for row in hydrated.entries
                    entries.Push({index: row["index"], row: row["row"]})
                hydrated.entries := entries
            }
            this.model.undo.Push(hydrated)
        }
        this.clientDirty := !!state.Get("clientDirty", false)
    }

    InstallationId() => this.owner.HasOwnProp("settingsBridge") ? this.owner.settingsBridge.installationId : SettingsBridge.InstallationId(A_ScriptDir)

    Snapshot() {
        if this.HasOwnProp("_runtimeFacade") {
            this._runtimeFacade.model := this.model
            this._runtimeFacade._JSON := this.model.config
            SettingsSoundLibrary.RefreshPreviewGain(this._runtimeFacade)
        }
        values := Map(), profiles := [], inheritance := Map()
        for name in this.model.Profiles
            profiles.Push(Map("name", name, "active", SettingsContract.Bool(name = this.model.config["LastUsedProfile"])))
        for id, field in SettingsSchema.Fields {
            inherited := !field.global && this.model.Inherited(field.group)
            value := this.model.Get(id)
            values[id] := Map("value", field.kind = "bool" ? SettingsContract.Bool(value) : value,
                "inherited", SettingsContract.Bool(inherited), "enabled", SettingsContract.Bool(!inherited && (field.dependency = "" || this.model.Get(field.dependency))))
        }
        for group in SettingsModel.Groups
            inheritance[group] := SettingsContract.Bool(this.model.Inherited(group))
        creation := Map()
        for group, value in this.model.CreationInheritance()
            creation[group] := SettingsContract.Bool(value)
        return Map("session", this.id, "profile", this.model.profile, "profiles", profiles, "dirty", SettingsContract.Bool(this.Dirty), "changeCount", this.model.ChangeCount,
            "clientDirty", SettingsContract.Bool(this.clientDirty), "values", values, "canUndo", SettingsContract.Bool(this.model.undo.Length > 0),
            "runtime", SettingsContract.Bool(this.runtime), "draft", this.Recovery(), "inheritance", inheritance, "creationInheritance", creation)
    }

    Rows(kind, group := "", rows?) {
        SettingsTables.Group(kind)
        if IsSet(rows) {
            if !(rows is Array) || rows.Length > 10000
                throw ValueError("Enter at most 10000 table rows.")
            before := this.model.Clone(this.model.Rows(kind, group))
            if kind = "group-members" || kind = "app-group-members" {
                groups := this.model.Groups(kind = "group-members" ? "characters" : "apps")
                if !groups.Has(group)
                    throw ValueError("Select an existing group before editing its members.")
            }
            this.model.SaveRows(kind, rows, group)
            if !SettingsModel.Equal(before, rows)
                this.model.undo.Push({kind: "table-state", profile: this.model.profile, tableKind: kind, groupName: group, rows: before})
            return this.Snapshot()
        }
        names := [], shortcuts := []
        if kind = "group-members" || kind = "app-group-members" {
            groups := this.model.Groups(kind = "group-members" ? "characters" : "apps")
            for name in groups
                names.Push(name)
            if groups.Has(group) {
                labels := kind = "group-members" ? Map("ForwardsHotkey", "Next character", "BackwardsHotkey", "Previous character", "FirstCharHotkey", "First available character")
                    : Map("fkey", "Next application", "bkey", "Previous application")
                for key, label in labels
                    shortcuts.Push(Map("key", key, "label", label, "value", groups[group].Get(key, "")))
            }
        }
        return Map("columns", SettingsContract.Wire(SettingsTables.Columns(kind)), "rows", this.model.Rows(kind, group),
            "groups", names, "groupShortcuts", shortcuts, "inherited", SettingsContract.Bool(this.model.Inherited(SettingsTables.Group(kind))))
    }

    Group(args) {
        kind := args.Get("kind", "")
        if kind != "characters" && kind != "apps"
            throw ValueError("Choose character groups or application groups.")
        action := args.Get("action", ""), name := args.Get("name", "")
        switch action {
            case "add": this.model.AddGroup(kind, name)
            case "rename": this.model.RenameGroup(kind, name, args.Get("newName", ""))
            case "delete": this.model.DeleteGroup(kind, name)
            case "shortcut": this.model.SetGroupShortcut(kind, name, args.Get("key", ""), args.Get("value", ""))
            default: throw ValueError("Unknown group action.")
        }
        return this.Snapshot()
    }

    static EventFields(id) {
        if !SettingsEventTable.Names.Has(id)
            throw ValueError("Unknown monitored event.")
        fields := [
            Map("id", "enabled", "label", "Enable this event alert", "kind", "bool"),
            Map("id", "color", "label", "Border color", "kind", "color"),
            Map("id", "sound", "label", "Sound file (blank is silent)", "kind", "sound"),
            Map("id", "soundVolume", "label", "Event volume", "kind", "integer", "min", 0, "max", 300, "unit", "%")]
        if id = "stoppedShooting"
            fields.Push(Map("id", "shootingInterval", "label", "Milliseconds without an outgoing shot", "kind", "integer", "min", 100, "max", 3600000, "unit", "ms"))
        if id = "underAttackByPlayer" || id = "underAttackByNPC" {
            fields.Push(Map("id", "includeNeutralization", "label", "Include neutralization and Nosferatu drains", "kind", "bool"))
            fields.Push(Map("id", "includeMisses", "label", "Include shots that miss", "kind", "bool"))
        }
        if id = "underAttackByPlayer"
            fields.Push(Map("id", "ignoreSmartbombDamage", "label", "Ignore player smartbomb damage", "kind", "bool"))
        static defaults := JSON.Load(default_JSON)
        eventDefaults := defaults["_Profiles"]["Default"]["Game Logs Monitoring"]["monitoredEvents"][id]
        for field in fields {
            key := field["id"]
            field["default"] := key = "shootingInterval" ? defaults["_Profiles"]["Default"]["Game Logs Monitoring"]["shootingInterval"]
                : eventDefaults.Get(key, key = "soundVolume" ? 100 : key = "sound" ? "" : 0)
            field["help"] := field["label"]
        }
        return fields
    }

    Event(id, values?) {
        fields := SettingsSession.EventFields(id), events := this.model.Group("events", "Game Logs Monitoring")["monitoredEvents"]
        config := events[id]
        if IsSet(values) {
            if !(values is Map)
                throw ValueError("Event values must be named fields.")
            if this.model.Inherited("events")
                throw ValueError("Customize Event alerts before editing this event.")
            prepared := this.PrepareEvent(id, values), updated := prepared.config, idle := prepared.idle
            if updated.Get("sound", "") != "" && values.Has("sound") && !EventSounds.SameFile(updated["sound"], config.Get("sound", ""))
                EventSounds.Add(this.model.config, updated["sound"])
            if idle != ""
                this.model.Set("shootingInterval", idle)
            events[id] := updated
            return this.Snapshot()
        }
        data := SettingsModel.Clone(config)
        data["sound"] := config.Get("sound", "")
        data["soundVolume"] := config.Get("soundVolume", 100)
        for field in fields {
            key := field["id"]
            if key = "shootingInterval"
                data[key] := this.model.Get(key)
            else if field["kind"] = "bool"
                data[key] := SettingsContract.Bool(data.Get(key, field["default"]))
        }
        return Map("id", id, "label", SettingsEventTable.Names[id], "values", data, "fields", fields,
            "inherited", SettingsContract.Bool(this.model.Inherited("events")))
    }

    PrepareEvent(id, values) {
        fields := SettingsSession.EventFields(id)
        config := this.model.Group("events", "Game Logs Monitoring")["monitoredEvents"][id]
        updated := SettingsModel.Clone(config), idle := ""
        for key, value in values {
            descriptor := 0
            for field in fields
                if field["id"] = key
                    descriptor := field
            if !IsObject(descriptor) {
                if config.Has(key) && SettingsModel.Equal(config[key], value)
                    continue
                throw ValueError("Unknown event field: " key)
            }
            kind := descriptor["kind"]
            if kind = "sound" {
                value := Trim(value)
                if value != "" && !EventSounds.SameFile(value, config.Get("sound", ""))
                    EventSounds.Validate(value)
                updated[key] := value
            } else {
                parsed := SettingsValue.Parse({kind: kind, min: descriptor.Get("min", 0), max: descriptor.Get("max", 0), unit: descriptor.Get("unit", "")}, value)
                if key = "shootingInterval"
                    idle := parsed
                else
                    updated[key] := parsed
            }
        }
        return {config: updated, idle: idle}
    }

    EventRows(keys, rows) {
        if this.model.Inherited("events")
            throw ValueError("Customize Event alerts before editing inherited values.")
        if rows.Length != keys.Length
            throw ValueError("The event table has missing rows.")
        prepared := Map()
        for index, row in rows {
            if row.Length != 6
                throw ValueError("The event row has missing columns.")
            prepared[keys[index]] := this.PrepareEvent(keys[index], Map("enabled", row[1], "color", row[4], "sound", row[5], "soundVolume", row[6]))
        }
        events := this.model.Group("events", "Game Logs Monitoring")["monitoredEvents"]
        for id, item in prepared
            events[id] := item.config
    }

    SoundRows(rows) {
        paths := [], gains := Map(), used := Map()
        used.CaseSense := "Off"
        for index, row in rows {
            path := Trim(row[1])
            if path = ""
                throw ValueError("Row " index ": enter a full sound-file path, or remove the unfinished row.")
            known := false
            for saved in EventSounds.Library(this.model.config)
                known := known || EventSounds.SameFile(path, saved)
            if !known
                EventSounds.Validate(path)
            path := EventSounds.FullPath(path)
            if used.Has(path)
                throw ValueError("Row " index " duplicates an existing sound file.")
            used[path] := true
            gains[StrLower(path)] := SettingsValue.Parse({kind: "integer", min: 0, max: 300, unit: "%"}, row[3])
            paths.Push(path)
            SplitPath(path,, &directory)
            row[2] := directory
        }
        this.model.config["SoundLibrary"] := paths
        for path, gain in gains
            EventSounds.Volumes(this.model.config)[path] := gain
    }

    Sound(args) {
        action := args.Get("action", "list"), path := args.Get("path", "")
        events := this.model.Group("events", "Game Logs Monitoring")["monitoredEvents"]
        switch action {
            case "list":
                rows := []
                for saved in EventSounds.Library(this.model.config) {
                    assigned := []
                    for id, config in events
                        if EventSounds.SameFile(saved, config.Get("sound", ""))
                            assigned.Push(id)
                    rows.Push(Map("path", saved, "volume", EventSounds.SoundVolume(this.model.config, saved), "assigned", assigned))
                }
                return Map("rows", rows)
            case "add": EventSounds.Add(this.model.config, path)
            case "remove":
                this.model.undo.Push({kind: "library", paths: SettingsModel.Clone(EventSounds.Library(this.model.config))})
                EventSounds.Remove(this.model.config, path)
            case "gain":
                known := false
                for saved in EventSounds.Library(this.model.config)
                    known := known || EventSounds.SameFile(saved, path)
                if !known
                    throw ValueError("Choose a file in the sound library.")
                EventSounds.Volumes(this.model.config)[StrLower(path)] := SettingsValue.Parse({kind: "integer", min: 0, max: 300, unit: "%"}, args.Get("volume", ""))
            case "assign":
                if this.model.Inherited("events")
                    throw ValueError("Customize Event alerts before changing event assignments.")
                checked := args.Get("events", Map())
                if checked is Array {
                    selected := Map()
                    for id in checked
                        selected[id] := true
                    checked := selected
                }
                if !(checked is Map) || path = ""
                    throw ValueError("Select a sound and the events to assign.")
                for id in checked
                    if !SettingsEventTable.Names.Has(id)
                        throw ValueError("Unknown event assignment: " id)
                known := false
                for saved in EventSounds.Library(this.model.config)
                    known := known || EventSounds.SameFile(saved, path)
                if !known
                    EventSounds.Validate(path)
                plan := EventSounds.AssignmentPlan(events, path, checked)
                for id, sound in plan.changes
                    events[id]["sound"] := sound
            default: throw ValueError("Unknown sound-library action.")
        }
        return this.Snapshot()
    }

    Validate() {
        issues := [], conflicts := this.model.Conflicts(this.owner._JSON), baseline := SettingsModel(this.model.original)
        for issue in this.model.Relations()
            issues.Push(Map("id", issue.id, "profile", issue.profile, "message", issue.message, "severity",
                this.FieldUnchanged(baseline, issue.profile, issue.id) && this.FieldUnchanged(baseline, issue.profile, issue.relatedId) ? "warning" : "error"))
        for name in this.model.Profiles {
            for id, field in SettingsSchema.Fields {
                try SettingsValue.Parse(field, this.model.GetFor(id, name))
                catch as err
                    issues.Push(Map("id", id, "profile", name, "message", err.Message, "severity", this.FieldUnchanged(baseline, name, id) ? "warning" : "error"))
            }
            selected := this.model.profile
            try {
                this.model.SelectProfile(name)
                for kind in SettingsContract.TableKinds {
                    if kind = "group-members" || kind = "app-group-members" {
                        for group in this.model.Groups(kind = "group-members" ? "characters" : "apps") {
                            try SettingsTables.Validate(kind, this.model.Rows(kind, group))
                            catch as err
                                issues.Push(Map("id", kind, "profile", name, "message", err.Message, "severity", this.TableUnchanged(baseline, name, kind, group) ? "warning" : "error"))
                        }
                    } else {
                        try SettingsTables.Validate(kind, this.model.Rows(kind))
                        catch as err
                            issues.Push(Map("id", kind, "profile", name, "message", err.Message, "severity", this.TableUnchanged(baseline, name, kind) ? "warning" : "error"))
                    }
                }
                events := this.model.Group("events", "Game Logs Monitoring")["monitoredEvents"]
                for id in SettingsEventTable.Names {
                    if !events.Has(id) {
                        issues.Push(Map("id", id, "profile", name, "message", "The monitored event is missing.", "severity", "error"))
                        continue
                    }
                    for field in SettingsSession.EventFields(id) {
                        key := field["id"]
                        if key = "sound" || key = "shootingInterval"
                            continue ; Existing sound references may be moved or absent.
                        try SettingsValue.Parse({kind: field["kind"], min: field.Get("min", 0), max: field.Get("max", 0), unit: field.Get("unit", "")}, events[id].Get(key, field["default"]))
                        catch as err
                            issues.Push(Map("id", id ":" key, "profile", name, "message", err.Message, "severity", this.EventUnchanged(baseline, name, id) ? "warning" : "error"))
                    }
                }
            } finally this.model.SelectProfile(selected)
        }
        ; Shortcut overlaps are allowed intentionally, matching the internal UI.
        try {
            facade := this.Facade()
            for id, labels in HotkeyBindings.Conflicts(HotkeyBindings.Collect(facade))
                issues.Push(Map("id", id, "profile", this.model.profile, "section", HotkeyBindings.SectionFor(id), "message", "Shortcut overlaps: " this.Join(labels), "severity", "warning", "kind", "shortcut-conflict"))
        }
        hasErrors := false
        for issue in issues
            hasErrors := hasErrors || issue["severity"] = "error"
        return Map("issues", issues, "conflicts", conflicts, "canApply", SettingsContract.Bool(!hasErrors && !conflicts.Length && !this.clientDirty))
    }

    FieldUnchanged(baseline, profile, id) {
        origin := this.model.profileOrigins.Get(profile, profile)
        try return baseline.Profiles.Has(origin) && SettingsModel.Equal(this.model.GetFor(id, profile), baseline.GetFor(id, origin))
        catch
            return false
    }

    TableUnchanged(baseline, profile, kind, group := "") {
        origin := this.model.profileOrigins.Get(profile, profile)
        if !baseline.Profiles.Has(origin)
            return false
        baseline.SelectProfile(origin)
        try return SettingsModel.Equal(this.model.Rows(kind, group), baseline.Rows(kind, group))
        catch
            return false
    }

    EventUnchanged(baseline, profile, id) {
        origin := this.model.profileOrigins.Get(profile, profile)
        if !baseline.Profiles.Has(origin)
            return false
        baseline.SelectProfile(origin)
        try return SettingsModel.Equal(this.model.Group("events", "Game Logs Monitoring")["monitoredEvents"][id], baseline.Group("events", "Game Logs Monitoring")["monitoredEvents"][id])
        catch
            return false
    }

    Join(values) {
        result := ""
        for value in values
            result .= (result = "" ? "" : "; ") value
        return result
    }

    Apply() {
        validation := this.Validate()
        if validation["conflicts"].Length
            throw SettingsServiceError("conflict", "Settings changed while this draft was open. Discard and reopen, or resolve the conflicting changes.", validation["conflicts"])
        if !validation["canApply"]
            throw SettingsServiceError("validation", "Complete or correct the pending settings before applying.", validation["issues"])
        SettingsSession.Commit(this.owner, this.model)
        this.clientDirty := false
        selected := this.model.profile
        this.model := SettingsModel(this.owner._JSON)
        this.model.SelectProfile(selected)
        return Map("applied", JSON.true, "restartRequired", SettingsContract.Bool(this.runtime), "snapshot", this.Snapshot())
    }

    static Commit(owner, model) {
        previousCritical := A_IsCritical
        Critical(true)
        try {
            conflicts := model.Conflicts(owner._JSON)
            if conflicts.Length
                throw SettingsServiceError("conflict", "Settings changed while this draft was open. Review the changed settings before applying.", conflicts)
            before := owner._JSON
            owner._JSON := model.Merge(before)
            try owner.SaveJsonToFile()
            catch as err {
                owner._JSON := before
                if err is SettingsServiceError
                    throw err
                throw SettingsServiceError("save-failed", "Settings could not be saved. Your draft is still open. " err.Message)
            }
        } finally Critical(previousCritical)
    }

    Facade() {
        if !this.HasOwnProp("_runtimeFacade")
            this._runtimeFacade := SettingsSessionRuntime(this)
        this._runtimeFacade._JSON := this.model.config
        this._runtimeFacade.model := this.model
        this._runtimeFacade.ProfileOverride()
        return this._runtimeFacade
    }

    Command(name, args := 0) {
        args := args is Map ? args : Map(), descriptor := 0
        for command in SettingsContract.Commands()
            if command["id"] = name
                descriptor := command
        if !IsObject(descriptor)
            throw SettingsServiceError("unknown-method", "Unknown settings command.")
        if descriptor["runtimeOnly"] && !this.runtime
            throw SettingsServiceError("runtime-unavailable", "Start EVE-X-Preview to use this runtime action.")
        facade := this.Facade()
        switch name {
            case "recover-offscreen-thumbnails": facade.RecoverRuntimeThumbnails()
            case "reset-thumbnail-sizes":
                facade.StageSizeReset()
                return this.Snapshot()
            case "capture-client-placement":
                facade.StageClientPlacement()
                return this.Snapshot()
            case "test-thumbnail":
                if facade.HasOwnProp("settingsThumbnail") && IsObject(facade.settingsThumbnail)
                    facade.StopTestThumbnail()
                else {
                    if facade.ownsPreviewOwner
                        facade.S_Gui.Show("NoActivate x0 y0 w1 h1")
                    facade.settingsThumbnail := SettingsThumbnailPreview(facade)
                }
            case "stop-test-thumbnail": facade.StopTestThumbnail()
            case "reshuffle-preview":
                facade.StopTestThumbnail()
                if facade.ownsPreviewOwner
                    facade.S_Gui.Show("NoActivate x0 y0 w1 h1")
                facade.settingsThumbnail := SettingsThumbnailPreview(facade)
            case "toggle-live-previews": this.owner.ToggleLivePreviews()
            case "show-log": ProgramLog.ShowOwned(facade.S_Gui.Hwnd)
            case "export-diagnostics": SetTimer((*) => ProgramLog.ExportDialog(), -20)
            case "preview-sound":
                path := args.Get("path", "")
                event := args.Get("event", "")
                if event != "" {
                    events := this.model.Group("events", "Game Logs Monitoring")["monitoredEvents"]
                    if !events.Has(event)
                        throw ValueError("Choose a monitored event.")
                    previewConfig := SettingsModel.Clone(events[event])
                    path := args.Get("path", previewConfig.Get("sound", ""))
                    if path != "" && !EventSounds.SameFile(path, previewConfig.Get("sound", ""))
                        EventSounds.Validate(path)
                    if args.Has("volume")
                        previewConfig["soundVolume"] := SettingsValue.Parse({kind: "integer", min: 0, max: 300, unit: "%"}, args["volume"])
                    SettingsSoundLibrary.Preview(facade, path, previewConfig)
                } else
                    SettingsSoundLibrary.Preview(facade, path)
                if facade.settingsSoundPlayer.lastError != ""
                    throw ValueError(facade.settingsSoundPlayer.lastError)
            case "stop-sound": facade.settingsSoundPlayer.Stop()
            case "running-windows": return Map("rows", SettingsContract.Wire(facade.RunningWindows(args.Get("applications", true))))
            case "detected-sources": return Map("game", facade.ResolveGameLogsDirectory(), "chat", facade.ResolveChatLogsDirectory())
            case "shortcut-conflicts": return Map("groups", SettingsContract.Wire(HotkeyBindings.ConflictGroups(HotkeyBindings.Collect(facade))))
            case "check-updates": facade.updateChecker.Start(args.Get("manual", true))
            case "download-settings":
                if this.Dirty
                    throw SettingsServiceError("busy", "Apply or discard the settings draft before replacing the editor.")
                facade.SettingsDownloader.Start(facade.updateChecker.settingsUrl, facade.updateChecker.settingsTag)
            case "update-status": return Map("status", facade.SettingsDownloader.status != "" ? facade.SettingsDownloader.status : facade.updateChecker.status, "busy", SettingsContract.Bool(facade.updateChecker.busy || facade.SettingsDownloader.busy),
                "stable", facade.updateChecker.releaseTag, "prerelease", facade.updateChecker.preReleaseTag, "installed", facade.programVersion,
                "settingsVersion", facade.updateChecker.settingsTag, "settingsUrl", facade.updateChecker.settingsUrl,
                "settingsCurrent", SettingsContract.Bool(facade.SettingsDownloader.Current(facade.updateChecker.settingsTag)),
                "updateAvailable", SettingsContract.Bool(UpdateChecker.IsNewer(facade.updateChecker.releaseTag, facade.programVersion)),
                "stableLabel", facade.InstallLabel(facade.updateChecker.releaseTag, false), "prereleaseLabel", facade.InstallLabel(facade.updateChecker.preReleaseTag, true),
                "stableAvailable", SettingsContract.Bool(facade.updateChecker.releaseTag != "" && VerCompare(facade.updateChecker.releaseTag, facade.programVersion) != 0),
                "prereleaseAvailable", SettingsContract.Bool(facade.updateChecker.preReleaseTag != "" && VerCompare(facade.updateChecker.preReleaseTag, facade.programVersion) != 0))
            case "install-stable", "install-prerelease":
                if this.Dirty
                    throw SettingsServiceError("busy", "Apply or discard your settings draft before installing another version.")
                SetTimer((*) => facade.InstallVersion(name = "install-prerelease"), -20)
            case "changelog": SetTimer((*) => ChangelogWindow.Show(false, facade.S_Gui.Hwnd), -20)
            case "documentation": Run("https://github.com/khivus/EVE-X-Preview/blob/dev/README.MD")
            case "report-problem": Run("https://github.com/khivus/EVE-X-Preview/issues/new")
        }
        return Map("message", facade.message != "" ? facade.message : "Action completed.")
    }

    Dispose() {
        if this.HasOwnProp("_runtimeFacade")
            this._runtimeFacade.Dispose()
    }
}

; Reuse runtime adapters, update/preview implementation and profile projection
; without constructing SettingsWorkbench or the normal Main_Class runtime.
class SettingsSessionRuntime extends SettingsWorkbenchPages {
    message := ""
    __New(session) {
        this.session := session
        this.owner := session.owner
        this.model := session.model
        this._JSON := session.model.config
        this.default_JSON := JSON.Load(default_JSON)
        this.programVersion := UpdateChecker.CurrentVersion()
        this.updateChecker := UpdateChecker()
        this.settingsSoundPlayer := EventSoundPlayer()
        this.ownsPreviewOwner := true
        this.S_Gui := Gui("+ToolWindow -Caption", "Settings preview owner")
        ; Use the real editor window as the native dialog/preview owner when connected.
        if this.owner.HasOwnProp("settingsBridge") && this.owner.settingsBridge.HasExternal {
            for hwnd in WinGetList("ahk_pid " this.owner.settingsBridge.ExternalPid) {
                if DllCall("IsWindowVisible", "Ptr", hwnd) && InStr(WinGetTitle(hwnd), "EVE-X-Preview") {
                    this.S_Gui.Destroy()
                    this.S_Gui := {Hwnd: hwnd}
                    this.ownsPreviewOwner := false
                    break
                }
            }
        }

    }
    LastUsedProfile {
        get => this.model.profile
        set => this.model.profile := value
    }
    HasDraftChanges() => this.session.Dirty
    Changed() => this._JSON := this.model.config
    UpdateStatus(message) => this.message := message
    ShowError(message) => this.message := message
    StopTestThumbnail() {
        if this.HasOwnProp("settingsThumbnail") && IsObject(this.settingsThumbnail) {
            this.settingsThumbnail.Close()
            this.settingsThumbnail := 0
        }
        if this.ownsPreviewOwner
            this.S_Gui.Hide()
    }
    Dispose() {
        this.StopTestThumbnail()
        this.settingsSoundPlayer.Dispose()
        if this.updateChecker.busy
            this.updateChecker.Finish("Settings closed.")
        if this.ownsPreviewOwner
            this.S_Gui.Destroy()
    }
}
