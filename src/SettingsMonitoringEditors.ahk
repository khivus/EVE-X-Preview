class SettingsEventTable {
    static Categories := Map(
        "Combat", ["stoppedShooting", "underAttackByPlayer", "underAttackByNPC", "engagedWithFactionBSNPC", "engagedWithOfficerNPC", "engagedWithCapitalNPC", "warpDisrupted", "decloaked"],
        "Travel", ["gateJumped", "undockedFromNPCStation", "conduited"],
        "Fleet and communication", ["convoRequest", "fleetInvited", "fleetWarped", "fleetRegrouped"],
        "Mining", ["crystalBroke", "miningStopped", "miningBayIsFull"])
    static Names := Map("stoppedShooting", "Shooting idle", "underAttackByPlayer", "Player attack", "underAttackByNPC", "NPC attack",
        "engagedWithFactionBSNPC", "Faction battleship combat", "engagedWithOfficerNPC", "Officer NPC combat", "engagedWithCapitalNPC", "Capital NPC combat",
        "warpDisrupted", "Warp disrupted", "decloaked", "Decloaked", "gateJumped", "Gate jump", "undockedFromNPCStation", "Undocked from NPC station",
        "conduited", "Conduit jump", "convoRequest", "Conversation request", "fleetInvited", "Fleet invitation", "fleetWarped", "Fleet warp",
        "fleetRegrouped", "Fleet regroup", "crystalBroke", "Mining crystal broke", "miningStopped", "Mining stopped", "miningBayIsFull", "Mining hold full")

    __New(app) {
        this.app := app
        this.events := app.model.Group("Monitored Events", "Game Logs Monitoring")["monitoredEvents"]
        this.inherited := app.model.Inherited("Monitored Events")
    }

    Build() {
        app := this.app
        buttons := app.ButtonRow([
            {label: "Event details...", width: 150, callback: (*) => this.EditSelected()},
            {label: "Preview event sound", width: 176, callback: (*) => this.PreviewSelected()},
            {label: "Manage shared sounds", width: 184, callback: (*) => app.Navigate("sounds")}])
        this.edit := buttons[1], this.test := buttons[2]
        this.list := app.MainFrame.AddListView("x20 y" app.y " w" (app.contentWidth - 40) " h340 NoSortHdr", ["Enabled", "Event", "Category", "Border color", "Sound path", "Volume %"])
        this.list.OnEvent("ItemSelect", (*) => this.SelectionChanged())
        columns := [{label: "Enabled", kind: "bool"}, {label: "Event", kind: "readonly"}, {label: "Category", kind: "readonly"},
            {label: "Border color", kind: "color"}, {label: "Sound path", kind: "sound"}, {label: "Volume", kind: "number"}]
        rows := this.Rows()
        this.grid := SettingsGrid(app, "events", this.list, columns, rows, ObjBindMethod(this, "SaveRows"),
            {append: false, readonly: this.inherited, group: "Monitored Events", selection: ObjBindMethod(this, "SelectionChanged"), read: ObjBindMethod(this, "Rows")})
        this.Widths()
        app.y += 348
        this.detail := app.Note("Click values to edit. Shift/Ctrl+click selects cells; Ctrl+C/V copies/pastes; typing fills selected cells. Event names stay fixed.")
        app.eventTable := this
        this.SelectionChanged()
    }
    Rows() {
        rows := [], this.keys := []
        for category in ["Combat", "Travel", "Fleet and communication", "Mining"]
            for key in SettingsEventTable.Categories[category] {
                config := this.events[key]
                this.keys.Push(key)
                rows.Push([config["enabled"], SettingsEventTable.Names[key], category = "Fleet and communication" ? "Fleet / comms" : category,
                    config["color"], config.Get("sound", ""), EventSounds.Percent(config.Get("soundVolume", 100))])
            }
        return rows
    }
    SaveRows(rows) {
        if this.inherited
            throw ValueError("Customize Event alerts before editing inherited values.")
        normalized := []
        for index, row in rows {
            color := SettingsValue.Color(row[4])
            path := Trim(row[5])
            if path != "" && !EventSounds.SameFile(path, this.events[this.keys[index]].Get("sound", ""))
                EventSounds.Validate(path)
            volume := SettingsValue.Parse({kind: "integer", min: 0, max: 300, unit: "%"}, row[6])
            normalized.Push([!!row[1], color, path, volume])
        }
        for index, row in normalized {
            config := this.events[this.keys[index]]
            config["enabled"] := row[1], config["color"] := row[2], config["sound"] := row[3], config["soundVolume"] := row[4]
        }
    }
    Widths() {
        width := this.app.contentWidth - 62
        for index, fraction in [0.09, 0.25, 0.17, 0.15, 0.24, 0.10]
            this.list.ModifyCol(index, Max(index = 1 ? 58 : 75, Round(width * fraction)))
    }
    Refresh(select := 0) {
        this.grid.Reset(this.Rows(), select)
        this.Widths()
        this.SelectionChanged()
    }
    SelectionChanged() {
        row := this.list.GetNext()
        SettingsTheme.Assign(this.edit, "Enabled", row && !this.inherited)
        SettingsTheme.Assign(this.test, "Enabled", row && this.events[this.keys[row]].Get("sound", "") != "")
        if this.HasOwnProp("detail")
            SettingsTheme.Assign(this.detail, "Value", row ? SettingsTooltips.Events[this.keys[row]] : "Click a cell to edit. Shift/Ctrl selects cells; Ctrl+C/V copies/pastes.")
    }
    EditSelected() {
        this.grid.CommitEditor()
        if this.app.tableErrors.Has(this.grid.key) {
            this.app.LocateTableIssue()
            return
        }
        row := this.list.GetNext()
        if row && !this.inherited
            SettingsEventEditor(this, this.keys[row], row).Show()
    }
    PreviewSelected() {
        row := this.list.GetNext()
        if row {
            config := this.events[this.keys[row]]
            SettingsSoundLibrary.Preview(this.app, config.Get("sound", ""), config)
        }
    }
}

class SettingsEventEditor {
    __New(table, event, row) {
        this.table := table
        this.app := table.app
        this.event := event
        this.row := row
        this.config := SettingsModel.Clone(table.events[event])
    }

    Show() {
        app := this.app
        dialog := app.ChildDialog(SettingsEventTable.Names[this.event])
        this.dialog := dialog
        this.enabled := dialog.AddCheckbox("x20 y18 w660 h28 Checked" this.config["enabled"], "Enable this event alert")
        dialog.Help(this.enabled, SettingsTooltips.Events[this.event])
        dialog.AddText("x20 y56 w660 h40 c555555", SettingsTooltips.Events[this.event])
        dialog.AddText("x20 y106 w660 h20", "Border color")
        this.color := dialog.AddEdit("x20 y132 w480 h26", this.config["color"])
        dialog.Help(this.color, "Border color for this event. Enter an RGB hex code or use Choose color.")
        dialog.AddButton("x512 y132 w168 h26", "Choose color...").OnEvent("Click", (*) => this.PickColor())
        dialog.AddText("x20 y174 w660 h20", "Sound for this event")
        this.sound := dialog.AddDropDownList("x20 y200 w540", [])
        dialog.Help(this.sound, "Sound played for this event. None keeps the event silent.")
        dialog.AddButton("x572 y200 w108 h26", "Browse...").OnEvent("Click", (*) => this.Browse())
        this.path := dialog.AddText("x20 y234 w660 h38 c555555", "")
        this.sound.OnEvent("Change", (*) => this.UpdatePreview())
        dialog.AddText("x20 y282 w660 h20", "Event volume (0–300%)")
        this.volume := dialog.AddEdit("x598 y308 w82 h26", this.config.Get("soundVolume", 100))
        this.slider := dialog.AddSlider("x20 y305 w570 h32 Range0-300 AltSubmit TickInterval30 Line5 Page30", EventSounds.Percent(this.config.Get("soundVolume", 100)))
        SettingsLayout.SliderTicks(dialog, this.slider)
        dialog.Help(this.volume, "Volume for this event, multiplied by master volume and file gain. 0 mutes it.")
        dialog.Help(this.slider, this.volume.SettingsHelp)
        this.volume.OnEvent("Change", (*) => this.UpdatePreview())
        this.slider.OnEvent("Change", ObjBindMethod(this, "SliderChanged"))
        this.effective := dialog.AddText("x20 y346 w660 h40 c555555", "")
        dialog.AddButton("x20 y392 w220 h30", "Preview this event sound").OnEvent("Click", (*) => this.Preview())
        y := 438
        if this.event = "stoppedShooting" {
            dialog.AddText("x20 y" y " w660 h20", "Milliseconds without an outgoing shot (100–3600000)")
            this.idle := dialog.AddEdit("x20 y" (y + 26) " w220 h26", this.app.model.Get("shootingInterval"))
            dialog.Help(this.idle, "Trigger Shooting idle after this many milliseconds without an outgoing shot.")
            y += 68
        }
        if this.event = "underAttackByPlayer" || this.event = "underAttackByNPC" {
            this.neuts := dialog.AddCheckbox("x20 y" y " w660 h28 Checked" this.config.Get("includeNeutralization", 0), "Include incoming neutralization and Nosferatu drains")
            this.misses := dialog.AddCheckbox("x20 y" (y + 36) " w660 h28 Checked" this.config.Get("includeMisses", 1), "Include incoming shots that miss")
            dialog.Help(this.neuts, "Treat incoming energy neutralization and Nosferatu drains as this attack event.")
            dialog.Help(this.misses, "Trigger this attack alert even when the incoming shot misses.")
            y += 80
            if this.event = "underAttackByPlayer" {
                this.smartbomb := dialog.AddCheckbox("x20 y" y " w660 h28 Checked" this.config.Get("ignoreSmartbombDamage", 1), "Ignore player smartbomb damage")
                dialog.Help(this.smartbomb, "Exclude player smartbomb damage from this alert. DPS still counts that damage.")
                y += 40
            }
        }
        this.error := dialog.AddText("x20 y" y " w660 h26 cB42318", "")
        y += 34
        dialog.AddButton("x398 y" y " w136 h32 Default", "Save to draft").OnEvent("Click", (*) => this.Save())
        dialog.AddButton("x546 y" y " w134 h32", "Cancel").OnEvent("Click", (*) => app.CloseChild(dialog))
        this.pendingFile := ""
        this.RefreshSounds()
        dialog.Show("w700 h" (y + 52))
    }

    RefreshSounds(preferred?) {
        if !IsSet(preferred)
            preferred := this.config.Get("sound", "")
        this.paths := [""]
        labels := ["None — silent event"]
        library := EventSounds.Library(this.app.model.config).Clone()
        if preferred != "" {
            found := false
            for path in library
                if EventSounds.SameFile(path, preferred)
                    found := true
            if !found
                library.Push(preferred)
        }
        chosen := 1
        for path in library {
            SplitPath(path, &file, &directory)
            this.paths.Push(path)
            labels.Push(file " — " directory)
            if EventSounds.SameFile(path, preferred)
                chosen := this.paths.Length
        }
        this.sound.Delete()
        this.sound.Add(labels)
        this.sound.Choose(chosen)
        SendMessage(0x160, 760, 0, this.sound)
        this.UpdatePreview()
    }

    SelectedPath() => this.sound.Value ? this.paths[this.sound.Value] : ""

    ReadVolume() => SettingsValue.Parse({kind: "integer", min: 0, max: 300, unit: "%"}, this.volume.Value)

    SliderChanged(control := 0, info := 8) {
        this.sliderTracking := info = 5
        value := Round(this.slider.Value / 5) * 5
        if !this.sliderTracking
            SettingsTheme.Assign(this.slider, "Value", value)
        SettingsTheme.Assign(this.volume, "Value", value)
        this.UpdatePreview()
    }

    UpdatePreview() {
        path := this.SelectedPath()
        SettingsTheme.Assign(this.path, "Value", path = "" ? "This event will be silent." : path)
        try {
            volume := this.ReadVolume()
            if !this.HasOwnProp("sliderTracking") || !this.sliderTracking
                SettingsTheme.Assign(this.slider, "Value", volume)
            multiplier := EventSounds.SoundVolume(this.app.model.config, path)
            master := EventSounds.MasterVolume(this.app.model.config)
            this.effective.Value := "Effective preview volume: " Round(master * multiplier * volume / 10000, 2)
                . "% (master " master "% × file " multiplier "% × event " volume "%)."
        } catch
            this.effective.Value := "Enter a whole event percentage from 0 to 300."
    }

    PickColor() {
        color := ColorPicker.Show(this.dialog.Hwnd, this.color.Value)
        if color != ""
            this.color.Value := color
    }

    Browse() {
        this.dialog.Opt("+OwnDialogs")
        path := FileSelect(1,, "Choose an event sound", "Audio files (*.wav; *.mp3)")
        if path = ""
            return
        try {
            EventSounds.Validate(path)
            this.pendingFile := EventSounds.FullPath(path)
            this.RefreshSounds(this.pendingFile)
        } catch as err
            this.dialog.SetError(this.error, err.Message)
    }

    Preview() {
        try {
            config := this.config.Clone()
            config["soundVolume"] := this.ReadVolume()
            SettingsSoundLibrary.Preview(this.app, this.SelectedPath(), config)
        } catch as err
            this.dialog.SetError(this.error, err.Message)
    }

    Save() {
        try {
            config := SettingsModel.Clone(this.config)
            config["color"] := SettingsValue.Color(this.color.Value)
            config["enabled"] := this.enabled.Value
            config["sound"] := this.SelectedPath()
            config["soundVolume"] := this.ReadVolume()
            if config["sound"] != ""
                EventSounds.Validate(config["sound"])
            if this.HasOwnProp("idle")
                idle := SettingsValue.Parse(SettingsSchema.Fields["shootingInterval"], this.idle.Value)
            if this.HasOwnProp("neuts") {
                config["includeNeutralization"] := this.neuts.Value
                config["includeMisses"] := this.misses.Value
            }
            if this.HasOwnProp("smartbomb")
                config["ignoreSmartbombDamage"] := this.smartbomb.Value
            if this.app.model.Inherited("Monitored Events")
                throw ValueError("Customize Event alerts before editing this event.")
            if this.HasOwnProp("idle")
                this.app.model.Set("shootingInterval", idle)
            if this.pendingFile != "" && EventSounds.SameFile(this.pendingFile, config["sound"])
                EventSounds.Add(this.app.model.config, this.pendingFile)
            this.table.events[this.event] := config
            this.app.Changed()
            this.table.Refresh(this.row)
            this.app.CloseChild(this.dialog)
        } catch as err
            this.dialog.SetError(this.error, err.Message)
    }
}

class SettingsSoundLibrary {
    __New(app) => this.app := app

    Build() {
        app := this.app
        app.Subheading("Shared sound library")
        buttons := app.ButtonRow([
            {label: "Add sound...", width: 128, callback: (*) => this.Add()},
            {label: "Preview file", width: 128, callback: (*) => this.PreviewFile()},
            {label: "Remove reference", width: 152, callback: (*) => this.Remove()},
            {label: "Advanced file gain...", width: 176, callback: (*) => this.FileGain()},
            {label: "Batch assignments...", width: 176, callback: (*) => this.Batch()}])
        this.preview := buttons[2], this.removeButton := buttons[3], this.volume := buttons[4], this.batchButton := buttons[5]
        this.preview.ToolTip := "Preview this file using master volume and file gain; no event multiplier is used."
        this.list := app.MainFrame.AddListView("x20 y" app.y " w" (app.contentWidth - 40) " h230 NoSortHdr", ["Sound path", "Folder", "File gain %"])
        this.list.OnEvent("ItemSelect", (*) => this.Selected())
        this.paths := EventSounds.Library(app.model.config).Clone()
        this.grid := SettingsGrid(app, "sound-library", this.list,
            [{label: "Sound path", kind: "file"}, {label: "Folder", kind: "readonly"}, {label: "File gain", kind: "number"}],
            this.Rows(), ObjBindMethod(this, "SaveRows"), {shared: true, read: ObjBindMethod(this, "Rows"),
                selection: ObjBindMethod(this, "Selected"), blank: (*) => ["", "", 100], remove: ObjBindMethod(this, "RemoveRows")})
        app.y += 238
        this.notice := app.Note("Paste a full sound-file path into the empty row, or use Add sound. Ctrl+C/V copies/pastes; Ctrl+Del removes selected references.")
        this.Widths()
        this.Selected()
        app.soundLibrary := this
    }
    Rows() {
        rows := []
        for path in EventSounds.Library(this.app.model.config) {
            SplitPath(path,, &directory)
            rows.Push([path, directory, EventSounds.SoundVolume(this.app.model.config, path)])
        }
        return rows
    }
    SaveRows(rows) {
        paths := [], gains := Map(), used := Map()
        used.CaseSense := "Off"
        for index, row in rows {
            path := Trim(row[1])
            if path = ""
                throw ValueError("Row " index ": enter a full sound-file path, or remove the unfinished row.")
            known := false
            for saved in EventSounds.Library(this.app.model.config)
                if EventSounds.SameFile(path, saved)
                    known := true
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
        this.app.model.config["SoundLibrary"] := paths
        for path, gain in gains
            EventSounds.Volumes(this.app.model.config)[path] := gain
        this.paths := paths.Clone()
    }
    Widths() {
        width := this.app.contentWidth - 62
        this.list.ModifyCol(1, Round(width * 0.50))
        this.list.ModifyCol(2, Round(width * 0.32))
        this.list.ModifyCol(3, Max(90, Round(width * 0.18)))
    }
    Refresh(preferred := "") {
        this.paths := EventSounds.Library(this.app.model.config).Clone()
        this.grid.Reset(this.Rows())
        this.Widths()
        for index, path in this.paths
            if EventSounds.SameFile(path, preferred)
                this.list.Modify(index, "Select Focus Vis")
        this.Selected()
    }
    Path() {
        row := this.list.GetNext()
        return row && row <= this.paths.Length ? this.paths[row] : ""
    }
    RemoveRows(selected) {
        old := SettingsModel.Clone(EventSounds.Library(this.app.model.config)), rows := SettingsModel.Clone(this.grid.rows)
        index := rows.Length
        while index > 0 {
            if selected.Has(index)
                rows.RemoveAt(index)
            index--
        }
        if this.grid.SaveCandidate(rows) {
            this.app.model.undo.Push({kind: "library", paths: old})
            this.app.UpdateStatus("Library references removed. Assignments and original files are kept. Undo restores the references.")
        }
    }

    Selected() {
        if !this.HasOwnProp("notice")
            return
        path := this.Path()
        selected := path != ""
        for control in [this.preview, this.volume, this.removeButton]
            control.Enabled := selected
        this.batchButton.Enabled := selected && !this.app.model.Inherited("Monitored Events")
        this.notice.Value := !this.paths.Length ? "No sounds yet. Add your first sound, then assign it from Event alerts."
            : selected ? path (FileExist(path) ? "" : "`nFile missing. Add its new location or choose another file.")
            : "Select a sound file to preview or manage it. Removing a reference keeps assignments and the original file."
    }

    Ready() {
        this.grid.CommitEditor()
        if this.app.tableErrors.Has(this.grid.key) {
            this.app.LocateTableIssue()
            return false
        }
        return true
    }
    Add() {
        if !this.Ready()
            return
        this.app.S_Gui.Opt("+OwnDialogs")
        path := FileSelect(1,, "Add a sound to the shared library", "Audio files (*.wav; *.mp3)")
        if path = ""
            return
        try {
            path := EventSounds.Add(this.app.model.config, path)
            this.app.Changed()
            this.Refresh(path)
        } catch as err
            this.app.ShowError(err.Message)
    }

    Remove() => this.grid.RemoveRows()

    PreviewFile() => SettingsSoundLibrary.Preview(this.app, this.Path())

    static Preview(app, path, eventConfig?) {
        if path = ""
            return
        if !app.HasOwnProp("settingsSoundPlayer")
            app.settingsSoundPlayer := EventSoundPlayer()
        volume := IsSet(eventConfig) ? EventSounds.EffectiveVolume(app.model.config, path, eventConfig) : EventSounds.EffectiveVolume(app.model.config, path)
        app.previewSoundContext := {path: path, eventVolume: IsSet(eventConfig) ? EventSounds.Percent(eventConfig.Get("soundVolume", 100)) : 100}
        played := app.settingsSoundPlayer.Play(path, true, volume)
        app.UpdateStatus(!volume ? "Preview is muted by the draft volume settings."
            : played ? "Preview playing at " Round(volume, 2) "%." : app.settingsSoundPlayer.lastError)
    }

    static RefreshPreviewGain(app) {
        if !app.HasOwnProp("settingsSoundPlayer") || !app.HasOwnProp("previewSoundContext")
            return
        context := app.previewSoundContext
        if app.settingsSoundPlayer.activePath = "" || !EventSounds.SameFile(app.settingsSoundPlayer.activePath, context.path)
            return
        volume := EventSounds.MasterVolume(app.model.config) * EventSounds.SoundVolume(app.model.config, context.path) * context.eventVolume / 10000
        try app.settingsSoundPlayer.SetVolume(volume)
    }

    FileGain() {
        if !this.Ready()
            return
        path := this.Path()
        if path = ""
            return
        dialog := this.app.ChildDialog("Advanced sound-file gain")
        dialog.AddText("x20 y18 w560 h54", "This multiplier affects every event using this file, across profiles. Most sounds can stay at 100%.`n" path)
        dialog.AddText("x20 y86 w560 h22", "File gain (0–300%)")
        value := dialog.AddEdit("x20 y116 w120 h26", EventSounds.SoundVolume(this.app.model.config, path))
        dialog.Help(value, "Shared volume multiplier for this file. 100% keeps its original level; 0 mutes it.")
        error := dialog.AddText("x20 y156 w560 h42 cB42318", "")
        dialog.AddButton("x294 y214 w138 h32 Default", "Save gain to draft").OnEvent("Click", Save)
        dialog.AddButton("x444 y214 w136 h32", "Cancel").OnEvent("Click", (*) => this.app.CloseChild(dialog))
        dialog.Show("w600 h266")
        Save(*) {
            try {
                gain := SettingsValue.Parse({kind: "integer", min: 0, max: 300, unit: "%"}, value.Value)
                EventSounds.Volumes(this.app.model.config)[StrLower(path)] := gain
                this.app.Changed()
                this.Refresh(path)
                this.app.CloseChild(dialog)
            } catch as err
                dialog.SetError(error, err.Message)
        }
    }

    Batch() {
        if !this.Ready()
            return
        path := this.Path()
        if path = "" || this.app.model.Inherited("Monitored Events")
            return
        events := this.app.model.Group("Monitored Events", "Game Logs Monitoring")["monitoredEvents"]
        dialog := this.app.ChildDialog("Batch event sound assignments")
        dialog.AddText("x20 y18 w720 h66", "Checked events will use the selected sound. Unchecking removes this sound only; other sounds are kept. Nothing changes until Save to draft.`n" path)
        list := dialog.AddListView("x20 y98 w720 h340 NoSortHdr", ["Assign", "Event", "Current sound"])
        keys := [], holder := {rows: []}
        for key, config in events {
            keys.Push(key)
            current := config.Get("sound", "")
            SplitPath(current, &file)
            holder.rows.Push([EventSounds.SameFile(current, path) ? 1 : 0, SettingsEventTable.Names[key], current = "" ? "None" : file])
        }
        grid := SettingsGrid(this.app, "sound-batch|" path, list,
            [{label: "Assign", kind: "bool"}, {label: "Event", kind: "readonly"}, {label: "Current sound", kind: "readonly"}],
            holder.rows, (rows) => holder.rows := SettingsModel.Clone(rows),
            {append: false, local: true, parent: dialog.content, scroller: dialog.scroll, read: (*) => holder.rows})
        list.ModifyCol(1, 70), list.ModifyCol(2, 320), list.ModifyCol(3, 310)
        notice := dialog.AddText("x20 y450 w720 h42 c555555", "Select Assign cells and type 1 or 0 to change several events. Enabled states and volumes are kept.")
        dialog.AddButton("x20 y508 w120 h30", "Check all").OnEvent("Click", (*) => All(1))
        dialog.AddButton("x152 y508 w120 h30", "Uncheck all").OnEvent("Click", (*) => All(0))
        dialog.AddButton("x438 y508 w146 h30 Default", "Save to draft").OnEvent("Click", Save)
        dialog.AddButton("x596 y508 w144 h30", "Cancel").OnEvent("Click", (*) => this.app.CloseChild(dialog))
        dialog.Show("w760 h558")
        All(value) {
            grid.CommitEditor()
            grid.ColumnSelection(list, 1)
            grid.SetCells(grid.SelectedChanges(value))
        }
        Save(*) {
            grid.CommitEditor()
            if this.app.tableErrors.Has(grid.key) {
                notice.Value := this.app.tableErrors[grid.key].message
                return
            }
            checked := Map()
            for index, row in holder.rows
                if row[1]
                    checked[keys[index]] := true
            plan := EventSounds.AssignmentPlan(events, path, checked)
            for key, newPath in plan.changes
                events[key]["sound"] := newPath
            this.app.Changed()
            this.app.CloseChild(dialog)
            this.app.UpdateStatus(plan.changes.Count " event sound assignments saved to the draft. Apply & restart activates them.")
        }
    }
}
