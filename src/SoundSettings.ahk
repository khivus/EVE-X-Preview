#Include "ColorSettings.ahk"
#Include "EventSounds.ahk"

class SoundSettingsGui extends ColorSettingsGui {
    RefreshEventSoundButtons() {
        for event, config in this.monitoredEvents {
            button := this.MainFrame["Sound" event]
            path := config.Get("sound", "")
            button.Text := path = "" ? Chr(0x266A) : Chr(0x266B)
            button.ToolTip := path = "" ? "Choose an event sound (none selected)." : "Event sound: " path
        }
    }

    SaveEventSound(event, path, addToLibrary := false) {
        if addToLibrary
            path := EventSounds.Add(this._JSON, path)
        this.monitoredEvents[event]["sound"] := path
        this.SoundSettingsChanged(true)
    }

    SoundSettingsChanged(restart := false) {
        if restart {
            this.NeedRestart := 1
            this.RefreshEventSoundButtons()
        }
        SetTimer(this.Save_Settings_Delay_Timer, -200)
        this.RefreshPlayingSoundVolumes()
    }

    OpenEventSoundDialog(event, *) {
        if !this.MainFrame["Sound" event].Enabled
            return
        this.CancelHotkeyRecording()
        this.ClearSettingsTooltip()
        this.CloseEventSoundDialog()
        dialog := Gui("+Owner" this.S_Gui.Hwnd " -MinimizeBox -MaximizeBox", "Event Sounds - " this.monitoredEventsTexts[event])
        this.eventSoundDialog := dialog
        dialog.event := event
        dialog.selectedSound := ""
        dialog.refreshing := false
        dialog.SetFont("s10", "Segoe UI")
        dialog.Add("Text", "x16 y16 w104 h20", "Master volume")
        dialog.Add("Slider", "x120 y8 w224 h32 Range0-300 Line5 Page25 TickInterval100 AltSubmit vMasterVolume", EventSounds.MasterVolume(this._JSON))
        dialog.Add("Edit", "x352 y12 w72 h24 vMasterVolumeText")
        dialog.Add("Text", "x440 y8 w444 h36", "Volume = master " Chr(0xD7) " sound " Chr(0xD7) " event" Chr(10) "100% = original volume; above 100% = boost")
        dialog.Add("Text", "x16 y52 w408 h20", "Sound library - shared by all profiles")
        dialog.Add("Button", "x16 y76 w128 h28 vBrowseSound", "Browse / Add...")
        dialog.Add("Button", "x156 y76 w128 h28 vTestSound", "Test selected")
        dialog.Add("Button", "x296 y76 w128 h28 vRemoveSound", "Remove saved")
        dialog.Add("ListView", "x16 y112 w408 h220 -Multi NoSort NoSortHdr vSoundLibrary", ["Sound", "Folder"])
        dialog.Add("Edit", "x16 y344 w408 h42 ReadOnly -Wrap vSelectedSound")
        dialog.Add("Text", "x16 y404 w112 h20", "Sound volume")
        dialog.Add("Slider", "x136 y396 w224 h32 Range0-300 Line5 Page25 TickInterval100 AltSubmit vSoundVolume", 100)
        dialog.Add("Edit", "x368 y400 w56 h24 vSoundVolumeText")
        dialog.Add("Text", "x16 y436 w408 h20", "Opened event: " this.monitoredEventsTexts[event])
        dialog.Add("Edit", "x16 y460 w408 h42 ReadOnly -Wrap vAssignedSound")
        dialog.Add("Button", "x16 y512 w196 h28 vNoSound", "Clear opened event")
        dialog.Add("Button", "x228 y512 w196 h28 vTestAssigned", "Test opened event")
        dialog.Add("Text", "x16 y556 w408 h40 vSoundStatus", "Choose a sound, check events, then Apply. Removing a library entry keeps assignments and the original file.")
        dialog.Add("Text", "x440 y52 w444 h20", "Assign in profile: " this.ProfileMonitoredEvents)
        dialog.Add("Button", "x440 y76 w104 h28 vSelectAllEvents", "Select all")
        dialog.Add("Button", "x552 y76 w112 h28 vClearEventChecks", "Clear checks")
        dialog.Add("ListView", "x440 y112 w444 h300 Checked -Multi NoSort NoSortHdr vSoundEvents", ["Event", "Current sound", "Volume"])
        dialog.Add("Text", "x440 y424 w444 h20 vEventVolumeLabel")
        dialog.Add("Slider", "x440 y448 w380 h32 Range0-300 Line5 Page25 TickInterval100 AltSubmit vEventVolume", 100)
        dialog.Add("Edit", "x828 y452 w56 h24 vEventVolumeText")
        dialog.Add("Text", "x440 y488 w444 h72 vAssignmentNotice")
        dialog.Add("Button", "x440 y568 w288 h28 vApplySounds", "Apply assignments")
        dialog.Add("Button", "x736 y568 w148 h28 Default", "Done").OnEvent("Click", (*) => this.CloseEventSoundDialog())
        dialog["BrowseSound"].OnEvent("Click", (*) => this.BrowseEventSound(event))
        dialog["NoSound"].OnEvent("Click", (*) => this.AssignLibrarySound(event, ""))
        dialog["TestAssigned"].OnEvent("Click", (*) => this.TestSettingsSound(this.monitoredEvents[event].Get("sound", ""), event))
        dialog["SoundLibrary"].OnEvent("ItemSelect", (*) => this.RefreshSoundLibrarySelection())
        dialog["TestSound"].OnEvent("Click", (*) => this.TestSettingsSound(this.SelectedLibrarySound()))
        dialog["RemoveSound"].OnEvent("Click", (*) => this.RemoveLibrarySound())
        dialog["SelectAllEvents"].OnEvent("Click", (*) => this.CheckAllSoundEvents(true))
        dialog["ClearEventChecks"].OnEvent("Click", (*) => this.CheckAllSoundEvents(false))
        dialog["SoundEvents"].OnEvent("ItemCheck", (*) => this.RefreshSoundAssignmentNotice())
        dialog["SoundEvents"].OnEvent("ItemSelect", (*) => this.RefreshEventVolumeControl())
        dialog["ApplySounds"].OnEvent("Click", (*) => this.ApplySoundAssignments())
        for name, scope in Map("MasterVolume", "master", "SoundVolume", "sound", "EventVolume", "event") {
            dialog[name].OnEvent("Change", ObjBindMethod(this, "ChangeSoundVolume", scope))
            dialog[name "Text"].OnEvent("Change", ObjBindMethod(this, "ChangeSoundVolumeText", scope))
            dialog[name "Text"].OnEvent("LoseFocus", ObjBindMethod(this, "CommitSoundVolumeText", scope))
            dialog[name].ToolTip := SettingsTooltips.Description(name)
            dialog[name "Text"].ToolTip := SettingsTooltips.Description(name "Text")
        }
        dialog.OnEvent("Close", (*) => this.CloseEventSoundDialog())
        dialog.OnEvent("Escape", (*) => this.CloseEventSoundDialog())
        this.AddSoundPlaybackControls()
        SettingsTooltips.Apply(dialog)
        this.RefreshSoundLibrary(event,, true)
        this.S_Gui.Opt("+Disabled")
        dialog.Show("w900 h652")
    }

    AddSoundPlaybackControls() {
        dialog := this.eventSoundDialog
        ; Insert a shared playback row below the master volume controls.
        for control in dialog {
            control.GetPos(, &y)
            if y >= 52
                control.Move(, y + 40)
        }
        dialog.Add("Text", "x16 y56 w168 h20", "Simultaneous sounds")
        dialog.Add("Edit", "x192 y52 w56 h24 vMaxSimultaneousSounds", EventSounds.MaxSimultaneous(this._JSON))
        dialog.Add("Text", "x256 y56 w168 h20", "1-16; oldest replaced")
        dialog.Add("Text", "x440 y56 w208 h20", "Repeat timeout (ms)")
        dialog.Add("Edit", "x656 y52 w96 h24 vSoundRepeatTimeout", EventSounds.RepeatTimeout(this._JSON))
        dialog.Add("Text", "x760 y56 w124 h20", "0 = no timeout")
        for name in ["MaxSimultaneousSounds", "SoundRepeatTimeout"] {
            dialog[name].OnEvent("Change", ObjBindMethod(this, "ChangeSoundPlaybackSetting"))
            dialog[name].OnEvent("LoseFocus", ObjBindMethod(this, "CommitSoundPlaybackSetting"))
        }
        dialog["MaxSimultaneousSounds"].ToolTip := SettingsTooltips.Description("MaxSimultaneousSounds")
        dialog["SoundRepeatTimeout"].ToolTip := SettingsTooltips.Description("SoundRepeatTimeout")
    }

    ReadSoundPlaybackSetting(name) => name = "MaxSimultaneousSounds" ? EventSounds.MaxSimultaneous(this._JSON) : EventSounds.RepeatTimeout(this._JSON)

    ChangeSoundPlaybackSetting(control, *) {
        if !RegExMatch(control.Value, "^\s*[+-]?\d+\s*$")
            return
        input := Map(control.Name, Trim(control.Value))
        value := control.Name = "MaxSimultaneousSounds" ? EventSounds.MaxSimultaneous(input) : EventSounds.RepeatTimeout(input)
        if value = this.ReadSoundPlaybackSetting(control.Name)
            return
        this._JSON[control.Name] := value
        this.SoundSettingsChanged()
    }

    CommitSoundPlaybackSetting(control, *) {
        this.ChangeSoundPlaybackSetting(control)
        control.Value := this.ReadSoundPlaybackSetting(control.Name)
    }

    RefreshOpenedEventSound() {
        dialog := this.eventSoundDialog
        path := this.monitoredEvents[dialog.event].Get("sound", "")
        dialog["AssignedSound"].Value := path = "" ? "None" : path
        dialog["TestAssigned"].Enabled := path != ""
    }

    RefreshSoundLibrary(event, preferred?, includeOpened := false) {
        dialog := this.eventSoundDialog
        preferred := IsSet(preferred) ? preferred : this.monitoredEvents[event].Get("sound", "")
        dialog.refreshing := true
        try {
            list := dialog["SoundLibrary"]
            list.Delete()
            dialog.soundPaths := []
            selected := 0
            for saved in EventSounds.Library(this._JSON) {
                if !(saved is String) || saved = ""
                    continue
                SplitPath(saved, &name, &directory)
                row := list.Add(, name, directory)
                dialog.soundPaths.Push(saved)
                if EventSounds.SameFile(saved, preferred)
                    selected := row
            }
            if !selected && dialog.soundPaths.Length
                selected := 1
            if selected
                list.Modify(selected, "Select Focus")
            list.ModifyCol(1, 150)
            list.ModifyCol(2, 230)
        } finally dialog.refreshing := false
        this.RefreshOpenedEventSound()
        this.RefreshSoundLibrarySelection(true, includeOpened)
    }

    SelectedLibrarySound() {
        list := this.eventSoundDialog["SoundLibrary"]
        row := list.GetNext()
        return row ? this.eventSoundDialog.soundPaths[row] : ""
    }

    RefreshSoundLibrarySelection(force := false, includeOpened := true) {
        dialog := this.eventSoundDialog
        if dialog.refreshing
            return
        path := this.SelectedLibrarySound()
        if !force && EventSounds.SameFile(path, dialog.selectedSound)
            return
        dialog.selectedSound := path
        enabled := path != ""
        for name in ["TestSound", "RemoveSound", "SoundVolume", "SoundVolumeText", "SelectAllEvents", "ClearEventChecks"]
            dialog[name].Enabled := enabled
        dialog["SelectedSound"].Value := enabled ? path : "No library sound selected"
        dialog["SoundVolume"].Value := EventSounds.SoundVolume(this._JSON, path)
        dialog["SoundVolumeText"].Text := dialog["SoundVolume"].Value "%"
        dialog["MasterVolumeText"].Text := EventSounds.MasterVolume(this._JSON) "%"
        this.RefreshSoundEventList(includeOpened)
    }

    RefreshSoundEventList(includeOpened := false) {
        dialog := this.eventSoundDialog
        dialog.refreshing := true
        try {
            list := dialog["SoundEvents"]
            list.Delete()
            dialog.assignmentEvents := []
            path := dialog.selectedSound
            for event, config in this.monitoredEvents {
                current := config.Get("sound", "")
                checked := path != "" && (EventSounds.SameFile(current, path) || (includeOpened && event = dialog.event))
                SplitPath(current, &name)
                row := list.Add(checked ? "Check" : "", this.monitoredEventsTexts[event], current = "" ? "None" : name,
                    EventSounds.Percent(config.Get("soundVolume", 100)) "%")
                dialog.assignmentEvents.Push(event)
                if event = dialog.event
                    list.Modify(row, "Select Focus Vis")
            }
            list.ModifyCol(1, 218)
            list.ModifyCol(2, 150)
            list.ModifyCol(3, 50)
        } finally dialog.refreshing := false
        this.RefreshEventVolumeControl()
        this.RefreshSoundAssignmentNotice()
    }

    SelectedSoundEvent() {
        dialog := this.eventSoundDialog
        row := dialog["SoundEvents"].GetNext()
        return row ? dialog.assignmentEvents[row] : dialog.event
    }

    RefreshEventVolumeControl() {
        dialog := this.eventSoundDialog
        if dialog.refreshing
            return
        event := this.SelectedSoundEvent()
        dialog["EventVolumeLabel"].Text := "Event volume: " this.monitoredEventsTexts[event]
        dialog["EventVolume"].Value := EventSounds.Percent(this.monitoredEvents[event].Get("soundVolume", 100))
        dialog["EventVolumeText"].Text := dialog["EventVolume"].Value "%"
    }

    CheckedSoundEvents() {
        dialog := this.eventSoundDialog
        checked := Map()
        row := 0
        while row := dialog["SoundEvents"].GetNext(row, "C")
            checked[dialog.assignmentEvents[row]] := true
        return checked
    }

    CheckAllSoundEvents(checked) {
        if this.eventSoundDialog.selectedSound = ""
            return
        this.eventSoundDialog["SoundEvents"].Modify(0, checked ? "Check" : "-Check")
        this.RefreshSoundAssignmentNotice()
    }

    RefreshSoundAssignmentNotice() {
        dialog := this.eventSoundDialog
        if dialog.refreshing
            return
        plan := EventSounds.AssignmentPlan(this.monitoredEvents, dialog.selectedSound, this.CheckedSoundEvents())
        dialog["ApplySounds"].Enabled := plan.changes.Count > 0
        count := plan.replacements.Length
        dialog["ApplySounds"].Text := count ? "Apply - replace " count " other sound(s)" : "Apply assignments"
        notice := dialog["AssignmentNotice"]
        notice.SetFont(count ? "cA34D00" : "c444444")
        if dialog.selectedSound = "" {
            notice.Text := "Choose a library sound to assign it to events."
            return
        }
        names := ""
        for index, event in plan.replacements {
            if index > 2 {
                names .= ", and " (count - 2) " more"
                break
            }
            names .= (names = "" ? "" : ", ") this.monitoredEventsTexts[event]
        }
        notice.Text := count ? "Will replace " count " other sound(s): " names "." : "Checked events will use this sound."
        notice.Text .= Chr(10) "Unchecking removes this sound only; other sounds are kept."
    }

    ApplySoundAssignments() {
        dialog := this.eventSoundDialog
        path := dialog.selectedSound
        plan := EventSounds.AssignmentPlan(this.monitoredEvents, path, this.CheckedSoundEvents())
        if !plan.changes.Count
            return
        for event, newPath in plan.changes
            this.monitoredEvents[event]["sound"] := newPath
        this.SoundSettingsChanged(true)
        this.RefreshOpenedEventSound()
        this.RefreshSoundEventList()
        dialog["SoundStatus"].Text := "Applied " plan.changes.Count " assignment change(s); replaced " plan.replacements.Length " other sound(s)."
        if !FileExist(path)
            dialog["SoundStatus"].Text .= " File is missing; Browse to select its new location."
    }

    BrowseEventSound(event) {
        try {
            path := this.SelectEventSoundFile()
            if path = ""
                return
            path := EventSounds.Add(this._JSON, path)
            this.SoundSettingsChanged()
            this.RefreshSoundLibrary(event, path, true)
            this.eventSoundDialog["SoundStatus"].Text := "Added to the library. Check events and Apply to assign this sound."
        } catch as err
            this.eventSoundDialog["SoundStatus"].Text := err.Message
    }

    SelectEventSoundFile() {
        this.eventSoundDialog.Opt("+OwnDialogs")
        return FileSelect(1,, "Select an event sound", "Audio files (*.wav; *.mp3)")
    }

    AssignLibrarySound(event, path) {
        this.SaveEventSound(event, path)
        this.RefreshOpenedEventSound()
        this.RefreshSoundEventList()
        this.eventSoundDialog["SoundStatus"].Text := path = "" ? "The opened event is silent." : "Sound assigned."
    }

    RemoveLibrarySound() {
        path := this.SelectedLibrarySound()
        if path = ""
            return
        EventSounds.Remove(this._JSON, path)
        this.SoundSettingsChanged()
        this.RefreshSoundLibrary(this.eventSoundDialog.event)
        this.eventSoundDialog["SoundStatus"].Text := "Removed from the library. Assignments, sound volume and the original file are kept."
    }

    ChangeSoundVolume(scope, control, *) {
        value := EventSounds.Percent(Round(control.Value / 5) * 5)
        control.Value := value
        this.eventSoundDialog[control.Name "Text"].Value := value "%"
        this.ApplySoundVolume(scope, value)
    }

    ReadSoundVolume(scope) {
        if scope = "master"
            return EventSounds.MasterVolume(this._JSON)
        if scope = "sound"
            return EventSounds.SoundVolume(this._JSON, this.eventSoundDialog.selectedSound)
        return EventSounds.Percent(this.monitoredEvents[this.SelectedSoundEvent()].Get("soundVolume", 100))
    }

    ChangeSoundVolumeText(scope, control, *) {
        if !control.Enabled || !RegExMatch(control.Value, "^\s*([+-]?\d+)\s*%?\s*$", &match)
            return ; Let the user clear or finish a field without overwriting it.
        value := EventSounds.Percent(match[1])
        name := SubStr(control.Name, 1, -4) ; Remove the Text suffix.
        this.eventSoundDialog[name].Value := value
        this.ApplySoundVolume(scope, value)
    }

    CommitSoundVolumeText(scope, control, *) {
        this.ChangeSoundVolumeText(scope, control)
        control.Value := this.ReadSoundVolume(scope) "%"
    }

    ApplySoundVolume(scope, value) {
        dialog := this.eventSoundDialog
        if value = this.ReadSoundVolume(scope)
            return
        if scope = "master"
            this._JSON["MasterSoundVolume"] := value
        else if scope = "sound" {
            if dialog.selectedSound = ""
                return
            EventSounds.Volumes(this._JSON)[StrLower(dialog.selectedSound)] := value
        } else {
            event := this.SelectedSoundEvent()
            this.monitoredEvents[event]["soundVolume"] := value
            for row, name in dialog.assignmentEvents {
                if name = event {
                    dialog["SoundEvents"].Modify(row,,,, value "%")
                    break
                }
            }
        }
        this.SoundSettingsChanged()
    }

    TestSettingsSound(path, event := "") {
        if path = ""
            return
        if !this.HasOwnProp("settingsSoundPlayer")
            this.settingsSoundPlayer := EventSoundPlayer()
        this.settingsPreviewEvent := event
        volume := event = "" ? EventSounds.EffectiveVolume(this._JSON, path) : EventSounds.EffectiveVolume(this._JSON, path, this.monitoredEvents[event])
        played := this.settingsSoundPlayer.Play(path, true, volume)
        this.eventSoundDialog["SoundStatus"].Text := !volume ? "Muted by the volume settings." : played ? "Playing at " Round(volume, 2) "%." : this.settingsSoundPlayer.lastError
    }

    RefreshPlayingSoundVolumes() {
        if this.HasOwnProp("eventSoundMixer") {
            this.eventSoundMixer.Configure(this._JSON)
            for channel in this.eventSoundMixer.channels {
                player := channel.player
                event := channel.event
                volume := this.monitoredEvents.Has(event) ? EventSounds.EffectiveVolume(this._JSON, player.activePath, this.monitoredEvents[event]) : EventSounds.EffectiveVolume(this._JSON, player.activePath)
                try player.SetVolume(volume)
            }
        }
        if this.HasOwnProp("settingsSoundPlayer") && this.settingsSoundPlayer.activePath != "" {
            player := this.settingsSoundPlayer
            event := this.settingsPreviewEvent
            volume := this.monitoredEvents.Has(event) ? EventSounds.EffectiveVolume(this._JSON, player.activePath, this.monitoredEvents[event]) : EventSounds.EffectiveVolume(this._JSON, player.activePath)
            try player.SetVolume(volume)
        }
    }

    CloseEventSoundDialog() {
        if this.HasOwnProp("eventSoundDialog") {
            for name in ["MaxSimultaneousSounds", "SoundRepeatTimeout"]
                this.CommitSoundPlaybackSetting(this.eventSoundDialog[name])
            for name, scope in Map("MasterVolume", "master", "SoundVolume", "sound", "EventVolume", "event")
                this.CommitSoundVolumeText(scope, this.eventSoundDialog[name "Text"])
        }
        if this.HasOwnProp("settingsSoundPlayer") {
            this.settingsSoundPlayer.Dispose()
            this.DeleteProp("settingsSoundPlayer")
        }
        if this.HasOwnProp("eventSoundDialog") {
            this.ClearSettingsTooltip()
            this.S_Gui.Opt("-Disabled")
            this.eventSoundDialog.Destroy()
            this.DeleteProp("eventSoundDialog")
        }
    }
}
