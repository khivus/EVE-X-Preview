; Generate silent audio in the temporary folder; the app ships no sound assets.
CreateTestSound(name := "alert", milliseconds := 500) {
    path := A_Temp "\EVE-X-Preview-" name "-" A_TickCount ".wav"
    samples := milliseconds * 8
    wave := Buffer(44 + samples, 128)
    StrPut("RIFF", wave, 5, "CP0")
    NumPut("UInt", 36 + samples, wave, 4)
    StrPut("WAVEfmt ", wave.Ptr + 8, 9, "CP0")
    NumPut("UInt", 16, "UShort", 1, "UShort", 1, "UInt", 8000, "UInt", 8000, "UShort", 1, "UShort", 8, wave, 16)
    StrPut("data", wave.Ptr + 36, 5, "CP0")
    NumPut("UInt", samples, wave, 40)
    file := FileOpen(path, "w")
    file.RawWrite(wave)
    file.Close()
    return path
}

class EventSoundSettingsFixture extends ColorSettingsFixture {
    nextFile := ""
    SelectEventSoundFile() => this.nextFile
    SelectLibrary(row) {
        this.eventSoundDialog["SoundLibrary"].Modify(row, "Select Focus")
        this.RefreshSoundLibrarySelection()
    }
    EventRow(event) {
        for row, name in this.eventSoundDialog.assignmentEvents
            if name = event
                return row
        throw Error("Missing sound event row: " event)
    }
    ChangeVolume(scope, name, value) {
        control := this.eventSoundDialog[name]
        control.Value := value
        ; WM_HSCROLL / TB_THUMBTRACK exercises the actual slider event handler.
        PostMessage(0x114, 5 | (value << 16), control.Hwnd,, "ahk_id " control.Gui.Hwnd)
        Sleep(30)
    }
    EditVolume(name, text, commit := false) {
        control := this.eventSoundDialog[name]
        control.Value := text
        id := DllCall("GetDlgCtrlID", "Ptr", control.Hwnd, "Int")
        PostMessage(0x111, (0x300 << 16) | (id & 0xFFFF), control.Hwnd,, "ahk_id " control.Gui.Hwnd)
        Sleep(30) ; EN_CHANGE, as emitted by editing or pasting into the field.
        if commit {
            PostMessage(0x111, (0x200 << 16) | (id & 0xFFFF), control.Hwnd,, "ahk_id " control.Gui.Hwnd)
            Sleep(30) ; EN_KILLFOCUS commits formatting and range limits.
        }
    }
    Dispose() {
        this.CloseEventSoundDialog()
        if this.HasOwnProp("eventSoundMixer")
            this.eventSoundMixer.Dispose()
        super.Dispose()
    }
}

TestRunner.Register("Event sounds default to silent with an empty shared library and support older profiles", EventSoundDefaultsTest)
EventSoundDefaultsTest() {
    app := CombatEventsFixture()
    AssertEqual(0, EventSounds.Library(app._JSON).Length)
    for event, config in app.monitoredEvents
        AssertEqual("", config["sound"])
    app._JSON.Delete("SoundLibrary")
    app.monitoredEvents["warpDisrupted"].Delete("sound")
    AssertEqual(0, EventSounds.Library(app._JSON).Length)
    AssertFalse(EventSounds.AnyEnabled(app.monitoredEvents))
    app.monitoredEvents["warpDisrupted"]["sound"] := "missing.wav"
    AssertFalse(EventSounds.AnyEnabled(app.monitoredEvents), "Disabled events must not enable audio monitoring.")
    app.monitoredEvents["warpDisrupted"]["enabled"] := true
    AssertTrue(EventSounds.AnyEnabled(app.monitoredEvents))
}

TestRunner.Register("Sound library deduplicates file paths and removal preserves files and profile assignments", EventSoundLibraryTest)
EventSoundLibraryTest() {
    app := CombatEventsFixture()
    path := CreateTestSound("library space")
    try {
        saved := EventSounds.Add(app._JSON, path)
        AssertEqual(saved, EventSounds.Add(app._JSON, StrUpper(path)))
        SplitPath(path, &name, &directory)
        AssertEqual(saved, EventSounds.Add(app._JSON, directory "\.\" name))
        AssertEqual(1, EventSounds.Library(app._JSON).Length)
        app.monitoredEvents["warpDisrupted"]["sound"] := saved
        app.Profiles["Second"] := JSON.Load(JSON.Dump(app.Profiles["Default"]))
        app._JSON := JSON.Load(JSON.Dump(app._JSON))
        EventSounds.Remove(app._JSON, StrUpper(path))
        AssertEqual(0, EventSounds.Library(app._JSON).Length)
        for profile, config in app.Profiles
            AssertEqual(saved, config["Game Logs Monitoring"]["monitoredEvents"]["warpDisrupted"]["sound"])
        AssertTrue(FileExist(path))
        AssertThrows(() => EventSounds.Add(app._JSON, path ".missing"))
        AssertEqual(0, EventSounds.Library(app._JSON).Length)
    } finally FileDelete(path)
}

TestRunner.Register("Every event has a sound button beside its color swatch within the settings panel", EventSoundButtonsTest)
EventSoundButtonsTest() {
    app := EventSoundSettingsFixture()
    try {
        count := 0
        for event in app.monitoredEventsTexts {
            button := app.MainFrame["Sound" event]
            button.GetPos(&x, &y, &w, &h)
            app.MainFrame["PreviewC" event].GetPos(&cx, &cy, &cw)
            app.MainFrame["C" event].GetPos(&fx, &fy, &fw)
            AssertTrue(x >= cx + cw && x + w <= app.contentW - 10)
            AssertEqual(app.baseGrid, cx - fx - fw, "Color fields and swatches need the standard gap.")
            AssertEqual(app.baseGrid, x - cx - cw, "Sound buttons need the same gap after the swatch.")
            AssertEqual(app.contentGap, app.contentW - x - w, "Sound buttons need the same panel padding as the left edge.")
            AssertEqual(h, w, "Sound buttons should be square.")
            AssertEqual(cy - 1, y)
            AssertEqual(26, h)
            AssertTrue(InStr(button.ToolTip, "none selected"))
            AssertFalse(button.Visible)
            count++
        }
        AssertEqual(18, count)
        app.MainFrame["SoundwarpDisrupted"].Enabled := false
        app.OpenEventSoundDialog("warpDisrupted")
        AssertFalse(app.HasOwnProp("eventSoundDialog"), "Inherited settings must not be editable.")
    } finally app.Dispose()
}

TestRunner.Register("Sound dialog adds reuses clears and removes sounds without deleting original files", EventSoundDialogTest)
EventSoundDialogTest() {
    app := EventSoundSettingsFixture()
    path := CreateTestSound("UI")
    other := CreateTestSound("other UI")
    exportPath := A_Temp "\EVE-X-Preview-sound-config-" A_TickCount ".json"
    try {
        app.OpenEventSoundDialog("warpDisrupted")
        AssertTrue(DllCall("GetWindowLongW", "Ptr", app.S_Gui.Hwnd, "Int", -16, "UInt") & 0x08000000)
        app.BrowseEventSound("warpDisrupted") ; Cancel.
        AssertEqual("", app.monitoredEvents["warpDisrupted"]["sound"])
        AssertFalse(app.NeedRestart)
        app.nextFile := path
        app.BrowseEventSound("warpDisrupted")
        AssertEqual("", app.monitoredEvents["warpDisrupted"]["sound"], "Browsing stages assignments until Apply.")
        app.ApplySoundAssignments()
        AssertEqual(path, app.monitoredEvents["warpDisrupted"]["sound"])
        AssertEqual(1, EventSounds.Library(app._JSON).Length)
        AssertTrue(app.NeedRestart)
        AssertTrue(InStr(app.MainFrame["SoundwarpDisrupted"].ToolTip, path))
        app.CloseEventSoundDialog()
        AssertFalse(DllCall("GetWindowLongW", "Ptr", app.S_Gui.Hwnd, "Int", -16, "UInt") & 0x08000000)
        app.OpenEventSoundDialog("decloaked")
        AssertEqual(1, app.eventSoundDialog["SoundLibrary"].GetCount())
        app.SelectLibrary(1)
        app.ApplySoundAssignments()
        AssertEqual(path, app.monitoredEvents["decloaked"]["sound"])
        app.nextFile := other
        app.BrowseEventSound("decloaked")
        app.ApplySoundAssignments()
        app.SelectLibrary(1)
        app.RemoveLibrarySound()
        AssertEqual([other], EventSounds.Library(app._JSON))
        AssertEqual(other, app.monitoredEvents["decloaked"]["sound"])
        AssertEqual(path, app.monitoredEvents["warpDisrupted"]["sound"])
        AssertTrue(FileExist(path))
        app.AssignLibrarySound("decloaked", "")
        AssertEqual("", app.monitoredEvents["decloaked"]["sound"])
        AssertEqual([other], EventSounds.Library(app._JSON))
        app.SaveJsonToFile(exportPath)
        AssertDeepEqual(app._JSON, JSON.Load(FileRead(exportPath)))
        Sleep(220)
        AssertTrue(app.saves > 0)
    } finally {
        app.Dispose()
        FileDelete(path)
        FileDelete(other)
        if FileExist(exportPath)
            FileDelete(exportPath)
    }
}

TestRunner.Register("Sound settings handle missing files and failed selections without losing assignments", EventSoundMissingFileTest)
EventSoundMissingFileTest() {
    app := EventSoundSettingsFixture()
    try {
        path := A_Temp "\EVE-X-Preview-missing-" A_TickCount ".wav"
        app._JSON["SoundLibrary"] := [path]
        app.OpenEventSoundDialog("warpDisrupted")
        app.SelectLibrary(1)
        app.ApplySoundAssignments()
        AssertEqual(path, app.monitoredEvents["warpDisrupted"]["sound"])
        AssertTrue(InStr(app.eventSoundDialog["SoundStatus"].Text, "missing"))
        app.TestSettingsSound(path)
        AssertTrue(InStr(app.eventSoundDialog["SoundStatus"].Text, "not found"))
        app.nextFile := path
        app.BrowseEventSound("warpDisrupted")
        AssertEqual(path, app.monitoredEvents["warpDisrupted"]["sound"])
        AssertEqual([path], EventSounds.Library(app._JSON))
    } finally app.Dispose()
}

class RecordingEventSoundPlayer {
    activePath := ""
    activeVolume := 100
    lastError := ""
    __New(app) => this.app := app
    Play(path, restart := false, volume := 100) {
        this.activePath := path
        this.activeVolume := volume
        this.app.played.Push(path)
        this.app.volumes.Push(volume)
        return true
    }
    IsPlaying() => this.activePath != ""
    SetVolume(volume) => this.activeVolume := volume
    Dispose() => this.activePath := ""
}

class RecordingEventSoundMixer extends EventSoundMixer {
    __New(app) {
        this.app := app
        super.__New(app._JSON)
    }
    CreatePlayer() => RecordingEventSoundPlayer(this.app)
    ValidateFile(path) {
        ; Parser/dispatch fixtures record audio without requiring real media.
    }
}

class EventSoundRuntimeFixture extends CombatEventsFixture {
    focused := false
    sourceExists := true
    played := []
    volumes := []
    __New() {
        super.__New()
        this.ProfileOverride()
        this.HideThumbnails := false
        this.HideThumbForActiveWin := false
        this.supressForFocused := false
        this.flashBorderEnabled := false
        this.showEventText := false
        this.eventDisplayDuration := 2000
        this.eventMethods := Map()
        this.flashMethod := Map()
        this.shootingChars := Map()
        this.debugToolTipMethod := (*) => 0
        this.debugToolTipDelay := 60000
        this.monitoredEvents["warpDisrupted"]["enabled"] := true
        this.monitoredEvents["warpDisrupted"]["sound"] := "alert.wav"
        this.monitoredEvents["decloaked"]["enabled"] := true
        this.monitoredEvents["decloaked"]["sound"] := "cloak.wav"
        this.monitoredChars["Pilot"]["event"] := "warpDisrupted"
    }
    EventSourceWindow(title) => this.sourceExists ? 123 : 0
    EventSourceActive(hwnd) => this.focused
    CreateEventSoundMixer() => RecordingEventSoundMixer(this)
    Dispose() {
        for character, callback in this.eventMethods
            SetTimer(callback, 0)
        for character, callback in this.shootingChars
            SetTimer(callback, 0)
        SetTimer(this.debugToolTipMethod, 0)
        if this.HasOwnProp("monitorMethod")
            SetTimer(this.monitorMethod, 0)
        if this.HasOwnProp("eventSoundMixer")
            this.eventSoundMixer.Dispose()
    }
}

TestRunner.Register("Event sound assignments follow the effective event profile while sharing the library", EventSoundProfilesTest)
EventSoundProfilesTest() {
    app := EventSoundSettingsFixture()
    path := CreateTestSound("profiles")
    try {
        app.Profiles["Second"] := JSON.Load(JSON.Dump(app.Profiles["Default"]))
        app.SaveEventSound("warpDisrupted", path, true)
        app.ProfileMonitoredEvents := "Second"
        app.RefreshEventSoundButtons()
        AssertEqual("", app.monitoredEvents["warpDisrupted"]["sound"])
        AssertTrue(InStr(app.MainFrame["SoundwarpDisrupted"].ToolTip, "none selected"))
        app.SaveEventSound("decloaked", path, true)
        AssertEqual([path], EventSounds.Library(app._JSON))
        app.ProfileMonitoredEvents := "Default"
        app.RefreshEventSoundButtons()
        AssertEqual(path, app.monitoredEvents["warpDisrupted"]["sound"])
        AssertEqual("", app.monitoredEvents["decloaked"]["sound"])
        AssertTrue(InStr(app.MainFrame["SoundwarpDisrupted"].ToolTip, path))
    } finally {
        app.Dispose()
        FileDelete(path)
    }
}

TestRunner.Register("A batch of matching game log lines dispatches one asynchronous audio alert", EventSoundLogBatchTest)
EventSoundLogBatchTest() {
    app := EventSoundRuntimeFixture()
    path := A_Temp "\EVE-X-Preview-sound-log-" A_TickCount ".txt"
    file := 0
    try {
        content := ""
        loop 30
            content .= "[ 2026.10.16 12:00:00 ] (notify) Someone has started trying to warp scramble you!`r`n"
        FileAppend(content, path, "UTF-8")
        file := FileOpen(path, "r", "UTF-8")
        app.checkNPCs := false
        app.enabledMonitoredEvents := Map("warpDisrupted", Map("pattern", "you!", "needRegex", false))
        app.monitoredChars["Pilot"] := Map("event", "", "file", file, "pending", "")
        app.ReadGameLogUpdates("Pilot", false)
        AssertEqual(["alert.wav"], app.played)
        AssertEqual(1, app.eventMethods.Count)
        app.ReadGameLogUpdates("Pilot", false)
        AssertEqual(1, app.played.Length, "An unchanged log must not repeat its sound.")
    } finally {
        app.Dispose()
        if IsObject(file)
            file.Close()
        if FileExist(path)
            FileDelete(path)
    }
}

TestRunner.Register("Enabled event sounds initialize monitoring with event text borders and DPS disabled", EventSoundOnlyMonitoringTest)
EventSoundOnlyMonitoringTest() {
    app := EventSoundRuntimeFixture()
    try {
        app.gameLogsMonitoringEnabled := true
        app.monitorOnlySelectedChars := true
        app.charsToMonitor := []
        app.gameLogsMonitoring()
        AssertTrue(app.monitoringInitialized)
        AssertTrue(app.eventMonitoringEnabled)
        AssertFalse(app.dpsMonitoringEnabled)
        AssertTrue(app.enabledMonitoredEvents.Has("warpDisrupted"))
        AssertFalse(app.flashBorderEnabled)
        AssertFalse(app.showEventText)
    } finally app.Dispose()
}

TestRunner.Register("Audio alerts respect first last priority focus suppression and closed clients", EventSoundDispatchTest)
EventSoundDispatchTest() {
    app := EventSoundRuntimeFixture()
    try {
        app.HideThumbnails := true ; Audio works even when thumbnails are hidden.
        app.focused := true
        app.supressForFocused := true
        app.handleEventActivation("Pilot")
        AssertEqual([], app.played)
        app.supressForFocused := false
        app.HideThumbForActiveWin := true
        app.handleEventActivation("Pilot")
        AssertEqual([], app.played)
        app.HideThumbForActiveWin := false
        app.sourceExists := false
        app.handleEventActivation("Pilot")
        AssertEqual([], app.played)
        app.sourceExists := true
        app.lastEventPriority := false
        app.handleEventActivation("Pilot")
        AssertEqual(["alert.wav"], app.played)
        app.monitoredChars["Pilot"]["event"] := "decloaked"
        app.handleEventActivation("Pilot")
        AssertEqual(["alert.wav"], app.played)
        app.lastEventPriority := true
        app.handleEventActivation("Pilot")
        AssertEqual(["alert.wav", "cloak.wav"], app.played)
        app.handleEventActivation("Pilot")
        AssertEqual(2, app.played.Length, "Repeated hits must not restart audio.")
        app.endEvent("Pilot", 123)
        app.monitoredEvents["stoppedShooting"]["sound"] := "stopped.wav"
        app.handleEventActivation("Pilot", true)
        AssertEqual("stopped.wav", app.played[3])
    } finally app.Dispose()
}

TestRunner.Register("Repeated sound timeout is shared across characters and expires independently of event visuals", EventSoundCooldownTest)
EventSoundCooldownTest() {
    app := EventSoundRuntimeFixture()
    try {
        app.PlayEventSound("Pilot", "warpDisrupted", 10000)
        app.PlayEventSound("Pilot", "warpDisrupted", 11999)
        AssertEqual(1, app.played.Length)
        app.PlayEventSound("Pilot", "decloaked", 11000)
        app.PlayEventSound("Other", "warpDisrupted", 11000)
        app.PlayEventSound("Pilot", "warpDisrupted", 12000)
        AssertEqual(["alert.wav", "cloak.wav", "alert.wav"], app.played)
        app.monitoredEvents["warpDisrupted"]["sound"] := ""
        app.PlayEventSound("Pilot", "warpDisrupted", 15000)
        AssertEqual(3, app.played.Length)
    } finally app.Dispose()
}

TestRunner.Register("Native audio playback is asynchronous releases completed files and handles invalid media", EventSoundNativePlaybackTest)
EventSoundNativePlaybackTest() {
    path := CreateTestSound("Unicode-" Chr(0x416), 500)
    badPath := A_Temp "\EVE-X-Preview-bad-audio-" A_TickCount ".wav"
    player := EventSoundPlayer()
    try {
        started := A_TickCount
        AssertTrue(player.Play(path), player.lastError)
        AssertTrue(A_TickCount - started < 450, "Playback must not wait for the sound duration.")
        AssertTrue(player.IsPlaying())
        AssertFalse(player.Play(path), "An identical sound already playing must not restart.")
        AssertEqual("", player.lastError)
        deadline := A_TickCount + 3000
        while player.activePath != "" && A_TickCount < deadline
            Sleep(20)
        AssertEqual("", player.activePath, "Completed playback must close its media handle.")
        FileDelete(path)
        AssertFalse(player.Play(path))
        AssertTrue(InStr(player.lastError, "not found"))
        FileAppend("This is not an audio file.", badPath)
        AssertFalse(player.Play(badPath))
        AssertTrue(player.lastError != "")
        AssertEqual("", player.activePath)
    } finally {
        player.Dispose()
        if FileExist(path)
            FileDelete(path)
        if FileExist(badPath)
            FileDelete(badPath)
    }
}

TestRunner.Register("Six native alerts start within one input-latency budget after audio preparation", EventSoundBurstStartupTest)
EventSoundBurstStartupTest() {
    path := CreateTestSound("burst startup", 4000)
    preparationStarted := A_TickCount
    warmup := EventSoundPlayer()
    preparationMs := A_TickCount - preparationStarted
    mixer := EventSoundMixer(Map("MaxSimultaneousSounds", 6, "SoundRepeatTimeout", 0))
    try {
        AssertTrue(warmup.Play(path), warmup.lastError)
        warmup.Stop()
        started := A_TickCount
        loop 6
            AssertTrue(mixer.Play(path), mixer.lastError)
        elapsed := A_TickCount - started
        AssertTrue(elapsed < 450, "Preparing output once must keep a six-alert burst responsive; elapsed=" elapsed " ms.")
        TestRunner.Log("Audio startup: preparation=" preparationMs " ms; six-alert burst=" elapsed " ms.")
        AssertEqual(6, mixer.ActiveCount(), "Overlapping alerts must retain independent playback.")
        mixer.channels[1].player.Stop()
        AssertEqual(5, mixer.ActiveCount(), "Stopping one alert must leave the other voices playing.")
        mixer.Dispose()
        FileDelete(path)
        AssertFalse(FileExist(path), "Completed/disposed sources must release their original file.")
    } finally {
        warmup.Dispose()
        mixer.Dispose()
        if FileExist(path)
            FileDelete(path)
    }
}

TestRunner.Register("Master sound and event volumes multiply independently and retain backward-compatible defaults", EventSoundVolumeScopesTest)
EventSoundVolumeScopesTest() {
    app := EventSoundRuntimeFixture()
    try {
        config := app.monitoredEvents["warpDisrupted"]
        AssertEqual(100, EventSounds.EffectiveVolume(app._JSON, "alert.wav", config))
        app._JSON["MasterSoundVolume"] := 50
        EventSounds.Volumes(app._JSON)["alert.wav"] := 50
        config["soundVolume"] := 50
        AssertEqual(12.5, EventSounds.EffectiveVolume(app._JSON, "ALERT.WAV", config))
        AssertEqual(25, EventSounds.EffectiveVolume(app._JSON, "alert.wav"))
        AssertEqual(50, EventSounds.EffectiveVolume(app._JSON, "unlisted.wav"))
        AssertEqual(25, EventSounds.EffectiveVolume(app._JSON, "alert.wav", app.monitoredEvents["decloaked"]))
        app.PlayEventSound("Pilot", "warpDisrupted", 10000)
        AssertEqual([12.5], app.volumes)
        app.Profiles["Second"] := JSON.Load(JSON.Dump(app.Profiles["Default"]))
        app.ProfileMonitoredEvents := "Second"
        app.monitoredEvents["warpDisrupted"]["soundVolume"] := 20
        AssertEqual(5, EventSounds.EffectiveVolume(app._JSON, "alert.wav", app.monitoredEvents["warpDisrupted"]))
        app.ProfileMonitoredEvents := "Default"
        AssertEqual(12.5, EventSounds.EffectiveVolume(app._JSON, "alert.wav", app.monitoredEvents["warpDisrupted"]))
        app._JSON := JSON.Load(JSON.Dump(app._JSON))
        AssertEqual(12.5, EventSounds.EffectiveVolume(app._JSON, "alert.wav", app.monitoredEvents["warpDisrupted"]))
        app._JSON.Delete("MasterSoundVolume")
        app._JSON.Delete("SoundVolumes")
        app.monitoredEvents["warpDisrupted"].Delete("soundVolume")
        AssertEqual(100, EventSounds.EffectiveVolume(app._JSON, "alert.wav", app.monitoredEvents["warpDisrupted"]))
        for pair in [[-10, 0], [125, 125], [300, 300], [350, 300], ["bad", 100], [24.6, 25]]
            AssertEqual(pair[2], EventSounds.Percent(pair[1]))
    } finally app.Dispose()
}

TestRunner.Register("Muting any volume level skips audio and unmuting permits the next event", EventSoundMuteTest)
EventSoundMuteTest() {
    app := EventSoundRuntimeFixture()
    try {
        app._JSON["MasterSoundVolume"] := 0
        app.PlayEventSound("Pilot", "warpDisrupted", 10000)
        AssertEqual([], app.played)
        app._JSON["MasterSoundVolume"] := 100
        EventSounds.Volumes(app._JSON)["alert.wav"] := 0
        app.PlayEventSound("Pilot", "warpDisrupted", 10000)
        AssertEqual([], app.played)
        EventSounds.Volumes(app._JSON)["alert.wav"] := 100
        app.monitoredEvents["warpDisrupted"]["soundVolume"] := 0
        app.PlayEventSound("Pilot", "warpDisrupted", 10000)
        AssertEqual([], app.played)
        app.monitoredEvents["warpDisrupted"]["soundVolume"] := 100
        app.PlayEventSound("Pilot", "warpDisrupted", 10001)
        AssertEqual(["alert.wav"], app.played)
        AssertEqual([100], app.volumes)
    } finally app.Dispose()
}

TestRunner.Register("Bulk assignment warns before replacing other sounds and updates all checked events in one profile", EventSoundBulkAssignmentsTest)
EventSoundBulkAssignmentsTest() {
    app := EventSoundSettingsFixture()
    path := CreateTestSound("bulk")
    other := CreateTestSound("old bulk")
    try {
        app.SaveEventSound("warpDisrupted", other, true)
        app.SaveEventSound("decloaked", other)
        app.monitoredEvents["decloaked"]["soundVolume"] := 35
        app.Profiles["Second"] := JSON.Load(JSON.Dump(app.Profiles["Default"]))
        app.OpenEventSoundDialog("warpDisrupted")
        app.nextFile := path
        app.BrowseEventSound("warpDisrupted")
        AssertEqual(other, app.monitoredEvents["warpDisrupted"]["sound"])
        AssertTrue(InStr(app.eventSoundDialog["AssignmentNotice"].Text, "Will replace 1"))
        AssertTrue(InStr(app.eventSoundDialog["AssignmentNotice"].Text, "Warp Disrupted"))
        app.CheckAllSoundEvents(true)
        AssertEqual(18, app.CheckedSoundEvents().Count)
        AssertTrue(InStr(app.eventSoundDialog["ApplySounds"].Text, "replace 2"))
        AssertTrue(InStr(app.eventSoundDialog["AssignmentNotice"].Text, "Decloaked"))
        app.ApplySoundAssignments()
        for event, config in app.monitoredEvents {
            AssertEqual(path, config["sound"])
            AssertFalse(config["enabled"], "Assigning a sound must not silently enable its event.")
        }
        AssertEqual(35, app.monitoredEvents["decloaked"]["soundVolume"])
        AssertEqual(other, app.Profiles["Second"]["Game Logs Monitoring"]["monitoredEvents"]["warpDisrupted"]["sound"])
        AssertFalse(app.eventSoundDialog["ApplySounds"].Enabled)
        AssertTrue(InStr(app.eventSoundDialog["SoundStatus"].Text, "replaced 2"))
        app.CloseEventSoundDialog()
        app.OpenEventSoundDialog("warpDisrupted")
        AssertEqual(18, app.CheckedSoundEvents().Count)
    } finally {
        app.Dispose()
        FileDelete(path)
        FileDelete(other)
    }
}

TestRunner.Register("Unchecking events removes only the chosen sound and leaves other assignments intact", EventSoundUnassignTest)
EventSoundUnassignTest() {
    app := EventSoundSettingsFixture()
    path := CreateTestSound("unassign")
    other := CreateTestSound("keep")
    try {
        app.SaveEventSound("warpDisrupted", path, true)
        app.SaveEventSound("decloaked", path)
        app.SaveEventSound("gateJumped", other, true)
        app.OpenEventSoundDialog("warpDisrupted")
        app.CheckAllSoundEvents(false)
        AssertEqual(0, app.CheckedSoundEvents().Count)
        AssertFalse(InStr(app.eventSoundDialog["AssignmentNotice"].Text, "Will replace"))
        app.ApplySoundAssignments()
        AssertEqual("", app.monitoredEvents["warpDisrupted"]["sound"])
        AssertEqual("", app.monitoredEvents["decloaked"]["sound"])
        AssertEqual(other, app.monitoredEvents["gateJumped"]["sound"])
        AssertEqual([path, other], EventSounds.Library(app._JSON))
        AssertTrue(FileExist(path))
    } finally {
        app.Dispose()
        FileDelete(path)
        FileDelete(other)
    }
}

TestRunner.Register("Native sound dialog volume sliders save their own scopes without resetting pending assignments", EventSoundVolumeControlsTest)
EventSoundVolumeControlsTest() {
    app := EventSoundSettingsFixture()
    path := CreateTestSound("sliders")
    exportPath := A_Temp "\EVE-X-Preview-volume-settings-" A_TickCount ".json"
    try {
        app.SaveEventSound("warpDisrupted", path, true)
        app.SaveEventSound("decloaked", path)
        app.OpenEventSoundDialog("warpDisrupted")
        app.CheckAllSoundEvents(true)
        app.ChangeVolume("master", "MasterVolume", 50)
        AssertEqual(50, app._JSON["MasterSoundVolume"])
        app.ChangeVolume("sound", "SoundVolume", 40)
        AssertEqual(40, EventSounds.SoundVolume(app._JSON, StrUpper(path)))
        row := app.EventRow("decloaked")
        app.eventSoundDialog["SoundEvents"].Modify(row, "Select Focus")
        app.RefreshEventVolumeControl()
        app.ChangeVolume("event", "EventVolume", 25)
        AssertEqual(25, app.monitoredEvents["decloaked"]["soundVolume"])
        AssertEqual("25%", app.eventSoundDialog["SoundEvents"].GetText(row, 3))
        AssertEqual(100, app.monitoredEvents["warpDisrupted"]["soundVolume"])
        AssertEqual(5, EventSounds.EffectiveVolume(app._JSON, path, app.monitoredEvents["decloaked"]))
        AssertEqual(18, app.CheckedSoundEvents().Count, "Volume changes must preserve pending assignment checks.")
        AssertEqual("", app.monitoredEvents["gateJumped"]["sound"], "Checking alone must not save assignments.")
        app.ApplySoundAssignments()
        app.SaveJsonToFile(exportPath)
        AssertDeepEqual(app._JSON, JSON.Load(FileRead(exportPath)))
        app.CloseEventSoundDialog()
        app.OpenEventSoundDialog("decloaked")
        AssertEqual(50, app.eventSoundDialog["MasterVolume"].Value)
        AssertEqual(40, app.eventSoundDialog["SoundVolume"].Value)
        AssertEqual(25, app.eventSoundDialog["EventVolume"].Value)
        for control in app.eventSoundDialog {
            control.GetPos(&x, &y, &w, &h)
            AssertTrue(x >= 16 && x + w <= 884 && y + h <= 640, "Expanded sound controls must fit inside dialog padding.")
        }
        EventSounds.Remove(app._JSON, path)
        AssertEqual(40, EventSounds.SoundVolume(app._JSON, path), "Removing a library entry must retain the gain of existing assignments.")
    } finally {
        app.Dispose()
        FileDelete(path)
        if FileExist(exportPath)
            FileDelete(exportPath)
    }
}

TestRunner.Register("Native WAV volume controls the clip while live preview changes apply all three gains", EventSoundNativeVolumeTest)
EventSoundNativeVolumeTest() {
    app := EventSoundSettingsFixture()
    path := CreateTestSound("native volume", 2500)
    try {
        app.SaveEventSound("warpDisrupted", path, true)
        app._JSON["MasterSoundVolume"] := 50
        EventSounds.Volumes(app._JSON)[StrLower(path)] := 50
        app.monitoredEvents["warpDisrupted"]["soundVolume"] := 50
        app.OpenEventSoundDialog("warpDisrupted")
        systemVolume := SoundGetVolume()
        app.TestSettingsSound(path, "warpDisrupted")
        player := app.settingsSoundPlayer
        AssertEqual(12.5, player.activeVolume)
        AssertTrue(Abs(player.NativeVolume() - 12.5) < 0.001)
        app.ChangeVolume("event", "EventVolume", 100)
        AssertEqual(25, player.activeVolume)
        app.ChangeVolume("sound", "SoundVolume", 80)
        AssertEqual(40, player.activeVolume)
        app.ChangeVolume("master", "MasterVolume", 25)
        AssertEqual(20, player.activeVolume)
        AssertTrue(Abs(player.NativeVolume() - 20) < 0.001)
        AssertEqual(systemVolume, SoundGetVolume(), "Per-clip gains must not change the Windows volume.")
        app.ChangeVolume("master", "MasterVolume", 0)
        AssertEqual(0, player.activeVolume)
        AssertEqual(0, player.NativeVolume())
        app.TestSettingsSound(path)
        AssertEqual("", player.activePath, "A muted preview must stop instead of playing at full volume.")
        AssertTrue(InStr(app.eventSoundDialog["SoundStatus"].Text, "Muted"))
    } finally {
        app.Dispose()
        FileDelete(path)
    }
}

TestRunner.Register("All sound sliders persist 300 percent and the native mixer applies multiplied boost live", EventSoundBoostControlsTest)
EventSoundBoostControlsTest() {
    app := EventSoundSettingsFixture()
    path := CreateTestSound("boost controls", 2500)
    try {
        app.SaveEventSound("warpDisrupted", path, true)
        app.OpenEventSoundDialog("warpDisrupted")
        systemVolume := SoundGetVolume()
        app.ChangeVolume("master", "MasterVolume", 300)
        app.TestSettingsSound(path, "warpDisrupted")
        player := app.settingsSoundPlayer
        AssertTrue(Abs(player.NativeVolume() - 300) < 0.001, "300% must reach the mixer as gain 3.0.")
        app.ChangeVolume("sound", "SoundVolume", 300)
        AssertTrue(Abs(player.NativeVolume() - 900) < 0.001)
        app.ChangeVolume("event", "EventVolume", 300)
        AssertTrue(Abs(player.NativeVolume() - 2700) < 0.001, "Combined boost must not be clipped to the old 100% ceiling.")
        AssertEqual(2700, player.activeVolume)
        AssertEqual(2700, EventSounds.EffectiveVolume(app._JSON, path, app.monitoredEvents["warpDisrupted"]))
        AssertEqual(systemVolume, SoundGetVolume(), "Boost must not alter Windows volume.")
        app.EditVolume("MasterVolumeText", "123%", true)
        AssertTrue(Abs(player.NativeVolume() - 1107) < 0.001, "Exact text entry must update native gain without slider snapping.")
        app.EditVolume("MasterVolumeText", "300%", true)
        app._JSON := JSON.Load(JSON.Dump(app._JSON))
        app.CloseEventSoundDialog()
        app.OpenEventSoundDialog("warpDisrupted")
        for name in ["MasterVolume", "SoundVolume", "EventVolume"] {
            AssertEqual(300, app.eventSoundDialog[name].Value)
            AssertEqual("300%", app.eventSoundDialog[name "Text"].Text)
        }
    } finally {
        app.Dispose()
        FileDelete(path)
    }
}

TestRunner.Register("Long boosted sounds keep only a bounded streaming queue and release the original file on stop", EventSoundBoostStreamingTest)
EventSoundBoostStreamingTest() {
    path := CreateTestSound("long boost", 30000)
    player := EventSoundPlayer()
    try {
        originalSize := FileGetSize(path)
        AssertTrue(player.Play(path, false, 300), player.lastError)
        stream := player.stream
        AssertFalse(stream.endOfStream, "Playback must start before decoding an entire long file.")
        AssertTrue(stream.chunks.Length <= 32)
        AssertTrue(stream.queuedBytes <= EventAudioStream.maxQueuedBytes + EventAudioStream.maxSampleBytes)
        Sleep(150)
        AssertTrue(player.IsPlaying())
        AssertTrue(Abs(player.NativeVolume() - 300) < 0.001)
        AssertEqual(originalSize, FileGetSize(path), "Boost must not rewrite the user's file.")
        player.Stop()
        AssertEqual(0, stream.reader)
        AssertEqual(0, stream.voice)
        AssertEqual(0, stream.engine)
        AssertEqual(0, stream.queuedBytes)
        AssertEqual(0, stream.chunks.Length)
        FileDelete(path)
    } finally {
        player.Dispose()
        if FileExist(path)
            FileDelete(path)
    }
}

TestRunner.Register("MP3 event sounds decode stream and complete with native gain above 100 percent", EventSoundMP3BoostTest)
EventSoundMP3BoostTest() {
    path := A_Temp "\EVE-X-Preview-MP3-boost-" A_TickCount ".mp3"
    ; Independent silent MPEG-1 Layer III frames, 128 kbit/s at 44.1 kHz.
    ; Generate the fixture at runtime rather than distributing an audio asset.
    frame := Buffer(417, 0)
    NumPut("UChar", 0xFF, "UChar", 0xFB, "UChar", 0x90, "UChar", 0, frame)
    file := FileOpen(path, "w")
    loop 40
        file.RawWrite(frame)
    file.Close()
    player := EventSoundPlayer()
    try {
        AssertTrue(player.Play(path, false, 300), player.lastError)
        AssertTrue(player.IsPlaying())
        AssertTrue(Abs(player.NativeVolume() - 300) < 0.001)
        deadline := A_TickCount + 4000
        while player.activePath != "" && A_TickCount < deadline
            Sleep(20)
        AssertEqual("", player.activePath)
        AssertEqual("", player.lastError)
        FileDelete(path)
    } finally {
        player.Dispose()
        if FileExist(path)
            FileDelete(path)
    }
}

TestRunner.Register("Every sound slider snaps mouse input to five percent and reaches exact 100 percent", EventSoundSliderSnapTest)
EventSoundSliderSnapTest() {
    app := EventSoundSettingsFixture()
    path := CreateTestSound("snap")
    try {
        app.SaveEventSound("warpDisrupted", path, true)
        app.OpenEventSoundDialog("warpDisrupted")
        for name, scope in Map("MasterVolume", "master", "SoundVolume", "sound", "EventVolume", "event") {
            for pair in [[99, 100], [101, 100], [102, 100], [103, 105], [1, 0], [297, 295], [299, 300]] {
                app.ChangeVolume(scope, name, pair[1])
                AssertEqual(pair[2], app.eventSoundDialog[name].Value)
                AssertEqual(pair[2] "%", app.eventSoundDialog[name "Text"].Value)
                AssertEqual(pair[2], app.ReadSoundVolume(scope))
            }
        }
    } finally {
        app.Dispose()
        FileDelete(path)
    }
}

TestRunner.Register("Editable volume percentages accept exact values clamp bounds and recover unfinished input", EventSoundTextVolumeTest)
EventSoundTextVolumeTest() {
    app := EventSoundSettingsFixture()
    path := CreateTestSound("exact volume")
    try {
        app.SaveEventSound("warpDisrupted", path, true)
        app.OpenEventSoundDialog("warpDisrupted")
        app.CheckAllSoundEvents(true)
        for name, scope in Map("MasterVolume", "master", "SoundVolume", "sound", "EventVolume", "event") {
            field := app.eventSoundDialog[name "Text"]
            AssertEqual("Edit", field.Type)
            app.EditVolume(name "Text", "103", true)
            AssertEqual(103, app.ReadSoundVolume(scope))
            AssertEqual(103, app.eventSoundDialog[name].Value)
            AssertEqual("103%", field.Value)
            app.EditVolume(name "Text", "100%", true)
            AssertEqual(100, app.ReadSoundVolume(scope))
            app.EditVolume(name "Text", "")
            AssertEqual(100, app.ReadSoundVolume(scope), "Clearing an edit temporarily must not mute or reset it.")
            app.EditVolume(name "Text", "invalid", true)
            AssertEqual(100, app.ReadSoundVolume(scope))
            AssertEqual("100%", field.Value)
            app.EditVolume(name "Text", "999%", true)
            AssertEqual(300, app.ReadSoundVolume(scope))
            AssertEqual("300%", field.Value)
            app.EditVolume(name "Text", "-10%", true)
            AssertEqual(0, app.ReadSoundVolume(scope))
            AssertEqual("0%", field.Value)
            app.EditVolume(name "Text", " 123 % ", true)
            AssertEqual(123, app.ReadSoundVolume(scope))
            AssertEqual("123%", field.Value)
        }
        AssertEqual(18, app.CheckedSoundEvents().Count, "Text editing must keep pending assignments.")
        AssertEqual("", app.monitoredEvents["gateJumped"]["sound"])
        app._JSON := JSON.Load(JSON.Dump(app._JSON))
        app.CloseEventSoundDialog()
        app.OpenEventSoundDialog("warpDisrupted")
        for name in ["MasterVolume", "SoundVolume", "EventVolume"] {
            AssertEqual(123, app.eventSoundDialog[name].Value)
            AssertEqual("123%", app.eventSoundDialog[name "Text"].Value)
        }
        app.CloseEventSoundDialog()
        EventSounds.Remove(app._JSON, path)
        app.OpenEventSoundDialog("warpDisrupted")
        AssertFalse(app.eventSoundDialog["SoundVolume"].Enabled)
        AssertFalse(app.eventSoundDialog["SoundVolumeText"].Enabled)
    } finally {
        app.Dispose()
        FileDelete(path)
    }
}

TestRunner.Register("Removing the second saved sound during preview never indexes zero and keeps files and assignments", EventSoundRemovePlayingSecondTest)
EventSoundRemovePlayingSecondTest() {
    app := EventSoundSettingsFixture()
    first := CreateTestSound("first saved", 2500)
    second := CreateTestSound("second saved", 2500)
    try {
        app.SaveEventSound("warpDisrupted", first, true)
        app.SaveEventSound("decloaked", second, true)
        app.OpenEventSoundDialog("decloaked")
        app.TestSettingsSound(second)
        AssertTrue(app.settingsSoundPlayer.IsPlaying())
        app.RemoveLibrarySound()
        AssertEqual([first], EventSounds.Library(app._JSON))
        AssertEqual(1, app.eventSoundDialog["SoundLibrary"].GetCount())
        AssertEqual(second, app.monitoredEvents["decloaked"]["sound"])
        AssertTrue(app.settingsSoundPlayer.IsPlaying())
        AssertTrue(FileExist(first) && FileExist(second))
        ; Also cover multiple matching entries, where several removals shrink
        ; the array before the backward traversal reaches the first element.
        app._JSON["SoundLibrary"] := [first, second, StrUpper(second), second]
        EventSounds.Remove(app._JSON, second)
        AssertEqual([first], EventSounds.Library(app._JSON))
        EventSounds.Remove(app._JSON, first)
        AssertEqual([], EventSounds.Library(app._JSON))
        EventSounds.Remove(app._JSON, second)
        AssertEqual([], EventSounds.Library(app._JSON))
    } finally {
        app.Dispose()
        FileDelete(first)
        FileDelete(second)
    }
}

TestRunner.Register("Event sound mixer enforces capacity and shares repeated-file timeouts across events and characters", EventSoundMixerPolicyTest)
EventSoundMixerPolicyTest() {
    app := EventSoundRuntimeFixture()
    try {
        app._JSON["MaxSimultaneousSounds"] := 2
        app._JSON["SoundRepeatTimeout"] := 5000
        app.monitoredEvents["gateJumped"]["sound"] := "gate.wav"
        app.PlayEventSound("Pilot", "warpDisrupted", 10000)
        app.PlayEventSound("Other", "decloaked", 10001)
        mixer := app.eventSoundMixer
        AssertEqual(2, mixer.ActiveCount())
        oldest := mixer.channels[1].player
        second := mixer.channels[2].player
        app.PlayEventSound("Third", "gateJumped", 10002)
        AssertEqual(2, mixer.ActiveCount())
        AssertFalse(oldest.IsPlaying())
        AssertTrue(second.IsPlaying())
        app.monitoredEvents["decloaked"]["sound"] := "ALERT.WAV"
        app.PlayEventSound("Fourth", "decloaked", 12000)
        app.PlayEventSound("Other", "warpDisrupted", 14999)
        AssertEqual(3, app.played.Length, "A different event or character must not bypass the same-file timeout.")
        app.eventDisplayDuration := 60000
        app.PlayEventSound("Other", "warpDisrupted", 15000)
        AssertEqual(4, app.played.Length, "Visual alert duration must not extend the sound timeout.")
        app._JSON["SoundRepeatTimeout"] := 0
        app.PlayEventSound("Other", "warpDisrupted", 15001)
        app.PlayEventSound("Pilot", "decloaked", 15001)
        AssertEqual(6, app.played.Length)
        AssertEqual(2, mixer.ActiveCount())
        AssertTrue(EventSounds.SameFile(mixer.channels[1].player.activePath, mixer.channels[2].player.activePath))
        app._JSON["MaxSimultaneousSounds"] := 1
        mixer.Configure(app._JSON)
        AssertEqual(1, mixer.ActiveCount())
        mixer.channels[1].player.Dispose()
        AssertEqual(0, mixer.ActiveCount())
    } finally app.Dispose()
}

TestRunner.Register("Playback limits and timeout fields save globally preserve unfinished edits and clamp their ranges", EventSoundPlaybackSettingsTest)
EventSoundPlaybackSettingsTest() {
    app := EventSoundSettingsFixture()
    try {
        AssertEqual(1, EventSounds.MaxSimultaneous(app._JSON))
        AssertEqual(2000, EventSounds.RepeatTimeout(app._JSON))
        app.OpenEventSoundDialog("warpDisrupted")
        app.EditVolume("MaxSimultaneousSounds", "3", true)
        app.EditVolume("SoundRepeatTimeout", "1250", true)
        AssertEqual(3, EventSounds.MaxSimultaneous(app._JSON))
        AssertEqual(1250, EventSounds.RepeatTimeout(app._JSON))
        app.EditVolume("SoundRepeatTimeout", "")
        AssertEqual(1250, EventSounds.RepeatTimeout(app._JSON))
        app.EditVolume("SoundRepeatTimeout", "invalid", true)
        AssertEqual("1250", app.eventSoundDialog["SoundRepeatTimeout"].Value)
        app.EditVolume("MaxSimultaneousSounds", "100", true)
        AssertEqual("16", app.eventSoundDialog["MaxSimultaneousSounds"].Value)
        app.EditVolume("MaxSimultaneousSounds", "0", true)
        AssertEqual("1", app.eventSoundDialog["MaxSimultaneousSounds"].Value)
        app.EditVolume("SoundRepeatTimeout", "9999999", true)
        AssertEqual(3600000, EventSounds.RepeatTimeout(app._JSON))
        app.EditVolume("SoundRepeatTimeout", "-1", true)
        AssertEqual(0, EventSounds.RepeatTimeout(app._JSON))
        app.EditVolume("MaxSimultaneousSounds", "3", true)
        app.EditVolume("SoundRepeatTimeout", "1250", true)
        app._JSON := JSON.Load(JSON.Dump(app._JSON))
        app.CloseEventSoundDialog()
        app.Profiles["Second"] := JSON.Load(JSON.Dump(app.Profiles["Default"]))
        app.ProfileMonitoredEvents := "Second"
        app.OpenEventSoundDialog("warpDisrupted")
        AssertEqual("3", app.eventSoundDialog["MaxSimultaneousSounds"].Value)
        AssertEqual("1250", app.eventSoundDialog["SoundRepeatTimeout"].Value)
        app.CloseEventSoundDialog()
        app._JSON.Delete("MaxSimultaneousSounds")
        app._JSON.Delete("SoundRepeatTimeout")
        AssertEqual(1, EventSounds.MaxSimultaneous(app._JSON))
        AssertEqual(2000, EventSounds.RepeatTimeout(app._JSON))
    } finally app.Dispose()
}

TestRunner.Register("Native mixer plays concurrent clips updates all gains lowers limits live and permits overlapping repeats", EventSoundNativeMixerTest)
EventSoundNativeMixerTest() {
    app := EventSoundSettingsFixture()
    first := CreateTestSound("mix first", 4000)
    second := CreateTestSound("mix second", 4000)
    third := CreateTestSound("mix third", 4000)
    try {
        app.SaveEventSound("warpDisrupted", first, true)
        app.SaveEventSound("decloaked", second, true)
        app.SaveEventSound("gateJumped", third, true)
        app._JSON["MaxSimultaneousSounds"] := 3
        app._JSON["SoundRepeatTimeout"] := 3000
        app.PlayEventSoundFile(first, 100, "warpDisrupted", 10000)
        app.PlayEventSoundFile(second, 100, "decloaked", 10001)
        app.PlayEventSoundFile(third, 100, "gateJumped", 10002)
        mixer := app.eventSoundMixer
        AssertEqual(3, mixer.ActiveCount())
        oldest := mixer.channels[1].player
        missing := first ".missing.wav"
        AssertFalse(mixer.Play(missing, 100, "warpDisrupted", 10003))
        AssertTrue(InStr(mixer.lastError, "not found"))
        AssertEqual(3, mixer.ActiveCount(), "A missing file must not evict a working clip.")
        AssertFalse(mixer.lastStarted.Has(StrLower(missing)))
        app.OpenEventSoundDialog("warpDisrupted")
        app.EditVolume("MasterVolumeText", "50%", true)
        for channel in mixer.channels {
            AssertTrue(channel.player.IsPlaying())
            AssertTrue(Abs(channel.player.NativeVolume() - 50) < 0.001)
        }
        app.PlayEventSoundFile(first, 50, "warpDisrupted", 12999)
        AssertEqual(oldest, mixer.channels[1].player)
        app.PlayEventSoundFile(first, 50, "warpDisrupted", 13000)
        AssertFalse(oldest.IsPlaying())
        AssertEqual(3, mixer.ActiveCount())
        app.EditVolume("MaxSimultaneousSounds", "2", true)
        AssertEqual(2, mixer.ActiveCount())
        app.EditVolume("SoundRepeatTimeout", "0", true)
        app.PlayEventSoundFile(first, 50, "warpDisrupted", 13001)
        app.PlayEventSoundFile(first, 50, "warpDisrupted", 13001)
        AssertEqual(2, mixer.ActiveCount())
        for channel in mixer.channels
            AssertEqual(first, channel.player.activePath)
        app.EditVolume("MaxSimultaneousSounds", "1", true)
        AssertEqual(1, mixer.ActiveCount())
        app.TestSettingsSound(second)
        AssertTrue(app.settingsSoundPlayer.IsPlaying(), "Manual previews must ignore the event mixer limit.")
        AssertEqual(1, mixer.ActiveCount())
    } finally {
        app.Dispose()
        FileDelete(first)
        FileDelete(second)
        FileDelete(third)
    }
}

TestRunner.Register("Completed concurrent native clips are pruned together without invalid array indexes", EventSoundMixerCompletionTest)
EventSoundMixerCompletionTest() {
    first := CreateTestSound("complete first", 500)
    second := CreateTestSound("complete second", 500)
    mixer := EventSoundMixer(Map("MaxSimultaneousSounds", 2, "SoundRepeatTimeout", 0))
    try {
        AssertTrue(mixer.Play(first))
        AssertTrue(mixer.Play(second))
        AssertEqual(2, mixer.ActiveCount())
        deadline := A_TickCount + 3000
        while mixer.ActiveCount() && A_TickCount < deadline
            Sleep(20)
        AssertEqual(0, mixer.ActiveCount())
        FileDelete(first)
        FileDelete(second)
    } finally {
        mixer.Dispose()
        if FileExist(first)
            FileDelete(first)
        if FileExist(second)
            FileDelete(second)
    }
}
