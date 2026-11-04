class SettingsWorkbenchOwnerFixture extends CombatEventsFixture {
    saves := 0
    failSave := false
    lifecycleActions := []
    __New() {
        super.__New()
        this.ProfileOverride()
    }
    SaveJsonToFile(*) {
        if this.failSave
            throw Error("Fixture save failed")
        this.saves++
    }
    PerformRuntimeAction(action) => this.lifecycleActions.Push(action)
}

class SettingsWorkbenchFixture extends SettingsWorkbench {
    messages := []
    restarts := 0
    played := []
    volumes := []
    __New(owner?) {
        super.__New(IsSet(owner) ? owner : SettingsWorkbenchOwnerFixture(), false)
    }
    ShowError(message) => this.messages.Push(message)
    RestartAfterApply() => this.restarts++
    Binding(id) {
        for binding in this.bindings
            if binding.field.id = id
                return binding
        throw Error("Missing workbench field: " id)
    }
}

TestRunner.Register("Workbench builds every subsection with seven persistent navigation areas", WorkbenchAllSectionsTest)
WorkbenchAllSectionsTest() {
    app := SettingsWorkbenchFixture()
    try {
        AssertEqual(7, app.navRoots.Length)
        AssertTrue(WinGetStyle(app.nav.Hwnd) & 0x20, "Navigation selection remains visible after focus moves.")
        AssertFalse(app.model.Dirty)
        for section in SettingsSchema.Sections {
            app.Navigate(section.id)
            AssertEqual(section.id, app.renderedSection)
            AssertTrue(app.pageControls.Length >= 2)
            for binding in app.bindings {
                AssertTrue(binding.control.SettingsHelp != "")
                AssertTrue(binding.field.help != "", "A setting needs an explanation beyond its label: " binding.field.id)
                if binding.field.kind = "bool"
                    AssertEqual(binding.field.label, binding.control.Text, "Checkboxes must have meaningful accessible names.")
            }
        }
        AssertFalse(app.model.Dirty, "Browsing settings must not change the draft.")
        AssertEqual(0, app.owner.saves)
        app.ToggleTestThumbnail()
        preview := app.settingsThumbnail
        app.model.Set("ThumbnailTextColor", "123456")
        app.Changed()
        preview.Refresh(A_TickCount + 400)
        AssertEqual("123456", preview.sample["TextOverlay"].SystemTextDefaultColor)
        app.StopTestThumbnail()
        AssertTrue(preview.closed)
        AssertNotEqual("123456", app.owner.ThumbnailTextColor, "Previewing the draft must not modify running appearance settings.")
    } finally app.Dispose()
}

TestRunner.Register("Workbench unfinished invalid edits stay visible block Apply and recover without clamping", WorkbenchValidationTest)
WorkbenchValidationTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.Navigate("performance")
        field := app.Binding("ThumbnailPollingInterval")
        field.control.Value := "5000"
        app.TextChanged(field, field.control)
        AssertEqual("5000", field.control.Value)
        AssertEqual(50, app.model.Get("ThumbnailPollingInterval"))
        AssertEqual(1, app.errors.Count)
        AssertFalse(app.applyButton.Enabled)
        app.Navigate("labels")
        app.Navigate("performance")
        field := app.Binding("ThumbnailPollingInterval")
        AssertEqual("5000", field.control.Value, "Rebuilding or changing pages must not hide unfinished input.")
        field.control.Value := "250"
        app.CommitField(field, field.control)
        AssertEqual(250, app.model.Get("ThumbnailPollingInterval"))
        AssertEqual(0, app.errors.Count)
        AssertTrue(app.applyButton.Enabled)
        AssertEqual(50, app.owner.ThumbnailPollingInterval)
        app.Navigate("dps")
        AssertFalse(app.Binding("incomingDPSThreshold").control.Enabled)
        parent := app.Binding("incomingDPSEnabled")
        parent.control.Value := 1
        app.BooleanChanged(parent, parent.control)
        AssertTrue(app.Binding("incomingDPSThreshold").control.Enabled)
    } finally app.Dispose()
}

TestRunner.Register("Workbench preserves page scroll positions and fits common monitor scaling", WorkbenchScrollBoundsTest)
WorkbenchScrollBoundsTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.settingsScroller.ScrollTo(240)
        original := app.settingsScroller.offset
        AssertTrue(original > 0)
        app.Navigate("layout")
        app.Navigate("appearance")
        AssertEqual(original, app.settingsScroller.offset)
        for monitor in [[0, 0, 1280, 720, 96], [0, 0, 1920, 1040, 192], [-1920, 0, 0, 1040, 144]] {
            bounds := SettingsWindowSizing.Fit(monitor*)
            scale := monitor[5] / 96
            AssertTrue(bounds.width * scale <= monitor[3] - monitor[1])
            AssertTrue(bounds.height * scale <= monitor[4] - monitor[2])
            AssertTrue(bounds.x >= monitor[1] && bounds.y >= monitor[2])
        }
        app.width := 800, app.height := 500
        app.Layout()
        app.closeButton.GetPos(&x, &y, &w, &h)
        app.status.GetPos(&sx, &sy, &sw, &sh)
        AssertTrue(sy + sh <= y, "Compact layout places the status above footer buttons.")
        AssertTrue(x + w < app.width)
    } finally app.Dispose()
}

TestRunner.Register("Workbench profile browsing inheritance and activation are explicit draft operations", WorkbenchProfileTest)
WorkbenchProfileTest() {
    owner := SettingsWorkbenchOwnerFixture()
    owner.Profiles["Fleet"] := SettingsModel.Clone(owner.Profiles["Default"])
    owner.Profiles["Default"]["Other"]["Global_Groups"]["Thumbnails Visuals"] := 1
    owner.Profiles["Fleet"]["Other"]["Global_Groups"]["Thumbnails Visuals"] := 0
    before := JSON.Dump(owner._JSON)
    app := SettingsWorkbenchFixture(owner)
    try {
        app.SelectProfile("Fleet")
        AssertEqual("Default", owner.LastUsedProfile)
        AssertEqual("Default", app.model.config["LastUsedProfile"])
        app.Navigate("appearance")
        AssertFalse(app.Binding("ThumbnailTextSize").control.Enabled)
        app.model.SetInheritance("Thumbnails Visuals", false)
        app.Changed()
        app.Render()
        AssertTrue(app.Binding("ThumbnailTextSize").control.Enabled)
        app.model.Set("ThumbnailTextSize", 18)
        app.model.ActivateProfile()
        AssertEqual("Fleet", app.model.config["LastUsedProfile"])
        AssertEqual(before, JSON.Dump(owner._JSON))
    } finally app.Dispose()
}

TestRunner.Register("Workbench row dialogs save a complete assignment and Cancel leaves it untouched", WorkbenchRowDialogTest)
WorkbenchRowDialogTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.model.SaveRows("character-shortcuts", [["Alice", "^F1"], ["Bob", "!F2"]])
        app.Navigate("character-shortcuts")
        editor := app.tableEditors[1]
        before := SettingsModel.Clone(editor.rows)
        editor.Edit(1)
        editor.inputs[1].Value := "Changed Pilot"
        app.CloseChild(editor.dialog)
        AssertDeepEqual(before, app.model.Rows("character-shortcuts"))
        editor.Edit(1)
        editor.inputs[1].Value := "Changed Pilot"
        editor.SetShortcut(2, "^F23")
        editor.SaveDialog()
        AssertFalse(app.HasOwnProp("activeDialog"))
        AssertDeepEqual(["Changed Pilot", "^F23"], app.model.Rows("character-shortcuts")[1])
        editor.list.Modify(1, "Select Focus")
        editor.RemoveSelected()
        AssertNotEqual("Changed Pilot", app.model.Rows("character-shortcuts")[1][1])
        app.Undo()
        AssertDeepEqual(["Changed Pilot", "^F23"], app.model.Rows("character-shortcuts")[1])
    } finally app.Dispose()
}

TestRunner.Register("Workbench event editor stages conditions sound and volume together", WorkbenchEventEditorTest)
WorkbenchEventEditorTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.Navigate("events")
        table := app.eventTable
        before := SettingsModel.Clone(table.events["underAttackByPlayer"])
        edit := SettingsEventEditor(table, "underAttackByPlayer", 2)
        edit.Show()
        canvas := app.MainFrame.Hwnd
        app.Render()
        AssertEqual(canvas, app.MainFrame.Hwnd, "A resize/redraw must preserve the table underneath an open editor until it saves or cancels.")
        for control in [edit.enabled, edit.color, edit.sound, edit.volume, edit.slider, edit.neuts, edit.misses, edit.smartbomb]
            AssertTrue(control.SettingsHelp != "", "Every event setting must offer concise keyboard and hover help.")
        edit.dialog.SetError(edit.error, "A long validation explanation that must remain fully visible at a narrow dialog width. " . "Correct the highlighted cell before saving. " . "All other settings remain staged and can still be changed.")
        edit.error.GetPos(,, &ew, &eh)
        AssertTrue(eh >= SettingsLayout.TextHeight(edit.error, edit.error.Value, ew))
        edit.volume.Value := "145"
        edit.neuts.Value := 1
        app.CloseChild(edit.dialog)
        table := app.eventTable
        AssertDeepEqual(before, table.events["underAttackByPlayer"])
        edit := SettingsEventEditor(table, "underAttackByPlayer", 2)
        edit.Show()
        edit.enabled.Value := 1
        edit.volume.Value := "145"
        edit.color.Value := "00AAFF"
        edit.neuts.Value := 1
        edit.misses.Value := 0
        edit.smartbomb.Value := 0
        edit.Save()
        AssertFalse(app.HasOwnProp("activeDialog"), app.HasOwnProp("activeDialog") ? edit.error.Value : "The event editor must close after a valid save.")
        result := table.events["underAttackByPlayer"]
        AssertEqual(145, result["soundVolume"])
        AssertEqual("00AAFF", result["color"])
        AssertEqual(1, result["enabled"])
        AssertEqual(1, result["includeNeutralization"])
        AssertEqual(0, result["includeMisses"])
        AssertEqual(0, result["ignoreSmartbombDamage"])
        AssertDeepEqual(before, app.owner.monitoredEvents["underAttackByPlayer"])
    } finally app.Dispose()
}

TestRunner.Register("Workbench Apply saves atomically and save failure keeps the running JSON and draft", WorkbenchApplyTest)
WorkbenchApplyTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.model.Set("ThumbnailStartLocationwidth", 400)
        app.owner.failSave := true
        AssertFalse(app.Apply())
        AssertFalse(app.closed)
        AssertEqual(250, app.owner.ThumbnailStartLocation["width"])
        AssertEqual(400, app.model.Get("ThumbnailStartLocationwidth"))
        AssertEqual(0, app.restarts)
        app.owner.failSave := false
        app.owner.Profiles["Default"]["Thumbnail Positions"]["EVE - Runtime Pilot"] := Map("x", 1, "y", 2, "width", 3, "height", 4)
        AssertTrue(app.Apply())
        AssertEqual(1, app.owner.saves)
        AssertEqual(1, app.restarts)
        AssertEqual(400, app.owner.Profiles["Default"]["Thumbnails Behavior"]["ThumbnailStartLocation"]["width"])
        AssertDeepEqual(Map("x", 1, "y", 2, "width", 3, "height", 4), app.owner.Profiles["Default"]["Thumbnail Positions"]["EVE - Runtime Pilot"])
    } finally app.Dispose()
}

TestRunner.Register("Workbench focused-client preferences independently control audio and visual alerts", WorkbenchFocusedAlertsTest)
WorkbenchFocusedAlertsTest() {
    app := EventSoundRuntimeFixture()
    try {
        app.focused := true
        app.supressForFocused := true
        app.HideThumbForActiveWin := true
        app.SuppressFocusedEventSounds := false
        app.handleEventActivation("Pilot")
        AssertEqual(["alert.wav"], app.played, "Hiding a thumbnail must not mute explicitly allowed sound alerts.")
        app.endEvent("Pilot", 123)
        app.supressForFocused := false
        app.HideThumbForActiveWin := false
        app.SuppressFocusedEventSounds := true
        app.handleEventActivation("Pilot")
        AssertEqual(["alert.wav"], app.played, "Muting focused-client sounds must work independently of visibility.")
        app.HideDPSForFocused := false
        app.supressForFocused := true
        AssertFalse(app.HideDPSForFocused)
    } finally app.Dispose()
}

TestRunner.Register("Workbench close protection and Discard keep every running setting unchanged", WorkbenchCloseDiscardTest)
WorkbenchCloseDiscardTest() {
    app := SettingsWorkbenchFixture()
    before := JSON.Dump(app.owner._JSON)
    try {
        app.model.Set("ThumbnailTextSize", 24)
        app.Changed()
        app.RequestClose()
        AssertTrue(app.HasOwnProp("closePrompt"))
        app.FinishClose("keep")
        AssertFalse(app.closed)
        AssertTrue(app.model.Dirty)
        app.DiscardDraft()
        AssertFalse(app.model.Dirty)
        AssertEqual(12, app.model.Get("ThumbnailTextSize"))
        app.model.Set("ThumbnailTextSize", 26)
        app.RequestClose()
        app.FinishClose("discard")
        AssertTrue(app.closed)
        AssertEqual(before, JSON.Dump(app.owner._JSON))
        AssertEqual(0, app.owner.saves)
    } finally app.Dispose()
    unchanged := SettingsWorkbenchFixture()
    try {
        unchanged.pending["Default|AnonymousNames"] := unchanged.model.Get("AnonymousNames")
        unchanged.RequestClose()
        unchanged.FinishClose("apply")
        AssertTrue(unchanged.closed, "Closing with a harmless pending text edit must not leave the window open.")
        AssertEqual(0, unchanged.owner.saves)
        AssertEqual(0, unchanged.restarts)
    } finally unchanged.Dispose()
}

TestRunner.Register("Workbench guards tray reload exit and profile switching while a draft is open", WorkbenchLifecycleTest)
WorkbenchLifecycleTest() {
    owner := SettingsWorkbenchOwnerFixture()
    owner.Profiles["Fleet"] := SettingsModel.Clone(owner.Profiles["Default"])
    app := SettingsWorkbenchFixture(owner)
    owner.settingsWorkbench := app
    try {
        app.model.Set("ThumbnailTextSize", 24)
        app.Navigate("character-shortcuts")
        app.tableEditors[1].Edit(0)
        AssertTrue(owner.SettingsOwnsWindow(app.activeDialog.Hwnd))
        AssertFalse(owner.ActivateProfile("Fleet", "Tray"))
        AssertEqual("Default", app.model.profile)
        AssertFalse(app.Apply(), "A local row edit must finish before the outer draft can apply.")
        app.CloseChild(app.activeDialog)
        AssertFalse(owner.RequestRuntimeAction("reload"))
        AssertFalse(owner.RequestRuntimeAction("exit"))
        AssertEqual([], owner.lifecycleActions)
        AssertEqual(0, owner.saves)
        AssertFalse(owner.ActivateProfile("Fleet", "Tray"))
        AssertEqual("Default", owner.LastUsedProfile)
        AssertEqual("Fleet", app.model.profile)
        AssertEqual("Fleet", app.model.config["LastUsedProfile"])
        app.DiscardDraft()
        AssertTrue(owner.RequestRuntimeAction("reload"))
        AssertEqual(["reload"], owner.lifecycleActions)
        AssertEqual(1, owner.saves)
        owner.livePreviewsPaused := true
        AssertTrue(InStr(owner.LivePreviewStatus(), "paused"))
    } finally app.Dispose()
}

TestRunner.Register("Workbench validates related sizes and Undo restores rows without reverting later edits", WorkbenchRelationsUndoTest)
WorkbenchRelationsUndoTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.model.Set("ThumbnailStartLocationwidth", 20)
        app.Changed()
        AssertFalse(app.applyButton.Enabled)
        AssertTrue(app.undoButton.Enabled)
        app.UndoOrLocate()
        AssertEqual("layout", app.currentSection)
        app.model.Set("ThumbnailMinimumSizewidth", 10)
        app.Changed()
        AssertTrue(app.applyButton.Enabled)
        app.model.SaveRows("character-shortcuts", [["Alice", "^F1"], ["Bob", "!F2"]])
        app.Navigate("character-shortcuts")
        table := app.tableEditors[1]
        table.list.Modify(1, "Select Focus")
        table.RemoveSelected()
        app.model.SaveRows("character-shortcuts", [["Bob", "!F23"]])
        app.Undo()
        AssertDeepEqual([["Alice", "^F1"], ["Bob", "!F23"]], app.model.Rows("character-shortcuts"))
        app.Navigate("profiles")
        app.CreateProfile("Temporary")
        app.RemoveProfile()
        app.Undo()
        AssertEqual("Temporary", app.profileSelector.Text)
        AssertTrue(app.model.Profiles.Has("Temporary"))
        app.model.Profiles["Default"]["Thumbnail Positions"]["Default Pilot"] := Map("x", 10, "y", 20, "width", 100, "height", 80)
        app.model.Profiles["Temporary"]["Thumbnail Positions"]["Local Pilot"] := Map("x", 30, "y", 40, "width", 100, "height", 80)
        app.model.SetInheritance("Thumbnails Behavior", true)
        app.StageSizeReset()
        AssertEqual(100, app.model.Profiles["Default"]["Thumbnail Positions"]["Default Pilot"]["width"])
        AssertEqual(20, app.model.Profiles["Temporary"]["Thumbnail Positions"]["Local Pilot"]["width"], "Size resets target the editing profile's positions even when layout preferences are inherited.")
    } finally app.Dispose()
}

class SettingsWorkbenchSoundPlayerFixture extends RecordingEventSoundPlayer {
    Stop() => this.activePath := ""
}

WorkbenchDialogControl(dialog, kind, text := "") {
    for control in dialog.content
        if control.Type = kind && (text = "" || control.Text = text)
            return control
    throw Error("Missing dialog control: " kind " " text)
}

TestRunner.Register("Workbench library batch gain and preview workflows keep all changes in the draft", WorkbenchSoundWorkflowTest)
WorkbenchSoundWorkflowTest() {
    app := SettingsWorkbenchFixture()
    path := CreateTestSound("workbench audio")
    try {
        EventSounds.Add(app.model.config, path)
        app.Navigate("sounds")
        library := app.soundLibrary
        library.list.Modify(1, "Select Focus")
        before := JSON.Dump(app.model.Group("Monitored Events", "Game Logs Monitoring"))
        library.Batch()
        app.grids[-1].SetCells([[1, 1, 1]])
        app.CloseChild(app.activeDialog)
        AssertEqual(before, JSON.Dump(app.model.Group("Monitored Events", "Game Logs Monitoring")))
        library.Batch()
        list := WorkbenchDialogControl(app.activeDialog, "ListView")
        app.grids[-1].SetCells([[1, 1, 1]])
        first := list.GetText(1, 2)
        save := WorkbenchDialogControl(app.activeDialog, "Button", "Save to draft")
        SendMessage(0xF5, 0, 0, save)
        Sleep(30)
        AssertFalse(app.HasOwnProp("activeDialog"))
        assigned := false
        for key, name in SettingsEventTable.Names
            if name = first {
                AssertEqual(path, app.model.Group("Monitored Events", "Game Logs Monitoring")["monitoredEvents"][key]["sound"])
                assigned := true
            }
        AssertTrue(assigned)
        library.FileGain()
        WorkbenchDialogControl(app.activeDialog, "Edit").Value := 200
        save := WorkbenchDialogControl(app.activeDialog, "Button", "Save gain to draft")
        SendMessage(0xF5, 0, 0, save)
        Sleep(30)
        AssertEqual(200, EventSounds.SoundVolume(app.model.config, path))
        app.settingsSoundPlayer := SettingsWorkbenchSoundPlayerFixture(app)
        SettingsSoundLibrary.Preview(app, path, Map("soundVolume", 50))
        AssertEqual(100, app.settingsSoundPlayer.activeVolume)
        app.model.Set("MasterSoundVolume", 200)
        app.Changed()
        AssertEqual(200, app.settingsSoundPlayer.activeVolume)
        AssertEqual(100, EventSounds.MasterVolume(app.owner._JSON))
        AssertEqual([], EventSounds.Library(app.owner._JSON))
        AssertEqual(0, app.owner.saves)
    } finally {
        app.Dispose()
        FileDelete(path)
    }
}

TestRunner.Register("Workbench existing group editors support readable shortcuts and atomic member reorder", WorkbenchGroupEditorTest)
WorkbenchGroupEditorTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.model.AddGroup("characters", "Fleet")
        app.model.SaveRows("group-members", [["Alice"], ["Bob"]], "Fleet")
        app.Navigate("character-groups")
        AssertEqual(3, app.recorders.Length)
        helper := app.recorders[1]
        app.FinishShortcut(helper, "^!F23")
        AssertEqual("^!F23", app.model.Groups("characters")["Fleet"]["ForwardsHotkey"])
        AssertTrue(InStr(helper.field.Value, "F23"))
        app.RecordShortcut(helper)
        AssertTrue(HotkeyCapture.Active)
        app.Render()
        AssertFalse(HotkeyCapture.Active)
        AssertEqual("^!F23", app.model.Groups("characters")["Fleet"]["ForwardsHotkey"], "Cancelling recording during a redraw must preserve the old assignment.")
        table := app.tableEditors[1]
        table.list.Modify(2, "Select Focus")
        table.Move(-1)
        AssertDeepEqual([["Bob"], ["Alice"]], app.model.Rows("group-members", "Fleet"))
        AssertThrows(() => app.model.AddGroup("characters", "FLEET"), "already exists")
        AssertTrue(SettingsSchema.Search("smartbomb").Length > 0)
        app.width := 700
        app.Layout()
        app.Render()
        AssertTrue(app.recorders[1].extraHeight > 0)
    } finally app.Dispose()
}
