TestRunner.Register("Settings appearance color cells draw the configured RGB swatch in both themes", SettingsSwatchDrawingTest)
SettingsSwatchDrawingTest() {
    app := SettingsWorkbenchFixture()
    dc := 0, memory := 0, bitmap := 0, previous := 0
    try {
        app.model.SaveRows("character-colors", [["Alice", "12AB34", "FFFFFF", "000000"]])
        app.Navigate("character-appearance")
        grid := app.tableEditors[1].grid
        dc := DllCall("GetDC", "Ptr", grid.list.Hwnd, "Ptr")
        memory := DllCall("CreateCompatibleDC", "Ptr", dc, "Ptr")
        bitmap := DllCall("CreateCompatibleBitmap", "Ptr", dc, "Int", 1200, "Int", 400, "Ptr")
        previous := DllCall("SelectObject", "Ptr", memory, "Ptr", bitmap, "Ptr")
        for mode in ["light", "dark"] {
            app.model.Set("SettingsTheme", mode)
            app.Changed()
            draw := Buffer(160, 0)
            NumPut("UInt", 0x30001, draw, A_PtrSize * 3)
            NumPut("Ptr", memory, draw, A_PtrSize * 4)
            NumPut("UPtr", 0, draw, A_PtrSize = 8 ? 56 : 36)
            NumPut("Int", 1, draw, A_PtrSize = 8 ? 88 : 56)
            AssertEqual(4, grid.Draw(grid.list, draw))
            rect := Buffer(16, 0)
            NumPut("Int", 0, "Int", 1, rect)
            SendMessage(0x1038, 0, rect.Ptr, grid.list)
            pixel := DllCall("GetPixel", "Ptr", memory, "Int", NumGet(rect, 0, "Int") + 10, "Int", NumGet(rect, 4, "Int") + 7, "UInt")
            AssertEqual(SettingsTheme.BGR("12AB34"), pixel, "The swatch must show the actual RGB color in " mode ".")
        }
        AssertEqual("", grid.ColorPreview("invalid color"))
        AssertEqual("12AB34", app.model.Rows("character-colors")[1][2])
    } finally {
        if previous
            DllCall("SelectObject", "Ptr", memory, "Ptr", previous)
        if bitmap
            DllCall("DeleteObject", "Ptr", bitmap)
        if memory
            DllCall("DeleteDC", "Ptr", memory)
        if dc
            DllCall("ReleaseDC", "Ptr", app.tableEditors[1].list.Hwnd, "Ptr", dc)
        app.Dispose()
    }
}

TestRunner.Register("Settings appearance border preview opens with legacy color formats without changing saved values", SettingsBorderPreviewColorsTest)
SettingsBorderPreviewColorsTest() {
    for pair in [["#E36A0D", "E36A0D"], ["0xe36a0d", "E36A0D"], ["e36a0d", "E36A0D"],
        ["rgb(12, 34, 56)", "0C2238"], ["12, 34, 56", "0C2238"], ["red", "FF0000"],
        ["  #12ab34  ", "12AB34"], ["0", "000000"], ["invalid color", "E36A0D"], ["", "E36A0D"]] {
        owner := SettingsWorkbenchOwnerFixture()
        owner._JSON["_Profiles"]["Default"]["Thumbnails Visuals"]["ClientHighligtColor"] := pair[1]
        before := JSON.Dump(owner._JSON)
        app := SettingsWorkbenchFixture(owner)
        try {
            AssertFalse(app.model.Dirty, "Opening Settings must not normalize the stored color.")
            for mode in [1, 2, 3] {
                app.model.Set("EventActiveBorderMode", mode)
                for theme in ["light", "dark"] {
                    app.theme.SetMode(theme)
                    app.Navigate("appearance")
                    app.Render()
                    rings := Map()
                    for control in app.MainFrame {
                        if control.Type != "Text" || !control.ThemePreserve
                            continue
                        control.GetPos(,, &width, &height)
                        if width = 210 && height = 76
                            rings["outer"] := control.ThemeBackground
                        if width = 194 && height = 60
                            rings["inner"] := control.ThemeBackground
                    }
                    AssertEqual(2, rings.Count, "Both preview rings must render for " pair[1] ".")
                    AssertEqual(mode = 2 ? "E9A23B" : pair[2], rings["outer"])
                    AssertEqual(mode = 1 ? "E9A23B" : pair[2], rings["inner"])
                }
            }
            AssertEqual(pair[1], app.model.Get("ClientHighligtColor"))
            AssertEqual(before, JSON.Dump(owner._JSON))
            AssertEqual(0, owner.saves)
        } finally app.Dispose()
    }
}

TestRunner.Register("Settings appearance dark theme previews across pages and child editors and Discard restores it", SettingsDarkThemeTest)
SettingsDarkThemeTest() {
    owner := SettingsWorkbenchOwnerFixture()
    owner._JSON["SettingsTheme"] := "light"
    app := SettingsWorkbenchFixture(owner)
    try {
        app.model.Set("SettingsTheme", "dark")
        app.Changed()
        AssertTrue(app.theme.dark)
        AssertEqual("202124", app.S_Gui.BackColor)
        for section in SettingsSchema.Sections
            app.Navigate(section.id)
        app.Navigate("dps")
        incoming := app.Binding("incomingDPSEnabled").control
        before := app.model.Get("incomingDPSEnabled")
        SendMessage(0xF5, 0, 0, incoming)
        Sleep(35)
        AssertEqual(!before, app.model.Get("incomingDPSEnabled"), "Dark checkboxes must retain native click behavior.")
        app.Navigate("events")
        editor := SettingsEventEditor(app.eventTable, "underAttackByPlayer", 2)
        editor.Show()
        AssertEqual("202124", app.activeDialog.root.BackColor)
        app.CloseChild(app.activeDialog)
        AssertEqual("light", owner._JSON["SettingsTheme"], "Theme previews must leave running configuration unchanged.")
        app.DiscardDraft()
        AssertFalse(app.theme.dark)
        AssertEqual("F7F7F7", app.S_Gui.BackColor)
        AssertEqual(0, owner.saves)
    } finally app.Dispose()
}

class SettingsStartupUpdatesFixture extends SettingsWorkbenchFixture {
    __New() => SettingsWorkbench.Prototype.__New.Call(this, SettingsWorkbenchOwnerFixture(), true)
    CreateUpdateChecker() => UpdateCheckerFixture()
    Show() {
        this.S_Gui.Show("NoActivate x-3000 y-3000 w1000 h740")
        this.Layout()
    }
}
TestRunner.Register("Settings appearance opening settings starts an asynchronous update check and shows a stable-update notice", SettingsStartupUpdatesTest)
SettingsStartupUpdatesTest() {
    app := SettingsStartupUpdatesFixture()
    try {
        AssertTrue(app.updateChecker.busy)
        AssertEqual(2, app.updateChecker.requests.Length)
        app.programVersion := "1.0"
        for request in app.updateChecker.requests
            request.done := true
        app.updateChecker.Poll()
        app.UpdateVersionControls()
        AssertTrue(app.updateNotice.Visible)
        AssertTrue(InStr(app.updateNotice.Text, "1.6.0.14"))
        AssertFalse(app.model.Dirty, "An update check must not alter the settings draft.")
        SendMessage(0xF5, 0, 0, app.updateNotice)
        Sleep(35)
        AssertEqual("about", app.currentSection)
        app.StartUpdateCheck()
        AssertEqual(4, app.updateChecker.requests.Length)
    } finally app.Dispose()
    AssertFalse(app.updateChecker.busy, "Closing settings must stop its background check.")
}

TestRunner.Register("Settings appearance update notices exclude current older and pre-release-only versions", SettingsUpdateNoticeTest)
SettingsUpdateNoticeTest() {
    app := SettingsWorkbenchFixture()
    try {
        for tag in ["", "1.0", app.programVersion] {
            app.updateChecker.releaseTag := tag
            app.updateChecker.preReleaseTag := "99.0"
            app.UpdateVersionControls()
            AssertFalse(app.updateNotice.Visible)
        }
        AssertFalse(app.HasOwnProp("activeDialog"))
    } finally app.Dispose()
}

TestRunner.Register("Settings appearance inherited scalar settings can edit Default in place without activating it", SettingsDefaultScalarTest)
SettingsDefaultScalarTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.CreateProfile("Fleet")
        app.StageProfileActivation()
        running := app.model.config["LastUsedProfile"]
        app.model.Set("ThumbnailTextSize", 17)
        app.SetSectionInheritance("appearance", true)
        app.Navigate("appearance")
        app.EditDefaultHere("appearance")
        AssertEqual("Fleet", app.profileSelector.Text)
        AssertEqual("Fleet", app.SelectedProfile)
        AssertEqual("Default", app.model.profile)
        AssertTrue(InStr(app.settingsScroller.heading.Text, "editing DEFAULT"))
        binding := app.Binding("ThumbnailTextSize")
        AssertTrue(binding.control.Enabled)
        binding.control.Value := 24
        app.CommitField(binding, binding.control)
        AssertEqual(24, app.model.Profiles["Default"]["Thumbnails Visuals"]["ThumbnailTextSize"])
        AssertEqual(17, app.model.Profiles["Fleet"]["Thumbnails Visuals"]["ThumbnailTextSize"])
        app.EndDefaultEditing()
        AssertEqual("Fleet", app.model.profile)
        AssertEqual(24, app.model.Get("ThumbnailTextSize"))
        AssertEqual(running, app.model.config["LastUsedProfile"])
        AssertEqual(0, app.owner.saves)
        AssertEqual(0, app.restarts)
        AssertTrue(app.Apply())
        AssertEqual(running, app.owner.LastUsedProfile)
        AssertEqual(1, app.restarts)
    } finally app.Dispose()
}

TestRunner.Register("Settings appearance Default editing routes inherited tables groups and event editors to Default", SettingsDefaultTablesTest)
SettingsDefaultTablesTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.model.AddGroup("characters", "Fleet group")
        app.CreateProfile("Fleet")
        for group in ["character-shortcuts", "character-groups", "events"]
            app.SetSectionInheritance(group, true)
        app.Navigate("character-shortcuts")
        app.EditDefaultHere("character-shortcuts")
        AssertTrue(app.tableEditors[1].grid.SetCells([[1, 1, "Alice"], [1, 2, "F23"]]))
        app.tableEditors[1].grid.Select(1, 1)
        app.tableEditors[1].grid.RemoveRows()
        app.Undo()
        AssertEqual("Fleet", app.profileSelector.Text, "Undo must keep the selected profile during Default editing.")
        AssertEqual(1, app.tableEditors[1].grid.rows.Length)
        app.Navigate("character-groups")
        AssertEqual("Fleet", app.model.profile, "Leaving the subsection must end Default editing.")
        app.EditDefaultHere("character-groups")
        AssertTrue(app.tableEditors[1].grid.SetCells([[1, 1, "Alice"]]))
        app.SetGroupShortcut("characters", "Fleet group", "ForwardsHotkey", "F24")
        app.RemoveGroup("characters", "Fleet group")
        app.Undo()
        AssertEqual("Fleet", app.profileSelector.Text)
        app.Navigate("events")
        app.EditDefaultHere("events")
        AssertTrue(app.eventTable.grid.SetCells([[2, 4, "12AB34"]]))
        app.EndDefaultEditing()
        AssertEqual("12AB34", app.eventTable.grid.rows[2][4])
        AssertEqual("Fleet", app.profileSelector.Text)
        AssertEqual([], app.model.Profiles["Fleet"]["Hotkeys Settings"]["CharacterHotkeys"])
        AssertEqual(1, app.model.Profiles["Default"]["Hotkeys Settings"]["CharacterHotkeys"].Length)
        AssertEqual("F24", app.model.Profiles["Default"]["Hotkey Groups"]["Fleet group"]["ForwardsHotkey"])
        AssertEqual("", app.model.Profiles["Fleet"]["Hotkey Groups"]["Fleet group"]["ForwardsHotkey"])
    } finally app.Dispose()
}

TestRunner.Register("Settings appearance invalid Default drafts return to the same view and Discard clears edit mode", SettingsDefaultValidationTest)
SettingsDefaultValidationTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.CreateProfile("Fleet")
        app.SetSectionInheritance("layout", true)
        app.Navigate("layout")
        app.EditDefaultHere("layout")
        binding := app.Binding("ThumbnailStartLocationwidth")
        binding.control.Value := "invalid"
        app.TextChanged(binding, binding.control)
        app.EndDefaultEditing()
        AssertFalse(app.Apply())
        AssertEqual("Fleet", app.profileSelector.Text)
        AssertTrue(app.HasOwnProp("defaultEditContext"))
        AssertEqual("invalid", app.Binding("ThumbnailStartLocationwidth").control.Value)
        app.DiscardDraft()
        AssertFalse(app.HasOwnProp("defaultEditContext"))
        AssertEqual(0, app.errors.Count)
        AssertEqual(app.model.profile, app.profileSelector.Text)
    } finally app.Dispose()
}

TestRunner.Register("Settings appearance duplicated Quick Group gestures edit Default Mouse actions without changing local Quick Group settings", SettingsDefaultQuickGestureTest)
SettingsDefaultQuickGestureTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.CreateProfile("Fleet")
        app.model.Set("QuickGroupColor", "123456")
        app.SetSectionInheritance("mouse-actions", true)
        app.Navigate("quick-group")
        app.EditDefaultHere("mouse-actions")
        AssertEqual("mouse-actions", app.currentSection)
        AssertEqual("Default", app.model.profile)
        app.EndDefaultEditing()
        app.Navigate("quick-group")
        AssertEqual("123456", app.model.Get("QuickGroupColor"))
        AssertEqual("Fleet", app.profileSelector.Text)
    } finally app.Dispose()
}
