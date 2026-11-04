TestRunner.Register("Settings followup sidebar uses pane scrollbar dragging without changing navigation selection", SettingsFollowupTreeDragTest)
SettingsFollowupTreeDragTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.model.Set("SettingsTheme", "dark")
        app.Changed()
        app.S_Gui.Show("NoActivate x-3000 y-3000 w800 h500")
        app.width := 800, app.height := 500
        for root in app.navRoots
            app.nav.Modify(root, "Expand")
        app.Layout()
        app.theme.Repaint()
        source := app.theme.chrome.windows[app.nav.Hwnd]
        AssertTrue(source.bars.Has(1), "Navigation needs the same separate scrollbar as the right pane.")
        handle := source.bars[1].handle
        AssertEqual("PaneScrollbar", app.theme.chrome.windows[handle].kind)
        selected := app.nav.GetSelection()
        app.theme.chrome.scrollbars.Scroll(source, 1, 5, 0)
        first := SendMessage(0x110A, 5, 0, app.nav)
        point := SettingsChromeStartDrag(app, handle, 1)
        point[2] += Round(app.theme.chrome.scrollbars.drag.travel * 0.8)
        SettingsChromeMoveDrag(handle, point)
        AssertNotEqual(first, SendMessage(0x110A, 5, 0, app.nav), "Dragging must move the tree's first visible row.")
        AssertEqual(selected, app.nav.GetSelection(), "Scrolling must not activate another settings page.")
        AssertEqual(SettingsTheme.BGR("25282D"), SettingsChromeTrackPixel(app, handle, 1))
        SettingsChromeMoveDrag(handle, point, true)
        AssertFalse(IsObject(app.theme.chrome.scrollbars.drag))
        app.width := 960, app.height := 620
        app.Layout()
        AssertTrue(DllCall("IsWindowVisible", "Ptr", handle))
        app.theme.SetMode("light")
        AssertFalse(DllCall("IsWindowVisible", "Ptr", handle), "Light mode restores the native scrollbar.")
    } finally app.Dispose()
}

TestRunner.Register("Settings followup table text and list scrollbars share pane controls and move their content", SettingsFollowupControlBarsTest)
SettingsFollowupControlBarsTest() {
    app := SettingsWorkbenchFixture()
    dialog := 0
    try {
        app.model.Set("SettingsTheme", "dark")
        app.Changed()
        app.model.AddGroup("characters", "Fleet")
        rows := []
        loop 80
            rows.Push(["Pilot " A_Index])
        app.model.SaveRows("group-members", rows, "Fleet")
        app.Navigate("character-groups")
        app.S_Gui.Show("NoActivate x-3000 y-3000 w1000 h740")
        grid := app.grids[1], list := grid.list
        list.ModifyCol(1, 1600)
        app.theme.Repaint()
        grid.Select(1, 1)
        grid.BeginEdit()
        grid.editor.Value := "Edited pilot"
        source := app.theme.chrome.windows[list.Hwnd]
        for axis in [1, 0] {
            AssertTrue(source.bars.Has(axis), "Tables need separate bars on both axes.")
            handle := source.bars[axis].handle
            point := SettingsChromeStartDrag(app, handle, axis)
            AssertFalse(IsObject(grid.editor), "Dragging the visible scrollbar must commit its active cell.")
            point[axis = 1 ? 2 : 1] += Round(app.theme.chrome.scrollbars.drag.travel * 0.7)
            SettingsChromeMoveDrag(handle, point)
            AssertTrue(DllCall("GetScrollPos", "Ptr", list.Hwnd, "Int", axis, "Int") > 0)
            AssertEqual(SettingsTheme.BGR("25282D"), SettingsChromeTrackPixel(app, handle, axis))
            SettingsChromeMoveDrag(handle, point, true)
        }
        AssertEqual("Edited pilot", app.model.Rows("group-members", "Fleet")[1][1])
        dialog := app.ChildDialog("Scrollbar controls")
        text := "", choices := []
        loop 100 {
            text .= (A_Index = 1 ? "" : "`n") "Line " A_Index " with enough text to scroll."
            choices.Push("Choice " A_Index)
        }
        edit := dialog.AddEdit("x20 y20 w280 h130 Multi -Wrap", text)
        listbox := dialog.AddListBox("x320 y20 w280 h130", choices)
        dialog.Show("w640 h320")
        app.theme.Repaint()
        for control in [edit, listbox] {
            item := app.theme.chrome.windows[control.Hwnd]
            AssertTrue(item.bars.Has(1))
            handle := item.bars[1].handle
            point := SettingsChromeStartDrag(app, handle, 1)
            point[2] += Round(app.theme.chrome.scrollbars.drag.travel * 0.7)
            SettingsChromeMoveDrag(handle, point)
            AssertTrue(DllCall("GetScrollPos", "Ptr", control.Hwnd, "Int", 1, "Int") > 0, control.Type " content must scroll.")
            SettingsChromeMoveDrag(handle, point, true)
        }
        app.CloseChild(dialog)
        dialog := 0
        AssertFalse(DllCall("IsWindow", "Ptr", handle), "Owned scrollbars must be destroyed with their controls.")
    } finally app.Dispose()
}

TestRunner.Register("Settings followup dropdowns use pane scrollbars and retain popup capture after dragging", SettingsFollowupPopupBarsTest)
SettingsFollowupPopupBarsTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.model.Set("SettingsTheme", "dark")
        app.Changed()
        app.S_Gui.Show("NoActivate x-3000 y-3000 w1000 h740")
        control := app.Binding("ThumbnailTextFont").control
        SendMessage(0x14F, 1, 0, control)
        info := app.theme.chrome.ComboInfo(control.Hwnd)
        popup := NumGet(info, A_PtrSize = 8 ? 56 : 48, "Ptr")
        app.theme.Repaint()
        source := app.theme.chrome.windows[popup]
        AssertTrue(source.bars.Has(1), "An open dropdown needs a separate scrollbar.")
        handle := source.bars[1].handle
        before := DllCall("GetCapture", "Ptr")
        point := SettingsChromeStartDrag(app, handle, 1)
        point[2] += Round(app.theme.chrome.scrollbars.drag.travel * 0.6)
        SettingsChromeMoveDrag(popup, point)
        AssertTrue(DllCall("GetScrollPos", "Ptr", popup, "Int", 1, "Int") > 0, "Dragging must move the popup content.")
        SettingsChromeMoveDrag(popup, point, true)
        AssertTrue(SendMessage(0x157, 0, 0, control), "The popup must stay open after dragging.")
        AssertEqual(before, DllCall("GetCapture", "Ptr"))
        SendMessage(0x14F, 0, 0, control)
        AssertFalse(DllCall("IsWindowVisible", "Ptr", handle), "Closing the dropdown must hide its scrollbar.")
    } finally app.Dispose()
}

TestRunner.Register("Settings followup large conflict reports group equivalent keys and fit a scrollable window", SettingsFollowupConflictsTest)
SettingsFollowupConflictsTest() {
    entries := []
    HotkeyBindings.Add(entries, "one", "One", "^!F23")
    HotkeyBindings.Add(entries, "two", "Two", "!^f23")
    HotkeyBindings.Add(entries, "three", "Three", "F24")
    HotkeyBindings.Add(entries, "four", "Four", "F24")
    HotkeyBindings.Add(entries, "five", "No conflict", "F22")
    groups := HotkeyBindings.ConflictGroups(entries)
    AssertEqual(2, groups.Length)
    for group in groups
        AssertEqual(2, group.entries.Length)
    app := SettingsWorkbenchFixture()
    try {
        rows := []
        loop 90
            rows.Push(["Pilot " A_Index, "1"])
        app.model.SaveRows("character-shortcuts", rows)
        app.ShowShortcutConflicts()
        dialog := app.activeDialog
        AssertEqual(1, dialog.conflictGroups.Length)
        AssertEqual(90, dialog.conflictGroups[1].entries.Length)
        AssertTrue(dialog.scroll.maximum > 0)
        work := SettingsWindowSizing.WorkArea(app.S_Gui.Hwnd)
        WinGetPos(&x, &y, &width, &height, dialog.Hwnd)
        AssertTrue(x >= work.left && y >= work.top)
        AssertTrue(x + width <= work.right && y + height <= work.bottom)
        dialog.scroll.ScrollTo(dialog.scroll.maximum)
        AssertTrue(dialog.scroll.offset > 0)
        app.CloseChild(dialog)
        AssertTrue(DllCall("IsWindowEnabled", "Ptr", app.S_Gui.Hwnd))
    } finally app.Dispose()
}

TestRunner.Register("Settings followup monitored character switches share values and inheritance with log sources", SettingsFollowupMonitoringTest)
SettingsFollowupMonitoringTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.Navigate("monitored-characters")
        for id in ["gameLogsMonitoringEnabled", "monitorOnlySelectedChars"] {
            binding := app.Binding(id)
            AssertTrue(binding.control.Enabled)
            binding.control.Value := 1
            app.BooleanChanged(binding, binding.control)
        }
        AssertTrue(InStr(app.monitoringNotice.Value, "Only the listed"), "Enabled monitoring status: " app.monitoringNotice.Value)
        app.Navigate("sources")
        AssertEqual(1, app.Binding("gameLogsMonitoringEnabled").control.Value)
        AssertEqual(1, app.Binding("monitorOnlySelectedChars").control.Value)
        app.CreateProfile("Fleet")
        app.SetSectionInheritance("sources", true)
        app.Navigate("monitored-characters")
        AssertFalse(app.Binding("monitorOnlySelectedChars").control.Enabled)
        app.SetSectionInheritance("sources", false)
        app.Render()
        AssertTrue(app.Binding("monitorOnlySelectedChars").control.Enabled, "Custom source switch must become editable.")
        app.Binding("monitorOnlySelectedChars").control.Value := 0
        app.BooleanChanged(app.Binding("monitorOnlySelectedChars"), app.Binding("monitorOnlySelectedChars").control)
        AssertTrue(InStr(app.monitoringNotice.Value, "All detected"), "Disabled filter status: " app.monitoringNotice.Value)
        AssertEqual(1, app.model.Profiles["Default"]["Game Logs Monitoring"]["monitorOnlySelectedChars"])
        AssertEqual(0, app.owner.saves)
    } finally app.Dispose()
}

TestRunner.Register("Settings followup Apply captures and restores the selected page window and scroll once", SettingsFollowupResumeTest)
SettingsFollowupResumeTest() {
    app := SettingsWorkbenchFixture(), restored := 0
    previous := EnvGet(SettingsWorkbench.ResumeVariable)
    try {
        app.width := 850, app.height := 560
        app.S_Gui.Show("NoActivate x-3000 y-3000 w850 h560")
        app.Navigate("sources")
        app.settingsScroller.ScrollTo(100)
        offset := app.settingsScroller.offset
        app.model.Set("monitoringInterval", 500)
        AssertTrue(app.Apply())
        AssertTrue(app.closed)
        AssertEqual(1, app.restarts)
        AssertEqual("sources", app.resumeState["section"])
        restored := SettingsWorkbenchFixture(app.owner)
        restored.RestoreSession(app.resumeState, false)
        AssertEqual("sources", restored.currentSection)
        AssertEqual(offset, restored.settingsScroller.offset)
        AssertFalse(restored.model.Dirty)
        EnvSet(SettingsWorkbench.ResumeVariable, JSON.Dump(app.resumeState))
        AssertEqual("sources", SettingsWorkbench.TakeResume()["section"])
        AssertEqual(0, SettingsWorkbench.TakeResume())
        EnvSet(SettingsWorkbench.ResumeVariable, "invalid")
        AssertEqual(0, SettingsWorkbench.TakeResume())
    } finally {
        EnvSet(SettingsWorkbench.ResumeVariable, previous)
        if IsObject(restored)
            restored.Dispose()
        app.Dispose()
    }
}
