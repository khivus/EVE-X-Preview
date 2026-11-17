; Observe real native paint messages, not only the final WM_PRINT snapshot.
class SettingsPaintObserver {
    __New(handles) {
        this.counts := Map(), this.callback := CallbackCreate(ObjBindMethod(this, "WindowProc"),, 6)
        for handle in handles {
            this.counts[handle] := {paints: 0, erases: 0, updates: []}
            DllCall("comctl32\SetWindowSubclass", "Ptr", handle, "Ptr", this.callback, "UPtr", 91, "UPtr", 0)
        }
    }
    Reset() {
        for handle, count in this.counts
            count.paints := 0, count.erases := 0, count.updates := []
    }
    WindowProc(handle, message, wParam, lParam, subclass, data) {
        if this.counts.Has(handle) {
            if message = 0xF {
                this.counts[handle].paints++
                rect := Buffer(16, 0)
                if DllCall("GetUpdateRect", "Ptr", handle, "Ptr", rect, "Int", false)
                    this.counts[handle].updates.Push([NumGet(rect, 0, "Int"), NumGet(rect, 4, "Int"), NumGet(rect, 8, "Int"), NumGet(rect, 12, "Int")])
            } else if message = 0x14
                this.counts[handle].erases++
        }
        return DllCall("comctl32\DefSubclassProc", "Ptr", handle, "UInt", message, "Ptr", wParam, "Ptr", lParam, "Ptr")
    }
    Dispose() {
        for handle in this.counts
            if DllCall("IsWindow", "Ptr", handle)
                DllCall("comctl32\RemoveWindowSubclass", "Ptr", handle, "Ptr", this.callback, "UPtr", 91)
        CallbackFree(this.callback)
    }
}

TestRunner.Register("Settings scroll rendering selection redraws cells without erasing or repainting the table header", SettingsTableSelectionPaintingTest)
SettingsTableSelectionPaintingTest() {
    app := SettingsWorkbenchFixture(), observer := 0
    try {
        app.model.Set("SettingsTheme", "dark")
        app.Changed()
        app.model.SaveRows("character-shortcuts", [["Alpha", "F13"], ["Beta", "F14"], ["Gamma", "F15"]])
        app.Navigate("character-shortcuts")
        app.S_Gui.Show("NoActivate x-3000 y-3000 w1000 h740")
        Sleep(220)
        grid := app.grids[1], list := grid.list
        header := SendMessage(0x101F, 0, 0, list)
        DllCall("RedrawWindow", "Ptr", list.Hwnd, "Ptr", 0, "Ptr", 0, "UInt", 0x585)
        observer := SettingsPaintObserver([list.Hwnd, header])
        loop 30
            grid.Select(Mod(A_Index - 1, 3) + 1, 1)
        AssertTrue(observer.counts[list.Hwnd].paints > 0)
        AssertEqual(0, observer.counts[header].paints, "Changing selected cells must not repaint the header.")
        AssertEqual(0, observer.counts[list.Hwnd].erases, "Selection must not clear the table between frames.")
        headerRect := Buffer(16, 0)
        DllCall("GetWindowRect", "Ptr", header, "Ptr", headerRect)
        DllCall("MapWindowPoints", "Ptr", 0, "Ptr", list.Hwnd, "Ptr", headerRect, "UInt", 2)
        for update in observer.counts[list.Hwnd].updates
            AssertTrue(update[2] >= NumGet(headerRect, 12, "Int"), "Cell selection must not invalidate the parent's pixels under the header.")
        AssertTrue(SendMessage(0x1037, 0, 0, list) & 0x10000, "Use the table's own double buffer.")
        AssertEqual("Alpha", list.GetText(1, 1))
        AssertEqual("F14", list.GetText(2, 2))
        AssertTrue(grid.selected.Has("3:1"))
    } finally {
        if IsObject(observer)
            observer.Dispose()
        app.Dispose()
    }
}

TestRunner.Register("Settings scroll rendering repeated cell clicks preserve the header through editor focus and palette changes", SettingsHeaderClickLifecycleTest)
SettingsHeaderClickLifecycleTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.model.SaveRows("character-shortcuts", [["Alpha", "F13"], ["Beta", "F14"], ["Gamma", "F15"]])
        app.Navigate("character-shortcuts")
        app.S_Gui.Show("NoActivate x-3000 y-3000 w1000 h740")
        Sleep(220)
        for mode in ["dark", "light", "dark"] {
            app.model.Set("SettingsTheme", mode)
            app.Changed()
            grid := app.grids[1], list := grid.list
            header := SendMessage(0x101F, 0, 0, list)
            before := Buffer(16, 0)
            DllCall("GetWindowRect", "Ptr", header, "Ptr", before)
            headerHeight := NumGet(before, 12, "Int") - NumGet(before, 4, "Int")
            loop 6 {
                row := Mod(A_Index - 1, 3) + 1
                GridPostPointer(grid, row, 1, 1, 0x201)
                GridPostPointer(grid, row, 1, 0, 0x202)
                AssertTrue(IsObject(grid.editor))
                AssertTrue(WinGetStyle(list.Hwnd) & 0x02000000, "Native body/focus painting must clip out the child header.")
                AssertTrue(SettingsNativeChrome.Long(header, -16) & 0x10000000)
                after := Buffer(16, 0)
                DllCall("GetWindowRect", "Ptr", header, "Ptr", after)
                AssertEqual(headerHeight, NumGet(after, 12, "Int") - NumGet(after, 4, "Int"))
                AssertTrue(NumGet(after, 8, "Int") > NumGet(after, 0, "Int"))
                DllCall("MapWindowPoints", "Ptr", 0, "Ptr", list.Hwnd, "Ptr", after, "UInt", 2)
                AssertTrue(NumGet(after, 4, "Int") >= 0 && NumGet(after, 4, "Int") <= 2, "Header must remain at the top of the table's client area.")
                AssertEqual(SettingsTheme.BGR(mode = "dark" ? "343941" : "EFEFEF"), SettingsPaintPixel({Hwnd: header}, 2, 2))
            }
            grid.CommitEditor()
            AssertEqual("Alpha", list.GetText(1, 1))
            AssertEqual("Beta", list.GetText(2, 1))
        }
    } finally app.Dispose()
}

TestRunner.Register("Settings scroll rendering tables and headers become idle after repaint and scrolling", SettingsTablePaintIdleTest)
SettingsTablePaintIdleTest() {
    app := SettingsWorkbenchFixture(), observer := 0
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
        Sleep(220)
        list := app.grids[1].list, header := SendMessage(0x101F, 0, 0, list)
        observer := SettingsPaintObserver([list.Hwnd, header])
        SendMessage(0x1014, 0, 320, list)
        DllCall("RedrawWindow", "Ptr", list.Hwnd, "Ptr", 0, "Ptr", 0, "UInt", 0x585)
        Sleep(80)
        observer.Reset()
        Sleep(150)
        AssertEqual(0, observer.counts[header].paints, "Header painting must stop when input stops.")
        AssertEqual(0, observer.counts[list.Hwnd].paints, "Table painting must stop when input stops.")
        AssertTrue(SendMessage(0x1027, 0, 0, list) > 0)
    } finally {
        if IsObject(observer)
            observer.Dispose()
        app.Dispose()
    }
}

TestRunner.Register("Settings scroll rendering moving the pane repaints separators and preserves control geometry", SettingsPaneChildPaintingTest)
SettingsPaneChildPaintingTest() {
    app := SettingsWorkbenchFixture(), observer := 0
    try {
        app.model.Set("SettingsTheme", "dark")
        app.Changed()
        app.Navigate("appearance")
        app.S_Gui.Show("NoActivate x-3000 y-3000 w1000 h740")
        Sleep(220)
        pane := app.settingsScroller
        pane.ScrollTo(0)
        for control in app.MainFrame {
            control.GetPos(, &y,, &height)
            if control.Type = "Text" && control.HasOwnProp("ThemeBackground") && control.ThemeBackground = "DDE1E7"
                && y > 180 && y < pane.page - 100 {
                separator := control
                break
            }
        }
        AssertTrue(IsSet(separator), "Fixture must have a visible separator.")
        binding := app.Binding("ThumbnailTextFont")
        binding.control.GetPos(&xBefore, &yBefore, &widthBefore, &heightBefore)
        observer := SettingsPaintObserver([separator.Hwnd])
        for offset in [12, 28, 44, 60, 28, 0]
            pane.ScrollTo(offset)
        AssertTrue(observer.counts[separator.Hwnd].paints > 0, "Moving a clipped canvas must repaint thin separators, not copy cached child pixels.")
        binding.control.GetPos(&xAfter, &yAfter, &widthAfter, &heightAfter)
        AssertEqual([xBefore, yBefore, widthBefore, heightBefore], [xAfter, yAfter, widthAfter, heightAfter])
        AssertEqual(SettingsTheme.BGR(app.theme.line), SettingsPaintPixel(separator, 0, 0))
    } finally {
        if IsObject(observer)
            observer.Dispose()
        app.Dispose()
    }
}

TestRunner.Register("Settings scroll rendering held table thumb drag preserves selected cells and stops repainting", SettingsSelectedTableDragPaintingTest)
SettingsSelectedTableDragPaintingTest() {
    app := SettingsWorkbenchFixture(), observer := 0
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
        Sleep(220)
        grid := app.grids[1], list := grid.list
        grid.Select(2, 1)
        grid.Select(4, 1, true)
        selected := JSON.Dump(grid.selected)
        header := SendMessage(0x101F, 0, 0, list)
        observer := SettingsPaintObserver([header])
        point := SettingsChromeStartDrag(app, list.Hwnd, 1)
        point[2] += Round(app.theme.chrome.scrollbars.drag.travel * 0.65)
        SettingsChromeMoveDrag(list.Hwnd, point)
        AssertEqual(selected, JSON.Dump(grid.selected))
        AssertEqual(SettingsTheme.BGR("25282D"), SettingsChromeTrackPixel(app, list.Hwnd, 1))
        SettingsChromeMoveDrag(list.Hwnd, point, true)
        Sleep(80)
        observer.Reset()
        Sleep(150)
        AssertEqual(0, observer.counts[header].paints)
        AssertEqual("Pilot 2", list.GetText(2, 1))
    } finally {
        if IsObject(observer)
            observer.Dispose()
        app.Dispose()
    }
}
