class SettingsColorChoiceFixture extends SettingsWorkbenchFixture {
    colorChoice := ""
    ChooseColor(value) => this.colorChoice
}

TestRunner.Register("Settings input pulse scalar color swatches update from typing and the picker in both palettes", SettingsLiveSwatchTest)
SettingsLiveSwatchTest() {
    app := SettingsColorChoiceFixture()
    try {
        for mode in ["dark", "light"] {
            app.model.Set("SettingsTheme", mode)
            app.Changed()
            app.Navigate("appearance")
            binding := app.Binding("ThumbnailTextColor")
            binding.control.Value := "34AF62"
            app.TextChanged(binding, binding.control)
            AssertEqual(SettingsTheme.BGR("34AF62"), SettingsPaintPixel(binding.swatch, 4, 4))
            binding.control.Value := "bad color"
            app.TextChanged(binding, binding.control)
            AssertEqual(SettingsTheme.BGR("34AF62"), SettingsPaintPixel(binding.swatch, 4, 4), "Incomplete text keeps the last valid preview.")
            app.colorChoice := "E24C89"
            app.PickColor(binding, binding.control)
            AssertEqual(SettingsTheme.BGR("E24C89"), SettingsPaintPixel(binding.swatch, 4, 4))
            AssertFalse(app.errors.Has(binding.key))
        }
    } finally app.Dispose()
}

TestRunner.Register("Settings input pulse numeric arrows retain plain four-digit and negative integers", SettingsPlainSpinnerTest)
SettingsPlainSpinnerTest() {
    app := SettingsWorkbenchFixture()
    try {
        for value in [2400, -2400] {
            app.Navigate("layout")
            binding := app.Binding("ThumbnailStartLocationx")
            binding.control.Value := value
            app.CommitField(binding, binding.control)
            for control in binding.inputs
                if control.Type = "UpDown"
                    spin := control
            AssertTrue(WinGetStyle(spin.Hwnd) & 0x80)
            rect := Buffer(16, 0)
            DllCall("GetClientRect", "Ptr", spin.Hwnd, "Ptr", rect)
            point := (NumGet(rect, 8, "Int") // 2) | ((NumGet(rect, 12, "Int") // 4) << 16)
            PostMessage(0x201, 1, point, spin)
            PostMessage(0x202, 0, point, spin)
            Sleep(35)
            AssertEqual(value + 1, app.model.Get("ThumbnailStartLocationx"))
            AssertEqual(String(value + 1), binding.control.Value)
            AssertFalse(app.errors.Has(binding.key))
        }
    } finally app.Dispose()
}

TestRunner.Register("Settings input pulse sliders keep native drag positions and do not repaint unchanged help", SettingsSliderTrackingTest)
SettingsSliderTrackingTest() {
    app := SettingsWorkbenchFixture(), observer := 0
    try {
        app.S_Gui.Show("NoActivate x-3000 y-3000 w1000 h740")
        Sleep(180)
        app.Navigate("sounds")
        binding := app.Binding("MasterSoundVolume")
        binding.slider.Value := 121
        app.SliderChanged(binding, binding.slider, binding.control, binding.slider, 5)
        Sleep(25)
        observer := SettingsPaintObserver([binding.hint.Hwnd, binding.label.Hwnd, app.status.Hwnd])
        for value in [122, 123, 124, 126, 127, 128, 131, 134, 137, 138] {
            binding.slider.Value := value
            app.SliderChanged(binding, binding.slider, binding.control, binding.slider, 5)
            AssertEqual(value, binding.slider.Value, "Do not reposition a captured thumb while Windows is tracking it.")
            AssertEqual(Round(value / 5) * 5, app.model.Get("MasterSoundVolume"))
        }
        for handle, count in observer.counts
            AssertEqual(0, count.paints, "Unchanged labels, help and status must not repaint on every slider movement.")
        app.SliderChanged(binding, binding.slider, binding.control, binding.slider, 8)
        AssertEqual(140, binding.slider.Value)
        binding.control.Value := 137
        app.TextChanged(binding, binding.control)
        AssertEqual(137, binding.slider.Value)
    } finally {
        if IsObject(observer)
            observer.Dispose()
        app.Dispose()
    }
}

TestRunner.Register("Settings input pulse event volume sliders also avoid repositioning a captured thumb", SettingsEventSliderTrackingTest)
SettingsEventSliderTrackingTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.Navigate("events")
        event := SettingsEventEditor(app.eventTable, "underAttackByPlayer", 2)
        event.Show()
        for value in [121, 123, 127, 132, 138] {
            event.slider.Value := value
            event.SliderChanged(event.slider, 5)
            AssertEqual(value, event.slider.Value)
            AssertEqual(Round(value / 5) * 5, Integer(event.volume.Value))
        }
        event.SliderChanged(event.slider, 8)
        AssertEqual(140, event.slider.Value)
        app.CloseChild(event.dialog)
    } finally app.Dispose()
}

TestRunner.Register("Settings input pulse rapid cell switches always open the clicked editor and ignore stale focus loss", SettingsRapidCellSwitchTest)
SettingsRapidCellSwitchTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.model.SaveRows("character-shortcuts", [["Alpha", "F13"], ["Beta", "F14"], ["Gamma", "F15"]])
        app.Navigate("character-shortcuts")
        app.S_Gui.Show("NoActivate x-3000 y-3000 w1000 h740")
        Sleep(180)
        grid := app.grids[1]
        for target in [[1, 1], [2, 1], [3, 1], [1, 2], [2, 2], [3, 2], [1, 1], [2, 1]] {
            point := GridTestPoint(grid, target[1], target[2])
            PostMessage(0x201, 1, point, grid.list)
            PostMessage(0x202, 0, point, grid.list)
            Sleep(35)
            ; Posted input shares the test thread with native repaint and
            ; validation callbacks. Wait briefly for that queued pair to finish.
            deadline := A_TickCount + 200
            while !IsObject(grid.editor) && A_TickCount < deadline
                Sleep(10)
            AssertTrue(IsObject(grid.editor), "Every ordinary cell click must open an editor; target=" target[1] ":" target[2]
                "; current=" grid.row ":" grid.column "; focus=" DllCall("GetFocus", "Ptr") "; dragging=" IsObject(grid.drag) "; disposed=" grid.disposed)
            AssertEqual(target, [grid.editRow, grid.editColumn])
            grid.EditorLostFocus(grid.editor)
            AssertTrue(IsObject(grid.editor), "An old focus event must not close the editor which has regained focus.")
        }
        app.search.Focus()
        Sleep(25)
        AssertFalse(IsObject(grid.editor), "Leaving the table still commits the current cell.")
    } finally app.Dispose()
}

TestRunner.Register("Settings input pulse dragging within an unchanged range does not repaint selected cells", SettingsDragNoopPaintingTest)
SettingsDragNoopPaintingTest() {
    app := SettingsWorkbenchFixture(), observer := 0
    try {
        app.model.SaveRows("custom-names", [["Alpha", "A"], ["Beta", "B"], ["Gamma", "C"]])
        app.Navigate("labels")
        app.S_Gui.Show("NoActivate x-3000 y-3000 w1000 h740")
        Sleep(180)
        grid := app.grids[1]
        grid.Select(1, 1)
        grid.Select(2, 2, true)
        GridPostPointer(grid, 1, 1, 1, 0x201)
        GridPostPointer(grid, 3, 2, 1, 0x200)
        selected := JSON.Dump(grid.selected)
        observer := SettingsPaintObserver([grid.list.Hwnd])
        point := GridTestPoint(grid, 3, 2)
        loop 15
            grid.Pointer(1, point + Mod(A_Index, 4), 0x200, grid.list.Hwnd)
        AssertEqual(selected, JSON.Dump(grid.selected))
        AssertEqual(0, observer.counts[grid.list.Hwnd].paints)
        GridPostPointer(grid, 3, 2, 0, 0x202)
    } finally {
        if IsObject(observer)
            observer.Dispose()
        app.Dispose()
    }
}

TestRunner.Register("Settings input pulse pane scrollbar owns its first paint and first drag without a native nonclient bar", SettingsPaneFirstScrollTest)
SettingsPaneFirstScrollTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.model.Set("SettingsTheme", "dark")
        app.Changed()
        app.Navigate("appearance")
        app.S_Gui.Show("NoActivate x-3000 y-3000 w1000 h740")
        Sleep(180)
        pane := app.settingsScroller, handle := pane.scrollbarHwnd
        AssertFalse(WinGetStyle(pane.viewportHwnd) & 0x200000, "No system nonclient scrollbar can repaint behind the custom pane bar.")
        AssertEqual("PaneScrollbar", app.theme.chrome.windows[handle].kind)
        loop 2 {
            AssertEqual(SettingsTheme.BGR("25282D"), SettingsChromeTrackPixel(app, handle, 1))
            point := SettingsChromeStartDrag(app, handle, 1)
            point[2] += 30
            SettingsChromeMoveDrag(handle, point)
            AssertTrue(pane.offset > 0)
            AssertEqual(SettingsTheme.BGR("25282D"), SettingsChromeTrackPixel(app, handle, 1))
            SettingsChromeMoveDrag(handle, point, true)
        }
        DllCall("SendMessageW", "Ptr", handle, "UInt", 0x100, "Ptr", 0x24, "Ptr", 0)
        AssertEqual(0, pane.offset)
        DllCall("SendMessageW", "Ptr", handle, "UInt", 0x100, "Ptr", 0x28, "Ptr", 0)
        AssertTrue(pane.offset > 0)
        AssertEqual(pane.offset, app.theme.chrome.scrollbars.Range(handle, 1).position)
        accessible := Buffer(60, 0)
        NumPut("UInt", 60, accessible)
        AssertTrue(DllCall("SendMessageW", "Ptr", handle, "UInt", 0xEB, "Ptr", 0, "Ptr", accessible))
        AssertEqual(NumGet(app.theme.chrome.ScrollInfo(handle, -5), 24, "Int"), NumGet(accessible, 24, "Int"))
    } finally app.Dispose()
}
