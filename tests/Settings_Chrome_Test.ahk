SettingsWindowPixel(handle, x, y) {
    rect := Buffer(16, 0)
    DllCall("GetWindowRect", "Ptr", handle, "Ptr", rect)
    width := NumGet(rect, 8, "Int") - NumGet(rect, 0, "Int")
    height := NumGet(rect, 12, "Int") - NumGet(rect, 4, "Int")
    dc := DllCall("GetWindowDC", "Ptr", handle, "Ptr")
    memory := DllCall("CreateCompatibleDC", "Ptr", dc, "Ptr")
    bitmap := DllCall("CreateCompatibleBitmap", "Ptr", dc, "Int", Max(1, width), "Int", Max(1, height), "Ptr")
    previous := DllCall("SelectObject", "Ptr", memory, "Ptr", bitmap, "Ptr")
    try {
        DllCall("SendMessageW", "Ptr", handle, "UInt", 0x317, "Ptr", memory, "Ptr", 0x1E)
        return DllCall("GetPixel", "Ptr", memory, "Int", x, "Int", y, "UInt")
    } finally {
        DllCall("SelectObject", "Ptr", memory, "Ptr", previous)
        DllCall("DeleteObject", "Ptr", bitmap)
        DllCall("DeleteDC", "Ptr", memory)
        DllCall("ReleaseDC", "Ptr", handle, "Ptr", dc)
    }
}

TestRunner.Register("Settings chrome dark borders and arrows are flat and palette switching preserves editable hotkeys", SettingsChromeBordersTest)
SettingsChromeBordersTest() {
    owner := SettingsWorkbenchOwnerFixture()
    owner._JSON["SettingsTheme"] := "light"
    app := SettingsWorkbenchFixture(owner)
    try {
        app.Navigate("shortcuts")
        binding := app.Binding("HideThumbnailsHotkey")
        app.AdvancedShortcut(binding.helper)
        originalBorder := SettingsNativeChrome.Long(binding.control.Hwnd, -16) & 0x800000
        originalEdge := SettingsNativeChrome.Long(binding.control.Hwnd, -20) & 0x200
        app.model.Set("SettingsTheme", "dark")
        app.Changed()
        AssertEqual(0, SettingsNativeChrome.Long(binding.control.Hwnd, -20) & 0x200)
        AssertEqual(SettingsTheme.BGR(app.theme.line), SettingsWindowPixel(binding.control.Hwnd, 0, 0))
        info := app.theme.chrome.ComboInfo(app.profileSelector.Hwnd)
        AssertEqual(SettingsTheme.BGR("30343B"), SettingsPaintPixel(app.profileSelector, NumGet(info, 20, "Int") + 2, NumGet(info, 24, "Int") + 2))
        app.model.Set("SettingsTheme", "light")
        app.Changed()
        AssertEqual(originalBorder, SettingsNativeChrome.Long(binding.control.Hwnd, -16) & 0x800000)
        AssertEqual(originalEdge, SettingsNativeChrome.Long(binding.control.Hwnd, -20) & 0x200)
        AssertFalse(WinGetStyle(binding.control.Hwnd) & 0x800, "Restoring borders must not restore an old ReadOnly flag.")
        AssertTrue(binding.helper.rawMode)
    } finally app.Dispose()
}

TestRunner.Register("Settings chrome spinner and slider drawing retain native value changes", SettingsChromeValuesTest)
SettingsChromeValuesTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.S_Gui.Show("NoActivate x-3000 y-3000 w1000 h740")
        Sleep(180)
        app.model.Set("SettingsTheme", "dark")
        app.Changed()
        app.Navigate("appearance")
        binding := app.Binding("ThumbnailTextSize")
        for control in binding.inputs
            if control.Type = "UpDown"
                spin := control
        AssertEqual(SettingsTheme.BGR("30343B"), SettingsPaintPixel(spin, 2, 2))
        before := app.model.Get("ThumbnailTextSize")
        rect := Buffer(16, 0)
        DllCall("GetClientRect", "Ptr", spin.Hwnd, "Ptr", rect)
        point := (NumGet(rect, 8, "Int") // 2) | ((NumGet(rect, 12, "Int") // 4) << 16)
        PostMessage(0x201, 1, point, spin)
        PostMessage(0x202, 0, point, spin)
        Sleep(50)
        AssertEqual(before + 1, app.model.Get("ThumbnailTextSize"))
        slider := app.Binding("ThumbnailOpacity").slider
        SendMessage(0x41A, 0, rect.Ptr, slider)
        AssertEqual(SettingsTheme.BGR("424750"), SettingsPaintPixel(slider, NumGet(rect, 0, "Int") + 1, (NumGet(rect, 4, "Int") + NumGet(rect, 12, "Int")) // 2))
        before := app.model.Get("ThumbnailOpacity")
        PostMessage(0x100, 0x25, 0, slider)
        PostMessage(0x101, 0x25, 0, slider)
        Sleep(35)
        AssertEqual(before - 5, app.model.Get("ThumbnailOpacity"))
    } finally app.Dispose()
}

TestRunner.Register("Settings chrome dropdown popups keep native selection and dark borders", SettingsChromeDropdownTest)
SettingsChromeDropdownTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.S_Gui.Show("NoActivate x-3000 y-3000 w1000 h740")
        Sleep(180)
        app.model.Set("SettingsTheme", "dark")
        app.Changed()
        app.Navigate("appearance")
        binding := app.Binding("EventActiveBorderMode")
        before := binding.control.Value
        SendMessage(0x14F, 1, 0, binding.control)
        info := app.theme.chrome.ComboInfo(binding.control.Hwnd)
        popup := NumGet(info, A_PtrSize = 8 ? 56 : 48, "Ptr")
        AssertTrue(app.theme.chrome.windows.Has(popup))
        AssertEqual(SettingsTheme.BGR(app.theme.line), SettingsWindowPixel(popup, 0, 0))
        PostMessage(0x100, 0x28, 0, binding.control)
        PostMessage(0x101, 0x28, 0, binding.control)
        Sleep(35)
        ; A changed border layout rebuilds the section and its native controls.
        binding := app.Binding("EventActiveBorderMode")
        SendMessage(0x14F, 0, 0, binding.control)
        AssertEqual(before + 1, binding.control.Value)
        AssertEqual(binding.field.choices[binding.control.Value].value, app.model.Get("EventActiveBorderMode"))
    } finally app.Dispose()
}

TestRunner.Register("Settings chrome native scrollbars use the dark track while retaining table positions", SettingsChromeScrollbarsTest)
SettingsChromeScrollbarsTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.model.Set("SettingsTheme", "dark")
        app.Changed()
        app.model.AddGroup("characters", "Fleet")
        rows := []
        loop 80
            rows.Push(["Pilot " A_Index])
        app.model.SaveRows("group-members", rows, "Fleet")
        app.Navigate("character-groups")
        list := app.tableEditors[1].list
        list.ModifyCol(1, 1200)
        SendMessage(0x1014, 90, 600, list)
        first := SendMessage(0x1027, 0, 0, list)
        AssertTrue(first > 0)
        window := Buffer(16, 0)
        DllCall("GetWindowRect", "Ptr", list.Hwnd, "Ptr", window)
        info := app.theme.chrome.ScrollInfo(list.Hwnd, -5)
        AssertTrue(IsObject(info))
        x := NumGet(info, 4, "Int") - NumGet(window, 0, "Int") + 1
        y := NumGet(info, 8, "Int") - NumGet(window, 4, "Int") + NumGet(info, 20, "Int") + 2
        AssertEqual(SettingsTheme.BGR("25282D"), SettingsWindowPixel(list.Hwnd, x, y))
        horizontal := app.theme.chrome.ScrollInfo(list.Hwnd, -6)
        cornerY := NumGet(horizontal, 8, "Int") - NumGet(window, 4, "Int") + 2
        AssertEqual(SettingsTheme.BGR("25282D"), SettingsWindowPixel(list.Hwnd, x, cornerY))
        AssertEqual(first, SendMessage(0x1027, 0, 0, list))
        AssertTrue(DllCall("GetScrollPos", "Ptr", list.Hwnd, "Int", 0, "Int") > 0)
    } finally app.Dispose()
}

TestRunner.Register("Settings chrome tooltips match both palettes without activating a window", SettingsChromeTooltipTest)
SettingsChromeTooltipTest() {
    app := SettingsWorkbenchFixture()
    try {
        foreground := DllCall("GetForegroundWindow", "Ptr")
        for mode in ["dark", "light"] {
            app.model.Set("SettingsTheme", mode)
            app.Changed()
            app.theme.ShowTooltip("A short explanation of this setting.", -3000, -3000)
            tip := app.theme.tooltip
            AssertTrue(DllCall("IsWindowVisible", "Ptr", tip.Hwnd))
            AssertEqual(foreground, DllCall("GetForegroundWindow", "Ptr"))
            AssertEqual(SettingsTheme.BGR(mode = "dark" ? app.theme.input : "FFFFFF"), tip.Send(0x416, 0, 0))
            AssertEqual(SettingsTheme.BGR(app.theme.foreground), tip.Send(0x417, 0, 0))
            if mode = "dark"
                AssertEqual(SettingsTheme.BGR(app.theme.input), SettingsPaintPixel(tip, 3, 3))
            app.ClearSettingsTooltip()
            AssertFalse(DllCall("IsWindowVisible", "Ptr", tip.Hwnd))
        }
    } finally app.Dispose()
}

SettingsChromePopupFits(app, control) {
    SendMessage(0x14F, 1, 0, control)
    info := app.theme.chrome.ComboInfo(control.Hwnd)
    popup := NumGet(info, A_PtrSize = 8 ? 56 : 48, "Ptr")
    bounds := Buffer(16, 0), rect := Buffer(16, 0)
    root := DllCall("GetAncestor", "Ptr", control.Hwnd, "UInt", 2, "Ptr")
    DllCall("GetClientRect", "Ptr", root, "Ptr", bounds)
    DllCall("MapWindowPoints", "Ptr", root, "Ptr", 0, "Ptr", bounds, "UInt", 2)
    DllCall("GetWindowRect", "Ptr", popup, "Ptr", rect)
    AssertTrue(NumGet(rect, 0, "Int") >= NumGet(bounds, 0, "Int"), "Dropdown must stay inside the left edge.")
    AssertTrue(NumGet(rect, 4, "Int") >= NumGet(bounds, 4, "Int"), "Dropdown must stay inside the top edge.")
    AssertTrue(NumGet(rect, 8, "Int") <= NumGet(bounds, 8, "Int"), "Dropdown must stay inside the right edge.")
    AssertTrue(NumGet(rect, 12, "Int") <= NumGet(bounds, 12, "Int"), "Dropdown must stay inside the bottom edge.")
    AssertTrue(NumGet(rect, 12, "Int") > NumGet(rect, 4, "Int"))
    return rect
}

TestRunner.Register("Settings chrome dropdowns fit compact windows in both palettes including grid editors", SettingsChromePopupBoundsTest)
SettingsChromePopupBoundsTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.S_Gui.Show("NoActivate x-3000 y-3000 w800 h500")
        app.width := 800, app.height := 500
        app.Layout()
        Sleep(180)
        for mode in ["light", "dark"] {
            app.model.Set("SettingsTheme", mode)
            app.Changed()
            app.Navigate("interface")
            control := app.Binding("SettingsTheme").control
            SendMessage(0x160, 1200, 0, control)
            SettingsChromePopupFits(app, control)
            SendMessage(0x14F, 0, 0, control)
            app.Navigate("mouse-actions")
            grid := app.grids[1]
            grid.Select(1, 2)
            grid.BeginEdit()
            SettingsChromePopupFits(app, grid.editor)
            grid.CancelEditor()
        }
    } finally app.Dispose()
}

TestRunner.Register("Settings chrome owned editor dropdowns open above the field when needed", SettingsChromePopupAboveTest)
SettingsChromePopupAboveTest() {
    app := SettingsWorkbenchFixture()
    dialog := app.theme.Surface(Gui("+Owner" app.S_Gui.Hwnd, "Dropdown bounds fixture"))
    try {
        dialog.SetFont("s10", "Segoe UI")
        choices := []
        loop 20
            choices.Push("Choice " A_Index)
        control := dialog.AddDropDownList("x180 y190 w180 h280 Choose1", choices)
        dialog.Show("NoActivate x-3000 y-3000 w400 h240")
        Sleep(30)
        rect := SettingsChromePopupFits(app, control)
        field := Buffer(16, 0)
        DllCall("GetWindowRect", "Ptr", control.Hwnd, "Ptr", field)
        AssertTrue(NumGet(rect, 12, "Int") <= NumGet(field, 4, "Int"), "List should open above a field near the bottom.")
        PostMessage(0x100, 0x28, 0, control)
        PostMessage(0x101, 0x28, 0, control)
        Sleep(25)
        AssertEqual(2, control.Value)
    } finally {
        dialog.Destroy()
        app.Dispose()
    }
}

SettingsChromeStartDrag(app, handle, bar) {
    info := app.theme.chrome.ScrollInfo(handle, bar = 1 ? -5 : -6)
    AssertTrue(IsObject(info))
    x := bar = 1 ? (NumGet(info, 4, "Int") + NumGet(info, 12, "Int")) // 2
        : NumGet(info, 4, "Int") + (NumGet(info, 24, "Int") + NumGet(info, 28, "Int")) // 2
    y := bar = 1 ? NumGet(info, 8, "Int") + (NumGet(info, 24, "Int") + NumGet(info, 28, "Int")) // 2
        : (NumGet(info, 8, "Int") + NumGet(info, 16, "Int")) // 2
    range := app.theme.chrome.scrollbars.Range(handle, bar)
    hit := app.theme.chrome.scrollbars.Hit(handle, bar, [x, y])
    if app.theme.chrome.windows[handle].kind = "PaneScrollbar" {
        point := Buffer(8, 0)
        NumPut("Int", x, "Int", y, point)
        DllCall("ScreenToClient", "Ptr", handle, "Ptr", point)
        DllCall("SendMessageW", "Ptr", handle, "UInt", 0x201, "Ptr", 1,
            "Ptr", (NumGet(point, 0, "Int") & 0xFFFF) | ((NumGet(point, 4, "Int") & 0xFFFF) << 16))
    } else
        DllCall("SendMessageW", "Ptr", handle, "UInt", 0xA1, "Ptr", bar = 1 ? 7 : 6, "Ptr", (x & 0xFFFF) | ((y & 0xFFFF) << 16))
    AssertTrue(IsObject(app.theme.chrome.scrollbars.drag), "Dark dragging must not enter the system light tracking loop. Range=" (IsObject(range) ? JSON.Dump(range) : "none") "; hit=" (IsObject(hit) ? JSON.Dump(hit) : "none"))
    AssertEqual(handle, DllCall("GetCapture", "Ptr"))
    return [x, y]
}
SettingsChromeMoveDrag(handle, point, release := false) {
    coordinates := Buffer(8, 0)
    NumPut("Int", point[1], "Int", point[2], coordinates)
    DllCall("ScreenToClient", "Ptr", handle, "Ptr", coordinates)
    packed := (NumGet(coordinates, 0, "Int") & 0xFFFF) | ((NumGet(coordinates, 4, "Int") & 0xFFFF) << 16)
    DllCall("SendMessageW", "Ptr", handle, "UInt", release ? 0x202 : 0x200, "Ptr", release ? 0 : 1, "Ptr", packed)
}
SettingsChromeTrackPixel(app, handle, bar) {
    info := app.theme.chrome.ScrollInfo(handle, bar = 1 ? -5 : -6)
    window := Buffer(16, 0)
    DllCall("GetWindowRect", "Ptr", handle, "Ptr", window)
    x := NumGet(info, 4, "Int") - NumGet(window, 0, "Int") + (bar = 1 ? 1 : NumGet(info, 20, "Int") + 1)
    y := NumGet(info, 8, "Int") - NumGet(window, 4, "Int") + (bar = 1 ? NumGet(info, 20, "Int") + 1 : 1)
    ; Render while capture is still held; offscreen window DCs are clipped.
    return SettingsWindowPixel(handle, x, y)
}

TestRunner.Register("Settings chrome held pane dragging stays dark and retains 32-bit scroll offsets", SettingsChromePaneDragTest)
SettingsChromePaneDragTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.model.Set("SettingsTheme", "dark")
        app.Changed()
        app.S_Gui.Show("NoActivate x-3000 y-3000 w1000 h740")
        app.Navigate("appearance")
        Sleep(180)
        pane := app.settingsScroller
        pane.extent := 500000
        pane.UpdateRange()
        point := SettingsChromeStartDrag(app, pane.scrollbarHwnd, 1)
        point[2] += Round(app.theme.chrome.scrollbars.drag.travel * 0.75)
        SettingsChromeMoveDrag(pane.scrollbarHwnd, point)
        AssertTrue(pane.offset > 65535, "Thumb dragging must keep the full scroll position.")
        AssertEqual(SettingsTheme.BGR("25282D"), SettingsChromeTrackPixel(app, pane.scrollbarHwnd, 1))
        AssertTrue(IsObject(app.theme.chrome.scrollbars.drag), "The held state is tested before mouse release.")
        SettingsChromeMoveDrag(pane.scrollbarHwnd, point, true)
        AssertFalse(IsObject(app.theme.chrome.scrollbars.drag))
        AssertEqual(0, DllCall("GetCapture", "Ptr"))
    } finally app.Dispose()
}

TestRunner.Register("Settings chrome held table dragging works on both axes and commits active cells", SettingsChromeTableDragTest)
SettingsChromeTableDragTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.model.Set("SettingsTheme", "dark")
        app.Changed()
        app.model.AddGroup("characters", "Fleet")
        rows := []
        loop 100
            rows.Push(["Pilot " A_Index])
        app.model.SaveRows("group-members", rows, "Fleet")
        app.S_Gui.Show("NoActivate x-3000 y-3000 w1000 h740")
        app.Navigate("character-groups")
        Sleep(180)
        grid := app.grids[1], list := grid.list
        list.ModifyCol(1, 1400)
        grid.Select(1, 1)
        grid.BeginEdit()
        grid.editor.Value := "Edited pilot"
        for bar in [1, 0] {
            point := SettingsChromeStartDrag(app, list.Hwnd, bar)
            AssertFalse(IsObject(grid.editor))
            point[bar = 1 ? 2 : 1] += Round(app.theme.chrome.scrollbars.drag.travel * 0.6)
            SettingsChromeMoveDrag(list.Hwnd, point)
            AssertTrue(DllCall("GetScrollPos", "Ptr", list.Hwnd, "Int", bar, "Int") > 0)
            AssertEqual(SettingsTheme.BGR("25282D"), SettingsChromeTrackPixel(app, list.Hwnd, bar))
            SettingsChromeMoveDrag(list.Hwnd, point, true)
        }
        AssertEqual("Edited pilot", app.model.Rows("group-members", "Fleet")[1][1])
    } finally app.Dispose()
}

TestRunner.Register("Settings chrome scrollbar clicks repeat and capture loss or palette changes end tracking", SettingsChromeScrollCaptureTest)
SettingsChromeScrollCaptureTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.model.Set("SettingsTheme", "dark")
        app.Changed()
        app.S_Gui.Show("NoActivate x-3000 y-3000 w1000 h740")
        app.Navigate("appearance")
        Sleep(180)
        pane := app.settingsScroller, handle := pane.scrollbarHwnd
        pane.extent := 10000
        pane.UpdateRange()
        info := app.theme.chrome.ScrollInfo(handle, -5)
        x := (NumGet(info, 4, "Int") + NumGet(info, 12, "Int")) // 2
        y := NumGet(info, 16, "Int") - 3
        click := Buffer(8, 0)
        NumPut("Int", x, "Int", y, click)
        DllCall("ScreenToClient", "Ptr", handle, "Ptr", click)
        DllCall("SendMessageW", "Ptr", handle, "UInt", 0x201, "Ptr", 1,
            "Ptr", (NumGet(click, 0, "Int") & 0xFFFF) | ((NumGet(click, 4, "Int") & 0xFFFF) << 16))
        first := pane.offset
        AssertTrue(first > 0)
        app.theme.chrome.scrollbars.Repeat()
        AssertTrue(pane.offset > first)
        DllCall("ReleaseCapture")
        AssertFalse(IsObject(app.theme.chrome.scrollbars.drag))
        SettingsChromeStartDrag(app, handle, 1)
        app.theme.SetMode("light")
        AssertFalse(IsObject(app.theme.chrome.scrollbars.drag))
        AssertEqual(0, DllCall("GetCapture", "Ptr"))
    } finally app.Dispose()
}

TestRunner.Register("Settings chrome dropdown scrollbars retain the open list and native selection", SettingsChromePopupDragTest)
SettingsChromePopupDragTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.model.Set("SettingsTheme", "dark")
        app.Changed()
        app.S_Gui.Show("NoActivate x-3000 y-3000 w1000 h740")
        app.Navigate("appearance")
        Sleep(180)
        control := app.Binding("ThumbnailTextFont").control
        SendMessage(0x14F, 1, 0, control)
        info := app.theme.chrome.ComboInfo(control.Hwnd)
        popup := NumGet(info, A_PtrSize = 8 ? 56 : 48, "Ptr")
        before := DllCall("GetCapture", "Ptr")
        point := SettingsChromeStartDrag(app, popup, 1)
        point[2] += Round(app.theme.chrome.scrollbars.drag.travel * 0.6)
        SettingsChromeMoveDrag(popup, point)
        AssertTrue(DllCall("GetScrollPos", "Ptr", popup, "Int", 1, "Int") > 0)
        SettingsChromeMoveDrag(popup, point, true)
        AssertTrue(SendMessage(0x157, 0, 0, control), "Releasing a dropdown scrollbar must keep its list open.")
        AssertEqual(before, DllCall("GetCapture", "Ptr"))
        SendMessage(0x14F, 0, 0, control)
    } finally app.Dispose()
}
