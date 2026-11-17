SettingsPaintPixel(control, x := 1, y := 1) {
    rect := Buffer(16, 0)
    DllCall("GetClientRect", "Ptr", control.Hwnd, "Ptr", rect)
    dc := DllCall("GetDC", "Ptr", control.Hwnd, "Ptr")
    memory := DllCall("CreateCompatibleDC", "Ptr", dc, "Ptr")
    bitmap := DllCall("CreateCompatibleBitmap", "Ptr", dc, "Int", Max(1, NumGet(rect, 8, "Int")), "Int", Max(1, NumGet(rect, 12, "Int")), "Ptr")
    previous := DllCall("SelectObject", "Ptr", memory, "Ptr", bitmap, "Ptr")
    try {
        SendMessage(0x318, memory, 4, control)
        return DllCall("GetPixel", "Ptr", memory, "Int", x, "Int", y, "UInt")
    } finally {
        DllCall("SelectObject", "Ptr", memory, "Ptr", previous)
        DllCall("DeleteObject", "Ptr", bitmap)
        DllCall("DeleteDC", "Ptr", memory)
        DllCall("ReleaseDC", "Ptr", control.Hwnd, "Ptr", dc)
    }
}

TestRunner.Register("Settings polish palette switches repaint existing labels and buttons in both launch themes", SettingsPaletteRepaintTest)
SettingsPaletteRepaintTest() {
    for initial in ["dark", "light"] {
        owner := SettingsWorkbenchOwnerFixture()
        owner._JSON["SettingsTheme"] := initial
        app := SettingsWorkbenchFixture(owner)
        try {
            app.Navigate("interface")
            label := app.Binding("SettingsTheme").label
            for mode in ["dark", "light", "dark", "light"] {
                app.model.Set("SettingsTheme", mode)
                app.Changed()
                AssertEqual(SettingsTheme.BGR(app.theme.background), SettingsPaintPixel(label, 0, 0), "A label must not retain its previous palette after " initial " -> " mode ".")
                AssertEqual(SettingsTheme.BGR(app.theme.dark ? "30343B" : "EFEFEF"), SettingsPaintPixel(app.manageProfiles))
                AssertEqual(mode, app.model.Get("SettingsTheme"))
            }
        } finally app.Dispose()
    }
}

TestRunner.Register("Settings polish changing enabled state immediately repaints inherited text without scrolling", SettingsDisabledRepaintTest)
SettingsDisabledRepaintTest() {
    app := SettingsWorkbenchFixture()
    callback := 0, handle := 0, observed := {paints: 0}
    try {
        app.S_Gui.Show("NoActivate x-3000 y-3000 w1000 h740")
        Sleep(180)
        app.model.Set("SettingsTheme", "dark")
        app.CreateProfile("Fleet")
        app.SetSectionInheritance("appearance", true)
        app.Navigate("appearance")
        label := app.Binding("ThumbnailTextFont").label
        handle := label.Hwnd
        callback := CallbackCreate(Observe,, 6)
        DllCall("comctl32\SetWindowSubclass", "Ptr", handle, "Ptr", callback, "UPtr", 99, "UPtr", 0)
        label.Enabled := true
        observed.paints := 0
        label.Enabled := false
        AssertTrue(observed.paints > 0, "WM_ENABLE must be followed by a full custom repaint in the same state change.")
        AssertEqual(SettingsTheme.BGR(app.theme.background), SettingsPaintPixel(label, 0, 0))
    } finally {
        if callback {
            DllCall("comctl32\RemoveWindowSubclass", "Ptr", handle, "Ptr", callback, "UPtr", 99)
            CallbackFree(callback)
        }
        app.Dispose()
    }
    Observe(hwnd, message, wParam, lParam, subclass, data) {
        if message = 0xF
            observed.paints++
        return DllCall("comctl32\DefSubclassProc", "Ptr", hwnd, "UInt", message, "Ptr", wParam, "Ptr", lParam, "Ptr")
    }
}

TestRunner.Register("Settings polish search matches are literal case-insensitive and merge overlapping words", SettingsSearchRangesTest)
SettingsSearchRangesTest() {
    matches := SettingsSearchText.Matches("Label COLOR color", "col COLOR")
    AssertEqual(2, matches.Length)
    AssertEqual(7, matches[1].start)
    AssertEqual(11, matches[1].end)
    AssertEqual(13, matches[2].start)
    AssertEqual(17, matches[2].end)
    AssertEqual(1, SettingsSearchText.Matches("Use [x] literally", "[x]").Length)
    AssertEqual(0, SettingsSearchText.Matches("No change", "").Length)
}

TestRunner.Register("Settings polish search highlights visible text in both palettes without changing field values", SettingsSearchPaintTest)
SettingsSearchPaintTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.S_Gui.Show("NoActivate x-3000 y-3000 w1000 h740")
        Sleep(180)
        for mode in ["dark", "light"] {
            app.model.Set("SettingsTheme", mode)
            app.Changed()
            app.search.Value := "background"
            app.RunSearch()
            binding := app.Binding("ThumbnailBackgroundColor")
            before := binding.control.Value
            AssertEqual(SettingsTheme.BGR(mode = "dark" ? "6B571A" : "FFE08A"), SettingsPaintPixel(binding.label, 1, 0))
            AssertEqual(before, binding.control.Value)
            app.search.Value := ""
            app.RunSearch()
            AssertEqual(SettingsTheme.BGR(app.theme.background), SettingsPaintPixel(app.Binding("ThumbnailBackgroundColor").label, 1, 0))
            app.search.Value := "beta"
            SetTimer(app.searchTimer, 0)
            wrapped := app.MainFrame.AddText("x20 y900 w120 h80 +0x80", "alpha`nbeta")
            dc := DllCall("GetDC", "Ptr", wrapped.Hwnd, "Ptr")
            previous := DllCall("SelectObject", "Ptr", dc, "Ptr", SendMessage(0x31, 0, 0, wrapped), "Ptr")
            size := Buffer(8, 0)
            DllCall("GetTextExtentPoint32W", "Ptr", dc, "WStr", "Mg", "Int", 2, "Ptr", size)
            height := NumGet(size, 4, "Int")
            DllCall("SelectObject", "Ptr", dc, "Ptr", previous)
            DllCall("ReleaseDC", "Ptr", wrapped.Hwnd, "Ptr", dc)
            AssertEqual(SettingsTheme.BGR(mode = "dark" ? "6B571A" : "FFE08A"), SettingsPaintPixel(wrapped, 1, height), "Matches must follow multiline text layout.")
            app.search.Value := ""
            app.RunSearch()
        }
    } finally app.Dispose()
}

class SettingsPaneCaptureFixture extends HotkeyCapture {
    __New(root) => this.expectedRoot := root
    HasFocus() => this.ownerHwnd = this.expectedRoot
}
TestRunner.Register("Settings polish recording uses the top-level settings or row-dialog window and stays active", SettingsRecordingRootTest)
SettingsRecordingRootTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.S_Gui.Show("NoActivate x-3000 y-3000 w1000 h740")
        Sleep(180)
        app.Navigate("shortcuts")
        helper := app.Binding("HideThumbnailsHotkey").helper
        capture := SettingsPaneCaptureFixture(app.S_Gui.Hwnd)
        capture.Start(helper.gui, helper, ObjBindMethod(app, "RecordedShortcut", helper))
        AssertEqual(app.S_Gui.Hwnd, capture.ownerHwnd)
        AssertNotEqual(helper.gui.Hwnd, capture.ownerHwnd)
        Sleep(160)
        AssertTrue(HotkeyCapture.Active, "The 100 ms focus check must not cancel a recording in a child pane.")
        capture.KeyDown(capture.hook, GetKeyVK("F23"), GetKeySC("F23"))
        capture.KeyUp(capture.hook, GetKeyVK("F23"), GetKeySC("F23"))
        AssertFalse(HotkeyCapture.Active)
        AssertEqual("F23", app.model.Get("HideThumbnailsHotkey"))
        app.model.SaveRows("character-shortcuts", [["Alice", "F24"]])
        app.Navigate("character-shortcuts")
        app.tableEditors[1].Edit(1)
        helper := app.recorders[-1]
        capture := SettingsPaneCaptureFixture(app.activeDialog.Hwnd)
        capture.Start(helper.gui, helper, ObjBindMethod(app, "RecordedShortcut", helper))
        AssertEqual(app.activeDialog.Hwnd, capture.ownerHwnd)
        AssertNotEqual(app.S_Gui.Hwnd, capture.ownerHwnd, "An owned editor must retain its own activation root.")
        capture.Cancel()
        app.CloseChild(app.activeDialog)
    } finally {
        if HotkeyCapture.Active
            HotkeyCapture.Current.Cancel()
        app.Dispose()
    }
}

TestRunner.Register("Settings polish mouse drag toggles previous cell selections and overlapping ranges", SettingsDragToggleTest)
SettingsDragToggleTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.model.SaveRows("character-shortcuts", [["Alice", "F20"], ["Bob", "F21"], ["Charlie", "F22"], ["Dave", "F23"]])
        app.Navigate("character-shortcuts")
        grid := app.tableEditors[1].grid
        loop 2 {
            GridPostPointer(grid, 1, 1, 1, 0x201)
            GridPostPointer(grid, 2, 2, 1, 0x200)
            GridPostPointer(grid, 2, 2, 0, 0x202)
            AssertEqual(A_Index = 1 ? 4 : 0, grid.selected.Count)
        }
        GridPostPointer(grid, 1, 1, 1, 0x201)
        GridPostPointer(grid, 2, 2, 1, 0x200)
        GridPostPointer(grid, 2, 2, 0, 0x202)
        GridPostPointer(grid, 2, 2, 1, 0x201)
        GridPostPointer(grid, 3, 2, 1, 0x200)
        GridPostPointer(grid, 3, 2, 0, 0x202)
        AssertEqual(4, grid.selected.Count)
        AssertFalse(grid.selected.Has("2:2"))
        AssertTrue(grid.selected.Has("3:2"))
        AssertTrue(grid.selected.Has("1:1"), "Cells outside the second drag must keep their selection.")
        GridPostPointer(grid, 1, 1, 9, 0x201)
        GridPostPointer(grid, 3, 2, 9, 0x200)
        GridPostPointer(grid, 3, 2, 8, 0x202)
        AssertEqual(2, grid.selected.Count)
        AssertTrue(grid.selected.Has("2:2"))
        AssertTrue(grid.selected.Has("3:1"))
    } finally app.Dispose()
}
