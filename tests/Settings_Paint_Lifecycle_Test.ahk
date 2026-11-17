TestRunner.Register("Settings painting buttons keep their palette through default focus pressed and enabled states", SettingsButtonStatesTest)
SettingsButtonStatesTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.S_Gui.Show("NoActivate x-3000 y-3000 w1000 h740")
        for mode in ["dark", "light", "dark"] {
            app.model.Set("SettingsTheme", mode)
            app.Changed()
            button := app.applyButton
            button.Enabled := true
            for style in [0, 1, 0, 1] {
                SendMessage(0xF4, style, 1, button)
                AssertEqual(style, WinGetStyle(button.Hwnd) & 0xF)
                SendMessage(0xF3, 1, 0, button)
                AssertEqual(SettingsTheme.BGR(mode = "dark" ? "3D4859" : "CEE2F6"), SettingsPaintPixel(button))
                SendMessage(0xF3, 0, 0, button)
                AssertEqual(SettingsTheme.BGR(mode = "dark" ? "30343B" : "EFEFEF"), SettingsPaintPixel(button))
                button.Enabled := false
                AssertEqual(SettingsTheme.BGR(mode = "dark" ? "30343B" : "EFEFEF"), SettingsPaintPixel(button))
                button.Enabled := true
            }
            app.closeButton.Focus()
            button.Focus()
            AssertEqual(SettingsTheme.BGR(mode = "dark" ? "30343B" : "EFEFEF"), SettingsPaintPixel(button))
            AssertTrue(SendMessage(0x87, 0, 0, button) & 0x10, "Native button dialog-key handling must remain available.")
        }
    } finally app.Dispose()
}

TestRunner.Register("Settings painting native button clicks and hover feedback survive custom drawing", SettingsButtonInputTest)
SettingsButtonInputTest() {
    app := SettingsWorkbenchFixture()
    clicked := {count: 0}
    try {
        app.model.Set("SettingsTheme", "dark")
        app.Changed()
        button := app.S_Gui.AddButton("x260 y700 w100 h26", "Test action")
        button.OnEvent("Click", (*) => clicked.count++)
        PostMessage(0xF5, 0, 0, button)
        Sleep(45)
        AssertEqual(1, clicked.count)
        SendMessage(0x200, 0, 0, button)
        AssertEqual(SettingsTheme.BGR("39404B"), SettingsPaintPixel(button))
        SendMessage(0x2A3, 0, 0, button)
        AssertEqual(SettingsTheme.BGR("30343B"), SettingsPaintPixel(button))
    } finally app.Dispose()
}

TestRunner.Register("Settings painting tray profile activation and reactivation keep Apply dark", SettingsTrayButtonRepaintTest)
SettingsTrayButtonRepaintTest() {
    owner := SettingsWorkbenchOwnerFixture()
    owner._JSON["SettingsTheme"] := "dark"
    owner.Profiles["Fleet"] := SettingsModel.Clone(owner.Profiles["Default"])
    app := SettingsWorkbenchFixture(owner)
    owner.settingsWorkbench := app
    try {
        app.S_Gui.Show("NoActivate x-3000 y-3000 w1000 h740")
        AssertFalse(owner.ActivateProfile("Fleet", "Tray"))
        AssertEqual("Fleet", app.SelectedProfile)
        AssertTrue(app.applyButton.Enabled)
        for active in [0, 1, 0, 1] {
            DllCall("SendMessageW", "Ptr", app.S_Gui.Hwnd, "UInt", 0x6, "Ptr", active, "Ptr", 0)
            SendMessage(0xF4, 1, 1, app.applyButton)
            AssertEqual(SettingsTheme.BGR("30343B"), SettingsPaintPixel(app.applyButton))
        }
        AssertEqual("Fleet", app.model.config["LastUsedProfile"])
        AssertEqual("Default", owner.LastUsedProfile)
    } finally app.Dispose()
}

TestRunner.Register("Settings painting interactive resize retains the pane until the final width is known", SettingsResizeLifecycleTest)
SettingsResizeLifecycleTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.S_Gui.Show("NoActivate x-3000 y-3000 w1000 h740")
        Sleep(180)
        app.Navigate("appearance")
        canvas := app.MainFrame.Hwnd
        app.SizeMove(0, 0, 0x231, app.S_Gui.Hwnd)
        for width in [900, 780, 1100, 800] {
            app.OnResize(app.S_Gui, 0, width, 600)
            app.ResizeContent()
            AssertEqual(canvas, app.MainFrame.Hwnd, "Sizing must not destroy and recreate the content between drag events.")
        }
        app.SizeMove(0, 0, 0x232, app.S_Gui.Hwnd)
        AssertFalse(app.sizing)
        AssertEqual(app.contentWidth, app.renderedContentWidth)
        AssertTrue(DllCall("IsWindow", "Ptr", app.MainFrame.Hwnd))
        AssertEqual(0, app.theme.updates.Count)
    } finally app.Dispose()
}

TestRunner.Register("Settings painting height-only resizes retain inline table editors and positions", SettingsHeightResizeTest)
SettingsHeightResizeTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.Navigate("character-shortcuts")
        grid := app.grids[1]
        grid.Select(1, 1)
        grid.BeginEdit("Unfinished pilot")
        canvas := app.MainFrame.Hwnd, editor := grid.editor.Hwnd
        app.OnResize(app.S_Gui, 0, app.width, app.height + 120)
        app.ResizeContent()
        AssertEqual(canvas, app.MainFrame.Hwnd)
        AssertEqual(editor, grid.editor.Hwnd)
        AssertEqual("Unfinished pilot", grid.editor.Value)
        AssertEqual(0, app.theme.updates.Count)
    } finally app.Dispose()
}

TestRunner.Register("Settings painting update batches nest safely without global child compositing", SettingsPaintBatchTest)
SettingsPaintBatchTest() {
    app := SettingsWorkbenchFixture()
    try {
        before := !!DllCall("IsWindowVisible", "Ptr", app.S_Gui.Hwnd)
        app.theme.BeginUpdate(app.S_Gui)
        try {
            app.theme.BeginUpdate(app.S_Gui)
            try app.Layout()
            finally app.theme.EndUpdate(app.S_Gui)
            AssertEqual(1, app.theme.updates[app.S_Gui.Hwnd].depth)
        } finally app.theme.EndUpdate(app.S_Gui)
        AssertEqual(before, !!DllCall("IsWindowVisible", "Ptr", app.S_Gui.Hwnd))
        AssertEqual(0, app.theme.updates.Count)
        AssertFalse(WinGetExStyle(app.S_Gui.Hwnd) & 0x02000000, "Whole-window compositing conflicts with local native/custom control buffers.")
        dialog := SettingsChildDialog(app, "Buffered child fixture")
        try {
            dialog.AddText("x20 y20 w300 h26", "Resize test")
            dialog.Show("w440 h300")
            dialog.Layout(500, 350)
            AssertFalse(WinGetExStyle(dialog.Hwnd) & 0x02000000)
            AssertEqual(0, app.theme.updates.Count)
        } finally dialog.Destroy()
    } finally app.Dispose()
}

TestRunner.Register("Settings painting repeated resize theme and profile changes retain dark scrollbar pixels", SettingsResizePaletteTest)
SettingsResizePaletteTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.model.Set("SettingsTheme", "dark")
        app.CreateProfile("Fleet")
        app.Navigate("appearance")
        app.S_Gui.Show("NoActivate x-3000 y-3000 w1000 h740")
        Sleep(180)
        for dimensions in [[800, 500], [1100, 740], [780, 620], [1000, 740]] {
            app.width := dimensions[1], app.height := dimensions[2]
            app.S_Gui.Show("NoActivate x-3000 y-3000 w" app.width " h" app.height)
            app.Layout()
            app.ResizeContent()
            AssertEqual(SettingsTheme.BGR("30343B"), SettingsPaintPixel(app.applyButton))
            handle := app.settingsScroller.scrollbarHwnd
            info := app.theme.chrome.ScrollInfo(handle, -5)
            if IsObject(info)
                AssertEqual(SettingsTheme.BGR("25282D"), SettingsChromeTrackPixel(app, handle, 1))
            AssertEqual(0, app.theme.updates.Count)
        }
    } finally app.Dispose()
}

TestRunner.Register("Settings painting visible update batches preserve top-level visibility and focus", SettingsVisibleBatchTest)
SettingsVisibleBatchTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.S_Gui.Show("NoActivate x-3000 y-3000 w1000 h740")
        Sleep(180)
        DllCall("SetActiveWindow", "Ptr", app.S_Gui.Hwnd, "Ptr")
        app.search.Focus()
        active := DllCall("GetActiveWindow", "Ptr"), focus := DllCall("GetFocus", "Ptr")
        app.theme.BeginUpdate(app.S_Gui)
        try {
            AssertTrue(DllCall("IsWindowVisible", "Ptr", app.S_Gui.Hwnd), "Settings must remain visible throughout an update, including to the taskbar.")
            AssertTrue(DllCall("IsWindowVisible", "Ptr", app.search.Hwnd), "The parent must retain its visible flag while child controls update.")
            app.theme.BeginUpdate(app.S_Gui)
            try app.Layout()
            finally app.theme.EndUpdate(app.S_Gui)
            AssertTrue(DllCall("IsWindowVisible", "Ptr", app.S_Gui.Hwnd))
            AssertEqual(1, app.theme.updates[app.S_Gui.Hwnd].depth)
            AssertEqual(active, DllCall("GetActiveWindow", "Ptr"))
            AssertEqual(focus, DllCall("GetFocus", "Ptr"))
        } finally app.theme.EndUpdate(app.S_Gui)
        AssertEqual(0, app.theme.updates.Count)

        dialog := SettingsChildDialog(app, "Visible update fixture")
        try {
            dialog.AddText("x20 y20 w300 h26", "Keep this owned dialog visible too")
            dialog.Show("w440 h300")
            app.theme.BeginUpdate(dialog.root)
            try {
                AssertTrue(DllCall("IsWindowVisible", "Ptr", dialog.Hwnd), "Owned editors are also top-level windows.")
                dialog.Layout(500, 350)
                AssertTrue(DllCall("IsWindowVisible", "Ptr", dialog.Hwnd))
            } finally app.theme.EndUpdate(dialog.root)
        } finally dialog.Destroy()
    } finally app.Dispose()
}

; Check visibility during native message dispatch; final-state checks miss the
; brief WS_VISIBLE removal that makes the shell remove and recreate a task button.
class SettingsVisibilityObserver {
    static WindowOrder(handles) {
        wanted := Map(), order := []
        for index, handle in handles
            wanted[handle] := index
        handle := DllCall("GetTopWindow", "Ptr", 0, "Ptr")
        while handle {
            if wanted.Has(handle)
                order.Push(wanted[handle])
            handle := DllCall("GetWindow", "Ptr", handle, "UInt", 2, "Ptr")
        }
        return order
    }
    __New(handle) {
        this.handle := handle, this.hidden := 0
        this.callback := CallbackCreate(ObjBindMethod(this, "WindowProc"),, 6)
        DllCall("comctl32\SetWindowSubclass", "Ptr", handle, "Ptr", this.callback, "UPtr", 97, "UPtr", 0)
    }
    WindowProc(handle, message, wParam, lParam, subclass, data) {
        result := DllCall("comctl32\DefSubclassProc", "Ptr", handle, "UInt", message, "Ptr", wParam, "Ptr", lParam, "Ptr")
        if !DllCall("IsWindowVisible", "Ptr", handle)
            this.hidden++
        return result
    }
    Dispose() {
        if DllCall("IsWindow", "Ptr", this.handle)
            DllCall("comctl32\RemoveWindowSubclass", "Ptr", this.handle, "Ptr", this.callback, "UPtr", 97)
        CallbackFree(this.callback)
    }
}

TestRunner.Register("Settings painting buttons and page rebuilds never hide or reorder the settings window", SettingsButtonWindowContinuityTest)
SettingsButtonWindowContinuityTest() {
    app := SettingsWorkbenchFixture(), observer := 0, backgrounds := []
    try {
        app.S_Gui.Show("NoActivate x-3000 y-3000 w1000 h740")
        Sleep(180)
        handles := [app.S_Gui.Hwnd]
        loop 2 {
            background := Gui("+ToolWindow", "Settings stacking fixture " A_Index)
            backgrounds.Push(background)
            background.Show("NoActivate x-3000 y-3000 w200 h100")
            DllCall("SetWindowPos", "Ptr", background.Hwnd, "Ptr", handles[-1],
                "Int", 0, "Int", 0, "Int", 0, "Int", 0, "UInt", 0x13)
            handles.Push(background.Hwnd)
        }
        AssertArrayEqual([1, 2, 3], SettingsVisibilityObserver.WindowOrder(handles))
        DllCall("SetActiveWindow", "Ptr", app.S_Gui.Hwnd, "Ptr")
        app.manageProfiles.Focus()
        active := DllCall("GetActiveWindow", "Ptr"), focus := DllCall("GetFocus", "Ptr")
        foreground := DllCall("GetForegroundWindow", "Ptr")
        observer := SettingsVisibilityObserver(app.S_Gui.Hwnd)
        SendMessage(0xF5, 0, 0, app.manageProfiles) ; BM_CLICK follows the real button handler.
        Sleep(35)
        AssertEqual("profiles", app.renderedSection)
        app.CreateProfile("Fleet")
        app.Navigate("appearance")
        app.EditDefaultHere("appearance")
        app.EndDefaultEditing()
        for mode in ["dark", "light"] {
            app.model.Set("SettingsTheme", mode)
            app.Changed()
            app.Navigate("layout")
            app.Layout()
        }
        app.DiscardDraft()
        AssertEqual(0, observer.hidden, "No intermediate message may see Settings hidden during ordinary actions.")
        AssertEqual(active, DllCall("GetActiveWindow", "Ptr"), "Page actions must retain the active window.")
        AssertEqual(focus, DllCall("GetFocus", "Ptr"), "A surviving sidebar button retains keyboard focus.")
        AssertEqual(foreground, DllCall("GetForegroundWindow", "Ptr"), "Page actions must retain the foreground window.")
        AssertArrayEqual([1, 2, 3], SettingsVisibilityObserver.WindowOrder(handles), "Settings and background fixtures must retain their relative window order.")
        AssertEqual(0, app.theme.updates.Count)
    } finally {
        if IsObject(observer)
            observer.Dispose()
        app.Dispose()
        for background in backgrounds
            background.Destroy()
    }
}
