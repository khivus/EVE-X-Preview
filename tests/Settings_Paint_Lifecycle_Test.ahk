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
