class SettingsLayoutFixture extends CombatEventsFixture {
    __New() {
        super.__New()
        this.ProfileOverride()
        this.SetState()
        this.S_Gui := Gui()
        this.CreateSidebar()
        this.CreateMainFrame()
        this.HotkeyGroups_Ctrl()
    }
    Dispose() {
        this.Sidebar.Destroy()
        this.MainFrame.Destroy()
        this.S_Gui.Destroy()
    }
}

SettingsNativeRect(hwnd, parentHwnd) {
    rect := Buffer(16, 0)
    if !DllCall("GetWindowRect", "Ptr", hwnd, "Ptr", rect)
        throw OSError()
    DllCall("MapWindowPoints", "Ptr", 0, "Ptr", parentHwnd, "Ptr", rect, "UInt", 2)
    return {left: NumGet(rect, 0, "Int"), top: NumGet(rect, 4, "Int"), right: NumGet(rect, 8, "Int"), bottom: NumGet(rect, 12, "Int")}
}

TestRunner.Register("Settings panels stay adjacent at simulated 100, 125, 150 and 200 percent scaling", SettingsScaledPanelsTest)
SettingsScaledPanelsTest() {
    app := SettingsLayoutFixture()
    try {
        ; Change only these hidden test windows, never Windows display settings.
        ; Win32 pixel sizes reproduce the scaled sidebar/unscaled X mismatch.
        app.S_Gui.Show("Hide x-500 y-200 w700 h822")
        app.Sidebar.Show("Hide x0 y0 w240 h822")
        for scale in [1, 1.25, 1.5, 2] {
            sidebarWidth := Round(app.sidebarW * scale)
            AssertTrue(DllCall("SetWindowPos", "Ptr", app.Sidebar.Hwnd, "Ptr", 0,
                "Int", 0, "Int", 0, "Int", sidebarWidth, "Int", Round(app.guiHeight * scale), "UInt", 0x14))
            sidebar := SettingsNativeRect(app.Sidebar.Hwnd, app.S_Gui.Hwnd)
            app.MainFrame.Show(Format("x{} y0 w450 h812", app.sidebarW))
            oldContent := SettingsNativeRect(app.MainFrame.Hwnd, app.S_Gui.Hwnd)
            if scale > 1
                AssertTrue(oldContent.left < sidebar.right, "Reproduce the reported leftward overlap at " (scale * 100) "%.")
            app.ShowSettingsContent()
            content := SettingsNativeRect(app.MainFrame.Hwnd, app.S_Gui.Hwnd)
            AssertEqual(sidebar.right, content.left, "No overlap or gap at " (scale * 100) "%.")
            AssertEqual(sidebar.top, content.top)
            heading := SettingsNativeRect(app.MainFrame.Group["Hotkey Groups"][1].Hwnd, app.S_Gui.Hwnd)
            AssertTrue(heading.left > sidebar.right, "The heading must stay clear of the sidebar.")
        }
    } finally {
        app.Dispose()
    }
}

TestRunner.Register("Settings content placement uses parent coordinates including sidebar offsets", SettingsPanelParentCoordinatesTest)
TestRunner.Register("Logout membership controls fit Thumbnails Interactions and update their profile settings", SettingsGroupPersistenceControlsTest)
SettingsGroupPersistenceControlsTest() {
    app := SettingsLayoutFixture()
    app.Save_Settings_Delay_Timer := (*) => 0
    try {
        app.ThumbnailsInteractions_Ctrl()
        for name in ["PreserveDisabledOnLogout", "PreserveQuickGroupOnLogout"] {
            control := app.MainFrame[name]
            control.GetPos(&x, &y, &w, &h)
            AssertTrue(y + h <= app.guiHeight - 10, "The logout controls must fit without clipping.")
            AssertEqual(1, control.Value, "Membership persistence must be enabled by default.")
            SendMessage(0xF5, 0, 0, control) ; BM_CLICK follows the real settings event path.
            Sleep(20)
            AssertEqual(0, app.%name%, "The checkbox must update its own profile setting.")
        }
    } finally {
        SetTimer(app.Save_Settings_Delay_Timer, 0)
        app.Dispose()
    }
}

SettingsPanelParentCoordinatesTest() {
    app := SettingsLayoutFixture()
    try {
        app.S_Gui.Show("Hide x-800 y-300 w700 h822")
        app.Sidebar.Show("Hide x13 y7 w240 h822")
        app.ShowSettingsContent()
        sidebar := SettingsNativeRect(app.Sidebar.Hwnd, app.S_Gui.Hwnd)
        content := SettingsNativeRect(app.MainFrame.Hwnd, app.S_Gui.Hwnd)
        AssertEqual(sidebar.right, content.left)
        AssertEqual(sidebar.top, content.top)
    } finally {
        app.Dispose()
    }
}

TestRunner.Register("Sidebar status responds to saved changes without moving UI or taking focus", SettingsSaveNoticeTest)
SettingsSaveNoticeTest() {
    app := SettingsLayoutFixture()
    path := A_Temp "\EVE-X-Preview-save-notice-" A_TickCount ".json"
    try {
        app.S_Gui.Show("NoActivate x-3000 y-3000 w700 h822")
        app.Sidebar.Show("NoActivate x0 y0 w240 h822")
        app.MainFrame.Show("NoActivate x240 y0 w450 h812")
        AssertEqual("", app.updateNotice.Text, "No save instructions are shown before an edit.")
        AssertFalse(DllCall("IsWindowVisible", "Ptr", app.updateNotice.Hwnd))
        fixedWindows := [app.S_Gui, app.Sidebar, app.MainFrame, app.firstBtn, app.testThumbnailButton, app.sidebarStatusBlock]
        originalRects := []
        for window in fixedWindows
            originalRects.Push(SettingsNativeRect(window.Hwnd, 0))
        app.SaveJsonToFile(path)
        AssertEqual("", app.updateNotice.Text)
        app.TrackClientPossitions := !app.TrackClientPossitions
        AssertThrows(() => app.SaveJsonToFile(path "\missing\settings.json"))
        AssertEqual("", app.updateNotice.Text, "A failed save must not claim success.")
        foreground := DllCall("GetForegroundWindow", "Ptr")
        app.SaveJsonToFile(path)
        AssertEqual(foreground, DllCall("GetForegroundWindow", "Ptr"), "Autosaving must not activate Settings.")
        AssertEqual(app.TrackClientPossitions, JSON.Load(FileRead(path))["_Profiles"][app.ProfileClientSettings]["Client Settings"]["TrackClientPossitions"])
        AssertEqual("Changes saved.`nClose Settings to apply.", app.updateNotice.Text)
        AssertTrue(DllCall("IsWindowVisible", "Ptr", app.updateNotice.Hwnd))
        banner := SettingsNativeRect(app.updateNotice.Hwnd, app.Sidebar.Hwnd)
        sidebar := SettingsNativeRect(app.Sidebar.Hwnd, app.S_Gui.Hwnd)
        button := SettingsNativeRect(app.testThumbnailButton.Hwnd, app.Sidebar.Hwnd)
        AssertTrue(banner.left >= 0 && banner.right <= sidebar.right - sidebar.left)
        AssertTrue(banner.top > button.bottom, "Save feedback sits below the Test Thumbnail button.")
        AssertTrue(banner.bottom <= sidebar.bottom - sidebar.top, "The notice fits inside the sidebar.")
        app.BeginSidebarUpdateStatus()
        app.SetSidebarStatus("No stable updates found.")
        app.SaveJsonToFile(path)
        AssertEqual("No stable updates found.", app.updateNotice.Text, "An unchanged save must not replace the latest action.")
        app.TrackClientPossitions := !app.TrackClientPossitions
        app.SaveJsonToFile(path)
        for index, window in fixedWindows {
            current := SettingsNativeRect(window.Hwnd, 0)
            original := originalRects[index]
            for edge in ["left", "top", "right", "bottom"]
                AssertEqual(original.%edge%, current.%edge%, "Saving must not move or resize the window, panels or buttons.")
        }
        AssertEqual("Changes saved.`nClose Settings to apply.", app.updateNotice.Text)
        app.S_Gui.Destroy()
        app.TrackClientPossitions := !app.TrackClientPossitions
        app.SaveJsonToFile(path) ; A pending save is safe after the window closes.
        app.S_Gui := Gui()
        app.CreateSidebar()
        app.CreateMainFrame()
        app.SaveJsonToFile(path)
        AssertEqual("", app.updateNotice.Text, "Reopening Settings resets the previous action.")
    } finally {
        app.Dispose()
        for file in [path, path ".save.tmp"]
            if FileExist(file)
                FileDelete(file)
    }
}
