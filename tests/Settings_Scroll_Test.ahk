class SettingsScrollFixture extends SettingsLayoutFixture {
    __New() {
        super.__New()
        this.topControl := this.MainFrame.AddEdit("x16 y16 w300 vScrollTop", "Top field")
        this.nativeEdit := this.MainFrame.AddEdit("x16 y80 w300 h100 Multi vScrollNative", "")
        Loop 80
            this.nativeEdit.Value .= "Native line " A_Index "`r`n"
        this.bottomControl := this.MainFrame.AddCheckbox("x16 y1600 w300 vScrollBottom", "Last setting")
        this.bottomButton := this.MainFrame.AddButton("x16 y1800 w200 vScrollAction", "Last action")
        this.shortControl := this.MainFrame.AddText("x16 y16 w200", "Short category")
        this.MainFrame.Group["Long"] := [this.topControl, this.nativeEdit, this.bottomControl, this.bottomButton]
        this.MainFrame.Group["Short"] := [this.shortControl]
        this.S_Gui.Show("NoActivate x-3000 y-3000 w700 h430")
        this.Sidebar.Show("NoActivate x0 y0 w240 h430")
        this.settingsScroller.Show(Round(240 * A_ScreenDPI / 96), 0, 450, 300)
        this.Select("Long")
    }
    Select(name) {
        for category, controls in this.MainFrame.Group
            for control in controls
                control.Visible := category = name
        this.settingsScroller.Select(this.MainFrame.Group[name])
    }
    Wheel(delta, inside := true) {
        rect := SettingsNativeRect(this.settingsScroller.viewportHwnd, 0)
        x := inside ? rect.left + 5 : rect.left - 10
        y := rect.top + 5
        SendMessage(0x20A, (delta & 0xFFFF) << 16, (x & 0xFFFF) | ((y & 0xFFFF) << 16), this.settingsScroller.viewportHwnd)
    }
}

TestRunner.Register("Native settings scrollbar reaches final controls without moving sidebar or changing control coordinates", SettingsNativeScrollTest)
SettingsNativeScrollTest() {
    app := SettingsScrollFixture()
    try {
        scroll := app.settingsScroller
        sidebar := SettingsNativeRect(app.Sidebar.Hwnd, 0)
        app.bottomButton.GetPos(, &originalY)
        AssertTrue(scroll.maximum > 0)
        SendMessage(0x115, 7, 0, scroll.viewportHwnd) ; SB_BOTTOM
        AssertEqual(scroll.maximum, scroll.offset)
        button := SettingsNativeRect(app.bottomButton.Hwnd, scroll.viewportHwnd)
        AssertTrue(button.top >= 0 && button.bottom <= scroll.page)
        app.bottomButton.GetPos(, &actualY)
        AssertEqual(originalY, actualY, "Scrolling the content must not rewrite child geometry.")
        current := SettingsNativeRect(app.Sidebar.Hwnd, 0)
        AssertEqual(sidebar.top, current.top)
        AssertEqual(sidebar.bottom, current.bottom)
        SendMessage(0x115, 6, 0, scroll.viewportHwnd) ; SB_TOP
        AssertEqual(0, scroll.offset)
        SendMessage(0x115, 1, 0, scroll.viewportHwnd) ; SB_LINEDOWN
        AssertTrue(scroll.offset > 0)
        SendMessage(0x115, 3, 0, scroll.viewportHwnd) ; SB_PAGEDOWN
        AssertTrue(scroll.offset > scroll.page - 50 * A_ScreenDPI / 96)
    } finally app.Dispose()
}

TestRunner.Register("Wheel input scrolls only over the right settings pane and accumulates partial notches", SettingsWheelScrollTest)
SettingsWheelScrollTest() {
    app := SettingsScrollFixture()
    try {
        scroll := app.settingsScroller
        app.Wheel(-60)
        AssertEqual(0, scroll.offset)
        app.Wheel(-60)
        AssertTrue(scroll.offset > 0)
        previous := scroll.offset
        app.Wheel(-120, false)
        AssertEqual(previous, scroll.offset, "Wheel over the sidebar must not scroll the settings page.")
        AssertTrue(scroll.NativeWheelTarget(app.nativeEdit.Hwnd), "A populated multiline edit retains its own scrollbar input.")
    } finally app.Dispose()
}

TestRunner.Register("Keyboard focus reveals offscreen fields and category changes reset scroll range", SettingsScrollFocusTest)
SettingsScrollFocusTest() {
    app := SettingsScrollFixture()
    try {
        scroll := app.settingsScroller
        app.bottomControl.Focus()
        Sleep(20)
        AssertTrue(scroll.offset > 0, "The native Focus event must bring the tabbed-to field into view.")
        control := SettingsNativeRect(app.bottomControl.Hwnd, scroll.viewportHwnd)
        AssertTrue(control.top >= 0 && control.bottom <= scroll.page)
        app.Select("Short")
        AssertEqual(0, scroll.offset)
        AssertEqual(0, scroll.maximum)
        AssertFalse(WinGetStyle(scroll.viewportHwnd) & 0x200000, "Short pages should not need a scrollbar.")
        app.Select("Long")
        AssertEqual(0, scroll.offset)
        AssertTrue(scroll.maximum > 0)
    } finally app.Dispose()
    AssertTrue(scroll.stopped, "Destroying the settings GUI must unregister native message handlers.")
}

TestRunner.Register("Anonymous settings migrate to visuals without replacing explicit visuals values", AnonymousSettingsCategoryMigrationTest)
AnonymousSettingsCategoryMigrationTest() {
    settings := JSON.Load(default_JSON)
    profile := settings["_Profiles"]["Default"]
    for key in ["AnonymizeThumbnails", "AnonymousNames"]
        profile["Thumbnails Visuals"].Delete(key)
    profile["Thumbnails Behavior"]["AnonymizeThumbnails"] := true
    profile["Thumbnails Behavior"]["AnonymousNames"] := "Preserved pool"
    MigrateThumbnailNameSettings(settings)
    settings := JsonMergeNoOverwrite(JSON.Load(default_JSON), settings)
    visuals := settings["_Profiles"]["Default"]["Thumbnails Visuals"]
    AssertTrue(visuals["AnonymizeThumbnails"])
    AssertEqual("Preserved pool", visuals["AnonymousNames"])
    AssertFalse(settings["_Profiles"]["Default"]["Thumbnails Behavior"].Has("AnonymousNames"))
    settings["_Profiles"]["Default"]["Thumbnails Behavior"]["AnonymousNames"] := "Old pool"
    MigrateThumbnailNameSettings(settings)
    AssertEqual("Preserved pool", visuals["AnonymousNames"])
}

TestRunner.Register("Actual hotkey behavior and visuals pages can scroll to every final setting", SettingsCategoryScrollReachabilityTest)
SettingsCategoryScrollReachabilityTest() {
    app := ColorSettingsFixture()
    try {
        app.ThumbnailsBehavior_Ctrl()
        app.HotkeyGroups_Ctrl()
        app.S_Gui.Show("NoActivate x-3000 y-3000 w700 h500")
        app.settingsScroller.Show(0, 0, 450, 350)
        for category in ["Hotkey Groups", "Thumbnails Behavior", "Thumbnails Visuals"] {
            controls := app.MainFrame.Group[category]
            for other, items in app.MainFrame.Group
                for control in items
                    control.Visible := other = category
            app.settingsScroller.Select(controls)
            app.settingsScroller.ScrollTo(app.settingsScroller.maximum)
            lowest := 0
            for control in controls {
                rect := SettingsNativeRect(control.Hwnd, app.settingsScroller.viewportHwnd)
                lowest := Max(lowest, rect.bottom)
            }
            AssertTrue(app.settingsScroller.maximum > 0, category " must have a scroll range in a short viewport.")
            AssertTrue(lowest <= app.settingsScroller.page, "The final controls of " category " must be reachable.")
        }
    } finally app.Dispose()
}

TestRunner.Register("Scrollbar outer rectangle fits the settings window and heading stays pinned", SettingsPinnedHeadingBoundsTest)
SettingsPinnedHeadingBoundsTest() {
    app := ColorSettingsFixture()
    try {
        app.S_Gui.Show("NoActivate x-3000 y-3000 w700 h500")
        scroll := app.settingsScroller
        scroll.Show(Round(240 * A_ScreenDPI / 96), 0, 450, 350)
        for category in ["Thumbnails Visuals", "Thumbnails Interactions"] {
            for name, items in app.MainFrame.Group
                for control in items
                    control.Visible := name = category
            scroll.Select(app.MainFrame.Group[category])
            viewport := SettingsNativeRect(scroll.viewportHwnd, app.S_Gui.Hwnd)
            parent := Buffer(16, 0)
            DllCall("GetClientRect", "Ptr", app.S_Gui.Hwnd, "Ptr", parent)
            AssertTrue(viewport.right <= NumGet(parent, 8, "Int"), "The entire scrollbar must fit within the parent client area.")
            AssertEqual(Round(450 * A_ScreenDPI / 96), viewport.right - viewport.left, "Requested width includes the scrollbar.")
            heading := SettingsNativeRect(scroll.heading.Hwnd, app.S_Gui.Hwnd)
            AssertEqual(app.MainFrame.Group[category][1].Text, scroll.heading.Text)
            scroll.ScrollTo(scroll.maximum)
            after := SettingsNativeRect(scroll.heading.Hwnd, app.S_Gui.Hwnd)
            AssertEqual(heading.top, after.top)
            AssertEqual(heading.bottom, after.bottom)
            AssertTrue(after.bottom <= viewport.top, "The heading remains above the scrolling body.")
        }
    } finally app.Dispose()
}

TestRunner.Register("About buttons fit without a scrollbar", SettingsAboutButtonsTest)
SettingsAboutButtonsTest() {
    app := SettingsLayoutFixture()
    try {
        app.About_Ctrl()
        app.S_Gui.Show("NoActivate x-3000 y-3000 w700 h822")
        app.Sidebar.Show("NoActivate x0 y0 w240 h822")
        app.ShowSettingsContent()
        for category, controls in app.MainFrame.Group
            for control in controls
                control.Visible := category = "About"
        scroll := app.settingsScroller
        scroll.Select(app.MainFrame.Group["About"])
        AssertEqual(0, scroll.maximum, "Bottom-right About buttons must not introduce vertical overflow.")
        for name in ["debugModeBtn", "funnyBtn"] {
            button := SettingsNativeRect(app.MainFrame[name].Hwnd, scroll.viewportHwnd)
            rect := Buffer(16, 0)
            DllCall("GetClientRect", "Ptr", scroll.viewportHwnd, "Ptr", rect)
            AssertTrue(button.right <= NumGet(rect, 8, "Int") && button.bottom <= scroll.page)
        }
    } finally app.Dispose()
}

TestRunner.Register("New settings preserve existing row offsets gaps and category ordering", SettingsNewRowsStyleTest)
SettingsNewRowsStyleTest() {
    app := ColorSettingsFixture()
    try {
        app.HotkeyGroups_Ctrl()
        app.ThumbnailsBehavior_Ctrl()
        for name in ["AutoWildcardWindowHotkeys", "LockThumbnailPositions", "ThumbnailPollingInterval", "AnonymizeThumbnails"] {
            control := app.MainFrame[name]
            control.GetPos(&x, &y, &w)
            AssertEqual(app.contentGap + app.offsetX, x, name " must use the original input column.")
            if control.Type = "CheckBox"
                AssertEqual("On/Off", control.Text)
        }
        app.MainFrame["AutoWildcardWindowHotkeys"].GetPos(, &starY)
        app.MainFrame["PreserveHotkeysOnLogout"].GetPos(, &preserveY)
        AssertEqual(app.xlGap, preserveY - starY)
        app.MainFrame["LockThumbnailPositions"].GetPos(, &lockY)
        app.MainFrame["ThumbnailSnap"].GetPos(, &snapY)
        AssertEqual(app.xlGap, snapY - lockY)
        AssertEqual("Hotkey Groups", app.MainFrame["AutoWildcardWindowHotkeys"].SettingsCategory)
    } finally app.Dispose()
}
