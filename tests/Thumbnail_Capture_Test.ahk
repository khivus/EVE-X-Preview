TestRunner.Register("Capture exclusion defaults off, persists and follows behavior overrides", ThumbnailCaptureProfilesTest)
ThumbnailCaptureProfilesTest() {
    app := CombatEventsFixture()
    app.ProfileOverride()
    AssertFalse(app.ExcludeThumbnailsFromCapture)
    app._JSON["_Profiles"]["Existing"] := DeepClone(app._JSON["_Profiles"]["Default"])
    app._JSON["_Profiles"]["Existing"]["Thumbnails Behavior"].Delete("ExcludeThumbnailsFromCapture")
    app._JSON := JsonMergeNoOverwrite(JSON.Load(default_JSON), app._JSON)
    app.LastUsedProfile := "Existing"
    app.ProfileOverride()
    AssertFalse(app.ExcludeThumbnailsFromCapture)
    app.ExcludeThumbnailsFromCapture := true
    app._JSON := JSON.Load(JSON.Dump(app._JSON))
    AssertTrue(app.ExcludeThumbnailsFromCapture)
    AssertFalse(app._JSON["_Profiles"]["Default"]["Thumbnails Behavior"]["ExcludeThumbnailsFromCapture"])
    app.Global_Groups["Thumbnails Behavior"] := 1
    app.AntiGlobalGroups["Thumbnails Behavior"] := 0
    app.ProfileOverride()
    AssertFalse(app.ExcludeThumbnailsFromCapture)
    app.ExcludeThumbnailsFromCapture := true
    AssertTrue(app._JSON["_Profiles"]["Default"]["Thumbnails Behavior"]["ExcludeThumbnailsFromCapture"])
    app.AntiGlobalGroups["Thumbnails Behavior"] := 1
    app.ProfileOverride()
    app.ExcludeThumbnailsFromCapture := false
    AssertFalse(app._JSON["_Profiles"]["Existing"]["Thumbnails Behavior"]["ExcludeThumbnailsFromCapture"])
    AssertTrue(app._JSON["_Profiles"]["Default"]["Thumbnails Behavior"]["ExcludeThumbnailsFromCapture"])
}

TestRunner.Register("Capture affinity covers thumbnail, text and both borders and can be cleared", ThumbnailCaptureWindowsTest)
ThumbnailCaptureWindowsTest() {
    oldHidden := A_DetectHiddenWindows
    DetectHiddenWindows(true)
    try {
        for enabled in [false, true, false] {
            app := ThumbnailBorderFixture(1, enabled)
            try {
                for key in ["Window", "TextOverlay", "Border", "EventBorder"] {
                    affinity := 0
                    AssertTrue(DllCall("user32\GetWindowDisplayAffinity", "Ptr", app.thumb[key].Hwnd, "UInt*", &affinity, "Int"), "Read affinity for " key)
                    AssertEqual(enabled ? 0x11 : 0, affinity, "Creation affinity for " key)
                    app.ExcludeThumbnailsFromCapture := false
                    AssertTrue(app.ApplyThumbnailCaptureAffinity(app.thumb[key].Hwnd))
                    AssertTrue(DllCall("user32\GetWindowDisplayAffinity", "Ptr", app.thumb[key].Hwnd, "UInt*", &affinity, "Int"))
                    AssertEqual(0, affinity, "Clear affinity for " key)
                    app.ExcludeThumbnailsFromCapture := enabled
                }
            } finally app.Dispose()
        }
    } finally DetectHiddenWindows(oldHidden)
}

TestRunner.Register("Capture checkbox uses behavior save/reload and fits the settings panel", ThumbnailCaptureControlTest)
ThumbnailCaptureControlTest() {
    app := ColorSettingsFixture()
    try {
        app.ThumbnailsBehavior_Ctrl()
        control := app.MainFrame["ExcludeThumbnailsFromCapture"]
        AssertFalse(control.Value)
        control.Value := true
        id := DllCall("GetDlgCtrlID", "Ptr", control.Hwnd, "Int")
        PostMessage(0x111, id & 0xFFFF, control.Hwnd,, "ahk_id " app.MainFrame.Hwnd)
        Sleep(250)
        AssertTrue(app.ExcludeThumbnailsFromCapture)
        AssertTrue(app.NeedRestart)
        AssertEqual(1, app.saves)
        for item in app.MainFrame.Group["Thumbnails Behavior"] {
            item.GetPos(&x, &y, &w, &h)
            AssertTrue(x + w <= app.contentW)
            AssertTrue(y + h <= app.guiHeight - 10)
        }
    } finally app.Dispose()
}
