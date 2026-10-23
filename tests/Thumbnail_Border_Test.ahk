class ThumbnailBorderFixture extends PreviewLifecycleFixture {
    focused := true
    __New(mode := 1, excludeCapture := false) {
        super.__New()
        this.ExcludeThumbnailsFromCapture := excludeCapture
        this.gameLogsMonitoringEnabled := true
        this.flashBorderEnabled := true
        this.EventActiveBorderMode := mode
        this.ShowAllColoredBorders := false
        this.ShowClientHighlightBorder := true
        this.ClientHighligtColor := "ABCDEF"
        this.ClientHighligtBorderthickness := 4
        this.InactiveClientBorderColor := "778899"
        this.InactiveClientBorderthickness := 2
        this.stopDisplayingOnSwitch := false
        this.supressForFocused := false
        this.eventMethods := Map()
        this.flashMethod := Map()
        this.LastActiveThumbHwnd := 0
        this.DisabledFromGroupsEnabled := false
        this.QuickGroupEnabled := false
        this.title := "EVE - Border Test"
        this.monitoredEvents["warpDisrupted"]["enabled"] := true
        this.source := Gui(, this.title)
        this.source.Show("Hide w320 h200")
        this.key := this.source.Hwnd
        this.ThumbWindows.%this.key% := this.Create_Thumbnail(this.key, this.title)
        this.thumb := this.ThumbWindows.%this.key%
        this.ThumbHwnd_EvEHwnd[this.thumb["Window"].Hwnd] := this.key
        this.ShowThumb(this.key, "Show")
        this.monitoredEvents["warpDisrupted"]["color"] := "123456"
        this.flashMethod[this.title] := Map("event", "warpDisrupted", "isOn", false, "flashMethod", (*) => 0)
        this.eventMethods[this.title] := (*) => 0
    }
    EventSourceWindow(title) => this.key
    EventSourceActive(hwnd) => this.focused
    Dispose() {
        for callback in this.eventMethods
            SetTimer(this.eventMethods[callback], 0)
        for title, state in this.flashMethod
            SetTimer(state["flashMethod"], 0)
        SetTimer(this.debugToolTipMethod, 0)
        this.DisposePreview(this.key)
        this.source.Destroy()
    }
}

BorderRegionContains(gui, x, y) {
    region := DllCall("gdi32\CreateRectRgn", "Int", 0, "Int", 0, "Int", 0, "Int", 0, "Ptr")
    try {
        AssertTrue(DllCall("GetWindowRgn", "Ptr", gui.Hwnd, "Ptr", region, "Int") > 0)
        return !!DllCall("gdi32\PtInRegion", "Ptr", region, "Int", x, "Int", y, "Int")
    } finally DllCall("gdi32\DeleteObject", "Ptr", region)
}

TestRunner.Register("Default and swapped event borders occupy distinct rings while active highlighting stays visible", ThumbnailDoubleBorderTest)
ThumbnailDoubleBorderTest() {
    oldHidden := A_DetectHiddenWindows
    DetectHiddenWindows(true)
    try {
        for mode in [1, 2] {
            app := ThumbnailBorderFixture(mode)
            try {
                app.DrawFlashingBorder(app.title)
                AssertTrue(DllCall("IsWindowVisible", "Ptr", app.thumb["Border"].Hwnd))
                AssertTrue(DllCall("IsWindowVisible", "Ptr", app.thumb["EventBorder"].Hwnd))
                AssertEqual("ABCDEF", app.thumb["Border"].BackColor)
                AssertEqual("123456", app.thumb["EventBorder"].BackColor)
                AssertEqual(0x80020, WinGetExStyle(app.thumb["EventBorder"].Hwnd) & 0x80020, "The event ring must stay layered and click-through.")
                AssertEqual(mode = 1, BorderRegionContains(app.thumb["Border"], 1, 20))
                AssertEqual(mode = 2, BorderRegionContains(app.thumb["Border"], 5, 20))
                AssertEqual(mode = 2, BorderRegionContains(app.thumb["EventBorder"], 1, 20))
                AssertEqual(mode = 1, BorderRegionContains(app.thumb["EventBorder"], 5, 20))
                AssertFalse(BorderRegionContains(app.thumb["EventBorder"], 9, 20))
                app.DrawFlashingBorder(app.title)
                AssertFalse(DllCall("IsWindowVisible", "Ptr", app.thumb["EventBorder"].Hwnd))
                AssertTrue(DllCall("IsWindowVisible", "Ptr", app.thumb["Border"].Hwnd), "An off event flash must leave the active ring visible.")
                AssertEqual("ABCDEF", app.thumb["Border"].BackColor)
                app.DrawFlashingBorder(app.title)
                app.endEvent(app.title, app.key)
                AssertFalse(app.thumb["EventBorder"].EventVisible)
                AssertTrue(DllCall("IsWindowVisible", "Ptr", app.thumb["Border"].Hwnd), "Event expiry must retain active highlighting.")
            } finally app.Dispose()
        }
    } finally DetectHiddenWindows(oldHidden)
}

TestRunner.Register("Combined mode retains one shared border and inactive event priority", ThumbnailCombinedBorderTest)
ThumbnailCombinedBorderTest() {
    oldHidden := A_DetectHiddenWindows
    DetectHiddenWindows(true)
    app := ThumbnailBorderFixture(3)
    try {
        AssertFalse(app.thumb.Has("EventBorder"))
        app.activateBorder(app.key, app.title)
        app.focused := false
        app.DrawFlashingBorder(app.title)
        AssertEqual("123456", app.thumb["Border"].BackColor)
        AssertTrue(DllCall("IsWindowVisible", "Ptr", app.thumb["Border"].Hwnd))
        app.DrawFlashingBorder(app.title)
        AssertEqual("123456", app.thumb["Border"].BackColor, "Combined mode keeps the previous event-priority behavior.")
    } finally {
        app.Dispose()
        DetectHiddenWindows(oldHidden)
    }
}

TestRunner.Register("Separate event borders respect focus suppression, hidden previews and inactive/group borders", ThumbnailBorderVisibilityTest)
ThumbnailBorderVisibilityTest() {
    oldHidden := A_DetectHiddenWindows
    DetectHiddenWindows(true)
    app := ThumbnailBorderFixture()
    try {
        app.DrawFlashingBorder(app.title)
        app.supressForFocused := true
        app.DrawFlashingBorder(app.title)
        AssertFalse(app.thumb["EventBorder"].EventVisible)
        AssertEqual("ABCDEF", app.thumb["Border"].BackColor)
        app.supressForFocused := false
        app.focused := false
        app.ShowAllColoredBorders := true
        app.DrawFlashingBorder(app.title)
        AssertEqual("778899", app.thumb["Border"].BackColor)
        AssertTrue(DllCall("IsWindowVisible", "Ptr", app.thumb["Border"].Hwnd))
        app.SetEventBorder(app.key, "123456", true)
        app.ShowThumb(app.key, "Hide")
        app.SetEventBorder(app.key, "123456", true)
        AssertFalse(DllCall("IsWindowVisible", "Ptr", app.thumb["EventBorder"].Hwnd), "An event timer must not reveal an independently hidden border.")
        app.ShowThumb(app.key, "Show")
        AssertTrue(DllCall("IsWindowVisible", "Ptr", app.thumb["EventBorder"].Hwnd))
        app.SetThumbnailText[app.key] := "EVE - Another Pilot"
        AssertFalse(app.thumb["EventBorder"].EventVisible, "A changed character must not inherit the previous character's event border.")
        app.DisabledFromGroupsEnabled := true
        app.DisabledChars[app.key] := app.title
        app.clearBorder(app.key, app.title)
        AssertEqual(app.DisableFromGroupsColor, app.thumb["Border"].BackColor)
        app.ToggleLivePreviews()
        AssertFalse(DllCall("IsWindowVisible", "Ptr", app.thumb["EventBorder"].Hwnd))
    } finally {
        app.Dispose()
        DetectHiddenWindows(oldHidden)
    }
}

TestRunner.Register("Both border layers move, resize, clip on tiny thumbnails and release on cleanup", ThumbnailBorderGeometryTest)
ThumbnailBorderGeometryTest() {
    oldHidden := A_DetectHiddenWindows
    DetectHiddenWindows(true)
    baseline := LiveThumb.OBJ_COUNTER
    app := ThumbnailBorderFixture()
    try {
        app.DrawFlashingBorder(app.title)
        windows := app.CaptureThumbnailDragWindows(app.key, false)
        AssertEqual(4, windows.Count)
        app.MoveThumbnailDragWindows(windows, 30, 40)
        for hwnd, position in windows {
            WinGetPos(&x, &y,,, hwnd)
            AssertEqual([position.x + 30, position.y + 40], [x, y])
        }
        app.ThumbMove(-100, 50, 200, 120, app.thumb)
        for key in ["Window", "Border", "EventBorder", "TextOverlay"] {
            WinGetPos(&x, &y, &w, &h, app.thumb[key].Hwnd)
            AssertEqual([-100, 50, 200, 120], [x, y, w, h])
        }
        AssertTrue(BorderRegionContains(app.thumb["EventBorder"], 5, 20))
        app.ThumbMove(20, 30, 3, 3, app.thumb)
        AssertFalse(BorderRegionContains(app.thumb["EventBorder"], 1, 1), "Insufficient space must produce an empty inner ring, rather than inverted geometry.")
        eventHwnd := app.thumb["EventBorder"].Hwnd
        app.DisposePreview(app.key)
        AssertFalse(DllCall("IsWindow", "Ptr", eventHwnd))
        AssertEqual(baseline, LiveThumb.OBJ_COUNTER)
    } finally {
        app.Dispose()
        DetectHiddenWindows(oldHidden)
    }
}

TestRunner.Register("Event border mode defaults for old profiles, persists and follows visuals inheritance", ThumbnailBorderSettingsTest)
ThumbnailBorderSettingsTest() {
    app := CombatEventsFixture()
    app.ProfileOverride()
    AssertEqual(1, app.EventActiveBorderMode)
    app.Profiles["Default"]["Thumbnails Visuals"].Delete("EventActiveBorderMode")
    AssertEqual(1, app.EventActiveBorderMode)
    for mode in [1, 2, 3] {
        app.EventActiveBorderMode := mode
        app._JSON := JSON.Load(JSON.Dump(app._JSON))
        AssertEqual(mode, app.EventActiveBorderMode)
    }
    app.Profiles["Other"] := JSON.Load(JSON.Dump(app.Profiles["Default"]))
    app.Profiles["Other"]["Thumbnails Visuals"]["EventActiveBorderMode"] := 2
    app.Profiles["Default"]["Other"]["Global_Groups"]["Thumbnails Visuals"] := 1
    app.Profiles["Other"]["Other"]["Global_Groups"]["Thumbnails Visuals"] := 0
    app.LastUsedProfile := "Other"
    app.ProfileOverride()
    AssertEqual(3, app.EventActiveBorderMode)
    app.Profiles["Other"]["Other"]["Global_Groups"]["Thumbnails Visuals"] := 1
    app.ProfileOverride()
    AssertEqual(2, app.EventActiveBorderMode)
}

TestRunner.Register("Border layout control saves selections and stays within the visuals panel", ThumbnailBorderSettingsControlTest)
ThumbnailBorderSettingsControlTest() {
    app := ColorSettingsFixture()
    try {
        control := app.MainFrame["EventActiveBorderMode"]
        control.GetPos(&x, &y, &w, &h)
        AssertTrue(x + w <= app.contentW)
        AssertTrue(control.ToolTip != "")
        control.Choose(2)
        id := DllCall("GetDlgCtrlID", "Ptr", control.Hwnd, "Int")
        PostMessage(0x111, (1 << 16) | (id & 0xFFFF), control.Hwnd,, "ahk_id " app.MainFrame.Hwnd)
        Sleep(30)
        AssertEqual(2, app.EventActiveBorderMode)
        AssertTrue(app.NeedRestart)
        for item in app.MainFrame.Group["Thumbnails Visuals"] {
            item.GetPos(&x, &y, &w, &h)
            AssertTrue(y + h <= app.guiHeight - 10, "The extra setting must not push visual controls outside the panel.")
        }
    } finally app.Dispose()
}

TestRunner.Register("Disabled monitoring, flashing or all event types keep a single border", ThumbnailBorderCreationGatesTest)
ThumbnailBorderCreationGatesTest() {
    for disabled in ["monitoring", "flashing", "events"] {
        app := SettingsThumbnailFixture()
        try {
            app.EventActiveBorderMode := 2
            app.gameLogsMonitoringEnabled := disabled != "monitoring"
            app.flashBorderEnabled := disabled != "flashing"
            app.monitoredEvents["warpDisrupted"]["enabled"] := disabled != "events"
            preview := app.OpenPreview()
            AssertFalse(preview.sample.Has("EventBorder"))
            AssertTrue(BorderRegionContains(preview.sample["Border"], 1, 20), "A single border must stay at the thumbnail edge without an unused outer slot.")
            preview.sample["TextOverlay"]["OverlayText"].GetPos(&x, &y)
            AssertEqual([5, 5], [x, y], "Single-border setups must retain the configured text margins.")
        } finally app.Dispose()
    }
}

TestRunner.Register("Double borders add only the inner ring width to all text margins without accumulating or changing settings", ThumbnailBorderTextMarginsTest)
ThumbnailBorderTextMarginsTest() {
    oldHidden := A_DetectHiddenWindows
    DetectHiddenWindows(true)
    try {
        for mode in [1, 2, 3] {
            app := ThumbnailBorderFixture(mode)
            try {
                app.ShowAllColoredBorders := true
                app.focused := false
                overlay := app.thumb["TextOverlay"]
                overlay.SystemTrackingEnabled := true
                overlay.MarginX := 5
                overlay.MarginY := 6
                app.ThumbnailTextMargins["x"] := 5
                app.ThumbnailTextMargins["y"] := 6
                app.ThumbMove(20, 30, 260, 220, app.thumb)
                app.updateThumbnailSystemText("Jita", app.key)
                app.updateThumbnailDPSText("DPS: In: 2500 | Out: 1800", app.key, "% EM 50 Th 50")
                app.updateThumbnailEventText("Warp Disrupted", app.key)
                Sleep(20)
                before := JSON.Dump(app._JSON)
                app.DrawFlashingBorder(app.title)
                extra := mode = 1 ? 4 : (mode = 2 ? 2 : 0)
                mx := 5 + extra / (A_ScreenDPI / 96)
                my := 6 + extra / (A_ScreenDPI / 96)
                for name in ["OverlayText", "EventText", "DPSText", "DPSDamageText", "SystemText"] {
                    overlay[name].GetPos(&x,, &w)
                    AssertTrue(Abs(x - mx) <= 1, name " must use the extra inner-border margin.")
                    AssertTrue(Abs(w - (overlay.TextLayoutWidth - 2 * mx)) <= 1)
                }
                overlay["OverlayText"].GetPos(, &y)
                AssertTrue(Abs(y - my) <= 1)
                overlay["SystemText"].GetPos(, &sy,, &sh)
                AssertTrue(Abs(sy + sh - (overlay.TextLayoutHeight - my)) <= 1, "Bottom text must clear the inner border too.")
                Loop 4
                    app.DrawFlashingBorder(app.title)
                overlay["OverlayText"].GetPos(&x, &y)
                AssertTrue(Abs(x - mx) <= 1 && Abs(y - my) <= 1, "Flashing must not accumulate padding or shift the text.")
                AssertEqual([5, 6], [overlay.MarginX, overlay.MarginY])
                AssertEqual(before, JSON.Dump(app._JSON))
                if mode = 2 {
                    app.activateBorder(app.key, app.title)
                    overlay["OverlayText"].GetPos(&x)
                    AssertTrue(Abs(x - (5 + 4 / (A_ScreenDPI / 96))) <= 1, "A thicker active inner border must update the extra margin.")
                }
                app.ThumbMove(20, 30, 3, 3, app.thumb)
                Sleep(20)
                for name in ["OverlayText", "SystemText", "DPSText", "DPSDamageText", "EventText"] {
                    overlay[name].GetPos(&x, &y, &w, &h)
                    AssertTrue(x >= 0 && y >= 0 && w >= 0 && h >= 0)
                }
            } finally app.Dispose()
        }
    } finally DetectHiddenWindows(oldHidden)
}

TestRunner.Register("Test thumbnail demonstrates both border modes and follows appearance/geometry edits", ThumbnailBorderDemoTest)
ThumbnailBorderDemoTest() {
    app := SettingsThumbnailFixture()
    try {
        app.gameLogsMonitoringEnabled := true
        app.monitoredEvents["warpDisrupted"]["enabled"] := true
        preview := app.OpenPreview()
        Loop 4
            preview.NextSample()
        AssertTrue(preview.sample.Has("EventBorder"))
        AssertTrue(DllCall("IsWindowVisible", "Ptr", preview.sample["Border"].Hwnd))
        AssertTrue(DllCall("IsWindowVisible", "Ptr", preview.sample["EventBorder"].Hwnd))
        AssertTrue(BorderRegionContains(preview.sample["EventBorder"], 5, 20))
        preview.PaintSample(preview.sampleSince + preview.flashBorderInterval + 1)
        AssertFalse(DllCall("IsWindowVisible", "Ptr", preview.sample["EventBorder"].Hwnd))
        AssertTrue(DllCall("IsWindowVisible", "Ptr", preview.sample["Border"].Hwnd))
        app.EventActiveBorderMode := 2
        preview.Refresh(A_TickCount + 500)
        AssertTrue(BorderRegionContains(preview.sample["EventBorder"], 1, 20))
        preview.ApplyGeometry(100, 150, 260, 180)
        preview.PaintSample(A_TickCount)
        AssertSampleGeometry(preview, [100, 150, 260, 180])
        app.EventActiveBorderMode := 3
        preview.Refresh(A_TickCount + 1000)
        AssertFalse(preview.sample.Has("EventBorder"))
    } finally app.Dispose()
}
