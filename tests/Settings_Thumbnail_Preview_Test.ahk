class SettingsThumbnailFixture extends SettingsLayoutFixture {
    OpenPreview() {
        this.settingsThumbnail := SettingsThumbnailPreview(this)
        SetTimer(this.settingsThumbnail.timer, 0) ; Advance sample time explicitly in tests.
        return this.settingsThumbnail
    }
    Dispose() {
        this.StopTestThumbnail()
        super.Dispose()
    }
}

class SettingsThumbnailGesturePreview extends SettingsThumbnailPreview {
    point := {x: 100, y: 100, buttons: 0}
    geometryUpdates := 0
    PointerState() => this.point
    BeginGesture(action, flags) {
        super.BeginGesture(action, flags)
        SetTimer(this.gestureTimer, 0) ; Supply deterministic pointer samples.
    }
    ApplyGeometry(args*) {
        this.geometryUpdates++
        super.ApplyGeometry(args*)
    }
}

; Inject mouse callbacks at deterministic lifecycle boundaries, with real GUI windows.
class SettingsThumbnailDestroyProbe {
    __New(preview, ownerHwnd) {
        this.preview := preview
        this.window := preview.sample["Window"]
        this.handles := [this.window.Hwnd, this.window["SampleStatus"].Hwnd, ownerHwnd]
        this.errors := []
    }
    Hwnd => this.window.Hwnd
    Destroy() {
        this.window.Destroy()
        this.detached := !IsObject(this.preview.sample)
        for hwnd in this.handles
            for message in SettingsThumbnailPreview.MouseMessages
                try this.preview.SampleMouseMessage(1, 0, message, hwnd)
                catch as err
                    this.errors.Push(err.Message)
    }
}

class SettingsThumbnailConstructionProbe extends SettingsThumbnailPreview {
    errors := []
    clicksIgnored := true
    ApplyGeometry(args*) {
        step := this.sampleStep
        for hwnd in [this.sample["Window"].Hwnd, this.sample["Window"]["SampleStatus"].Hwnd]
            for message in SettingsThumbnailPreview.MouseMessages
                try this.SampleMouseMessage(1, 0, message, hwnd)
                catch as err
                    this.errors.Push(err.Message)
        this.clicksIgnored := this.clicksIgnored && step = this.sampleStep
        super.ApplyGeometry(args*)
    }
}

TestRunner.Register("Event checkbox rebuilds ignore mouse callbacks after destroying the old test thumbnail", SettingsThumbnailRebuildMouseTest)
SettingsThumbnailRebuildMouseTest() {
    app := SettingsThumbnailFixture()
    try {
        app.gameLogsMonitoringEnabled := true
        app.showEventText := true
        preview := app.OpenPreview()
        now := A_TickCount
        for enabled in [1, 0, 1, 0] {
            probe := SettingsThumbnailDestroyProbe(preview, app.S_Gui.Hwnd)
            preview.sample["Window"] := probe
            app.monitoredEvents["warpDisrupted"]["enabled"] := enabled
            preview.Refresh(now += 500)
            AssertEqual(0, probe.errors.Length, "Queued mouse callbacks must not read destroyed GUI objects.")
            AssertTrue(probe.detached, "Detach the old sample before Destroy can dispatch messages.")
            AssertFalse(preview.closed)
            AssertEqual(enabled, preview.monitoredEvents["warpDisrupted"]["enabled"])
            step := preview.sampleStep
            preview.SampleMouseMessage(0, 0, 0x208, preview.sample["Window"]["SampleStatus"].Hwnd)
            AssertEqual(step + 1, preview.sampleStep, "The replacement must still respond to sample cycling.")
        }
    } finally {
        app.Dispose()
    }
}

TestRunner.Register("Test-thumbnail mouse gestures wait until construction and rebuild finish", SettingsThumbnailConstructionMouseTest)
SettingsThumbnailConstructionMouseTest() {
    app := SettingsThumbnailFixture()
    try {
        app.settingsThumbnail := SettingsThumbnailConstructionProbe(app)
        preview := app.settingsThumbnail
        SetTimer(preview.timer, 0)
        app.ThumbnailTextColor := "123456"
        preview.Refresh(A_TickCount + 500)
        AssertEqual(0, preview.errors.Length)
        AssertTrue(preview.clicksIgnored)
        step := preview.sampleStep
        preview.SampleMouseMessage(0, 0, 0x208, preview.sample["Window"].Hwnd)
        AssertEqual(step + 1, preview.sampleStep)
    } finally {
        app.Dispose()
    }
}

OpenGesturePreview(app) {
    app.settingsThumbnail := SettingsThumbnailGesturePreview(app)
    SetTimer(app.settingsThumbnail.timer, 0)
    return app.settingsThumbnail
}

AssertSampleGeometry(preview, expected) {
    for key in ["Window", "TextOverlay", "Border", "EventBorder"] {
        if !preview.sample.Has(key)
            continue
        WinGetPos(&x, &y, &w, &h, preview.sample[key].Hwnd)
        AssertEqual(expected, [x, y, w, h], key " must move/resize with the complete test thumbnail.")
    }
}

TestRunner.Register("Configured test-thumbnail drag moves every layer without cycling or saving and survives appearance edits", SettingsThumbnailDragTest)
TestRunner.Register("Test-thumbnail drag ignores inherited window delays and restores the caller", SettingsThumbnailDragDelayTest)
SettingsThumbnailDragDelayTest() {
    app := SettingsThumbnailFixture()
    previousDelay := A_WinDelay
    try {
        preview := OpenGesturePreview(app)
        preview.BeginGesture("MoveThumbnail", 2)
        preview.point := {x: 150, y: 170, buttons: 2}
        SetWinDelay(250)
        started := A_TickCount
        preview.GestureTick()
        AssertTrue(A_TickCount - started < 400, "Moving the layer stack must not incur four per-window delays.")
        AssertEqual(250, A_WinDelay)
        AssertSampleGeometry(preview, [70, 90, 250, 140])
        preview.point.buttons := 0
        preview.GestureTick()
    } finally {
        SetWinDelay(previousDelay)
        app.Dispose()
    }
}
SettingsThumbnailDragTest() {
    app := SettingsThumbnailFixture()
    try {
        preview := OpenGesturePreview(app)
        before := JSON.Dump(app._JSON)
        startIndex := preview.sampleIndex
        preview.point := {x: 100, y: 100, buttons: 2}
        PostMessage(0x204, 2, 0, preview.sample["Window"]["SampleStatus"])
        Sleep(30)
        AssertTrue(IsObject(preview.gesture), "Right-click over the label must use the configured Move gesture.")
        AssertFalse(preview.gesture.resize)
        updates := preview.geometryUpdates
        preview.GestureTick()
        AssertEqual(updates, preview.geometryUpdates, "A stationary pointer must not repaint the window stack.")
        preview.point := {x: 160, y: 170, buttons: 2}
        preview.GestureTick()
        AssertSampleGeometry(preview, [80, 90, 250, 140])
        preview.Refresh(A_TickCount + 5000)
        preview.NextSample()
        AssertEqual(startIndex, preview.sampleIndex, "Sample cycling must pause while dragging.")
        preview.point.buttons := 0
        preview.GestureTick()
        AssertFalse(IsObject(preview.gesture))
        AssertEqual(before, JSON.Dump(app._JSON))
        app.ThumbnailTextColor := "ABCDEF"
        preview.Refresh(A_TickCount + 6000)
        AssertSampleGeometry(preview, [80, 90, 250, 140])
        app.ThumbnailStartLocation["x"] := 40
        preview.Refresh(A_TickCount + 6500)
        AssertSampleGeometry(preview, [40, 20, 250, 140])
        AssertEqual(0, app.ThumbHwnd_EvEHwnd.Count, "Preview gestures must not enter the live-client interaction map.")
    } finally {
        app.Dispose()
    }
}

TestRunner.Register("Configured test-thumbnail resize and resize-all respect minimum size and retain text layout", SettingsThumbnailResizeTest)
SettingsThumbnailResizeTest() {
    app := SettingsThumbnailFixture()
    try {
        app.ThumbnailMinimumSize["width"] := 120
        app.ThumbnailMinimumSize["height"] := 80
        app.systemTrackingEnabled := true
        preview := OpenGesturePreview(app)
        for flags in [11, 3] { ; Ctrl+both = Resize; both = Resize All.
            preview.point := {x: 100, y: 100, buttons: 3}
            preview.SampleMouseMessage(flags, 0, 0x204, preview.sample["Window"].Hwnd)
            AssertTrue(preview.gesture.resize)
            preview.point := {x: 250, y: 200, buttons: 3}
            preview.GestureTick()
            AssertSampleGeometry(preview, [20, 20, 400, 240])
            preview.point := {x: -1000, y: -1000, buttons: 3}
            preview.GestureTick()
            AssertSampleGeometry(preview, [20, 20, 120, 80])
            Sleep(20)
            preview.sample["TextOverlay"]["SystemText"].GetPos(, &sy,, &sh)
            AssertTrue(sy + sh <= 80, "Overlays must relayout inside the resized thumbnail.")
            preview.point.buttons := 0
            preview.GestureTick()
            preview.ApplyGeometry(20, 20, 250, 140)
        }
    } finally {
        app.Dispose()
    }
}

TestRunner.Register("Custom preview gestures override cycling and middle-click still advances samples", SettingsThumbnailCustomGesturesTest)
SettingsThumbnailCustomGesturesTest() {
    app := SettingsThumbnailFixture()
    try {
        app.ThumbnailsInteractions["MoveThumbnail"] := Map("lmb", 1, "rmb", 0, "ctrl", 1, "shift", 0)
        app.ThumbnailsInteractions["ResizeThumbnail"] := Map("lmb", 0, "rmb", 1, "ctrl", 0, "shift", 1)
        app.ThumbnailsInteractions["MinimizeClient"] := Map("lmb", 0, "rmb", 0, "ctrl", 0, "shift", 0)
        app.ThumbnailsInteractions["QuickGroup"] := Map("lmb", 0, "rmb", 0, "ctrl", 0, "shift", 0)
        preview := OpenGesturePreview(app)
        preview.point := {x: 100, y: 100, buttons: 1}
        preview.SampleMouseMessage(9, 0, 0x201, preview.sample["Window"].Hwnd)
        AssertFalse(preview.gesture.resize, "Custom Ctrl+left must move, rather than the default minimize action.")
        preview.point := {x: 130, y: 140, buttons: 1}
        preview.GestureTick()
        AssertSampleGeometry(preview, [50, 60, 250, 140])
        preview.point.buttons := 0
        preview.GestureTick()
        preview.point := {x: 100, y: 100, buttons: 2}
        preview.SampleMouseMessage(6, 0, 0x204, preview.sample["Window"]["SampleStatus"].Hwnd)
        AssertTrue(preview.gesture.resize, "Custom Shift+right must resize, rather than the default Quick Group action.")
        preview.point.buttons := 0
        preview.GestureTick()
        index := preview.sampleIndex
        preview.SampleMouseMessage(0, 0, 0x208, preview.sample["Window"].Hwnd)
        AssertEqual(index + 1, preview.sampleIndex)
        preview.SampleMouseMessage(1, 0, 0x201, preview.sample["Window"]["SampleStatus"].Hwnd)
        AssertEqual(index + 2, preview.sampleIndex, "Activate-click must retain manual cycling.")
    } finally {
        app.Dispose()
    }
}

TestRunner.Register("Preview interactions simulate group, minimize, hide and close actions without changing real clients", SettingsThumbnailActionsTest)
SettingsThumbnailActionsTest() {
    app := SettingsThumbnailFixture()
    try {
        app.ThumbnailsInteractions["HideThumbnail"] := Map("lmb", 1, "rmb", 0, "ctrl", 1, "shift", 1)
        app.ThumbnailsInteractions["CloseClient"] := Map("lmb", 0, "rmb", 1, "ctrl", 1, "shift", 1)
        preview := OpenGesturePreview(app)
        window := preview.sample["Window"].Hwnd
        before := JSON.Dump(app._JSON)
        preview.SampleMouseMessage(5, 0, 0x201, window)
        AssertEqual("DisableFromGroups", preview.samples[preview.sampleIndex].id)
        preview.SampleMouseMessage(6, 0, 0x204, window)
        AssertEqual("QuickGroup", preview.samples[preview.sampleIndex].id)
        preview.SampleMouseMessage(9, 0, 0x201, window)
        AssertEqual("inactive", preview.samples[preview.sampleIndex].id)
        preview.SampleMouseMessage(13, 0, 0x201, window)
        AssertFalse(DllCall("IsWindowVisible", "Ptr", window))
        preview.NextSample()
        AssertTrue(DllCall("IsWindowVisible", "Ptr", window))
        preview.SampleMouseMessage(14, 0, 0x204, window)
        AssertTrue(preview.closed)
        AssertFalse(preview.HasOwnProp("mouseCallback"))
        AssertFalse(preview.HasOwnProp("gestureTimer"))
        AssertFalse(DllCall("IsWindow", "Ptr", window))
        AssertEqual(before, JSON.Dump(app._JSON))
    } finally {
        app.Dispose()
    }
}

TestRunner.Register("Preview samples cycle at one-second intervals and changes remain temporary", SettingsThumbnailFastCycleTest)
SettingsThumbnailFastCycleTest() {
    app := SettingsThumbnailFixture()
    try {
        app.gameLogsMonitoringEnabled := true
        app.monitoredEvents["warpDisrupted"]["enabled"] := 1
        app.eventDisplayDuration := 10000
        preview := app.OpenPreview()
        for sample in preview.samples
            AssertTrue(sample.duration <= 1000)
        start := preview.sampleIndex
        preview.Refresh(preview.sampleSince + 1000)
        AssertEqual(start + 1, preview.sampleIndex)
        preview.ApplyGeometry(100, 150, 300, 200)
        app.StopTestThumbnail()
        preview := app.OpenPreview()
        AssertSampleGeometry(preview, [20, 20, 250, 140])
    } finally {
        app.Dispose()
    }
}

TestRunner.Register("Preview cycles all requested systems with distinct demo colors without changing live assignments", SettingsThumbnailSystemVarietyTest)
SettingsThumbnailSystemVarietyTest() {
    for liveColors in [false, true] {
        app := SettingsThumbnailFixture()
        try {
            app.systemTrackingEnabled := true
            app.systemColorsEnabled := liveColors
            app.systemColors.Update(["Live client system"])
            liveMap := JSON.Dump(app.systemColors.active)
            settings := JSON.Dump(app._JSON)
            preview := app.OpenPreview()
            hwnd := preview.sample["Window"].Hwnd
            colors := Map()
            for expected in ["Jita", "Dodixie", "Amarr", "Hek", "Tama", "Turnur", "Ahbazon"] {
                overlay := preview.sample["TextOverlay"]
                AssertEqual(expected, overlay["SystemText"].Text)
                AssertTrue(overlay.SystemTextColor != "")
                AssertFalse(colors.Has(overlay.SystemTextColor), "Each test system must have its own color.")
                colors[overlay.SystemTextColor] := true
                preview.NextSample()
            }
            AssertEqual("Jita", preview.sample["TextOverlay"]["SystemText"].Text)
            AssertEqual(7, colors.Count)
            AssertEqual(hwnd, preview.sample["Window"].Hwnd, "Cycling must update the existing thumbnail without rebuilding its windows.")
            AssertEqual(liveMap, JSON.Dump(app.systemColors.active))
            AssertEqual("", app.systemColors.Get("Jita"), "Demo colors must not reserve a color for a live client.")
            AssertEqual(settings, JSON.Dump(app._JSON))
        } finally {
            app.Dispose()
        }
    }
}

TestRunner.Register("Preview varies incoming and outgoing DPS, whole-number damage shares and character text", SettingsThumbnailCombatVarietyTest)
SettingsThumbnailCombatVarietyTest() {
    app := SettingsThumbnailFixture()
    try {
        app.gameLogsMonitoringEnabled := true
        app.dpsMonitoring["incomingDPSEnabled"] := 1
        app.dpsMonitoring["outgoingDPSEnabled"] := 1
        app.dpsMonitoring["showIncomingDPSResistances"] := 1
        preview := app.OpenPreview()
        values := Map(), shares := Map(), names := Map()
        Loop 7 {
            overlay := preview.sample["TextOverlay"]
            AssertTrue(RegExMatch(overlay["DPSText"].Text, "^DPS: In: \d+ \| Out: \d+$"))
            AssertFalse(values.Has(overlay["DPSText"].Text))
            values[overlay["DPSText"].Text] := true
            percentages := overlay["DPSDamageText"].Text
            AssertFalse(shares.Has(percentages))
            shares[percentages] := true
            AssertTrue(RegExMatch(percentages, "^% (?:EM|Th|Ki|Ex) \d+(?: (?:EM|Th|Ki|Ex) \d+)*$"))
            AssertFalse(RegExMatch(percentages, "(?:EM|Th|Ki|Ex) 0\b"), "Zero damage types must stay omitted, matching real thumbnails.")
            total := 0
            for value in StrSplit(RegExReplace(percentages, "[^\d ]", ""), " ")
                if value != ""
                    total += Integer(value)
            AssertEqual(100, total)
            names[overlay["OverlayText"].Text] := true
            preview.NextSample()
        }
        AssertEqual(7, names.Count, "Fallback names should exercise different lengths and wrapping.")
        AssertEqual("DPS: In: 2500 | Out: 1800", preview.sample["TextOverlay"]["DPSText"].Text)
    } finally {
        app.Dispose()
    }
}

TestRunner.Register("Preview rotates configured characters with their matching aliases and custom colors", SettingsThumbnailCharacterVarietyTest)
SettingsThumbnailCharacterVarietyTest() {
    app := SettingsThumbnailFixture()
    try {
        app.CustomColorsActive := true
        app.CustomColors_AllCharNames := "Alice`nBob"
        app.CustomColors_AllTColors := "112233`n445566"
        app.CustomColors_AllBColors := "223344`n556677"
        app.CustomColors_IABorder_Colors := "334455`n667788"
        app.CustomThumbnailNames := Map("Characters", "Alice`nBob", "Names", "Scout`nLogistics Wing Commander")
        app.ShowAllColoredBorders := true
        preview := app.OpenPreview()
        AssertEqual("Scout", preview.sample["TextOverlay"]["OverlayText"].Text)
        preview.ApplyGeometry(100, 150, 140, 180)
        preview.NextSample()
        overlay := preview.sample["TextOverlay"]
        AssertEqual("Logistics Wing Commander", overlay["OverlayText"].Text)
        AssertEqual("445566", overlay.SystemTextDefaultColor)
        AssertEqual("667788", preview.sample["Border"].BackColor)
        AssertSampleGeometry(preview, [100, 150, 140, 180])
        AssertTrue(preview.MeasureThumbnailNameHeight(overlay, 130) > overlay.EventTextHeight)
        preview.NextSample()
        AssertEqual("Scout", overlay["OverlayText"].Text)
        AssertEqual("112233", overlay.SystemTextDefaultColor)
    } finally {
        app.Dispose()
    }
}

TestRunner.Register("Preview keeps its current sample data through settings edits and honors DPS gates and thresholds", SettingsThumbnailVarietySettingsTest)
SettingsThumbnailVarietySettingsTest() {
    app := SettingsThumbnailFixture()
    try {
        app.systemTrackingEnabled := true
        app.gameLogsMonitoringEnabled := true
        app.dpsMonitoring["incomingDPSEnabled"] := 1
        app.dpsMonitoring["outgoingDPSEnabled"] := 1
        app.dpsMonitoring["showIncomingDPSResistances"] := 1
        app.dpsMonitoring["incomingDPSThreshold"] := 100000
        app.dpsMonitoring["outgoingDPSThreshold"] := 200000
        preview := app.OpenPreview()
        previous := ""
        Loop 4 {
            text := preview.sample["TextOverlay"]["DPSText"].Text
            AssertTrue(RegExMatch(text, "In: (\d+) \| Out: (\d+)", &rates))
            AssertTrue(Integer(rates[1]) >= 100000 && Integer(rates[2]) >= 200000)
            AssertNotEqual(previous, text)
            previous := text
            preview.NextSample()
        }
        overlay := preview.sample["TextOverlay"]
        system := overlay["SystemText"].Text
        color := overlay.SystemTextColor
        text := overlay["DPSText"].Text
        step := preview.sampleStep
        app.ThumbnailTextColor := "AABBCC"
        preview.Refresh(A_TickCount + 500)
        overlay := preview.sample["TextOverlay"]
        AssertEqual(step, preview.sampleStep, "Appearance edits must not reset or advance sample data.")
        AssertEqual(system, overlay["SystemText"].Text)
        AssertEqual(color, overlay.SystemTextColor)
        AssertEqual(text, overlay["DPSText"].Text)
        app.dpsMonitoring["incomingDPSEnabled"] := 0
        app.systemTrackingEnabled := false
        preview.Refresh(A_TickCount + 1000)
        overlay := preview.sample["TextOverlay"]
        AssertTrue(RegExMatch(overlay["DPSText"].Text, "^DPS: Out: \d+$"))
        AssertEqual("", overlay["DPSDamageText"].Text)
        AssertEqual("", overlay["SystemText"].Text)
        AssertEqual("", overlay.SystemTextColor)
        app.dpsMonitoring["outgoingDPSEnabled"] := 0
        preview.Refresh(A_TickCount + 1500)
        AssertEqual("", preview.sample["TextOverlay"]["DPSText"].Text)
    } finally {
        app.Dispose()
    }
}

TestRunner.Register("Test thumbnail uses exact default geometry without registering a client or saving positions", SettingsThumbnailPositionTest)
SettingsThumbnailPositionTest() {
    app := SettingsThumbnailFixture()
    baseline := LiveThumb.OBJ_COUNTER
    try {
        for key, value in Map("x", -300, "y", 70, "width", 260, "height", 180)
            app.ThumbnailStartLocation[key] := value
        app.AutoSaveThumbnailPositions := true
        app.Profiles["Default"]["Thumbnail Positions"]["EVE - Test Character"] := Map("x", 400, "y", 300, "width", 80, "height", 50)
        before := JSON.Dump(app._JSON)
        preview := app.OpenPreview()
        window := preview.sample["Window"]
        WinGetPos(&x, &y, &w, &h, window.Hwnd)
        AssertEqual([-300, 70, 260, 180], [x, y, w, h], "Saved client positions must not replace the default test position.")
        AssertEqual(app.S_Gui.Hwnd, DllCall("GetWindow", "Ptr", window.Hwnd, "UInt", 4, "Ptr"))
        AssertEqual(0, app.ThumbHwnd_EvEHwnd.Count)
        AssertEqual(baseline, LiveThumb.OBJ_COUNTER, "Test thumbnails must not consume DWM registrations.")
        AssertFalse(app.ThumbWindows.HasOwnProp("sample"))
        AssertEqual(before, JSON.Dump(app._JSON), "Test settings and positions must never be written into the real profile.")
        hwnd := window.Hwnd
        app.StopTestThumbnail()
        AssertFalse(DllCall("IsWindow", "Ptr", hwnd))
        AssertTrue(preview.closed)
        AssertFalse(preview.HasOwnProp("timer"), "Closing must release the callback and its renderer reference.")
        preview.Tick() ; A queued tick after close must be harmless.
    } finally {
        app.Dispose()
    }
}

TestRunner.Register("Test thumbnail respects independent inheritance of event settings and game-log controls", SettingsThumbnailEventProfileTest)
SettingsThumbnailEventProfileTest() {
    app := SettingsThumbnailFixture()
    try {
        app.monitoredEvents["warpDisrupted"]["enabled"] := 1
        app.monitoredEvents["warpDisrupted"]["color"] := "123456"
        app.Profiles["Other"] := JSON.Load(JSON.Dump(app.Profiles["Default"]))
        app.Profiles["Other"]["Game Logs Monitoring"]["monitoredEvents"]["warpDisrupted"]["enabled"] := 0
        app.Profiles["Other"]["Game Logs Monitoring"]["monitoredEvents"]["underAttackByNPC"]["enabled"] := 1
        app.Profiles["Default"]["Other"]["Global_Groups"]["Monitored Events"] := 1
        app.Profiles["Other"]["Other"]["Global_Groups"]["Monitored Events"] := 0
        app.LastUsedProfile := "Other"
        app.ProfileOverride()
        app.gameLogsMonitoringEnabled := true
        app.showEventText := true
        preview := app.OpenPreview()
        AssertEqual(5, preview.samples.Length)
        AssertEqual("event:warpDisrupted", preview.samples[5].id)
        AssertEqual("123456", preview.monitoredEvents["warpDisrupted"]["color"])
        app.Profiles["Other"]["Other"]["Global_Groups"]["Monitored Events"] := 1
        app.ProfileOverride()
        preview.Refresh(A_TickCount + 500)
        AssertEqual("event:underAttackByNPC", preview.samples[5].id)
    } finally {
        app.Dispose()
    }
}

TestRunner.Register("Test thumbnail button stays above the shared sidebar status footer", SettingsThumbnailButtonTest)
SettingsThumbnailButtonTest() {
    app := SettingsThumbnailFixture()
    try {
        button := app.testThumbnailButton
        button.GetPos(&x, &y, &w, &h)
        app.sidebarStatusBlock.GetPos(, &noticeY,, &noticeH)
        lastGroupBottom := 0
        for control in app.Sidebar {
            if control.Type != "Button" || control = button
                continue
            control.GetPos(, &cy,, &ch)
            lastGroupBottom := Max(lastGroupBottom, cy + ch)
        }
        AssertTrue(y >= lastGroupBottom + app.baseGrid && y + h <= noticeY)
        AssertTrue(noticeY + noticeH <= app.guiHeight - 8)
        AssertTrue(x + w <= app.sidebarW && y + h = app.guiHeight - 56, "The footer reserves a compact fixed status block.")
        id := DllCall("GetDlgCtrlID", "Ptr", button.Hwnd, "Int")
        PostMessage(0x111, id & 0xFFFF, button.Hwnd,, "ahk_id " app.Sidebar.Hwnd)
        Sleep(30)
        AssertTrue(IsObject(app.settingsThumbnail))
        AssertEqual("Close Test Thumbnail", button.Text)
        hwnd := app.settingsThumbnail.sample["Window"].Hwnd
        PostMessage(0x111, id & 0xFFFF, button.Hwnd,, "ahk_id " app.Sidebar.Hwnd)
        Sleep(30)
        AssertEqual("Test Thumbnail", button.Text)
        AssertFalse(IsObject(app.settingsThumbnail))
        AssertFalse(DllCall("IsWindow", "Ptr", hwnd))
    } finally {
        app.Dispose()
    }
}

TestRunner.Register("Test thumbnail shows only enabled system and DPS overlays and follows overlay visibility", SettingsThumbnailOverlayTest)
SettingsThumbnailOverlayTest() {
    app := SettingsThumbnailFixture()
    try {
        app.systemTrackingEnabled := true
        app.systemColorsEnabled := true
        app.dpsMonitoring["incomingDPSEnabled"] := 1
        app.dpsMonitoring["outgoingDPSEnabled"] := 1
        app.dpsMonitoring["showIncomingDPSResistances"] := 1
        preview := app.OpenPreview()
        overlay := preview.sample["TextOverlay"]
        AssertEqual("Jita", overlay["SystemText"].Text, "Local tracking works independently of game-log monitoring.")
        AssertTrue(overlay.SystemTextColor != "")
        AssertEqual("", overlay["DPSText"].Text, "Disabled game-log monitoring must suppress DPS samples.")
        app.gameLogsMonitoringEnabled := true
        preview.Refresh(A_TickCount + 500)
        overlay := preview.sample["TextOverlay"]
        AssertEqual("DPS: In: 2500 | Out: 1800", overlay["DPSText"].Text)
        AssertTrue(InStr(overlay["DPSDamageText"].Text, "EM 25"))
        AssertTrue(InStr(overlay["DPSDamageText"].Text, "Ex 25"))
        app.ShowThumbnailTextOverlay := false
        preview.Refresh(A_TickCount + 1000)
        AssertFalse(DllCall("IsWindowVisible", "Ptr", preview.sample["TextOverlay"].Hwnd))
        app.ShowThumbnailTextOverlay := true
        app.systemTrackingEnabled := false
        app.dpsMonitoring["incomingDPSEnabled"] := 0
        app.dpsMonitoring["outgoingDPSEnabled"] := 0
        preview.Refresh(A_TickCount + 1500)
        overlay := preview.sample["TextOverlay"]
        AssertEqual("", overlay["DPSText"].Text)
        AssertEqual("", overlay["SystemText"].Text)
        AssertFalse(overlay.HasOwnProp("SystemTextColor") && overlay.SystemTextColor != "")
    } finally {
        app.Dispose()
    }
}

TestRunner.Register("Test thumbnail cycles enabled interaction and event borders with configured flashing and text", SettingsThumbnailSamplesTest)
SettingsThumbnailSamplesTest() {
    app := SettingsThumbnailFixture()
    try {
        app.gameLogsMonitoringEnabled := true
        app.showEventText := true
        app.flashBorderEnabled := true
        app.monitoredEvents["warpDisrupted"]["enabled"] := 1
        app.monitoredEvents["warpDisrupted"]["color"] := "A1B2C3"
        app.ShowAllColoredBorders := false
        app.EventActiveBorderMode := 3 ; Verify the retained combined-border behavior here.
        preview := app.OpenPreview()
        AssertEqual(5, preview.samples.Length, "Active, inactive, both assigned interactions and one enabled event.")
        Loop 4
            preview.NextSample()
        AssertEqual("event:warpDisrupted", preview.samples[preview.sampleIndex].id)
        AssertEqual("Warp Disrupted", preview.sample["TextOverlay"]["EventText"].Text)
        AssertEqual("A1B2C3", preview.sample["Border"].BackColor)
        AssertTrue(DllCall("IsWindowVisible", "Ptr", preview.sample["Border"].Hwnd))
        preview.PaintSample(preview.sampleSince + preview.flashBorderInterval + 1)
        AssertFalse(DllCall("IsWindowVisible", "Ptr", preview.sample["Border"].Hwnd), "An off flash restores the inactive border policy.")
        app.showEventText := false
        app.flashBorderEnabled := false
        preview.Refresh(A_TickCount + 500)
        AssertEqual(4, preview.samples.Length, "Events without text or flashing have no visual sample.")
        AssertEqual("", preview.sample["TextOverlay"]["EventText"].Text)
    } finally {
        app.Dispose()
    }
}

TestRunner.Register("Test thumbnail uses configured custom character colors, aliases and click-through state", SettingsThumbnailAppearanceTest)
SettingsThumbnailAppearanceTest() {
    app := SettingsThumbnailFixture()
    try {
        app.CustomColorsActive := true
        app.CustomColors_AllCharNames := "Alice"
        app.CustomColors_AllTColors := "112233"
        app.CustomColors_AllBColors := "445566"
        app.CustomColors_IABorder_Colors := "778899"
        app.CustomThumbnailNames := Map("Characters", "Alice", "Names", "Scout")
        app.ShowAllColoredBorders := true
        app.ClickThroughActive := true
        app.ShowThumbnailsAlwaysOnTop := true
        app._JSON["_Profiles"]["Default"]["Thumbnails Visuals"]["ThumbnailOpacity"] := 65
        preview := app.OpenPreview()
        AssertEqual("Scout", preview.sample["TextOverlay"]["OverlayText"].Text)
        AssertEqual("112233", preview.sample["TextOverlay"].SystemTextDefaultColor)
        AssertEqual("445566", preview.sample["Border"].BackColor)
        AssertTrue(WinGetExStyle(preview.sample["Window"].Hwnd) & 0x20)
        AssertTrue(WinGetExStyle(preview.sample["Window"].Hwnd) & 0x8)
        AssertEqual(app.ThumbnailOpacity, WinGetTransparent(preview.sample["Window"].Hwnd))
        preview.NextSample()
        AssertEqual("778899", preview.sample["Border"].BackColor)
    } finally {
        app.Dispose()
    }
}

TestRunner.Register("Test thumbnail refreshes effective profile settings and tolerates incomplete edits", SettingsThumbnailRefreshTest)
SettingsThumbnailRefreshTest() {
    app := SettingsThumbnailFixture()
    try {
        app.Profiles["Other"] := JSON.Load(JSON.Dump(app.Profiles["Default"]))
        app.Profiles["Other"]["Thumbnails Visuals"]["ThumbnailTextColor"] := "ABCDEF"
        app.Profiles["Other"]["Thumbnails Behavior"]["ThumbnailStartLocation"]["x"] := 120
        preview := app.OpenPreview()
        oldWindow := preview.sample["Window"].Hwnd
        app.LastUsedProfile := "Other"
        app.ProfileOverride()
        preview.Refresh(A_TickCount + 500)
        AssertEqual("ABCDEF", preview.sample["TextOverlay"].SystemTextDefaultColor)
        WinGetPos(&x,,, , preview.sample["Window"].Hwnd)
        AssertEqual(120, x)
        AssertFalse(DllCall("IsWindow", "Ptr", oldWindow))
        app.Profiles["Default"]["Other"]["Global_Groups"]["Thumbnails Visuals"] := 1
        app.Profiles["Other"]["Other"]["Global_Groups"]["Thumbnails Visuals"] := 0
        app.ProfileOverride()
        preview.Refresh(A_TickCount + 800)
        AssertEqual("FAC57A", preview.sample["TextOverlay"].SystemTextDefaultColor, "Inherited groups must preview Default's effective values.")
        currentWindow := preview.sample["Window"].Hwnd
        app.ThumbnailStartLocation["width"] := ""
        preview.Refresh(A_TickCount + 1000)
        AssertEqual(currentWindow, preview.sample["Window"].Hwnd, "Keep the previous preview during an incomplete numeric edit.")
        app.ThumbnailStartLocation["width"] := 280
        app.ThumbnailBackgroundColor := "incomplete"
        before := JSON.Dump(app._JSON)
        preview.Refresh(A_TickCount + 1500)
        WinGetPos(,, &width,, preview.sample["Window"].Hwnd)
        AssertEqual(280, width)
        AssertEqual("57504E", preview.sample["Window"].BackColor)
        AssertEqual(before, JSON.Dump(app._JSON), "Preview fallback colors must not rewrite the user's in-progress edits.")
    } finally {
        app.Dispose()
    }
}

TestRunner.Register("Open test-thumbnail timer applies edits automatically and stops on Settings cleanup", SettingsThumbnailTimerTest)
SettingsThumbnailTimerTest() {
    app := SettingsThumbnailFixture()
    try {
        app.ToggleTestThumbnail()
        preview := app.settingsThumbnail
        app.ThumbnailTextColor := "123456"
        Sleep(400)
        AssertFalse(preview.closed)
        AssertEqual("123456", preview.sample["TextOverlay"].SystemTextDefaultColor)
        app.StopTestThumbnail()
        Sleep(150)
        AssertTrue(preview.closed)
        AssertFalse(IsObject(app.settingsThumbnail))
    } finally {
        app.Dispose()
    }
}

TestRunner.Register("Destroyed Settings owners stop test-thumbnail timers and destroy the entire window stack", SettingsThumbnailOwnerCloseTest)
SettingsThumbnailOwnerCloseTest() {
    app := SettingsThumbnailFixture()
    try {
        preview := app.OpenPreview()
        handles := []
        for key in ["Window", "TextOverlay", "Border"]
            handles.Push(preview.sample[key].Hwnd)
        app.S_Gui.Destroy()
        preview.Tick()
        AssertTrue(preview.closed)
        AssertFalse(IsObject(app.settingsThumbnail))
        for hwnd in handles
            AssertFalse(DllCall("IsWindow", "Ptr", hwnd))
    } finally {
        app.Dispose()
    }
}
