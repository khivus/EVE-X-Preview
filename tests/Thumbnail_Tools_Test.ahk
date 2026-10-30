TestRunner.Register("Automatic wildcard opt-out preserves explicit syntax and updates conflict warnings", ThumbnailWildcardOptOutTest)
ThumbnailWildcardOptOutTest() {
    app := CombatEventsFixture()
    app.ProfileOverride()
    app.AutoWildcardWindowHotkeys := false
    app.SwitchToPreviousWindow_Hotkey := "F23"
    app.Exit_Program_Hotkey := "^F23"
    conflicts := HotkeyBindings.Conflicts(HotkeyBindings.Collect(app))
    AssertFalse(conflicts.Has("setting:SwitchToPreviousWindow_Hotkey"))
    AssertEqual("F23", HotkeyBindings.ForWindowSwitch("F23", app.AutoWildcardWindowHotkeys))
    AssertEqual("*F23", HotkeyBindings.ForWindowSwitch("*F23", false))
    AssertEqual("XButton1 & 1", HotkeyBindings.ForWindowSwitch("XButton1 & 1", false))
    app.AutoWildcardWindowHotkeys := true
    AssertTrue(HotkeyBindings.Conflicts(HotkeyBindings.Collect(app)).Has("setting:SwitchToPreviousWindow_Hotkey"))
}

class ThumbnailExactHotkeyFixture extends CombatEventsFixture {
    keys := []
    RegisterHotkeyWithLanguageRetry(key, *) {
        this.keys.Push(key)
        return true
    }
}
TestRunner.Register("Window hotkey registration honors exact and wildcard policies without rewriting settings", ThumbnailExactRegistrationTest)
ThumbnailExactRegistrationTest() {
    app := ThumbnailExactHotkeyFixture()
    app.ProfileOverride()
    app.SwitchLangOnErr := true
    for enabled in [false, true] {
        app.AutoWildcardWindowHotkeys := enabled
        before := JSON.Dump(app._JSON)
        app.RegisterWindowHotkey("^F23", (*) => 0, "Test")
        AssertEqual(enabled ? "*^F23" : "^F23", app.keys[-1])
        AssertEqual(before, JSON.Dump(app._JSON))
    }
}

TestRunner.Register("Anonymous labels are unique stable reversible and cover distinct login windows", AnonymousThumbnailLabelsTest)
AnonymousThumbnailLabelsTest() {
    app := CombatEventsFixture()
    app.ProfileOverride()
    app.CustomThumbnailNames := Map("Characters", "Alice", "Names", "Custom Scout")
    app.AnonymousNames := "Nova`n`nNOVA`nComet"
    app.AnonymizeThumbnails := true
    aliases := Map()
    aliases.CaseSense := "Off"
    for source in ["Alice", "Bob", "Carol"] {
        alias := app.GetThumbnailDisplayText("EVE - " source)
        AssertFalse(aliases.Has(alias), "Names must not repeat when the pool runs out.")
        aliases[alias] := true
        AssertEqual(alias, app.GetThumbnailDisplayText("EVE - " StrLower(source)))
        AssertEqual(alias "`nUnder attack", app.GetThumbnailDisplayText("EVE - " source "`nUnder attack"))
    }
    AssertNotEqual(app.GetThumbnailDisplayText("EVE", 101), app.GetThumbnailDisplayText("EVE", 102))
    AssertNotEqual(app.GetThumbnailDisplayText("Char Screen", 101), app.GetThumbnailDisplayText("Char Screen", 102))
    app.AnonymizeThumbnails := false
    AssertEqual("Custom Scout", app.GetThumbnailDisplayText("EVE - Alice"))
    AssertEqual("Bob", app.GetThumbnailDisplayText("EVE - Bob"))
    AssertEqual("Custom Scout", app.CustomThumbnailNames["Names"])
}

TestRunner.Register("Thumbnail tool settings survive JSON and respect behavior profile overrides", ThumbnailToolsProfileTest)
ThumbnailToolsProfileTest() {
    app := CombatEventsFixture()
    app.ProfileOverride()
    app.LockThumbnailPositions := true
    app.AnonymizeThumbnails := true
    app.AnonymousNames := "One`nTwo"
    app.ThumbnailPollingInterval := 250
    app._JSON := JSON.Load(JSON.Dump(app._JSON))
    AssertTrue(app.LockThumbnailPositions)
    AssertTrue(app.AnonymizeThumbnails)
    AssertEqual(250, app.ThumbnailPollingInterval)
    AssertEqual("One`nTwo", app.AnonymousNames)
    app.Profiles["OtherTools"] := DeepClone(app.Profiles["Default"])
    app.ProfileThumbnailsBehavior := "OtherTools"
    app.ProfileThumbnailsVisuals := "OtherTools"
    app.LockThumbnailPositions := false
    app.AnonymousNames := "Other"
    AssertEqual("Other", app.GetThumbnailDisplayText("Alice"))
    app.ProfileThumbnailsBehavior := "Default"
    app.ProfileThumbnailsVisuals := "Default"
    AssertTrue(app.LockThumbnailPositions)
    AssertTrue(app.GetThumbnailDisplayText("Alice") != "Other")
}

TestRunner.Register("Position lock blocks single all resize and snap gestures before touching windows", ThumbnailPositionLockTest)
ThumbnailPositionLockTest() {
    app := ThumbnailDragFixture()
    app.ProfileThumbnailsBehavior := "Default"
    app.Profiles["Default"]["Thumbnails Behavior"] := Map("LockThumbnailPositions", 1)
    try {
        hwnd := app.ThumbWindows.%1%["Window"].Hwnd
        for all in [false, true] {
            app.Mouse_DragMove(hwnd, all)
            app.Mouse_ResizeThumb(hwnd, all)
        }
        app.Window_Snap(0, {})
        AssertEqual(0, app.positionIndex)
        AssertEqual(0, app.moves)
        app.LockThumbnailPositions := false
        AssertEqual(6, app.CaptureThumbnailDragWindows(1, true).Count)
    } finally app.Cleanup()
}

TestRunner.Register("Window polling interval remains bounded without changing thumbnail dimensions", WindowPollingSettingsTest)
WindowPollingSettingsTest() {
    app := CombatEventsFixture()
    app.ProfileOverride()
    app.ThumbnailPollingInterval := "invalid"
    AssertEqual(50, app.ThumbnailPollingInterval)
    app.ThumbnailPollingInterval := 5000
    AssertEqual(1000, app.ThumbnailPollingInterval)
    app.ThumbnailPollingInterval := 100
    AssertEqual(100, app.ThumbnailPollingInterval)
}

TestRunner.Register("Losing client focus clears active highlighting and preserves inactive and group borders", ThumbnailFocusBorderTest)
ThumbnailFocusBorderTest() {
    previousHidden := A_DetectHiddenWindows
    DetectHiddenWindows(true)
    app := ThumbnailBorderFixture()
    try {
        app.flashMethod.Clear()
        app.BorderActive := app.key
        app.activateBorder(app.key, app.title)
        AssertEqual(app.ClientHighligtColor, app.thumb["Border"].BackColor)
        app.ClearActiveThumbnailBorder()
        AssertEqual(0, app.BorderActive)
        AssertFalse(DllCall("IsWindowVisible", "Ptr", app.thumb["Border"].Hwnd))
        app.ShowAllColoredBorders := true
        app.BorderActive := app.key
        app.activateBorder(app.key, app.title)
        app.ClearActiveThumbnailBorder()
        AssertEqual(app.InactiveClientBorderColor, app.thumb["Border"].BackColor)
        app.QuickGroupEnabled := true
        app.QuickGroupChars[app.key] := app.title
        app.BorderActive := app.key
        app.ClearActiveThumbnailBorder()
        AssertEqual(app.QuickGroupColor, app.thumb["Border"].BackColor)
    } finally {
        app.Dispose()
        DetectHiddenWindows(previousHidden)
    }
}

TestRunner.Register("New thumbnail options live in their matching settings categories and save through native controls", ThumbnailToolsGuiTest)
ThumbnailToolsGuiTest() {
    app := ColorSettingsFixture()
    try {
        app.ThumbnailsBehavior_Ctrl()
        app.HotkeyGroups_Ctrl()
        AssertFalse(app.MainFrame.Group.Has("Thumbnail Tools"))
        for name, category in Map("AutoWildcardWindowHotkeys", "Hotkey Groups", "LockThumbnailPositions", "Thumbnails Behavior",
            "ThumbnailPollingInterval", "Thumbnails Behavior",
            "AnonymizeThumbnails", "Thumbnails Visuals", "AnonymousNames", "Thumbnails Visuals") {
            control := app.MainFrame[name]
            AssertEqual(category, control.SettingsCategory)
            found := false
            for item in app.MainFrame.Group[category]
                found := found || item.Hwnd = control.Hwnd
            AssertTrue(found, "Missing " name " in " category)
            if control.Type = "CheckBox" {
                before := app.%name%
                SendMessage(0xF5, 0, 0, control)
                Sleep(20)
                AssertEqual(!before, app.%name%)
            }
        }
        AssertEqual(app.AnonymousNames, app.MainFrame["AnonymousNames"].Value)
        AssertEqual(50, app.MainFrame["ThumbnailPollingInterval"].Value)
    } finally {
        SetTimer(app.Save_Settings_Delay_Timer, 0)
        app.Dispose()
    }
}

TestRunner.Register("Focus loss clears painted active borders even after activation caches were reset", ThumbnailPaintedFocusResetTest)
ThumbnailPaintedFocusResetTest() {
    previous := A_DetectHiddenWindows
    DetectHiddenWindows(true)
    app := ThumbnailBorderFixture()
    try {
        app.flashMethod.Clear()
        app.activateBorder(app.key, app.title)
        AssertTrue(app.thumb["Border"].ActiveHighlight)
        app.BorderActive := 0
        app.lastActiveThumbHwnd := 0
        app.focused := false
        app.ClearActiveThumbnailBorder()
        AssertFalse(app.thumb["Border"].ActiveHighlight)
        AssertFalse(DllCall("IsWindowVisible", "Ptr", app.thumb["Border"].Hwnd))
        app.focused := true
        app.ShowActiveBorder(app.key)
        AssertTrue(app.thumb["Border"].ActiveHighlight, "Returning focus must restore the highlight.")
    } finally {
        app.Dispose()
        DetectHiddenWindows(previous)
    }
}

TestRunner.Register("Disabled group features cannot preserve active borders through remembered membership", ThumbnailDormantMembershipFocusTest)
ThumbnailDormantMembershipFocusTest() {
    previous := A_DetectHiddenWindows
    DetectHiddenWindows(true)
    app := ThumbnailBorderFixture()
    try {
        app.flashMethod.Clear()
        app.DisabledChars[app.key] := app.title
        app.QuickGroupChars[app.key] := app.title
        for allBorders in [false, true] {
            app.focused := true
            app.ShowAllColoredBorders := allBorders
            app.activateBorder(app.key, app.title)
            app.focused := false
            app.ClearActiveThumbnailBorder()
            AssertFalse(app.thumb["Border"].ActiveHighlight)
            if allBorders
                AssertEqual(app.InactiveClientBorderColor, app.thumb["Border"].BackColor)
            else
                AssertFalse(DllCall("IsWindowVisible", "Ptr", app.thumb["Border"].Hwnd))
        }
    } finally {
        app.Dispose()
        DetectHiddenWindows(previous)
    }
}

TestRunner.Register("Background event callbacks cannot redraw active highlights while event borders still work", ThumbnailBackgroundDrawGuardTest)
ThumbnailBackgroundDrawGuardTest() {
    previous := A_DetectHiddenWindows
    DetectHiddenWindows(true)
    app := ThumbnailBorderFixture()
    try {
        app.activateBorder(app.key, app.title)
        app.focused := false
        app.ClearActiveThumbnailBorder()
        app.activateBorder(app.key, app.title)
        AssertFalse(app.thumb["Border"].ActiveHighlight)
        AssertFalse(DllCall("IsWindowVisible", "Ptr", app.thumb["Border"].Hwnd))
        app.DrawFlashingBorder(app.title)
        AssertFalse(app.thumb["Border"].ActiveHighlight)
        AssertFalse(DllCall("IsWindowVisible", "Ptr", app.thumb["Border"].Hwnd))
        AssertTrue(DllCall("IsWindowVisible", "Ptr", app.thumb["EventBorder"].Hwnd), "Separate event alerts are independent of active highlighting.")
    } finally {
        app.Dispose()
        DetectHiddenWindows(previous)
    }
}

class ThumbnailForegroundPollFixture extends ThumbnailShiftFixture {
    foreground := 0
    __New() {
        super.__New()
        this.BorderActive := 0
        this.lastActiveThumbHwnd := 0
        this.flashMethod := Map()
        this.eventMethods := Map()
        this.allTrackedApps := Map("fixture.exe", true)
        this.ShowAllColoredBorders := false
        this.ShowClientHighlightBorder := true
    }
    _GetForegroundHwnd() => this.foreground
    _GetForegroundInfo(*) => {hwnd: this.foreground, exe: this.ThumbWindows.HasOwnProp(this.foreground) ? "fixture.exe" : ""}
    EventSourceActive(hwnd) => hwnd = this.foreground
}

TestRunner.Register("Main window polling clears stale highlights on external focus and restores them on return", ThumbnailForegroundPollingTest)
ThumbnailForegroundPollingTest() {
    previous := A_DetectHiddenWindows
    DetectHiddenWindows(true)
    app := ThumbnailForegroundPollFixture(), sources := []
    outside := Gui()
    try {
        source := app.CreateSource("EVE - Focus polling")
        sources.Push(source)
        app.EVE_WIN_Created(source.Hwnd, source.Title)
        app.foreground := source.Hwnd
        app.PollMainWindows()
        border := app.ThumbWindows.%source.Hwnd%["Border"]
        AssertTrue(border.ActiveHighlight)
        app.PollMainWindows()
        AssertTrue(border.ActiveHighlight, "Repeated polling must retain the focused client's highlight.")
        ; Event/group updates invalidate this cache without erasing the ring.
        app.BorderActive := 0
        app.lastActiveThumbHwnd := 0
        app.foreground := outside.Hwnd
        app.PollMainWindows()
        AssertFalse(border.ActiveHighlight)
        AssertFalse(DllCall("IsWindowVisible", "Ptr", border.Hwnd))
        app.foreground := source.Hwnd
        app.PollMainWindows()
        AssertTrue(border.ActiveHighlight)
    } finally {
        app.Cleanup(sources)
        outside.Destroy()
        DetectHiddenWindows(previous)
    }
}
