class ThumbnailShiftFixture extends PreviewLifecycleFixture {
    __New() {
        super.__New()
        static groupId := 0
        this.sourceGroup := "ThumbnailShiftSources" (++groupId)
        this.EVEExe := "ahk_group " this.sourceGroup
        for key, value in Map("x", 100, "y", 100, "width", 100, "height", 80)
            this.ThumbnailStartLocation[key] := value
        this.allLoginClosed := true
        this.skipShiftThumbs := false
        this.HideThumbnailsOnLostFocus := false
        this.HideThumbForActiveWin := false
        this.DisabledFromGroupsEnabled := false
        this.QuickGroupEnabled := false
        this.ActiveNonEVEApps := []
        this.allTrackedApps := Map()
        this.DestroyThumbnailsToggle := 1
        this.ProfActive := false
        this.SlowThumbnailCreation := false
        this.previewQueue := PreviewStartupQueue()
    }
    GetMonitorBounds(x, y, &left, &top, &right, &bottom) {
        left := 0, top := 0, right := 1920, bottom := 1080
    }
    UpdateActiveNonEVEApps() {
    }
    _GetForegroundInfo(*) {
        return {hwnd: 0, exe: ""}
    }
    _ObserveEVEActivation(*) {
    }
    CreateSource(title, tracked := true) {
        source := Gui(, title)
        source.Show("Hide w160 h100")
        if tracked
            GroupAdd(this.sourceGroup, "ahk_id " source.Hwnd)
        return source
    }
    Cleanup(sources) {
        for source in sources {
            this.DisposePreview(source.Hwnd)
            source.Destroy()
        }
        SetTimer(this.debugToolTipMethod, 0)
    }
}

AssertShiftPosition(app, source, expectedX, expectedY, expectedW := 100, expectedH := 80) {
    WinGetPos(&x, &y, &w, &h, app.ThumbWindows.%source.Hwnd%["Window"].Hwnd)
    AssertEqual([expectedX, expectedY, expectedW, expectedH], [x, y, w, h])
}

TestRunner.Register("Logged-in startup shifts unsaved positions independently and preserves saved positions", LoggedInThumbnailShiftTest)
LoggedInThumbnailShiftTest() {
    oldHidden := A_DetectHiddenWindows
    DetectHiddenWindows(true)
    app := ThumbnailShiftFixture(), sources := []
    try {
        app.ShiftThumbsForLoginScreen := 0
        app.ShiftThumbsDirection := 5
        app.ShiftThumbVerticalStep := 95
        for index in [1, 2, 3, 4] {
            source := app.CreateSource("EVE - Shift test " index)
            sources.Push(source)
        }
        app._JSON["_Profiles"]["Default"]["Thumbnail Positions"]["EVE - Shift test 1"] := Map("x", 100, "y", 100, "width", 100, "height", 80)
        app.EVE_WIN_Created(sources[1].Hwnd, "EVE - Shift test 1")
        app.EVE_WIN_Created(sources[2].Hwnd, "EVE - Shift test 2")
        app.EVE_WIN_Created(sources[3].Hwnd, "EVE - Shift test 3")
        AssertShiftPosition(app, sources[1], 100, 100)
        AssertShiftPosition(app, sources[2], 100, 195)
        AssertShiftPosition(app, sources[3], 100, 290)
        AssertEqual(1, app.ThumbnailPositions.Count, "Automatic shifting must not save positions.")
        app.ShiftThumbsForLoggedInCharacters := 0
        app.EVE_WIN_Created(sources[4].Hwnd, "EVE - Shift test 4")
        AssertShiftPosition(app, sources[4], 100, 100)
    } finally {
        app.Cleanup(sources)
        DetectHiddenWindows(oldHidden)
    }
}

TestRunner.Register("Logged-in sequential shifting advances across polls with collision avoidance disabled", LoggedInSequentialShiftTest)
LoggedInSequentialShiftTest() {
    oldHidden := A_DetectHiddenWindows
    DetectHiddenWindows(true)
    app := ThumbnailShiftFixture(), sources := []
    try {
        app.ShiftThumbsCollisionCheck := 0
        app.ShiftThumbsForLoginScreen := 0
        app.ShiftThumbHorizontalStep := 125
        for index in [1, 2, 3] {
            source := app.CreateSource("EVE - Shift test " index)
            sources.Push(source)
            app.PollMainWindows()
            AssertShiftPosition(app, source, 100 + (index - 1) * 125, 100)
            AssertFalse(app.allLoginClosed, "Polling must not reset the next position just because no login screens remain.")
        }
    } finally {
        app.Cleanup(sources)
        DetectHiddenWindows(oldHidden)
    }
}

TestRunner.Register("Login shifting remains independent and non-EVE previews stay at their default position", LoginThumbnailShiftIndependenceTest)
LoginThumbnailShiftIndependenceTest() {
    oldHidden := A_DetectHiddenWindows
    DetectHiddenWindows(true)
    app := ThumbnailShiftFixture(), sources := []
    try {
        app.ShiftThumbsForLoggedInCharacters := 0
        for index in [1, 2, 3] {
            source := app.CreateSource("Other app " index, false)
            sources.Push(source)
        }
        app.EVE_WIN_Created(sources[1].Hwnd, "EVE")
        app.EVE_WIN_Created(sources[2].Hwnd, "EVE")
        AssertShiftPosition(app, sources[1], 100, 100)
        AssertShiftPosition(app, sources[2], 200, 100)
        app.ShiftThumbsForLoggedInCharacters := 1
        app.EVE_WIN_Created(sources[3].Hwnd, "Other app 3")
        AssertShiftPosition(app, sources[3], 100, 100)
    } finally {
        app.Cleanup(sources)
        DetectHiddenWindows(oldHidden)
    }
}

TestRunner.Register("Logged-in shifting defaults on for older profiles, persists opt-out and sits beside login shifting", LoggedInThumbnailShiftSettingTest)
LoggedInThumbnailShiftSettingTest() {
    app := ThumbnailShiftFixture()
    AssertEqual(1, app.ShiftThumbsForLoggedInCharacters)
    app._JSON["_Profiles"][app.ProfileThumbnailsBehavior]["Thumbnails Behavior"].Delete("ShiftThumbsForLoggedInCharacters")
    AssertEqual(1, app.ShiftThumbsForLoggedInCharacters)
    app.ShiftThumbsForLoggedInCharacters := 0
    app._JSON := JsonMergeNoOverwrite(JSON.Load(default_JSON), JSON.Load(JSON.Dump(app._JSON)))
    AssertEqual(0, app.ShiftThumbsForLoggedInCharacters)
    app._JSON["_Profiles"]["ShiftTest"] := Map("Thumbnails Behavior", Map())
    app.ProfileThumbnailsBehavior := "ShiftTest"
    AssertEqual(1, app.ShiftThumbsForLoggedInCharacters)
    app.ProfileThumbnailsBehavior := "Default"
    AssertEqual(0, app.ShiftThumbsForLoggedInCharacters)
    app.SetState()
    app.MainFrame := Gui()
    app.MainFrame.Group := Map()
    try {
        app.ThumbnailsBehavior_Ctrl()
        AssertEqual(0, app.MainFrame["ShiftThumbsForLoggedInCharacters"].Value)
        app.MainFrame["ShiftThumbsForLoginScreen"].GetPos(, &loginY, , &loginH)
        app.MainFrame["ShiftThumbsForLoggedInCharacters"].GetPos(&shiftX, &shiftY, , &shiftH)
        app.MainFrame["ShiftThumbsCollisionCheck"].GetPos(, &collisionY)
        AssertTrue(shiftY >= loginY + loginH && shiftY + shiftH <= collisionY)
        controls := app.MainFrame.Group["Thumbnails Behavior"]
        for index, control in controls {
            if control.Name = "ShiftThumbsForLoggedInCharacters" {
                controls[index - 1].GetPos(&labelX, , &labelW)
                AssertTrue(labelX + labelW <= shiftX,
                    Format("The new label must fit before its checkbox: left={1}, width={2}, checkbox={3}; DPI={4}.", labelX, labelW, shiftX, A_ScreenDPI))
            }
        }
        app.MainFrame["PreserveThumbPosOnLogout"].GetPos(, &bottom, , &height)
        AssertTrue(bottom + height <= app.guiHeight - 10)
    } finally {
        app.MainFrame.Destroy()
    }
}

TestRunner.Register("Login-to-character title changes shift unsaved thumbnails even when login shifting is off", LoggedInRenamePlacementTest)
LoggedInRenamePlacementTest() {
    previous := A_DetectHiddenWindows
    DetectHiddenWindows(true)
    app := ThumbnailShiftFixture(), sources := []
    try {
        app.ShiftThumbsForLoginScreen := false
        Loop 2 {
            source := app.CreateSource("EVE")
            sources.Push(source)
            app.EVE_WIN_Created(source.Hwnd, "EVE")
        }
        Loop 2 {
            source := sources[A_Index]
            source.Title := "EVE - Login placement " A_Index
            app.EVENameChange(source.Hwnd, source.Title)
        }
        WinGetPos(&x1, &y1, &w1, &h1, app.ThumbWindows.%sources[1].Hwnd%["Window"].Hwnd)
        WinGetPos(&x2, &y2, &w2, &h2, app.ThumbWindows.%sources[2].Hwnd%["Window"].Hwnd)
        AssertFalse(x1 < x2 + w2 && x2 < x1 + w1 && y1 < y2 + h2 && y2 < y1 + h1)
        AssertEqual(0, app.ThumbnailPositions.Count, "Automatic placement must not rewrite saved layouts.")
    } finally {
        app.Cleanup(sources)
        DetectHiddenWindows(previous)
    }
}

TestRunner.Register("Unsaved previews reserve saved positions of clients still waiting for startup", LoggedInReservedPlacementTest)
LoggedInReservedPlacementTest() {
    previous := A_DetectHiddenWindows
    DetectHiddenWindows(true)
    app := ThumbnailShiftFixture(), sources := []
    try {
        saved := app.CreateSource("EVE - Reserved placement")
        unsaved := app.CreateSource("EVE - Unsaved placement")
        sources.Push(saved, unsaved)
        app.Profiles["Default"]["Thumbnail Positions"]["EVE - Reserved placement"] := Map("x", 100, "y", 100, "width", 100, "height", 80)
        app.EVE_WIN_Created(unsaved.Hwnd, unsaved.Title)
        AssertShiftPosition(app, unsaved, 200, 100)
        app.EVE_WIN_Created(saved.Hwnd, saved.Title)
        AssertShiftPosition(app, saved, 100, 100)
        AssertShiftPosition(app, unsaved, 200, 100)
    } finally {
        app.Cleanup(sources)
        DetectHiddenWindows(previous)
    }
}

TestRunner.Register("A later login restoring saved geometry relocates overlapping unsaved previews", LoggedInLateSavedPlacementTest)
LoggedInLateSavedPlacementTest() {
    previous := A_DetectHiddenWindows
    DetectHiddenWindows(true)
    app := ThumbnailShiftFixture(), sources := []
    try {
        app.ShiftThumbsForLoginScreen := false
        saved := app.CreateSource("EVE")
        unsaved := app.CreateSource("EVE - Late saved neighbor")
        sources.Push(saved, unsaved)
        app.EVE_WIN_Created(saved.Hwnd, "EVE")
        app.EVE_WIN_Created(unsaved.Hwnd, unsaved.Title)
        AssertShiftPosition(app, unsaved, 200, 100)
        saved.Title := "EVE - Late saved placement"
        app.Profiles["Default"]["Thumbnail Positions"][saved.Title] := Map("x", 200, "y", 100, "width", 100, "height", 80)
        app.EVENameChange(saved.Hwnd, saved.Title)
        app.EVE_WIN_Created(saved.Hwnd, saved.Title)
        AssertShiftPosition(app, saved, 200, 100)
        AssertShiftPosition(app, unsaved, 100, 100)
    } finally {
        app.Cleanup(sources)
        DetectHiddenWindows(previous)
    }
}
