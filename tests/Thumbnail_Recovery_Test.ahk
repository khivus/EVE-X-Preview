TestRunner.Register("Thumbnail recovery preserves second third negative and spanning monitor placements", ThumbnailRecoveryVisibleTest)
ThumbnailRecoveryVisibleTest() {
    monitors := [{left: 0, top: 0, right: 1920, bottom: 1080},
        {left: 1920, top: 0, right: 3840, bottom: 1080},
        {left: -1280, top: -200, right: 0, bottom: 824},
        {left: 0, top: -1080, right: 1920, bottom: 0}]
    for rect in [[50, 50, 200, 120], [2500, 200, 300, 180], [-1200, -100, 200, 120],
        [100, -1000, 200, 120], [1820, 100, 200, 120], [-100, 100, 200, 120],
        [100, -60, 200, 120], [1720, 960, 200, 120]]
        AssertFalse(ThumbnailRecovery.Position(rect[1], rect[2], rect[3], rect[4], monitors),
            "A thumbnail fully covered by connected displays must stay in place.")
}

TestRunner.Register("Thumbnail recovery handles every screen edge disconnected displays and desktop gaps", ThumbnailRecoveryEdgesTest)
ThumbnailRecoveryEdgesTest() {
    monitors := [{left: 0, top: 0, right: 1920, bottom: 1080},
        {left: 1920, top: 0, right: 3840, bottom: 1080},
        {left: -1280, top: 0, right: 0, bottom: 1024}]
    for item in [
        {rect: [-1350, 100, 200, 120], x: -1280, y: 100},
        {rect: [3750, 100, 200, 120], x: 3640, y: 100},
        {rect: [2100, -50, 200, 120], x: 2100, y: 0},
        {rect: [2100, 1030, 200, 120], x: 2100, y: 960},
        {rect: [6000, 400, 200, 120], x: 3640, y: 400},
        {rect: [2000, -2000, 200, 120], x: 2000, y: 0},
        {rect: [3840, 1080, 200, 120], x: 3640, y: 960}] {
        rect := item.rect
        target := ThumbnailRecovery.Position(rect[1], rect[2], rect[3], rect[4], monitors)
        AssertTrue(IsObject(target))
        AssertEqual(item.x, target.x)
        AssertEqual(item.y, target.y)
        AssertFalse(ThumbnailRecovery.Position(target.x, target.y, rect[3], rect[4], monitors))
    }
    ; Being inside the virtual desktop bounding box is insufficient in an L layout.
    staggered := [monitors[1], {left: 1920, top: 1080, right: 3840, bottom: 2160}]
    AssertTrue(IsObject(ThumbnailRecovery.Position(2100, 100, 200, 120, staggered)))
    AssertTrue(IsObject(ThumbnailRecovery.Position(1850, 1000, 200, 120, staggered)))
    gap := [monitors[1], {left: 2000, top: 0, right: 3920, bottom: 1080}]
    AssertTrue(IsObject(ThumbnailRecovery.Position(1900, 100, 200, 120, gap)))
    ; Duplicate/mirrored rectangles cannot hide an uncovered area by double-counting it.
    AssertTrue(IsObject(ThumbnailRecovery.Position(1850, 100, 200, 120, [monitors[1], monitors[1]])))
}

TestRunner.Register("Thumbnail recovery keeps oversized previews reachable and tolerates no displays", ThumbnailRecoveryOversizedTest)
ThumbnailRecoveryOversizedTest() {
    monitors := [{left: -800, top: 0, right: 0, bottom: 600},
        {left: 0, top: 0, right: 1920, bottom: 1080}]
    target := ThumbnailRecovery.Position(-900, 100, 1000, 700, monitors)
    AssertEqual(0, target.x, "Prefer a display that fits the thumbnail.")
    AssertEqual(100, target.y)
    target := ThumbnailRecovery.Position(4000, 2000, 5000, 3000, monitors)
    AssertEqual(0, target.x)
    AssertEqual(0, target.y)
    AssertFalse(ThumbnailRecovery.Position(4000, 2000, 200, 120, []))
}

class ThumbnailRecoveryFixture extends CombatEventsFixture {
    __New() {
        super.__New()
        this.ProfileOverride()
        this.guis := []
        this.ThumbWindows := {}
        this.Save_Settings_Delay_Timer := (*) => 0
        this.monitors := [{left: 0, top: 0, right: 1000, bottom: 1000},
            {left: 1000, top: 0, right: 2000, bottom: 1000},
            {left: -1000, top: 0, right: 0, bottom: 1000}]
        this.AddThumb(1, 2500, 200)
        this.AddThumb(2, 1200, 200)
        this.AddThumb(3, -800, 200)
        this.AddThumb(4, 950, 200)
        this.LockThumbnailPositions := true
        this.AutoSaveThumbnailPositions := false
    }
    AddThumb(id, x, y) {
        thumb := Map()
        for type in ["Window", "TextOverlay", "Border", "EventBorder"] {
            preview := Gui("-Caption +ToolWindow -DPIScale", "EVE - Recovery " id)
            preview.Show("Hide w100 h80")
            WinMove(x, y, 100, 80, preview.Hwnd)
            thumb[type] := preview
            this.guis.Push(preview)
        }
        this.ThumbWindows.%id% := thumb
    }
    ThumbnailRecoveryMonitors() => this.monitors
    Cleanup() {
        SetTimer(this.Save_Settings_Delay_Timer, 0)
        for preview in this.guis
            preview.Destroy()
    }
}

TestRunner.Register("Thumbnail recovery moves and saves only lost stacks even when hidden and locked", ThumbnailRecoveryWindowsTest)
ThumbnailRecoveryWindowsTest() {
    previousHidden := A_DetectHiddenWindows
    DetectHiddenWindows(true)
    app := ThumbnailRecoveryFixture()
    source := Gui("-Caption")
    source.Show("Hide x500 y500 w100 h80")
    try {
        app.ThumbWindows.%1%["Thumbnail"] := {Hwnd: source.Hwnd}
        sourceBefore := app.CaptureThumbnailDragWindows(1, false)
        AssertFalse(sourceBefore.Has(source.Hwnd), "Source clients must never be moved.")
        app.ThumbnailPositions["EVE - Recovery 2"] := [1200, 200, 100, 80]
        app.ThumbnailPositions["EVE - Offline"] := [-8000, 200, 100, 80]
        savedVisible := JSON.Dump(app.ThumbnailPositions["EVE - Recovery 2"])
        savedOffline := JSON.Dump(app.ThumbnailPositions["EVE - Offline"])
        DetectHiddenWindows(false)
        AssertEqual(1, app.RecoverOffscreenThumbnails())
        AssertFalse(A_DetectHiddenWindows, "Recovery must restore the caller's hidden-window setting.")
        DetectHiddenWindows(true)
        for id, thumb in app.ThumbWindows.OwnProps() {
            for type in ["Window", "TextOverlay", "Border", "EventBorder"] {
                WinGetPos(&x, &y, &w, &h, thumb[type].Hwnd)
                AssertEqual(id = 1 ? 1900 : (id = 2 ? 1200 : (id = 3 ? -800 : 950)), x)
                AssertEqual(200, y)
                AssertEqual(100, w)
                AssertEqual(80, h)
                AssertFalse(DllCall("IsWindowVisible", "Ptr", thumb[type].Hwnd))
            }
        }
        WinGetPos(&sourceX, &sourceY,,, source.Hwnd)
        AssertEqual(500, sourceX), AssertEqual(500, sourceY)
        AssertEqual(1900, app.ThumbnailPositions["EVE - Recovery 1"]["x"])
        AssertEqual(savedVisible, JSON.Dump(app.ThumbnailPositions["EVE - Recovery 2"]))
        AssertEqual(savedOffline, JSON.Dump(app.ThumbnailPositions["EVE - Offline"]))
        AssertEqual(0, app.RecoverOffscreenThumbnails(), "Repeating recovery must leave the recovered layout alone.")
    } finally {
        source.Destroy()
        app.Cleanup()
        DetectHiddenWindows(previousHidden)
    }
}

TestRunner.Register("Thumbnail recovery Layout button runs immediately and survives unrelated draft apply", ThumbnailRecoveryButtonTest)
ThumbnailRecoveryButtonTest() {
    for query in ["recover", "offscreen", "off-screen", "lost thumbnails"] {
        found := false
        for section in SettingsSchema.Search(query)
            found := found || section.id = "layout"
        AssertTrue(found, "Recovery must be discoverable through Settings search.")
    }
    owner := ThumbnailRecoveryFixture()
    workbench := SettingsWorkbenchFixture(owner)
    try {
        workbench.Navigate("layout")
        AssertTrue(IsObject(workbench.recoverThumbnailsButton))
        workbench.model.Set("ThumbnailSnap", !workbench.model.Get("ThumbnailSnap"))
        ; Dispatch the actual native button, even while the settings fixture is hidden.
        SendMessage(0x00F5, 0, 0, workbench.recoverThumbnailsButton.Hwnd)
        Sleep(40)
        AssertEqual(1900, owner.ThumbnailPositions["EVE - Recovery 1"]["x"])
        merged := workbench.model.Merge(owner._JSON)
        AssertEqual(1900, merged["_Profiles"][owner.LastUsedProfile]["Thumbnail Positions"]["EVE - Recovery 1"]["x"])
        AssertTrue(workbench.model.Dirty)
    } finally {
        workbench.Dispose()
        owner.Cleanup()
    }
}

TestRunner.Register("Thumbnail recovery modern command is exported and rejects offline hosts", ThumbnailRecoveryContractTest)
ThumbnailRecoveryContractTest() {
    descriptor := 0
    for command in SettingsContract.Commands()
        if command["id"] = "recover-offscreen-thumbnails"
            descriptor := command
    AssertTrue(IsObject(descriptor))
    AssertEqual("layout", descriptor["section"])
    AssertTrue(descriptor["runtimeOnly"])
    AssertFalse(descriptor["destructive"])
    exported := false
    for section in SettingsContract.Export()["sections"]
        if section["id"] = "layout"
            for command in section["commands"]
                exported := exported || command = descriptor["id"]
    AssertTrue(exported, "The modern editor must discover recovery from the canonical layout metadata.")
    session := SettingsSession(SettingsServiceFixture(JSON.Load(default_JSON)), 0, false)
    AssertThrows(() => session.Command("recover-offscreen-thumbnails"), "Start EVE-X-Preview")
}

TestRunner.Register("Thumbnail recovery modern command updates running profile and preserves a different profile draft", ThumbnailRecoveryModernCommandTest)
ThumbnailRecoveryModernCommandTest() {
    owner := ThumbnailRecoveryFixture()
    session := SettingsSession(owner)
    try {
        session.model.AddProfile("Recovery draft", true)
        session.model.SelectProfile("Recovery draft")
        session.model.Set("ThumbnailSnap", !session.model.Get("ThumbnailSnap"))
        draftBefore := JSON.Dump(session.model.config)
        result := session.Command("recover-offscreen-thumbnails")
        AssertTrue(InStr(result["message"], "Recovered 1"))
        AssertEqual("Default", owner.LastUsedProfile)
        AssertEqual(1900, owner.ThumbnailPositions["EVE - Recovery 1"]["x"])
        AssertEqual(draftBefore, JSON.Dump(session.model.config), "Immediate recovery must not mutate the selected profile's draft.")
        AssertTrue(session.Dirty)
        AssertEqual(0, session.model.Conflicts(owner._JSON).Length)
        merged := session.model.Merge(owner._JSON)
        AssertEqual(1900, merged["_Profiles"]["Default"]["Thumbnail Positions"]["EVE - Recovery 1"]["x"])
        AssertFalse(merged["_Profiles"]["Recovery draft"]["Thumbnail Positions"].Has("EVE - Recovery 1"))
    } finally {
        session.Dispose()
        owner.Cleanup()
    }
}
