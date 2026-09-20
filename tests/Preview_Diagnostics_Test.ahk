TestRunner.Register("Preview startup waits for stable sources and spaces registrations", PreviewQueueTest)
PreviewQueueTest() {
    q := PreviewStartupQueue(0, true)
    q.Observe(11, "first", 640, 480, 0)
    q.Observe(12, "second", 640, 480, 0)
    AssertEqual(0, q.Next(2999))
    AssertEqual(11, q.Next(3000))
    q.Remove(11)
    AssertEqual(0, q.Next(3149))
    q.Observe(12, "second", 800, 600, 3100)
    AssertEqual(0, q.Next(3500), "Resized sources must settle again.")
    AssertEqual(12, q.Next(3600))
    q.Failed(12, 3600)
    AssertEqual(0, q.Next(8599))
    AssertEqual(12, q.Next(8600))
    q.Failed(12, 8600)
    AssertEqual(0, q.Next(18599))
    AssertEqual(12, q.Next(18600))
    q.Failed(12, 18600)
    AssertEqual(0, q.Next(33599))
    AssertEqual(12, q.Next(33600), "Recovery continues after three failures.")
    Loop 10
        q.Failed(12, 33600)
    AssertEqual(63600, q.items[12].retryAt, "Retry delay is capped at 30 seconds.")
    AssertEqual(6, q.items[12].failures)
    q.Remove(12)
    q.Observe(13, "empty", 0, 0, 0)
    AssertEqual(0, q.Next(999999))
    q.Remove(12)
    q.Remove(13)
    AssertEqual(0, q.items.Count)
}

TestRunner.Register("Live DWM wrapper uses aligned fields and releases registrations", LiveThumbnailTest)
LiveThumbnailTest() {
    src := Gui(, "Preview diagnostic source")
    dst := Gui(, "Preview diagnostic destination")
    live := 0
    baseline := LiveThumb.OBJ_COUNTER
    try {
        src.Show("Hide w160 h100")
        dst.Show("Hide w80 h50")
        live := LiveThumb(src.Hwnd, dst.Hwnd)
        live.Source := [0, 0, 160, 100]
        live.Destination := [0, 0, 80, 50]
        live.Opacity := 255
        live.Visible := true
        live.SourceClientAreaOnly := true
        AssertEqual(1, NumGet(live.THUMB_UPD_PROP_PTR, 40, "Int"))
        AssertEqual(1, NumGet(live.THUMB_UPD_PROP_PTR, 44, "Int"))
        AssertEqual(255, NumGet(live.THUMB_UPD_PROP_PTR, 36, "UChar"))
        AssertTrue(live.Update())
        AssertEqual(baseline + 1, LiveThumb.OBJ_COUNTER)
        live.Source := [0, 0, 160, 100]
        live.Destination := [0, 0, 80, 50]
        AssertFalse(live.THUMB_PENDING_UPDATE, "Identical geometry must not submit another DWM update.")
        live.Visible := false
        live.Update()
        AssertEqual(0, live.Visible)
        AssertEqual(1, live.SourceClientAreaOnly, "Visibility must not overwrite the client-area flag.")
        live.Close()
        live.Close()
        AssertEqual(baseline, LiveThumb.OBJ_COUNTER)
        AssertFalse(live.Update())
        AssertThrows(() => LiveThumb(0, dst.Hwnd), "DwmRegisterThumbnail")
        AssertEqual(baseline, LiveThumb.OBJ_COUNTER, "Failed registration must not increment the live count.")
    } finally {
        if IsObject(live)
            live.Close()
        src.Destroy()
        dst.Destroy()
    }
}

TestRunner.Register("Program events persist with debug disabled and errors export snapshots", ProgramLogTest)
ProgramLogTest() {
    directory := A_Temp "\EVE-X-Preview-test-" A_TickCount
    previousText := ProgramLog.Text
    previousPending := ProgramLog.Pending
    try {
        ProgramLog.Text := "", ProgramLog.Pending := ""
        ProgramLog.Init(directory)
        app := CombatEventsFixture()
        app.DebugMode := false
        app.debugToolTipText := ""
        app.debugToolTipText .= "diagnostic first`n"
        app.debugToolTipText .= "diagnostic second`n"
        app.debugToolTip()
        AssertEqual("", app.debugToolTipText)
        AssertTrue(InStr(ProgramLog.Text, "diagnostic first"))
        AssertEqual(2, StrSplit(ProgramLog.Text, "diagnostic first").Length, "Appending debug text must not duplicate previous entries.")
        ProgramLog.Flush()
        AssertTrue(InStr(FileRead(directory "\current.log"), "diagnostic second"))
        ProgramLog.LastErrorTick := -60000
        AssertThrows(() => Error_Handler(Error("intentional diagnostic test"), "Return"), "intentional diagnostic test")
        AssertTrue(FileExist(ProgramLog.LastErrorExport))
        AssertTrue(InStr(FileRead(ProgramLog.LastErrorExport), "intentional diagnostic test"))
        ProgramLog.Export(directory "\manual.log")
        AssertTrue(InStr(FileRead(directory "\manual.log"), "diagnostic first"))
        FileAppend(Format("{:01050000}", 0), directory "\current.log")
        ProgramLog.Add("rotation test")
        ProgramLog.Flush()
        AssertTrue(FileExist(directory "\previous.log"))
        ProgramLog.Add(Format("{:0300000}", 0))
        AssertTrue(StrLen(ProgramLog.Text) <= ProgramLog.MaxChars)
        AssertTrue(StrLen(ProgramLog.Pending) <= ProgramLog.MaxChars)
    } finally {
        if ProgramLog.HasOwnProp("FlushTimer")
            SetTimer(ProgramLog.FlushTimer, 0)
        if ProgramLog.HasOwnProp("ExitHandler")
            OnExit(ProgramLog.ExitHandler, 0)
        ProgramLog.Directory := ""
        ProgramLog.Text := previousText, ProgramLog.Pending := previousPending
        loop files directory "\*.log"
            FileDelete(A_LoopFileFullPath)
        DirDelete(directory)
    }
}

TestRunner.Register("Log viewer opens a full buffer without truncation", ProgramLogViewerTest)
ProgramLogViewerTest() {
    previous := ProgramLog.Text
    viewer := 0
    try {
        tail := "`r`nLast event: тест"
        full := Format("{:0" (ProgramLog.MaxChars - StrLen(tail)) "}", 0) tail
        multiline := ""
        loop 6000
            multiline .= "[INFO] Event " A_Index ": тест`r`n"
        for snapshot in ["", "A short log", full, multiline] {
            ProgramLog.Text := snapshot
            viewer := ProgramLog.CreateViewer()
            AssertEqual(StrLen(snapshot), DllCall("GetWindowTextLengthW", "Ptr", viewer["Log"].Hwnd), "Native log display must not truncate snapshots.")
            ; AHK's multiline Edit.Value normalizes CRLF to LF on retrieval.
            AssertEqual(StrReplace(snapshot, "`r`n", "`n"), viewer["Log"].Value)
            viewer.Destroy()
            viewer := 0
        }
    } finally {
        if IsObject(viewer)
            viewer.Destroy()
        ProgramLog.Text := previous
    }
}

TestRunner.Register("About exposes diagnostics and preview pause within the panel", PreviewAboutLayoutTest)
PreviewAboutLayoutTest() {
    app := CombatEventsFixture()
    app.ProfileOverride()
    app.SetState()
    app.MainFrame := Gui()
    app.MainFrame.Group := Map()
    try {
        app.About_Ctrl()
        for name in ["showProgramLogBtn", "exportProgramLogBtn", "pausePreviewsBtn"] {
            app.MainFrame[name].GetPos(&x, &y, &w, &h)
            AssertTrue(x + w <= app.contentW && y + h <= app.guiHeight - 10)
        }
        app.ToggleLivePreviews()
        AssertTrue(app.livePreviewsPaused)
        AssertEqual("Resume Live Previews", app.MainFrame["pausePreviewsBtn"].Text)
        app.ToggleLivePreviews()
        AssertFalse(app.livePreviewsPaused)
    } finally {
        app.MainFrame.Destroy()
    }
}

class PreviewLifecycleFixture extends CombatEventsFixture {
    failOverlay := false
    __New() {
        super.__New()
        this.ProfileOverride()
        this.margins := Buffer(16, 0)
        this.debugToolTipMethod := (*) => 0
        this.debugToolTipDelay := 60000
        this.ClickThroughActive := false
        this.HidedThumbs := Map()
        this.DisabledChars := Map()
        this.QuickGroupChars := Map()
        this.TrackClientPossitions := false
        this.HideThumbnailsOnLostFocus := true
        this.EVEExe := "ahk_exe nonexistent-preview-test.exe"
        this.monitoringInitialized := false
        this.SlowThumbnailCreation := true
        this.previewQueue := PreviewStartupQueue(A_TickCount, true)
    }
    RegisterHotkeys(*) {
    }
    RegisterNonEVEHotkeys(*) {
    }
    AddThumbnailTextControls(args*) {
        if this.failOverlay
            throw Error("Injected overlay failure")
        super.AddThumbnailTextControls(args*)
    }
}

class PreviewCloseFixture extends PreviewLifecycleFixture {
    failCloseHwnd := 0
    failCreateHwnd := 0
    DisposePreview(hwnd) {
        if hwnd = this.failCloseHwnd
            throw Error("Injected cleanup failure")
        super.DisposePreview(hwnd)
    }
    EVE_WIN_Created(hwnd, title) {
        if hwnd = this.failCreateHwnd
            throw Error("Injected registration failure")
        super.EVE_WIN_Created(hwnd, title)
    }
    ActivateEVEWindow(hwnd?, title?) {
        this.clicked := hwnd
    }
}

TestRunner.Register("Client cleanup isolates failures and preserves surviving previews and clicks", PreviewCloseIsolationTest)
PreviewCloseIsolationTest() {
    oldHidden := A_DetectHiddenWindows
    DetectHiddenWindows(true)
    app := PreviewCloseFixture()
    first := Gui(), second := Gui(), survivor := Gui()
    firstHwnd := first.Hwnd, secondHwnd := second.Hwnd, liveHwnd := survivor.Hwnd
    baseline := LiveThumb.OBJ_COUNTER
    try {
        for source in [first, second, survivor] {
            source.Show("Hide w160 h100")
            app.EVE_WIN_Created(source.Hwnd, "Close test " source.Hwnd)
        }
        firstPreview := app.ThumbWindows.%firstHwnd%["Window"].Hwnd
        secondPreview := app.ThumbWindows.%secondHwnd%["Window"].Hwnd
        live := app.ThumbWindows.%liveHwnd%["Thumbnail"]
        liveId := live.THUMB_ID
        app.QuickGroupChars[firstHwnd] := "Closed pilot"
        app.DisabledChars[firstHwnd] := "Closed pilot"
        app.HidedThumbs[firstHwnd] := "Closed pilot"
        first.Destroy(), second.Destroy()
        app.failCloseHwnd := firstHwnd
        app.DestroyThumbnailsToggle := 0
        AssertThrows(() => app.EvEWindowDestroy(firstHwnd), "Injected cleanup failure")
        AssertEqual(1, app.DestroyThumbnailsToggle, "Explicit cleanup failures must release the gate.")
        app.DestroyThumbnailsToggle := 0
        app.EvEWindowDestroy()
        AssertEqual(1, app.DestroyThumbnailsToggle)
        AssertTrue(app.ThumbWindows.HasProp(firstHwnd), "Failed disposal remains available for retry.")
        AssertFalse(app.ThumbWindows.HasProp(secondHwnd), "One failure must not block another closed client.")
        AssertFalse(app.ThumbHwnd_EvEHwnd.Has(secondPreview))
        app.failCloseHwnd := 0
        app.EvEWindowDestroy()
        AssertFalse(app.ThumbWindows.HasProp(firstHwnd))
        AssertFalse(app.ThumbHwnd_EvEHwnd.Has(firstPreview))
        AssertEqual(0, app.QuickGroupChars.Count)
        AssertEqual(0, app.DisabledChars.Count)
        AssertEqual(0, app.HidedThumbs.Count)
        AssertEqual(baseline + 1, LiveThumb.OBJ_COUNTER)
        AssertEqual(liveId, live.THUMB_ID)
        live.Opacity := 254
        AssertTrue(live.Update(), "Closing other sources must not invalidate the surviving registration.")
        app.ThumbnailsInteractionsMap := Map(1, "ActivateThumbnail")
        app._OnMessage(1, 0, 0x201, app.ThumbWindows.%liveHwnd%["Window"].Hwnd)
        AssertEqual(liveHwnd, app.clicked)
    } finally {
        app.failCloseHwnd := 0
        for hwnd in [firstHwnd, secondHwnd, liveHwnd]
            app.DisposePreview(hwnd)
        SetTimer(app.debugToolTipMethod, 0)
        first.Destroy(), second.Destroy(), survivor.Destroy()
        DetectHiddenWindows(oldHidden)
    }
}

TestRunner.Register("One preview registration failure does not delay other ready previews", PreviewRecoveryIsolationTest)
PreviewRecoveryIsolationTest() {
    oldHidden := A_DetectHiddenWindows
    DetectHiddenWindows(true)
    app := PreviewCloseFixture()
    app.previewQueue := PreviewStartupQueue()
    first := Gui(), second := Gui()
    try {
        first.Show("Hide w160 h100"), second.Show("Hide w160 h100")
        app.failCreateHwnd := first.Hwnd
        app.QueuePreview(first.Hwnd, "Failing source")
        app.QueuePreview(second.Hwnd, "Healthy source")
        app.ProcessPreviewQueue([first.Hwnd, second.Hwnd])
        AssertTrue(app.ThumbWindows.HasProp(second.Hwnd), "Healthy previews start in the same poll.")
        AssertEqual(1, app.previewQueue.items[first.Hwnd].failures)
        app.failCreateHwnd := 0
        app.previewQueue.items[first.Hwnd].retryAt := 0
        app.ProcessPreviewQueue([first.Hwnd, second.Hwnd])
        AssertTrue(app.ThumbWindows.HasProp(first.Hwnd))
        AssertEqual(0, app.previewQueue.items.Count)
    } finally {
        app.DisposePreview(first.Hwnd), app.DisposePreview(second.Hwnd)
        SetTimer(app.debugToolTipMethod, 0)
        first.Destroy(), second.Destroy()
        DetectHiddenWindows(oldHidden)
    }
}

TestRunner.Register("Fast previews have no startup waits and slow mode persists", FastPreviewModeTest)
FastPreviewModeTest() {
    q := PreviewStartupQueue(0)
    q.Observe(1, "one", 100, 100, 0)
    q.Observe(2, "two", 100, 100, 0)
    AssertEqual(1, q.Next(0))
    q.Remove(1)
    AssertEqual(2, q.Next(0), "Fast mode must not space ready previews.")
    q.Failed(2, 0)
    AssertEqual(0, q.Next(4999), "Failure backoff remains in fast mode.")
    app := CombatEventsFixture()
    app.ProfileOverride()
    AssertEqual(0, app.SlowThumbnailCreation)
    app._JSON["_Profiles"][app.ProfileOther]["Other"].Delete("SlowThumbnailCreation")
    AssertEqual(0, app.SlowThumbnailCreation, "Older profiles default to fast mode.")
    app.SlowThumbnailCreation := 1
    app._JSON := JsonMergeNoOverwrite(JSON.Load(default_JSON), JSON.Load(JSON.Dump(app._JSON)))
    AssertEqual(1, app.SlowThumbnailCreation)
    app.SetState()
    app.MainFrame := Gui()
    app.MainFrame.Group := Map()
    try {
        app.Other_Ctrl()
        app.MainFrame["SlowThumbnailCreation"].GetPos(&x, &y, &w, &h)
        app.MainFrame["SwitchLangOnErr"].GetPos(, &languageY)
        AssertTrue(y + h <= languageY, "Slow mode must appear above the language option.")
        AssertEqual(1, app.MainFrame["SlowThumbnailCreation"].Value)
        app.MainFrame["UpdateThumbnails"].GetPos(, &bottom, , &buttonHeight)
        AssertTrue(bottom + buttonHeight <= app.guiHeight - 10)
    } finally {
        app.MainFrame.Destroy()
    }
}

TestRunner.Register("Fast mode creates all ready previews in one polling pass", FastPreviewBatchTest)
FastPreviewBatchTest() {
    previousHidden := A_DetectHiddenWindows
    DetectHiddenWindows(true)
    app := PreviewLifecycleFixture()
    app.SlowThumbnailCreation := false
    app.previewQueue := PreviewStartupQueue()
    first := Gui(, "Fast preview one"), second := Gui(, "Fast preview two")
    try {
        first.Show("Hide w160 h100"), second.Show("Hide w160 h100")
        app.QueuePreview(first.Hwnd, "Fast preview one")
        app.QueuePreview(second.Hwnd, "Fast preview two")
        app.ProcessPreviewQueue([first.Hwnd, second.Hwnd])
        AssertTrue(app.ThumbWindows.HasProp(first.Hwnd))
        AssertTrue(app.ThumbWindows.HasProp(second.Hwnd))
        AssertEqual(0, app.previewQueue.items.Count)
    } finally {
        app.DisposePreview(first.Hwnd), app.DisposePreview(second.Hwnd)
        SetTimer(app.debugToolTipMethod, 0)
        first.Destroy(), second.Destroy()
        DetectHiddenWindows(previousHidden)
    }
}

TestRunner.Register("Runtime diagnostic counters aggregate and reset without per-frame logs", RuntimeLogCountersTest)
RuntimeLogCountersTest() {
    ProgramLog.Activity := Map()
    ProgramLog.ActivityAt := A_TickCount
    before := ProgramLog.Text
    loop 20
        ProgramLog.Count("test-DWM-updates")
    ProgramLog.ReportActivity(ProgramLog.ActivityAt + 9999)
    AssertEqual(before, ProgramLog.Text)
    ProgramLog.ReportActivity(ProgramLog.ActivityAt + 10000)
    AssertTrue(InStr(ProgramLog.Text, "test-DWM-updates=20"))
    AssertTrue(InStr(ProgramLog.Text, "GDI="))
    AssertEqual(0, ProgramLog.Activity.Count)
    ProgramLog.ActivityAt := A_TickCount
}

TestRunner.Register("Preview lifecycle cleans partial failures and resumes paused registrations", PreviewLifecycleTest)
PreviewLifecycleTest() {
    previousHidden := A_DetectHiddenWindows
    DetectHiddenWindows(true)
    app := PreviewLifecycleFixture()
    source := Gui(, "Preview lifecycle source")
    baseline := LiveThumb.OBJ_COUNTER
    try {
        source.Show("Hide w160 h100")
        hwnd := source.Hwnd
        app.failOverlay := true
        beforeWindows := WinGetList("ahk_pid " ProcessExist()).Length
        AssertThrows(() => app.Create_Thumbnail(hwnd, "Partial preview test"), "Injected overlay failure")
        AssertEqual(baseline, LiveThumb.OBJ_COUNTER)
        AssertEqual(beforeWindows, WinGetList("ahk_pid " ProcessExist()).Length, "Partial creation must destroy its GUI windows.")
        app.failOverlay := false
        app.QueuePreview(hwnd, "Preview lifecycle source")
        app.previewQueue.nextAt := A_TickCount
        app.previewQueue.items[hwnd].stableAt := A_TickCount - 1000
        app.ProcessPreviewQueue([hwnd])
        AssertTrue(app.ThumbWindows.HasProp(hwnd), "Queued source must create a preview.")
        AssertEqual(baseline + 1, LiveThumb.OBJ_COUNTER)
        app.ToggleLivePreviews()
        AssertEqual(baseline, LiveThumb.OBJ_COUNTER, "Pause must unregister rather than merely hide.")
        app.ProcessPreviewQueue([hwnd])
        AssertEqual(baseline, LiveThumb.OBJ_COUNTER)
        app.ToggleLivePreviews()
        app.ProcessPreviewQueue([hwnd])
        AssertEqual(baseline, LiveThumb.OBJ_COUNTER, "Resume retains its grace period.")
        app.previewQueue.nextAt := A_TickCount
        app.previewQueue.items[hwnd].stableAt := A_TickCount - 1000
        app.ProcessPreviewQueue([hwnd])
        AssertEqual(baseline + 1, LiveThumb.OBJ_COUNTER)
        app.DisposePreview(hwnd)
        AssertEqual(baseline, LiveThumb.OBJ_COUNTER)
        AssertEqual(0, app.ThumbHwnd_EvEHwnd.Count)
    } finally {
        app.DisposePreview(source.Hwnd)
        SetTimer(app.debugToolTipMethod, 0)
        source.Destroy()
        DetectHiddenWindows(previousHidden)
    }
}
