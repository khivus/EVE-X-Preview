TestRunner.Register("RunnerPaint regression log readers cannot take away the writer", RunnerPaintLogSharingTest)
RunnerPaintLogSharingTest() {
    path := A_Temp "\evex-log-sharing-" DllCall("GetCurrentProcessId") ".txt"
    runner := TestRunnerClass(path), reader := 0
    try {
        runner.Log("first")
        ; A live snapshot reader must share the existing writer's access.
        ; Before the fix this incompatible reader opened successfully and
        ; made the next Log call fail after its short retry window.
        try reader := FileOpen(path, "r-w", "UTF-8")
        catch OSError as err {
            if err.Number != 32
                throw err
        }
        runner.Log("second")
        if IsObject(reader)
            reader.Close(), reader := 0
        snapshot := FileOpen(path, "r", "UTF-8")
        try AssertEqual("first`nsecond`n", snapshot.Read())
        finally snapshot.Close()
    } finally {
        if IsObject(reader)
            reader.Close()
        if HasMethod(runner, "CloseLog")
            runner.CloseLog()
        FileDelete(path)
    }
}

TestRunner.Register("RunnerPaint regression live snapshots see complete Unicode log lines", RunnerPaintLiveLogTest)
RunnerPaintLiveLogTest() {
    path := A_Temp "\evex-live-log-" DllCall("GetCurrentProcessId") ".txt"
    runner := TestRunnerClass(path), reader := FileOpen(path, "r", "UTF-8")
    try {
        loop 100 {
            line := "Snapshot " A_Index " — 日本語"
            runner.Log(line)
            AssertEqual(line "`n", reader.Read(), "A shared reader must see the flushed line while the writer remains open.")
        }
        runner.CloseLog()
        runner.CloseLog()
        AssertThrows(() => runner.Log("after close"), "closed")
    } finally {
        reader.Close()
        runner.CloseLog()
        FileDelete(path)
    }
}

TestRunner.Register("RunnerPaint regression destroyed static paint cancels and restores its DC", RunnerPaintDestroyedControlTest.Bind(false))
TestRunner.Register("RunnerPaint regression native paint callback tolerates control destruction", RunnerPaintDestroyedControlTest.Bind(true))
RunnerPaintDestroyedControlTest(viaCallback) {
    theme := SettingsTheme({closed: false}, "dark"), window := Gui()
    control := window.AddText("w100 h30", "Paint me")
    control.ThemePreserve := false, control.ThemeBackground := "202124", control.ThemeColor := "FFFFFF"
    handle := control.Hwnd, dc := DllCall("CreateCompatibleDC", "Ptr", 0, "Ptr")
    state := {window: window, destroyed: false}
    callback := CallbackCreate(RunnerPaintDestroyOnFont.Bind(state),, 6)
    DllCall("comctl32\SetWindowSubclass", "Ptr", handle, "Ptr", callback, "UPtr", 99, "UPtr", 0)
    originalColor := DllCall("GetTextColor", "Ptr", dc, "UInt")
    originalCritical := A_IsCritical
    try {
        if viaCallback {
            theme.statics[handle] := control
            theme.StaticProc(handle, 0x318, dc, 0, 1, 0)
        } else
            theme.PaintStatic(control, dc)
        AssertTrue(state.destroyed, "Destroy the real native control synchronously during WM_GETFONT.")
        AssertFalse(DllCall("IsWindow", "Ptr", handle))
        AssertEqual(originalColor, DllCall("GetTextColor", "Ptr", dc, "UInt"), "Cancelled paint must restore the caller's DC.")
        AssertEqual(originalCritical, A_IsCritical, "Paint cancellation must restore thread interruptibility.")
    } finally {
        if DllCall("IsWindow", "Ptr", handle)
            DllCall("comctl32\RemoveWindowSubclass", "Ptr", handle, "Ptr", callback, "UPtr", 99)
        window.Destroy()
        CallbackFree(callback)
        DllCall("DeleteDC", "Ptr", dc)
        theme.app.closed := true
        theme.Dispose()
    }
}

TestRunner.Register("RunnerPaint regression live control errors propagate with drawing state restored", RunnerPaintLiveControlErrorTest)
RunnerPaintLiveControlErrorTest() {
    theme := SettingsTheme({closed: false}, "dark"), window := Gui()
    control := window.AddText("w100 h30", "Paint me")
    control.ThemePreserve := false, control.ThemeBackground := "202124", control.ThemeColor := "FFFFFF"
    handle := control.Hwnd, dc := DllCall("CreateCompatibleDC", "Ptr", 0, "Ptr")
    originalColor := DllCall("GetTextColor", "Ptr", dc, "UInt"), originalCritical := A_IsCritical
    theme.statics[handle] := control
    try {
        ; First prove ordinary painting still works, then inject a renderer bug.
        theme.StaticProc(handle, 0x318, dc, 0, 1, 0)
        theme.DefineProp("Text", {Call: RunnerPaintThrow})
        AssertThrows(() => theme.StaticProc(handle, 0x318, dc, 0, 1, 0), "renderer failure")
        AssertTrue(DllCall("IsWindow", "Ptr", handle))
        AssertEqual(originalColor, DllCall("GetTextColor", "Ptr", dc, "UInt"))
        AssertEqual(originalCritical, A_IsCritical)
    } finally {
        window.Destroy()
        DllCall("DeleteDC", "Ptr", dc)
        theme.app.closed := true
        theme.Dispose()
    }
}
RunnerPaintThrow(*) {
    throw Error("Expected renderer failure on a live control.")
}
RunnerPaintDestroyOnFont(state, hwnd, msg, wParam, lParam, subclass, data) {
    if msg = 0x31 && !state.destroyed {
        state.destroyed := true
        state.window.Destroy()
        return 0
    }
    return DllCall("comctl32\DefSubclassProc", "Ptr", hwnd, "UInt", msg, "Ptr", wParam, "Ptr", lParam, "Ptr")
}
