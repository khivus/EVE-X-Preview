TestRunner.Register("Simultaneous slow log workers leave GUI timers and settings controls responsive", LogWorkerResponsiveTest)
LogWorkerResponsiveTest() {
    path := A_ScriptDir "\SlowLogWorker-" A_TickCount ".ahk"
    source := '#Requires AutoHotkey v2.0`n#SingleInstance Off`n#NoTrayIcon`nSleep(900)`nFileAppend("", A_Args[1] "\output.txt", "UTF-8")`nExitApp()`n'
    FileAppend(source, path, "UTF-8")
    game := LogScanJob(path), chat := LogScanJob(path)
    app := DPSLogFixture()
    state := {ticks: 0}
    timer := (*) => state.ticks++
    try {
        started := A_TickCount
        AssertTrue(game.Start("game`nignored`n"))
        AssertTrue(chat.Start("chat`nignored`n"))
        AssertTrue(A_TickCount - started < 500, "Starting both scans must not wait for the worker's slow disk work.")
        SetTimer(timer, 10)
        app.SetState()
        app.MainFrame := Gui()
        app.MainFrame.Group := Map()
        started := A_TickCount
        app.GameLogsMonitoring_Ctrl()
        AssertTrue(A_TickCount - started < 500, "Settings controls must build while both scans are pending.")
        AssertTrue(game.busy && chat.busy)
        Sleep(150)
        AssertTrue(state.ticks >= 4, "GUI interpreter timers must continue while scans are blocked.")
        app.MainFrame["systemTrackingEnabled"].Value := 1
        AssertEqual(1, app.MainFrame["systemTrackingEnabled"].Value)
        for job in [game, chat] {
            deadline := A_TickCount + 10000
            result := 0
            while !result && A_TickCount < deadline {
                result := job.Poll()
                Sleep(10)
            }
            AssertTrue(result && result["ok"])
        }
    } finally {
        SetTimer(timer, 0)
        if app.HasOwnProp("MainFrame")
            app.MainFrame.Destroy()
        app.Cleanup()
        game.Close(), chat.Close()
        FileDelete(path)
    }
}

TestRunner.Register("Stopping monitoring terminates pending workers and removes job files", LogWorkerCancellationTest)
LogWorkerCancellationTest() {
    path := A_ScriptDir "\CancelLogWorker-" A_TickCount ".ahk"
    FileAppend("#Requires AutoHotkey v2.0`n#SingleInstance Off`n#NoTrayIcon`nSleep(10000)`n", path, "UTF-8")
    job := LogScanJob(path)
    try {
        AssertTrue(job.Start("game`nignored`n"))
        directory := job.directory
        handle := job.process
        job.Close()
        AssertFalse(job.busy)
        AssertFalse(DirExist(directory))
        AssertEqual(0xFFFFFFFF, DllCall("WaitForSingleObject", "Ptr", handle, "UInt", 0, "UInt"))
    } finally {
        job.Close()
        FileDelete(path)
    }
}

TestRunner.Register("Worker failures are reported without waiting or opening an error dialog", LogWorkerFailureTest)
LogWorkerFailureTest() {
    job := LogScanJob()
    try {
        AssertTrue(job.Start("invalid-mode`nignored`n"))
        deadline := A_TickCount + 10000, result := 0
        while !result && A_TickCount < deadline {
            result := job.Poll()
            Sleep(10)
        }
        AssertTrue(result && !result["ok"])
        AssertFalse(job.busy)
        AssertTrue(InStr(FileRead(job.directory "\error.txt", "UTF-8"), "Unknown log scan mode"))
    } finally job.Close()
}

TestRunner.Register("Long Local sessions retain the system without reading their history in the UI", LocalChatLargeSessionTest)
LocalChatLargeSessionTest() {
    app := SystemTrackingFixture()
    try {
        path := app.WriteLog("Local_20260908_123456_111.txt", "Alice", "2026.09.08 23:28:20", app.Message("Jita"), "UTF-16")
        writer := FileOpen(path, "a", "UTF-16")
        block := ""
        Loop 1024
            block .= "[ 2026.09.08 23:28:24 ] Someone > Ordinary chat`r`n"
        Loop 64
            writer.Write(block)
        writer.Close()
        app.monitorCharacterSystems()
        AssertEqual("Jita", app.System("alice"), "The worker must recover a system announcement older than the recent tail.")
        reader := app.systemLogMonitor.readers["EVE - Alice"]
        before := reader.file.Pos
        app.systemLogMonitor.ReadUpdates(reader)
        AssertTrue(reader.file.Pos - before <= 65536, "A Local read must consume at most one bounded character chunk.")
        AssertTrue(StrLen(reader.pending) <= 8192)
    } finally app.Cleanup()
}

TestRunner.Register("Bounded combat reads continue on quiet polls until all pending damage is processed", GameLogChunkDrainTest)
GameLogChunkDrainTest() {
    previousHidden := A_DetectHiddenWindows
    DetectHiddenWindows(true)
    app := DPSLogFixture()
    now := DPSMeter.CurrentSecond()
    try {
        path := app.WriteLog("20260912_000000_12345.txt", "")
        app.startLogMonitoring(app.character, 12345)
        batch := ""
        Loop 600
            batch .= DPSLine(now, 100)
        app.Append(path, batch)
        reader := app.monitoredChars[app.character]
        before := reader["file"].Pos
        app.monitorChanges(app.character, now)
        AssertTrue(reader["file"].Pos - before <= 65536)
        AssertFalse(reader["file"].AtEOF)
        reads := 1
        while !reader["file"].AtEOF {
            app.monitorChanges(app.character, now)
            if ++reads > 100
                throw Error("A reader must drain its remaining chunks even without further writes")
        }
        AssertEqual(6000, reader["dps"].Rates(now).incoming)
    } finally {
        app.Cleanup()
        DetectHiddenWindows(previousHidden)
    }
}

TestRunner.Register("Local matches become available before the background archive scan finishes", LogWorkerPartialResultTest)
LogWorkerPartialResultTest() {
    path := A_ScriptDir "\PartialLogWorker-" A_TickCount ".ahk"
    FileAppend('#Requires AutoHotkey v2.0`n#SingleInstance Off`n#NoTrayIcon`nFileAppend("Alice``tC:\example.txt``t20261008000000``tJita``t0``t``n", A_Args[1] "\output.txt", "UTF-8")`nFileAppend("1", A_Args[1] "\progress.txt", "UTF-8")`nSleep(10000)`n', path, "UTF-8")
    discovery := LogFileDiscovery(A_ScriptDir, "chat")
    discovery.job.workerScript := path
    try {
        discovery.Refresh(false, "Alice`n")
        deadline := A_TickCount + 5000
        while !discovery.generation && A_TickCount < deadline {
            discovery.Refresh(false, "Alice`n")
            Sleep(10)
        }
        AssertTrue(discovery.busy, "The deliberately slow archive scan is still running.")
        AssertTrue(discovery.byId.Has("Alice"), "An early match must be published without waiting for process exit.")
        AssertEqual("Jita", discovery.byId["Alice"].system)
    } finally {
        discovery.Close()
        FileDelete(path)
    }
}

TestRunner.Register("Changing window order does not restart a pending Local archive scan", LocalScanWindowOrderTest)
LocalScanWindowOrderTest() {
    app := SystemTrackingFixture()
    path := A_ScriptDir "\WindowOrderLogWorker-" A_TickCount ".ahk"
    FileAppend('#Requires AutoHotkey v2.0`n#SingleInstance Off`n#NoTrayIcon`nSleep(10000)`n', path, "UTF-8")
    app.systemLogMonitor.directory := app.directory
    app.systemLogMonitor.discovery := LogFileDiscovery(app.directory, "chat")
    app.systemLogMonitor.discovery.job.workerScript := path
    try {
        app.systemLogMonitor.FindLogs(app.directory, app.active)
        handle := app.systemLogMonitor.discovery.job.process
        started := app.systemLogMonitor.discovery.job.started
        app.systemLogMonitor.FindLogs(app.directory, Map("EVE - Bob", "bob", "EVE - Alice", "alice"))
        AssertEqual(handle, app.systemLogMonitor.discovery.job.process)
        AssertEqual(started, app.systemLogMonitor.discovery.job.started)
        discovery := app.systemLogMonitor.discovery
        app.systemLogMonitor.Poll(Map())
        AssertFalse(app.systemLogMonitor.HasOwnProp("discovery"))
        AssertFalse(discovery.job.busy, "Logging out all characters must stop pending archive work.")
    } finally {
        app.Cleanup()
        FileDelete(path)
    }
}
