TestRunner.Register("Quiet game logs remain attached without restart churn", GameLogQuietReaderTest)
GameLogQuietReaderTest() {
    previousHidden := A_DetectHiddenWindows
    DetectHiddenWindows(true)
    app := DPSLogFixture()
    try {
        path := app.WriteLog("20260912_000000_12345.txt", "")
        app.startLogMonitoring(app.character, 12345)
        reader := app.monitoredChars[app.character]
        reader["launchTime"] := A_TickCount - 60000
        reader["fileUpdated"] := false
        before := app.debugToolTipText
        Loop 4 {
            app.getFilesList(true)
            app.monitorAllChars()
        }
        AssertEqual(reader, app.monitoredChars[app.character], "Quiet polls must preserve the reader, not reattach at EOF.")
        AssertEqual(before, app.debugToolTipText, "No stop/start messages for an unchanged session.")
        now := DPSMeter.CurrentSecond()
        app.Append(path, DPSLine(now, 300))
        app.monitorChanges(app.character, now)
        AssertEqual(30, reader["dps"].Rates(now).incoming)
    } finally {
        app.Cleanup()
        DetectHiddenWindows(previousHidden)
    }
}

TestRunner.Register("Game log discovery chooses the newest session rather than last write", GameLogSessionOrderTest)
GameLogSessionOrderTest() {
    app := DPSLogFixture()
    try {
        older := app.WriteLog("20260912_000000_12345.txt", "")
        newer := app.WriteLog("20260912_000100_12345.txt", "")
        FileSetTime("20261008000000", older, "M")
        FileSetTime("20260912000100", newer, "M")
        app.startLogMonitoring(app.character, 12345)
        AssertEqual(newer, app.monitoredChars[app.character]["fileName"])
        AssertEqual(1, app.getFilesList().Length, "Historical sessions of the same ID must not require header reads.")
    } finally app.Cleanup()
}

TestRunner.Register("Game log filename notifications follow rotation after updates", GameLogRotationTest)
GameLogRotationTest() {
    previousHidden := A_DetectHiddenWindows
    DetectHiddenWindows(true)
    app := DPSLogFixture()
    now := DPSMeter.CurrentSecond()
    try {
        old := app.WriteLog("20260912_000000_12345.txt", "")
        app.startLogMonitoring(app.character, 12345)
        app.Append(old, DPSLine(now, 100))
        app.monitorChanges(app.character, now)
        newer := app.WriteLog("20260912_000100_12345.txt", DPSLine(now, 500))
        Loop 20 {
            app.monitorAllChars()
            if app.monitoredChars[app.character]["fileName"] = newer
                break
            Sleep(10)
        }
        reader := app.monitoredChars[app.character]
        AssertEqual(newer, reader["fileName"], "Rotation must not depend on the old reader being quiet or never updated.")
        AssertEqual(50, reader["dps"].Rates(now).incoming, "The new session's first hit must not be skipped at attachment.")
        generation := app.gameLogDiscovery.generation
        app.Append(newer, DPSLine(now, 200))
        app.monitorAllChars()
        AssertEqual(generation, app.gameLogDiscovery.generation, "Combat writes must not rescan the archive.")
        AssertEqual(70, reader["dps"].Rates(now).incoming)
    } finally {
        app.Cleanup()
        DetectHiddenWindows(previousHidden)
    }
}

TestRunner.Register("Incomplete and mismatched replacement headers preserve the active reader", GameLogIncompleteRotationTest)
GameLogIncompleteRotationTest() {
    previousHidden := A_DetectHiddenWindows
    DetectHiddenWindows(true)
    app := DPSLogFixture()
    now := DPSMeter.CurrentSecond()
    try {
        old := app.WriteLog("20260912_000000_12345.txt", "")
        app.startLogMonitoring(app.character, 12345)
        reader := app.monitoredChars[app.character]
        path := app.directory "\20260912_000100_12345.txt"
        app.files.Push(path)
        FileAppend("----------------`r`nGamelog`r`nListener: DPS Test Pilot`r`nSession Started: ", path, "UTF-8")
        app.getFilesList(true)
        app.monitorAllChars()
        AssertEqual(reader, app.monitoredChars[app.character])
        app.Append(old, DPSLine(now, 300))
        app.monitorChanges(app.character, now)
        AssertEqual(30, reader["dps"].Rates(now).incoming)
        app.Append(path, "2026.09.12 00:01:00`r`n" DPSLine(now, 600))
        app.monitorAllChars()
        AssertEqual(path, app.monitoredChars[app.character]["fileName"])
        AssertEqual(60, app.monitoredChars[app.character]["dps"].Rates(now).incoming)
        wrong := app.WriteLog("20260912_000200_12345.txt", "")
        content := StrReplace(FileRead(wrong, "UTF-8"), "Listener: DPS Test Pilot", "Listener: Other Pilot")
        FileDelete(wrong)
        FileAppend(content, wrong, "UTF-8")
        app.getFilesList(true)
        app.monitorAllChars()
        AssertEqual(path, app.monitoredChars[app.character]["fileName"], "A cached ID must still validate the selected log's listener.")
    } finally {
        app.Cleanup()
        DetectHiddenWindows(previousHidden)
    }
}

TestRunner.Register("Replacing a game log at the same path follows the new file identity", GameLogSamePathReplacementTest)
GameLogSamePathReplacementTest() {
    previousHidden := A_DetectHiddenWindows
    DetectHiddenWindows(true)
    app := DPSLogFixture()
    now := DPSMeter.CurrentSecond()
    try {
        path := app.WriteLog("20260912_000000_12345.txt", "")
        app.startLogMonitoring(app.character, 12345)
        old := app.monitoredChars[app.character]
        FileDelete(path)
        FileAppend("----------------`r`nGamelog`r`nListener: DPS Test Pilot`r`nSession started: 2026.09.12 00:00:00`r`n" DPSLine(now, 700), path, "UTF-8")
        app.getFilesList(true)
        app.monitorAllChars()
        reader := app.monitoredChars[app.character]
        AssertNotEqual(old["identity"], reader["identity"])
        AssertEqual(70, reader["dps"].Rates(now).incoming)
    } finally {
        app.Cleanup()
        DetectHiddenWindows(previousHidden)
    }
}

TestRunner.Register("Game log discovery releases its watcher and ignores malformed filenames", GameLogWatcherCleanupTest)
GameLogWatcherCleanupTest() {
    app := DPSLogFixture()
    try {
        app.WriteLog("not_a_session.txt", "")
        app.WriteLog("20260912_000000.txt", "")
        AssertEqual(0, app.getFilesList().Length)
        discovery := app.gameLogDiscovery
        AssertTrue(discovery.watch != 0)
        handle := discovery.watch
        app.CloseGameLogDiscovery()
        AssertEqual(0, discovery.watch)
        AssertEqual(0xFFFFFFFF, DllCall("WaitForSingleObject", "Ptr", handle, "UInt", 0, "UInt"))
    } finally app.Cleanup()
}

WaitGameLogListing(app) {
    files := app.getFilesList()
    deadline := A_TickCount + 10000
    while app.gameLogDiscovery.busy {
        if A_TickCount > deadline
            throw Error("Game discovery worker did not finish")
        Sleep(10)
        files := app.getFilesList()
    }
    return files
}
