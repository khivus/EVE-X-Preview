TestRunner.Register("Simultaneous slow log workers leave GUI timers and settings controls responsive", LogWorkerResponsiveTest)

TestRunner.Register("Statistics totals survive restart and deduplicate saved characters", StatisticsPersistenceTest)
StatisticsPersistenceTest() {
    directory := A_Temp "\EVE-X-Preview-StatisticsTest-" A_TickCount
    config := JSON.Load(default_JSON)
    profile := config["_Profiles"]["Default"]
    profile["Thumbnail Positions"]["EVE - Alice"] := Map()
    profile["Client Possitions"]["EVE - Bob"] := Map()
    profile["Hotkeys Settings"]["CharacterHotkeys"].Push(Map("alice", "F1"))
    profile["Thumbnail Positions"]["EVE"] := Map()
    config["_Profiles"]["Second"] := SettingsModel.Clone(profile)
    owner := {_JSON: config, usageStatistics: 0}
    try {
        AssertFalse(config["CollectStatistics"])
        AssertEqual(2, UsageStatistics.SavedCharacters(config))
        stats := owner.usageStatistics := UsageStatistics(owner, directory)
        stats.switchesRequested := 5
        stats.characterLaunches := 2
        stats.ConfirmSwitch(100, 100)
        stats.ConfirmSwitch(100, 200)
        stats.ConfirmSwitch(100, 800)
        stats.Flush()
        stats.switchesRequested += 1
        stats.Close()
        owner.usageStatistics := 0
        saved := UsageStatistics.Snapshot(owner, directory)
        AssertFalse(saved["enabled"])
        AssertEqual(6, saved["categories"][2]["metrics"][1]["value"])
        stats := owner.usageStatistics := UsageStatistics(owner, directory)
        stats.ConfirmSwitch(100, 100)
        lifetime := UsageStatistics.Snapshot(owner, directory)
        AssertEqual(4, lifetime["categories"][2]["metrics"][2]["value"])
        AssertEqual(2, lifetime["categories"][2]["metrics"][3]["value"])
        AssertEqual(2, lifetime["categories"][3]["metrics"][1]["value"])
    } finally {
        if owner.usageStatistics
            owner.usageStatistics.Close()
        if DirExist(directory)
            DirDelete(directory, true)
    }
}

TestRunner.Register("Statistics archive scans resume and count only complete new log records", StatisticsArchiveResumeTest)
StatisticsArchiveResumeTest() {
    directory := A_Temp "\EVE-X-Preview-StatisticsArchiveTest-" A_TickCount
    game := directory "\game", chat := directory "\chat", storage := directory "\statistics"
    DirCreate(game), DirCreate(chat)
    job := LogScanJob()
    try {
        timestamp := "[ 2026.11.19 00:00:00 ]"
        combat := game "\20261119_000000_100.txt"
        localLog := chat "\Local_20261119_000000_100.txt"
        FileAppend(timestamp " (combat) <b>25</b> from pilot`n" timestamp " (combat) 40 to pilot", combat, "UTF-8-RAW")
        FileAppend(timestamp " EVE System > Channel changed to Local : Jita`n" timestamp " Pilot > Привет`n", localLog, "UTF-16")
        request := "statistics`n" game "`n" chat "`n" storage
        RunStatisticsFixtureScan(job, request)
        RunStatisticsFixtureScan(job, request)
        scanLog := FileRead(storage "\scan.log", "UTF-8")
        AssertTrue(InStr(scanLog, "Loading saved scan checkpoints"), "Log the checkpoint-loading phase.")
        AssertTrue(InStr(scanLog, "Finding EVE log files"), "Log the file-discovery phase.")
        AssertTrue(InStr(scanLog, "Ended;"), "Log the scan outcome.")
        totals := UsageStatistics.ReadArchive(storage)
        AssertEqual(1, totals[3])
        AssertEqual(25, totals[4])
        AssertEqual(0, totals[5])
        AssertEqual(1, totals[8])
        AssertEqual(1, totals[9])
        FileAppend("`n" timestamp " (notify) Undocking from station`n", combat, "UTF-8-RAW")
        FileAppend(timestamp " Pilot > World`n", localLog, "UTF-16")
        RunStatisticsFixtureScan(job, request)
        FileAppend("torn record", storage "\archive-index.tsv", "UTF-8-RAW")
        RunStatisticsFixtureScan(job, request)
        totals := UsageStatistics.ReadArchive(storage)
        AssertEqual(1, totals[1])
        AssertEqual(1, totals[2])
        AssertEqual(2, totals[3])
        AssertEqual(25, totals[4])
        AssertEqual(40, totals[5])
        AssertEqual(2, totals[9])
        AssertEqual(1, totals[10])
    } finally {
        job.Close()
        DirDelete(directory, true)
    }
}

RunStatisticsFixtureScan(job, request) {
    AssertTrue(job.Start(request))
    deadline := A_TickCount + 10000
    while job.busy && A_TickCount < deadline {
        result := job.Poll()
        Sleep(10)
    }
    AssertFalse(job.busy)
    AssertTrue(result && result["ok"])
    return result["text"]
}

TestRunner.Register("Statistics filters aggregate dated mining and peak damage without duplicating appended records", StatisticsFilterTest)
TestRunner.Register("Statistics live indexing leaves older records for manual scans available with collection off", StatisticsLiveBackfillTest)
StatisticsLiveBackfillTest() {
    directory := A_Temp "\EVE-X-Preview-StatisticsLive-" A_TickCount
    game := directory "\game", chat := directory "\chat", storage := directory "\statistics"
    DirCreate(game), DirCreate(chat)
    job := LogScanJob()
    owner := {_JSON: JSON.Load(default_JSON), usageStatistics: 0}
    try {
        today := FormatTime(A_NowUTC, "yyyy.MM.dd")
        old := FormatTime(DateAdd(A_NowUTC, -3, "Days"), "yyyy.MM.dd")
        path := game "\" StrReplace(old, ".") "_120000_1.txt"
        FileAppend("Listener: Alice`nSession Started: " old " 12:00:00`n"
            . "[ " old " 12:00:01 ] (combat) 100 from pilot`n"
            . "[ " today " 12:00:01 ] (combat) 20 from pilot`n", path, "UTF-8-RAW")
        request := "statistics-live`n" game "`n" chat "`n" storage "`n"
            . DateAdd(A_Now, -1, "Days") "`n" SubStr(A_NowUTC, 1, 8) "000000"
        RunStatisticsFixtureScan(job, request)
        RunStatisticsFixtureScan(job, request)
        AssertEqual(20, UsageStatistics.ReadArchive(storage)[4])
        info := UsageStatistics.Snapshot(owner, storage, Map())
        AssertFalse(info["enabled"])
        AssertTrue(info["available"], "Saved statistics must open without starting live collection.")
        stats := UsageStatistics.ForOwner(owner, storage)
        resolver := {ResolveGameLogsDirectory: (*) => game, ResolveChatLogsDirectory: (*) => chat}
        stats.StartScan(false, resolver)
        AssertTrue(stats.job && stats.job.busy, "Manual scans can start with live collection disabled.")
        deadline := A_TickCount + 10000
        while stats.job.busy && A_TickCount < deadline {
            Sleep(20)
            stats.Poll()
        }
        AssertFalse(stats.job.busy)
        AssertEqual(120, UsageStatistics.ReadArchive(storage)[4], "Backfill includes older hits without repeating live hits.")
    } finally {
        job.Close()
        if owner.HasOwnProp("statisticsBrowser")
            owner.statisticsBrowser.Close()
        DirDelete(directory, true)
    }
}

StatisticsFilterTest() {
    directory := A_Temp "\EVE-X-Preview-StatisticsFilters-" A_TickCount
    game := directory "\game", chat := directory "\chat", storage := directory "\statistics"
    DirCreate(game), DirCreate(chat)
    job := LogScanJob()
    try {
        today := FormatTime(A_NowUTC, "yyyy.MM.dd")
        older := FormatTime(DateAdd(A_NowUTC, -45, "Days"), "yyyy.MM.dd")
        filenameDate := StrReplace(today, ".")
        stamp := "[ " today " 12:00:01 ] "
        alice := game "\" filenameDate "_120000_1.txt"
        bob := game "\" filenameDate "_120000_2.txt"
        previous := game "\" StrReplace(older, ".") "_120000_1.txt"
        FileAppend("Listener: Alice`nSession Started: " today " 12:00:00`n"
            . stamp "(combat) 10 from hostile`n" stamp "(combat) 20 to hostile`n"
            . stamp "(mining) You mined 2,000 units of Fullerite-C32 with a lost residue of 30 units`n", alice, "UTF-8-RAW")
        FileAppend("Listener: Bob`nSession Started: " today " 12:00:00`n" stamp "(combat) 900 to hostile`n", bob, "UTF-8-RAW")
        FileAppend("Listener: Alice`nSession Started: " older " 12:00:00`n[ " older " 12:00:01 ] (combat) 300 to hostile`n", previous, "UTF-8-RAW")
        request := "statistics`n" game "`n" chat "`n" storage
        progress := RunStatisticsFixtureScan(job, request)
        AssertTrue(InStr(progress, "`n3`t3`t"))
        FileAppend("[ " today " 12:00:08 ] (combat) 15 from hostile`n"
            . "[ " today " 12:00:18 ] (combat) 22 from hostile`n"
            . "[ " today " 12:00:18 ] (mining) You mined 42 units of Gneiss`n", alice, "UTF-8-RAW")
        RunStatisticsFixtureScan(job, request)
        todayResult := RunStatisticsFixtureScan(job, "statistics-view`n" storage "`nAlice`nmonth")
        AssertTrue(InStr(todayResult, "P`tBob"))
        allView := RunStatisticsFixtureScan(job, "statistics-view`n" storage "`n`nall")
        AssertTrue(InStr(allView, "T`tcount1`tAlice"), "Alice has more game sessions indexed.")
        AssertTrue(InStr(todayResult, "M`tFullerite-C32`t2000`t30"))
        AssertTrue(InStr(todayResult, "M`tGneiss`t42`t0"))
        monthly := StrSplit(StrSplit(todayResult, "`n")[1], "`t")
        AssertEqual(47, Number(monthly[5]))
        AssertEqual(20, Number(monthly[6]))
        AssertEqual(2.5, Number(monthly[15]), "Hits within ten seconds combine across a saved checkpoint.")
        AssertEqual(22, Number(monthly[18]), "Alpha uses the largest hit even when its DPS window is weaker.")
        AssertEqual(20, Number(monthly[19]))
        lifetime := StrSplit(StrSplit(RunStatisticsFixtureScan(job, "statistics-view`n" storage "`nAlice`nall"), "`n")[1], "`t")
        AssertEqual(320, Number(lifetime[6]))
        AssertEqual(30, Number(lifetime[16]))
        AssertEqual(300, Number(lifetime[19]))
        latest := StrSplit(StrSplit(RunStatisticsFixtureScan(job, "statistics-view`n" storage "`nAlice`nsession"), "`n")[1], "`t")
        AssertEqual(20, Number(latest[6]))

        stats := UsageStatistics({_JSON: JSON.Load(default_JSON)}, storage)
        ; Simulate the archive writer running while Settings switches filters.
        stats.job := {busy: true}
        try {
            AssertTrue(stats.FilteredView("Alice", "month").Get("pending", false))
            view := stats.FilteredView("Bob", "all")
            deadline := A_TickCount + 10000
            while view.Get("pending", false) && A_TickCount < deadline {
                Sleep(15)
                view := stats.FilteredView("Bob", "all")
            }
            AssertFalse(view.Get("deferred", false), "A completed snapshot can be read during a scan.")
            AssertEqual(900, view["counts"][5], "The latest character selection wins while an earlier query is pending.")
            AssertEqual(20, stats.FilteredView("Alice", "month")["counts"][5], "The earlier filter result can be reused.")
        } finally {
            stats.job := 0
            stats.Close()
        }
    } finally {
        job.Close()
        DirDelete(directory, true)
    }
}

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

TestRunner.Register("Statistics extended events survive resume and attribute independent peaks", StatisticsExtendedTest)
StatisticsExtendedTest() {
    directory := A_Temp "\EVE-X-Preview-Extended-" A_TickCount
    game := directory "\game", chat := directory "\chat", storage := directory "\statistics"
    DirCreate(game), DirCreate(chat)
    job := LogScanJob(), stats := UsageStatistics({_JSON: JSON.Load(default_JSON)}, storage)
    try {
        stamp := "[ 2026.11.20 12:00:01 ] "
        alice := game "\20261120_120000_1.txt"
        FileAppend("Listener: Alice`nSession Started: 2026.11.20 12:00:00`n"
            . stamp "(combat) 100 to Target - Laser II - Hits`n"
            . stamp "(combat) 100 to Target - Laser II - Hits`n"
            . stamp "(combat) 1 to Target - Ogre II - Hits`n"
            . stamp "(combat) 1 to Target - Hammerhead II - Hits`n"
            . stamp "(combat) 20 remote armor repaired to Friend - Repairer II`n"
            . stamp "(combat) 30 remote shield boosted by Friend - Booster II`n"
            . stamp "(combat) 40 remote hull repaired to Friend - Hull Repairer II`n"
            . stamp "(combat) 50 remote structure repaired by Friend - Hull Repairer II`n"
            . stamp "(notify) Your <b>Mining Foreman Burst II</b> has applied bonuses to <b>6</b> fleet members.`n"
            . stamp "(notify) Your Armor Command Burst II has applied bonuses to 1 fleet member.`n"
            . stamp "(notify) Following Friend in warp`n"
            . stamp "(notify) Regrouping to Friend`n"
            . stamp "(notify) Jumping from Jita to Perimeter`n"
            . stamp "(notify) Undocking from station`n"
            . stamp "(notify) A Conduit Field activated by Friend`n"
            . stamp "(notify) Please wait...`n", alice, "UTF-8-RAW")
        FileAppend("Listener: Bob`nSession Started: 2026.11.20 12:00:00`n"
            . stamp "(combat) 150 to Target - Laser II - Hits`n"
            . stamp "(notify) Your Shield Command Burst II has applied bonuses to 20 fleet members.`n", game "\20261120_120000_2.txt", "UTF-8-RAW")
        FileAppend("Listener: Alice`nSession Started: 2026.11.20 12:00:00`n"
            . stamp "EVE System > Channel changed to Local : Jita`n"
            . stamp "EVE System > Channel changed to Local : Amarr`n"
            . stamp "EVE System > Channel changed to Local : Jita`n", chat "\Local_20261120_120000_1.txt", "UTF-16")
        request := "statistics`n" game "`n" chat "`n" storage
        RunStatisticsFixtureScan(job, request)
        RunStatisticsFixtureScan(job, request)
        view := stats.ParseView(RunStatisticsFixtureScan(job, "statistics-view`n" storage "`n`nall"))
        AssertEqual(20, view["extra"]["armorGiven"])
        AssertEqual(30, view["extra"]["shieldReceived"])
        AssertEqual(40, view["extra"]["hullGiven"])
        AssertEqual(50, view["extra"]["hullReceived"])
        AssertEqual(4, view["extra"]["logisticsRecords"])
        AssertEqual(3, view["extra"]["burstApplications"])
        AssertEqual(27, view["extra"]["burstRecipients"])
        AssertEqual(20, view["extra"]["burstMaxRecipients"])
        AssertEqual("Alice", view["leaders"]["burstApplications"])
        AssertEqual("Bob", view["leaders"]["burstRecipients"])
        AssertEqual("Bob", view["leaders"]["burstMaxRecipients"])
        AssertEqual(3, view["bursts"].Length)
        AssertEqual(1, view["extra"]["fleetWarps"])
        AssertEqual(1, view["extra"]["regroups"])
        AssertEqual(1, view["extra"]["pleaseWait"])
        AssertEqual("Alice", view["leaders"]["peak2"])
        AssertEqual("Bob", view["leaders"]["alpha2"])
        AssertEqual("Alice", view["leaders"]["count3"])
        AssertEqual("Alice", view["leaders"]["count5"])
        AssertEqual("Alice", view["leaders"]["fleetWarps"])
        AssertEqual("Alice", view["leaders"]["count8"])
        AssertEqual("Alice", view["leaders"]["count7"])
        AssertEqual("Alice", view["leaders"]["count10"])
        AssertEqual("Alice", view["leaders"]["count13"])
        AssertEqual(2, view["drones"].Length)
        for drone in view["drones"]
            AssertEqual("Alice", drone["character"])
        for module in view["modules"]
            AssertFalse(module["name"] = "Ogre II" || module["name"] = "Hammerhead II")
        for module in view["modules"]
            if module["name"] = "Laser II" {
                AssertEqual(3, module["count"])
                AssertEqual("Alice", module["character"])
            }
        AssertEqual(2, view["systems"].Length)
        FileAppend(stamp "(notify) Please wait...`n"
            . stamp "(notify) Your Mining Foreman Burst II has applied bonuses to 2 fleet members.`n", alice, "UTF-8-RAW")
        RunStatisticsFixtureScan(job, request)
        view := stats.ParseView(RunStatisticsFixtureScan(job, "statistics-view`n" storage "`nAlice`nall"))
        AssertEqual(2, view["extra"]["pleaseWait"])
        AssertEqual(20, view["extra"]["armorGiven"])
        AssertEqual(3, view["extra"]["burstApplications"])
        AssertEqual(9, view["extra"]["burstRecipients"])
        AssertEqual(6, view["extra"]["burstMaxRecipients"])
        ; Downgrade the checkpoint marker to verify backfill replaces aggregates.
        indexPath := storage "\archive-index.tsv"
        old := StrReplace(FileRead(indexPath, "UTF-8"), ";V4", ";V3")
        old := RegExReplace(old, ";@extra,hull(?:Given|Received),[^;|`t`r`n]*")
        FileDelete(indexPath)
        FileAppend(old, indexPath, "UTF-8-RAW")
        RunStatisticsFixtureScan(job, request)
        view := stats.ParseView(RunStatisticsFixtureScan(job, "statistics-view`n" storage "`n`nall"))
        AssertEqual(2, view["extra"]["pleaseWait"], "Backfill replaces old counts instead of adding duplicates.")
        AssertFalse(view["extendedNeedsRescan"])
        AssertEqual(4, view["extra"]["burstApplications"])
        AssertEqual(29, view["extra"]["burstRecipients"])
        AssertEqual(20, view["extra"]["burstMaxRecipients"])
        for burst in view["bursts"]
            if burst["name"] = "Mining Foreman Burst II" {
                AssertEqual(2, burst["count"])
                AssertEqual("Alice", burst["character"])
            }
        AssertEqual(40, view["extra"]["hullGiven"], "Backfill upgrades pre-hull checkpoints.")
        AssertEqual(50, view["extra"]["hullReceived"])
        for system in view["systems"]
            if system["name"] = "Jita"
                AssertEqual(2, system["count"])
    } finally {
        job.Close(), stats.Close()
        DirDelete(directory, true)
    }
}

TestRunner.Register("Statistics character counters persist once and merge new activity", StatisticsCharacterCountersTest)
StatisticsCharacterCountersTest() {
    directory := A_Temp "\EVE-X-Preview-CharacterCounters-" A_TickCount
    owner := {_JSON: JSON.Load(default_JSON)}
    stats := UsageStatistics(owner, directory)
    try {
        stats.switchesRequested += 3
        stats.Attribute("switchesRequested", "EVE - Alice", 2)
        stats.Attribute("switchesRequested", "EVE - Bob", 1)
        stats.ConfirmSwitch(100, 100, "Alice")
        stats.Flush()
        stats.Flush()
        AssertEqual(2, UsageStatistics.ReadTotals(stats.path)["characters"]["Alice"]["switchesRequested"])
        AssertEqual("Alice", stats.CharacterLeaders()["switchesRequested"])
        stats.Attribute("switchesRequested", "Bob", 4)
        AssertEqual("Bob", stats.CharacterLeaders()["switchesRequested"])
        stats.Flush()
        AssertEqual(5, UsageStatistics.ReadTotals(stats.path)["characters"]["Bob"]["switchesRequested"])
        stats.Close()
        stats := UsageStatistics(owner, directory)
        AssertEqual("Bob", stats.CharacterLeaders()["switchesRequested"])
        AssertEqual("Alice", stats.CharacterLeaders()["switchesConfirmed"])
    } finally {
        stats.Close()
        DirDelete(directory, true)
    }
}
