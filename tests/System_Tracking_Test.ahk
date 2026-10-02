class SystemTrackingFixture extends CombatEventsFixture {
    __New() {
        super.__New()
        this.ProfileGameLogsMonitoring := "Default"
        this.systemTrackingEnabled := 1
        this.directory := A_ScriptDir "\SystemLogsFixture-" A_TickCount
        DirCreate(this.directory)
        this.chatLogsDirectory := this.directory
        this.files := []
        this.active := Map("EVE - Alice", "alice", "EVE - Bob", "bob")
        for character, hwnd in this.active
            this.ThumbWindows.%hwnd% := Map("Window", {Title: character}, "TextOverlay", Map("OverlayText", {Text: this.CleanTitle(character)}, "EventText", {Text: ""}, "SystemText", {Text: ""}))
        this.systemLogMonitor := LocalChatMonitor(this)
    }

    WriteLog(name, listener, session, text, encoding := "UTF-8") {
        path := this.directory "\" name
        this.files.Push(path)
        header := "`r`n`r`n`r`n`r`n    ---------------------------------------------------------------`r`n`r`n      Channel ID:      local`r`n      Channel Name:    Local`r`n      Listener:        " listener "`r`n      Session started: " session "`r`n    ---------------------------------------------------------------`r`n`r`n"
        FileAppend(header text, path, encoding)
        return path
    }

    Message(system) => "[ 2026.09.08 23:28:23 ] EVE System > Channel changed to Local : " system "`r`n"
    Append(path, text) {
        file := FileOpen(path, "a", "UTF-8") ; Share read access with the live monitor, like the game writer.
        file.Write(text)
        file.Close()
    }
    System(hwnd) => this.ThumbWindows.%hwnd%["TextOverlay"]["SystemText"].Text
    monitorCharacterSystems() => this.systemLogMonitor.Poll(this.active)

    Cleanup() {
        this.stopSystemTracking()
        for path in this.files {
            if FileExist(path)
                FileDelete(path)
        }
        DirDelete(this.directory) ; Empty fixture directory only, never recursive.
    }
}

TestRunner.Register("Local logs match listeners and track only system announcements", LocalChatReadTest)
LocalChatReadTest() {
    app := SystemTrackingFixture()
    try {
        alice := app.WriteLog("Local_20260908_123456_111.txt", "Alice", "2026.09.08 23:28:20", Chr(0xFEFF) app.Message("Dodixie") app.Message("Jita"), "UTF-16")
        app.WriteLog("Local_20260908_123456_222.txt", "Bob", "2026.09.08 23:28:20", app.Message("Amarr"))
        app.WriteLog("Local_20260908_123456_333.txt", "Alice Two", "2026.09.08 23:29:20", app.Message("Wrong"))
        app.startSystemTracking() ; Combat monitoring remains disabled in defaults.
        AssertEqual("Jita", app.System("alice"))
        AssertEqual("Amarr", app.System("bob"))
        app.Append(alice, "[ 2026.09.08 23:29:23 ] Someone > EVE System > Channel changed to Local : Fake`r`n")
        app.monitorCharacterSystems()
        AssertEqual("Jita", app.System("alice"))
        app.Append(alice, app.Message("Rens"))
        app.monitorCharacterSystems()
        AssertEqual("Rens", app.System("alice"))
        app.updateThumbnailEventText("Warp Disrupted", "alice")
        app.updateThumbnailEventText("", "alice")
        AssertEqual("Rens", app.System("alice"), "Ending an event must preserve the current system.")
        AssertEqual("Alice", app.ThumbWindows.alice["TextOverlay"]["OverlayText"].Text)
        app.active.Delete("EVE - Alice")
        app.monitorCharacterSystems()
        AssertFalse(app.systemLogMonitor.readers.Has("EVE - Alice"))
        AssertEqual("", app.System("alice"), "Logout clears the previous system.")
    } finally {
        app.Cleanup()
    }
}

TestRunner.Register("Local logs retry discovery, rotate sessions and buffer partial writes", LocalChatLifecycleTest)
LocalChatLifecycleTest() {
    app := SystemTrackingFixture()
    try {
        app.monitorCharacterSystems() ; No files yet.
        old := app.WriteLog("Local_20260908_999999_111.txt", "Alice", "2026.09.08 20:00:00", app.Message("Dodixie"))
        app.systemLogMonitor.lastDiscovery := -5000
        app.monitorCharacterSystems()
        AssertEqual("Dodixie", app.System("alice"))
        newer := app.WriteLog("Local_20260908_000001_111.txt", "Alice", "2026.09.08 23:00:00", app.Message("Jita"))
        app.systemLogMonitor.lastDiscovery := -5000
        app.monitorCharacterSystems()
        AssertEqual("Jita", app.System("alice"), "Session headers determine which log is newest.")
        app.Append(old, app.Message("Wrong"))
        app.systemLogMonitor.lastDiscovery := -5000
        app.monitorCharacterSystems()
        AssertEqual("Jita", app.System("alice"), "Old log activity cannot replace the current session.")
        app.Append(newer, "[ 2026.09.08 23:28:23 ] EVE System > Channel changed to Local : Oursu")
        app.monitorCharacterSystems()
        AssertEqual("Jita", app.System("alice"))
        app.Append(newer, "laert`r`n")
        app.monitorCharacterSystems()
        AssertEqual("Oursulaert", app.System("alice"))
        file := FileOpen(newer, "w", "UTF-8")
        file.Write(app.Message("Rens"))
        file.Close()
        app.monitorCharacterSystems()
        AssertEqual("Rens", app.System("alice"), "Truncated logs restart at the beginning.")
    } finally {
        app.Cleanup()
    }
}

TestRunner.Register("System and event lines stack inside thumbnails", SystemTrackingLayoutTest)
SystemTrackingLayoutTest() {
    app := ThumbnailNamesFixture()
    app.systemTrackingEnabled := 1
    overlay := Gui("-Caption")
    try {
        overlay.MarginX := 5
        overlay.MarginY := 5
        overlay.SetFont("s12", "Gill Sans MT")
        app.AddThumbnailTextControls(overlay, "EVE - Alice", 200, 150)
        for size in [[200, 150], [400, 300], [50, 50], [8, 8]] {
            app.LayoutThumbnailText(overlay, 0, size[1], size[2])
            overlay["OverlayText"].GetPos(&nx, &ny, &nw, &nh)
            overlay["EventText"].GetPos(&ex, &ey, &ew, &eh)
            overlay["SystemText"].GetPos(&sx, &sy, &sw, &sh)
            AssertTrue(ny + nh <= ey && ey + eh <= sy, "The name, event, and system must not overlap.")
            AssertEqual(size[2] - Min(5, size[2] / 2), sy + sh)
            AssertTrue(sx >= 0 && sy >= 0 && sx + sw <= size[1] && sy + sh <= size[2])
        }
    } finally {
        overlay.Destroy()
    }
}

TestRunner.Register("System tracking controls fit in Game Logs Monitoring", SystemTrackingSettingsTest)
SystemTrackingSettingsTest() {
    app := CombatEventsFixture()
    app.ProfileGameLogsMonitoring := "Default"
    AssertEqual(0, app.systemTrackingEnabled)
    app.systemTrackingEnabled := 1
    app.chatLogsDirectory := "D:\Custom Chatlogs"
    app._JSON := JsonMergeNoOverwrite(JSON.Load(default_JSON), JSON.Load(JSON.Dump(app._JSON)))
    AssertEqual(1, app.systemTrackingEnabled)
    AssertEqual("D:\Custom Chatlogs", app.ResolveChatLogsDirectory())
    app.SetState()
    app.MainFrame := Gui()
    app.MainFrame.Group := Map()
    try {
        app.GameLogsMonitoring_Ctrl()
        for control in app.MainFrame.Group["Game Logs Monitoring"] {
            control.GetPos(&x, &y, &w, &h)
            AssertTrue(x + w <= app.contentW - 10 && y + h <= app.guiHeight - 10, "Game Logs Monitoring controls must fit the panel: " control.Text)
        }
    } finally {
        app.MainFrame.Destroy()
    }
}

TestRunner.Register("Game and chat directories share automatic detection without saving paths", LogDirectoryResolutionTest)
LogDirectoryResolutionTest() {
    app := CombatEventsFixture()
    app.ProfileGameLogsMonitoring := "Default"
    root := A_ScriptDir "\LogDirectories-" A_TickCount
    chat := root "\Chatlogs"
    game := root "\Gamelogs"
    alternate := root "\Alternate"
    DirCreate(chat)
    DirCreate(game)
    DirCreate(alternate)
    try {
        app.chatLogsDirectory := chat
        app.gameLogsDirectory := ""
        AssertEqual(game, app.ResolveGameLogsDirectory())
        AssertEqual("", app.gameLogsDirectory)
        app.gameLogsDirectory := game "\"
        app.chatLogsDirectory := ""
        AssertEqual(chat, app.ResolveChatLogsDirectory())
        AssertEqual("", app.chatLogsDirectory)
        app.gameLogsDirectory := ""
        app.gameLogsMonitoring() ; Disabled monitoring must not populate an automatic path.
        AssertEqual("", app.gameLogsDirectory)
        FileAppend("sample", game "\one.txt")
        FileAppend("sample", alternate "\two.txt")
        app.chatLogsDirectory := chat
        AssertArrayEqual([game "\one.txt"], app.getFilesList())
        app.gameLogsDirectory := alternate
        AssertArrayEqual([alternate "\two.txt"], app.getFilesList(), "A same-size listing from another folder must not reuse cached paths.")
        app.gameLogsDirectory := root "\Missing"
        app.gameLogsMonitoringEnabled := 1
        app.waitingMonitoringChars["Pilot"] := (*) => 0
        app.startLogMonitoring("Pilot", 0)
        AssertFalse(app.waitingMonitoringChars.Has("Pilot"), "Missing folders must allow discovery to retry.")
        AssertEqual(root "\Missing", app.gameLogsDirectory, "Failed discovery must not overwrite the user's setting.")
    } finally {
        for path in [game "\one.txt", alternate "\two.txt"] {
            if FileExist(path)
                FileDelete(path)
        }
        DirDelete(chat)
        DirDelete(game)
        DirDelete(alternate)
        DirDelete(root)
    }
}
