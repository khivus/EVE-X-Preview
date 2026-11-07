class Update1704ShortcutFixture extends DirectPriorityFixture {
    quickWindows := Map()
    targets := Map()
    trayRefreshes := 0
    utilities := 0
    QuickGroupTargetAvailable(hwnd) => this.quickWindows.Has(hwnd)
    ShortcutTargetAvailable(title) => this.targets.Has(title)
    SettingsEditingActive() => false
    TrayMenu() => this.trayRefreshes++
}

class Update1704DispatchFixture extends Update1704ShortcutFixture {
    shortcutTestActive := true
    calls := []
    SettingsEditingActive() => !this.shortcutTestActive
    ActivateEVEWindow(*) => this.calls.Push("direct")
    Cycle_Hotkey_Groups(*) => this.calls.Push("group")
    CycleQuickGroup(*) => this.calls.Push("quick")
}

TestRunner.Register("Update1704 native hotkey dispatch follows Quick Group direct and group priority as windows change", Update1704QuickDispatchTest)
Update1704QuickDispatchTest() {
    app := Update1704DispatchFixture(), direct := Gui(, "EVE - Priority fixture " A_TickCount)
    previousHidden := A_DetectHiddenWindows, previousLevel := A_SendLevel
    DetectHiddenWindows(true)
    try {
        app.EVEExe := "ahk_pid " DllCall("GetCurrentProcessId")
        app.Global_Hotkeys := true
        app._Hotkeys := [Map(direct.Title, "F16")]
        app.targets := Map(direct.Title, true, "EVE - Fleet fixture", true)
        app.Profiles["Default"]["Hotkey Groups"] := Map("Fleet", Map("Characters", ["EVE - Fleet fixture"], "ForwardsHotkey", "F16", "BackwardsHotkey", "", "FirstCharHotkey", ""))
        app.QuickGroupEnabled := true
        app.QuickGroupHotkey := "F16"
        app.QuickGroupChars := Map()
        app.debugToolTipMethod := (*) => 0
        app.debugToolTipDelay := -1000
        app.Register_Hotkey_Groups()
        app.RegisterQuickGroup()
        app.RegisterHotkeys(direct.Title)
        SendLevel(1)
        SendEvent("{F16}")
        Sleep(80)
        AssertEqual(["direct"], app.calls)
        app.QuickGroupChars[123] := "EVE - Quick fixture"
        app.quickWindows[123] := true
        SendEvent("{F16}")
        Sleep(80)
        AssertEqual(["direct", "quick"], app.calls)
        app.quickWindows.Clear()
        SendEvent("{F16}")
        Sleep(80)
        AssertEqual(["direct", "quick", "direct"], app.calls)
        app.targets.Delete(direct.Title)
        direct.Destroy()
        SendEvent("{F16}")
        Sleep(80)
        AssertEqual(["direct", "quick", "direct", "group"], app.calls)
        app.quickWindows[123] := true
        SendEvent("{F16}")
        Sleep(80)
        AssertEqual(["direct", "quick", "direct", "group", "quick"], app.calls)
        app.activationHotkeysSuspended := true
        SendEvent("{F16}")
        Sleep(80)
        AssertEqual(5, app.calls.Length)
    } finally {
        app.shortcutTestActive := false
        SetTimer(app.debugToolTipMethod, 0)
        HotIf()
        SendLevel(previousLevel)
        DetectHiddenWindows(previousHidden)
        try direct.Destroy()
    }
}

TestRunner.Register("Update1704 populated Quick Group wins shared shortcuts and empty or closed groups fall back", Update1704QuickPriorityTest)
Update1704QuickPriorityTest() {
    app := Update1704ShortcutFixture()
    app._Hotkeys := [Map("EVE - Direct", "CapsLock")]
    app.targets := Map("EVE - Direct", true, "EVE - Fleet", true)
    app.QuickGroupEnabled := true
    app.QuickGroupHotkey := "CapsLock"
    app.QuickGroupChars := Map()
    AssertFalse(app.QuickGroupShortcutAvailable("*CapsLock"))
    AssertTrue(app.DirectShortcutAvailable("*CapsLock"))
    app.targets.Delete("EVE - Direct")
    AssertTrue(app.OnWinExist(["EVE - Fleet"], "*CapsLock"))
    app.targets["EVE - Direct"] := true
    app.QuickGroupChars[123] := "EVE - Quick"
    AssertFalse(app.QuickGroupShortcutAvailable("*CapsLock"), "A closed/stale HWND must not reserve the key.")
    app.quickWindows[123] := true
    AssertTrue(app.QuickGroupShortcutAvailable("*CapsLock"))
    AssertFalse(app.DirectShortcutAvailable("*CapsLock"))
    AssertFalse(app.OnWinExist(["EVE - Fleet"], "*CapsLock"))
    AssertFalse(app.DirectShortcutAvailable("*F23"))
    app.quickWindows.Clear()
    AssertTrue(app.DirectShortcutAvailable("*CapsLock"))
    app.QuickGroupEnabled := false
    AssertFalse(app.QuickGroupShortcutAvailable("*CapsLock"))
    entries := []
    HotkeyBindings.Add(entries, "quick", "Quick Group", "CapsLock", true)
    HotkeyBindings.Add(entries, "character:1", "Direct", "CapsLock", true)
    HotkeyBindings.Add(entries, "group:Fleet:ForwardsHotkey", "Fleet", "CapsLock", true)
    AssertEqual(0, HotkeyBindings.Conflicts(entries).Count)
}

TestRunner.Register("Update1704 narrow pause keeps utility hotkeys and full pause keeps only its resume shortcut", Update1704PauseModesTest)
Update1704PauseModesTest() {
    app := Update1704ShortcutFixture()
    app.Suspend_Hotkeys_Hotkey := "F13"
    app.Suspend_All_Hotkeys_Hotkey := "F14"
    utilityCondition := (*) => !HotkeyCapture.Active
    previousLevel := A_SendLevel
    try {
        Suspend(false)
        app.RegisterPauseHotkeys()
        HotIf(utilityCondition)
        Hotkey("F15", (*) => app.utilities++, "S0 On")
        SendLevel(1)
        SendEvent("{F13}")
        Sleep(80)
        AssertTrue(app.activationHotkeysSuspended)
        AssertFalse(A_IsSuspended)
        AssertFalse(app.DirectShortcutAvailable("F23"))
        SendEvent("{F15}")
        Sleep(80)
        AssertEqual(1, app.utilities, "Narrow pause must leave utility shortcuts working.")
        SendEvent("{F14}")
        Sleep(80)
        AssertTrue(A_IsSuspended)
        SendEvent("{F13}{F15}")
        Sleep(80)
        AssertTrue(app.activationHotkeysSuspended, "The narrow-pause shortcut must be suspended with everything else.")
        AssertEqual(1, app.utilities)
        SendEvent("{F14}")
        Sleep(80)
        AssertFalse(A_IsSuspended, "The all-hotkeys shortcut must resume itself.")
        AssertTrue(app.activationHotkeysSuspended, "Resuming all hotkeys preserves the independent narrow pause.")
        SendEvent("{F13}{F15}")
        Sleep(80)
        AssertFalse(app.activationHotkeysSuspended)
        AssertEqual(2, app.utilities)
    } finally {
        Suspend(false)
        SendLevel(previousLevel)
        if app.HasOwnProp("pauseHotkeyCondition") {
            HotIf(app.pauseHotkeyCondition)
            try Hotkey("F13", "Off")
            try Hotkey("F14", "Off")
        }
        HotIf(utilityCondition)
        try Hotkey("F15", "Off")
        HotIf()
    }
}

TestRunner.Register("Update1704 pause bindings and tray controls are present in both settings editors", Update1704PauseSettingsTest)
Update1704PauseSettingsTest() {
    AssertEqual("Pause EVE switching hotkeys", SettingsSchema.Fields["Suspend_Hotkeys_Hotkey"].label)
    AssertEqual("Pause all hotkeys", SettingsSchema.Fields["Suspend_All_Hotkeys_Hotkey"].label)
    AssertTrue(SettingsSchema.Fields.Has("TrayCommand:Suspend All Hotkeys"))
    app := HotkeySettingsFixture()
    try {
        AssertTrue(IsObject(app.Find("Suspend_All_Hotkeys_Hotkey")))
        model := SettingsModel(app._JSON)
        model.Set("Suspend_All_Hotkeys_Hotkey", "F24")
        app._JSON := model.config
        AssertEqual("F24", app.Suspend_All_Hotkeys_Hotkey)
    } finally app.Dispose()
}

TestRunner.Register("Update1704 migration archives move into the backups folder and survive snapshot rotation", Update1704MigrationArchiveTest)
Update1704MigrationArchiveTest() {
    previousRoot := SettingsBackup.Root
    folder := A_ScriptDir "\..\.vs\update1704-tests\archives-" A_TickCount
    DirCreate(folder)
    SettingsBackup.Root := folder "\snapshots"
    try {
        old := folder "\EVE-X-Preview.PreV1.6.0.20260101.backup.json"
        FileAppend("original migration copy", old, "UTF-8-RAW")
        SettingsBackup.CollectMigrationBackups(folder)
        archive := SettingsBackup.Directory(folder) "\EVE-X-Preview.PreV1.6.0.20260101.backup.json"
        AssertFalse(FileExist(old))
        AssertEqual("original migration copy", FileRead(archive))
        source := folder "\EVE-X-Preview.json"
        FileAppend("current settings", source, "UTF-8-RAW")
        loop 7
            AssertTrue(SettingsBackup.Capture(source) != "")
        AssertEqual("original migration copy", FileRead(archive))
        AssertTrue(SettingsContract.Commands().Length > 0)
        found := false
        for descriptor in SettingsContract.Commands()
            found := found || (descriptor["id"] = "open-backups" && !descriptor["runtimeOnly"])
        AssertTrue(found)
    } finally SettingsBackup.Root := previousRoot
}

TestRunner.Register("Update1704 startup backups retain the last five exact snapshots per installation", Update1704BackupRetentionTest)
Update1704BackupRetentionTest() {
    previousRoot := SettingsBackup.Root
    folder := A_ScriptDir "\..\.vs\update1704-tests\backups-" A_TickCount
    first := folder "\first", second := folder "\second"
    DirCreate(first), DirCreate(second)
    SettingsBackup.Root := folder "\snapshots"
    try {
        for index in [1, 2, 3, 4, 5, 6, 7] {
            path := first "\EVE-X-Preview.json"
            if FileExist(path)
                FileDelete(path)
            FileAppend("snapshot " index, path, "UTF-8-RAW")
            AssertTrue(SettingsBackup.Capture(path) != "")
            AssertEqual("snapshot " index, FileRead(path))
        }
        kept := Map(), count := 0
        loop files SettingsBackup.Directory(first) "\*.json", "F" {
            kept[FileRead(A_LoopFileFullPath)] := true
            count++
        }
        AssertEqual(5, count)
        for index in [3, 4, 5, 6, 7]
            AssertTrue(kept.Has("snapshot " index))
        FileAppend("second installation", second "\EVE-X-Preview.json", "UTF-8-RAW")
        AssertTrue(SettingsBackup.Capture(second "\EVE-X-Preview.json") != "")
        AssertNotEqual(SettingsBackup.Directory(first), SettingsBackup.Directory(second))
    } finally SettingsBackup.Root := previousRoot
}

TestRunner.Register("Update1704 loader backs up original settings before normalization and backup failures do not block startup", Update1704BackupLoaderTest)
Update1704BackupLoaderTest() {
    previousRoot := SettingsBackup.Root, previousDirectory := A_WorkingDir
    folder := A_ScriptDir "\..\.vs\update1704-tests\loader-" A_TickCount
    DirCreate(folder)
    SettingsBackup.Root := folder "\snapshots"
    original := JSON.Dump(JSON.Load(default_JSON))
    try {
        SetWorkingDir(folder)
        FileAppend(original, "EVE-X-Preview.json", "UTF-8-RAW")
        Load_JSONUnlocked(false)
        count := 0
        loop files SettingsBackup.Directory(folder) "\*.json", "F" {
            AssertEqual(original, FileRead(A_LoopFileFullPath))
            count++
        }
        AssertEqual(1, count)
        blocked := folder "\blocked"
        FileAppend("file blocks directory creation", blocked)
        SettingsBackup.Root := blocked
        loaded := Load_JSONUnlocked(false)
        AssertTrue(loaded.Has("_Profiles"))
        AssertTrue(FileExist("EVE-X-Preview.json"))
    } finally {
        SetWorkingDir(previousDirectory)
        SettingsBackup.Root := previousRoot
    }
}

TestRunner.Register("Update1704 unfocused border delay expires resets on return and does not delay switching clients", Update1704BorderDelayTest)
Update1704BorderDelayTest() {
    previousHidden := A_DetectHiddenWindows
    DetectHiddenWindows(true)
    app := ThumbnailForegroundPollFixture(), sources := []
    outside := Gui()
    try {
        first := app.CreateSource("EVE - Delay first"), second := app.CreateSource("EVE - Delay second")
        sources.Push(first, second)
        app.EVE_WIN_Created(first.Hwnd, first.Title)
        app.EVE_WIN_Created(second.Hwnd, second.Title)
        app.foreground := first.Hwnd
        app.ShowActiveBorder(first.Hwnd)
        border := app.ThumbWindows.%first.Hwnd%["Border"]
        app.UnfocusedBorderClearDelay := 1000
        app.UpdateThumbnailFocusLoss(outside.Hwnd, 100)
        app.UpdateThumbnailFocusLoss(outside.Hwnd, 1099)
        AssertTrue(border.ActiveHighlight)
        app.UpdateThumbnailFocusLoss(first.Hwnd, 1100)
        app.UpdateThumbnailFocusLoss(outside.Hwnd, 1200)
        AssertTrue(border.ActiveHighlight)
        app.UpdateThumbnailFocusLoss(outside.Hwnd, 2200)
        AssertFalse(border.ActiveHighlight)
        app.ShowActiveBorder(first.Hwnd)
        app.UpdateThumbnailFocusLoss(second.Hwnd, 2300)
        AssertFalse(border.ActiveHighlight, "Another tracked client clears the previous active border immediately.")
        app.foreground := first.Hwnd
        app.ShowActiveBorder(first.Hwnd)
        app.UnfocusedBorderClearDelay := 0
        app.UpdateThumbnailFocusLoss(outside.Hwnd, 2400)
        AssertFalse(border.ActiveHighlight)
    } finally {
        app.Cleanup(sources)
        outside.Destroy()
        DetectHiddenWindows(previousHidden)
    }
}

TestRunner.Register("Update1704 both editors accept 30 ms group hold delay and reject smaller values", Update1704MinimumDelayTest)
Update1704MinimumDelayTest() {
    field := SettingsSchema.Fields["GroupsHoldDelay"]
    AssertEqual(30, SettingsValue.Parse(field, 30))
    AssertThrows(() => SettingsValue.Parse(field, 29))
    app := ColorSettingsFixture()
    try {
        app.HotkeyGroups_Ctrl()
        app.MainFrame["GroupsHoldDelay"].Value := 30
        SendMessage(0x111, (0x300 << 16) | DllCall("GetDlgCtrlID", "Ptr", app.MainFrame["GroupsHoldDelay"].Hwnd), app.MainFrame["GroupsHoldDelay"].Hwnd, app.MainFrame.Hwnd)
        Sleep(30)
        AssertEqual(30, app.GroupsHoldDelay)
    } finally {
        SetTimer(app.Save_Settings_Delay_Timer, 0)
        app.Dispose()
    }
}
