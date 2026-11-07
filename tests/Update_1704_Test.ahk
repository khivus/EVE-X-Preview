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
