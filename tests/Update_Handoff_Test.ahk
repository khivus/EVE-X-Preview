TestRunner.Register("Update launch reports deferred installer errors through status", UpdateLaunchStatusTest)
UpdateLaunchStatusTest() {
    owner := SettingsServiceFixture(JSON.Load(default_JSON))
    session := SettingsSession(owner, 0, true)
    try {
        facade := session.Facade()
        facade.updateChecker.releaseTag := "1.7.0.8"
        session.Command("install-stable")
        AssertTrue(session.Command("update-status")["busy"], "The editor must poll while the deferred handoff is pending.")
        AssertThrows(() => session.Command("install-stable"), "already in progress")
        Sleep(100)
        ; The source runner cannot install itself. This real deferred failure
        ; must survive the next status poll, just like staging/UAC errors.
        status := session.Command("update-status")
        AssertTrue(InStr(status["status"], "compiled installation"), "The editor must receive the install failure instead of an apparent no-op.")
        AssertFalse(status["busy"])
        AssertTrue(status["installError"])
        AssertEqual(status["status"], session.Command("update-status")["status"], "Polling must not erase the error.")
    } finally session.Dispose()
}

TestRunner.Register("Update launch reports unexpected deferred exceptions and allows retry", UpdateLaunchExceptionTest)
UpdateLaunchExceptionTest() {
    owner := SettingsServiceFixture(JSON.Load(default_JSON))
    session := SettingsSession(owner, 0, true)
    try {
        facade := session.Facade()
        facade.DefineProp("InstallVersion", {Call: UpdateLaunchThrow})
        session.Command("install-prerelease")
        Sleep(100)
        status := session.Command("update-status")
        AssertTrue(InStr(status["status"], "handoff fixture failure"))
        AssertTrue(status["installError"])
        AssertFalse(status["busy"])
        facade.DeleteProp("InstallVersion")
        session.Command("install-stable")
        status := session.Command("update-status")
        AssertTrue(status["busy"])
        AssertFalse(status["installError"], "Retry clears the previous failure.")
        Sleep(100)
        AssertTrue(InStr(session.Command("update-status")["status"], "Check for updates"))
    } finally session.Dispose()
}
UpdateLaunchThrow(*) {
    throw Error("handoff fixture failure")
}

TestRunner.Register("Settings update waits for the previous editor process after its lease is released", UpdateEditorExitTest)
UpdateEditorExitTest() {
    folder := A_Temp "\evex-editor-exit-" A_TickCount
    DirCreate(folder)
    child := folder "\editor.ahk", signal := folder "\closing.txt"
    FileAppend('
    (
    #Requires AutoHotkey v2.0
    #SingleInstance Off
    #NoTrayIcon
    window := Gui(, "Update handoff fixture")
    window.OnEvent("Close", Closing)
    window.Show("NoActivate x-3000 y-3000 w200 h80")
    Closing(*) {
        window.Hide()
        FileAppend("closing", A_Args[1])
        SetTimer((*) => ExitApp(), -600)
    }
    )', child)
    owner := SettingsServiceFixture(JSON.Load(default_JSON))
    bridge := SettingsBridge(owner, "runtime", folder, false)
    pid := 0
    ReleaseLease(*) {
        if !FileExist(signal)
            return
        if IsObject(bridge.session)
            bridge.session.Dispose()
        bridge.session := 0
        bridge.ExternalPid := 0
    }
    try {
        Run('"' A_AhkPath '" /ErrorStdOut "' child '" "' signal '"', folder,, &pid)
        AssertTrue(WinWait("Update handoff fixture ahk_pid " pid,, 5))
        bridge.ExternalPid := pid
        bridge.session := SettingsSession(owner)
        bridge.session.clientDirty := true
        AssertFalse(bridge.CloseExternalForUpdate(), "An unfinished settings draft must prevent closure.")
        AssertTrue(ProcessExist(pid))
        AssertFalse(FileExist(signal))
        bridge.session.clientDirty := false
        SetTimer(ReleaseLease, 20)
        started := A_TickCount
        AssertTrue(bridge.CloseExternalForUpdate())
        AssertFalse(ProcessExist(pid), "Releasing the lease is insufficient: the previous editor must exit.")
        AssertTrue(A_TickCount - started >= 500, "Update handoff returned before the old editor exited.")
    } finally {
        SetTimer(ReleaseLease, 0)
        if pid && ProcessExist(pid)
            ProcessClose(pid)
        bridge.Dispose()
        DirDelete(folder, true)
    }
}

TestRunner.Register("Character group hotkeys are presented Next Previous First by the shared settings service", CharacterGroupShortcutOrderTest)
CharacterGroupShortcutOrderTest() {
    owner := SettingsServiceFixture(JSON.Load(default_JSON))
    session := SettingsSession(owner)
    try {
        session.model.AddGroup("characters", "Ordered group")
        session.model.SetGroupShortcut("characters", "Ordered group", "ForwardsHotkey", "F13")
        session.model.SetGroupShortcut("characters", "Ordered group", "BackwardsHotkey", "F14")
        session.model.SetGroupShortcut("characters", "Ordered group", "FirstCharHotkey", "F15")
        shortcuts := session.Rows("group-members", "Ordered group")["groupShortcuts"]
        keys := [], labels := [], values := []
        for row in shortcuts {
            keys.Push(row["key"]), labels.Push(row["label"]), values.Push(row["value"])
        }
        AssertEqual(["ForwardsHotkey", "BackwardsHotkey", "FirstCharHotkey"], keys)
        AssertEqual(["Next character", "Previous character", "First available character"], labels)
        AssertEqual(["F13", "F14", "F15"], values, "Reordering must preserve each hotkey's mapping.")
    } finally session.Dispose()
}
