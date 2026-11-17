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

class CharacterGroupPriorityFixture extends Update1704ShortcutFixture {
    loggedOut := Map()
    scopeActive := true
    hasMathcingOldTitle(title) => this.loggedOut.Get(title, 0)
    ShortcutScopeActive() => this.scopeActive
}

TestRunner.Register("Character group first-available shortcuts yield to populated next and previous groups", CharacterGroupFallbackPriorityTest)
CharacterGroupFallbackPriorityTest() {
    app := CharacterGroupPriorityFixture()
    app._Hotkeys := []
    for direction in ["ForwardsHotkey", "BackwardsHotkey"] {
        for cycleName in ["A cycle", "Z cycle"] {
            fallback := Map("Characters", ["EVE - Fallback"], "ForwardsHotkey", "", "BackwardsHotkey", "", "FirstCharHotkey", "F16")
            cycle := Map("Characters", ["EVE - Cycle", "EVE - Other"], "ForwardsHotkey", "", "BackwardsHotkey", "", "FirstCharHotkey", "")
            cycle[direction] := "F16"
            app.Profiles["Default"]["Hotkey Groups"] := Map(cycleName, cycle, "M fallback", fallback)
            app.targets := Map("EVE - Fallback", true, "EVE - Other", true)
            app._Hotkeys := [Map("EVE - Direct", "F16")]
            app.targets["EVE - Direct"] := true
            AssertTrue(app.DirectShortcutAvailable("*F16"))
            AssertFalse(app.CharacterGroupCycleAvailable(cycle["Characters"], "*F16"), "A running direct character overrides group cycling.")
            AssertFalse(app.FirstCharacterGroupAvailable(fallback["Characters"], "*F16"), "A running direct character overrides first-available fallback.")
            app.targets.Delete("EVE - Direct")
            AssertFalse(app.DirectShortcutAvailable("*F16"))
            AssertTrue(app.CharacterGroupCycleAvailable(cycle["Characters"], "*F16"))
            AssertFalse(app.FirstCharacterGroupAvailable(fallback["Characters"], "*F16"), "Any available cycle member takes priority, regardless of group order.")
            app.targets.Delete("EVE - Other")
            AssertFalse(app.CharacterGroupCycleAvailable(cycle["Characters"], "*F16"))
            AssertTrue(app.FirstCharacterGroupAvailable(fallback["Characters"], "*F16"), "Closing the last cycle member enables first-available fallback immediately.")
            app.loggedOut["EVE - Cycle"] := 123
            AssertFalse(app.CharacterGroupCycleAvailable(cycle["Characters"], "*F16"), "A preserved login-screen title must not block an available fallback.")
            AssertTrue(app.FirstCharacterGroupAvailable(fallback["Characters"], "*F16"))
            app.loggedOut.Clear()
            app.targets["EVE - Cycle"] := true
            AssertFalse(app.FirstCharacterGroupAvailable(fallback["Characters"], "*F16"), "Launching a cycle member restores cycle priority immediately.")
            cycle[direction] := "F17"
            AssertTrue(app.FirstCharacterGroupAvailable(fallback["Characters"], "*F16"), "An unrelated cycle key must not reserve the fallback shortcut.")
            cycle[direction] := "^F16"
            AssertTrue(app.FirstCharacterGroupAvailable(fallback["Characters"], "*F16"), "An inactive modifier chord must not reserve the fallback shortcut.")
            app.scopeActive := false
            AssertFalse(app.CharacterGroupCycleAvailable(cycle["Characters"], "*F16"))
            AssertFalse(app.FirstCharacterGroupAvailable(fallback["Characters"], "*F16"))
            app.scopeActive := true
            app.activationHotkeysSuspended := true
            AssertFalse(app.FirstCharacterGroupAvailable(fallback["Characters"], "*F16"))
            app.activationHotkeysSuspended := false
            app.targets.Clear()
            AssertFalse(app.FirstCharacterGroupAvailable(fallback["Characters"], "*F16"))
        }
    }
    entries := []
    HotkeyBindings.Add(entries, "group:Cycle:ForwardsHotkey", "Next", "F16", true)
    HotkeyBindings.Add(entries, "group:Fallback:FirstCharHotkey", "First available", "F16", true)
    AssertEqual(0, HotkeyBindings.Conflicts(entries).Count, "Availability-based fallback is an intentional shared shortcut.")
    HotkeyBindings.Add(entries, "group:Other:BackwardsHotkey", "Previous", "F16", true)
    AssertTrue(HotkeyBindings.Conflicts(entries).Has("group:Cycle:ForwardsHotkey"), "Competing cycle actions still warn.")
}

class HeldRefreshFixture extends Update1704ShortcutFixture {
    foreground := 101
    refreshed := []
    borders := []
    during := 0
    held := false
    complete := false
    visibility := []
    _GetForegroundHwnd() => this.foreground
    ShowActiveBorder(hwnd) => this.borders.Push(hwnd)
    ShowThumb(hwnd, state) => this.visibility.Push([hwnd,state])
    UpdateThumb_AfterActivation(event?, hwnd?) {
        this.refreshed.Push(hwnd)
        if this.held
            this.during++
    }
    __New() {
        super.__New()
        this.ThumbnailPollingInterval := 50
        this.BorderActive := 0
        this.HideThumbnails := false
        this.HideThumbForActiveWin := false
        this.ThumbWindows := {}
        for hwnd in [101,102,103,104]
            this.ThumbWindows.%hwnd% := Map("Window",{},"Thumbnail",{THUMB_ID:1},"Border",{})
    }
    Hold(*) {
        this.held := true
        for hwnd in [101,102,103,104] {
            this.foreground := hwnd
            Sleep(90)
        }
        this.held := false
        this.complete := true
    }
}
TestRunner.Register("HeldThumbnailRefresh foreground refresh progresses inside a P1 switching callback", HeldThumbnailRefreshPriorityTest)
HeldThumbnailRefreshPriorityTest() {
    app := HeldRefreshFixture()
    try {
        app.StartForegroundThumbnailRefresh()
        RunHeldThumbnailCallback(app, "complete")
        AssertTrue(app.during>=3,"Refresh must progress before the held switching callback ends.")
        AssertEqual(104,app.refreshed[-1])
        AssertTrue(app.refreshed.Length<=4,"Unchanged foreground must not trigger repeated DWM work.")
    } finally {
        SetTimer(app.foregroundRefreshTimer,0)
    }
}

RunHeldThumbnailCallback(app, completionProperty) {
    ; Exercise the same P1 scheduling as a held switching hotkey without relying
    ; on synthetic keyboard input reaching a disconnected Sandbox desktop.
    callback := ObjBindMethod(app, "Hold")
    SetTimer(callback, -1, 1)
    try {
        deadline := A_TickCount + 2000
        while !app.%completionProperty% && A_TickCount < deadline
            Sleep(10)
        AssertTrue(app.%completionProperty%, "The scheduled P1 held callback must complete.")
    } finally SetTimer(callback, 0)
}

TestRunner.Register("HeldThumbnailRefresh priority-zero control cannot refresh during a P1 hold", HeldThumbnailRefreshStarvationTest)
HeldThumbnailRefreshStarvationTest() {
    app := HeldRefreshFixture()
    try {
        app.StartForegroundThumbnailRefresh()
        SetTimer(app.foregroundRefreshTimer,, 0)
        RunHeldThumbnailCallback(app, "complete")
        AssertEqual(0, app.during, "A lower-priority timer must be starved until the P1 callback returns.")
        Sleep(100)
        AssertEqual(104, app.refreshed[-1], "The deferred refresh resumes after the held callback.")
    } finally SetTimer(app.foregroundRefreshTimer, 0)
}
TestRunner.Register("HeldThumbnailRefresh pause hidden incomplete and DWM-busy states skip refresh", HeldThumbnailRefreshGuardTest)
HeldThumbnailRefreshGuardTest() {
    app := HeldRefreshFixture()
    for property in ["livePreviewsPaused","releasingLivePreviews","previewDwmCalls","HideThumbnails"] {
        app.%property% := true
        app.RefreshForegroundThumbnail()
        app.%property% := false
    }
    AssertEqual(0,app.refreshed.Length)
    app.foreground := 999
    app.RefreshForegroundThumbnail()
    app.ThumbWindows.999 := Map("Window",{})
    app.RefreshForegroundThumbnail()
    AssertEqual(0,app.refreshed.Length)
    app.foreground := 101
    app.ThumbWindows.101["Thumbnail"].THUMB_ID := 0
    app.RefreshForegroundThumbnail()
    AssertEqual(0,app.refreshed.Length)
    app.ThumbWindows.101["Thumbnail"].THUMB_ID := 1
    app.foreground := 101
    app.RefreshForegroundThumbnail()
    app.RefreshForegroundThumbnail()
    AssertEqual([101],app.refreshed)
}
TestRunner.Register("HeldThumbnailRefresh hiding the active thumbnail restores and refreshes the previous one", HeldThumbnailRefreshHiddenTest)
HeldThumbnailRefreshHiddenTest() {
    app := HeldRefreshFixture()
    app.HideThumbForActiveWin := true
    app.RefreshForegroundThumbnail()
    app.foreground := 102
    app.RefreshForegroundThumbnail()
    app.RefreshForegroundThumbnail()
    AssertEqual([[101,"Hide"],[101,"Show"],[102,"Hide"]],app.visibility)
    AssertEqual([101],app.refreshed)
}

class RealHeldRefreshFixture extends ThumbnailForegroundPollFixture {
    held := false
    during := 0
    completed := false
    sequence := []
    UpdateThumb_AfterActivation(event?, hwnd?) {
        thumbnail := this.ThumbWindows.%hwnd%["Thumbnail"]
        thumbnail.THUMB_UPDATED := false
        super.UpdateThumb_AfterActivation(,hwnd)
        if this.held && thumbnail.THUMB_UPDATED
            this.during++
    }
    Hold(*) {
        this.held := true
        for hwnd in this.sequence {
            this.foreground := hwnd
            Sleep(100)
        }
        this.held := false
        this.completed := true
    }
}
TestRunner.Register("HeldThumbnailRefresh real DWM previews update before a held P1 callback finishes", HeldThumbnailRefreshDwmTest)
HeldThumbnailRefreshDwmTest() {
    previous := A_DetectHiddenWindows
    DetectHiddenWindows(true)
    app := RealHeldRefreshFixture(), sources := []
    try {
        loop 3 {
            source := app.CreateSource("EVE - Held DWM " A_Index)
            sources.Push(source)
            app.EVE_WIN_Created(source.Hwnd,source.Title)
            app.sequence.Push(source.Hwnd)
        }
        app.foreground := app.sequence[1]
        app.StartForegroundThumbnailRefresh()
        RunHeldThumbnailCallback(app, "completed")
        AssertTrue(app.during>=2,"Real DWM update calls must complete while switching remains held.")
    } finally {
        if app.HasOwnProp("foregroundRefreshTimer")
            SetTimer(app.foregroundRefreshTimer,0)
        app.Cleanup(sources)
        DetectHiddenWindows(previous)
    }
}

class FeedbackCommandFixture extends Update1704ShortcutFixture {
    messages := []
    HotkeyFeedback(message) => this.messages.Push(message)
    ShowHideThumbnails() => this.HideThumbnails := !this.HideThumbnails
    Toggle_ClickThrough(*) => this.ClickThroughActive := !this.ClickThroughActive
    ToggleLivePreviews(*) => this.livePreviewsPaused := !this.livePreviewsPaused
    CloseActiveEVEWin(*) => false
    CloseAllEVEWindows(*) => 2
    RequestRuntimeAction(action, hotkey := false) => false
}

class FeedbackRuntimeFixture extends SettingsWorkbenchOwnerFixture {
    messages := []
    HotkeyFeedback(message) => this.messages.Push(message)
}
class FeedbackBlockingEditor {
    closed := false
    activeDialog := true
    UpdateStatus(*) {
    }
}
TestRunner.Register("CommandFeedback reload and exit receipts defer execution and recheck draft protection", CommandFeedbackRuntimeTest)
CommandFeedbackRuntimeTest() {
    app := FeedbackRuntimeFixture()
    AssertTrue(app.RequestRuntimeAction("reload",true))
    AssertEqual("Reload requested",app.messages[-1])
    AssertEqual(0,app.lifecycleActions.Length)
    Sleep(800)
    AssertEqual(["reload"],app.lifecycleActions)
    AssertTrue(app.RequestRuntimeAction("exit",true))
    app.settingsWorkbench := FeedbackBlockingEditor()
    Sleep(800)
    AssertEqual(["reload"],app.lifecycleActions)
    AssertTrue(InStr(app.messages[-1],"cancelled"))
}

TestRunner.Register("FirstEverFollowup recovery notice excludes existing users updates and additional installations", FirstEverNoticeTest)
FirstEverNoticeTest() {
    previous := SettingsBackup.Root
    folder := A_ScriptDir "\..\.vs\update1704-tests\first-ever-" A_TickCount
    SettingsBackup.Root := folder "\new-user"
    try {
        AssertTrue(SettingsBackup.PrepareFirstRun(folder "\install-one",false))
        AssertTrue(SettingsBackup.TakeFirstRunNotice(Map("RecoveryGuidancePending",true)))
        AssertFalse(SettingsBackup.TakeFirstRunNotice(Map("RecoveryGuidancePending",true)),"Restoring an old first-run snapshot must not repeat the notice.")
        AssertFalse(SettingsBackup.PrepareFirstRun(folder "\install-one",true))
        AssertFalse(SettingsBackup.PrepareFirstRun(folder "\install-two",false))
        SettingsBackup.Root := folder "\existing-user"
        AssertFalse(SettingsBackup.PrepareFirstRun(folder "\existing-install",true))
        AssertFalse(SettingsBackup.TakeFirstRunNotice(Map("RecoveryGuidancePending",true)))
        AssertFalse(SettingsBackup.PrepareFirstRun(folder "\new-install",false))
        SettingsBackup.Root := folder "\lost-settings"
        DirCreate(SettingsBackup.Directory(folder "\lost-install"))
        AssertFalse(SettingsBackup.PrepareFirstRun(folder "\lost-install",false))
    } finally SettingsBackup.Root := previous
}

TestRunner.Register("CommandFeedback command hotkeys confirm their resulting states and exclude switching", CommandFeedbackStateTest)
CommandFeedbackStateTest() {
    app := FeedbackCommandFixture()
    try {
        HotkeyFeedback.Run(app,"Suspend_Hotkeys_Hotkey")
        AssertEqual("EVE switching hotkeys paused",app.messages[-1])
        HotkeyFeedback.Run(app,"Suspend_Hotkeys_Hotkey")
        AssertEqual("EVE switching hotkeys resumed",app.messages[-1])
        HotkeyFeedback.Run(app,"Suspend_All_Hotkeys_Hotkey")
        AssertTrue(A_IsSuspended)
        AssertTrue(InStr(app.messages[-1],"same shortcut"))
        HotkeyFeedback.Run(app,"Suspend_All_Hotkeys_Hotkey")
        AssertFalse(A_IsSuspended)
        AssertEqual("All hotkeys resumed",app.messages[-1])
        HotkeyFeedback.Run(app,"HideThumbnailsHotkey")
        AssertEqual("Thumbnails hidden",app.messages[-1])
        HotkeyFeedback.Run(app,"HideThumbnailsHotkey")
        AssertEqual("Thumbnails shown",app.messages[-1])
        HotkeyFeedback.Run(app,"ClickThroughHotkey")
        AssertEqual("Thumbnail click-through enabled",app.messages[-1])
        HotkeyFeedback.Run(app,"ToggleLivePreviewsHotkey")
        AssertEqual("Live previews paused",app.messages[-1])
        HotkeyFeedback.Run(app,"ToggleLivePreviewsHotkey")
        AssertEqual("Live previews resumed",app.messages[-1])
        HotkeyFeedback.Run(app,"Close_Active_EVE_Win_Hotkey")
        AssertEqual("No focused EVE client to close",app.messages[-1])
        HotkeyFeedback.Run(app,"Close_All_EVE_Win_Hotkey")
        AssertEqual("Close requested for 2 EVE clients",app.messages[-1])
        HotkeyFeedback.Run(app,"Reload_Program_Hotkey")
        AssertTrue(InStr(app.messages[-1],"blocked"))
        HotkeyFeedback.Run(app,"Exit_Program_Hotkey")
        AssertTrue(InStr(app.messages[-1],"blocked"))
        count := app.messages.Length
        for id in ["SwitchToPreviousWindow_Hotkey","CycleEveryLoggedIn_Hotkey","Login_Screen_Cycle_Hotkey","QuickGroupHotkey","ActivateThumbnail"]
            HotkeyFeedback.Run(app,id)
        AssertEqual(count,app.messages.Length)
    } finally Suspend(false)
}

TestRunner.Register("CommandFeedback a newer receipt replaces the previous tooltip and restarts its expiry", CommandFeedbackTooltipTest)
CommandFeedbackTooltipTest() {
    visible(text) {
        for hwnd in WinGetList("ahk_class tooltips_class32 ahk_pid " DllCall("GetCurrentProcessId"))
            if InStr(WinGetTitle("ahk_id " hwnd),text)
                return true
        return false
    }
    try {
        HotkeyFeedback.Show("First command receipt")
        AssertTrue(visible("First command receipt"))
        Sleep(1000)
        HotkeyFeedback.Show("Latest command receipt")
        Sleep(650)
        AssertTrue(visible("Latest command receipt"))
        Sleep(1000)
        AssertFalse(visible("Latest command receipt"))
    } finally HotkeyFeedback.Hide()
}

TestRunner.Register("RecoveryFollowup guidance chooses a readable backup and skips damaged newer copies", RecoveryFollowupSuggestionTest)
RecoveryFollowupSuggestionTest() {
    root := SettingsBackup.Root
    folder := A_ScriptDir "\..\.vs\update1704-tests\suggestion-" A_TickCount
    DirCreate(folder)
    SettingsBackup.Root := folder "\backups"
    try {
        good := folder "\EVE-X-Preview.old.backup.json"
        FileAppend(default_JSON,good,"UTF-8")
        bad := folder "\EVE-X-Preview.new.backup.json"
        FileAppend("broken json",bad,"UTF-8")
        FileSetTime("20250101000000",good,"M")
        FileSetTime("20260101000000",bad,"M")
        AssertEqual(SettingsBridge.NormalizeInstallation(good),SettingsBridge.NormalizeInstallation(SettingsBackup.LatestValid(folder)))
        SettingsBackup.CollectMigrationBackups(folder)
        SettingsBackup.SuggestedRepair := SettingsBackup.LatestValid(folder)
        AssertTrue(InStr(SettingsBridge.NormalizeInstallation(SettingsBackup.SuggestedRepair),SettingsBridge.NormalizeInstallation(SettingsBackup.Directory(folder))))
        text := SettingsBackup.HelpText(folder)
        AssertTrue(InStr(text,"discord.gg/d9EzgUXnT6"))
        AssertTrue(InStr(text,"github.com/khivus/EVE-X-Preview/issues/new"))
        AssertTrue(InStr(text,"older installation"))
        window := SettingsBackup.ShowGuidance(folder)
        AssertTrue(DllCall("IsWindowVisible","Ptr",window.Hwnd))
        window.Destroy()
    } finally SettingsBackup.Root := root
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
    AssertTrue(app.DirectShortcutAvailable("*CapsLock"), "Quick Group priority is resolved in the callback, not in this predicate.")
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
        app.foreground := second.Hwnd
        app.ShowActiveBorder(second.Hwnd)
        AssertTrue(app.ThumbWindows.%second.Hwnd%["Border"].ActiveHighlight, "Confirmed EVE activation refreshes its border immediately.")
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
