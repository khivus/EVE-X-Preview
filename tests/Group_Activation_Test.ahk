; Hidden real HWNDs for title/existence checks; foreground sequences are simulated
; so tests do not steal focus from the user's running EVE clients.
class GroupActivationFixture extends ThumbnailNamesFixture {
    __New() {
        super.__New()
        this._JSON := JSON.Load(default_JSON)
        this.ProfileHotkeysGroups := "Default"
        this.PreserveHotkeysOnLogout := false
        this.MaxActiveWindowRetries := 3
        this.ActiveWindowRetryInterval := 1
        this.ThumbWindows := {}
        this.foregrounds := []
        this.reads := 0
    }
    _GetForegroundHwnd() {
        this.reads++
        return this.foregrounds[Min(this.reads, this.foregrounds.Length)]
    }
    Check(group, sequence, pending := 0) {
        this.foregrounds := sequence
        this.reads := 0
        this.pendingEVEActivation := pending
        return this._GetCurrentGroupIndex(group)
    }
}

class QuickGroupOrderFixture extends GroupActivationFixture {
    __New() {
        super.__New()
        this.ProfileThumbnailsInteractions := "Default"
        this.QuickGroupEnabled := true
        this.QuickGroupChars := Map()
        this.QuickGroupOrder := []
        this.DisabledChars := Map()
        this.BorderActive := 0
    }
    toggleColorBorder(*) {
    }
}

TestRunner.Register("Quick Group cycles in added order or name order", QuickGroupOrderTest)
QuickGroupOrderTest() {
    app := QuickGroupOrderFixture()
    AssertEqual("Added order", app.QuickGroupSortOrder)
    app.AddToQuickGroup(1003, "Zelda")
    app.AddToQuickGroup(1001, "Alice")
    app.AddToQuickGroup(1002, "Alice")
    AssertEqual([1003, 1001, 1002], app.GetQuickGroupWindowOrder())
    app.QuickGroupSortOrder := "Name"
    AssertEqual([1001, 1002, 1003], app.GetQuickGroupWindowOrder())
    app.DeleteFromQuickGroup(1001)
    app.AddToQuickGroup(1001, "Alice")
    app.QuickGroupSortOrder := "Added order"
    AssertEqual([1003, 1002, 1001], app.GetQuickGroupWindowOrder(), "Readded windows belong at the end.")
    app.DisabledFromGroupsEnabled := true
    app.AddToDisabled(1002, "Alice")
    AssertEqual([1003, 1001], app.GetQuickGroupWindowOrder(), "Disabling a member removes it from Quick Group order.")
}

TestRunner.Register("Foreground polling recovers from missing and closed windows", ForegroundPollingTest)
TestRunner.Register("Close-all protects group HWNDs on login screens without protecting unrelated login clients", GroupCloseAllProtectionTest)
GroupCloseAllProtectionTest() {
    oldHidden := A_DetectHiddenWindows
    DetectHiddenWindows(true)
    app := QuickGroupOrderFixture()
    app.ProfileClientSettings := "Default"
    app.DontCloseOnLoginScreen := false
    app.dontCloseActiveClient := false
    app.DontCloseClients := ""
    app.DontCloseQuickGroupClients := true
    app.DontCloseDisabledClients := true
    app.DisabledFromGroupsEnabled := true
    app.EVEExe := "ahk_group CloseProtectionFixture"
    windows := [], closed := Map()
    try {
        loop 3 {
            window := Gui(, "EVE")
            windows.Push(window)
            window.Show("Hide w100 h60")
            GroupAdd("CloseProtectionFixture", "ahk_id " window.Hwnd)
            window.OnEvent("Close", RecordClose.Bind(window.Hwnd))
        }
        quick := windows[1].Hwnd, disabled := windows[2].Hwnd, ordinary := windows[3].Hwnd
        app.AddToQuickGroup(quick, "EVE - Old Quick Pilot")
        app.AddToDisabled(disabled, "EVE - Old Disabled Pilot")
        ; The title has changed, but the main poll has not updated either stored name.
        AssertTrue(app.DontCloseWIn("EVE", quick))
        AssertTrue(app.DontCloseWIn("EVE", disabled))
        AssertFalse(app.DontCloseWIn("EVE", ordinary))
        app.CloseAllEVEWindows()
        Sleep(50)
        AssertFalse(closed.Has(quick), "Quick Group must be protected even with login-screen protection off.")
        AssertFalse(closed.Has(disabled), "Disabled protection must use the window, not its previous character name.")
        AssertTrue(closed.Has(ordinary), "A protected login client must not protect every window titled EVE.")
        app.DontCloseQuickGroupClients := false
        AssertFalse(app.DontCloseWIn("EVE", quick))
        app.DontCloseOnLoginScreen := true
        AssertTrue(app.DontCloseWIn("EVE", ordinary))
        app.DontCloseOnLoginScreen := false
        app.DontCloseClients := "Alice"
        AssertTrue(app.DontCloseWIn("EVE - Alice", ordinary), "The named-client exclusion must still work.")
    } finally {
        for window in windows
            window.Destroy()
        DetectHiddenWindows(oldHidden)
    }

    RecordClose(hwnd, *) {
        closed[hwnd] := true
        return true ; Keep these test windows alive for the remaining checks.
    }
}

TestRunner.Register("Logout membership defaults migrate without overwriting per-profile opt-outs", GroupPersistenceDefaultsTest)
GroupPersistenceDefaultsTest() {
    defaults := JSON.Load(default_JSON)
    settings := DeepClone(defaults)
    profile := settings["_Profiles"]["Default"]
    profile["Hotkeys Settings"].Delete("PreserveQuickGroupOnLogout")
    profile["Hotkeys Settings"].Delete("PreserveDisabledOnLogout")
    settings["_Profiles"]["Opt out"] := DeepClone(profile)
    settings["_Profiles"]["Opt out"]["Hotkeys Settings"]["PreserveQuickGroupOnLogout"] := 0
    settings["_Profiles"]["Opt out"]["Hotkeys Settings"]["PreserveDisabledOnLogout"] := 0
    merged := JsonMergeNoOverwrite(defaults, settings)
    app := QuickGroupOrderFixture()
    app._JSON := JSON.Load(JSON.Dump(merged))
    AssertEqual(1, app.PreserveQuickGroupOnLogout)
    AssertEqual(1, app.PreserveDisabledOnLogout)
    app.ProfileThumbnailsInteractions := "Opt out"
    AssertEqual(0, app.PreserveQuickGroupOnLogout)
    AssertEqual(0, app.PreserveDisabledOnLogout)
}

ForegroundPollingTest() {
    previousDetectHidden := A_DetectHiddenWindows
    DetectHiddenWindows(true)
    app := GroupActivationFixture()
    window := Gui(, "Foreground polling test")
    closed := Gui(, "Closed foreground test")
    closedHwnd := closed.Hwnd
    closed.Destroy()
    savedLog := ProgramLog.Text
    try {
        hwnd := window.Hwnd
        exe := WinGetProcessName("ahk_id " hwnd)
        app.foregrounds := [0, closedHwnd, hwnd, hwnd, 0, hwnd]
        info := app._GetForegroundInfo(hwnd, "stale.exe")
        AssertEqual(0, info.hwnd)
        AssertEqual("", info.exe, "Missing foreground clears the old process name.")
        info := app._GetForegroundInfo(info.hwnd, info.exe)
        AssertEqual(0, info.hwnd)
        AssertEqual("", info.exe, "A window closed before lookup is an expected race.")
        info := app._GetForegroundInfo(info.hwnd, info.exe)
        AssertEqual(hwnd, info.hwnd)
        AssertEqual(exe, info.exe, "The next valid foreground must resolve normally.")
        info := app._GetForegroundInfo(info.hwnd, info.exe)
        AssertEqual(exe, info.exe)
        info := app._GetForegroundInfo(info.hwnd, info.exe)
        AssertEqual("", info.exe)
        info := app._GetForegroundInfo(info.hwnd, info.exe)
        AssertEqual(exe, info.exe, "Returning to the same window after a gap refreshes the cache.")
        AssertEqual(6, app.reads, "Capture the foreground only once per poll.")
        AssertEqual(savedLog, ProgramLog.Text, "Expected focus transitions must not generate errors.")
    } finally {
        window.Destroy()
        DetectHiddenWindows(previousDetectHidden)
    }
}

class ForegroundLookupFailureFixture extends GroupActivationFixture {
    failure := 0
    lookups := 0
    _GetWindowProcessName(hwnd) {
        this.lookups++
        if IsObject(this.failure)
            throw this.failure
        return super._GetWindowProcessName(hwnd)
    }
}

TestRunner.Register("Access-denied foreground lookups stay untracked and retry without log floods", ForegroundAccessDeniedTest)
ForegroundAccessDeniedTest() {
    previousDetectHidden := A_DetectHiddenWindows
    DetectHiddenWindows(true)
    app := ForegroundLookupFailureFixture()
    window := Gui()
    savedLog := ProgramLog.Text
    try {
        app.foregrounds := [window.Hwnd]
        app.failure := OSError(5)
        info := {hwnd: 0, exe: "stale.exe"}
        loop 20 {
            info := app._GetForegroundInfo(info.hwnd, info.exe)
            AssertEqual(window.Hwnd, info.hwnd)
            AssertEqual("", info.exe)
        }
        AssertEqual(20, app.lookups, "A denied lookup must not cache a stale process name.")
        AssertEqual(savedLog, ProgramLog.Text)
        app.failure := 0
        info := app._GetForegroundInfo(info.hwnd, info.exe)
        AssertEqual(WinGetProcessName("ahk_id " window.Hwnd), info.exe)
        app.failure := OSError(6)
        AssertThrows(() => app._GetForegroundInfo(0, ""), app.failure.Message, "Other OS failures remain visible.")
        app.failure := MemoryError("Out of memory.")
        AssertThrows(() => app._GetForegroundInfo(0, ""), "Out of memory.")
    } finally {
        window.Destroy()
        DetectHiddenWindows(previousDetectHidden)
    }
}

TestRunner.Register("Group lookup waits only for a relevant pending activation", GroupActivationTest)
GroupActivationTest() {
    previousDetectHidden := A_DetectHiddenWindows
    DetectHiddenWindows(true)
    app := GroupActivationFixture()
    first := Gui(, "Group test first")
    second := Gui(, "Group test second")
    other := Gui(, "Group test other")
    try {
        a := first.Hwnd, b := second.Hwnd, c := other.Hwnd
        pending() => {hwnd: b, source: a, expires: A_TickCount + 250}
        AssertEqual(-1, app.Check([b], [a]))
        AssertEqual(1, app.reads, "Switching groups must not retry without a pending request.")
        AssertEqual(-1, app.Check([b], [c]))
        AssertEqual(1, app.reads, "Entering EVE from another app must not retry.")
        AssertEqual(-1, app.Check([c], [a], pending()))
        AssertEqual(1, app.reads, "A request into another group must not delay this group.")
        AssertEqual(2, app.Check([a, b], [a, b], pending()))
        AssertEqual(2, app.reads, "Wait for a pending member even when the old foreground is in the group.")
        AssertEqual(0, app.pendingEVEActivation, "Observed focus confirms and clears the pending request.")
        AssertEqual(1, app.Check([b], [0, b], pending()))
        AssertEqual(-1, app.Check([b], [c], pending()))
        AssertEqual(1, app.reads)
        AssertEqual(0, app.pendingEVEActivation, "Unrelated focus cancels the wait.")
        expired := pending()
        expired.expires := A_TickCount - 1
        AssertEqual(-1, app.Check([b], [a], expired))
        AssertEqual(1, app.reads)
        AssertEqual(1, app.Check([a, b], [a], pending()))
        AssertEqual(0, app.pendingEVEActivation, "Exhausted retries cannot leave a stale pending request.")
        AssertTrue(app.reads <= 4, "Retry count remains bounded.")
        app.MaxActiveWindowRetries := 1
        AssertEqual(-1, app.Check([b], [a], pending()))
        AssertEqual(0, app.pendingEVEActivation)
        app.PreserveHotkeysOnLogout := true
        app.ThumbWindows.%a% := Map("Window", {OldTitle: "EVE - Logged Out Pilot"})
        AssertEqual(1, app.Check([a], [a]), "Preservation must not resolve the HWND through an old title.")
        AssertEqual(1, app.Check(["EVE - Logged Out Pilot"], [a]))
        AssertEqual(1, app.Check(["Group test other"], [c]), "Untracked foreground titles must not throw.")
        second.Destroy()
        AssertEqual(-1, app.Check([b], [a], pending()))
        AssertEqual(1, app.reads, "Destroyed targets cannot be pending.")
    } finally {
        first.Destroy()
        second.Destroy()
        other.Destroy()
        DetectHiddenWindows(previousDetectHidden)
    }
}
