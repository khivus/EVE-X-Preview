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
