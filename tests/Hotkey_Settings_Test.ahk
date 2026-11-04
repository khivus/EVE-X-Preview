TestRunner.Register("Hotkey conflicts recognize aliases, modifier order, custom prefixes and wildcards", HotkeyOverlapTest)
HotkeyOverlapTest() {
    for pair in [["^!F23", "!^f23"], ["Ctrl & 1", "^1"], ["Escape", "Esc"], ["^a", "^vk41"],
        ["~$^F1", "^F1"], ["*F1", "^!F1"], ["<^F1", "^F1"], ["XButton1 & 1", "xbutton1 & 1"],
        ["^XButton1 & 1", "^XButton1 & 1"], ["^XButton1 & 1", "XButton1 & 1"], ["Ctrl & 1", "^!1"],
        ["^~$F1", "^F1"], ["Ctrl", "RControl"]] {
        AssertTrue(HotkeyBindings.Overlap(HotkeyBindings.Parse(pair[1]), HotkeyBindings.Parse(pair[2])), pair[1] " should overlap " pair[2])
    }
    for pair in [["^F1", "!F1"], ["F1", "^F1"], ["F1", "F1 Up"], ["<^F1", ">^F1"],
        ["XButton1 & 1", "XButton2 & 1"], ["Enter", "NumpadEnter"]] {
        AssertFalse(HotkeyBindings.Overlap(HotkeyBindings.Parse(pair[1]), HotkeyBindings.Parse(pair[2])), pair[1] " should differ from " pair[2])
    }
}

TestRunner.Register("Hotkey recording preserves modifiers, mouse chords, wheel, cancellation and key release", HotkeyRecordingTest)
HotkeyRecordingTest() {
    recorder := HotkeyCapture()
    for key in ["LControl", "RAlt", "LShift", "LWin", "F23"]
        recorder.Pressed(key)
    AssertEqual("^!+#F23", recorder.result)
    recorder.Released("F23")
    AssertFalse(recorder.finished, "Wait until held modifiers are released.")
    for key in ["LControl", "RAlt", "LShift", "LWin"]
        recorder.Released(key)
    AssertTrue(recorder.finished)
    AssertFalse(recorder.cancelled)
    recorder.Cancel()
    AssertFalse(recorder.cancelled, "A queued hook-end notification must not cancel an already completed recording.")
    recorder := HotkeyCapture()
    recorder.Pressed("LControl")
    recorder.Pressed("XButton1")
    recorder.Pressed("1")
    AssertEqual("^XButton1 & 1", recorder.result)
    for key in ["1", "XButton1", "LControl"]
        recorder.Released(key)
    AssertTrue(recorder.finished)
    recorder := HotkeyCapture()
    recorder.Pressed("WheelDown")
    recorder.Released("WheelDown")
    AssertEqual("WheelDown", recorder.result)
    AssertTrue(recorder.finished)
    recorder := HotkeyCapture()
    recorder.Pressed("Escape")
    AssertTrue(recorder.cancelled)
    AssertEqual("", recorder.result)
    recorder := HotkeyCapture()
    recorder.Pressed("LControl")
    recorder.Pressed("Escape")
    recorder.Released("Escape")
    recorder.Released("LControl")
    AssertEqual("^Escape", recorder.result)
    AssertFalse(recorder.cancelled)
    AssertEqual("a`n^F23`n`nc`n", HotkeyBindings.ReplaceLine("a`r`nb`r`n`r`nc`r`n", 2, "^F23"))
    AssertEqual("a`n`nF24", HotkeyBindings.ReplaceLine("a", 3, "F24"))
}

class HotkeySettingsFixture extends CombatEventsFixture {
    __New() {
        super.__New()
        this.ProfileOverride()
        this.SetState()
        this.saves := 0
        this.Save_Settings_Delay_Timer := (*) => this.saves++
        this.S_Gui := Gui()
        this.CreateSidebar()
        this.CreateMainFrame()
        this.HotkeyGroups_Ctrl()
        this.HotkeysSettings_Ctrl()
        this.ThumbnailsInteractions_Ctrl()
        this.NonEVEApps_Ctrl()
        this.InitHotkeyHelpers()
    }
    Find(name) {
        for helper in this.hotkeyFields
            if helper.field.Name = name
                return helper
        throw Error("Missing hotkey helper: " name)
    }
    Dispose() {
        this.StopHotkeyHelpers()
        SetTimer(this.Save_Settings_Delay_Timer, 0)
        this.Sidebar.Destroy()
        this.MainFrame.Destroy()
        this.S_Gui.Destroy()
    }
}

TestRunner.Register("Every hotkey field has compact recorder and warning controls within its row", HotkeyHelpersLayoutTest)
HotkeyHelpersLayoutTest() {
    app := HotkeySettingsFixture()
    try {
        AssertEqual(20, app.hotkeyFields.Length)
        for helper in app.hotkeyFields {
            helper.field.GetPos(&x, &y, &w, &h)
            helper.button.GetPos(&bx, &by, &bw, &bh)
            helper.warning.GetPos(&wx, &wy, &ww, &wh)
            helper.waiting.GetPos(&px, &py, &pw, &ph)
            AssertTrue(x + w <= bx && bx + bw <= wx)
            AssertEqual(x + w + 1, bx)
            AssertEqual(y - 1, by)
            AssertEqual(y, wy)
            AssertEqual(26, bw)
            AssertEqual(Min(h, 25) + 2, bh)
            limit := helper.field.Name = "ProfileActivationHotkey" ? app.sidebarW : app.contentW
            AssertTrue(wx + ww <= limit, helper.field.Name " exceeds its panel.")
            AssertEqual([x, y, w, h], [px, py, pw, ph], "Waiting prompt must use the existing field rectangle.")
        }
        app.enableCtrlsInGroupsSettings(0)
        app.enableCtrlsInNonEVEGroupsSettings(0)
        app.RefreshHotkeyWarnings()
        AssertFalse(app.Find("ForwardsKey").button.Enabled)
        AssertFalse(app.Find("NonEVEForwardsKey").button.Enabled)
        app.enableCtrlsInGroupsSettings(1)
        app.RefreshHotkeyWarnings()
        AssertTrue(app.Find("ForwardsKey").button.Enabled)
    } finally {
        app.Dispose()
    }
}

TestRunner.Register("Captured scalar and selected-row binds use existing saves; duplicate imports remain allowed", HotkeyHelpersSaveTest)
HotkeyHelpersSaveTest() {
    app := HotkeySettingsFixture()
    path := A_Temp "\EVE-X-Preview-duplicate-hotkeys-" A_TickCount ".json"
    try {
        app.ApplyRecordedHotkey(app.Find("Exit_Program_Hotkey"), "^!F23")
        app.ApplyRecordedHotkey(app.Find("Suspend_Hotkeys_Hotkey"), "!^f23")
        Sleep(30)
        AssertEqual("^!F23", app.Exit_Program_Hotkey)
        AssertEqual("!^f23", app.Suspend_Hotkeys_Hotkey)
        app.RefreshHotkeyWarnings()
        AssertEqual("!", app.Find("Exit_Program_Hotkey").warning.Text)
        AssertTrue(InStr(app.Find("Exit_Program_Hotkey").message, "Suspend all hotkeys"))
        AssertTrue(InStr(app.Find("Exit_Program_Hotkey").message, "Saving and importing are still allowed"))
        app.MainFrame["HotkeyCharList"].Value := "Alice`n`nBob"
        app.MainFrame["HotkeyList"].Value := "F24`n`nF22"
        app.ApplyRecordedHotkey(app.Find("HotkeyList"), "^!F23", 3)
        Sleep(30)
        AssertEqual("F24`n`n^!F23", app.MainFrame["HotkeyList"].Value)
        AssertEqual("F24", app._Hotkeys["EVE - Alice"])
        AssertEqual("^!F23", app._Hotkeys["EVE - Bob"])
        app.RefreshHotkeyWarnings()
        AssertTrue(InStr(app.Find("HotkeyList").message, "Line 3:"))
        app.SaveJsonToFile(path)
        imported := JSON.Load(FileRead(path))
        AssertEqual("^!F23", imported["_Profiles"]["Default"]["Hotkeys Settings"]["Exit_Program_Hotkey"])
        AssertEqual("!^f23", imported["_Profiles"]["Default"]["Hotkeys Settings"]["Suspend_Hotkeys_Hotkey"])
        app._JSON := imported
        app.ProfileOverride()
        app.RefreshHotkeyWarnings()
        AssertEqual("!", app.Find("Exit_Program_Hotkey").warning.Text)
        app.ApplyRecordedHotkey(app.Find("Exit_Program_Hotkey"), "F21")
        Sleep(30)
        app.RefreshHotkeyWarnings()
        AssertEqual("", app.Find("Exit_Program_Hotkey").warning.Text)
    } finally {
        app.Dispose()
        if FileExist(path)
            FileDelete(path)
    }
}

TestRunner.Register("Conflict collection follows effective settings and all profile activation shortcuts", HotkeyCollectionTest)
HotkeyCollectionTest() {
    app := CombatEventsFixture()
    app.Profiles["Other"] := JSON.Load(JSON.Dump(app.Profiles["Default"]))
    app.LastUsedProfile := "Other"
    app.Profiles["Default"]["Other"]["Global_Groups"]["Hotkeys Settings"] := 1
    app.Profiles["Other"]["Other"]["Global_Groups"]["Hotkeys Settings"] := 0
    app.Profiles["Default"]["ActivationHotkey"] := "F23"
    app.Profiles["Default"]["Hotkeys Settings"]["Exit_Program_Hotkey"] := "F23"
    app.Profiles["Other"]["Hotkeys Settings"]["Exit_Program_Hotkey"] := "F24"
    app.ProfileOverride()
    conflicts := HotkeyBindings.Conflicts(HotkeyBindings.Collect(app))
    AssertTrue(conflicts.Has("profile:Default"))
    AssertTrue(conflicts.Has("setting:Exit_Program_Hotkey"))
    app.Profiles["Other"]["Other"]["Global_Groups"]["Hotkeys Settings"] := 1
    app.ProfileOverride()
    conflicts := HotkeyBindings.Conflicts(HotkeyBindings.Collect(app))
    AssertFalse(conflicts.Has("setting:Exit_Program_Hotkey"), "An inactive profile's ordinary shortcut cannot conflict.")
}

TestRunner.Register("Profile, group, Quick Group and app recorders save and show cross-section conflicts", HotkeyHelpersCategoriesTest)
HotkeyHelpersCategoriesTest() {
    app := HotkeySettingsFixture()
    try {
        app.Profiles["Default"]["Hotkey Groups"]["Fleet"] := Map("Characters", ["EVE - Alice"],
            "ForwardsHotkey", "", "BackwardsHotkey", "", "FirstCharHotkey", "")
        app.MainFrame["HotkeyGroupDDL"].Add(["Fleet"])
        app.MainFrame["HotkeyGroupDDL"].Choose("Fleet")
        app.Profiles["Default"]["Non-EVE Applications"]["NonEVEGroups"]["Tools"] := Map(
            "exe", ["notepad.exe"], "title", [""], "fkey", "", "bkey", "")
        app.MainFrame["NonEVEGroupsDDL"].Add(["Tools"])
        app.MainFrame["NonEVEGroupsDDL"].Choose("Tools")
        app.NonEVEHotkeys := Map("exe", ["notepad.exe"], "title", [""], "hotkey", [""])
        for name in ["ProfileActivationHotkey", "ForwardsKey", "QuickGroupHotkey", "NonEVEBackwardsdKey", "NonEVEHotkeysHtks"]
            app.ApplyRecordedHotkey(app.Find(name), "^F23")
        Sleep(30)
        AssertEqual("^F23", app.ProfileActivationHotkey)
        AssertEqual("^F23", app.Hotkey_Groups["Fleet"]["ForwardsHotkey"])
        AssertEqual("^F23", app.QuickGroupHotkey)
        AssertEqual("^F23", app.NonEVEGroups["Tools"]["bkey"])
        AssertEqual("^F23", app.NonEVEHotkeys["hotkey"][1])
        app.RefreshHotkeyWarnings()
        for name in ["ProfileActivationHotkey", "ForwardsKey", "QuickGroupHotkey", "NonEVEBackwardsdKey", "NonEVEHotkeysHtks"]
            AssertEqual("!", app.Find(name).warning.Text, name " must show the cross-section conflict.")
        AssertTrue(InStr(app.Find("ProfileActivationHotkey").message, "Group Fleet"))
        AssertTrue(InStr(app.Find("ForwardsKey").message, "App group Tools"))
    } finally {
        app.Dispose()
    }
}

TestRunner.Register("Inline recording keeps the original value until release and restores hotkey state", InlineHotkeyRecordingTest)
InlineHotkeyRecordingTest() {
    app := HotkeySettingsFixture()
    wasSuspended := A_IsSuspended
    try {
        helper := app.Find("Exit_Program_Hotkey")
        for suspended in [false, true] {
            Suspend(suspended)
            helper.field.Value := "F24"
            app.Exit_Program_Hotkey := "F24"
            wasVisible := helper.field.Visible
            app.RecordHotkey(helper)
            capture := HotkeyCapture.Current
            AssertTrue(HotkeyCapture.Active)
            AssertFalse(capture.HasOwnProp("gui"), "Recording must not create another GUI window.")
            AssertTrue(InStr(helper.waiting.Text, "Press"))
            AssertTrue(DllCall("GetWindowLongW", "Ptr", helper.waiting.Hwnd, "Int", -16, "UInt") & 0x10000000)
            AssertFalse(DllCall("GetWindowLongW", "Ptr", app.S_Gui.Hwnd, "Int", -16, "UInt") & 0x08000000, "Settings must remain interactive.")
            AssertEqual("F24", helper.field.Value)
            AssertEqual("F24", app.Exit_Program_Hotkey, "The waiting prompt must never enter saved settings.")
            capture.Pressed("LControl")
            capture.Pressed("F23")
            capture.Released("F23")
            AssertTrue(HotkeyCapture.Active, "Wait for the modifier release.")
            AssertEqual("^F23", helper.waiting.Text)
            AssertEqual("F24", helper.field.Value)
            capture.Released("LControl")
            AssertFalse(HotkeyCapture.Active)
            AssertEqual(0, HotkeyCapture.Current)
            AssertEqual(suspended, !!A_IsSuspended)
            AssertFalse(capture.hook.InProgress)
            AssertEqual("", capture.hook.OnKeyDown)
            AssertEqual("", capture.hook.OnKeyUp)
            AssertEqual(wasVisible, helper.field.Visible)
            AssertFalse(DllCall("GetWindowLongW", "Ptr", helper.waiting.Hwnd, "Int", -16, "UInt") & 0x10000000)
            AssertEqual(Chr(0x2328), helper.button.Text)
            AssertEqual("^F23", helper.field.Value)
            Sleep(30)
            AssertEqual("^F23", app.Exit_Program_Hotkey)
        }
    } finally {
        app.Dispose()
        Suspend(wasSuspended)
    }
}

TestRunner.Register("Inline recording cancels without changing the binding on button, Esc, outside click or close", InlineHotkeyCancelTest)
InlineHotkeyCancelTest() {
    app := HotkeySettingsFixture()
    try {
        helper := app.Find("Exit_Program_Hotkey")
        helper.field.Value := "F24"
        app.Exit_Program_Hotkey := "F24"
        for action in ["button", "escape", "outside", "close"] {
            app.RecordHotkey(helper)
            capture := HotkeyCapture.Current
            AssertTrue(HotkeyCapture.Active)
            if action = "button"
                app.RecordHotkey(helper)
            else if action = "escape"
                capture.Pressed("Escape")
            else if action = "outside"
                capture.ClickOutside(0, 0, 0x201, app.Find("Suspend_Hotkeys_Hotkey").field.Hwnd)
            else
                app.StopHotkeyHelpers()
            AssertTrue(capture.cancelled)
            AssertFalse(HotkeyCapture.Active)
            AssertFalse(capture.hook.InProgress)
            AssertEqual("F24", helper.field.Value)
            AssertEqual("F24", app.Exit_Program_Hotkey)
            AssertEqual(Chr(0x2328), helper.button.Text)
        }
    } finally {
        app.Dispose()
    }
}

TestRunner.Register("Inline list recording targets the selected line and preserves the other bindings", InlineHotkeyListTest)
InlineHotkeyListTest() {
    app := HotkeySettingsFixture()
    try {
        helper := app.Find("HotkeyList")
        app.MainFrame["HotkeyCharList"].Value := "Alice`n`nBob"
        helper.field.Value := "F24`n`nF22"
        lineStart := SendMessage(0xBB, 2, 0, helper.field) ; EM_LINEINDEX
        SendMessage(0xB1, lineStart, lineStart, helper.field) ; EM_SETSEL
        app.RecordHotkey(helper)
        capture := HotkeyCapture.Current
        AssertTrue(InStr(helper.waiting.Text, "Line 3:"))
        AssertEqual("F24`n`nF22", helper.field.Value)
        capture.Pressed("F23")
        capture.Released("F23")
        AssertEqual("F24`n`nF23", helper.field.Value)
        Sleep(30)
        AssertEqual("F24", app._Hotkeys["EVE - Alice"])
        AssertEqual("F23", app._Hotkeys["EVE - Bob"])
    } finally {
        app.Dispose()
    }
}

TestRunner.Register("Window-switch wildcard preparation preserves existing stars and empty binds", WindowHotkeyWildcardTextTest)
WindowHotkeyWildcardTextTest() {
    for pair in [["", ""], ["F23", "*F23"], ["^!F23", "*^!F23"], ["F23 Up", "*F23 Up"],
        ["*F23", "*F23"], ["**F23", "**F23"], ["~*^F23", "~*^F23"], ["^*F23", "^*F23"],
        ["XButton1 & F23", "XButton1 & F23"], ["Ctrl & F23", "Ctrl & F23"]]
        AssertEqual(pair[2], HotkeyBindings.ForWindowSwitch(pair[1]))
}

class WindowHotkeyRegistrationFixture extends ProfileSwitchFixture {
    retry := false
    retryKeys := []
    SwitchLangOnErr {
        get => this.retry
    }
    RegisterHotkeyWithLanguageRetry(key, callback, options, context) {
        this.retryKeys.Push(key)
        return super.RegisterHotkeyWithLanguageRetry(key, callback, options, context)
    }
}

TestRunner.Register("Window-switch hotkeys register wildcard key, mouse and custom combos with language retry on or off", WindowHotkeyWildcardRegistrationTest)
WindowHotkeyWildcardRegistrationTest() {
    app := WindowHotkeyRegistrationFixture()
    condition := (*) => false ; Never intercept the user's keys or activate EVE.
    registered := []
    HotIf(condition)
    try {
        for retry in [false, true] {
            app.retry := retry
            for raw in ["F13", "^F14", "*F15", "~*^F16", "F17 Up", "XButton1 & F18", "XButton2 & F19", "Ctrl & F20", "$F21"] {
                expected := HotkeyBindings.ForWindowSwitch(raw)
                registered.Push(expected)
                AssertTrue(app.RegisterWindowHotkey(raw, (*) => 0, "Window wildcard regression"))
                try Hotkey(expected, "On")
                catch as err
                    throw Error("Wildcard lookup for " raw ": " err.Message)
                if retry
                    AssertEqual(expected, app.retryKeys[-1])
            }
        }
        AssertEqual(9, app.retryKeys.Length)
        AssertEqual(0, app.layoutSwitches, "Valid wildcard syntax must not require a keyboard layout retry.")
        app.retry := false
        AssertThrows(() => app.RegisterWindowHotkey("NotARealKeyName", (*) => 0, "Invalid wildcard regression"), "Invalid key name")
        app.retry := true
        AssertFalse(app.RegisterWindowHotkey("NotARealKeyName", (*) => 0, "Invalid wildcard regression"))
        AssertEqual("*NotARealKeyName", app.retryKeys[-1])
        AssertEqual(1, app.layoutSwitches)
    } finally {
        for key in registered
            try Hotkey(key, "Off")
        HotIf()
    }
}

class WindowHotkeyPathsFixture extends CombatEventsFixture {
    requested := []
    RegisterWindowHotkey(key, callback, context) {
        this.requested.Push(key)
        return true
    }
}

TestRunner.Register("Character, EVE groups, Quick Group and app activation share wildcard registration", WindowHotkeyWildcardPathsTest)
WindowHotkeyWildcardPathsTest() {
    app := WindowHotkeyPathsFixture()
    app.ProfileOverride()
    app.debugToolTipMethod := (*) => 0
    app.debugToolTipDelay := -1000
    app._Hotkeys := [Map("EVE - Wildcard Test", "F13")]
    app.Profiles["Default"]["Hotkey Groups"] := Map("Fleet", Map("Characters", ["EVE - Wildcard Test"],
        "ForwardsHotkey", "F14", "BackwardsHotkey", "*F15", "FirstCharHotkey", "F16"))
    app.QuickGroupEnabled := true
    app.QuickGroupHotkey := "F17"
    app.NonEVEGroups := Map("Tools", Map("exe", ["wildcard-test.exe"], "title", [""], "fkey", "F18", "bkey", "F19"))
    app.NonEVEHotkeys := Map("exe", ["wildcard-test.exe"], "title", [""], "hotkey", ["F20"])
    before := JSON.Dump(app._JSON)
    try {
        app.RegisterHotkeys("EVE - Wildcard Test")
        app.Register_Hotkey_Groups()
        app.RegisterQuickGroup()
        app.RegisterNonEVEGroups()
        app.RegisterNonEVEHotkeys()
        AssertEqual(8, app.requested.Length)
        for key in ["F13", "F14", "*F15", "F16", "F17", "F18", "F19", "F20"] {
            found := false
            for requested in app.requested
                found := found || requested = key
            AssertTrue(found, "Missing window registration path for " key)
        }
        AssertEqual(before, JSON.Dump(app._JSON), "Registration must not rewrite saved or displayed binds.")
    } finally {
        SetTimer(app.debugToolTipMethod, 0)
        HotIf()
    }
}

TestRunner.Register("Duplicate warnings distinguish modified shortcuts from automatic window wildcard fallbacks", WindowHotkeyWildcardConflictsTest)
WindowHotkeyWildcardConflictsTest() {
    app := CombatEventsFixture()
    app.ProfileOverride()
    app.SwitchToPreviousWindow_Hotkey := "F23"
    app.Exit_Program_Hotkey := "^F23"
    app.ProfileActivationHotkey := "F24"
    entries := HotkeyBindings.Collect(app)
    conflicts := HotkeyBindings.Conflicts(entries)
    AssertFalse(conflicts.Has("setting:SwitchToPreviousWindow_Hotkey"))
    AssertFalse(conflicts.Has("setting:Exit_Program_Hotkey"))
    for entry in entries {
        if entry.id = "setting:SwitchToPreviousWindow_Hotkey" {
            AssertEqual("F23", entry.key)
            AssertTrue(entry.parsed.wildcard)
        } else if entry.id = "setting:Exit_Program_Hotkey" || entry.id = "profile:Default"
            AssertFalse(entry.parsed.wildcard, "Non-window actions must retain their entered modifier semantics.")
    }
    AssertEqual("F23", app.SwitchToPreviousWindow_Hotkey)
    AssertEqual("^F23", app.Exit_Program_Hotkey)
}
