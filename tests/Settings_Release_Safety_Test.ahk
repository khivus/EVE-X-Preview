TestRunner.Register("Settings release safety modified chords coexist with wildcard fallbacks and duplicate chords still warn", SettingsSafeModifierWarningsTest)
SettingsSafeModifierWarningsTest() {
    for pair in [["F23", "!F23"], ["F23", "^F23"], ["^F23", "^!F23"], ["F23", "+#F23"]] {
        entries := []
        HotkeyBindings.Add(entries, "one", "One", pair[1], true)
        HotkeyBindings.Add(entries, "two", "Two", pair[2], true)
        AssertEqual(0, HotkeyBindings.Conflicts(entries).Count)
    }
    for pair in [["^!F23", "!^f23"], ["F23", "*F23"], ["*^F23", "*!F23"], ["<^F23", "^F23"], ["Ctrl & F23", "^F23"]] {
        entries := []
        HotkeyBindings.Add(entries, "one", "One", pair[1])
        HotkeyBindings.Add(entries, "two", "Two", pair[2])
        AssertEqual(2, HotkeyBindings.Conflicts(entries).Count)
    }
    app := SettingsWorkbenchFixture()
    try {
        tick := Chr(96)
        app.model.Set("Login_Screen_Cycle_Hotkey", "!" tick)
        app.model.AddGroup("characters", "Fleet")
        app.model.Groups("characters")["Fleet"]["ForwardsHotkey"] := tick
        for automatic in [true, false] {
            app.model.Set("AutoWildcardWindowHotkeys", automatic)
            app.Changed()
            conflicts := HotkeyBindings.Conflicts(HotkeyBindings.Collect(app))
            AssertFalse(conflicts.Has("setting:Login_Screen_Cycle_Hotkey"))
            AssertFalse(conflicts.Has("group:Fleet:ForwardsHotkey"))
        }
        app.model.Groups("characters")["Fleet"]["ForwardsHotkey"] := "!" tick
        app.Changed()
        conflicts := HotkeyBindings.Conflicts(HotkeyBindings.Collect(app))
        AssertTrue(conflicts.Has("setting:Login_Screen_Cycle_Hotkey"))
        AssertTrue(conflicts.Has("group:Fleet:ForwardsHotkey"))
    } finally app.Dispose()
}

; Place this observer below the chrome subclass to inspect the redraw state
; at the moment the original common-control procedure receives an update.
class SettingsScrollUpdateObserver {
    __New(app, control) {
        this.app := app, this.handle := control.Hwnd, this.redraw := true, this.protected := [], this.transitions := [], this.erases := []
        this.callback := CallbackCreate(ObjBindMethod(this, "WindowProc"),, 6)
        chrome := app.theme.chrome
        DllCall("comctl32\RemoveWindowSubclass", "Ptr", this.handle, "Ptr", chrome.callback, "UPtr", 1)
        DllCall("comctl32\SetWindowSubclass", "Ptr", this.handle, "Ptr", this.callback, "UPtr", 96, "UPtr", 0)
        DllCall("comctl32\SetWindowSubclass", "Ptr", this.handle, "Ptr", chrome.callback, "UPtr", 1, "UPtr", 0)
    }
    WindowProc(handle, message, wParam, lParam, subclass, data) {
        if message = 0xB {
            this.redraw := !!wParam
            this.transitions.Push(this.redraw)
        }
        if message = 0x115 || message = 0x114 || message = 0x20A || message = 0x1014 || message = 0x1102 || message = 0x1114 || message = 0x201
            this.protected.Push(!this.redraw)
        if message = 0x14 {
            item := this.app.theme.chrome.windows[handle]
            this.erases.Push([this.redraw, item.HasOwnProp("scrollUpdating") && item.scrollUpdating, DllCall("GetObjectType", "Ptr", wParam, "UInt")])
        }
        return DllCall("comctl32\DefSubclassProc", "Ptr", handle, "UInt", message, "Ptr", wParam, "Ptr", lParam, "Ptr")
    }
    AssertProtected() {
        AssertTrue(this.protected.Length > 0)
        for state in this.protected
            AssertTrue(state, "Native range changes must run with intermediate drawing disabled.")
        AssertTrue(this.redraw, "Redraw must be restored at the end of every operation.")
        AssertEqual(0, Mod(this.transitions.Length, 2))
        AssertTrue(DllCall("IsWindowVisible", "Ptr", this.handle))
    }
    Dispose() {
        if DllCall("IsWindow", "Ptr", this.handle)
            DllCall("comctl32\RemoveWindowSubclass", "Ptr", this.handle, "Ptr", this.callback, "UPtr", 96)
        CallbackFree(this.callback)
    }
}

TestRunner.Register("Settings release safety event table scrolling suppresses native light draws and retains its header", SettingsSafeEventScrollTest)
SettingsSafeEventScrollTest() {
    app := SettingsWorkbenchFixture(), observer := 0, paints := 0
    try {
        app.model.Set("SettingsTheme", "dark")
        app.Changed()
        app.Navigate("events")
        app.S_Gui.Show("NoActivate x-3000 y-3000 w1000 h740")
        Sleep(180)
        list := app.eventTable.list, grid := app.eventTable.grid
        list.Move(,,, 190)
        list.ModifyCol(5, 1200)
        grid.Select(1, 4), grid.Select(3, 4, true)
        selected := JSON.Dump(grid.selected)
        DllCall("RedrawWindow", "Ptr", list.Hwnd, "Ptr", 0, "Ptr", 0, "UInt", 0x585)
        observer := SettingsScrollUpdateObserver(app, list)
        header := SendMessage(0x101F, 0, 0, list)
        paints := SettingsPaintObserver([header, list.Hwnd])
        point := SettingsChromeStartDrag(app, list.Hwnd, 1)
        point[2] += Round(app.theme.chrome.scrollbars.drag.travel * 0.6)
        SettingsChromeMoveDrag(list.Hwnd, point)
        AssertTrue(SendMessage(0x1027, 0, 0, list) > 0)
        AssertEqual(selected, JSON.Dump(grid.selected))
        AssertEqual(SettingsTheme.BGR("25282D"), SettingsChromeTrackPixel(app, list.Hwnd, 1))
        SettingsChromeMoveDrag(list.Hwnd, point, true)
        observer.AssertProtected()
        AssertEqual(0, paints.counts[header].paints, "Vertical scrolling must leave the header's drawing alone.")
        for erase in observer.erases
            AssertEqual(10, erase[3], "Any native background clearing must stay inside the table's memory buffer.")
        SendMessage(0x1014, 90, 0, list)
        AssertTrue(DllCall("GetScrollPos", "Ptr", list.Hwnd, "Int", 0, "Int") > 0)
        AssertEqual("Event", list.GetText(0, 2))
    } finally {
        if IsObject(paints)
            paints.Dispose()
        if IsObject(observer)
            observer.Dispose()
        app.Dispose()
    }
}

TestRunner.Register("Settings release safety owned profile table scrolling preserves editable inheritance", SettingsSafeOwnedScrollTest)
SettingsSafeOwnedScrollTest() {
    app := SettingsWorkbenchFixture(), observer := 0
    try {
        app.model.Set("SettingsTheme", "dark")
        app.Changed()
        app.AskName("Create profile", "", ObjBindMethod(app, "CreateProfile"), true)
        dialog := app.activeDialog, grid := dialog.profileCreation.grid
        dialog.root.Show("NoActivate x-3000 y-3000 w640 h634")
        observer := SettingsScrollUpdateObserver(app, grid.list)
        point := SettingsChromeStartDrag(app, grid.hwnd, 1)
        point[2] += Round(app.theme.chrome.scrollbars.drag.travel * 0.75)
        SettingsChromeMoveDrag(grid.hwnd, point)
        SettingsChromeMoveDrag(grid.hwnd, point, true)
        AssertTrue(SendMessage(0x1027, 0, 0, grid.list) > 0)
        observer.AssertProtected()
        AssertEqual(SettingsTheme.BGR("25282D"), SettingsChromeTrackPixel(app, grid.hwnd, 1))
        AssertTrue(grid.SetCells([[SettingsModel.Groups.Length, 2, 1]]))
        AssertEqual(1, grid.rows[-1][2])
        observer.Dispose(), observer := 0
        app.CloseChild(dialog)
        AssertFalse(WinGetStyle(app.S_Gui.Hwnd) & 0x08000000)
    } finally {
        if IsObject(observer)
            observer.Dispose()
        app.Dispose()
    }
}

TestRunner.Register("Settings release safety sidebar expansion wheel and ensure-visible retain dark scrollbars", SettingsSafeSidebarScrollTest)
SettingsSafeSidebarScrollTest() {
    app := SettingsWorkbenchFixture(), observer := 0
    try {
        app.model.Set("SettingsTheme", "dark")
        app.Changed()
        app.S_Gui.Show("NoActivate x-3000 y-3000 w800 h500")
        app.width := 800, app.height := 500
        app.Layout()
        Sleep(180)
        observer := SettingsScrollUpdateObserver(app, app.nav)
        AssertTrue(SendMessage(0x112D, 0, 0, app.nav) & 4, "Dark navigation uses its own native paint buffer.")
        for root in app.navRoots
            app.nav.Modify(root, "Expand")
        last := app.navRoots[-1]
        app.nav.Modify(last, "-Expand")
        SendMessage(0x1114, 0, last, app.nav)
        point := SettingsSafeTreeButtonPoint(app.nav, last)
        SendMessage(0x201, 1, point, app.nav)
        SendMessage(0x202, 0, point, app.nav)
        Sleep(20)
        AssertTrue(app.expandedPages[app.navPages[last]], "The actual bottom-group expander click must still work.")
        SendMessage(0x1114, 0, app.navSections["about"], app.nav)
        SendMessage(0x20A, (-120 & 0xFFFF) << 16, 0, app.nav)
        observer.AssertProtected()
        AssertTrue(IsObject(app.theme.chrome.ScrollInfo(app.nav.Hwnd, -5)))
        AssertEqual(SettingsTheme.BGR("25282D"), SettingsChromeTrackPixel(app, app.nav.Hwnd, 1))
        AssertEqual("about", app.navItems[app.navSections["about"]])
    } finally {
        if IsObject(observer)
            observer.Dispose()
        app.Dispose()
    }
}

SettingsSafeTreeButtonPoint(tree, item) {
    rect := Buffer(16, 0)
    NumPut("Ptr", item, rect)
    AssertTrue(SendMessage(0x1104, 1, rect.Ptr, tree))
    y := (NumGet(rect, 4, "Int") + NumGet(rect, 12, "Int")) // 2
    loop Max(1, NumGet(rect, 0, "Int")) {
        hit := Buffer(A_PtrSize = 8 ? 24 : 16, 0)
        NumPut("Int", A_Index - 1, "Int", y, hit)
        SendMessage(0x1111, 0, hit.Ptr, tree)
        if NumGet(hit, 8, "UInt") & 0x10
            return (A_Index - 1) | ((y & 0xFFFF) << 16)
    }
    throw Error("No visible tree expander hit area.")
}

TestRunner.Register("Settings release safety light and externally disabled redraw paths remain untouched", SettingsSafeScrollRestorationTest)
SettingsSafeScrollRestorationTest() {
    app := SettingsWorkbenchFixture(), observer := 0
    try {
        app.model.Set("SettingsTheme", "light")
        app.Changed()
        app.Navigate("events")
        app.S_Gui.Show("NoActivate x-3000 y-3000 w1000 h740")
        Sleep(180)
        list := app.eventTable.list
        observer := SettingsScrollUpdateObserver(app, list)
        SendMessage(0x1014, 0, 30, list)
        AssertEqual(0, observer.transitions.Length, "Light mode must continue using native drawing.")
        app.model.Set("SettingsTheme", "dark")
        app.Changed()
        list.Opt("-Redraw")
        observer.transitions := []
        SendMessage(0x1014, 0, 30, list)
        AssertFalse(observer.redraw)
        AssertEqual(0, observer.transitions.Length, "A caller's redraw lock must not be re-enabled by the scroll guard.")
        list.Opt("+Redraw")
        AssertTrue(observer.redraw)
    } finally {
        if IsObject(observer)
            observer.Dispose()
        app.Dispose()
    }
}
