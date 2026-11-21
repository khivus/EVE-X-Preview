TestRunner.Register("CharacterList settings validation inheritance and migration", CharacterListSettingsTest)
CharacterListSettingsTest() {
    config := JSON.Load(default_JSON)
    config["CharacterListEnabled"] := 1
    model := SettingsModel(config)
    AssertTrue(model.Get("CharacterListEnabled"))
    AssertTrue(!model.config.Has("CharacterListEnabled"))
    model.Set("CharacterListColumns", 3)
    model.Set("CharacterListColumnWidth", 50)
    model.Set("CharacterListDirection", "horizontal")
    AssertEqual(50, model.Get("CharacterListColumnWidth"))
    AssertEqual(3, model.Merge(config)["_Profiles"]["Default"]["Character List"]["CharacterListColumns"])
    model.AddProfile("Second")
    model.SetInheritance("character-list", true)
    AssertEqual("Default", model.Source("character-list"))
    AssertThrows(() => model.Set("CharacterListColumns", 2))
    model.SetInheritance("character-list", false)
    model.Set("CharacterListColumns", 2)
    AssertEqual(3, model.GetFor("CharacterListColumns", "Default"))
    AssertEqual(2, model.Get("CharacterListColumns"))
    AssertThrows(() => model.Set("CharacterListColumnWidth", 49))
    AssertThrows(() => model.Set("CharacterListOpacity", 101))
    AssertThrows(() => model.Set("CharacterListSort", "invalid"))
    model.config["LastUsedProfile"] := "Second"
    owner := {_JSON: model.config, Save_Settings_Delay_Timer: (*) => 0}
    list := {owner: owner}
    ObjSetBase(list, CharacterList.Prototype)
    model.SetInheritance("character-list", true)
    AssertEqual(3, list.Get("Columns", 0))
    list.Put("AlwaysOnTop", 0)
    AssertEqual(0, list.Get("AlwaysOnTop", 1))
    AssertEqual(1, model.GetFor("CharacterListAlwaysOnTop", "Default"))
    SetTimer(owner.Save_Settings_Delay_Timer, 0)
}
TestRunner.Register("CharacterList sorting respects selected order without grouping", CharacterListSortTest)
CharacterListSortTest() {
    owner := {_JSON: JSON.Load(default_JSON)}
    list := {owner: owner}
    ObjSetBase(list, CharacterList.Prototype)
    a := {name: "Alice", key: "Alice", hwnd: 200, created: 10, loggedOut: false, system: "Jita"}
    b := {name: "Bob", key: "Bob", hwnd: 100, created: 20, loggedOut: false, system: "Amarr"}
    logout := {name: "EVE", key: "hwnd:300", hwnd: 300, created: 30, loggedOut: true, system: ""}
    AssertTrue(list.Compare(a, b, "name", []) < 0)
    AssertTrue(list.Compare(a, b, "descending", []) > 0)
    AssertTrue(list.Compare(a, b, "oldest", []) < 0)
    AssertTrue(list.Compare(a, b, "newest", []) > 0)
    AssertTrue(list.Compare(a, b, "manual", ["Bob", "Alice"]) > 0)
    AssertTrue(list.Compare(logout, a, "name", []) > 0)
    AssertTrue(!SettingsSchema.Fields.Has("CharacterListGroupSystems"))
    AssertTrue(!SettingsSchema.Fields.Has("CharacterListGroupLoggedOut"))
}
TestRunner.Register("CharacterList persistent borderless window refreshes without replacement", CharacterListWindowTest)
CharacterListWindowTest() {
    owner := {_JSON: JSON.Load(default_JSON), EVEExe: "ahk_exe CharacterListMissingFixture.exe", Save_Settings_Delay_Timer: (*) => 0}
    owner._JSON["_Profiles"]["Default"]["Character List"]["CharacterListEnabled"] := 1
    list := CharacterList(owner)
    try {
        original := list.gui.Hwnd
        popup := list.BuildMenu()
        AssertTrue(!!popup.Handle)
        AssertTrue(!(WinGetStyle(original) & 0xC00000))
        list.items := [{name: "Alice", key: "Alice", hwnd: 0, loggedOut: false, system: "Jita"}, {name: "EVE", key: "hwnd:5", hwnd: 5, loggedOut: true, system: ""}]
        list.Put("ShowHwnd", 1)
        list.Render()
        AssertEqual(2, list.cells.Count)
        AssertEqual("Alice", list.pool[1].Text)
        AssertTrue(InStr(list.pool[2].Text, "0x5"))
        firstCell := list.pool[1].Hwnd
        list.Render()
        AssertEqual(original, list.gui.Hwnd)
        AssertEqual(firstCell, list.pool[1].Hwnd)
        list.Collapse()
        AssertTrue(list.Get("Collapsed", false))
        AssertEqual(original, list.gui.Hwnd)
        list.Collapse()
        list.Put("Columns", 3)
        list.Refresh()
        list.Put("Enabled", 0)
        list.Refresh()
        AssertTrue(!DllCall("IsWindowVisible", "Ptr", list.gui.Hwnd))
    } finally {
        SetTimer(list.tick, 0), SetTimer(owner.Save_Settings_Delay_Timer, 0)
        OnMessage(0x2B, list.systemDrawCallback, 0)
        OnMessage(0x20, list.cursorCallback, 0)
        OnMessage(0x83, list.ncCallback, 0), OnMessage(0x84, list.ncCallback, 0)
        OnMessage(0x201, list.mouse, 0), OnMessage(0x20A, list.wheelCallback, 0), OnMessage(0x232, list.endMove, 0)
        if list.gui
            list.gui.Destroy()
    }
}
TestRunner.Register("CharacterList settings navigation follows placement with unique section icons", CharacterListNavigationTest)
CharacterListNavigationTest() {
    previous := "", icons := Map()
    for section in SettingsSchema.Sections {
        if section.id = "character-list"
            AssertEqual("placement", previous)
        icon := SettingsSchema.Icons[section.id]
        AssertTrue(!icons.Has(icon))
        icons[icon] := true
        previous := section.id
    }
    AssertEqual("appearance", SettingsSchema.Sections[1].id)
    AssertTrue(!SettingsSchema.Fields["CharacterListEnabled"].global)
}

TestRunner.Register("CharacterList manual lanes insertion and directional scrolling", CharacterListLayoutTest)
CharacterListLayoutTest() {
    owner := {_JSON: JSON.Load(default_JSON), EVEExe: "ahk_exe CharacterListMissingFixture.exe", Save_Settings_Delay_Timer: (*) => 0}
    owner._JSON["_Profiles"]["Default"]["Character List"]["CharacterListEnabled"] := 1
    list := CharacterList(owner)
    try {
        list.items := []
        Loop 12
            list.items.Push({name: "Pilot " A_Index, key: "pilot" A_Index, hwnd: A_Index+1000, loggedOut: false})
        list.Put("Columns", 2), list.Put("AutoFill", 0)
        layout := list.BuildLayout(list.items, 4)
        for slot in layout
            AssertEqual(0, slot.lane)
        list.Render()
        AssertTrue(list.maxOffset > 0)
        AssertTrue(list.track.Visible)
        list.ApplyDrop(3, {lane: 1, before: ""})
        AssertEqual(1, list.Get("Lanes", Map())["pilot3"])
        AssertEqual("pilot3", list.items[12].key)
        list.Render()
        list.ApplyDrop(12, {lane: 0, before: "pilot2"})
        AssertEqual("pilot1", list.items[1].key)
        AssertEqual("pilot3", list.items[2].key)
        AssertEqual("pilot2", list.items[3].key)
        AssertEqual(12, list.items.Length)
        list.Put("Direction", "horizontal")
        list.Render()
        list.track.GetPos(,, &width, &height)
        AssertTrue(width > height)
        list.Put("AutoFill", 1)
        AssertEqual(0, list.Get("Lanes", Map()).Count)
        rect := Buffer(16), client := Buffer(16)
        DllCall("GetWindowRect", "Ptr", list.gui.Hwnd, "Ptr", rect)
        DllCall("GetClientRect", "Ptr", list.gui.Hwnd, "Ptr", client)
        AssertEqual(NumGet(rect, 12, "Int")-NumGet(rect, 4, "Int"), NumGet(client, 12, "Int"))
    } finally {
        SetTimer(list.tick, 0), SetTimer(owner.Save_Settings_Delay_Timer, 0)
        OnMessage(0x2B, list.systemDrawCallback, 0)
        OnMessage(0x20, list.cursorCallback, 0)
        OnMessage(0x83, list.ncCallback, 0), OnMessage(0x84, list.ncCallback, 0)
        OnMessage(0x201, list.mouse, 0), OnMessage(0x20A, list.wheelCallback, 0), OnMessage(0x232, list.endMove, 0)
        list.gui.Destroy()
    }
}
TestRunner.Register("CharacterList events reuse event text colors suppression and duration", CharacterListEventsTest)
CharacterListEventsTest() {
    owner := {_JSON: JSON.Load(default_JSON), EVEExe: "ahk_exe CharacterListMissingFixture.exe", eventDisplayDuration: 1000, supressForFocused: false, showEventText: true, flashBorderEnabled: true, flashBorderInterval: 500, monitoredEventsTexts: Map("test", "Alert"), monitoredEvents: Map("test", Map("color", "FF0000"))}
    owner._JSON["_Profiles"]["Default"]["Character List"]["CharacterListEnabled"] := 1
    list := {owner: owner, events: Map()}
    ObjSetBase(list, CharacterList.Prototype)
    list.EventStart("test", 123)
    AssertEqual("Alert", list.EventState(123)[1])
    AssertEqual("FF0000", StrUpper(list.EventState(123)[2]))
    owner.showEventText := false, owner.flashBorderEnabled := false
    AssertEqual("", list.EventState(123)[1])
    AssertEqual("", list.EventState(123)[2])
    list.events[123].expires := A_TickCount-1
    AssertEqual("", list.EventState(123)[1])
    AssertEqual(0, list.events.Count)
    list.EventStart("test", 123), list.EventEnd(123)
    AssertEqual(0, list.events.Count)
}

TestRunner.Register("CharacterList nonblocking drag release inserts and leaves no native frame", CharacterListDragReleaseTest)
CharacterListDragReleaseTest() {
    owner := {_JSON: JSON.Load(default_JSON), EVEExe: "ahk_exe CharacterListMissingFixture.exe", Save_Settings_Delay_Timer: (*) => 0}
    owner._JSON["_Profiles"]["Default"]["Character List"]["CharacterListEnabled"] := 1
    list := CharacterList(owner)
    try {
        list.items := [{key: "one", name: "One", hwnd: 1, loggedOut: false}, {key: "two", name: "Two", hwnd: 2, loggedOut: false}, {key: "three", name: "Three", hwnd: 3, loggedOut: false}]
        list.Render()
        list.gui.GetClientPos(&gx, &gy)
        list.BeginDrag("row", gx+20, gy+48, 1)
        SetTimer(list.dragTimer, 0)
        AssertTrue(list.dragging)
        AssertEqual(list.gui.Hwnd, DllCall("GetCapture", "Ptr"))
        list.UpdateDrag(gx+20, gy+102, true)
        AssertTrue(list.insertion.Visible)
        list.UpdateDrag(gx+20, gy+102, false)
        AssertTrue(!list.dragging)
        AssertTrue(!list.insertion.Visible)
        AssertEqual("two", list.items[1].key)
        AssertEqual("one", list.items[2].key)
        AssertEqual("three", list.items[3].key)
        AssertTrue(DllCall("GetCapture", "Ptr") != list.gui.Hwnd)
        for active in [0, 1, 0] {
            SendMessage(0x86, active, 0,, list.gui.Hwnd)
            AssertEqual(0, WinGetStyle(list.gui.Hwnd) & 0xC40000)
        }
        list.gui.GetPos(&wx, &wy, &width, &height)
        list.BeginDrag("resize", wx+width-1, wy+60, 0, "R")
        SetTimer(list.dragTimer, 0)
        list.UpdateDrag(wx+width+19, wy+60, false)
        list.gui.GetPos(,, &resized)
        AssertEqual(width+20, resized)
        AssertEqual(0, WinGetStyle(list.gui.Hwnd) & 0xC40000)
        list.BeginDrag("row", gx+20, gy+48, 1)
        SetTimer(list.dragTimer, 0)
        list.FinishDrag(true)
        AssertEqual("two", list.items[1].key)
    } finally {
        list.FinishDrag(true)
        SetTimer(list.tick, 0), SetTimer(list.dragTimer, 0), SetTimer(owner.Save_Settings_Delay_Timer, 0)
        OnMessage(0x2B, list.systemDrawCallback, 0)
        OnMessage(0x20, list.cursorCallback, 0)
        OnMessage(0x83, list.ncCallback, 0), OnMessage(0x84, list.ncCallback, 0)
        OnMessage(0x201, list.mouse, 0), OnMessage(0x20A, list.wheelCallback, 0), OnMessage(0x232, list.endMove, 0)
        list.gui.Destroy()
    }
}

TestRunner.Register("CharacterList fits contents and right aligns existing system labels", CharacterListSystemLayoutTest)
CharacterListSystemLayoutTest() {
    owner := {_JSON: JSON.Load(default_JSON), EVEExe: "ahk_exe CharacterListMissingFixture.exe", Save_Settings_Delay_Timer: (*) => 0, systemTrackingEnabled: true, systemLogMonitor: {readers: Map("EVE - Alice", {system: "Jita"})}}
    owner._JSON["_Profiles"]["Default"]["Character List"]["CharacterListEnabled"] := 1
    list := CharacterList(owner)
    try {
        list.items := [{name: "Alice", title: "EVE - Alice", key: "Alice", hwnd: 1000, loggedOut: false}]
        list.Render()
        AssertEqual("Jita", list.systemPool[1].Text)
        AssertEqual(13, WinGetStyle(list.systemPool[1].Hwnd) & 15)
        list.pool[1].GetPos(&nx,, &nw)
        list.systemPool[1].GetPos(&sx,, &sw)
        AssertTrue(nx+nw <= sx)
        AssertEqual(8+list.Get("ColumnWidth", 220)-8, sx+sw)
        list.gui.GetPos(,,, &height)
        AssertEqual(76, height)
        Loop 3 {
            list.Render()
            list.gui.GetPos(,,, &height)
            AssertEqual(76, height)
        }
        AssertEqual(340, list.Get("Height", 0))
        owner.systemLogMonitor.readers["EVE - Alice"].system := "Amarr"
        list.Render()
        AssertEqual("Amarr", list.systemPool[1].Text)
        owner.systemLogMonitor.readers["EVE - Alice"].system := "Tekaima"
        list.Render()
        AssertEqual("Tekaima", list.systemPaint[list.systemPool[1].Hwnd].text)
        owner.systemLogMonitor.readers["EVE - Alice"].system := "Jita"
        list.Render()
        AssertEqual("Jita", list.systemPaint[list.systemPool[1].Hwnd].text)
        owner.systemTrackingEnabled := false
        list.Render()
        AssertTrue(!list.systemPool[1].Visible)
        list.Put("Direction", "horizontal"), list.Put("Columns", 2)
        list.Render()
        list.gui.GetPos(,,, &height)
        AssertEqual(112, height)
    } finally {
        list.FinishDrag(true)
        SetTimer(list.tick, 0), SetTimer(list.dragTimer, 0), SetTimer(owner.Save_Settings_Delay_Timer, 0)
        OnMessage(0x2B, list.systemDrawCallback, 0)
        OnMessage(0x20, list.cursorCallback, 0)
        OnMessage(0x83, list.ncCallback, 0), OnMessage(0x84, list.ncCallback, 0)
        OnMessage(0x201, list.mouse, 0), OnMessage(0x20A, list.wheelCallback, 0), OnMessage(0x232, list.endMove, 0)
        list.gui.Destroy()
    }
}
TestRunner.Register("CharacterList system label repaint removes old glyph pixels", CharacterListSystemPaintTest)
CharacterListSystemPaintTest() {
    owner := {_JSON: JSON.Load(default_JSON), EVEExe: "ahk_exe CharacterListMissingFixture.exe", Save_Settings_Delay_Timer: (*) => 0}
    owner._JSON["_Profiles"]["Default"]["Character List"]["CharacterListEnabled"] := 1
    list := CharacterList(owner)
    dc := DllCall("CreateCompatibleDC", "Ptr", 0, "Ptr")
    info := Buffer(40, 0), pixels := 0
    NumPut("UInt", 40, "Int", 140, "Int", -28, "UShort", 1, "UShort", 32, info)
    bitmap := DllCall("CreateDIBSection", "Ptr", dc, "Ptr", info, "UInt", 0, "Ptr*", &pixels, "Ptr", 0, "UInt", 0, "Ptr")
    previous := DllCall("SelectObject", "Ptr", dc, "Ptr", bitmap, "Ptr")
    try {
        list.items := [{key: "test", name: "Pilot", hwnd: 1000, loggedOut: false}]
        list.Render()
        control := list.systemPool[1].Hwnd
        draw := Buffer(A_PtrSize = 8 ? 64 : 48, 0)
        NumPut("Ptr", control, draw, A_PtrSize = 8 ? 24 : 20)
        NumPut("Ptr", dc, draw, A_PtrSize = 8 ? 32 : 24)
        rectOffset := A_PtrSize = 8 ? 40 : 28
        NumPut("Int", 140, "Int", 28, draw, rectOffset+8)
        list.systemPaint[control] := {text: "Jita", background: "202124", foreground: "ECEEF2"}
        list.DrawSystemLabel(0, draw.Ptr, 0x2B, list.gui.Hwnd)
        DllCall("GdiFlush")
        expected := Buffer(140*28*4)
        DllCall("RtlMoveMemory", "Ptr", expected, "Ptr", pixels, "UPtr", expected.Size)
        list.systemPaint[control].text := "Tekaima longer name"
        list.DrawSystemLabel(0, draw.Ptr, 0x2B, list.gui.Hwnd)
        list.systemPaint[control].text := "Jita"
        list.DrawSystemLabel(0, draw.Ptr, 0x2B, list.gui.Hwnd)
        DllCall("GdiFlush")
        AssertEqual(0, DllCall("msvcrt\memcmp", "Ptr", expected, "Ptr", pixels, "UPtr", expected.Size, "Cdecl Int"))
    } finally {
        DllCall("SelectObject", "Ptr", dc, "Ptr", previous)
        DllCall("DeleteObject", "Ptr", bitmap), DllCall("DeleteDC", "Ptr", dc)
        SetTimer(list.tick, 0), SetTimer(owner.Save_Settings_Delay_Timer, 0)
        OnMessage(0x2B, list.systemDrawCallback, 0), OnMessage(0x20, list.cursorCallback, 0)
        OnMessage(0x83, list.ncCallback, 0), OnMessage(0x84, list.ncCallback, 0)
        OnMessage(0x201, list.mouse, 0), OnMessage(0x20A, list.wheelCallback, 0), OnMessage(0x232, list.endMove, 0)
        list.gui.Destroy()
    }
}
