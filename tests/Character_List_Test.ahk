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
