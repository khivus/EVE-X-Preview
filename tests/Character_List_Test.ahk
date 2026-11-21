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
TestRunner.Register("CharacterList sorting groups logged out clients and solar systems", CharacterListSortTest)
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
    AssertTrue(list.Compare(logout, a, "descending", []) < 0)
    list.Data["CharacterListGroupSystems"] := 1
    AssertTrue(list.Compare(a, b, "name", []) > 0)
    AssertTrue(list.Compare(a, b, "manual", ["Alice", "Bob"]) < 0)
}
TestRunner.Register("CharacterList persistent borderless window refreshes without replacement", CharacterListWindowTest)
CharacterListWindowTest() {
    owner := {_JSON: JSON.Load(default_JSON), EVEExe: "ahk_exe CharacterListMissingFixture.exe", Save_Settings_Delay_Timer: (*) => 0}
    owner._JSON["_Profiles"]["Default"]["Character List"]["CharacterListEnabled"] := 1
    list := CharacterList(owner)
    try {
        original := list.gui.Hwnd
        AssertTrue(!(WinGetStyle(original) & 0xC00000))
        list.items := [{name: "Alice", key: "Alice", hwnd: 0, loggedOut: false, system: "Jita"}, {name: "EVE", key: "hwnd:5", hwnd: 5, loggedOut: true, system: ""}]
        list.Put("ShowHwnd", 1)
        list.Put("GroupSystems", 1)
        list.Render()
        AssertEqual(2, list.cells.Count)
        AssertTrue(InStr(list.pool[1].Text, "Jita"))
        AssertTrue(InStr(list.pool[2].Text, "0x5"))
        firstCell := list.pool[1].Hwnd
        list.Render()
        AssertEqual(original, list.gui.Hwnd)
        AssertEqual(firstCell, list.pool[1].Hwnd)
        list.Put("GroupSystems", 0)
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
