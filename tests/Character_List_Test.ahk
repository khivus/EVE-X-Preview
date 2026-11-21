TestRunner.Register("CharacterList settings validate and persist independently", CharacterListSettingsTest)
CharacterListSettingsTest() {
    model := SettingsModel(JSON.Load(default_JSON))
    model.Set("CharacterListEnabled", true)
    model.Set("CharacterListColumns", 3)
    model.Set("CharacterListColumnWidth", 250)
    model.Set("CharacterListDirection", "horizontal")
    model.Set("CharacterListSort", "manual")
    merged := model.Merge(JSON.Load(default_JSON))
    AssertEqual(3, merged["CharacterListColumns"])
    AssertEqual("horizontal", merged["CharacterListDirection"])
    AssertThrows(() => model.Set("CharacterListColumns", 0))
    AssertThrows(() => model.Set("CharacterListSort", "invalid"))
}
TestRunner.Register("CharacterList sorting respects names custom order and process age", CharacterListSortTest)
CharacterListSortTest() {
    list := CharacterList.Prototype
    a := {name: "Alice", hwnd: 200, created: 10}, b := {name: "Bob", hwnd: 100, created: 20}
    AssertTrue(list.Compare(a, b, "name", []) < 0)
    AssertTrue(list.Compare(a, b, "descending", []) > 0)
    AssertTrue(list.Compare(a, b, "oldest", []) < 0)
    AssertTrue(list.Compare(a, b, "newest", []) > 0)
    AssertTrue(list.Compare(a, b, "manual", ["Bob", "Alice"]) > 0)
}
TestRunner.Register("CharacterList window builds collapses and disables without clients", CharacterListWindowTest)
CharacterListWindowTest() {
    owner := {_JSON: JSON.Load(default_JSON), EVEExe: "ahk_exe CharacterListMissingFixture.exe", Save_Settings_Delay_Timer: (*) => 0}
    owner._JSON["CharacterListEnabled"] := 1
    list := CharacterList(owner)
    try {
        AssertTrue(!!list.gui)
        list.items := [{name: "Alice", hwnd: 0}, {name: "Bob", hwnd: 0}]
        list.Render()
        AssertEqual(2, list.cells.Count)
        list.Collapse()
        AssertTrue(list.Get("Collapsed", false))
        list.Collapse()
        list.Put("Columns", 3)
        list.signature := ""
        list.Refresh()
        list.Disable()
        list.Refresh()
        AssertTrue(!list.gui)
    } finally {
        SetTimer(list.tick, 0)
        SetTimer(owner.Save_Settings_Delay_Timer, 0)
        OnMessage(0x201, list.mouse, 0)
        OnMessage(0x20A, list.wheelCallback, 0)
        OnMessage(0x232, list.endMove, 0)
        if list.gui
            list.gui.Destroy()
    }
}
