TestRunner.Register("Settings interaction mouse focus preserves the current pane position", SettingsMouseFocusStableTest)
SettingsMouseFocusStableTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.height := 500
        app.S_Gui.Show("NoActivate x-3000 y-3000 w1000 h500")
        Sleep(180) ; Let the initial resize rebuild finish before retaining handles.
        app.Layout()
        app.Navigate("close-protection")
        scroll := app.settingsScroller
        scroll.ScrollTo(50)
        before := scroll.offset
        AssertTrue(before > 0, "This check must begin with a scrolled pane.")
        list := app.tableEditors[1].list
        grid := app.tableEditors[1].grid
        GridPostPointer(grid, 1, 1, 1, 0x201)
        Sleep(360) ; Releasing a held click must also preserve the pane position.
        GridPostPointer(grid, 1, 1, 0, 0x202)
        AssertEqual(before, scroll.offset)
        scroll.MouseFocus(9, 0, 0x100, list.Hwnd)
        AssertEqual(0, scroll.mouseFocusUntil, "Keyboard navigation must immediately restore focus visibility.")
        app.Sidebar.GetPos(,,, &height)
        AssertEqual(app.height, height, "The sidebar must use the complete settings client height.")
    } finally app.Dispose()
}

GridTestPoint(grid, row, column) {
    rect := Buffer(16, 0)
    NumPut("Int", 0, "Int", column - 1, rect)
    SendMessage(0x1038, row - 1, rect.Ptr, grid.list)
    x := NumGet(rect, 0, "Int") + 10
    if column = 1
        x := 10
    y := Round((NumGet(rect, 4, "Int") + NumGet(rect, 12, "Int")) / 2)
    return (x & 0xFFFF) | ((y & 0xFFFF) << 16)
}
GridPostPointer(grid, row, column, buttons, message) {
    PostMessage(message, buttons, GridTestPoint(grid, row, column), grid.list)
    Sleep(35)
}

TestRunner.Register("Settings interaction first typed character retains its caret when starting a cell edit", SettingsFirstCharacterCaretTest)
SettingsFirstCharacterCaretTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.S_Gui.Show("NoActivate x-3000 y-3000 w1000 h740")
        Sleep(180)
        app.model.SaveRows("character-shortcuts", [["Alice", "F23"], ["Bob", "F24"]])
        app.Navigate("character-shortcuts")
        grid := app.tableEditors[1].grid
        for sample in [[1, false], [2, false], [3, false], [1, true], [2, true], [3, true]] {
            row := sample[1]
            grid.Select(row, 1)
            grid.list.Focus()
            if sample[2] {
                GridPostPointer(grid, row, 1, 1, 0x201)
                GridPostPointer(grid, row, 1, 0, 0x202)
            }
            PostMessage(0x102, Ord("X"), 0, IsObject(grid.editor) ? grid.editor : grid.list)
            Sleep(35)
            AssertTrue(IsObject(grid.editor))
            selection := SendMessage(0xB0, 0, 0, grid.editor)
            AssertEqual(1, selection & 0xFFFF, "The first typed character must not remain selected.")
            AssertEqual(1, selection >> 16)
            PostMessage(0x102, Ord("Y"), 0, grid.editor)
            Sleep(35)
            AssertEqual("XY", grid.editor.Value, "The next character must append instead of replacing the first.")
            grid.CancelEditor()
        }
    } finally app.Dispose()
}

TestRunner.Register("Settings interaction selecting scrolled group members preserves the table viewport", SettingsListViewportTest)
SettingsListViewportTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.S_Gui.Show("NoActivate x-3000 y-3000 w1000 h740")
        Sleep(180)
        app.model.AddGroup("characters", "Large fleet")
        rows := []
        loop 80
            rows.Push(["Pilot " A_Index])
        app.model.SaveRows("group-members", rows, "Large fleet")
        app.Navigate("character-groups")
        grid := app.tableEditors[1].grid
        grid.Select(1, 1)
        app.search.Focus()
        SendMessage(0x1014, 0, 700, grid.list)
        first := SendMessage(0x1027, 0, 0, grid.list)
        AssertTrue(first > 10)
        row := first + 5
        GridPostPointer(grid, row, 1, 1, 0x201)
        GridPostPointer(grid, row, 1, 0, 0x202)
        AssertEqual(first, SendMessage(0x1027, 0, 0, grid.list), "Selecting a visible row must not scroll the list to its old focused row.")
        AssertEqual(row, grid.row)
        ; Commit an edited cell while clicking another visible row.
        grid.editor.Value := "Updated pilot"
        nextRow := row + 2
        GridPostPointer(grid, nextRow, 1, 1, 0x201)
        GridPostPointer(grid, nextRow, 1, 0, 0x202)
        AssertEqual(first, SendMessage(0x1027, 0, 0, grid.list), "Repainting a committed cell must preserve the table's viewport.")
        AssertEqual(nextRow, grid.row, "Hit testing must refer to the row under the original pointer position.")
        AssertEqual("Updated pilot", app.model.Rows("group-members", "Large fleet")[row][1])
        GridPostPointer(grid, row, 1, 1, 0x201)
        GridPostPointer(grid, row, 1, 0, 0x202)
        AssertEqual(first, SendMessage(0x1027, 0, 0, grid.list), "Leaving an unchanged cell must preserve the viewport too.")
        AssertEqual(row, grid.row)
        grid.CancelEditor()
        ; Preserve horizontal position too when a table has wide columns.
        grid.list.ModifyCol(1, 1200)
        SendMessage(0x1014, 120, 0, grid.list)
        horizontal := DllCall("GetScrollPos", "Ptr", grid.list.Hwnd, "Int", 0, "Int")
        AssertTrue(horizontal > 0)
        GridPostPointer(grid, row, 1, 1, 0x201)
        GridPostPointer(grid, row, 1, 0, 0x202)
        grid.editor.Value := "Another update"
        AssertTrue(grid.CommitEditor())
        AssertEqual(horizontal, DllCall("GetScrollPos", "Ptr", grid.list.Hwnd, "Int", 0, "Int"))
        AssertEqual(first, SendMessage(0x1027, 0, 0, grid.list))
    } finally app.Dispose()
}

TestRunner.Register("Settings interaction real mouse drag selects cell ranges including inheritance", SettingsPointerRangeTest)
SettingsPointerRangeTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.model.SaveRows("character-shortcuts", [["Alice", "^F1"], ["Bob", "!F2"], ["Charlie", "^F3"]])
        app.Navigate("character-shortcuts")
        grid := app.tableEditors[1].grid
        GridPostPointer(grid, 1, 1, 1, 0x201)
        GridPostPointer(grid, 3, 2, 1, 0x200)
        GridPostPointer(grid, 3, 2, 0, 0x202)
        AssertEqual(6, grid.selected.Count)
        AssertFalse(IsObject(grid.editor), "Dragging must select cells instead of opening an editor.")
        app.CreateProfile("Fleet")
        app.Navigate("inheritance")
        grid := app.grids[1]
        GridPostPointer(grid, 1, 2, 1, 0x201)
        GridPostPointer(grid, 3, 2, 1, 0x200)
        GridPostPointer(grid, 3, 2, 0, 0x202)
        AssertEqual(3, grid.selected.Count)
    } finally app.Dispose()
}
TestRunner.Register("Settings interaction clearing a target deletes its paired assignment without validation errors", SettingsBlankRowsTest)
SettingsBlankRowsTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.model.SaveRows("character-shortcuts", [["Alice", "^F1"], ["Bob", "!F2"]])
        app.Navigate("character-shortcuts")
        grid := app.tableEditors[1].grid
        AssertTrue(grid.SetCells([[1, 1, ""]]))
        AssertDeepEqual([["Bob", "!F2"]], app.model.Rows("character-shortcuts"))
        AssertEqual(0, app.tableErrors.Count)
        AssertTrue(grid.SetCells([[1, 1, ""], [1, 2, ""]]))
        AssertEqual([], app.model.Rows("character-shortcuts"))
        AssertEqual(1, grid.list.GetCount())
        AssertTrue(grid.SetCells([[1, 2, "NoSuchKey123"]]))
        AssertEqual(0, app.tableErrors.Count, "Targetless rows ignore leftover values too.")
        AssertFalse(grid.SetCells([[1, 1, "Charlie"]]), "Identified rows must validate their shortcut.")
        AssertTrue(grid.SetCells([[1, 1, ""]]))
        AssertEqual([], app.model.Rows("character-shortcuts"))
        AssertEqual(1, grid.list.GetCount())
        AssertEqual(0, app.tableErrors.Count)
    } finally app.Dispose()
}
TestRunner.Register("Settings interaction gesture dropdown has real popup space and opens on one click", SettingsGestureDropdownTest)
SettingsGestureDropdownTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.S_Gui.Show("NoActivate x-3000 y-3000 w1000 h740")
        app.Navigate("mouse-actions")
        grid := app.tableEditors[1].grid
        GridPostPointer(grid, 1, 2, 1, 0x201)
        GridPostPointer(grid, 1, 2, 0, 0x202)
        AssertTrue(IsObject(grid.editor))
        AssertEqual("DDL", grid.editor.Type)
        bounds := Buffer(16, 0)
        AssertTrue(SendMessage(0x152, 0, bounds.Ptr, grid.editor))
        AssertTrue(NumGet(bounds, 12, "Int") - NumGet(bounds, 4, "Int") >= 150)
        AssertTrue(SendMessage(0x157, 0, 0, grid.editor))
        grid.editor.Choose(1)
        grid.ChoiceCommitted()
        AssertEqual(0, app.model.Rows("gestures")[1][2])
    } finally app.Dispose()
}
TestRunner.Register("Settings interaction Advanced edits raw shortcuts inline and retains invalid drafts", SettingsInlineAdvancedTest)
SettingsInlineAdvancedTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.Navigate("shortcuts")
        binding := app.Binding("HideThumbnailsHotkey")
        helper := binding.helper
        app.AdvancedShortcut(helper)
        AssertTrue(helper.rawMode)
        AssertFalse(WinGetStyle(helper.field.Hwnd) & 0x800)
        helper.field.Value := "NoSuchKey123"
        app.InlineShortcutChanged(helper)
        AssertTrue(app.shortcutErrors.Count > 0)
        AssertFalse(app.applyButton.Enabled)
        app.Navigate("appearance")
        AssertEqual("shortcuts", app.shortcutErrors[helper.rawKey].section, "An invalid draft must retain its original page when navigating.")
        app.Navigate("shortcuts")
        helper := app.Binding("HideThumbnailsHotkey").helper
        AssertEqual("NoSuchKey123", helper.field.Value)
        helper.field.Value := "^!F23"
        AssertTrue(app.CommitInlineShortcut(helper))
        AssertEqual("^!F23", app.model.Get("HideThumbnailsHotkey"))
        AssertEqual(0, app.shortcutErrors.Count)
        app.AdvancedShortcut(helper)
        AssertFalse(helper.rawMode)
        AssertTrue(WinGetStyle(helper.field.Hwnd) & 0x800)
        AssertFalse(app.HasOwnProp("activeDialog"))
        app.AdvancedShortcut(helper)
        helper.field.Value := "NoSuchKey456"
        app.InlineShortcutChanged(helper)
        app.DiscardDraft()
        AssertEqual(0, app.shortcutErrors.Count, "Discard must not re-import destroyed inline fields into the new draft.")
        AssertEqual("", app.model.Get("HideThumbnailsHotkey"))
    } finally app.Dispose()
}
TestRunner.Register("Settings interaction new subsection inheritance preserves effective values and runtime routing", SettingsSectionSourcesTest)
SettingsSectionSourcesTest() {
    config := JSON.Load(default_JSON)
    primary := config["_Profiles"]["Default"]
    config["_Profiles"]["Fleet"] := SettingsModel.Clone(primary)
    fleet := config["_Profiles"]["Fleet"]
    primary["Other"]["Global_Groups"]["Thumbnails Visuals"] := 1
    fleet["Other"]["Global_Groups"]["Thumbnails Visuals"] := 0
    primary["Thumbnails Visuals"]["ThumbnailTextSize"] := 14
    fleet["Thumbnails Visuals"]["ThumbnailTextSize"] := 20
    primary["Thumbnails Visuals"]["CustomThumbnailNames"] := Map("Characters", "Alice", "Names", "Default Scout")
    fleet["Thumbnails Visuals"]["CustomThumbnailNames"] := Map("Characters", "Alice", "Names", "Local Scout")
    fleet["Thumbnails Visuals"]["ThumbnailMinimumSize"]["width"] := 99
    fleet["Thumbnails Behavior"]["ThumbnailStartLocation"]["width"] := 400
    fleet["Thumbnails Behavior"]["ShiftThumbHorizontalStep"] := 77
    primary["Thumbnails Behavior"]["ThumbnailPollingInterval"] := 100
    fleet["Thumbnails Behavior"]["ThumbnailPollingInterval"] := 200
    config["LastUsedProfile"] := "Fleet"
    model := SettingsModel(config)
    AssertEqual("Default", model.Source("appearance"))
    AssertEqual("Default", model.Source("labels"))
    AssertEqual("Fleet", model.Source("layout"))
    AssertEqual(primary["Thumbnails Visuals"]["ThumbnailMinimumSize"]["width"], model.Get("ThumbnailMinimumSizewidth"))
    AssertTrue(model.Profiles["Fleet"].Has("SettingsMigrationBackup"))
    model.SetInheritance("appearance", false)
    AssertEqual(20, model.Get("ThumbnailTextSize"))
    AssertEqual("Default", model.Source("labels"), "Appearance must not control label inheritance anymore.")
    model.SetInheritance("layout", true)
    runtime := SettingsWorkbenchOwnerFixture()
    runtime._JSON := model.config
    runtime.ProfileOverride()
    AssertEqual(20, runtime.ThumbnailTextSize)
    AssertEqual("Default Scout", runtime.CustomThumbnailNames["Names"])
    model.SetInheritance("labels", false)
    runtime.ProfileOverride()
    AssertEqual("Local Scout", runtime.CustomThumbnailNames["Names"])
    runtime.model := model
    preview := JSON.Load(SettingsThumbnailPreview.Prototype.ReadSettings.Call({settings: runtime}))
    AssertEqual("Local Scout", preview["Thumbnails Visuals"]["CustomThumbnailNames"]["Names"], "Preview labels follow Labels, independently of Appearance.")
    AssertEqual(77, runtime.ShiftThumbHorizontalStep, "Layout inheritance must not change automatic placement.")
    AssertEqual(primary["Thumbnails Behavior"]["ThumbnailStartLocation"]["width"], runtime.ThumbnailStartLocation["width"])
    model.SetInheritance("performance", true)
    runtime.ProfileOverride()
    AssertEqual(100, runtime.ThumbnailPollingInterval, "Polling follows Performance, independently of Layout.")
    model.SetInheritance("performance", false)
    runtime.ProfileOverride()
    AssertEqual(200, runtime.ThumbnailPollingInterval)
    model.SetInheritance("window-behavior", true)
    model.SetInheritance("close-protection", false)
    model.SaveRows("exceptions", [["Alice", 1, 1]])
    runtime.ProfileOverride()
    AssertTrue(runtime.Dont_Minimize_Clients.Length > 0)
    model.AddGroup("apps", "Local tools")
    model.SaveRows("app-shortcuts", [["local.exe", "", "F23"]])
    model.SetInheritance("app-shortcuts", true)
    runtime.ProfileOverride()
    AssertTrue(runtime.NonEVEGroups.Has("Local tools"))
    AssertEqual(0, runtime.NonEVEHotkeys["exe"].Length, "Application groups and direct shortcuts have independent sources.")
}

class DirectPriorityFixture extends SettingsWorkbenchOwnerFixture {
    present := true
    ShortcutTargetAvailable(title) => this.present
    ShortcutScopeActive() => true
    ShortcutModifierDown(symbol, side) => false
}

TestRunner.Register("Settings interaction inherited subsections consistently disable labels and inputs", SettingsInheritedControlsTest)
SettingsInheritedControlsTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.model.AddGroup("characters", "Fleet")
        app.model.AddGroup("apps", "Tools")
        app.CreateProfile("Inherited")
        for group in SettingsModel.Groups
            app.model.SetInheritance(group, true)
        app.Changed()
        for section in SettingsSchema.Sections {
            if section.group = ""
                continue
            app.Navigate(section.id)
            for binding in app.bindings {
                AssertFalse(binding.label.Enabled, "Inherited label is active: " binding.field.id)
                for control in binding.inputs
                    AssertFalse(control.Enabled, "Inherited input is active: " binding.field.id)
            }
            for grid in app.grids
                AssertTrue(grid.readonly, "Inherited table must remain read-only: " section.id)
            for helper in app.recorders
                AssertFalse(helper.field.Enabled, "Inherited shortcut is active: " section.id)
        }
    } finally app.Dispose()
}

TestRunner.Register("Settings interaction direct shortcuts override group fallbacks only when available", SettingsDirectPriorityTest)
SettingsDirectPriorityTest() {
    app := DirectPriorityFixture()
    app._Hotkeys := [Map("EVE - Alice", "F23")]
    AssertTrue(app.DirectShortcutAvailable("*F23"))
    app.present := false
    AssertFalse(app.DirectShortcutAvailable("*F23"))
    entries := []
    HotkeyBindings.Add(entries, "character:1", "Alice", "F23", true)
    HotkeyBindings.Add(entries, "group:Fleet:ForwardsHotkey", "Fleet", "F23", true)
    AssertEqual(0, HotkeyBindings.Conflicts(entries).Count)
    HotkeyBindings.Add(entries, "setting:Hide", "Hide", "F23")
    AssertTrue(HotkeyBindings.Conflicts(entries).Count > 0)
}
