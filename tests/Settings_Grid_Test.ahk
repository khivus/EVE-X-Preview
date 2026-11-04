TestRunner.Register("Grid entry rows edit inline and keep exactly one trailing empty row", GridEntryTest)
GridEntryTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.model.SaveRows("character-shortcuts", [["Alice", "^F1"]])
        app.Navigate("character-shortcuts")
        grid := app.tableEditors[1].grid
        AssertEqual(2, grid.list.GetCount())
        grid.Select(2, 1)
        grid.BeginEdit("Bob")
        AssertTrue(IsObject(grid.editor), "BeginEdit must open an inline editor.")
        AssertTrue(grid.CommitEditor(), "Bob entry must save: " app.status.Value)
        AssertEqual(3, grid.list.GetCount())
        AssertEqual("Bob", app.model.Rows("character-shortcuts")[2][1])
        grid.Select(2, 2)
        grid.BeginEdit("!F2")
        grid.Scrolled(0, 0, 0x20A, grid.list.Hwnd)
        AssertFalse(IsObject(grid.editor), "Finish a cell edit before the list scrolls underneath its overlay.")
        AssertEqual("!F2", app.model.Rows("character-shortcuts")[2][2])
        app.Navigate("appearance")
        app.Navigate("character-shortcuts")
        AssertEqual(3, app.tableEditors[1].list.GetCount())
        AssertEqual(0, app.owner.saves)
        grid := app.tableEditors[1].grid
        grid.Select(3, 1)
        grid.BeginEdit()
        grid.CommitEditor()
        AssertEqual(0, app.tableErrors.Count, "Leaving an untouched entry row must not create an invalid assignment.")
        app.Navigate("close-protection")
        grid := app.tableEditors[1].grid
        AssertTrue(grid.SetCells([[1, 2, 1]]))
        AssertEqual(0, app.tableErrors.Count, "Rows without a target are ignored, without warning, until a target is entered.")
        AssertTrue(grid.SetCells([[1, 1, "Charlie"]]), "Completing a boolean-first row must save: " app.status.Value)
        AssertEqual(1, app.model.Rows("exceptions")[1][2])
        AssertEqual(2, grid.list.GetCount())
    } finally app.Dispose()
}

TestRunner.Register("Grid column ranges support type-to-fill copy paste multirow removal and Undo", GridBulkTest)
GridBulkTest() {
    app := SettingsWorkbenchFixture()
    clipboard := ClipboardAll()
    try {
        app.model.SaveRows("exceptions", [["Alice", 0, 0], ["Bob", 0, 0]])
        ; Exception rows need at least one enabled flag to be represented.
        app.model.SaveRows("exceptions", [["Alice", 0, 1], ["Bob", 0, 1]])
        app.Navigate("close-protection")
        grid := app.tableEditors[1].grid
        grid.Select(1, 2)
        grid.Character(49, 0, 0x102, grid.list.Hwnd)
        AssertEqual(1, app.model.Rows("exceptions")[1][2])
        grid.Select(2, 2, true)
        AssertEqual(2, grid.selected.Count)
        grid.Character(49, 0, 0x102, grid.list.Hwnd)
        AssertDeepEqual([["Alice", 1, 1], ["Bob", 1, 1]], app.model.Rows("exceptions"))
        AssertEqual("1`r`n1", grid.Copy())
        grid.Select(1, 2)
        AssertTrue(grid.Paste("0`t1`n1`t0"))
        AssertDeepEqual([["Alice", 0, 1], ["Bob", 1, 0]], app.model.Rows("exceptions"))
        grid.Select(1, 1)
        grid.Select(2, 3, true)
        grid.RemoveRows()
        AssertEqual([], app.model.Rows("exceptions"))
        AssertEqual(1, grid.list.GetCount())
        app.Undo()
        AssertDeepEqual([["Alice", 0, 1], ["Bob", 1, 0]], app.model.Rows("exceptions"))
    } finally {
        app.Dispose()
        A_Clipboard := clipboard
    }
}

TestRunner.Register("Grid invalid and partial rows survive navigation block Apply and recover atomically", GridValidationTest)
GridValidationTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.model.SaveRows("app-shortcuts", [["alpha.exe", "", "^F1"]])
        app.Navigate("app-shortcuts")
        grid := app.tableEditors[1].grid
        grid.Select(2, 1)
        AssertFalse(grid.Paste("not an executable"))
        AssertEqual(1, app.model.Rows("app-shortcuts").Length)
        AssertEqual(1, app.tableErrors.Count)
        grid.BeginEdit()
        AssertFalse(grid.CommitEditor(), "An unchanged invalid cell must still report its validation failure.")
        grid.Select(2, 1)
        grid.RemoveRows()
        AssertEqual(0, app.tableErrors.Count)
        app.Undo()
        AssertEqual(1, app.tableErrors.Count, "Undo must restore an unfinished row without writing invalid data into the model.")
        AssertFalse(app.applyButton.Enabled)
        app.Navigate("appearance")
        app.Navigate("app-shortcuts")
        grid := app.tableEditors[1].grid
        AssertEqual("not an executable", grid.rows[2][1])
        AssertTrue(grid.SetCells([[2, 1, "beta.exe"], [2, 3, "!F2"]]))
        AssertEqual(0, app.tableErrors.Count)
        AssertEqual(2, app.model.Rows("app-shortcuts").Length)
        AssertTrue(app.applyButton.Enabled)
        AssertFalse(grid.SetCells([[2, 1, "alpha.exe"]]))
        AssertEqual("beta.exe", app.model.Rows("app-shortcuts")[2][1])
        app.DiscardDraft()
        AssertEqual(0, app.tableErrors.Count)
    } finally app.Dispose()
}

TestRunner.Register("Grid event values and shared sound gains edit inline without adding event names", GridMonitoringTest)
GridMonitoringTest() {
    app := SettingsWorkbenchFixture()
    path := CreateTestSound("grid audio")
    try {
        app.Navigate("events")
        grid := app.eventTable.grid
        AssertEqual(18, grid.list.GetCount())
        grid.Select(1, 6)
        grid.Select(2, 6, true)
        AssertTrue(grid.Paste("125"))
        AssertEqual(125, app.eventTable.events["stoppedShooting"]["soundVolume"])
        AssertEqual(125, app.eventTable.events["underAttackByPlayer"]["soundVolume"])
        AssertFalse(grid.SetCells([[1, 6, 301]]))
        AssertEqual(125, app.eventTable.events["stoppedShooting"]["soundVolume"])
        AssertTrue(grid.SetCells([[1, 6, 100]]))
        app.Navigate("sounds")
        grid := app.soundLibrary.grid
        AssertTrue(grid.SetCells([[1, 1, path], [1, 3, 127]]))
        AssertEqual([path], EventSounds.Library(app.model.config))
        AssertEqual(127, EventSounds.SoundVolume(app.model.config, path))
        AssertEqual(2, grid.list.GetCount())
        AssertEqual([], EventSounds.Library(app.owner._JSON))
    } finally {
        app.Dispose()
        FileDelete(path)
    }
}

TestRunner.Register("Grid navigation stays expanded and source labels precede descriptions", GridNavigationTest)
GridNavigationTest() {
    app := SettingsWorkbenchFixture()
    try {
        first := app.navRoots[1], second := app.navRoots[2]
        app.Navigate("character-groups")
        app.Navigate("appearance")
        AssertTrue(SendMessage(0x1127, first, 0x20, app.nav) & 0x20)
        AssertTrue(SendMessage(0x1127, second, 0x20, app.nav) & 0x20)
        app.nav.Modify(first, "-Expand")
        app.NavigationSelected(app.nav, first)
        AssertFalse(SendMessage(0x1127, first, 0x20, app.nav) & 0x20, "An explicit collapse must remain collapsed even when its category gains selection.")
        app.Navigate("events")
        sourceY := 100000, descriptionY := 0
        for control in app.MainFrame
            if control.Type = "Text" {
                if InStr(control.Text, "Source: Default")
                    control.GetPos(, &sourceY)
                if control.Text = SettingsSchema.FindSection("events").help
                    control.GetPos(, &descriptionY)
            }
        AssertTrue(sourceY < descriptionY)
        for width in [1000, 800, 1120, 780, 1000] {
            app.width := width, app.height := 600
            app.Layout()
            previous := 0
            for button in [app.closeButton, app.undoButton, app.discardButton, app.applyButton] {
                button.GetPos(&x,, &w)
                AssertTrue(x >= previous)
                previous := x + w
            }
            AssertTrue(previous <= width)
        }
    } finally app.Dispose()
}

TestRunner.Register("Grid slider scales snap to five while exact fields and wrapped DPS help stay precise", GridSlidersTest)
GridSlidersTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.Navigate("sounds")
        binding := app.Binding("MasterSoundVolume")
        AssertTrue(SendMessage(0x410, 0, 0, binding.slider) >= 10)
        binding.slider.GetPos(&sx,, &sw)
        binding.control.GetPos(&ex,, &ew)
        AssertEqual(8, ex - sx - sw, "Slider fields use the common eight-pixel gap.")
        AssertEqual(app.contentWidth - 20, ex + ew, "Slider numeric fields align to the setting row's right edge.")
        app.model.Set("MasterSoundVolume", 125)
        app.Changed()
        app.Navigate("events")
        event := SettingsEventEditor(app.eventTable, "underAttackByPlayer", 2)
        event.Show()
        AssertTrue(WinGetStyle(event.slider.Hwnd) & 1, "Event slider ticks must belong to the slider, so tracking cannot erase overlapping labels.")
        AssertFalse(WinGetStyle(event.slider.Hwnd) & 0x10)
        app.CloseChild(event.dialog)
        AssertTrue(app.applyButton.Enabled, "The footer must refresh after an owned dialog closes.")
        app.Navigate("sounds")
        binding := app.Binding("MasterSoundVolume")
        binding.slider.Value := 128
        app.SliderChanged(binding, binding.slider, binding.control)
        AssertEqual(130, app.model.Get("MasterSoundVolume"))
        binding.control.Value := 127
        app.TextChanged(binding, binding.control)
        AssertEqual(127, app.model.Get("MasterSoundVolume"))
        app.width := 780
        app.Layout()
        app.Navigate("dps")
        binding := app.Binding("showIncomingDPSResistances")
        binding.hint.GetPos(,, &width, &height)
        AssertTrue(height >= SettingsLayout.TextHeight(binding.hint, app.FieldHint(binding.field), width))
    } finally app.Dispose()
}

TestRunner.Register("Grid profile rename remove and Undo keep unfinished rows attached to their profile", GridProfileDraftTest)
GridProfileDraftTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.CreateProfile("Fleet")
        app.Navigate("app-shortcuts")
        grid := app.tableEditors[1].grid
        AssertFalse(grid.SetCells([[1, 1, "unfinished"]]))
        app.RenameProfile("Logistics")
        AssertEqual("Logistics", app.model.profile)
        AssertEqual(1, app.tableErrors.Count)
        for key, issue in app.tableErrors
            AssertEqual("Logistics", issue.profile)
        app.Navigate("profiles")
        app.RemoveProfile()
        AssertEqual(0, app.tableErrors.Count)
        app.Undo()
        AssertEqual("Logistics", app.model.profile)
        AssertEqual(1, app.tableErrors.Count)
        app.Navigate("app-shortcuts")
        AssertEqual("unfinished", app.tableEditors[1].grid.rows[1][1])
        AssertTrue(app.tableEditors[1].grid.SetCells([[1, 1, "logistics.exe"]]))
        AssertEqual(0, app.tableErrors.Count)
    } finally app.Dispose()
}

class GridPickerFixture extends SettingsWorkbenchFixture {
    RunningWindows(applications) => [{exe: "alpha.exe", title: "Window A", character: "Alice"}, {exe: "beta.exe", title: "Window B", character: "Bob"}]
}
TestRunner.Register("Grid picker selection and inheritance bulk source choices use the shared controls", GridPickerInheritanceTest)
GridPickerInheritanceTest() {
    app := GridPickerFixture()
    clipboard := ClipboardAll()
    try {
        holder := {windows: []}
        app.PickRunningWindows(false, (windows) => holder.windows := windows)
        grid := app.grids[-1]
        grid.Select(1, 1)
        grid.Select(2, 1, true)
        AssertEqual("Alice`r`nBob", grid.Copy())
        add := WorkbenchDialogControl(app.activeDialog, "Button", "Add selected")
        SendMessage(0xF5, 0, 0, add)
        Sleep(20)
        AssertEqual(2, holder.windows.Length)
        AssertFalse(app.HasOwnProp("activeDialog"))
        app.CreateProfile("Fleet")
        app.Navigate("inheritance")
        grid := app.grids[1]
        grid.ColumnSelection(grid.list, 2)
        AssertTrue(grid.Paste("1"))
        for group in SettingsModel.Groups
            AssertEqual("Default", app.model.Source(group))
        AssertTrue(grid.Paste("2"))
        for group in SettingsModel.Groups
            AssertEqual("Fleet", app.model.Source(group))
    } finally {
        app.Dispose()
        A_Clipboard := clipboard
    }
}
