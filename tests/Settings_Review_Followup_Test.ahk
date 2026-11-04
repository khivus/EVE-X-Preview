TestRunner.Register("Settings review scalar previews request immediate native paint without scrolling", SettingsReviewSwatchPaintTest)
SettingsReviewSwatchPaintTest() {
    app := SettingsColorChoiceFixture(), observer := 0
    try {
        app.S_Gui.Show("NoActivate x-3000 y-3000 w1000 h740")
        Sleep(180)
        for mode in ["dark", "light"] {
            app.model.Set("SettingsTheme", mode)
            app.Changed()
            app.Navigate("appearance")
            binding := app.Binding("ThumbnailTextColor")
            observer := SettingsPaintObserver([binding.swatch.Hwnd])
            binding.control.Value := "12AB34"
            app.TextChanged(binding, binding.control)
            AssertEqual("12AB34", binding.swatch.ThemeBackground)
            AssertTrue(observer.counts[binding.swatch.Hwnd].paints > 0, "Typing must repaint the swatch now, before any print, scroll or view reset.")
            observer.Reset()
            app.colorChoice := "AB3412"
            app.PickColor(binding, binding.control)
            AssertEqual("AB3412", binding.swatch.ThemeBackground)
            AssertTrue(observer.counts[binding.swatch.Hwnd].paints > 0)
            observer.Dispose(), observer := 0
        }
    } finally {
        if IsObject(observer)
            observer.Dispose()
        app.Dispose()
    }
}

class SettingsReviewHintProbe extends SettingsPaintObserver {
    __New(control) {
        this.control := control, this.lastColor := ""
        super.__New([control.Hwnd])
    }
    WindowProc(handle, message, wParam, lParam, subclass, data) {
        result := super.WindowProc(handle, message, wParam, lParam, subclass, data)
        if message = 0xF
            this.lastColor := this.control.ThemeColor
        return result
    }
}

TestRunner.Register("Settings review validation text and color follow the current input and related sizes", SettingsReviewValidationTest)
SettingsReviewValidationTest() {
    app := SettingsWorkbenchFixture(), probe := 0, statusProbe := 0
    try {
        app.S_Gui.Show("NoActivate x-3000 y-3000 w1000 h740")
        Sleep(180)
        statusProbe := SettingsReviewHintProbe(app.status)
        for mode in ["dark", "light"] {
            app.model.Set("SettingsTheme", mode)
            app.Changed()
            app.Navigate("layout")
            binding := app.Binding("ThumbnailStartLocationwidth")
            app.settingsScroller.EnsureVisible(binding.hint)
            probe := SettingsReviewHintProbe(binding.hint)
            minimum := app.model.Get("ThumbnailMinimumSizewidth")
            for value in [minimum - 1, minimum + 1, "invalid", minimum + 2] {
                binding.control.Value := value
                app.TextChanged(binding, binding.control)
                problem := value = "invalid" || value < minimum
                AssertEqual(problem, app.ValidationMessage(binding.key) != "")
                AssertEqual(problem ? "B42318" : "555555", binding.hint.ThemeColor)
                AssertEqual(problem ? app.ValidationMessage(binding.key) : app.FieldHint(binding.field), binding.hint.Value)
                AssertEqual(binding.hint.ThemeColor, probe.lastColor, "The final native paint must already use the current error/help color.")
                AssertEqual(problem ? "B42318" : "444444", statusProbe.lastColor, "Footer validation colors must recover with the field hints.")
            }
            probe.Dispose(), probe := 0
            binding := app.Binding("ThumbnailMinimumSizewidth")
            maximum := app.model.Get("ThumbnailStartLocationwidth")
            for value in [maximum + 1, maximum - 1] {
                binding.control.Value := value
                app.TextChanged(binding, binding.control)
                AssertEqual(value > maximum ? "B42318" : "555555", binding.hint.ThemeColor)
            }
        }
    } finally {
        if IsObject(probe)
            probe.Dispose()
        if IsObject(statusProbe)
            statusProbe.Dispose()
        app.Dispose()
    }
}

TestRunner.Register("Settings review slider marks belong to native paint throughout tracking", SettingsReviewTicksTest)
SettingsReviewTicksTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.model.Set("SettingsTheme", "dark")
        app.Changed()
        app.Navigate("sounds")
        binding := app.Binding("MasterSoundVolume")
        AssertTrue(WinGetStyle(binding.slider.Hwnd) & 1)
        AssertFalse(WinGetStyle(binding.slider.Hwnd) & 0x10)
        channel := Buffer(16, 0), rect := Buffer(16, 0)
        SendMessage(0x41A, 0, channel.Ptr, binding.slider)
        DllCall("GetClientRect", "Ptr", binding.slider.Hwnd, "Ptr", rect)
        x := Round(NumGet(channel, 0, "Int") + (NumGet(channel, 8, "Int") - NumGet(channel, 0, "Int") - 1) / 2)
        y := NumGet(rect, 12, "Int") - 3
        for value in [110, 137, 183, 71] {
            binding.slider.Value := value
            app.SliderChanged(binding, binding.slider, binding.control, binding.slider, 5)
            AssertEqual(SettingsTheme.BGR("7D8793"), SettingsPaintPixel(binding.slider, x, y))
        }
    } finally app.Dispose()
}

TestRunner.Register("Settings review invalid shortcut row details retain the table and ignore disposed editors", SettingsReviewInvalidDetailsTest)
SettingsReviewInvalidDetailsTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.model.SaveRows("character-shortcuts", [["Alpha", "F13"]])
        app.Navigate("character-shortcuts")
        table := app.tableEditors[1], grid := table.grid, hwnd := grid.hwnd
        grid.Select(1, 2), grid.BeginEdit()
        editor := grid.editor
        editor.Value := "DefinitelyNotAKey"
        table.Edit(1)
        AssertTrue(app.tableErrors.Has(grid.key))
        AssertEqual(hwnd, app.grids[1].hwnd, "Locating an error on this page must not destroy the active grid.")
        AssertFalse(app.HasOwnProp("activeDialog"))
        app.Navigate("about")
        grid.EditorLostFocus(editor)
        AssertEqual(0, grid.Viewport())
        AssertFalse(grid.SetCells([[1, 2, "F14"]]))
    } finally app.Dispose()
}

TestRunner.Register("Settings review commits tolerate a save callback disposing its own grid", SettingsReviewDisposeDuringSaveTest)
SettingsReviewDisposeDuringSaveTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.Navigate("character-shortcuts")
        grid := app.grids[1]
        grid.save := (*) => app.Navigate("about")
        grid.Select(1, 1), grid.BeginEdit()
        grid.editor.Value := "Alpha"
        AssertTrue(grid.CommitEditor())
        AssertTrue(grid.disposed)
        AssertEqual("about", app.currentSection)
    } finally app.Dispose()
}

TestRunner.Register("Settings review inheritance templates persist and new custom sections copy effective values", SettingsReviewInheritanceModelTest)
SettingsReviewInheritanceModelTest() {
    model := SettingsModel(JSON.Load(default_JSON))
    flags := model.CreationInheritance()
    flags["labels"] := 1, flags["quick-group"] := 1
    model.SetCreationInheritance(flags)
    model.AddProfile("Fleet", true, flags)
    AssertTrue(model.Inherited("labels"))
    AssertFalse(model.Inherited("layout"))
    model.Profiles["Default"]["Thumbnails Interactions"]["QuickGroupColor"] := "ABCDEF"
    expected := Map()
    for id, field in SettingsSchema.Fields
        if !field.global
            expected[id] := model.Get(id)
    none := Map()
    model.AddProfile("Independent", true, none)
    for group in SettingsModel.Groups
        AssertFalse(model.Inherited(group))
    for id, value in expected
        AssertDeepEqual(value, model.Get(id), "Copy effective value: " id)
    reopened := SettingsModel(model.Merge(model.original))
    AssertEqual(flags, reopened.CreationInheritance())
    AssertEqual(1, reopened.Profiles["Fleet"]["SettingsInheritance"]["labels"])
}

TestRunner.Register("Settings review profile creation offers individual inheritance rows and three bulk presets", SettingsReviewProfileDialogTest)
SettingsReviewProfileDialogTest() {
    app := SettingsWorkbenchFixture()
    try {
        app.Navigate("inheritance")
        grid := app.grids[1]
        AssertEqual("bool", grid.columns[2].kind)
        AssertTrue(grid.SetCells([[1, 2, 1]]))
        app.AskName("Create profile", "", ObjBindMethod(app, "CreateProfile"), true)
        dialog := app.activeDialog, form := dialog.profileCreation
        AssertEqual(1, form.grid.rows[1][2])
        for pair in [[3, 1], [2, 0], [1, -1]] {
            index := pair[1], expected := pair[2]
            PostMessage(0xF5, 0, 0, form.presets[index])
            Sleep(30)
            for row, value in form.grid.rows
                AssertEqual(expected = -1 ? (row = 1 ? 1 : 0) : expected, value[2])
        }
        form.name.Value := "Created with template"
        PostMessage(0xF5, 0, 0, form.create)
        Sleep(40)
        AssertFalse(app.HasOwnProp("activeDialog"))
        AssertEqual("Created with template", app.model.profile)
        AssertTrue(app.model.Inherited(SettingsModel.Groups[1]))
        AssertFalse(app.model.Inherited(SettingsModel.Groups[2]))
    } finally app.Dispose()
}

TestRunner.Register("Settings review popup centering stays inside the owner's monitor and version rows align", SettingsReviewAlignmentTest)
SettingsReviewAlignmentTest() {
    work := {left: -2560, top: 0, right: 0, bottom: 1400}
    position := SettingsWindowSizing.CenterPosition({left: -2200, top: 200, right: -1200, bottom: 940}, 500, 300, work)
    AssertEqual(-1950, position.x), AssertEqual(420, position.y)
    position := SettingsWindowSizing.CenterPosition({left: -100, top: 1200, right: 900, bottom: 1940}, 500, 300, work)
    AssertEqual(-500, position.x), AssertEqual(1100, position.y)
    app := SettingsWorkbenchFixture()
    popup := 0
    try {
        app.Navigate("about")
        app.installedVersion.GetPos(&installedX)
        app.stableVersion.GetPos(&stableX)
        app.preVersion.GetPos(&preX)
        AssertEqual(installedX, stableX), AssertEqual(stableX, preX)
        width := 0
        for control in app.MainFrame
            if control.Type = "Button" {
                control.GetPos(,, &w)
                if width
                    AssertEqual(width, w, "Stacked action buttons share a width.")
                width := w
            }
        AssertTrue(width > 0)
        popup := Gui("+Owner" app.S_Gui.Hwnd)
        popup.Show("Hide w400 h200")
        dialog := ColorPicker.BuildDialog(app.S_Gui.Hwnd, "123456")
        ColorPicker.CenterDialog(popup.Hwnd, 0x110, 0, dialog.Ptr)
        ownerRect := Buffer(16, 0), popupRect := Buffer(16, 0)
        DllCall("GetWindowRect", "Ptr", app.S_Gui.Hwnd, "Ptr", ownerRect)
        DllCall("GetWindowRect", "Ptr", popup.Hwnd, "Ptr", popupRect)
        expected := SettingsWindowSizing.CenterPosition({left: NumGet(ownerRect, 0, "Int"), top: NumGet(ownerRect, 4, "Int"),
            right: NumGet(ownerRect, 8, "Int"), bottom: NumGet(ownerRect, 12, "Int")},
            NumGet(popupRect, 8, "Int") - NumGet(popupRect, 0, "Int"), NumGet(popupRect, 12, "Int") - NumGet(popupRect, 4, "Int"),
            SettingsWindowSizing.WorkArea(app.S_Gui.Hwnd))
        AssertEqual(expected.x, NumGet(popupRect, 0, "Int"))
        AssertEqual(expected.y, NumGet(popupRect, 4, "Int"))
    } finally {
        if IsObject(popup)
            popup.Destroy()
        app.Dispose()
    }
}
