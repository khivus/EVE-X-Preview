class SettingsRowTable {
    __New(workbench, kind, groupName := "", ordered := false) {
        this.app := workbench
        this.kind := kind
        this.groupName := groupName
        this.ordered := ordered
        this.columns := SettingsTables.Columns(kind)
        this.rows := workbench.model.Rows(kind, groupName)
        this.inherited := workbench.model.Inherited(SettingsTables.Group(kind))
    }

    Build() {
        app := this.app
        headers := []
        for column in this.columns
            headers.Push(column.label)
        actions := []
        if this.kind != "gestures"
            actions.Push({label: "Add row", width: 96, callback: (*) => this.grid.AddRow()})
        actions.Push({label: "Row details...", width: 118, callback: (*) => this.EditSelected()})
        if this.kind != "gestures" {
            actions.Push({label: "Remove rows", width: 118, callback: (*) => this.grid.RemoveRows()})
            actions.Push({label: "Add running windows...", width: 184, callback: (*) => app.PickRunningWindows(!!InStr(this.kind, "app-"), ObjBindMethod(this, "ImportWindows"))})
        }
        if this.ordered {
            actions.Push({label: "Move up", width: 96, keepWithNext: true, callback: (*) => this.Move(-1)})
            actions.Push({label: "Move down", width: 96, callback: (*) => this.Move(1)})
        }
        buttons := app.ButtonRow(actions)
        offset := this.kind != "gestures" ? 1 : 0
        if offset
            this.add := buttons[1]
        this.editButton := buttons[1 + offset]
        if offset
            this.remove := buttons[3]
        if this.ordered
            this.up := buttons[-2], this.down := buttons[-1]
        for button in buttons
            button.Enabled := !this.inherited
        this.list := app.MainFrame.AddListView("x20 y" app.y " w" (app.contentWidth - 40) " h250 NoSortHdr", headers)
        this.list.OnEvent("ItemSelect", (*) => this.SelectionChanged())
        columns := []
        for column in this.columns
            columns.Push({label: column.label, kind: column.kind})
        this.grid := SettingsGrid(app, this.kind "|" this.groupName, this.list, columns, this.rows,
            ObjBindMethod(this, "SaveGridRows"), {append: this.kind != "gestures", readonly: this.inherited,
                group: SettingsTables.Group(this.kind), selection: ObjBindMethod(this, "SelectionChanged"), blank: ObjBindMethod(this, "BlankRow"),
                read: (*) => app.model.Rows(this.kind, this.groupName), remove: ObjBindMethod(this, "RemoveRows")})
        this.rows := this.grid.rows
        app.y += 258
        this.empty := app.Note("Click a cell to edit. Tab/Enter saves; Shift/Ctrl+click selects cells. Ctrl+C/V copies/pastes; Del clears; Ctrl+Del removes rows.", "555555", 36)
        this.ResizeColumns()
        this.SelectionChanged()
    }
    ResizeColumns() {
        available := this.app.contentWidth - 62
        for index, column in this.columns {
            width := this.columns.Length = 1 ? available : this.columns.Length = 2 ? Round(available / 2)
                : this.columns.Length = 4 ? Round(available * (index = 1 ? 0.34 : 0.22)) : Round(available / 3)
            this.list.ModifyCol(index, Max(80, width))
        }
    }
    SaveGridRows(rows) {
        this.app.model.SaveRows(this.kind, rows, this.groupName)
        this.rows := SettingsModel.Clone(rows)
        this.app.RefreshConflictNotice()
    }
    Refresh(select := 0) {
        this.rows := this.app.model.Rows(this.kind, this.groupName)
        this.grid.Reset(this.rows, select)
        this.ResizeColumns()
        this.SelectionChanged()
    }
    SelectionChanged() {
        if !this.HasOwnProp("grid") || this.grid.painting
            return
        selected := this.grid.selected.Count ? this.grid.row : 0
        SettingsTheme.Assign(this.editButton, "Enabled", selected && selected <= this.grid.rows.Length && !this.inherited)
        if this.HasOwnProp("remove")
            SettingsTheme.Assign(this.remove, "Enabled", selected && selected <= this.grid.rows.Length && !this.inherited)
        if this.ordered {
            SettingsTheme.Assign(this.up, "Enabled", selected > 1 && selected <= this.grid.rows.Length && !this.inherited)
            SettingsTheme.Assign(this.down, "Enabled", selected && selected < this.grid.rows.Length && !this.inherited)
        }
    }

    BlankRow() {
        row := []
        for index, column in this.columns {
            switch column.kind {
                case "bool": value := this.kind = "exceptions" && index = 3 ? 1 : 0
                case "color": value := this.app.model.Get(index = 2 ? "ClientHighligtColor" : index = 3 ? "InactiveClientBorderColor" : "ThumbnailTextColor")
                default: value := ""
            }
            row.Push(value)
        }
        return row
    }

    EditSelected() {
        selected := this.list.GetNext()
        if selected && !this.inherited
            this.Edit(selected)
    }

    Edit(index) {
        if this.inherited
            return
        this.grid.CommitEditor()
        if this.app.tableErrors.Has(this.grid.key) {
            this.app.LocateTableIssue()
            return
        }
        this.rows := this.grid.rows
        dialog := this.app.ChildDialog(index ? "Edit assignment" : "Add assignment")
        this.dialog := dialog
        this.editIndex := index
        this.editRow := index ? SettingsModel.Clone(this.rows[index]) : this.BlankRow()
        this.inputs := []
        y := 20
        width := 700
        for columnIndex, column in this.columns {
            value := this.editRow[columnIndex]
            if column.kind = "bool" {
                input := dialog.AddCheckbox("x20 y" y " w660 h28 Checked" (!!value), column.label)
                y += 36
            } else {
                dialog.AddText("x20 y" y " w660 h22", column.label)
                y += 26
                switch column.kind {
                    case "gesture":
                        labels := []
                        selected := 0
                        for optionIndex, option in SettingsTables.GestureOptions {
                            labels.Push(option.label)
                            if option.value = value
                                selected := optionIndex
                        }
                        input := dialog.AddDropDownList("x20 y" y " w660", labels)
                        if selected
                            input.Choose(selected)
                    case "hotkey":
                        input := dialog.AddEdit("x20 y" y " w398 h26 ReadOnly", SettingsValue.HotkeyText(value))
                        this.app.AddRecorder(dialog, input, y, 660, ObjBindMethod(this, "SetShortcut", columnIndex), value)
                    case "readonly":
                        input := dialog.AddEdit("x20 y" y " w660 h26 ReadOnly", this.kind = "gestures" ? SettingsTables.ActionName(value) : value)
                    default:
                        input := dialog.AddEdit("x20 y" y " w" (column.kind = "color" ? 480 : 660) " h26", value)
                        if column.kind = "color" {
                            picker := dialog.AddButton("x512 y" y " w168 h26", "Choose color...")
                            picker.OnEvent("Click", ObjBindMethod(this, "PickColor", input))
                        }
                }
                y += 34
            }
            this.inputs.Push(input)
            input.SettingsHelp := column.label ". " (column.kind = "hotkey" ? "Record a shortcut, clear it, or use Advanced syntax."
                : column.kind = "character" ? "Enter the original EVE character name."
                : column.kind = "exe" ? "Use the exact executable name, without a folder path." : "This value stays paired with the rest of this row.")
            input.ToolTip := input.SettingsHelp
        }
        if this.kind = "custom-names" {
            dialog.AddText("x20 y" y " w660 h34 c555555", "A blank displayed name keeps the character's normal label.")
            y += 44
        }
        if this.kind = "gestures" {
            dialog.AddText("x20 y" y " w660 h40 c555555", "Move and resize actions use dragging. Click actions trigger while the chosen buttons and modifiers are held.")
            y += 50
        }
        this.error := dialog.AddText("x20 y" y " w660 h34 cB42318", "")
        y += 42
        dialog.AddButton("x398 y" y " w136 h32 Default", "Save row to draft").OnEvent("Click", (*) => this.SaveDialog())
        dialog.AddButton("x546 y" y " w134 h32", "Cancel").OnEvent("Click", (*) => this.app.CloseChild(dialog))
        dialog.Show("w700 h" (y + 52))
    }

    SetShortcut(index, value) => this.editRow[index] := value

    PickColor(input, *) {
        value := ColorPicker.Show(this.dialog.Hwnd, input.Value)
        if value != ""
            input.Value := value
    }

    SaveDialog() {
        if !this.app.CommitInlineShortcuts(this.dialog) {
            this.dialog.SetError(this.error, "Complete the shortcut syntax before saving this row.")
            return
        }
        row := this.editRow.Clone()
        for index, column in this.columns {
            if column.kind = "hotkey" || column.kind = "readonly"
                continue
            row[index] := column.kind = "gesture" ? (this.inputs[index].Value ? SettingsTables.GestureOptions[this.inputs[index].Value].value : -1)
                : column.kind = "bool" ? this.inputs[index].Value : Trim(this.inputs[index].Value)
        }
        rows := SettingsModel.Clone(this.rows)
        if this.editIndex
            rows[this.editIndex] := row
        else
            rows.Push(row)
        try {
            this.app.model.SaveRows(this.kind, rows, this.groupName)
            this.app.Changed()
            this.Refresh(this.editIndex ? this.editIndex : rows.Length)
            this.app.RefreshConflictNotice()
            this.app.CloseChild(this.dialog)
        } catch as err
            this.dialog.SetError(this.error, err.Message)
    }

    RemoveSelected() => this.grid.RemoveRows()
    RemoveRows(selected) {
        if this.inherited || !selected.Count
            return
        old := SettingsModel.Clone(this.grid.rows), rows := SettingsModel.Clone(old), removed := []
        for index, row in old
            if selected.Has(index)
                removed.Push({index: index, row: row})
        index := rows.Length
        while index > 0 {
            if selected.Has(index)
                rows.RemoveAt(index)
            index--
        }
        if this.grid.SaveCandidate(rows) && removed.Length {
            this.app.model.undo.Push({kind: "grid-batch", profile: this.app.model.profile,
                section: this.app.currentSection, key: this.grid.key, entries: removed})
            this.rows := rows
            this.app.UpdateStatus("Rows removed. Undo restores the complete assignments.")
        }
    }

    Move(direction) {
        this.grid.CommitEditor()
        this.rows := this.grid.rows
        index := this.list.GetNext()
        destination := index + direction
        if !index || destination < 1 || destination > this.rows.Length || this.inherited
            return
        rows := SettingsModel.Clone(this.rows)
        row := rows.RemoveAt(index)
        rows.InsertAt(destination, row)
        this.app.model.SaveRows(this.kind, rows, this.groupName)
        this.app.Changed()
        this.Refresh(destination)
    }

    ImportWindows(windows) {
        this.grid.CommitEditor()
        this.rows := this.grid.rows
        rows := SettingsModel.Clone(this.rows)
        for item in windows {
            row := this.BlankRow()
            if InStr(this.kind, "app-")
                row[1] := item.exe, row[2] := item.title
            else
                row[1] := item.character
            duplicate := false
            for existing in rows {
                if StrLower(existing[1]) = StrLower(row[1])
                    && (!InStr(this.kind, "app-") || existing[2] = row[2])
                    duplicate := true
            }
            if !duplicate
                rows.Push(row)
        }
        this.app.model.SaveRows(this.kind, rows, this.groupName)
        this.app.Changed()
        this.Refresh(rows.Length)
        this.app.RefreshConflictNotice()
    }
}
