; Shared native spreadsheet interactions. Configuration remains staged in the
; workbench; invalid identified rows survive page changes and block Apply.
class SettingsGrid {
    static PointerMessage := DllCall("RegisterWindowMessageW", "Str", "EVE-X-Preview.SettingsGrid.Pointer", "UInt")

    __New(app, id, list, columns, rows, save, options?) {
        this.app := app, this.id := id, this.list := list, this.columns := columns, this.save := save
        this.hwnd := list.Hwnd
        this.options := IsSet(options) ? options : {}
        this.append := this.Option("append", true), this.readonly := this.Option("readonly", false)
        this.parent := this.Option("parent", app.MainFrame)
        this.key := (this.Option("shared", false) ? "*" : app.model.profile) "|" id
        this.profile := app.model.profile, this.section := app.currentSection
        this.group := this.Option("group", "")
        this.snapshot := JSON.Dump(rows)
        this.rows := SettingsModel.Clone(rows)
        if app.tableStates.Has(this.key) && app.tableStates[this.key].snapshot = this.snapshot
            this.rows := SettingsModel.Clone(app.tableStates[this.key].rows)
        this.selected := Map(), this.row := 1, this.column := 1, this.anchor := [1, 1]
        this.disposed := false, this.painting := false, this.editor := 0
        this.clickCallback := ObjBindMethod(this, "Click")
        this.drawCallback := ObjBindMethod(this, "Draw")
        this.keyCallback := ObjBindMethod(this, "HandleKey")
        this.charCallback := ObjBindMethod(this, "Character")
        this.scrollCallback := ObjBindMethod(this, "Scrolled")
        this.pointerCallback := ObjBindMethod(this, "Pointer")
        this.pointerDispatchCallback := ObjBindMethod(this, "DispatchPointer")
        this.pendingPointer := []
        this.processingPointer := false
        this.pointerDown := false
        this.drag := 0
        list.OnNotify(-2, this.clickCallback)
        list.OnNotify(-12, this.drawCallback)
        list.OnEvent("ItemSelect", ObjBindMethod(this, "NativeSelection"))
        list.OnEvent("ColClick", ObjBindMethod(this, "ColumnSelection"))
        OnMessage(0x100, this.keyCallback), OnMessage(0x102, this.charCallback)
        for message in [0x20A, 0x115, 0x114, 0x201]
            OnMessage(message, this.scrollCallback)
        for message in [0x201, 0x200, 0x202]
            OnMessage(message, this.pointerCallback)
        OnMessage(SettingsGrid.PointerMessage, this.pointerDispatchCallback)
        list.ToolTip := "Click a cell to edit. Shift/Ctrl+click selects cells; Ctrl+C/V copies/pastes. Del clears; Ctrl+Del removes rows."
        list.SettingsHelp := list.ToolTip
        SendMessage(0x1036, 0x4001, app.theme.dark ? 0x4000 : 0x4001, list) ; Native grid lines are too bright on dark surfaces.
        app.grids.Push(this)
        this.Paint()
    }
    Option(name, fallback) => this.options.HasOwnProp(name) ? this.options.%name% : fallback
    RedrawCells() {
        if this.disposed
            return
        rect := Buffer(16, 0)
        DllCall("GetClientRect", "Ptr", this.list.Hwnd, "Ptr", rect)
        header := SendMessage(0x101F, 0, 0, this.list)
        if header && SettingsNativeChrome.Long(header, -16) & 0x10000000 {
            headerRect := Buffer(16, 0)
            DllCall("GetWindowRect", "Ptr", header, "Ptr", headerRect)
            DllCall("MapWindowPoints", "Ptr", 0, "Ptr", this.list.Hwnd, "Ptr", headerRect, "UInt", 2)
            NumPut("Int", Max(0, Min(NumGet(rect, 12, "Int"), NumGet(headerRect, 12, "Int"))), rect, 4)
        }
        if NumGet(rect, 12, "Int") > NumGet(rect, 4, "Int")
            DllCall("RedrawWindow", "Ptr", this.list.Hwnd, "Ptr", rect, "Ptr", 0, "UInt", 0x141)
    }
    Editable(column) => !this.readonly && this.columns[column].kind != "readonly"
    RowCount => this.rows.Length + (this.append ? 1 : 0)
    Blank() {
        if this.options.HasOwnProp("blank")
            return this.options.blank.Call()
        result := []
        for column in this.columns
            result.Push(column.kind = "bool" ? 0 : "")
        return result
    }
    Display(value, column) {
        spec := this.columns[column]
        if spec.kind = "bool"
            return value = "" ? "" : value = 1 ? "Yes" : value = 0 ? "No" : value
        if spec.kind = "gesture" {
            for option in SettingsTables.GestureOptions
                if option.value = value
                    return option.label
        }
        if spec.kind = "hotkey"
            return SettingsValue.HotkeyText(value)
        if spec.kind = "choice" {
            for option in spec.choices
                if option.value = value
                    return option.label
        }
        if spec.kind = "readonly" && InStr(this.id, "gestures|") = 1
            return SettingsTables.ActionName(value)
        return value
    }
    Paint() {
        if this.disposed
            return
        this.painting := true
        try {
            this.list.Opt("-Redraw")
            ; Updating existing items retains the list's native scroll/focus
            ; state. Deleting and rebuilding them moves the next click's row.
            existing := this.list.GetCount()
            for index, row in this.rows {
                display := []
                for column, value in row
                    display.Push(this.Display(value, column))
                if index <= existing
                    this.list.Modify(index, "", display*)
                else
                    this.list.Add(, display*)
            }
            if this.append {
                display := []
                for column in this.columns
                    display.Push("")
                if this.RowCount <= existing
                    this.list.Modify(this.RowCount, "", display*)
                else
                    this.list.Add(, display*)
            }
            while this.list.GetCount() > this.RowCount
                this.list.Delete(this.list.GetCount())
        } finally {
            this.list.Opt("+Redraw")
            this.painting := false
        }
    }
    Select(row, column, extend := false, toggle := false, dragPrevious?) {
        row := Max(1, Min(this.RowCount, row)), column := Max(1, Min(this.columns.Length, column))
        if !this.RowCount
            return
        this.row := row, this.column := column
        if extend {
            this.selected := Map()
            loop Abs(row - this.anchor[1]) + 1 {
                r := Min(row, this.anchor[1]) + A_Index - 1
                loop Abs(column - this.anchor[2]) + 1 {
                    c := Min(column, this.anchor[2]) + A_Index - 1
                    this.selected[r ":" c] := [r, c]
                }
            }
            if IsSet(dragPrevious) {
                range := this.selected
                this.selected := dragPrevious.Clone()
                for key, value in range
                    if this.selected.Has(key)
                        this.selected.Delete(key)
                    else
                        this.selected[key] := value
            }
        } else {
            if !toggle
                this.selected := Map()
            key := row ":" column
            if toggle && this.selected.Has(key)
                this.selected.Delete(key)
            else
                this.selected[key] := [row, column]
            this.anchor := [row, column]
        }
        this.painting := true
        try {
            if this.list.GetNext() != row || this.list.GetNext(0, "F") != row {
                this.list.Modify(0, "-Select")
                this.list.Modify(row, "Select Focus Vis")
            }
        } finally this.painting := false
        this.RedrawCells()
        this.NotifySelection()
    }
    NotifySelection() {
        if this.options.HasOwnProp("selection")
            this.options.selection.Call()
    }
    NativeSelection(control, row, selected) {
        if !this.painting && selected && row > 0 && !GetKeyState("Ctrl", "P") && !GetKeyState("Shift", "P") {
            this.row := row
            this.selected := Map(row ":" this.column, [row, this.column])
            this.anchor := [row, this.column]
            this.NotifySelection()
        }
    }
    ColumnSelection(control, column) {
        this.CommitEditor()
        this.selected := Map(), this.column := column
        loop this.rows.Length
            this.selected[A_Index ":" column] := [A_Index, column]
        this.row := 1, this.anchor := [1, column]
        this.RedrawCells()
    }
    CellAt(x, y) {
        if !this.RowCount
            return 0
        hit := Buffer(24, 0)
        NumPut("Int", x, "Int", y, hit)
        row := SendMessage(0x1039, 0, hit.Ptr, this.list)
        column := NumGet(hit, 16, "Int") + 1
        actual := row >= 0
        if !actual {
            first := SendMessage(0x1027, 0, 0, this.list)
            rect := Buffer(16, 0)
            if !SendMessage(0x100E, first, rect.Ptr, this.list)
                return 0
            rowHeight := Max(1, NumGet(rect, 12, "Int") - NumGet(rect, 4, "Int"))
            row := first + Floor((y - NumGet(rect, 4, "Int")) / rowHeight)
        }
        position := x + DllCall("GetScrollPos", "Ptr", this.list.Hwnd, "Int", 0, "Int")
        column := 1
        loop this.columns.Length {
            position -= SendMessage(0x101D, A_Index - 1, 0, this.list)
            column := A_Index
            if position < 0
                break
        }
        return {row: Max(1, Min(this.RowCount, row + 1)), column: column, actual: actual}
    }
    Pointer(wParam, lParam, message, hwnd) {
        if this.disposed || hwnd != this.hwnd
            return
        if message != 0x201 && !IsObject(this.drag) && !this.pendingPointer.Length && !this.processingPointer && !this.pointerDown
            return
        ; System mouse messages can interrupt a commit even in a Critical
        ; thread. Capture them quickly; only the buffered private message may
        ; change selection, commit values or create an editor.
        previousCritical := A_IsCritical
        Critical(true)
        try {
            this.pendingPointer.Push([wParam, lParam, message, hwnd])
            ; Keep releases outside the list routed here while work is queued.
            if message = 0x201 {
                this.pointerDown := true
                DllCall("SetCapture", "Ptr", this.hwnd, "Ptr")
            } else if message = 0x202
                this.pointerDown := false
            PostMessage(SettingsGrid.PointerMessage, 0, 0, this.list)
        } finally Critical(previousCritical)
        return 0
    }
    DispatchPointer(wParam, lParam, message, hwnd) {
        Critical(true)
        if this.disposed || hwnd != this.hwnd
            return
        this.processingPointer := true
        try {
            while this.pendingPointer.Length && !this.disposed
                this.ApplyPointer(this.pendingPointer.RemoveAt(1)*)
        } finally this.processingPointer := false
        return 0
    }
    ApplyPointer(wParam, lParam, message, hwnd) {
        if this.disposed || hwnd != this.list.Hwnd
            return
        x := (lParam & 0xFFFF) << 48 >> 48
        y := ((lParam >> 16) & 0xFFFF) << 48 >> 48
        if message = 0x201 {
            this.CommitEditor()
            cell := this.CellAt(x, y)
            if !IsObject(cell)
                return 0
            extend := !!(wParam & 4), toggle := !!(wParam & 8)
            previous := this.selected.Clone()
            this.Select(cell.row, cell.column, extend, toggle)
            this.list.Focus()
            this.drag := {x: x, y: y, start: [cell.row, cell.column], previous: previous,
                extend: extend, toggle: toggle, moved: false, actual: cell.actual, endRow: 0, endColumn: 0}
            if this.pointerDown
                DllCall("SetCapture", "Ptr", this.list.Hwnd, "Ptr")
            return 0
        }
        if message = 0x202 {
            drag := this.drag
            this.drag := 0
            ; A newer raw press may already be queued. Its release must still
            ; reach this list while opening the older click's editor pumps UI.
            if !this.pointerDown && DllCall("GetCapture", "Ptr") = this.hwnd
                DllCall("ReleaseCapture")
            if !IsObject(drag)
                return 0
            if !drag.moved && !drag.extend && !drag.toggle && drag.actual {
                if this.columns[this.column].kind = "bool" && this.row <= this.rows.Length && !this.readonly
                    this.SetCells([[this.row, this.column, !this.rows[this.row][this.column]]])
                else
                    this.BeginEdit()
            }
            return 0
        }
        if !IsObject(this.drag)
            return
        if message = 0x200 && wParam & 1 {
            drag := this.drag
            if Abs(x - drag.x) + Abs(y - drag.y) < 4 && !drag.moved
                return 0
            drag.moved := true
            cell := this.CellAt(x, y)
            if IsObject(cell) {
                if drag.endRow = cell.row && drag.endColumn = cell.column
                    return 0
                drag.endRow := cell.row, drag.endColumn := cell.column
                this.anchor := drag.start
                if drag.extend
                    this.Select(cell.row, cell.column, true)
                else
                    this.Select(cell.row, cell.column, true, false, drag.previous)
            }
            return 0
        }
    }

    Click(control, data) {
        row := NumGet(data, A_PtrSize * 3, "Int") + 1
        column := NumGet(data, A_PtrSize * 3 + 4, "Int") + 1
        if row < 1 || column < 1 || column > this.columns.Length
            return
        this.CommitEditor()
        extend := GetKeyState("Shift", "P"), toggle := GetKeyState("Ctrl", "P")
        this.Select(row, column, extend, toggle)
        if !extend && !toggle {
            if this.columns[column].kind = "bool" && row <= this.rows.Length && !this.readonly
                this.SetCells([[row, column, !this.rows[row][column]]])
            else
                this.BeginEdit()
        }
        return 0
    }
    Draw(control, data) {
        if this.disposed
            return
        stage := NumGet(data, A_PtrSize * 3, "UInt")
        if stage = 1 || stage = 0x10001
            return 0x20
        if stage != 0x30001
            return 0
        row := NumGet(data, A_PtrSize = 8 ? 56 : 36, "UPtr") + 1
        column := NumGet(data, A_PtrSize = 8 ? 88 : 56, "Int") + 1
        if column > this.columns.Length
            return 0
        selected := this.selected.Has(row ":" column)
        background := selected ? 0xFBE4CC : row > this.rows.Length ? 0xF2F6F8 : Mod(row, 2) ? 0xFFFFFF : 0xFAF9F7
        foreground := this.Editable(column) ? 0x202020 : 0x76706A
        if this.app.theme.dark {
            background := SettingsTheme.BGR(selected ? "304C70" : row > this.rows.Length ? "27313D" : Mod(row, 2) ? "2B2E33" : "25282D")
            foreground := SettingsTheme.BGR(this.Editable(column) ? this.app.theme.foreground : this.app.theme.muted)
        }
        if this.app.tableErrors.Has(this.key) && RegExMatch(this.app.tableErrors[this.key].message, "i)^Row (\d+)", &problem) && problem[1] = row
            background := this.app.theme.dark ? 0x303D65 : 0xE6E6FF, foreground := this.app.theme.dark ? 0xCDCBFF : 0x1823B4
        NumPut("UInt", foreground, data, A_PtrSize = 8 ? 80 : 48)
        NumPut("UInt", background, data, A_PtrSize = 8 ? 84 : 52)
        if row <= this.rows.Length && this.columns[column].kind = "color" {
            color := this.ColorPreview(this.rows[row][column])
            if color != "" {
                this.DrawColor(data, row, column, color, background, foreground)
                return 4 ; CDRF_SKIPDEFAULT: swatch plus readable hex text.
            }
        }
        value := row <= this.rows.Length ? this.Display(this.rows[row][column], column) : ""
        if SettingsSearchText.Matches(value, this.app.theme.Query(this.list)).Length {
            this.DrawSearchCell(data, row, column, value, background, foreground)
            return 4
        }
        stateOffset := A_PtrSize = 8 ? 64 : 40
        NumPut("UInt", NumGet(data, stateOffset, "UInt") & ~0x11, data, stateOffset)
        return 2
    }
    ColorPreview(value) {
        if Trim(value) = ""
            return ""
        try return SettingsValue.Color(value)
        catch
            return ""
    }
    DrawSearchCell(data, row, column, value, background, foreground) {
        dc := NumGet(data, A_PtrSize * 4, "Ptr")
        rect := Buffer(16, 0)
        NumPut("Int", 0, "Int", column - 1, rect)
        if !SendMessage(0x1038, row - 1, rect.Ptr, this.list)
            return
        if column = 1
            NumPut("Int", NumGet(rect, 0, "Int") + SendMessage(0x101D, 0, 0, this.list), rect, 8)
        saved := DllCall("SaveDC", "Ptr", dc, "Int")
        brush := DllCall("CreateSolidBrush", "UInt", background, "Ptr")
        DllCall("FillRect", "Ptr", dc, "Ptr", rect, "Ptr", brush)
        DllCall("DeleteObject", "Ptr", brush)
        NumPut("Int", NumGet(rect, 0, "Int") + 6, rect)
        DllCall("SetTextColor", "Ptr", dc, "UInt", foreground)
        DllCall("SetBkMode", "Ptr", dc, "Int", 1)
        DllCall("SelectObject", "Ptr", dc, "Ptr", SendMessage(0x31, 0, 0, this.list), "Ptr")
        this.app.theme.Text(dc, value, rect, 0x8824, this.list)
        DllCall("RestoreDC", "Ptr", dc, "Int", saved)
    }
    DrawColor(data, row, column, color, background, foreground) {
        dc := NumGet(data, A_PtrSize * 4, "Ptr")
        rect := Buffer(16, 0)
        NumPut("Int", 0, "Int", column - 1, rect)
        if !SendMessage(0x1038, row - 1, rect.Ptr, this.list)
            return
        saved := DllCall("SaveDC", "Ptr", dc, "Int")
        brush := DllCall("CreateSolidBrush", "UInt", background, "Ptr")
        DllCall("FillRect", "Ptr", dc, "Ptr", rect, "Ptr", brush)
        DllCall("DeleteObject", "Ptr", brush)
        size := Min(20, NumGet(rect, 12, "Int") - NumGet(rect, 4, "Int") - 6)
        left := NumGet(rect, 0, "Int") + 6, top := NumGet(rect, 4, "Int") + 3
        swatch := Buffer(16, 0)
        NumPut("Int", left, "Int", top, "Int", left + size, "Int", top + size, swatch)
        rgb := Integer("0x" color)
        bgr := ((rgb & 255) << 16) | (rgb & 0xFF00) | ((rgb >> 16) & 255)
        brush := DllCall("CreateSolidBrush", "UInt", bgr, "Ptr")
        DllCall("FillRect", "Ptr", dc, "Ptr", swatch, "Ptr", brush)
        DllCall("DeleteObject", "Ptr", brush)
        brush := DllCall("CreateSolidBrush", "UInt", foreground, "Ptr")
        DllCall("FrameRect", "Ptr", dc, "Ptr", swatch, "Ptr", brush)
        DllCall("DeleteObject", "Ptr", brush)
        NumPut("Int", left + size + 8, rect)
        DllCall("SetTextColor", "Ptr", dc, "UInt", foreground)
        DllCall("SetBkMode", "Ptr", dc, "Int", 1)
        DllCall("SelectObject", "Ptr", dc, "Ptr", SendMessage(0x31, 0, 0, this.list), "Ptr")
        this.app.theme.Text(dc, "#" color, rect, 0x8824, this.list)
        DllCall("RestoreDC", "Ptr", dc, "Int", saved)
    }
    BeginEdit(initial?) {
        this.openingEditor := true
        try {
            if !this.Editable(this.column) || !this.RowCount
                return
            this.CommitEditor()
            SendMessage(0x1013, this.row - 1, 0, this.list) ; LVM_ENSUREVISIBLE
            rect := Buffer(16, 0)
            NumPut("Int", 0, "Int", this.column - 1, rect)
            if !SendMessage(0x1038, this.row - 1, rect.Ptr, this.list)
                return
            DllCall("MapWindowPoints", "Ptr", this.list.Hwnd, "Ptr", this.parent.Hwnd, "Ptr", rect, "UInt", 2)
            scale := 96 / A_ScreenDPI
            x := NumGet(rect, 0, "Int") * scale, y := NumGet(rect, 4, "Int") * scale
            width := Max(50, (NumGet(rect, 8, "Int") - NumGet(rect, 0, "Int")) * scale)
            if this.column = 1 {
                ; Subitem zero's rectangle covers the row; its real width is separate.
                width := SendMessage(0x101D, 0, 0, this.list) * scale
            }
            height := Max(23, (NumGet(rect, 12, "Int") - NumGet(rect, 4, "Int")) * scale)
            value := this.row <= this.rows.Length ? this.rows[this.row][this.column] : ""
            spec := this.columns[this.column]
            options := "x" x " y" y " w" width " h" height
            this.editChoices := []
            if spec.kind = "bool"
                this.editChoices := [{label: "No", value: 0}, {label: "Yes", value: 1}]
            else if spec.kind = "gesture"
                this.editChoices := SettingsTables.GestureOptions
            else if spec.kind = "choice"
                this.editChoices := spec.choices
            if this.editChoices.Length {
                choiceHeight := Min(340, Max(180, 24 * (this.editChoices.Length + 1)))
                options := "x" x " y" y " w" width " h" choiceHeight
                labels := [], chosen := 0
                for i, option in this.editChoices {
                    labels.Push(option.label)
                    if option.value = value
                        chosen := i
                }
                if !this.HasOwnProp("choiceEditor") {
                    this.choiceEditor := this.parent.AddDropDownList(options, labels)
                    this.choiceEditor.OnEvent("Change", (*) => this.ChoiceCommitted())
                    this.choiceEditor.OnEvent("LoseFocus", ObjBindMethod(this, "EditorLostFocus"))
                } else {
                    this.choiceEditor.Delete()
                    this.choiceEditor.Add(labels)
                    this.choiceEditor.Move(x, y, width, choiceHeight)
                }
                this.editor := this.choiceEditor
                this.editor.Visible := true
                if chosen
                    this.editor.Choose(chosen)
            } else {
                if !this.HasOwnProp("textEditor") {
                    this.textEditor := this.parent.AddEdit(options)
                    this.textEditor.OnEvent("LoseFocus", ObjBindMethod(this, "EditorLostFocus"))
                    this.textEditor.OnEvent("Change", (*) => this.app.UpdateStatus())
                } else
                    this.textEditor.Move(x, y, width, height)
                this.editor := this.textEditor
                this.editor.Value := IsSet(initial) ? initial : value
                this.editor.Visible := true
            }
            this.editRow := this.row, this.editColumn := this.column
            this.editor.SettingsHelp := "Edit " spec.label ". Enter saves; Escape cancels; Tab advances. Ctrl+Enter fills selected cells."
            this.app.theme.Control(this.editor)
            this.editor.ToolTip := this.editor.SettingsHelp
            this.editor.Focus()
            if !this.editChoices.Length {
                ; Edit controls may select their initial text when gaining focus.
                ; Set the selection afterwards so type-to-replace keeps its caret.
                if IsSet(initial)
                    SendMessage(0xB1, StrLen(this.editor.Value), StrLen(this.editor.Value), this.editor)
                else
                    SendMessage(0xB1, 0, -1, this.editor)
            }
            this.Option("scroller", this.app.settingsScroller).BindFocusControls()
            this.app.UpdateStatus()
            if this.editChoices.Length {
                SendMessage(0x160, Round(Max(width, Min(620, this.app.contentWidth - 40)) * A_ScreenDPI / 96), 0, this.editor)
                SendMessage(0x14F, 1, 0, this.editor)
            }
        } finally this.openingEditor := false
    }
    ChoiceCommitted() {
        this.CommitEditor()
        if !this.disposed
            this.list.Focus()
    }
    EditorLostFocus(control, *) {
        if this.disposed || !DllCall("IsWindow", "Ptr", this.hwnd)
            return
        focus := DllCall("GetFocus", "Ptr")
        if (!this.HasOwnProp("openingEditor") || !this.openingEditor)
            && IsObject(this.editor) && this.editor = control && focus && focus != control.Hwnd && focus != this.list.Hwnd
            && !IsObject(this.drag)
            this.CommitEditor()
    }
    CancelEditor() {
        if !IsObject(this.editor)
            return
        editor := this.editor
        this.editor := 0
        editor.Visible := false
        this.app.UpdateStatus()
    }
    CommitEditor(fill := false) {
        if this.HasOwnProp("committing") && this.committing
            return false
        this.committing := true
        try {
            if this.disposed || !IsObject(this.editor)
                return true
            value := this.editChoices.Length ? (this.editor.Value ? this.editChoices[this.editor.Value].value : "") : this.editor.Value
            row := this.editRow, column := this.editColumn
            view := this.Viewport()
            this.CancelEditor()
            if (row > this.rows.Length && value = "") || (!fill && row <= this.rows.Length && value == this.rows[row][column] && !this.app.tableErrors.Has(this.key)) {
                this.RestoreViewport(view)
                return true
            }
            changes := fill ? this.SelectedChanges(value) : [[row, column, value]]
            saved := this.SetCells(changes)
            this.RestoreViewport(view)
            return saved
        } finally {
            this.committing := false
        }
    }
    Viewport() {
        if this.disposed || !DllCall("IsWindow", "Ptr", this.hwnd)
            return 0
        first := SendMessage(0x1027, 0, 0, this.list) ; LVM_GETTOPINDEX
        rect := Buffer(16, 0)
        SendMessage(0x100E, first, rect.Ptr, this.list) ; LVM_GETITEMRECT
        return {row: first, top: NumGet(rect, 4, "Int"), x: DllCall("GetScrollPos", "Ptr", this.list.Hwnd, "Int", 0, "Int")}
    }
    RestoreViewport(view) {
        if !IsObject(view) || this.disposed || !DllCall("IsWindow", "Ptr", this.hwnd) || !this.list.GetCount()
            return
        row := Min(view.row, this.list.GetCount() - 1)
        rect := Buffer(16, 0)
        if SendMessage(0x100E, row, rect.Ptr, this.list) {
            x := DllCall("GetScrollPos", "Ptr", this.list.Hwnd, "Int", 0, "Int")
            SendMessage(0x1014, view.x - x, NumGet(rect, 4, "Int") - view.top, this.list) ; LVM_SCROLL
        }
    }
    Normalize(value, column) {
        kind := this.columns[column].kind
        value := Trim(value)
        if kind = "bool" {
            if value = "" || value = "0" || StrLower(value) = "no" || StrLower(value) = "false"
                return 0
            if value = "1" || StrLower(value) = "yes" || StrLower(value) = "true"
                return 1
            throw ValueError("Use 1/Yes or 0/No for " this.columns[column].label ".")
        }
        if kind = "choice" {
            for i, option in this.columns[column].choices
                if value = option.value || value = option.label || value = i
                    return option.value
            throw ValueError("Choose a listed value for " this.columns[column].label ".")
        }
        if kind = "gesture" {
            for option in SettingsTables.GestureOptions
                if value = option.value || value = option.label
                    return option.value
            throw ValueError("Choose a listed mouse gesture.")
        }
        if kind = "color" && value != ""
            return SettingsValue.Color(value)
        if kind = "file" || kind = "sound"
            return StrLen(value) > 1 && SubStr(value, 1, 1) = Chr(34) && SubStr(value, -1) = Chr(34) ? SubStr(value, 2, StrLen(value) - 2) : value
        if kind = "hotkey" {
            SettingsValue.Hotkey(value)
            return value
        }
        if kind = "number"
            return SettingsValue.Parse({kind: "integer", min: 0, max: 300, unit: "%"}, RTrim(value, "%"))
        return value
    }
    SetCells(changes) {
        if this.disposed || !DllCall("IsWindow", "Ptr", this.hwnd) || this.readonly
            return false
        rows := SettingsModel.Clone(this.rows), failure := "", cleared := Map()
        for change in changes {
            row := change[1], column := change[2]
            if !this.append && row > rows.Length {
                this.app.UpdateStatus("This table has fixed rows; paste within its existing entries.")
                return false
            }
            if !this.Editable(column) || row < 1
                continue
            if row > 2000
                throw ValueError("Paste up to 2000 rows at a time.")
            while rows.Length < row
                rows.Push(this.Blank())
            rows[row][column] := change[3]
            if this.append && column = 1 && Trim(change[3]) = ""
                cleared[row] := true

        }
        index := rows.Length
        while index > 0 {
            if cleared.Has(index)
                rows.RemoveAt(index)
            index--
        }
        return this.SaveCandidate(rows, failure)
    }
    SaveCandidate(rows, failure := "") {
        if this.disposed || !DllCall("IsWindow", "Ptr", this.hwnd)
            return false
        view := this.Viewport()
        persisted := []
        try {
            for index, row in rows {
                if this.append && Trim(row[1]) = ""
                    continue
                persisted.Push(row)
                for column, value in row
                    if this.Editable(column) {
                        try row[column] := this.Normalize(value, column)
                        catch as err
                            throw ValueError("Row " index ", " this.columns[column].label ": " err.Message)
                    }
            }
            if failure != ""
                throw ValueError(failure)
            this.save.Call(persisted)
            this.snapshot := JSON.Dump(this.options.HasOwnProp("read") ? this.options.read.Call() : rows)
            if this.app.tableErrors.Has(this.key)
                this.app.tableErrors.Delete(this.key)
        } catch as err {
            failure := err.Message
            this.app.tableErrors[this.key] := {profile: this.profile, section: this.section, message: failure}
        }
        if this.disposed || !DllCall("IsWindow", "Ptr", this.hwnd)
            return failure = ""
        this.list.ToolTip := failure != "" ? failure : "Click a cell to edit. Shift/Ctrl selects cells. Ctrl+C/V copies/pastes; Del clears; Ctrl+Del removes rows."
        this.list.SettingsHelp := this.list.ToolTip
        this.rows := rows
        this.app.tableStates[this.key] := {rows: SettingsModel.Clone(rows), snapshot: this.snapshot, group: this.group}
        this.Paint()
        if this.RowCount {
            this.painting := true
            try this.list.Modify(Min(this.row, this.RowCount), "Select Focus")
            finally this.painting := false
        }
        this.RestoreViewport(view)
        this.NotifySelection()
        this.app.Changed()
        if failure != ""
            this.app.UpdateStatus(failure)
        return failure = ""
    }
    Reset(rows, select := 0) {
        this.CancelEditor()
        this.rows := SettingsModel.Clone(rows), this.snapshot := JSON.Dump(rows)
        if this.app.tableStates.Has(this.key)
            this.app.tableStates.Delete(this.key)
        if this.app.tableErrors.Has(this.key)
            this.app.tableErrors.Delete(this.key)
        this.Paint()
        if select
            this.Select(Min(this.RowCount, select), this.column)
    }
    SelectedChanges(value) {
        changes := []
        if !this.selected.Count
            this.Select(this.row, this.column)
        for key, cell in this.selected
            changes.Push([cell[1], cell[2], value])
        return changes
    }
    Copy() {
        if !this.selected.Count
            this.Select(this.row, this.column)
        minRow := 2001, maxRow := 0, minCol := this.columns.Length, maxCol := 1
        for key, cell in this.selected
            minRow := Min(minRow, cell[1]), maxRow := Max(maxRow, cell[1]), minCol := Min(minCol, cell[2]), maxCol := Max(maxCol, cell[2])
        text := ""
        loop maxRow - minRow + 1 {
            row := minRow + A_Index - 1
            if A_Index > 1
                text .= "`r`n"
            loop maxCol - minCol + 1 {
                column := minCol + A_Index - 1
                if A_Index > 1
                    text .= "`t"
                if row <= this.rows.Length && this.selected.Has(row ":" column)
                    text .= this.rows[row][column]
            }
        }
        A_Clipboard := text
        return text
    }
    Paste(text?) {
        if !IsSet(text)
            text := A_Clipboard
        text := RTrim(StrReplace(text, "`r"), "`n")
        if text = ""
            return true
        if !InStr(text, "`t") && !InStr(text, "`n") && this.selected.Count > 1
            return this.SetCells(this.SelectedChanges(text))
        changes := [], originRow := this.row, originCol := this.column
        if this.selected.Count > 1
            for key, cell in this.selected
                originRow := Min(originRow, cell[1]), originCol := Min(originCol, cell[2])
        for r, line in StrSplit(text, "`n")
            for c, value in StrSplit(line, "`t")
                if originCol + c - 1 <= this.columns.Length
                    changes.Push([originRow + r - 1, originCol + c - 1, value])
        return this.SetCells(changes)
    }
    RemoveRows() {
        if this.readonly || !this.options.HasOwnProp("remove")
            return
        this.CommitEditor()
        selected := Map()
        for key, cell in this.selected
            if cell[1] <= this.rows.Length
                selected[cell[1]] := true
        if !selected.Count && this.row <= this.rows.Length
            selected[this.row] := true
        this.options.remove.Call(selected)
    }
    AddRow() {
        this.CommitEditor()
        this.Select(this.rows.Length + 1, 1)
        this.BeginEdit()
    }
    HandleKey(key, lParam, message, hwnd) {
        if this.disposed
            return
        editing := IsObject(this.editor) && hwnd = this.editor.Hwnd
        if hwnd != this.list.Hwnd && !editing
            return
        ctrl := GetKeyState("Ctrl", "P"), shift := GetKeyState("Shift", "P")
        if editing {
            if key = 27 {
                this.CancelEditor(), this.list.Focus()
                return 0
            }
            if key = 13 || key = 9 {
                this.CommitEditor(ctrl)
                this.list.Focus()
                if key = 9
                    this.Advance(shift ? -1 : 1)
                return 0
            }
            return
        }
        if ctrl && key = 65 {
            this.selected := Map()
            loop this.rows.Length {
                row := A_Index
                loop this.columns.Length
                    this.selected[row ":" A_Index] := [row, A_Index]
            }
            this.RedrawCells()
        } else if ctrl && key = 67
            this.Copy()
        else if ctrl && key = 86
            this.Paste()
        else if key = 46 {
            if ctrl
                this.RemoveRows()
            else
                this.SetCells(this.SelectedChanges(""))
        } else if key = 113 || key = 13
            this.BeginEdit()
        else if key = 9
            this.Advance(shift ? -1 : 1)
        else if key >= 37 && key <= 40 {
            this.Select(this.row + (key = 38 ? -1 : key = 40 ? 1 : 0), this.column + (key = 37 ? -1 : key = 39 ? 1 : 0), shift)
            this.Option("scroller", this.app.settingsScroller).EnsureVisible(this.list)
        } else
            return
        return 0
    }
    Advance(direction) {
        row := this.row, column := this.column
        loop this.columns.Length * 2 {
            column += direction
            if column > this.columns.Length
                row++, column := 1
            else if column < 1
                row--, column := this.columns.Length
            if row < 1 || row > this.RowCount
                break
            if this.Editable(column) {
                this.Select(row, column)
                this.BeginEdit()
                return
            }
        }
    }
    Character(key, lParam, message, hwnd) {
        if this.disposed || hwnd != this.list.Hwnd || key < 32 || GetKeyState("Ctrl", "P") || GetKeyState("Alt", "P")
            return
        if key = 32 && this.columns[this.column].kind = "bool" && this.row <= this.rows.Length
            this.SetCells(this.SelectedChanges(!this.rows[this.row][this.column]))
        else if this.selected.Count > 1 || this.columns[this.column].kind = "bool" || this.columns[this.column].kind = "choice" || this.columns[this.column].kind = "gesture"
            this.SetCells(this.SelectedChanges(Chr(key)))
        else
            this.BeginEdit(Chr(key))
        return 0
    }
    Scrolled(wParam, lParam, message, hwnd) {
        ; Pointer owns direct list clicks. Committing here can dispatch their
        ; queued button-up before Pointer has established the matching drag.
        if message = 0x201 && hwnd = this.hwnd
            return
        if !this.disposed && IsObject(this.editor) && (hwnd = this.list.Hwnd || lParam = this.list.Hwnd || DllCall("IsChild", "Ptr", this.list.Hwnd, "Ptr", hwnd))
            this.CommitEditor()
    }
    Dispose() {
        if this.disposed
            return
        this.disposed := true
        this.pendingPointer := []
        this.pointerDown := false
        if IsObject(this.drag) || DllCall("GetCapture", "Ptr") = this.hwnd {
            DllCall("ReleaseCapture")
            this.drag := 0
        }
        this.CancelEditor()
        OnMessage(0x100, this.keyCallback, 0), OnMessage(0x102, this.charCallback, 0)
        for message in [0x20A, 0x115, 0x114, 0x201]
            OnMessage(message, this.scrollCallback, 0)
        for message in [0x201, 0x200, 0x202]
            OnMessage(message, this.pointerCallback, 0)
        OnMessage(SettingsGrid.PointerMessage, this.pointerDispatchCallback, 0)
        try this.list.OnNotify(-2, this.clickCallback, 0)
        try this.list.OnNotify(-12, this.drawCallback, 0)
    }
}
