#Include "HotkeyBindings.ahk"
#Include "HotkeyCapture.ahk"

class HotkeySettingsGui {
    InitHotkeyHelpers() {
        this.hotkeyFields := []
        this.hotkeyRefreshTimer := ObjBindMethod(this, "RefreshHotkeyWarnings")
        this.AddHotkeyHelper(this.Sidebar, "ProfileActivationHotkey")
        for name, label in HotkeyBindings.Settings
            this.AddHotkeyHelper(this.MainFrame, name, "Hotkeys Settings")
        for name in ["ForwardsKey", "BackwardsdKey", "FirstCharHotkey"]
            this.AddHotkeyHelper(this.MainFrame, name, "Hotkey Groups")
        this.AddHotkeyHelper(this.MainFrame, "HotkeyList", "Hotkeys Settings", true)
        this.AddHotkeyHelper(this.MainFrame, "QuickGroupHotkey", "Thumbnails Interactions")
        for name in ["NonEVEForwardsKey", "NonEVEBackwardsdKey"]
            this.AddHotkeyHelper(this.MainFrame, name, "Non-EVE Applications")
        this.AddHotkeyHelper(this.MainFrame, "NonEVEHotkeysHtks", "Non-EVE Applications", true)
        for name in ["HotkeyGroupDDL", "NonEVEGroupsDDL"]
            this.MainFrame[name].OnEvent("Change", (*) => this.ScheduleHotkeyWarnings())
        for name in ["HotkeyCharList", "NonEVEProcessHtks", "NonEVETitlesHtks"]
            this.MainFrame[name].OnEvent("Change", (*) => this.ScheduleHotkeyWarnings())
        for group in ["Hotkey Groups", "Hotkeys Settings", "Non-EVE Applications"] {
            for control in this.MainFrame.Group[group]
                if control.Type = "Button"
                    control.OnEvent("Click", (*) => this.ScheduleHotkeyWarnings())
        }
        this.hotkeyTooltipCallback := ObjBindMethod(this, "On_WM_MOUSEMOVE")
        OnMessage(0x200, this.hotkeyTooltipCallback)
    }

    AddHotkeyHelper(gui, name, group := "", multiline := false) {
        field := gui[name]
        field.GetPos(&x, &y, &w, &h)
        buttonW := 24, warningW := 16, gap := 2
        field.Move(,, w - buttonW - warningW - gap * 2)
        button := gui.Add("Button", Format("x{} y{} w{} h{}", x + w - buttonW - warningW - gap - 1, y - 1, buttonW + 2, Min(h, 25) + 2), Chr(0x2328))
        button.SetFont("s11", "Segoe UI Symbol")
        button.ToolTip := multiline ? "Record a hotkey in the selected line (click that line first)." : "Record a hotkey."
        warning := gui.Add("Text", Format("x{} y{} w{} h{} Center +0x200 cB36B00", x + w - warningW, y, warningW, Min(h, 25)), "")
        warning.SetFont("s11 bold", "Segoe UI")
        field.GetPos(&fx, &fy, &fw, &fh)
        waiting := gui.Add("Text", Format("x{} y{} w{} h{} Hidden Border Center +0x300 BackgroundE0E0E0 c666666", fx, fy, fw, fh), "")
        waiting.SetFont("s9", "Segoe UI")
        waiting.ToolTip := "Press and release a shortcut, or use a mouse button or the wheel here. Esc or the record button cancels."
        helper := {field: field, button: button, warning: warning, waiting: waiting, multiline: multiline, message: ""}
        this.hotkeyFields.Push(helper)
        button.OnEvent("Click", (*) => this.RecordHotkey(helper))
        warning.OnEvent("Click", (*) => this.ShowHotkeyConflicts(helper))
        field.OnEvent("Change", (*) => this.ScheduleHotkeyWarnings())
        if group != "" {
            this.MainFrame.Group[group].Push(button, warning)
            button.Visible := field.Visible
            warning.Visible := field.Visible
        }
    }

    ScheduleHotkeyWarnings(*) {
        if this.HasOwnProp("hotkeyRefreshTimer")
            SetTimer(this.hotkeyRefreshTimer, -100)
    }

    StopHotkeyHelpers() {
        this.CancelHotkeyRecording()
        if this.HasOwnProp("hotkeyRefreshTimer")
            SetTimer(this.hotkeyRefreshTimer, 0)
        if this.HasOwnProp("hotkeyTooltipCallback")
            OnMessage(0x200, this.hotkeyTooltipCallback, 0)
        this.ClearSettingsTooltip()
    }

    HotkeyFieldIds(helper) {
        name := helper.field.Name
        if name = "ProfileActivationHotkey"
            return ["profile:" this.LastUsedProfile]
        if HotkeyBindings.Settings.Has(name)
            return ["setting:" name]
        if name = "QuickGroupHotkey"
            return ["quick"]
        groups := Map("ForwardsKey", "ForwardsHotkey", "BackwardsdKey", "BackwardsHotkey", "FirstCharHotkey", "FirstCharHotkey")
        if groups.Has(name)
            return ["group:" this.MainFrame["HotkeyGroupDDL"].Text ":" groups[name]]
        if name = "NonEVEForwardsKey" || name = "NonEVEBackwardsdKey"
            return ["nonGroup:" this.MainFrame["NonEVEGroupsDDL"].Text ":" (name = "NonEVEForwardsKey" ? "fkey" : "bkey")]
        ids := []
        if name = "HotkeyList" {
            storedRow := 0
            for row, character in StrSplit(this.MainFrame["HotkeyCharList"].Value, "`n") {
                if Trim(character, " `t`r") != "" {
                    storedRow++
                    ids.Push({id: "character:" storedRow, row: row})
                }
            }
        } else {
            for row, key in StrSplit(helper.field.Value, "`n")
                ids.Push({id: "nonApp:" row, row: row})
        }
        return ids
    }

    RefreshHotkeyWarnings(*) {
        if !this.HasOwnProp("hotkeyFields")
            return
        conflicts := HotkeyBindings.Conflicts(HotkeyBindings.Collect(this))
        for helper in this.hotkeyFields {
            helper.button.Enabled := helper.field.Enabled
            details := ""
            for value in this.HotkeyFieldIds(helper) {
                id := IsObject(value) ? value.id : value
                if !conflicts.Has(id)
                    continue
                if IsObject(value)
                    details .= "Line " value.row ":`n"
                for other in conflicts[id]
                    details .= "  " other "`n"
            }
            helper.message := details = "" ? "" : "This binding is also assigned to:`n" details
                . "`nRepeated hotkeys may cause unpredictable or inconsistent behavior. Saving and importing are still allowed."
            helper.warning.Text := details = "" ? "" : "!"
            helper.warning.ToolTip := helper.message
        }
    }

    ShowHotkeyConflicts(helper) {
        if helper.message != ""
            MsgBox(helper.message, "Repeated hotkey", "Icon! Owner" this.S_Gui.Hwnd)
    }

    RecordHotkey(helper) {
        if HotkeyCapture.Active {
            sameField := HotkeyCapture.Current.helper = helper
            HotkeyCapture.Current.Cancel()
            if sameField
                return
        }
        if !helper.field.Enabled
            return
        row := 1
        if helper.multiline {
            start := Buffer(4, 0)
            SendMessage(0xB0, start.Ptr, 0, helper.field) ; EM_GETSEL
            row := SendMessage(0xC9, NumGet(start, 0, "UInt"), 0, helper.field) + 1 ; EM_LINEFROMCHAR
        }
        HotkeyCapture().Start(this.S_Gui, helper, ObjBindMethod(this, "RecordedHotkey", helper, row), row)
    }

    RecordedHotkey(helper, row, key) {
        if key != ""
            this.ApplyRecordedHotkey(helper, key, row)
    }

    CancelHotkeyRecording() {
        if HotkeyCapture.Active
            HotkeyCapture.Current.Cancel()
    }

    ApplyRecordedHotkey(helper, key, row := 1) {
        field := helper.field
        field.Value := helper.multiline ? HotkeyBindings.ReplaceLine(field.Value, row, key) : key
        ; Notify the existing save handler, including for multiline WM_SETTEXT.
        id := DllCall("GetDlgCtrlID", "Ptr", field.Hwnd, "Int")
        PostMessage(0x111, (0x300 << 16) | (id & 0xFFFF), field.Hwnd,, "ahk_id " field.Gui.Hwnd)
        field.Focus()
        this.ScheduleHotkeyWarnings()
    }
}
