; Keep control coordinates stable by scrolling their child GUI inside a clipped
; viewport. All ranges and offsets below are physical pixels (Win32 units).
class SettingsScrollPane {
    offset := 0
    extent := 0
    page := 0
    maximum := 0
    stopped := false
    focusControls := []
    wheelRemainder := 0
    inset := 0
    padding := 16

    __New(owner) {
        this.ownerHwnd := owner.Hwnd
        this.viewport := Gui("+Parent" owner.Hwnd " -Caption -Border +0x02200000 +E0x10000") ; CLIPCHILDREN | VSCROLL
        this.viewport.BackColor := "F7F7F7"
        this.viewportHwnd := this.viewport.Hwnd
        this.content := Gui("+Parent" this.viewportHwnd " -Caption -Border +0x02000000 +E0x10000")
        this.content.BackColor := "F7F7F7"
        this.contentHwnd := this.content.Hwnd
        this.header := Gui("+Parent" owner.Hwnd " -Caption -Border +0x02000000")
        this.header.BackColor := "F7F7F7"
        this.header.SetFont("s12 w700 q5")
        this.heading := this.header.AddText("x16 y16 w410 h24", "")
        this.rule := this.header.AddText("x16 y40 w410 h2 +0x10", "")
        this.headerHwnd := this.header.Hwnd
        this.focusCallback := ObjBindMethod(this, "EnsureVisible")
        this.callbacks := Map(0x115, ObjBindMethod(this, "OnScroll"),
            0x20A, ObjBindMethod(this, "OnWheel"), 0x2, ObjBindMethod(this, "OnDestroy"))
        for message, callback in this.callbacks
            OnMessage(message, callback)
    }

    Show(x, y, width, height) {
        this.x := x, this.y := y
        this.width := Round(width * A_ScreenDPI / 96)
        this.height := Round(height * A_ScreenDPI / 96)
        this.Layout()
        this.UpdateRange()
    }

    Layout() {
        if !this.HasOwnProp("width")
            return
        ; Gui.Show w/h describe CLIENT dimensions; a native scrollbar adds to
        ; the outer width. Position the OUTER rectangle explicitly to avoid it
        ; being clipped by the settings window's right edge.
        DllCall("SetWindowPos", "Ptr", this.viewportHwnd, "Ptr", 0,
            "Int", this.x, "Int", this.y + this.inset, "Int", this.width,
            "Int", Max(1, this.height - this.inset), "UInt", 0x54)
        if this.inset {
            this.heading.Value := this.title
            this.heading.Move(, , this.width * 96 / A_ScreenDPI - 32)
            this.rule.Move(, , this.width * 96 / A_ScreenDPI - 20)
            DllCall("SetWindowPos", "Ptr", this.headerHwnd, "Ptr", 0,
                "Int", this.x, "Int", this.y, "Int", this.width, "Int", this.inset, "UInt", 0x54)
        } else
            this.header.Hide()
    }

    Select(controls) {
        this.offset := 0
        this.wheelRemainder := 0
        this.extent := 0
        pinned := controls.Length >= 2 && controls[1].Type = "Text" && controls[2].Type = "Text"
            && (WinGetStyle(controls[2].Hwnd) & 0x1F) = 0x10
        this.inset := pinned ? Round(48 * A_ScreenDPI / 96) : 0
        this.title := pinned ? controls[1].Text : ""
        for control in controls {
            ; Map real rectangles instead of assuming system DPI, including when
            ; Windows scales a settings window on another monitor.
            rect := Buffer(16, 0)
            DllCall("GetWindowRect", "Ptr", control.Hwnd, "Ptr", rect)
            DllCall("MapWindowPoints", "Ptr", 0, "Ptr", this.contentHwnd, "Ptr", rect, "UInt", 2)
            this.extent := Max(this.extent, NumGet(rect, 12, "Int"))
        }
        this.extent := Max(0, this.extent - this.inset) + Round(this.padding * A_ScreenDPI / 96)
        this.Layout()
        this.UpdateRange()
        this.BindFocusControls()
    }

    BindFocusControls() {
        ; Helpers added after the page controls (recorders, warnings and color
        ; pickers) share this GUI and need the same keyboard visibility handling.
        for hwnd, control in this.content {
            if control.HasOwnProp("SettingsScrollFocusBound")
                continue
            if !InStr("|Edit|Button|CheckBox|Radio|DDL|ComboBox|ListBox|DateTime|MonthCal|UpDown|ListView|TreeView|", "|" control.Type "|")
                continue
            control.OnEvent("Focus", this.focusCallback)
            control.SettingsScrollFocusBound := true
            this.focusControls.Push(control)
        }
    }

    UpdateRange() {
        if this.stopped
            return
        rect := Buffer(16, 0)
        DllCall("GetClientRect", "Ptr", this.viewportHwnd, "Ptr", rect)
        this.page := Max(1, NumGet(rect, 12, "Int"))
        this.maximum := Max(0, this.extent - this.page)
        this.offset := Min(this.offset, this.maximum)
        info := Buffer(28, 0)
        NumPut("UInt", 28, "UInt", 7, "Int", 0, "Int", Max(0, this.extent - 1),
            "UInt", this.page, "Int", this.offset, info)
        DllCall("SetScrollInfo", "Ptr", this.viewportHwnd, "Int", 1, "Ptr", info, "Int", true)
        ; Showing/hiding the bar changes the viewport's available client width.
        DllCall("GetClientRect", "Ptr", this.viewportHwnd, "Ptr", rect)
        DllCall("SetWindowPos", "Ptr", this.contentHwnd, "Ptr", 0, "Int", 0, "Int", -this.offset - this.inset,
            "Int", NumGet(rect, 8, "Int"), "Int", Max(this.page, this.extent) + this.inset, "UInt", 0x54)
    }

    ScrollTo(position) {
        if this.stopped
            return
        this.offset := Max(0, Min(this.maximum, Round(position)))
        info := Buffer(28, 0)
        NumPut("UInt", 28, "UInt", 4, info)
        NumPut("Int", this.offset, info, 20)
        DllCall("SetScrollInfo", "Ptr", this.viewportHwnd, "Int", 1, "Ptr", info, "Int", true)
        DllCall("SetWindowPos", "Ptr", this.contentHwnd, "Ptr", 0, "Int", 0, "Int", -this.offset - this.inset,
            "Int", 0, "Int", 0, "UInt", 0x15) ; NOACTIVATE | NOZORDER | NOSIZE
    }

    OnScroll(wParam, lParam, message, hwnd) {
        if this.stopped || hwnd != this.viewportHwnd || lParam
            return
        line := Round(32 * A_ScreenDPI / 96)
        switch wParam & 0xFFFF {
            case 0: this.ScrollTo(this.offset - line)
            case 1: this.ScrollTo(this.offset + line)
            case 2: this.ScrollTo(this.offset - this.page + line)
            case 3: this.ScrollTo(this.offset + this.page - line)
            case 4, 5:
                info := Buffer(28, 0)
                NumPut("UInt", 28, "UInt", 0x10, info)
                DllCall("GetScrollInfo", "Ptr", hwnd, "Int", 1, "Ptr", info)
                this.ScrollTo(NumGet(info, 24, "Int")) ; 32-bit thumb tracking position
            case 6: this.ScrollTo(0)
            case 7: this.ScrollTo(this.maximum)
        }
        return 0
    }

    OnWheel(wParam, lParam, message, hwnd) {
        if this.stopped || !this.maximum || (hwnd != this.ownerHwnd && !DllCall("IsChild", "Ptr", this.ownerHwnd, "Ptr", hwnd))
            return
        x := (lParam & 0xFFFF) << 48 >> 48
        y := ((lParam >> 16) & 0xFFFF) << 48 >> 48
        rect := Buffer(16, 0)
        DllCall("GetWindowRect", "Ptr", this.viewportHwnd, "Ptr", rect)
        if x < NumGet(rect, 0, "Int") || x >= NumGet(rect, 8, "Int")
            || y < NumGet(rect, 4, "Int") || y >= NumGet(rect, 12, "Int")
            return
        target := DllCall("WindowFromPoint", "Int64", (x & 0xFFFFFFFF) | (y << 32), "Ptr")
        if this.NativeWheelTarget(target)
            return ; Multiline edits, lists and open combo boxes keep their own input.
        delta := (wParam >> 16) & 0xFFFF
        if delta & 0x8000
            delta -= 0x10000
        lines := 3
        DllCall("SystemParametersInfoW", "UInt", 0x68, "UInt", 0, "UInt*", &lines, "UInt", 0)
        this.wheelRemainder += delta
        steps := Integer(this.wheelRemainder / 120)
        this.wheelRemainder -= steps * 120
        distance := lines = 0xFFFFFFFF ? this.page : lines * Round(32 * A_ScreenDPI / 96)
        if steps
            this.ScrollTo(this.offset - steps * distance)
        return 0
    }

    NativeWheelTarget(hwnd) {
        control := GuiCtrlFromHwnd(hwnd)
        if !control
            return false
        if control.Type = "DDL" || control.Type = "ComboBox"
            return !!SendMessage(0x157, 0, 0, control) ; CB_GETDROPPEDSTATE
        if control.Type != "Edit" && control.Type != "ListBox" && control.Type != "ListView" && control.Type != "TreeView"
            return false
        info := Buffer(28, 0)
        NumPut("UInt", 28, "UInt", 3, info)
        return DllCall("GetScrollInfo", "Ptr", hwnd, "Int", 1, "Ptr", info)
            && NumGet(info, 12, "Int") >= NumGet(info, 16, "UInt")
    }

    EnsureVisible(control, *) {
        if this.stopped || !control.Visible || !this.maximum
            return
        rect := Buffer(16, 0)
        DllCall("GetWindowRect", "Ptr", control.Hwnd, "Ptr", rect)
        DllCall("MapWindowPoints", "Ptr", 0, "Ptr", this.contentHwnd, "Ptr", rect, "UInt", 2)
        top := NumGet(rect, 4, "Int") - this.inset, bottom := NumGet(rect, 12, "Int") - this.inset
        padding := Round(8 * A_ScreenDPI / 96)
        if top < this.offset
            this.ScrollTo(top - padding)
        else if bottom > this.offset + this.page
            this.ScrollTo(bottom - this.page + padding)
    }

    OnDestroy(wParam, lParam, message, hwnd) {
        if hwnd = this.contentHwnd || hwnd = this.viewportHwnd
            this.Stop()
    }

    Stop() {
        if this.stopped
            return
        this.stopped := true
        for message, callback in this.callbacks
            OnMessage(message, callback, 0)
        for control in this.focusControls
            try control.OnEvent("Focus", this.focusCallback, 0)
        this.focusControls := []
    }
}
