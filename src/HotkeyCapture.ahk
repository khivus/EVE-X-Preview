class HotkeyCapture {
    static Active := false
    static Current := 0
    static MouseKeys := ["LButton", "RButton", "MButton", "XButton1", "XButton2", "WheelUp", "WheelDown", "WheelLeft", "WheelRight"]
    held := Map()
    result := ""
    firstKey := ""
    cancelled := false
    finished := false

    Start(owner, helper, onComplete, row := 1) {
        if HotkeyCapture.Active
            return false
        wasCritical := A_IsCritical
        Critical(true)
        ; Main settings forms live in child GUIs. Focus belongs to their
        ; top-level window; owned row dialogs keep their own top-level root.
        this.ownerHwnd := DllCall("GetAncestor", "Ptr", owner.Hwnd, "UInt", 2, "Ptr") ; GA_ROOT
        this.helper := helper
        this.area := helper.waiting
        this.onComplete := onComplete
        this.wasSuspended := A_IsSuspended
        this.fieldWasVisible := helper.field.Visible
        this.buttonText := helper.button.Text
        this.buttonTooltip := helper.button.ToolTip
        this.hook := InputHook("L0 T20")
        this.hook.KeyOpt("{All}", "NS")
        this.hook.OnKeyDown := ObjBindMethod(this, "KeyDown")
        this.hook.OnKeyUp := ObjBindMethod(this, "KeyUp")
        this.hook.OnEnd := (*) => this.Cancel()
        this.focusTimer := ObjBindMethod(this, "CheckFocus")
        this.clickCallback := ObjBindMethod(this, "ClickOutside")
        HotkeyCapture.Current := this
        HotkeyCapture.Active := true
        try {
            Suspend(true)
            helper.field.Visible := false
            helper.field.GetPos(,, &width)
            this.area.Text := helper.multiline ? "Line " row ": press shortcut..."
                : width >= 100 ? "Press shortcut..." : "Press keys..."
            this.area.Visible := true
            helper.button.Text := Chr(0xD7)
            helper.button.ToolTip := "Cancel recording (or press Esc). Mouse buttons and the wheel can be recorded over the gray field."
            helper.button.Focus()
            for key in ["LControl", "RControl", "LAlt", "RAlt", "LShift", "RShift", "LWin", "RWin"]
                if GetKeyState(key, "P")
                    this.held[key] := true
            this.hook.Start()
            HotkeyCapture.SetMouseHotkeys(true)
            OnMessage(0x201, this.clickCallback) ; Cancel before another setting handles a click.
            SetTimer(this.focusTimer, 100)
            return true
        } catch as err {
            this.Cancel()
            throw err
        } finally {
            Critical(wasCritical)
        }
    }

    static SetMouseHotkeys(enabled) {
        if !this.HasOwnProp("mouseCondition")
            this.mouseCondition := ObjBindMethod(this, "MouseAllowed")
        HotIf(this.mouseCondition)
        try {
            for key in this.MouseKeys {
                if enabled {
                    Hotkey("*" key, ObjBindMethod(this, "MouseEvent", key, false), "S1 On")
                    if !InStr(key, "Wheel")
                        Hotkey("*" key " Up", ObjBindMethod(this, "MouseEvent", key, true), "S1 On")
                } else {
                    try Hotkey("*" key, "Off")
                    if !InStr(key, "Wheel")
                        try Hotkey("*" key " Up", "Off")
                }
            }
        } finally {
            HotIf()
        }
    }

    static MouseAllowed(hotkeyName := "") {
        if !this.Active || !IsObject(this.Current)
            return false
        capture := this.Current
        if capture.finished
            return false
        if !capture.HasFocus()
            return false
        ; Keep receiving releases even when the pointer leaves the field.
        if InStr(hotkeyName, " Up") {
            for key in this.MouseKeys
                if capture.held.Has(key)
                    return true
        }
        MouseGetPos(,,, &control, 2)
        return control = capture.area.Hwnd
    }

    static MouseEvent(key, up, *) {
        if !IsObject(this.Current)
            return
        if up
            this.Current.Released(key)
        else {
            this.Current.Pressed(key)
            if InStr(key, "Wheel")
                this.Current.Released(key)
        }
    }

    KeyDown(ih, vk, sc) {
        if this.HasFocus()
            this.Pressed(this.KeyName(vk, sc))
        else
            this.Cancel()
    }
    KeyUp(ih, vk, sc) => this.Released(this.KeyName(vk, sc))

    KeyName(vk, sc) {
        code := Format("vk{:02X}sc{:03X}", vk, sc)
        name := GetKeyName(code)
        return name = "" || RegExMatch(name, "[^\x20-\x7E]|^[!#+^&<>~*$]$") ? code : name
    }

    Modifiers() {
        mods := ""
        for pair in [["^", ["LControl", "RControl", "Control"]], ["!", ["LAlt", "RAlt", "Alt"]],
            ["+", ["LShift", "RShift", "Shift"]], ["#", ["LWin", "RWin"]]] {
            for name in pair[2] {
                if this.held.Has(name) {
                    mods .= pair[1]
                    break
                }
            }
        }
        return mods
    }

    IsModifier(key) => RegExMatch(key, "i)^(?:[LR]?(?:Control|Ctrl|Shift|Alt)|[LR]Win)$")

    Pressed(key) {
        if this.finished || this.held.Has(key)
            return
        mods := this.Modifiers()
        if key = "Escape" && mods = "" {
            this.Cancel()
            return
        }
        if this.IsModifier(key) {
            if this.firstKey = ""
                this.result := mods key
        } else {
            if this.firstKey != "" && !this.held.Has(this.firstKey)
                this.firstKey := ""
            this.result := mods (this.firstKey != "" ? this.firstKey " & " : "") key
            if this.firstKey = ""
                this.firstKey := key
        }
        this.held[key] := true
        if this.HasOwnProp("area")
            this.area.Text := this.result
    }

    Released(key) {
        if this.finished || !this.held.Has(key)
            return
        this.held.Delete(key)
        if !this.held.Count && this.result != ""
            this.Finish()
    }

    CheckFocus(*) {
        if !this.finished && (!this.HasFocus() || !this.helper.button.Visible || !this.helper.field.Enabled)
            this.Cancel()
    }

    HasFocus() => !!WinActive("ahk_id " this.ownerHwnd)

    ClickOutside(wParam, lParam, msg, hwnd) {
        if !this.finished && hwnd != this.area.Hwnd && hwnd != this.helper.button.Hwnd
            this.Cancel()
    }

    Cancel(*) {
        if this.finished
            return
        this.cancelled := true
        this.Finish()
    }

    Finish() {
        if this.finished
            return
        this.finished := true
        if !this.HasOwnProp("hook")
            return
        try {
            SetTimer(this.focusTimer, 0)
            this.focusTimer := 0
            OnMessage(0x201, this.clickCallback, 0)
            this.clickCallback := 0
            this.hook.OnEnd := ""
            this.hook.OnKeyDown := ""
            this.hook.OnKeyUp := ""
            this.hook.Stop()
            HotkeyCapture.SetMouseHotkeys(false)
        } finally {
            HotkeyCapture.Active := false
            HotkeyCapture.Current := 0
            Suspend(this.wasSuspended)
            try {
                this.area.Visible := false
                this.helper.field.Visible := this.fieldWasVisible
                this.helper.button.Text := this.buttonText
                this.helper.button.ToolTip := this.buttonTooltip
            }
        }
        callback := this.onComplete
        this.onComplete := 0
        callback.Call(this.cancelled ? "" : this.result)
    }
}
