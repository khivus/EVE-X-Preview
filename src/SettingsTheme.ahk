; Palette and native-control drawing belong to the settings window only.
; Configuration colors and thumbnail previews retain their actual RGB values.
class SettingsTheme {
    __New(app, mode := "system") {
        this.app := app, this.surfaces := [], this.controls := Map(), this.headers := Map(), this.statics := Map()
        this.updates := Map()
        this.chrome := SettingsNativeChrome(this)
        this.headerCallback := CallbackCreate(ObjBindMethod(this, "HeaderProc"),, 6)
        this.staticCallback := CallbackCreate(ObjBindMethod(this, "StaticProc"),, 6)
        this.systemCallback := ObjBindMethod(this, "SystemChanged")
        this.activationCallback := ObjBindMethod(this, "ActivationChanged")
        this.drawItemCallback := ObjBindMethod(this, "DrawItem")
        this.systemTimer := ObjBindMethod(this, "RefreshSystem")
        this.repaintTimer := ObjBindMethod(this, "Repaint")
        OnMessage(0x1A, this.systemCallback)
        OnMessage(0x6, this.activationCallback)
        OnMessage(0x2B, this.drawItemCallback)
        this.SetMode(mode)
    }
    SetMode(mode) {
        dark := mode = "dark"
        if mode = "system" {
            try dark := !RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "AppsUseLightTheme", 1)
        }
        changed := !this.HasOwnProp("dark") || dark != this.dark
        this.mode := mode, this.dark := dark
        if !dark && this.HasOwnProp("chrome")
            this.chrome.scrollbars.Finish(false)
        this.background := dark ? "202124" : "F7F7F7"
        this.sidebar := dark ? "191B1F" : "EFEFEF"
        this.input := dark ? "2B2E33" : "FFFFFF"
        this.foreground := dark ? "ECEEF2" : "202020"
        this.muted := dark ? "A6ADB8" : "707070"
        this.line := dark ? "424750" : "DDE1E7"
        if changed {
            this.ApplyAll()
            this.Repaint()
            SetTimer(this.repaintTimer, -1)
        }
        return changed
    }
    SystemChanged(*) {
        if this.mode = "system"
            SetTimer(this.systemTimer, -1)
    }
    ActivationChanged(wParam, lParam, message, hwnd) {
        if this.app.closed || wParam & 0xFFFF || !this.app.HasOwnProp("S_Gui")
            return
        if hwnd = this.app.S_Gui.Hwnd || (this.app.HasOwnProp("activeDialog") && hwnd = this.app.activeDialog.Hwnd) {
            this.chrome.scrollbars.Finish(false)
            this.app.ClearSettingsTooltip()
        }
    }
    RefreshSystem() {
        if !this.app.closed && this.mode = "system"
            this.SetMode("system")
    }
    Color(value) {
        if !this.dark
            return value
        return Map("202020", this.foreground, "000000", this.foreground,
            "444444", "D0D5DD", "555555", "BAC2CE", "606A76", "BAC2CE",
            "707070", this.muted, "526778", "9BC5E8", "765B00", "E9C579",
            "9A5700", "F0C674", "B42318", "FF9990").Get(StrUpper(value), value)
    }
    static BGR(value) {
        rgb := Integer("0x" value)
        return ((rgb & 255) << 16) | (rgb & 0xFF00) | ((rgb >> 16) & 255)
    }
    static Assign(control, property, value) {
        if control.%property% != value
            control.%property% := value
    }
    SetSwatch(control, color) {
        color := SettingsValue.Color(color)
        if control.ThemeBackground = color
            return
        control.ThemeBackground := color
        control.Opt("Background" color)
        DllCall("RedrawWindow", "Ptr", control.Hwnd, "Ptr", 0, "Ptr", 0, "UInt", 0x101)
    }
    Options(options) {
        if RegExMatch(options, "i)(?:^|\s)c([0-9a-f]{6})(?=\s|$)", &m)
            options := RegExReplace(options, "i)(?<=\b)c" m[1] "(?=\s|$)", "c" this.Color(m[1]))
        return options
    }
    Surface(gui, sidebar := false) {
        surface := SettingsThemeSurface(this, gui, sidebar)
        this.surfaces.Push(surface)
        this.Window(gui, sidebar)
        return surface
    }
    Window(gui, sidebar := false) {
        gui.BackColor := sidebar ? this.sidebar : this.background
        if DllCall("GetAncestor", "Ptr", gui.Hwnd, "UInt", 2, "Ptr") != gui.Hwnd
            return
        ; These controls mix native custom draw with local buffered painting.
        ; Whole-window compositing replays stale child/nonclient surfaces.
        gui.Opt("-E0x02000000")
        dark := Buffer(4, 0)
        NumPut("Int", this.dark, dark)
        try DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", gui.Hwnd, "UInt", 20, "Ptr", dark, "UInt", 4)
        DllCall("SetWindowPos", "Ptr", gui.Hwnd, "Ptr", 0, "Int", 0, "Int", 0, "Int", 0, "Int", 0, "UInt", 0x37)
    }
    BeginUpdate(gui) {
        handle := gui.Hwnd
        if this.updates.Has(handle) {
            this.updates[handle].depth++
            return
        }
        visible := !!DllCall("IsWindowVisible", "Ptr", handle)
        this.updates[handle] := {depth: 1, visible: visible}
        if visible
            DllCall("SendMessageW", "Ptr", handle, "UInt", 0xB, "Ptr", 0, "Ptr", 0)
    }
    EndUpdate(gui) {
        handle := gui.Hwnd
        if !this.updates.Has(handle) || --this.updates[handle].depth
            return
        update := this.updates.Delete(handle)
        if !DllCall("IsWindow", "Ptr", handle)
            return
        ; Never enable redraw on a window that was hidden: WM_SETREDRAW(TRUE)
        ; would make it visible, including a fixture or an unfinished editor.
        if update.visible
            DllCall("SendMessageW", "Ptr", handle, "UInt", 0xB, "Ptr", 1, "Ptr", 0)
        DllCall("RedrawWindow", "Ptr", handle, "Ptr", 0, "Ptr", 0, "UInt", 0x585)
    }
    Scroll(scroll) {
        scroll.theme := this
        for gui in [scroll.viewport, scroll.content, scroll.header]
            this.Window(gui)
        this.chrome.AttachWindow(scroll.viewportHwnd, "ScrollPane", scroll)
        this.chrome.AttachWindow(scroll.scrollbarHwnd, "PaneScrollbar", scroll)
        this.Control(scroll.heading)
        this.Control(scroll.rule)
    }
    Font(control, options, name := "") {
        previousColor := control.HasOwnProp("ThemeColor") ? control.ThemeColor : ""
        if RegExMatch(options, "i)(?:^|\s)c([0-9a-f]{6})(?=\s|$)", &m)
            control.ThemeColor := StrUpper(m[1])
        mapped := this.Options(options), color := this.Color(control.ThemeColor)
        painted := InStr("|Text|Button|CheckBox|Radio|", "|" control.Type "|")
        if painted
            mapped := Trim(RegExReplace(mapped, "i)(?:^|\s)c[0-9a-f]{6}(?=\s|$)"))
        key := mapped "|" name
        if !control.HasOwnProp("ThemeFontKey") || control.ThemeFontKey != key || (!painted && control.ThemeAppliedColor != color) {
            if mapped != "" || name != ""
                control.SetFont(mapped, name)
            control.ThemeFontKey := key, control.ThemeAppliedColor := color
        }
        if previousColor != control.ThemeColor
            DllCall("RedrawWindow", "Ptr", control.Hwnd, "Ptr", 0, "Ptr", 0, "UInt", 0x101)
    }
    Control(control, options := "", defaultColor := "202020", created := false) {
        handle := control.Hwnd
        if !control.HasOwnProp("ThemeColor") {
            control.ThemeColor := RegExMatch(options, "i)(?:^|\s)c([0-9a-f]{6})(?=\s|$)", &m) ? StrUpper(m[1]) : defaultColor
            control.ThemePreserve := InStr(options, "ThemePreserve") != 0
            control.ThemeBackground := RegExMatch(options, "i)\bBackground([0-9a-f]{6})(?=\s|$)", &m) ? StrUpper(m[1]) : ""
        }
        this.controls[handle] := control
        kind := control.Type
        color := this.Color(control.ThemeColor)
        if !created && (!control.HasOwnProp("ThemeAppliedColor") || control.ThemeAppliedColor != color)
            && InStr("|Edit|ComboBox|ListBox|DateTime|MonthCal|Link|", "|" kind "|")
            control.SetFont("c" this.Color(control.ThemeColor))
        control.ThemeAppliedColor := color
        if !control.ThemePreserve {
            if kind = "Edit" || kind = "ListBox" || kind = "ComboBox" || kind = "DDL"
                control.Opt("Background" (this.dark ? this.input : "FFFFFF"))
            else if kind = "Slider"
                control.Opt("Background" this.background)
            else if kind = "Text" || kind = "CheckBox" || kind = "Radio" {
                color := control.ThemeBackground
                if color = "DDE1E7" || color = "AAAAAA"
                    color := this.line
                else if color = "F7F7F7" || color = "E0E0E0" || color = ""
                    color := control.Gui.BackColor
                control.Opt("Background" color)
            }
        }
        if kind = "DDL" || kind = "ComboBox"
            control.Opt(this.dark ? "-Theme" : "+Theme")
        if kind = "Button"
            control.Opt("-Theme") ; Prevent native theme animations from replaying a cached light button.
        if (kind = "Text" || kind = "CheckBox" || kind = "Radio" || kind = "Button")
            && (!control.ThemePreserve || kind = "Text" && control.ThemeBackground != "") {
            this.statics[handle] := control
            DllCall("comctl32\SetWindowSubclass", "Ptr", handle, "Ptr", this.staticCallback, "UPtr", 1, "UPtr", 0)
        } else if kind = "ListView" {
            control.Opt("+0x02000000") ; WS_CLIPCHILDREN: the body must not paint over its header child.
            bg := SettingsTheme.BGR(this.dark ? this.input : "FFFFFF")
            SendMessage(0x1001, 0, bg, control)
            SendMessage(0x1026, 0, bg, control)
            SendMessage(0x1024, 0, SettingsTheme.BGR(this.foreground), control)
            header := SendMessage(0x101F, 0, 0, control)
            this.headers[header] := control
            DllCall("comctl32\SetWindowSubclass", "Ptr", header, "Ptr", this.headerCallback, "UPtr", 1, "UPtr", 0)
            DllCall("uxtheme\SetWindowTheme", "Ptr", header, "WStr", "", "WStr", "")
            SendMessage(0x1036, 1, this.dark ? 0 : 1, control)
            SendMessage(0x1036, 0x10000, 0x10000, control) ; LVS_EX_DOUBLEBUFFER, local to this table.
        } else if kind = "TreeView" {
            control.Opt(this.dark ? "-Theme" : "+Theme")
            SendMessage(0x112C, 4, this.dark ? 4 : 0, control) ; TVS_EX_DOUBLEBUFFER, local to dark navigation.
            SendMessage(0x111D, 0, SettingsTheme.BGR(this.dark ? this.sidebar : "FFFFFF"), control)
            SendMessage(0x111E, 0, SettingsTheme.BGR(this.foreground), control)
            SendMessage(0x1128, 0, SettingsTheme.BGR(this.line), control)
            if !control.HasOwnProp("ThemeTreeDraw") {
                control.ThemeTreeDraw := ObjBindMethod(this, "TreeDraw")
                control.OnNotify(-12, control.ThemeTreeDraw)
            }
        } else if kind = "Slider" {
            if !control.HasOwnProp("ThemeSliderDraw") {
                control.ThemeSliderDraw := ObjBindMethod(this, "SliderDraw")
                control.OnNotify(-12, control.ThemeSliderDraw)
            }
        } else if kind = "Button" || kind = "CheckBox" || kind = "Radio" {
            if !control.HasOwnProp("ThemeDraw") {
                control.ThemeDraw := ObjBindMethod(this, "ButtonDraw")
                control.OnNotify(-12, control.ThemeDraw)
            }
        }
        this.chrome.Attach(control)
        control.Redraw()
        return control
    }
    Prune() {
        index := this.surfaces.Length
        while index > 0 {
            surface := this.surfaces[index]
            if !DllCall("IsWindow", "Ptr", surface.Hwnd)
                this.surfaces.RemoveAt(index)
            index--
        }
        for handle, control in this.controls.Clone() {
            if !DllCall("IsWindow", "Ptr", handle)
                this.controls.Delete(handle)
        }
    }
    ApplyAll() {
        this.Prune()
        for surface in this.surfaces
            this.Window(surface.gui, surface.sidebar)
        for handle, control in this.controls
            this.Control(control)
        if this.app.HasOwnProp("settingsScroller")
            this.Scroll(this.app.settingsScroller)
        if this.app.HasOwnProp("activeDialog")
            this.Scroll(this.app.activeDialog.scroll)
        if this.HasOwnProp("tooltip")
            this.tooltip.Apply()
    }
    Repaint() {
        if this.app.closed || this.updates.Count
            return
        this.Prune()
        for surface in this.surfaces
            DllCall("RedrawWindow", "Ptr", surface.Hwnd, "Ptr", 0, "Ptr", 0, "UInt", 0x585)
    }
    static Fill(dc, rect, color, frame := false) {
        brush := DllCall("CreateSolidBrush", "UInt", SettingsTheme.BGR(color), "Ptr")
        DllCall(frame ? "FrameRect" : "FillRect", "Ptr", dc, "Ptr", rect, "Ptr", brush)
        DllCall("DeleteObject", "Ptr", brush)
    }
    ShowTooltip(text, x?, y?) {
        if !this.HasOwnProp("tooltip")
            this.tooltip := SettingsTooltip(this)
        this.tooltip.Show(text, x?, y?)
    }
    HideTooltip() {
        if this.HasOwnProp("tooltip")
            this.tooltip.Hide()
    }
    SliderDraw(control, data) {
        if !this.dark
            return 0
        stage := NumGet(data, A_PtrSize * 3, "UInt")
        if stage = 1
            return 0x20
        if stage != 0x10001
            return 0
        part := NumGet(data, A_PtrSize = 8 ? 56 : 36, "UPtr")
        if part = 1 {
            dc := NumGet(data, A_PtrSize * 4, "Ptr"), rect := Buffer(16, 0), channel := Buffer(16, 0)
            DllCall("GetClientRect", "Ptr", control.Hwnd, "Ptr", rect)
            SendMessage(0x41A, 0, channel.Ptr, control)
            left := NumGet(channel, 0, "Int"), right := NumGet(channel, 8, "Int")
            bottom := NumGet(rect, 12, "Int"), size := Max(2, Round(4 * A_ScreenDPI / 96))
            loop 11 {
                tick := Buffer(16, 0), x := Round(left + (right - left - 1) * (A_Index - 1) / 10)
                NumPut("Int", x, "Int", bottom - size - 1, "Int", x + 1, "Int", bottom - 1, tick)
                SettingsTheme.Fill(dc, tick, "7D8793")
            }
            return 4
        }
        rect := Buffer(16, 0), dc := NumGet(data, A_PtrSize * 4, "Ptr")
        SendMessage(part = 2 ? 0x419 : 0x41A, 0, rect.Ptr, control)
        SettingsTheme.Fill(dc, rect, part = 2 ? control.Enabled ? "76B7FF" : this.muted : "424750")
        if part = 2 && DllCall("GetFocus", "Ptr") = control.Hwnd
            SettingsTheme.Fill(dc, rect, this.foreground, true)
        return 4
    }
    TreeDraw(control, data) {
        if !this.dark
            return 0
        stage := NumGet(data, A_PtrSize * 3, "UInt")
        if stage = 1
            return 0x20
        if stage = 0x10001 {
            stateOffset := A_PtrSize = 8 ? 64 : 40
            state := NumGet(data, stateOffset, "UInt")
            NumPut("UInt", SettingsTheme.BGR(this.foreground), data, A_PtrSize = 8 ? 80 : 48)
            NumPut("UInt", SettingsTheme.BGR(state & 1 ? "304C70" : this.sidebar), data, A_PtrSize = 8 ? 84 : 52)
            NumPut("UInt", state & ~0x11, data, stateOffset)
            return 0x12
        }
        if stage != 0x10002
            return 0
        item := NumGet(data, A_PtrSize = 8 ? 56 : 36, "UPtr")
        info := Buffer(A_PtrSize = 8 ? 56 : 40, 0)
        NumPut("UInt", 0x40, info)
        NumPut("Ptr", item, info, A_PtrSize = 8 ? 8 : 4)
        SendMessage(0x113E, 0, info.Ptr, control)
        if !NumGet(info, A_PtrSize = 8 ? 44 : 32, "Int")
            return 0
        rect := Buffer(16, 0), part := Buffer(A_PtrSize * 3, 0)
        NumPut("Ptr", item, "Ptr", rect.Ptr, part)
        NumPut("Int", 1, part, A_PtrSize * 2)
        if !SendMessage(0x1148, 0, part.Ptr, control)
            return 0
        dc := NumGet(data, A_PtrSize * 4, "Ptr")
        SettingsTheme.Fill(dc, rect, this.sidebar)
        expanded := SendMessage(0x1127, item, 0x20, control)
        this.chrome.Arrow(dc, rect, expanded ? "down" : "right", this.muted)
        return 0
    }
    Query(control) {
        if !this.app.HasOwnProp("search") || !this.app.HasOwnProp("settingsScroller")
            return ""
        parent := control.Gui.Hwnd
        scroll := this.app.settingsScroller
        if parent = scroll.contentHwnd || parent = scroll.headerHwnd
            return Trim(this.app.search.Value)
        if this.app.HasOwnProp("activeDialog") && parent = this.app.activeDialog.content.Hwnd
            return Trim(this.app.search.Value)
        return ""
    }
    Text(dc, text, rect, flags, control) {
        query := this.Query(control)
        return SettingsSearchText.Draw(dc, text, rect, flags, query, this.dark)
    }
    DrawItem(wParam, lParam, message, hwnd) {
        if !lParam || NumGet(lParam, 0, "UInt") != 3 ; ODT_COMBOBOX
            return
        handle := NumGet(lParam, A_PtrSize = 8 ? 24 : 20, "Ptr")
        if !this.controls.Has(handle)
            return
        control := this.controls[handle]
        dc := NumGet(lParam, A_PtrSize = 8 ? 32 : 24, "Ptr")
        rect := Buffer(16, 0)
        DllCall("RtlMoveMemory", "Ptr", rect, "Ptr", lParam + (A_PtrSize = 8 ? 40 : 28), "UPtr", 16)
        state := NumGet(lParam, 16, "UInt"), index := NumGet(lParam, 8, "Int")
        selected := state & 1, disabled := state & 4 || !control.Enabled
        SettingsTheme.Fill(dc, rect, this.dark ? selected ? "304C70" : this.input : selected ? "0078D4" : "FFFFFF")
        if index >= 0 {
            length := SendMessage(0x149, index, 0, control) ; CB_GETLBTEXTLEN
            text := Buffer(Max(2, (length + 1) * 2), 0)
            SendMessage(0x148, index, text.Ptr, control)
            saved := DllCall("SaveDC", "Ptr", dc, "Int")
            DllCall("SetBkMode", "Ptr", dc, "Int", 1)
            color := disabled ? this.muted : selected && !this.dark ? "FFFFFF" : this.foreground
            DllCall("SetTextColor", "Ptr", dc, "UInt", SettingsTheme.BGR(color))
            DllCall("SelectObject", "Ptr", dc, "Ptr", SendMessage(0x31, 0, 0, control), "Ptr")
            NumPut("Int", NumGet(rect, 0, "Int") + 4, rect)
            this.Text(dc, StrGet(text), rect, 0x8824, control)
            if state & 0x10 {
                if this.dark
                    SettingsTheme.Fill(dc, rect, "76B7FF", true)
                else
                    DllCall("DrawFocusRect", "Ptr", dc, "Ptr", rect)
            }
            DllCall("RestoreDC", "Ptr", dc, "Int", saved)
        }
        return 1
    }
    ButtonDraw(control, data) {
        if NumGet(data, A_PtrSize * 3, "UInt") != 1
            return 0
        dc := NumGet(data, A_PtrSize * 4, "Ptr")
        rect := Buffer(16, 0)
        DllCall("GetClientRect", "Ptr", control.Hwnd, "Ptr", rect)
        saved := DllCall("SaveDC", "Ptr", dc, "Int")
        state := NumGet(data, A_PtrSize = 8 ? 64 : 40, "UInt")
        background := control.Type = "Button" ? this.dark ? state & 1 ? "3D4859" : state & 0x40 ? "39404B" : "30343B"
            : state & 1 ? "CEE2F6" : state & 0x40 ? "E2EEF9" : "EFEFEF" : control.Gui.BackColor
        SettingsTheme.Fill(dc, rect, background)
        if control.Type = "Button"
            SettingsTheme.Fill(dc, rect, control.Enabled && (WinGetStyle(control.Hwnd) & 0xF) = 1 ? this.dark ? "76B7FF" : "0078D4" : this.line, true)
        else {
            box := Buffer(16, 0), height := NumGet(rect, 12, "Int")
            scale := A_ScreenDPI / 96, size := Round(14 * scale), left := Round(scale)
            top := Max(2, (height - size) // 2)
            NumPut("Int", left, "Int", top, "Int", left + size, "Int", top + size, box)
            SettingsTheme.Fill(dc, box, this.input)
            SettingsTheme.Fill(dc, box, control.Enabled ? this.muted : this.line, true)
            if SendMessage(0xF0, 0, 0, control) {
                pen := DllCall("CreatePen", "Int", 0, "Int", Max(1, Round(2 * scale)), "UInt", SettingsTheme.BGR(control.Enabled ? this.dark ? "76B7FF" : "0078D4" : this.muted), "Ptr")
                previous := DllCall("SelectObject", "Ptr", dc, "Ptr", pen, "Ptr")
                DllCall("MoveToEx", "Ptr", dc, "Int", left + Round(3 * scale), "Int", top + Round(7 * scale), "Ptr", 0)
                DllCall("LineTo", "Ptr", dc, "Int", left + Round(6 * scale), "Int", top + Round(10 * scale))
                DllCall("LineTo", "Ptr", dc, "Int", left + Round(11 * scale), "Int", top + Round(3 * scale))
                DllCall("SelectObject", "Ptr", dc, "Ptr", previous)
                DllCall("DeleteObject", "Ptr", pen)
            }
            NumPut("Int", Round(22 * scale), rect)
        }
        DllCall("SetBkMode", "Ptr", dc, "Int", 1)
        DllCall("SetTextColor", "Ptr", dc, "UInt", SettingsTheme.BGR(control.Enabled ? this.Color(control.ThemeColor) : this.muted))
        DllCall("SelectObject", "Ptr", dc, "Ptr", SendMessage(0x31, 0, 0, control), "Ptr")
        if control.Type != "Button" {
            measure := Buffer(16, 0)
            DllCall("RtlMoveMemory", "Ptr", measure, "Ptr", rect, "UPtr", 16)
            DllCall("DrawTextW", "Ptr", dc, "WStr", control.Text, "Int", -1, "Ptr", measure, "UInt", 0x410)
            textHeight := NumGet(measure, 12, "Int") - NumGet(measure, 4, "Int")
            NumPut("Int", Max(0, (NumGet(rect, 12, "Int") - textHeight) // 2), rect, 4)
        }
        this.Text(dc, control.Text, rect, control.Type = "Button" ? 0x25 : 0x10, control)
        if DllCall("GetFocus", "Ptr") = control.Hwnd {
            DllCall("InflateRect", "Ptr", rect, "Int", -3, "Int", -3)
            if this.dark
                SettingsTheme.Fill(dc, rect, "76B7FF", true)
            else
                DllCall("DrawFocusRect", "Ptr", dc, "Ptr", rect)
        }
        DllCall("RestoreDC", "Ptr", dc, "Int", saved)
        return 4
    }
    HeaderProc(header, message, wParam, lParam, subclass, data) {
        if message = 0x82 {
            if this.headers.Has(header)
                this.headers.Delete(header)
            DllCall("comctl32\RemoveWindowSubclass", "Ptr", header, "Ptr", this.headerCallback, "UPtr", 1)
        } else if this.headers.Has(header) && message = 0x14 {
            return 1 ; The buffered header paints every pixel; never erase it first.
        } else if this.headers.Has(header) && (message = 0xF || message = 0x317 || message = 0x318) {
            if message = 0xF {
                paint := Buffer(A_PtrSize = 8 ? 72 : 64, 0)
                dc := DllCall("BeginPaint", "Ptr", header, "Ptr", paint, "Ptr")
            } else
                dc := wParam
            rect := Buffer(16, 0)
            DllCall("GetClientRect", "Ptr", header, "Ptr", rect)
            memory := DllCall("CreateCompatibleDC", "Ptr", dc, "Ptr")
            bitmap := DllCall("CreateCompatibleBitmap", "Ptr", dc, "Int", Max(1, NumGet(rect, 8, "Int")), "Int", Max(1, NumGet(rect, 12, "Int")), "Ptr")
            previous := DllCall("SelectObject", "Ptr", memory, "Ptr", bitmap, "Ptr")
            try {
                this.PaintHeader(header, memory)
                DllCall("BitBlt", "Ptr", dc, "Int", 0, "Int", 0, "Int", NumGet(rect, 8, "Int"), "Int", NumGet(rect, 12, "Int"),
                    "Ptr", memory, "Int", 0, "Int", 0, "UInt", 0xCC0020)
            } finally {
                DllCall("SelectObject", "Ptr", memory, "Ptr", previous)
                DllCall("DeleteObject", "Ptr", bitmap)
                DllCall("DeleteDC", "Ptr", memory)
            }
            if message = 0xF
                DllCall("EndPaint", "Ptr", header, "Ptr", paint)
            return 0
        }
        return DllCall("comctl32\DefSubclassProc", "Ptr", header, "UInt", message, "Ptr", wParam, "Ptr", lParam, "Ptr")
    }
    StaticProc(handle, message, wParam, lParam, subclass, data) {
        if message = 0x82 {
            if this.statics.Has(handle)
                this.statics.Delete(handle)
            DllCall("comctl32\RemoveWindowSubclass", "Ptr", handle, "Ptr", this.staticCallback, "UPtr", 1)
        } else if this.statics.Has(handle) && message = 0x14 {
            return 1 ; Every custom paint fills its complete client area.
        } else if this.statics.Has(handle) && (message = 0xA || message = 0x30 || message = 0xC || message = 0x31A
            || message = 0x7 || message = 0x8 || message = 0xF4 || message = 0xF3 || message = 0xF1
            || message = 0x127 || message = 0x128
            || ((message = 0x200 || message = 0x2A3) && this.statics[handle].Type = "Button")
            || message = 0x201 || message = 0x202 || message = 0x100 || message = 0x101) {
            control := this.statics[handle]
            if control.Type = "Button" && message = 0x200 && (!control.HasOwnProp("ThemeHot") || !control.ThemeHot) {
                control.ThemeHot := true
                track := Buffer(A_PtrSize = 8 ? 24 : 16, 0)
                NumPut("UInt", track.Size, "UInt", 2, "Ptr", handle, track)
                DllCall("TrackMouseEvent", "Ptr", track)
            } else if message = 0x2A3
                control.ThemeHot := false
            ; Dialog activation changes BS_DEFPUSHBUTTON using BM_SETSTYLE.
            ; Keep its native semantics, but suppress the native immediate draw.
            result := DllCall("comctl32\DefSubclassProc", "Ptr", handle, "UInt", message, "Ptr", wParam, "Ptr", message = 0xF4 ? 0 : lParam, "Ptr")
            if !this.updates.Count
                DllCall("RedrawWindow", "Ptr", handle, "Ptr", 0, "Ptr", 0, "UInt", 0x101)
            return result
        } else if this.statics.Has(handle) && (message = 0xF || message = 0x317 || message = 0x318) {
            if message = 0xF {
                paint := Buffer(A_PtrSize = 8 ? 72 : 64, 0)
                dc := DllCall("BeginPaint", "Ptr", handle, "Ptr", paint, "Ptr")
            } else
                dc := wParam
            control := this.statics[handle]
            if control.Type = "Text"
                this.PaintStatic(control, dc)
            else {
                draw := Buffer(80, 0)
                NumPut("UInt", 1, draw, A_PtrSize * 3)
                NumPut("Ptr", dc, draw, A_PtrSize * 4)
                nativeState := SendMessage(0xF2, 0, 0, control)
                state := (nativeState & 4 ? 1 : 0) | (control.HasOwnProp("ThemeHot") && control.ThemeHot ? 0x40 : 0)
                NumPut("UInt", state, draw, A_PtrSize = 8 ? 64 : 40)
                this.ButtonDraw(control, draw)
            }
            if message = 0xF
                DllCall("EndPaint", "Ptr", handle, "Ptr", paint)
            return 0
        }
        return DllCall("comctl32\DefSubclassProc", "Ptr", handle, "UInt", message, "Ptr", wParam, "Ptr", lParam, "Ptr")
    }
    PaintStatic(control, dc) {
        rect := Buffer(16, 0)
        DllCall("GetClientRect", "Ptr", control.Hwnd, "Ptr", rect)
        if control.ThemePreserve {
            SettingsTheme.Fill(dc, rect, control.ThemeBackground)
            return
        }
        style := WinGetStyle(control.Hwnd), kind := style & 0x1F
        color := control.ThemeBackground
        if color = "DDE1E7" || color = "AAAAAA" || kind = 0x10
            color := this.line
        else if color = "F7F7F7" || color = "E0E0E0" || color = ""
            color := control.Gui.BackColor
        SettingsTheme.Fill(dc, rect, color)
        if kind = 0x10
            return
        saved := DllCall("SaveDC", "Ptr", dc, "Int")
        DllCall("SetBkMode", "Ptr", dc, "Int", 1)
        DllCall("SetTextColor", "Ptr", dc, "UInt", SettingsTheme.BGR(control.Enabled ? this.Color(control.ThemeColor) : this.muted))
        DllCall("SelectObject", "Ptr", dc, "Ptr", SendMessage(0x31, 0, 0, control), "Ptr")
        flags := (kind = 1 ? 1 : kind = 2 ? 2 : 0) | (style & 0x80 ? 0x800 : 0) | (style & 0x200 ? 0x24 : 0x10)
        this.Text(dc, control.Text, rect, flags, control)
        DllCall("RestoreDC", "Ptr", dc, "Int", saved)
    }
    PaintHeader(header, dc) {
        list := this.headers[header]
        rect := Buffer(16, 0)
        DllCall("GetClientRect", "Ptr", header, "Ptr", rect)
        SettingsTheme.Fill(dc, rect, this.dark ? "343941" : "EFEFEF")
        saved := DllCall("SaveDC", "Ptr", dc, "Int")
        DllCall("SetBkMode", "Ptr", dc, "Int", 1)
        DllCall("SetTextColor", "Ptr", dc, "UInt", SettingsTheme.BGR(this.foreground))
        DllCall("SelectObject", "Ptr", dc, "Ptr", SendMessage(0x31, 0, 0, list), "Ptr")
        loop SendMessage(0x1200, 0, 0,, "ahk_id " header) {
            index := A_Index - 1
            text := Buffer(1024, 0), item := Buffer(A_PtrSize = 8 ? 72 : 48, 0)
            NumPut("UInt", 2, item)
            NumPut("Ptr", text.Ptr, item, 8)
            NumPut("Int", 512, item, A_PtrSize = 8 ? 24 : 16)
            SendMessage(0x120B, index, item.Ptr,, "ahk_id " header)
            SendMessage(0x1207, index, rect.Ptr,, "ahk_id " header)
            SettingsTheme.Fill(dc, rect, this.line, true)
            NumPut("Int", NumGet(rect, 0, "Int") + 7, rect)
            this.Text(dc, StrGet(text), rect, 0x8824, list)
        }
        DllCall("RestoreDC", "Ptr", dc, "Int", saved)
    }
    Dispose() {
        if this.HasOwnProp("tooltip")
            this.tooltip.Dispose()
        this.chrome.Dispose()
        for header in this.headers
            if DllCall("IsWindow", "Ptr", header)
                DllCall("comctl32\RemoveWindowSubclass", "Ptr", header, "Ptr", this.headerCallback, "UPtr", 1)
        CallbackFree(this.headerCallback)
        for handle in this.statics
            if DllCall("IsWindow", "Ptr", handle)
                DllCall("comctl32\RemoveWindowSubclass", "Ptr", handle, "Ptr", this.staticCallback, "UPtr", 1)
        CallbackFree(this.staticCallback)
        OnMessage(0x1A, this.systemCallback, 0)
        OnMessage(0x6, this.activationCallback, 0)
        OnMessage(0x2B, this.drawItemCallback, 0)
        SetTimer(this.systemTimer, 0)
        SetTimer(this.repaintTimer, 0)
        this.controls.Clear(), this.headers.Clear(), this.statics.Clear(), this.updates.Clear(), this.surfaces := []
    }
}

class SettingsThemeSurface {
    __New(theme, gui, sidebar := false) {
        for name, value in Map("theme", theme, "gui", gui, "handle", gui.Hwnd, "sidebar", sidebar, "fontColor", "202020")
            this.DefineProp(name, {Value: value})
        this.SetFont("c202020")
    }
    Hwnd => this.handle
    __Enum(count) => this.gui.__Enum(count)
    __Get(name, params) => this.gui.%name%[params*]
    __Set(name, params, value) => this.gui.%name%[params*] := value
    SetFont(options := "", name := "") {
        if RegExMatch(options, "i)(?:^|\s)c([0-9a-f]{6})(?=\s|$)", &m)
            this.fontColor := StrUpper(m[1])
        this.gui.SetFont(this.theme.Options(options) " c" this.theme.Color(this.fontColor), name)
    }
    __Call(name, args) {
        if SubStr(name, 1, 3) = "Add" {
            original := args.Length ? args[1] : ""
            if args.Length
                args[1] := StrReplace(this.theme.Options(original), "ThemePreserve")
            if name = "AddDropDownList"
                args[1] .= " +0x210" ; Owner-draw with strings keeps native selection/keyboard behavior.
            if name = "AddUpDown"
                args[1] .= " +0x80" ; UDS_NOTHOUSANDS: buddy edits contain plain integers.
            control := this.gui.%name%(args*)
            return this.theme.Control(control, original, this.fontColor, true)
        }
        return this.gui.%name%(args*)
    }
}
