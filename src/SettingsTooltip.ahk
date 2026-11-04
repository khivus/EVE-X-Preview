; A native tracked tooltip keeps standard accessibility and never takes focus.
class SettingsTooltip {
    __New(theme) {
        this.theme := theme
        this.Hwnd := DllCall("CreateWindowExW", "UInt", 0x08000088, "WStr", "tooltips_class32", "Ptr", 0,
            "UInt", 0x80000003, "Int", 0, "Int", 0, "Int", 0, "Int", 0, "Ptr", theme.app.S_Gui.Hwnd,
            "Ptr", 0, "Ptr", DllCall("GetModuleHandleW", "Ptr", 0, "Ptr"), "Ptr", 0, "Ptr")
        if !this.Hwnd
            throw OSError()
        this.text := Buffer(2, 0)
        this.info := Buffer(A_PtrSize = 8 ? 72 : 48, 0)
        NumPut("UInt", this.info.Size, "UInt", 0x1A0, this.info) ; TRACK | ABSOLUTE | TRANSPARENT
        NumPut("Ptr", theme.app.S_Gui.Hwnd, this.info, 8)
        NumPut("UPtr", 1, this.info, A_PtrSize = 8 ? 16 : 12)
        NumPut("Ptr", this.text.Ptr, this.info, A_PtrSize = 8 ? 48 : 36)
        this.Send(0x432, 0, this.info.Ptr)
        this.Send(0x418, 0, Round(420 * A_ScreenDPI / 96))
        margin := Buffer(16, 0)
        NumPut("Int", 8, "Int", 6, "Int", 8, "Int", 6, margin)
        this.Send(0x41A, 0, margin.Ptr)
        this.Apply()
    }
    Send(message, wParam, lParam) => DllCall("SendMessageW", "Ptr", this.Hwnd, "UInt", message, "Ptr", wParam, "Ptr", lParam, "Ptr")
    Apply() {
        if !DllCall("IsWindow", "Ptr", this.Hwnd)
            return
        DllCall("uxtheme\SetWindowTheme", "Ptr", this.Hwnd, "WStr", "", "WStr", "")
        this.Send(0x413, SettingsTheme.BGR(this.theme.dark ? this.theme.input : "FFFFFF"), 0)
        this.Send(0x414, SettingsTheme.BGR(this.theme.foreground), 0)
        this.theme.chrome.AttachWindow(this.Hwnd, "Tooltip")
        DllCall("RedrawWindow", "Ptr", this.Hwnd, "Ptr", 0, "Ptr", 0, "UInt", 0x585)
    }
    Show(text, x?, y?) {
        previous := this.text
        this.text := Buffer(StrPut(text, "UTF-16") * 2, 0)
        StrPut(text, this.text, "UTF-16")
        NumPut("Ptr", this.text.Ptr, this.info, A_PtrSize = 8 ? 48 : 36)
        this.Send(0x439, 0, this.info.Ptr)
        if !IsSet(x) || !IsSet(y) {
            point := Buffer(8, 0)
            DllCall("GetCursorPos", "Ptr", point)
            x := NumGet(point, 0, "Int") + 14, y := NumGet(point, 4, "Int") + 22
        }
        this.Send(0x412, 0, (x & 0xFFFF) | ((y & 0xFFFF) << 16))
        this.Send(0x411, 1, this.info.Ptr)
    }
    Hide() {
        if DllCall("IsWindow", "Ptr", this.Hwnd)
            this.Send(0x411, 0, this.info.Ptr)
    }
    Dispose() {
        if DllCall("IsWindow", "Ptr", this.Hwnd)
            DllCall("DestroyWindow", "Ptr", this.Hwnd)
    }
}
