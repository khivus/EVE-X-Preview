class ColorPicker {
    static CustomColors := Buffer(16 * 4, 0)

    static Normalize(value, fallback := "FFFFFF") {
        value := convertToHex(value)
        if RegExMatch(value, "i)^(?:#|0x)?([0-9a-f]{1,6})$", &m)
            return Format("{:06X}", Integer("0x" m[1]))
        return fallback
    }

    static SwapChannels(value) => ((value & 0xFF) << 16) | (value & 0xFF00) | ((value >> 16) & 0xFF)

    static BuildDialog(ownerHwnd, initial) {
        ; CHOOSECOLORW: pointer alignment differs between 32-bit and 64-bit AHK.
        dialog := Buffer(A_PtrSize = 8 ? 72 : 36, 0)
        NumPut("UInt", dialog.Size, dialog, 0)
        NumPut("Ptr", ownerHwnd, dialog, A_PtrSize)
        NumPut("UInt", this.SwapChannels(Integer("0x" this.Normalize(initial))), dialog, A_PtrSize * 3)
        NumPut("Ptr", this.CustomColors.Ptr, dialog, A_PtrSize * 4)
        NumPut("UInt", 0x103, dialog, A_PtrSize * 5) ; RGBINIT | FULLOPEN | ANYCOLOR
        return dialog
    }

    static Show(ownerHwnd, initial) {
        dialog := this.BuildDialog(ownerHwnd, initial)
        hook := CallbackCreate(ObjBindMethod(this, "CenterDialog"),, 4)
        NumPut("UInt", 0x113, dialog, A_PtrSize * 5) ; Also enable the initialization hook.
        NumPut("Ptr", hook, dialog, A_PtrSize = 8 ? 56 : 28)
        try {
            if DllCall("comdlg32\ChooseColorW", "Ptr", dialog, "Int")
                return Format("{:06X}", this.SwapChannels(NumGet(dialog, A_PtrSize * 3, "UInt")))
            code := DllCall("comdlg32\CommDlgExtendedError", "UInt")
            if code
                throw Error("Cannot open color picker",, Format("Common dialog error 0x{:X}", code))
            return "" ; Cancel preserves the current field and saved settings.
        } finally CallbackFree(hook)
    }
    static CenterDialog(hwnd, message, wParam, lParam) {
        if message = 0x110 {
            owner := NumGet(lParam, A_PtrSize, "Ptr")
            if owner && DllCall("IsWindow", "Ptr", owner)
                SettingsWindowSizing.CenterWindow(hwnd, owner)
        }
        return 0
    }
}
