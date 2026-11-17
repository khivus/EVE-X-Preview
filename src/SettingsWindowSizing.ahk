class SettingsWindowSizing {
    static WorkArea(owner) {
        monitor := DllCall("MonitorFromWindow", "Ptr", owner, "UInt", 2, "Ptr")
        info := Buffer(40, 0)
        NumPut("UInt", 40, info)
        DllCall("GetMonitorInfoW", "Ptr", monitor, "Ptr", info)
        return {left: NumGet(info, 20, "Int"), top: NumGet(info, 24, "Int"), right: NumGet(info, 28, "Int"), bottom: NumGet(info, 32, "Int")}
    }
    static CenterPosition(ownerRect, width, height, work) {
        return {x: Max(work.left, Min(work.right - width, Round((ownerRect.left + ownerRect.right - width) / 2))),
            y: Max(work.top, Min(work.bottom - height, Round((ownerRect.top + ownerRect.bottom - height) / 2)))}
    }
    static CenterWindow(window, owner) {
        rect := Buffer(16, 0), parent := Buffer(16, 0)
        DllCall("GetWindowRect", "Ptr", window, "Ptr", rect)
        DllCall("GetWindowRect", "Ptr", owner, "Ptr", parent)
        bounds := {left: NumGet(parent, 0, "Int"), top: NumGet(parent, 4, "Int"), right: NumGet(parent, 8, "Int"), bottom: NumGet(parent, 12, "Int")}
        point := this.CenterPosition(bounds, NumGet(rect, 8, "Int") - NumGet(rect, 0, "Int"), NumGet(rect, 12, "Int") - NumGet(rect, 4, "Int"), this.WorkArea(owner))
        DllCall("SetWindowPos", "Ptr", window, "Ptr", 0, "Int", point.x, "Int", point.y, "Int", 0, "Int", 0, "UInt", 0x15)
    }
    static ShowCentered(gui, owner, width, height) {
        gui.Show("Hide w" width " h" height)
        this.CenterWindow(gui.Hwnd, owner)
        gui.Show()
    }
    static Fit(left, top, right, bottom, dpi := 96) {
        scale := dpi / 96
        availableW := Max(320, Floor((right - left - 24) / scale))
        availableH := Max(260, Floor((bottom - top - 40) / scale))
        width := Min(1000, availableW), height := Min(740, availableH)
        return {width: width, height: height, minWidth: Min(780, availableW), minHeight: Min(440, availableH),
            x: left + Max(0, Floor((right - left - width * scale) / 2)),
            y: top + Max(0, Floor((bottom - top - height * scale) / 2))}
    }
}
