; Native lists keep their ranges and keyboard behavior. Their visible bars use
; the same separate, app-painted SCROLLBAR controls as SettingsScrollPane.
class SettingsControlScrollbars {
    __New(chrome) => this.chrome := chrome

    SyncAll() {
        items := []
        for handle, item in this.chrome.windows
            if item.kind != "PaneScrollbar"
                items.Push(item)
        for item in items
            this.Sync(item)
    }

    Sync(item) {
        if !InStr("|TreeView|ListView|ListBox|ComboList|Edit|", "|" item.kind "|")
            || (item.HasOwnProp("syncingBars") && item.syncingBars)
            || (item.HasOwnProp("scrollUpdating") && item.scrollUpdating)
            || (item.HasOwnProp("redrawDisabled") && item.redrawDisabled)
            || this.chrome.theme.updates.Count
            return
        item.syncingBars := true
        try {
            if !item.HasOwnProp("bars")
                item.bars := Map()
            visible := this.chrome.theme.dark && DllCall("IsWindowVisible", "Ptr", item.handle)
            popup := item.kind = "ComboList"
            parent := popup ? item.handle : DllCall("GetParent", "Ptr", item.handle, "Ptr")
            for axis in [1, 0] {
                info := visible ? this.chrome.NativeScrollInfo(item.handle, axis = 1 ? -5 : -6) : 0
                if !IsObject(info) {
                    if item.bars.Has(axis) {
                        item.bars[axis].signature := ""
                        DllCall("ShowWindow", "Ptr", item.bars[axis].handle, "Int", 0)
                    }
                    continue
                }
                if !item.bars.Has(axis) {
                    bar := SettingsControlScrollbar(this.chrome, item, axis, parent, popup)
                    item.bars[axis] := bar
                    this.chrome.AttachWindow(bar.handle, "PaneScrollbar", bar)
                    this.chrome.windows[bar.handle].axis := axis
                }
                bar := item.bars[axis]
                range := Buffer(28, 0)
                NumPut("UInt", 28, "UInt", 7, range)
                DllCall("GetScrollInfo", "Ptr", item.handle, "Int", axis, "Ptr", range)
                rect := Buffer(16, 0)
                DllCall("RtlMoveMemory", "Ptr", rect, "Ptr", info.Ptr + 4, "UPtr", 16)
                if !popup
                    DllCall("MapWindowPoints", "Ptr", 0, "Ptr", parent, "Ptr", rect, "UInt", 2)
                enabled := DllCall("IsWindowEnabled", "Ptr", item.handle)
                signature := enabled
                for offset in [8, 12, 16, 20]
                    signature .= "|" NumGet(range, offset, "Int")
                for offset in [0, 4, 8, 12]
                    signature .= "|" NumGet(rect, offset, "Int")
                if bar.signature = signature
                    continue
                bar.signature := signature
                DllCall("SetScrollInfo", "Ptr", bar.handle, "Int", 2, "Ptr", range, "Int", false)
                DllCall("EnableWindow", "Ptr", bar.handle, "Int", enabled)
                DllCall("SetWindowPos", "Ptr", bar.handle, "Ptr", 0,
                    "Int", NumGet(rect, 0, "Int"), "Int", NumGet(rect, 4, "Int"),
                    "Int", NumGet(rect, 8, "Int") - NumGet(rect, 0, "Int"),
                    "Int", NumGet(rect, 12, "Int") - NumGet(rect, 4, "Int"), "UInt", 0x50)
                DllCall("InvalidateRect", "Ptr", bar.handle, "Ptr", 0, "Int", false)
            }
        } finally item.syncingBars := false
    }

    Destroy(item) {
        if !item.HasOwnProp("bars")
            return
        bars := item.bars, item.bars := Map()
        for axis, bar in bars
            if DllCall("IsWindow", "Ptr", bar.handle)
                DllCall("DestroyWindow", "Ptr", bar.handle)
    }
}

class SettingsControlScrollbar {
    __New(chrome, source, axis, parent, popup := false) {
        this.chrome := chrome, this.source := source, this.axis := axis
        this.signature := ""
        this.handle := DllCall("CreateWindowExW", "UInt", popup ? 0x08000080 : 0, "WStr", "SCROLLBAR", "Ptr", 0,
            "UInt", (popup ? 0x84000000 : 0x44000000) | axis, "Int", 0, "Int", 0, "Int", 1, "Int", 1,
            "Ptr", parent, "Ptr", 0, "Ptr", 0, "Ptr", 0, "Ptr")
        if !this.handle
            throw OSError()
        DllCall("uxtheme\SetWindowTheme", "Ptr", this.handle, "WStr", "", "WStr", "")
        ; Prevent the source's native nonclient draws from crossing this sibling.
        style := SettingsNativeChrome.Long(source.handle, -16)
        SettingsNativeChrome.SetLong(source.handle, -16, style | 0x4000000)
    }
    ScrollTo(position) => this.chrome.scrollbars.Scroll(this.source, this.axis, 5, position)
    OnScroll(code, *) => this.chrome.scrollbars.Scroll(this.source, this.axis, code)
}
