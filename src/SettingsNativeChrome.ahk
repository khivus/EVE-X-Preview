; Recolor native chrome locally; native window procedures retain input,
; accessibility, selection, scrolling and spinner behavior.
class SettingsNativeChrome {
    __New(theme) {
        this.theme := theme, this.windows := Map()
        this.scrollbars := SettingsScrollbars(this)
        this.controlScrollbars := SettingsControlScrollbars(this)
        this.callback := CallbackCreate(ObjBindMethod(this, "WindowProc"),, 6)
    }
    static Long(handle, index) => DllCall(A_PtrSize = 8 ? "GetWindowLongPtrW" : "GetWindowLongW", "Ptr", handle, "Int", index, "Ptr")
    static SetLong(handle, index, value) => DllCall(A_PtrSize = 8 ? "SetWindowLongPtrW" : "SetWindowLongW", "Ptr", handle, "Int", index, "Ptr", value, "Ptr")
    Attach(control) {
        kind := control.Type
        if kind = "Text" && control.ThemePreserve && SettingsNativeChrome.Long(control.Hwnd, -16) & 0x800000 {
            this.AttachWindow(control.Hwnd, "Swatch", control)
            return
        }
        if !InStr("|Edit|DDL|ComboBox|ListBox|ListView|TreeView|UpDown|", "|" kind "|")
            return
        this.AttachWindow(control.Hwnd, kind, control)
        if kind = "DDL" || kind = "ComboBox"
            this.ComboList(control)
    }
    AttachWindow(handle, kind, control := 0) {
        if !handle || !DllCall("IsWindow", "Ptr", handle)
            return
        if !this.windows.Has(handle) {
            this.windows[handle] := {handle: handle, kind: kind, control: control,
                border: SettingsNativeChrome.Long(handle, -16) & 0x800000,
                edge: SettingsNativeChrome.Long(handle, -20) & 0x200}
            DllCall("comctl32\SetWindowSubclass", "Ptr", handle, "Ptr", this.callback, "UPtr", 1, "UPtr", 0)
        }
        item := this.windows[handle]
        if kind != "ScrollPane" && kind != "PaneScrollbar" && kind != "UpDown" {
            style := SettingsNativeChrome.Long(handle, -16), extended := SettingsNativeChrome.Long(handle, -20)
            newStyle := (style & ~0x800000) | (this.theme.dark ? kind = "ComboEdit" ? 0 : 0x800000 : item.border)
            newExtended := (extended & ~0x200) | (this.theme.dark ? 0 : item.edge)
            if newStyle != style || newExtended != extended {
                SettingsNativeChrome.SetLong(handle, -16, newStyle)
                SettingsNativeChrome.SetLong(handle, -20, newExtended)
                DllCall("SetWindowPos", "Ptr", handle, "Ptr", 0, "Int", 0, "Int", 0, "Int", 0, "Int", 0, "UInt", 0x37)
            }
        }
        DllCall("RedrawWindow", "Ptr", handle, "Ptr", 0, "Ptr", 0, "UInt", 0x585)
        this.controlScrollbars.Sync(item)
    }
    ComboInfo(handle) {
        info := Buffer(A_PtrSize = 8 ? 64 : 52, 0)
        NumPut("UInt", info.Size, info)
        return DllCall("GetComboBoxInfo", "Ptr", handle, "Ptr", info) ? info : 0
    }
    ComboList(control) {
        info := this.ComboInfo(control.Hwnd)
        if IsObject(info) {
            edit := NumGet(info, A_PtrSize = 8 ? 48 : 44, "Ptr")
            if edit && edit != control.Hwnd
                this.AttachWindow(edit, "ComboEdit", control)
            list := NumGet(info, A_PtrSize = 8 ? 56 : 48, "Ptr")
            this.AttachWindow(list, "ComboList", control)
        }
    }
    WindowProc(handle, message, wParam, lParam, subclass, data) {
        if message = 0xB && this.windows.Has(handle)
            this.windows[handle].redrawDisabled := !wParam
        if message = 0x82 {
            if IsObject(this.scrollbars.drag) && this.scrollbars.drag.item.handle = handle
                this.scrollbars.Finish(false, false)
            if this.windows.Has(handle)
                this.controlScrollbars.Destroy(this.windows[handle])
            this.windows.Delete(handle)
            DllCall("comctl32\RemoveWindowSubclass", "Ptr", handle, "Ptr", this.callback, "UPtr", 1)
        }
        if this.windows.Has(handle) && this.windows[handle].kind = "ComboList" && message = 0x46 {
            this.ConstrainPopup(this.windows[handle], lParam)
            return 0
        }
        if this.windows.Has(handle) && (message = 0x18 || message = 0xA)
            this.controlScrollbars.Sync(this.windows[handle])
        if !this.theme.dark || !this.windows.Has(handle)
            return DllCall("comctl32\DefSubclassProc", "Ptr", handle, "UInt", message, "Ptr", wParam, "Ptr", lParam, "Ptr")
        item := this.windows[handle]
        if this.scrollbars.Handle(item, message, wParam, lParam)
            return 0
        if item.kind = "PaneScrollbar" {
            if message = 0x14
                return 1
            if message = 0xF || message = 0x317 || message = 0x318 {
                if message = 0xF {
                    paint := Buffer(A_PtrSize = 8 ? 72 : 64, 0)
                    dc := DllCall("BeginPaint", "Ptr", handle, "Ptr", paint, "Ptr")
                } else
                    dc := wParam
                this.PaintNonclient(handle, dc)
                if message = 0xF
                    DllCall("EndPaint", "Ptr", handle, "Ptr", paint)
                return 0
            }
            if message = 0x5 || message = 0x200 || message = 0x2A3 {
                DllCall("InvalidateRect", "Ptr", handle, "Ptr", 0, "Int", false)
                return 0
            }
            if message = 0xE9 { ; SBM_SETSCROLLINFO: update the native accessible range without its light paint.
                result := DllCall("comctl32\DefSubclassProc", "Ptr", handle, "UInt", message, "Ptr", 0, "Ptr", lParam, "Ptr")
                DllCall("InvalidateRect", "Ptr", handle, "Ptr", 0, "Int", false)
                return result
            }
            if message = 0xEB && lParam && NumGet(lParam, 0, "UInt") >= 60 { ; SBM_GETSCROLLBARINFO
                info := this.ScrollInfo(handle, -5)
                if IsObject(info) {
                    DllCall("RtlMoveMemory", "Ptr", lParam, "Ptr", info, "UPtr", 60)
                    return true
                }
            }
            if message = 0x100 && InStr("|33|34|35|36|37|38|39|40|", "|" wParam "|") {
                code := Map(37, 0, 38, 0, 39, 1, 40, 1, 33, 2, 34, 3, 36, 6, 35, 7)[wParam]
                this.scrollbars.Scroll(item, item.HasOwnProp("axis") ? item.axis : 1, code)
                return 0
            }
        }
        ; Default nonclient painting and its modal scrollbar tracking use the
        ; system light palette. Own these surfaces before Windows draws them.
        if message = 0x85 || (message = 0xA0 && (wParam = 6 || wParam = 7)) {
            this.RepaintNonclient(handle)
            return 0
        }
        if item.kind = "UpDown" {
            if message = 0x201 || message = 0x203 {
                rect := Buffer(16, 0)
                DllCall("GetClientRect", "Ptr", handle, "Ptr", rect)
                item.pressed := ((lParam >> 16) & 0xFFFF) < NumGet(rect, 12, "Int") // 2 ? 1 : 2
            } else if message = 0x202 || message = 0x215 || message = 0x1F
                item.pressed := 0
        }
        if message = 0xF && item.kind = "UpDown" {
            paint := Buffer(A_PtrSize = 8 ? 72 : 64, 0)
            dc := DllCall("BeginPaint", "Ptr", handle, "Ptr", paint, "Ptr")
            this.PaintSpin(item, dc)
            DllCall("EndPaint", "Ptr", handle, "Ptr", paint)
            return 0
        }
        if this.ScrollUpdateMessage(item, message, wParam, lParam) && DllCall("IsWindowVisible", "Ptr", handle)
            && !DllCall("GetPropW", "Ptr", handle, "WStr", "SysSetRedraw", "Ptr")
            && (!item.HasOwnProp("redrawDisabled") || !item.redrawDisabled)
            && (!item.HasOwnProp("scrollUpdating") || !item.scrollUpdating)
            result := this.UpdateScrollControl(item, message, wParam, lParam)
        else
            result := DllCall("comctl32\DefSubclassProc", "Ptr", handle, "UInt", message, "Ptr", wParam, "Ptr", lParam, "Ptr")
        if message = 0x14F && (item.kind = "DDL" || item.kind = "ComboBox")
            this.ComboList(item.control)
        if message = 0xF && (item.kind = "DDL" || item.kind = "ComboBox" || item.kind = "ComboEdit" || item.kind = "Tooltip") {
            dc := DllCall("GetDC", "Ptr", handle, "Ptr")
            this.PaintClient(item, dc)
            DllCall("ReleaseDC", "Ptr", handle, "Ptr", dc)
            this.RepaintNonclient(handle)
        } else if message = 0x317 || message = 0x318 {
            saved := DllCall("SaveDC", "Ptr", wParam, "Int")
            if message = 0x317 && lParam & 2 {
                window := Buffer(16, 0), client := Buffer(8, 0)
                DllCall("GetWindowRect", "Ptr", handle, "Ptr", window)
                DllCall("ClientToScreen", "Ptr", handle, "Ptr", client)
                DllCall("OffsetViewportOrgEx", "Ptr", wParam, "Int", NumGet(client, 0, "Int") - NumGet(window, 0, "Int"),
                    "Int", NumGet(client, 4, "Int") - NumGet(window, 4, "Int"), "Ptr", 0)
            }
            this.PaintClient(item, wParam)
            DllCall("RestoreDC", "Ptr", wParam, "Int", saved)
            if message = 0x317 && lParam & 2
                this.PaintNonclient(handle, wParam)
        }
        if message = 0x85 || message = 0x115 || message = 0x114 || message = 0x20A
            || message = 0xA0 || message = 0xA1 || message = 0xA2 || message = 0x2A2
            || message = 0x1013 || message = 0x1014 || message = 0x1060 || message = 0x5 || message = 0x47 || message = 0x7D
            this.RepaintNonclient(handle)
        if this.windows.Has(handle)
            this.controlScrollbars.Sync(item)
        if (item.kind = "UpDown" || item.kind = "DDL" || item.kind = "ComboBox")
            && (message = 0xA || message = 0x30 || message = 0x31A || message = 0x14F || message = 0x201 || message = 0x202)
            DllCall("RedrawWindow", "Ptr", handle, "Ptr", 0, "Ptr", 0, "UInt", 0x141)
        return result
    }
    ScrollUpdateMessage(item, message, key, point := 0) {
        ; TreeView refuses to move its visible rows while WM_SETREDRAW hides it.
        ; The separate scrollbar covers the system track throughout the update.
        if item.kind = "TreeView" && item.HasOwnProp("bars") && item.bars.Count
            return false
        if !InStr("|ListView|TreeView|ListBox|ComboList|Edit|", "|" item.kind "|")
            return false
        if message = 0x115 || message = 0x114 || message = 0x20A || message = 0x20E || message = 0x5 || message = 0x46 || message = 0x47
            return true
        if message = 0x100
            return key >= 0x21 && key <= 0x28 ; Page/Home/End/arrows.
        if item.kind = "ListView"
            return message = 0x1014 ; LVM_SCROLL; ordinary selection retains its existing paint path.
        if item.kind = "Edit"
            return message = 0xB6 || message = 0xC2 || message = 0xC || message = 0xB7 ; Scroll / replace / set / caret.
        if item.kind = "ListBox" || item.kind = "ComboList"
            return message = 0x197 || message = 0x186 || message = 0x18B ; Top index / selection / count.
        if message = 0x201 || message = 0x203 {
            hit := Buffer(A_PtrSize = 8 ? 24 : 16, 0)
            NumPut("Int", (point & 0xFFFF) << 48 >> 48, "Int", ((point >> 16) & 0xFFFF) << 48 >> 48, hit)
            DllCall("SendMessageW", "Ptr", item.handle, "UInt", 0x1111, "Ptr", 0, "Ptr", hit)
            flags := NumGet(hit, 8, "UInt")
            return !!(flags & 0x10) || (message = 0x203 && !!(flags & 0x46))
        }
        return message = 0x1102 || message = 0x110B || message = 0x1114 ; Expand / select / ensure visible.
    }
    UpdateScrollControl(item, message, wParam, lParam) {
        ; Common controls update and paint their system scrollbar from inside
        ; scrolling/expansion, before WM_NCPAINT. Suppress that intermediate
        ; draw locally, retaining native ranges, input and layout calculations.
        handle := item.handle, item.scrollUpdating := true
        DllCall("SendMessageW", "Ptr", handle, "UInt", 0xB, "Ptr", 0, "Ptr", 0)
        try return DllCall("comctl32\DefSubclassProc", "Ptr", handle, "UInt", message, "Ptr", wParam, "Ptr", lParam, "Ptr")
        finally {
            if DllCall("IsWindow", "Ptr", handle) {
                DllCall("SendMessageW", "Ptr", handle, "UInt", 0xB, "Ptr", 1, "Ptr", 0)
                item.scrollUpdating := false
                if !this.theme.updates.Count {
                    ; Do not erase or invalidate the header child while scrolling.
                    rect := Buffer(16, 0)
                    DllCall("GetClientRect", "Ptr", handle, "Ptr", rect)
                    if item.kind = "ListView" {
                        header := DllCall("SendMessageW", "Ptr", handle, "UInt", 0x101F, "Ptr", 0, "Ptr", 0, "Ptr")
                        if header && SettingsNativeChrome.Long(header, -16) & 0x10000000 {
                            bounds := Buffer(16, 0)
                            DllCall("GetWindowRect", "Ptr", header, "Ptr", bounds)
                            DllCall("MapWindowPoints", "Ptr", 0, "Ptr", handle, "Ptr", bounds, "UInt", 2)
                            NumPut("Int", Max(0, NumGet(bounds, 12, "Int")), rect, 4)
                        }
                    }
                    if NumGet(rect, 12, "Int") > NumGet(rect, 4, "Int")
                        DllCall("RedrawWindow", "Ptr", handle, "Ptr", rect, "Ptr", 0, "UInt", item.kind = "TreeView" ? 0x145 : 0x161)
                    this.RepaintNonclient(handle)
                }
            } else
                item.scrollUpdating := false
        }
    }
    ConstrainPopup(item, position) {
        if !item.control || !DllCall("IsWindow", "Ptr", item.control.Hwnd)
            return
        offset := A_PtrSize * 2, flags := NumGet(position, offset + 16, "UInt")
        if flags & 0x80 ; Hiding a list must not move or resize it.
            return
        root := DllCall("GetAncestor", "Ptr", item.control.Hwnd, "UInt", 2, "Ptr")
        bounds := Buffer(16, 0), control := Buffer(16, 0), current := Buffer(16, 0)
        DllCall("GetClientRect", "Ptr", root, "Ptr", bounds)
        DllCall("MapWindowPoints", "Ptr", root, "Ptr", 0, "Ptr", bounds, "UInt", 2)
        DllCall("GetWindowRect", "Ptr", item.control.Hwnd, "Ptr", control)
        DllCall("GetWindowRect", "Ptr", item.handle, "Ptr", current)
        left := NumGet(bounds, 0, "Int"), top := NumGet(bounds, 4, "Int")
        right := NumGet(bounds, 8, "Int"), bottom := NumGet(bounds, 12, "Int")
        monitor := Buffer(40, 0)
        NumPut("UInt", 40, monitor)
        if DllCall("GetMonitorInfoW", "Ptr", DllCall("MonitorFromWindow", "Ptr", root, "UInt", 2, "Ptr"), "Ptr", monitor) {
            workLeft := Max(left, NumGet(monitor, 20, "Int")), workTop := Max(top, NumGet(monitor, 24, "Int"))
            workRight := Min(right, NumGet(monitor, 28, "Int")), workBottom := Min(bottom, NumGet(monitor, 32, "Int"))
            if workRight > workLeft && workBottom > workTop
                left := workLeft, top := workTop, right := workRight, bottom := workBottom
        }
        if right <= left || bottom <= top
            return
        width := flags & 1 ? NumGet(current, 8, "Int") - NumGet(current, 0, "Int") : NumGet(position, offset + 8, "Int")
        height := flags & 1 ? NumGet(current, 12, "Int") - NumGet(current, 4, "Int") : NumGet(position, offset + 12, "Int")
        fieldLeft := NumGet(control, 0, "Int"), fieldRight := NumGet(control, 8, "Int")
        width := Max(1, Min(width, right - left, Max(fieldRight - fieldLeft, right - Max(left, fieldLeft))))
        x := Max(left, Min(fieldLeft, right - width))
        below := Max(0, bottom - NumGet(control, 12, "Int")), above := Max(0, NumGet(control, 4, "Int") - top)
        openBelow := height <= below || below >= above
        height := Max(1, Min(height, openBelow ? below : above, bottom - top))
        y := openBelow ? NumGet(control, 12, "Int") : NumGet(control, 4, "Int") - height
        y := Max(top, Min(y, bottom - height))
        NumPut("Int", x, "Int", y, "Int", width, "Int", height, position, offset)
        NumPut("UInt", flags & ~3, position, offset + 16)
    }
    PaintClient(item, dc) {
        if item.kind = "UpDown"
            this.PaintSpin(item, dc)
        else if item.kind = "DDL" || item.kind = "ComboBox"
            this.PaintCombo(item, dc)
        else if item.kind = "ComboEdit" {
            rect := Buffer(16, 0)
            DllCall("GetClientRect", "Ptr", item.handle, "Ptr", rect)
            SettingsTheme.Fill(dc, rect, this.theme.input, true)
        }
        else if item.kind = "Tooltip" {
            rect := Buffer(16, 0)
            DllCall("GetClientRect", "Ptr", item.handle, "Ptr", rect)
            SettingsTheme.Fill(dc, rect, this.theme.line, true)
            DllCall("InflateRect", "Ptr", rect, "Int", -1, "Int", -1)
            SettingsTheme.Fill(dc, rect, this.theme.input, true)
        }
    }
    PaintCombo(item, dc) {
        info := this.ComboInfo(item.handle)
        if !IsObject(info)
            return
        whole := Buffer(16, 0)
        DllCall("GetClientRect", "Ptr", item.handle, "Ptr", whole)
        saved := DllCall("SaveDC", "Ptr", dc, "Int")
        DllCall("ExcludeClipRect", "Ptr", dc, "Int", NumGet(info, 4, "Int"), "Int", NumGet(info, 8, "Int"),
            "Int", NumGet(info, 12, "Int"), "Int", NumGet(info, 16, "Int"))
        SettingsTheme.Fill(dc, whole, this.theme.input)
        DllCall("RestoreDC", "Ptr", dc, "Int", saved)
        rect := Buffer(16, 0)
        DllCall("RtlMoveMemory", "Ptr", rect, "Ptr", info.Ptr + 20, "UPtr", 16)
        SettingsTheme.Fill(dc, rect, NumGet(info, 36, "UInt") & 8 ? "3D4859" : "30343B")
        this.Arrow(dc, rect, "down", item.control.Enabled ? this.theme.foreground : this.theme.muted)
        SettingsTheme.Fill(dc, whole, this.theme.line, true)
        DllCall("InflateRect", "Ptr", whole, "Int", -1, "Int", -1)
        SettingsTheme.Fill(dc, whole, this.theme.input, true)
    }
    PaintSpin(item, dc) {
        rect := Buffer(16, 0)
        DllCall("GetClientRect", "Ptr", item.handle, "Ptr", rect)
        width := NumGet(rect, 8, "Int"), height := NumGet(rect, 12, "Int")
        horizontal := SettingsNativeChrome.Long(item.handle, -16) & 0x40
        for index, direction in horizontal ? ["left", "right"] : ["up", "down"] {
            part := Buffer(16, 0)
            NumPut("Int", horizontal && index = 2 ? width // 2 : 0, "Int", !horizontal && index = 2 ? height // 2 : 0,
                "Int", horizontal && index = 1 ? width // 2 : width, "Int", !horizontal && index = 1 ? height // 2 : height, part)
            SettingsTheme.Fill(dc, part, !item.control.Enabled ? this.theme.input : item.HasOwnProp("pressed") && item.pressed = index ? "3D4859" : "30343B")
            if item.control.Enabled
                SettingsTheme.Fill(dc, part, this.theme.line, true)
            this.Arrow(dc, part, direction, item.control.Enabled ? this.theme.foreground : this.theme.muted)
        }
    }
    Arrow(dc, rect, direction, color) {
        x := (NumGet(rect, 0, "Int") + NumGet(rect, 8, "Int")) // 2
        y := (NumGet(rect, 4, "Int") + NumGet(rect, 12, "Int")) // 2
        size := Max(2, Round(3 * A_ScreenDPI / 96)), points := Buffer(24, 0)
        coordinates := direction = "up" ? [x-size, y+1, x, y-size+1, x+size, y+1]
            : direction = "down" ? [x-size, y-1, x, y+size-1, x+size, y-1]
            : direction = "left" ? [x+1, y-size, x-size+1, y, x+1, y+size]
            : [x-1, y-size, x+size-1, y, x-1, y+size]
        for index, value in coordinates
            NumPut("Int", value, points, (index - 1) * 4)
        pen := DllCall("CreatePen", "Int", 0, "Int", Max(1, Round(A_ScreenDPI / 96)), "UInt", SettingsTheme.BGR(color), "Ptr")
        previous := DllCall("SelectObject", "Ptr", dc, "Ptr", pen, "Ptr")
        DllCall("Polyline", "Ptr", dc, "Ptr", points, "Int", 3)
        DllCall("SelectObject", "Ptr", dc, "Ptr", previous)
        DllCall("DeleteObject", "Ptr", pen)
    }
    ScrollInfo(handle, object) {
        paneBar := this.windows.Has(handle) && this.windows[handle].kind = "PaneScrollbar"
        if !paneBar && this.windows.Has(handle) {
            item := this.windows[handle], axis := object = -5 ? 1 : 0
            if item.HasOwnProp("bars") && item.bars.Has(axis) && DllCall("IsWindowVisible", "Ptr", item.bars[axis].handle)
                return this.ScrollInfo(item.bars[axis].handle, object)
        }
        if paneBar {
            axis := this.windows[handle].HasOwnProp("axis") ? this.windows[handle].axis : 1
            if object != (axis = 1 ? -5 : -6) || !(SettingsNativeChrome.Long(handle, -16) & 0x10000000)
                return 0
            range := this.scrollbars.Range(handle, axis)
            if !IsObject(range)
                return 0
            rect := Buffer(16, 0), info := Buffer(60, 0)
            DllCall("GetWindowRect", "Ptr", handle, "Ptr", rect)
            height := Max(1, axis = 1 ? NumGet(rect, 12, "Int") - NumGet(rect, 4, "Int") : NumGet(rect, 8, "Int") - NumGet(rect, 0, "Int"))
            arrow := Min(DllCall("GetSystemMetrics", "Int", 20, "Int"), height // 2)
            track := Max(0, height - 2 * arrow)
            thumb := Min(track, Max(Round(8 * A_ScreenDPI / 96), Round(track * range.page / Max(1, range.maximum - range.minimum + range.page))))
            top := arrow + Round((track - thumb) * (range.position - range.minimum) / Max(1, range.maximum - range.minimum))
            NumPut("UInt", info.Size, info)
            DllCall("RtlMoveMemory", "Ptr", info.Ptr + 4, "Ptr", rect, "UPtr", 16)
            NumPut("Int", arrow, "Int", top, "Int", top + thumb, info, 20)
            return info
        }
        return this.NativeScrollInfo(handle, object)
    }
    NativeScrollInfo(handle, object) {
        info := Buffer(60, 0)
        NumPut("UInt", info.Size, info)
        return DllCall("GetScrollBarInfo", "Ptr", handle, "Int", object, "Ptr", info)
            && !(NumGet(info, 36, "UInt") & 0x18000) ? info : 0
    }
    PaintNonclient(handle, dc) {
        window := Buffer(16, 0), rect := Buffer(16, 0)
        DllCall("GetWindowRect", "Ptr", handle, "Ptr", window)
        width := NumGet(window, 8, "Int") - NumGet(window, 0, "Int")
        height := NumGet(window, 12, "Int") - NumGet(window, 4, "Int")
        NumPut("Int", width, rect, 8), NumPut("Int", height, rect, 12)
        if SettingsNativeChrome.Long(handle, -16) & 0x800000
            SettingsTheme.Fill(dc, rect, this.theme.line, true)
        for object in [-5, -6] {
            info := this.ScrollInfo(handle, object)
            if !IsObject(info)
                continue
            vertical := object = -5
            left := NumGet(info, 4, "Int") - NumGet(window, 0, "Int")
            top := NumGet(info, 8, "Int") - NumGet(window, 4, "Int")
            right := NumGet(info, 12, "Int") - NumGet(window, 0, "Int")
            bottom := NumGet(info, 16, "Int") - NumGet(window, 4, "Int")
            ; Present the track, arrows and thumb together, so a compositor
            ; cannot show an empty dark track between individual GDI draws.
            memory := DllCall("CreateCompatibleDC", "Ptr", dc, "Ptr")
            bitmap := DllCall("CreateCompatibleBitmap", "Ptr", dc, "Int", Max(1, right - left), "Int", Max(1, bottom - top), "Ptr")
            buffered := memory && bitmap
            target := buffered ? memory : dc
            if buffered {
                previous := DllCall("SelectObject", "Ptr", memory, "Ptr", bitmap, "Ptr")
                DllCall("SetViewportOrgEx", "Ptr", memory, "Int", -left, "Int", -top, "Ptr", 0)
            }
            try {
                NumPut("Int", left, "Int", top, "Int", right, "Int", bottom, rect)
                SettingsTheme.Fill(target, rect, "25282D")
                button := NumGet(info, 20, "Int")
                for index, direction in vertical ? ["up", "down"] : ["left", "right"] {
                    part := Buffer(16, 0)
                    NumPut("Int", !vertical && index = 2 ? right - button : left, "Int", vertical && index = 2 ? bottom - button : top,
                        "Int", !vertical && index = 1 ? left + button : right, "Int", vertical && index = 1 ? top + button : bottom, part)
                    pressed := this.scrollbars.Pressed(handle, vertical ? 1 : 0, index - 1)
                    SettingsTheme.Fill(target, part, pressed ? "3D4859" : "30343B")
                    this.Arrow(target, part, direction, this.theme.muted)
                }
                thumbTop := NumGet(info, 24, "Int"), thumbBottom := NumGet(info, 28, "Int")
                if thumbBottom > thumbTop {
                    NumPut("Int", vertical ? left + 3 : left + thumbTop, "Int", vertical ? top + thumbTop : top + 3,
                        "Int", vertical ? right - 3 : left + thumbBottom, "Int", vertical ? top + thumbBottom : bottom - 3, rect)
                    SettingsTheme.Fill(target, rect, this.scrollbars.Pressed(handle, vertical ? 1 : 0, 5) ? "8B9BAD" : "66717E")
                }
            } finally {
                if buffered {
                    DllCall("SetViewportOrgEx", "Ptr", memory, "Int", 0, "Int", 0, "Ptr", 0)
                    DllCall("BitBlt", "Ptr", dc, "Int", left, "Int", top, "Int", right - left, "Int", bottom - top,
                        "Ptr", memory, "Int", 0, "Int", 0, "UInt", 0xCC0020)
                    DllCall("SelectObject", "Ptr", memory, "Ptr", previous)
                }
                if bitmap
                    DllCall("DeleteObject", "Ptr", bitmap)
                if memory
                    DllCall("DeleteDC", "Ptr", memory)
            }
        }
        vertical := this.ScrollInfo(handle, -5), horizontal := this.ScrollInfo(handle, -6)
        if IsObject(vertical) && IsObject(horizontal) {
            NumPut("Int", NumGet(vertical, 4, "Int") - NumGet(window, 0, "Int"),
                "Int", NumGet(horizontal, 8, "Int") - NumGet(window, 4, "Int"),
                "Int", NumGet(vertical, 12, "Int") - NumGet(window, 0, "Int"),
                "Int", NumGet(horizontal, 16, "Int") - NumGet(window, 4, "Int"), rect)
            SettingsTheme.Fill(dc, rect, "25282D")
        }
    }
    RepaintNonclient(handle) {
        if !this.theme.dark || this.theme.updates.Count || !DllCall("IsWindow", "Ptr", handle)
            return
        if !this.windows.Has(handle) || (this.windows[handle].HasOwnProp("painting") && this.windows[handle].painting)
            || (this.windows[handle].HasOwnProp("scrollUpdating") && this.windows[handle].scrollUpdating)
            return
        item := this.windows[handle], item.painting := true
        dc := DllCall("GetWindowDC", "Ptr", handle, "Ptr")
        try this.PaintNonclient(handle, dc)
        finally {
            DllCall("ReleaseDC", "Ptr", handle, "Ptr", dc)
            item.painting := false
        }
    }
    Dispose() {
        this.scrollbars.Dispose()
        items := []
        for handle, item in this.windows
            items.Push(item)
        for item in items
            this.controlScrollbars.Destroy(item)
        for handle in this.windows
            if DllCall("IsWindow", "Ptr", handle)
                DllCall("comctl32\RemoveWindowSubclass", "Ptr", handle, "Ptr", this.callback, "UPtr", 1)
        this.windows.Clear()
        CallbackFree(this.callback)
    }
}
