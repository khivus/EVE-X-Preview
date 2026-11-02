; Dark nonclient scrollbars keep the native ranges and scroll commands, but
; capture the pointer locally instead of entering Windows' light tracking loop.
class SettingsScrollbars {
    __New(chrome) {
        this.chrome := chrome, this.drag := 0
        this.repeatTimer := ObjBindMethod(this, "Repeat")
    }
    Handle(item, message, wParam, lParam) {
        if item.kind = "PaneScrollbar" && (message = 0x201 || message = 0x203) {
            this.Begin(item, 1, this.Point(item.handle, lParam))
            return true ; Never fall back to the system's modal light tracking loop.
        }
        if (message = 0xA1 || message = 0xA3) && (wParam = 6 || wParam = 7)
            return this.Begin(item, wParam = 7 ? 1 : 0, this.Point(item.handle, lParam, true))
        if !IsObject(this.drag) || this.drag.item.handle != item.handle
            return false
        switch message {
            case 0x200:
                this.Move(this.Point(item.handle, lParam))
                return true
            case 0x202:
                this.Move(this.Point(item.handle, lParam))
                this.Finish(true)
                return true
            case 0x215:
                if lParam != item.handle
                    this.Finish(false)
            case 0x1F:
                this.Finish(false)
            case 0x100:
                if wParam = 27 {
                    this.Finish(false)
                    return true
                }
        }
        return false
    }
    Point(handle, lParam, screen := false) {
        x := lParam & 0xFFFF, y := (lParam >> 16) & 0xFFFF
        point := Buffer(8, 0)
        NumPut("Int", x & 0x8000 ? x - 0x10000 : x, "Int", y & 0x8000 ? y - 0x10000 : y, point)
        if !screen
            DllCall("ClientToScreen", "Ptr", handle, "Ptr", point)
        return [NumGet(point, 0, "Int"), NumGet(point, 4, "Int")]
    }
    Range(handle, bar) {
        range := Buffer(28, 0)
        NumPut("UInt", 28, "UInt", 7, range)
        nativeBar := this.chrome.windows.Has(handle) && this.chrome.windows[handle].kind = "PaneScrollbar" ? 2 : bar
        if !DllCall("GetScrollInfo", "Ptr", handle, "Int", nativeBar, "Ptr", range)
            return 0
        minimum := NumGet(range, 8, "Int"), page := NumGet(range, 16, "UInt")
        return {minimum: minimum, maximum: Max(minimum, NumGet(range, 12, "Int") - Max(0, page - 1)),
            position: NumGet(range, 20, "Int"), page: page}
    }
    Hit(handle, bar, point) {
        info := this.chrome.ScrollInfo(handle, bar = 1 ? -5 : -6)
        if !IsObject(info) || NumGet(info, 36, "UInt") & 1
            return 0
        left := NumGet(info, 4, "Int"), top := NumGet(info, 8, "Int")
        right := NumGet(info, 12, "Int"), bottom := NumGet(info, 16, "Int")
        if point[1] < left || point[1] >= right || point[2] < top || point[2] >= bottom
            return 0
        origin := bar = 1 ? top : left, length := bar = 1 ? bottom - top : right - left
        coordinate := point[bar = 1 ? 2 : 1] - origin, arrow := NumGet(info, 20, "Int")
        thumbTop := NumGet(info, 24, "Int"), thumbBottom := NumGet(info, 28, "Int")
        code := coordinate < arrow ? 0 : coordinate >= length - arrow ? 1
            : coordinate < thumbTop ? 2 : coordinate >= thumbBottom ? 3 : 5
        return {code: code, origin: origin, length: length, arrow: arrow,
            thumbTop: thumbTop, thumbSize: thumbBottom - thumbTop, coordinate: coordinate}
    }
    Begin(item, bar, point) {
        ; Let grids finish an active cell before its native rows move.
        if item.kind = "ListView"
            for grid in this.chrome.theme.app.grids
                if !grid.disposed && grid.list.Hwnd = item.handle
                    grid.CommitEditor()
        hit := this.Hit(item.handle, bar, point), range := this.Range(item.handle, bar)
        if !IsObject(hit) || !IsObject(range) || range.maximum <= range.minimum
            return false
        this.Finish(false)
        this.drag := {item: item, bar: bar, code: hit.code, position: range.position, point: point,
            grab: hit.coordinate - hit.thumbTop, origin: hit.origin, arrow: hit.arrow,
            travel: Max(1, hit.length - 2 * hit.arrow - hit.thumbSize), range: range,
            previousCapture: DllCall("GetCapture", "Ptr")}
        if this.drag.previousCapture != item.handle
            DllCall("SetCapture", "Ptr", item.handle)
        if DllCall("GetCapture", "Ptr") != item.handle {
            this.drag := 0
            return false
        }
        if hit.code != 5 {
            this.Scroll(item, bar, hit.code)
            SetTimer(this.repeatTimer, -350)
        }
        this.chrome.RepaintNonclient(item.handle)
        return true
    }
    Move(point) {
        if !IsObject(this.drag)
            return
        drag := this.drag, drag.point := point
        if drag.code = 5 {
            fraction := (point[drag.bar = 1 ? 2 : 1] - drag.origin - drag.arrow - drag.grab) / drag.travel
            position := Max(drag.range.minimum, Min(drag.range.maximum,
                drag.range.minimum + Round(fraction * (drag.range.maximum - drag.range.minimum))))
            if position != drag.position {
                drag.position := position
                this.Scroll(drag.item, drag.bar, 5, position)
            }
            return
        }
        this.chrome.RepaintNonclient(drag.item.handle)
    }
    Repeat() {
        if !IsObject(this.drag)
            return
        drag := this.drag
        hit := this.Hit(drag.item.handle, drag.bar, drag.point)
        if IsObject(hit) && hit.code = drag.code
            this.Scroll(drag.item, drag.bar, drag.code)
        if IsObject(this.drag)
            SetTimer(this.repeatTimer, -60)
    }
    Scroll(item, bar, code, position := 0) {
        if !DllCall("IsWindow", "Ptr", item.handle)
            return
        tracking := code = 4 || code = 5
        if item.kind = "PaneScrollbar" {
            if tracking
                item.control.ScrollTo(position)
            else
                item.control.OnScroll(code, item.handle, 0x115, item.control.viewportHwnd)
        } else if tracking && item.kind = "ScrollPane"
            item.control.ScrollTo(position) ; Keep 32-bit pane offsets.
        else if tracking && item.kind = "ListView" {
            range := this.Range(item.handle, bar)
            if !IsObject(range)
                return
            delta := position - range.position
            if bar = 1 {
                rect := Buffer(16, 0)
                NumPut("Int", 0, rect)
                DllCall("SendMessageW", "Ptr", item.handle, "UInt", 0x100E, "Ptr", 0, "Ptr", rect)
                delta *= Max(1, NumGet(rect, 12, "Int") - NumGet(rect, 4, "Int"))
            }
            DllCall("SendMessageW", "Ptr", item.handle, "UInt", 0x1014, "Ptr", bar = 0 ? delta : 0, "Ptr", bar = 1 ? delta : 0)
        } else if tracking && bar = 1 && (item.kind = "ListBox" || item.kind = "ComboList")
            DllCall("SendMessageW", "Ptr", item.handle, "UInt", 0x197, "Ptr", position, "Ptr", 0) ; LB_SETTOPINDEX
        else if tracking && bar = 1 && item.kind = "Edit" {
            range := this.Range(item.handle, bar)
            if IsObject(range)
                DllCall("SendMessageW", "Ptr", item.handle, "UInt", 0xB6, "Ptr", 0, "Ptr", position - range.position) ; EM_LINESCROLL
        } else
            DllCall("SendMessageW", "Ptr", item.handle, "UInt", bar = 1 ? 0x115 : 0x114,
                "Ptr", code | ((position & 0xFFFF) << 16), "Ptr", 0)
        this.chrome.RepaintNonclient(item.handle)
    }
    Pressed(handle, bar, code) => IsObject(this.drag) && this.drag.item.handle = handle && this.drag.bar = bar && this.drag.code = code
    Finish(commit, notify := true) {
        if !IsObject(this.drag)
            return
        drag := this.drag, this.drag := 0
        SetTimer(this.repeatTimer, 0)
        if DllCall("GetCapture", "Ptr") = drag.item.handle {
            ; An open combo owns capture before its scrollbar is pressed.
            ; Preserve it so releasing the thumb does not close the list.
            if drag.item.kind = "ComboList" && drag.previousCapture && DllCall("IsWindow", "Ptr", drag.previousCapture) {
                if drag.previousCapture != drag.item.handle
                    DllCall("SetCapture", "Ptr", drag.previousCapture)
            } else
                DllCall("ReleaseCapture")
        }
        if notify && DllCall("IsWindow", "Ptr", drag.item.handle) {
            if commit && drag.code = 5
                this.Scroll(drag.item, drag.bar, 4, drag.position)
            this.Scroll(drag.item, drag.bar, 8)
        }
        this.chrome.RepaintNonclient(drag.item.handle)
    }
    Dispose() {
        this.Finish(false)
        SetTimer(this.repeatTimer, 0)
    }
}
