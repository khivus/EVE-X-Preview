; Screen rectangles use physical desktop coordinates, including negative ones.
class ThumbnailRecovery {
    static Monitors() {
        monitors := []
        loop MonitorGetCount() {
            MonitorGet(A_Index, &left, &top, &right, &bottom)
            monitors.Push({left: left, top: top, right: right, bottom: bottom})
        }
        return monitors
    }

    static Position(x, y, width, height, monitors) {
        if width <= 0 || height <= 0 || !monitors.Length
            return false
        if this.IsCovered(x, y, width, height, monitors)
            return false

        best := false, bestFits := false, bestDistance := 0
        for monitor in monitors {
            fits := width <= monitor.right - monitor.left && height <= monitor.bottom - monitor.top
            newX := Max(monitor.left, Min(x, monitor.right - width))
            newY := Max(monitor.top, Min(y, monitor.bottom - height))
            distance := (newX - x) ** 2 + (newY - y) ** 2
            if !best || (fits && !bestFits) || (fits = bestFits && distance < bestDistance) {
                best := {x: newX, y: newY}
                bestFits := fits, bestDistance := distance
            }
        }
        ; Oversized previews retain their size, with their top-left corner reachable.
        return best
    }

    static IsCovered(x, y, width, height, monitors) {
        uncovered := [{left: x, top: y, right: x + width, bottom: y + height}]
        for monitor in monitors {
            remaining := []
            for rect in uncovered {
                left := Max(rect.left, monitor.left), top := Max(rect.top, monitor.top)
                right := Min(rect.right, monitor.right), bottom := Min(rect.bottom, monitor.bottom)
                if left >= right || top >= bottom {
                    remaining.Push(rect)
                    continue
                }
                ; Subtract this monitor, retaining only the still-uncovered pieces.
                if rect.top < top
                    remaining.Push({left: rect.left, top: rect.top, right: rect.right, bottom: top})
                if bottom < rect.bottom
                    remaining.Push({left: rect.left, top: bottom, right: rect.right, bottom: rect.bottom})
                if rect.left < left
                    remaining.Push({left: rect.left, top: top, right: left, bottom: bottom})
                if right < rect.right
                    remaining.Push({left: right, top: top, right: rect.right, bottom: bottom})
            }
            uncovered := remaining
            if !uncovered.Length
                return true
        }
        return false
    }
}
