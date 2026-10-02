; Optional staggered startup; failure backoff applies in both modes.
class PreviewStartupQueue {
    items := Map()
    nextAt := 0
    __New(now := A_TickCount, slow := false) {
        this.slow := slow
        this.nextAt := now + (slow ? 3000 : 0)
    }
    Observe(hwnd, title, width, height, now := A_TickCount) {
        if !this.items.Has(hwnd)
            this.items[hwnd] := {title: title, w: width, h: height, stableAt: now, retryAt: 0, failures: 0}
        item := this.items[hwnd]
        if item.title != title || item.w != width || item.h != height {
            item.title := title, item.w := width, item.h := height, item.stableAt := now
        }
    }
    Next(now := A_TickCount) {
        if now < this.nextAt
            return 0
        for hwnd, item in this.items {
            if item.w <= 0 || item.h <= 0 || (this.slow && now - item.stableAt < 500) || now < item.retryAt
                continue
            this.nextAt := now + (this.slow ? 150 : 0)
            return hwnd
        }
        return 0
    }
    Failed(hwnd, now := A_TickCount) {
        if this.items.Has(hwnd) {
            item := this.items[hwnd]
            item.failures := Min(item.failures + 1, 6)
            item.retryAt := now + 5000 * item.failures ; Cap retries at once every 30 seconds.
        }
    }
    Remove(hwnd) {
        if this.items.Has(hwnd)
            this.items.Delete(hwnd)
    }
}
