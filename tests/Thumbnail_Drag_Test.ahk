class ThumbnailDragFixture extends ThumbnailNamesFixture {
    __New() {
        super.__New()
        this.guis := []
        this.ThumbWindows := {}
        this.ThumbHwnd_EvEHwnd := Map()
        this.positions := [[100, 100], [120, 130], [120, 130], [150, 160]]
        this.positionIndex := 0
        this.moves := 0
        for id in [1, 2] {
            stack := Map()
            for type in ["Window", "TextOverlay", "Border"] {
                previewGui := Gui("-Caption")
                previewGui.Show("Hide x" (100 * id) " y" (100 * id) " w100 h80")
                stack[type] := previewGui
                this.guis.Push(previewGui)
                this.ThumbHwnd_EvEHwnd[previewGui.Hwnd] := id
            }
            this.ThumbWindows.%id% := stack
        }
    }
    MouseDragPosition(&x, &y) {
        point := this.positions[++this.positionIndex]
        x := point[1], y := point[2]
    }
    MouseDragButtonsHeld() {
        return this.positionIndex < this.positions.Length
    }
    MoveThumbnailDragWindows(windows, dx, dy) {
        this.moves++
        AssertEqual(0, this.timerHits, "Timers must not interrupt movement, including its Sleep calls.")
        if this.failMove
            throw Error("Simulated drag failure")
        super.MoveThumbnailDragWindows(windows, dx, dy)
    }
    Cleanup() {
        for previewGui in this.guis
            previewGui.Destroy()
    }
}

TestRunner.Register("Dragging single and grouped thumbnails preserves offsets and defers timers", ThumbnailDragTest)
ThumbnailDragTest() {
    previousHidden := A_DetectHiddenWindows
    DetectHiddenWindows(true)
    try {
        for moveAll in [false, true] {
            app := ThumbnailDragFixture()
            app.timerHits := 0, app.failMove := false
            timer := (*) => app.timerHits++
            try {
                source := app.ThumbWindows.%1%["Window"].Hwnd
                windows := app.CaptureThumbnailDragWindows(1, true)
                SetTimer(timer, 5)
                app.Mouse_DragMove(source, moveAll)
                AssertEqual(2, app.moves, "Stationary pointer samples must not move the GUI stack again.")
                for id, thumb in app.ThumbWindows.OwnProps() {
                    for type, previewGui in thumb {
                        WinGetPos(&x, &y, , , previewGui.Hwnd)
                        deltaX := (moveAll || id = 1) ? 50 : 0
                        deltaY := (moveAll || id = 1) ? 60 : 0
                        AssertEqual(windows[previewGui.Hwnd].x + deltaX, x)
                        AssertEqual(windows[previewGui.Hwnd].y + deltaY, y)
                    }
                }
                Sleep(30)
                AssertTrue(app.timerHits > 0, "Timers must resume immediately after dragging ends.")
            } finally {
                SetTimer(timer, 0)
                app.Cleanup()
            }
        }
    } finally DetectHiddenWindows(previousHidden)
}

TestRunner.Register("Dragging restores timers after an error", ThumbnailDragFailureTest)
ThumbnailDragFailureTest() {
    previousHidden := A_DetectHiddenWindows
    DetectHiddenWindows(true)
    app := ThumbnailDragFixture()
    app.timerHits := 0, app.failMove := true
    timer := (*) => app.timerHits++
    try {
        SetTimer(timer, 5)
        AssertThrows(() => app.Mouse_DragMove(app.ThumbWindows.%1%["Window"].Hwnd), "Simulated drag failure")
        Sleep(30)
        AssertTrue(app.timerHits > 0)
    } finally {
        SetTimer(timer, 0)
        app.Cleanup()
        DetectHiddenWindows(previousHidden)
    }
}
