; Row and event editors keep their complete form reachable at high display
; scaling. Their canvas scrolls vertically; narrow screens scale column widths.
class SettingsChildDialog {
    __New(app, title) {
        this.app := app
        this.root := app.theme.Surface(Gui("+Owner" app.S_Gui.Hwnd " +Resize", title))
        this.scroll := SettingsScrollPane(this.root)
        app.theme.Scroll(this.scroll)
        this.content := app.theme.Surface(this.scroll.content)
        this.content.SetFont("s10", "Segoe UI")
        this.root.OnEvent("Close", (*) => app.CloseChild(this))
        this.root.OnEvent("Escape", (*) => app.CloseChild(this))
        this.content.OnEvent("Escape", (*) => app.CloseChild(this))
        this.root.OnEvent("Size", ObjBindMethod(this, "Resize"))
        this.shown := false
        this.closed := false
    }

    Hwnd => this.root.Hwnd
    __Call(name, args) {
        if name = "AddText" && args.Length
            args[1] .= " +0x80" ; Literal ampersands in descriptions and paths.
        return this.content.%name%(args*)
    }
    SetFont(options := "", name := "") => this.content.SetFont(options, name)
    Opt(options) => this.root.Opt(options)
    OnEvent(name, callback, addRemove := 1) => this.root.OnEvent(name, callback, addRemove)

    Help(control, text) {
        control.SettingsHelp := text
        control.ToolTip := text
        return control
    }

    SetError(control, message) {
        control.Value := message
        control.ToolTip := message
        control.GetPos(,, &width)
        height := Max(22, SettingsLayout.TextHeight(control, message, width))
        for index, item in this.controls
            if item = control {
                origin := this.rectangles[index]
                delta := height - origin.h
                origin.h := height
                for rectangle in this.rectangles
                    if rectangle.y > origin.y
                        rectangle.y += delta
                break
            }
        this.Layout(this.currentWidth, this.currentHeight)
    }

    Show(options := "") {
        this.canvasWidth := RegExMatch(options, "i)\bw(\d+)", &m) ? m[1] + 0 : 700
        canvasHeight := RegExMatch(options, "i)\bh(\d+)", &m) ? m[1] + 0 : 600
        this.controls := [], this.rectangles := []
        for control in this.content {
            control.GetPos(&x, &y, &w, &h)
            this.controls.Push(control)
            this.rectangles.Push({x: x, y: y, w: w, h: h})
        }
        work := SettingsWindowSizing.WorkArea(this.app.S_Gui.Hwnd)
        left := work.left, top := work.top, right := work.right, bottom := work.bottom
        fit := SettingsWindowSizing.Fit(left, top, right, bottom, A_ScreenDPI)
        width := Min(this.canvasWidth, fit.width), height := Min(canvasHeight, fit.height)
        this.root.Opt("+MinSize" Min(440, width) "x" Min(260, height))
        SettingsWindowSizing.ShowCentered(this.root, this.app.S_Gui.Hwnd, width, height)
        this.shown := true
        this.Layout(width, height)
    }

    Resize(gui, state, width, height) {
        if this.closed || !this.shown || state = -1
            return
        this.app.FlushTableEdits()
        this.Layout(width, height)
    }

    Layout(width, height) {
        this.app.theme.BeginUpdate(this.root)
        try {
            this.currentWidth := width, this.currentHeight := height
            position := this.scroll.offset
            ratio := Min(1, width / this.canvasWidth)
            for index, control in this.controls {
                rect := this.rectangles[index]
                minimum := control.Type = "Text" ? 1 : 20
                control.Move(Round(rect.x * ratio), rect.y, Max(minimum, Round(rect.w * ratio)), rect.h)
            }
            this.scroll.Show(0, 0, width, height)
            this.scroll.Select(this.controls)
            this.scroll.ScrollTo(position)
        } finally this.app.theme.EndUpdate(this.root)
    }

    Destroy() {
        if this.closed
            return
        this.closed := true
        this.scroll.Stop()
        this.root.Destroy()
    }
}
