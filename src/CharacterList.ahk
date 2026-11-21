; Independent character launcher. Configuration is global, separate from thumbnail visibility.
class CharacterList {
    __New(owner) {
        this.owner := owner, this.gui := 0, this.signature := "", this.offset := 0
        this.tick := ObjBindMethod(this, "Refresh")
        this.mouse := ObjBindMethod(this, "MouseDown")
        this.wheelCallback := ObjBindMethod(this, "Wheel")
        this.endMove := ObjBindMethod(this, "SaveBounds")
        OnMessage(0x201, this.mouse)
        OnMessage(0x20A, this.wheelCallback)
        OnMessage(0x232, this.endMove)
        SetTimer(this.tick, 400)
        this.Refresh()
    }
    Config => this.owner._JSON
    Get(key, fallback) => this.Config.Get("CharacterList" key, fallback)
    Put(key, value) {
        this.Config["CharacterList" key] := value
        SetTimer(this.owner.Save_Settings_Delay_Timer, -200)
    }
    Entries() {
        entries := []
        for hwnd in WinGetList(this.owner.EVEExe) {
            try {
                title := WinGetTitle(hwnd)
                name := RegExReplace(title, "^EVE - ", "")
                created := 0
                try created := this.owner.GetProcessCreationTime(WinGetPID(hwnd))[3]
                entries.Push({hwnd: hwnd, name: name, created: created})
            }
        }
        order := this.Get("Order", [])
        mode := this.Get("Sort", "name")
        ; Stable insertion sort gives deterministic ordering for duplicate login labels.
        Loop entries.Length {
            i := A_Index
            while i > 1 && this.Compare(entries[i], entries[i-1], mode, order) < 0 {
                temp := entries[i-1], entries[i-1] := entries[i], entries[i] := temp
                i -= 1
            }
        }
        return entries
    }
    Rank(name, order) {
        for i, item in order
            if item = name
                return i
        return 2147483647
    }
    Compare(a, b, mode, order) {
        if mode = "manual" {
            delta := this.Rank(a.name, order) - this.Rank(b.name, order)
            if delta
                return delta
        }
        if mode = "oldest" || mode = "newest"
            return mode = "oldest" ? a.created - b.created : b.created - a.created
        delta := StrCompare(a.name, b.name, false)
        return mode = "descending" ? -delta : delta
    }
    Refresh(*) {
        if GetKeyState("LButton", "P") || (this.HasOwnProp("dragging") && this.dragging)
            return
        if !this.Get("Enabled", false) {
            if this.gui
                this.gui.Destroy(), this.gui := 0, this.signature := ""
            return
        }
        entries := this.Entries(), signature := JSON.Dump(entries.Length ? this.EntryKeys(entries) : [])
        signature .= this.Config.Get("SettingsTheme", "system") this.Get("Collapsed", false) this.Get("Columns", 1) this.Get("ColumnWidth", 220) this.Get("Direction", "vertical") this.Get("Sort", "name") this.ActiveClient()
        if signature = this.signature
            return
        this.signature := signature, this.items := entries
        this.Render()
    }
    ActiveClient() {
        active := WinExist("A")
        for hwnd in WinGetList(this.owner.EVEExe)
            if hwnd = active
                return hwnd
        return 0
    }
    EntryKeys(entries) {
        keys := []
        for entry in entries
            keys.Push(entry.name "|" entry.hwnd)
        return keys
    }
    Colors() {
        mode := this.Config.Get("SettingsTheme", "system"), dark := mode != "light"
        if mode = "system"
            try dark := !RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "AppsUseLightTheme", 1)
        if SettingsPalettes.Items.Has(mode)
            return SettingsPalettes.Items[mode]
        return Map("background", dark ? "202124" : "F7F7F7", "foreground", dark ? "ECEEF2" : "202020", "accent", dark ? "76B7FF" : "0078D4")
    }
    Render() {
        if this.gui
            this.gui.Destroy()
        {
            this.gui := Gui("+AlwaysOnTop +Resize +ToolWindow -DPIScale", "EVE character list")
            this.gui.OnEvent("Close", (*) => this.Disable())
            this.gui.OnEvent("Size", ObjBindMethod(this, "Resize"))
        }
        g := this.gui, colors := this.Colors()

        g.BackColor := colors["background"]
        g.SetFont("s10 c" colors["foreground"], "Segoe UI")
        arrow := g.AddText("x8 y8 w24 h24 +0x100", this.Get("Collapsed", false) ? Chr(0x25B6) : Chr(0x25BC))
        arrow.OnEvent("Click", (*) => this.Collapse())
        g.AddText("x38 y8 w160 h24", "Characters · " this.items.Length)
        this.cells := Map()
        cols := Max(1, this.Get("Columns", 1)), cw := Max(100, this.Get("ColumnWidth", 220))
        h := this.Get("Height", 340)
        rows := Max(1, Floor((h - 48) / 32)), capacity := rows * cols
        this.offset := Min(this.offset, Max(0, this.items.Length - capacity))
        if !this.Get("Collapsed", false) {
            Loop Min(capacity, this.items.Length - this.offset) {
                i := A_Index, entry := this.items[i + this.offset]
                horizontal := this.Get("Direction", "vertical") = "horizontal"
                col := horizontal ? Mod(i-1, cols) : Floor((i-1) / rows)
                row := horizontal ? Floor((i-1) / cols) : Mod(i-1, rows)
                active := WinActive("ahk_id " entry.hwnd)
                cell := g.AddText("x" (8 + col*cw) " y" (40 + row*32) " w" (cw-12) " h28 +0x100 c" (active ? colors["accent"] : colors["foreground"]), (active ? "● " : "○ ") entry.name)
                this.cells[cell.Hwnd] := i + this.offset
            }
            if !this.items.Length
                g.AddText("x12 y44 w200", "No EVE clients open")
        }
        g.Show("NA x" this.Get("X", 40) " y" this.Get("Y", 80) " w" (cols*cw+16) " h" (this.Get("Collapsed", false) ? 38 : h))
    }
    Resize(guiObj, minMax, width, height) {
        if minMax = -1 || this.Get("Collapsed", false)
            return
        cw := Max(100, Floor((width-16)/Max(1, this.Get("Columns", 1))))
        if cw = this.Get("ColumnWidth", 220) && height = this.Get("Height", 340)
            return
        this.Put("ColumnWidth", cw), this.Put("Height", Max(100, height))
        this.signature := ""
    }
    SaveBounds(wParam, lParam, msg, hwnd) {
        if this.gui && hwnd = this.gui.Hwnd {
            this.gui.GetPos(&x, &y)
            this.Put("X", x), this.Put("Y", y)
        }
    }
    Collapse() {
        this.Put("Collapsed", !this.Get("Collapsed", false)), this.signature := ""
        this.Refresh()
    }
    Disable() => this.Put("Enabled", false)
    Wheel(wParam, lParam, msg, hwnd) {
        if !this.gui || (hwnd != this.gui.Hwnd && !this.cells.Has(hwnd))
            return
        this.offset := Max(0, this.offset + ((wParam >> 16 & 0x8000) ? 1 : -1))
        this.signature := ""
        this.Refresh()
        return 0
    }
    MouseDown(wParam, lParam, msg, hwnd) {
        if !this.gui || !this.cells.Has(hwnd)
            return
        index := this.cells[hwnd], entry := this.items[index]
        MouseGetPos(&startX, &startY)
        this.dragging := true
        KeyWait("LButton")
        this.dragging := false
        MouseGetPos(&endX, &endY,, &target, 2)
        if Abs(endX-startX) + Abs(endY-startY) < 6 {
            try this.owner.ActivateEVEWindow(entry.hwnd)
        } else if this.cells.Has(target) {
            destination := this.cells[target]
            this.items.RemoveAt(index), this.items.InsertAt(destination, entry)
            order := []
            for item in this.items
                order.Push(item.name)
            this.Put("Order", order), this.Put("Sort", "manual")
        }
        this.signature := ""
        return 0
    }
}
