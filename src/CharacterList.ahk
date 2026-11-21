class CharacterList {
    static Defaults() => Map("Enabled", 0, "Columns", 1, "ColumnWidth", 220, "Direction", "vertical", "Sort", "name", "Order", [], "AlwaysOnTop", 1, "LockPosition", 0, "AutoSavePosition", 1, "Opacity", 100, "ShowHwnd", 0, "GroupLoggedOut", 1, "GroupSystems", 0, "HideOnLostFocus", 0, "X", 40, "Y", 80, "Height", 340, "Collapsed", 0)
    static Migrate(config) {
        for name, profile in config["_Profiles"] {
            if profile.Has("SettingsInheritance") && !profile["SettingsInheritance"].Has("character-list")
                profile["SettingsInheritance"]["character-list"] := 0
            if !profile.Has("Character List")
                profile["Character List"] := Map()
            for key, value in this.Defaults() {
                if config.Has("CharacterList" key)
                    profile["Character List"]["CharacterList" key] := SettingsModel.Clone(config["CharacterList" key])
                else if !profile["Character List"].Has("CharacterList" key)
                    profile["Character List"]["CharacterList" key] := SettingsModel.Clone(value)
            }
        }
        for key in this.Defaults()
            if config.Has("CharacterList" key)
                config.Delete("CharacterList" key)
    }
    __New(owner) {
        this.owner := owner, this.gui := 0, this.signature := "", this.offset := 0, this.pool := [], this.cells := Map(), this.dragging := false
        CharacterList.Migrate(owner._JSON)
        this.tick := ObjBindMethod(this, "Refresh"), this.mouse := ObjBindMethod(this, "MouseDown"), this.wheelCallback := ObjBindMethod(this, "Wheel"), this.endMove := ObjBindMethod(this, "SaveBounds")
        OnMessage(0x201, this.mouse), OnMessage(0x20A, this.wheelCallback), OnMessage(0x232, this.endMove)
        SetTimer(this.tick, 250)
        this.Refresh()
    }
    Config => this.owner._JSON
    ProfileName {
        get {
            name := this.Config["LastUsedProfile"]
            flags := this.Config["_Profiles"][name].Get("SettingsInheritance", Map())
            return name != "Default" && flags.Get("character-list", 0) ? "Default" : name
        }
    }
    Data => this.Config["_Profiles"][this.ProfileName]["Character List"]
    Get(key, fallback) => this.Data.Get("CharacterList" key, fallback)
    Put(key, value) {
        name := this.Config["LastUsedProfile"]
        if name != this.ProfileName {
            profile := this.Config["_Profiles"][name]
            profile["Character List"] := SettingsModel.Clone(this.Data)
            profile["SettingsInheritance"]["character-list"] := 0
        }
        this.Data["CharacterList" key] := value
        this.signature := ""
        SetTimer(this.owner.Save_Settings_Delay_Timer, -200)
    }
    Entries() {
        entries := []
        for hwnd in WinGetList(this.owner.EVEExe) {
            try {
                title := WinGetTitle(hwnd), name := RegExReplace(title, "^EVE - ", ""), created := 0, system := ""
                loggedOut := title = "EVE" || title = "Char Screen"
                try created := this.owner.GetProcessCreationTime(WinGetPID(hwnd))[3]
                try system := this.owner.systemLogMonitor.readers[title].system
                entries.Push({hwnd: hwnd, name: name, created: created, loggedOut: loggedOut, system: system, key: loggedOut ? "hwnd:" hwnd : name})
            }
        }
        mode := this.Get("Sort", "name"), order := this.Get("Order", [])
        Loop entries.Length {
            i := A_Index
            while i > 1 && this.Compare(entries[i], entries[i-1], mode, order) < 0 {
                temp := entries[i-1], entries[i-1] := entries[i], entries[i] := temp, i -= 1
            }
        }
        return entries
    }
    Rank(key, order) {
        for i, item in order
            if item = key
                return i
        return 2147483647
    }
    Compare(a, b, mode, order) {
        if mode = "manual" {
            delta := this.Rank(a.key, order) - this.Rank(b.key, order)
            if delta
                return delta
        } else {
            if this.Get("GroupLoggedOut", 1) && a.loggedOut != b.loggedOut
                return (a.loggedOut ? 1 : -1) * (mode = "descending" ? -1 : 1)
            if this.Get("GroupSystems", 0) && !a.loggedOut && !b.loggedOut {
                delta := StrCompare(a.system, b.system, false)
                if delta
                    return mode = "descending" ? -delta : delta
            }
        }
        if mode = "oldest" || mode = "newest"
            return mode = "oldest" ? a.created - b.created : b.created - a.created
        delta := StrCompare(a.name, b.name, false)
        return delta ? (mode = "descending" ? -delta : delta) : a.created - b.created
    }
    ActiveClient() {
        active := WinExist("A")
        for hwnd in WinGetList(this.owner.EVEExe)
            if hwnd = active
                return hwnd
        return 0
    }
    Refresh(*) {
        if (this.HasOwnProp("menuOpen") && this.menuOpen) || this.dragging || GetKeyState("LButton", "P")
            return
        if !this.Get("Enabled", 0) {
            if this.gui
                this.gui.Hide()
            return
        }
        if this.Get("GroupSystems", 0) && !this.owner.HasOwnProp("systemLogMonitor")
            try this.owner.startSystemTracking()
        active := this.ActiveClient()
        if active
            this.lastActive := active
        if this.Get("HideOnLostFocus", 0) && !active && (!this.gui || !WinActive("ahk_id " this.gui.Hwnd)) {
            if this.gui
                this.gui.Hide()
            return
        }
        entries := this.Entries(), keys := []
        for entry in entries
            keys.Push([entry.key, entry.name, entry.system])
        signature := JSON.Dump([keys, this.Data, this.Colors(), this.ProfileName, active])
        if signature != this.signature {
            this.signature := signature, this.items := entries
            this.Render()
        } else if this.gui && !DllCall("IsWindowVisible", "Ptr", this.gui.Hwnd)
            this.gui.Show("NA")
    }
    Colors() {
        mode := this.Config.Get("SettingsTheme", "system"), dark := mode != "light"
        if mode = "system"
            try dark := !RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "AppsUseLightTheme", 1)
        if SettingsPalettes.Items.Has(mode)
            return SettingsPalettes.Items[mode]
        return Map("background", dark ? "202124" : "F7F7F7", "foreground", dark ? "ECEEF2" : "202020", "accent", dark ? "76B7FF" : "0078D4", "selection", dark ? "304C70" : "CCE4FF")
    }
    Render() {
        this.rendering := true
        try {
            if !this.gui {
                this.gui := Gui("-Caption +Resize +ToolWindow -DPIScale", "EVE character list")
                this.gui.OnEvent("Size", ObjBindMethod(this, "Resize"))
                this.gui.OnEvent("Close", (*) => this.Put("Enabled", 0))
                this.arrow := this.gui.AddText("x4 y5 w24 h24 +0x100", "")
                this.arrow.OnEvent("Click", (*) => this.Collapse())
                this.heading := this.gui.AddText("x30 y5 w130 h24 +0x100", "Characters")
                this.dots := this.gui.AddText("x170 y5 w24 h24 Center +0x100", "⋯")
                this.dots.OnEvent("Click", ObjBindMethod(this, "ShowMenu"))
            }
            g := this.gui, colors := this.Colors(), cols := Max(1, this.Get("Columns", 1)), cw := Max(50, this.Get("ColumnWidth", 220))
            width := cols*cw+16, height := Max(100, this.Get("Height", 340)), collapsed := this.Get("Collapsed", 0)
            g.BackColor := colors["background"]
            g.SetFont("s10 c" colors["foreground"], "Segoe UI")
            this.arrow.SetFont("c" colors["foreground"]), this.heading.SetFont("c" colors["foreground"]), this.dots.SetFont("s14 c" colors["foreground"])
            this.arrow.Text := collapsed ? "▶" : "▼"
            this.heading.Text := width >= 140 ? "Characters · " this.items.Length : ""
            this.heading.Move(30, 5, Max(1, width-64), 24), this.dots.Move(width-28, 2, 24, 28)
            g.Opt(this.Get("LockPosition", 0) ? "-Resize" : "+Resize")
            WinSetAlwaysOnTop(this.Get("AlwaysOnTop", 1), g.Hwnd)
            rows := Max(1, Floor((height-42)/32)), capacity := rows*cols
            this.offset := Min(this.offset, Max(0, this.items.Length-capacity))
            count := collapsed ? 0 : Min(capacity, this.items.Length-this.offset)
            this.cells := Map()
            Loop count {
                i := A_Index, entry := this.items[i+this.offset]
                if i > this.pool.Length
                    this.pool.Push(g.AddText("x0 y0 w1 h28 +0x100 +0x200", ""))
                cell := this.pool[i], horizontal := this.Get("Direction", "vertical") = "horizontal"
                col := horizontal ? Mod(i-1, cols) : Floor((i-1)/rows), row := horizontal ? Floor((i-1)/cols) : Mod(i-1, rows)
                cell.Move(8+col*cw, 36+row*32, cw-8, 28)
                active := this.HasOwnProp("lastActive") && entry.hwnd = this.lastActive
                cell.Opt("Background" (active ? colors["selection"] : colors["background"]))
                cell.SetFont((active ? "bold " : "norm ") "c" colors["foreground"])
                label := entry.name
                if entry.loggedOut && this.Get("ShowHwnd", 0)
                    label .= " [" Format("0x{:X}", entry.hwnd) "]"
                if this.Get("GroupSystems", 0) && !entry.loggedOut
                    label .= " · " (entry.system = "" ? "Unknown system" : entry.system)
                cell.Text := (active ? "▌ " : "  ") label
                cell.Visible := true, this.cells[cell.Hwnd] := i+this.offset
            }
            for i, cell in this.pool
                if i > count
                    cell.Visible := false
            firstShow := !this.HasOwnProp("shownProfile") || this.shownProfile != this.Config["LastUsedProfile"]
            if firstShow {
                position := " x" this.Get("X", 40) " y" this.Get("Y", 80)
                this.shownProfile := this.Config["LastUsedProfile"]
            } else
                position := ""
            g.Show("NA" position " w" width " h" (collapsed ? 32 : height))
            WinSetTransparent(Round(Max(10, this.Get("Opacity", 100))*2.55), g.Hwnd)
        } finally this.rendering := false
    }
    Resize(guiObj, minMax, width, height) {
        if minMax = -1 || this.Get("Collapsed", 0) || (this.HasOwnProp("rendering") && this.rendering)
            return
        this.Put("ColumnWidth", Max(50, Floor((width-16)/Max(1, this.Get("Columns", 1)))))
        this.Data["CharacterListHeight"] := Max(100, height)
    }
    SaveBounds(wParam, lParam, msg, hwnd) {
        if this.gui && hwnd = this.gui.Hwnd && this.Get("AutoSavePosition", 1) {
            this.gui.GetPos(&x, &y)
            this.Put("X", x), this.Put("Y", y)
        }
    }
    Collapse() {
        this.Put("Collapsed", !this.Get("Collapsed", 0))
        this.Refresh()
    }
    Wheel(wParam, lParam, msg, hwnd) {
        if !this.gui || (hwnd != this.gui.Hwnd && !this.cells.Has(hwnd))
            return
        this.offset := Max(0, this.offset+((wParam >> 16 & 0x8000) ? 1 : -1)), this.signature := ""
        this.Refresh()
        return 0
    }
    ShowMenu(*) {
        menu := Menu()
        for key, label in Map("AlwaysOnTop", "Always on top", "LockPosition", "Lock position", "AutoSavePosition", "Auto save position") {
            menu.Add(label, ObjBindMethod(this, "Toggle", key))
            if this.Get(key, key != "LockPosition")
                menu.Check(label)
        }
        sortMenu := Menu()
        for key, label in Map("name", "Name A–Z", "descending", "Name Z–A", "oldest", "Oldest client first", "newest", "Newest client first", "manual", "Custom order") {
            sortMenu.Add(label, ObjBindMethod(this, "Choose", "Sort", key))
            if this.Get("Sort", "name") = key
                sortMenu.Check(label)
        }
        menu.Add("Sort order", sortMenu)
        opacityMenu := Menu()
        for percent in [10, 25, 50, 75, 90, 100] {
            opacityMenu.Add(percent "%", ObjBindMethod(this, "Choose", "Opacity", percent))
            if this.Get("Opacity", 100) = percent
                opacityMenu.Check(percent "%")
        }
        opacityMenu.Add("Custom…", ObjBindMethod(this, "CustomOpacity"))
        menu.Add("Opacity", opacityMenu)
        fillMenu := Menu()
        for direction in ["vertical", "horizontal"] {
            fillMenu.Add(direction, ObjBindMethod(this, "Choose", "Direction", direction))
            if this.Get("Direction", "vertical") = direction
                fillMenu.Check(direction)
        }
        menu.Add("Fill direction", fillMenu)
        this.menuOpen := true
        try menu.Show()
        finally this.menuOpen := false
    }
    Toggle(key, *) => this.Put(key, !this.Get(key, 0))
    Choose(key, value, *) => this.Put(key, value)
    CustomOpacity(*) {
        result := InputBox("Window opacity (10–100 percent)", "Character list opacity", "w260 h120", this.Get("Opacity", 100))
        if result.Result = "OK" && IsNumber(result.Value) && result.Value >= 10 && result.Value <= 100
            this.Put("Opacity", Round(result.Value))
    }
    MouseDown(wParam, lParam, msg, hwnd) {
        if !this.gui
            return
        if hwnd = this.heading.Hwnd || hwnd = this.gui.Hwnd {
            if !this.Get("LockPosition", 0)
                PostMessage(0xA1, 2,,, this.gui.Hwnd)
            return
        }
        if !this.cells.Has(hwnd)
            return
        index := this.cells[hwnd], entry := this.items[index]
        MouseGetPos(&startX, &startY)
        this.dragging := true, moved := false, targetIndex := 0
        try {
            while GetKeyState("LButton", "P") {
                MouseGetPos(&x, &y,, &target, 2)
                if Abs(x-startX)+Abs(y-startY) >= 6 {
                    moved := true
                    nextIndex := this.cells.Get(target, 0)
                    if nextIndex != targetIndex {
                        for handle, itemIndex in this.cells
                            GuiCtrlFromHwnd(handle).Text := this.items[itemIndex].name
                        targetIndex := nextIndex
                        for handle, itemIndex in this.cells
                            if itemIndex = targetIndex {
                                cell := GuiCtrlFromHwnd(handle)
                                cell.Text := "➜ Drop here: " this.items[itemIndex].name
                            }
                    }
                    ToolTip("Moving " entry.name (targetIndex ? " → " this.items[targetIndex].name : " · choose a row"))
                }
                Sleep(16)
            }
            if !moved {
                try this.owner.ActivateEVEWindow(entry.hwnd)
            } else if targetIndex && targetIndex != index {
                this.items.RemoveAt(index), this.items.InsertAt(targetIndex, entry), order := []
                for item in this.items
                    order.Push(item.key)
                this.Put("Order", order), this.Put("Sort", "manual")
            }
        } finally {
            ToolTip(), this.dragging := false, this.signature := ""
        }
        return 0
    }
}
