class CharacterList {
    static Defaults() => Map("Enabled", 0, "Columns", 1, "ColumnWidth", 220, "Direction", "vertical", "Sort", "name", "Order", [], "AlwaysOnTop", 1, "LockPosition", 0, "AutoSavePosition", 1, "Opacity", 100, "ShowHwnd", 0, "AutoFill", 1, "Lanes", Map(), "HideOnLostFocus", 0, "X", 40, "Y", 80, "Height", 340, "Width", 0, "Collapsed", 0)
    static Migrate(config) {
        for name, profile in config["_Profiles"] {
            if profile.Has("SettingsInheritance") && !profile["SettingsInheritance"].Has("character-list")
                profile["SettingsInheritance"]["character-list"] := 0
            if !profile.Has("Character List")
                profile["Character List"] := Map()
            for obsolete in ["GroupLoggedOut", "GroupSystems"]
                if profile["Character List"].Has("CharacterList" obsolete)
                    profile["Character List"].Delete("CharacterList" obsolete)
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
        this.owner := owner, this.gui := 0, this.signature := "", this.offset := 0, this.pool := [], this.eventBorders := [], this.cells := Map(), this.dragging := false
        CharacterList.Migrate(owner._JSON)
        this.events := Map()
        this.ncCallback := ObjBindMethod(this, "NonClient")
        OnMessage(0x83, this.ncCallback), OnMessage(0x84, this.ncCallback)
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
        if key = "AutoFill" && value
            this.Data["CharacterListLanes"] := Map()
        this.Data["CharacterList" key] := value
        this.signature := ""
        SetTimer(this.owner.Save_Settings_Delay_Timer, -200)
    }
    Entries() {
        entries := []
        for hwnd in WinGetList(this.owner.EVEExe) {
            try {
                title := WinGetTitle(hwnd), name := RegExReplace(title, "^EVE - ", ""), created := 0
                loggedOut := title = "EVE" || title = "Char Screen"
                try created := this.owner.GetProcessCreationTime(WinGetPID(hwnd))[3]
                entries.Push({hwnd: hwnd, name: name, created: created, loggedOut: loggedOut, key: loggedOut ? "hwnd:" hwnd : name})
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
        active := this.ActiveClient()
        if active {
            if this.HasOwnProp("lastActive") && active != this.lastActive
                try {
                    if this.owner.stopDisplayingOnSwitch
                        this.EventEnd(active)
                }
            this.lastActive := active
        }
        if this.Get("HideOnLostFocus", 0) && !active && (!this.gui || !WinActive("ahk_id " this.gui.Hwnd)) {
            if this.gui
                this.gui.Hide()
            return
        }
        entries := this.Entries(), keys := []
        for entry in entries
            keys.Push([entry.key, entry.name, this.EventState(entry.hwnd)])
        signature := JSON.Dump([keys, this.Data, this.Colors(), this.ProfileName, active])
        if signature != this.signature {
            this.signature := signature, this.items := entries
            this.Render()
        } else if this.gui && !DllCall("IsWindowVisible", "Ptr", this.gui.Hwnd)
            DllCall("ShowWindow", "Ptr", this.gui.Hwnd, "Int", 4)
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
                this.arrow := this.gui.AddText("x4 y5 w24 h24 +0x100 BackgroundTrans", "")
                this.arrow.OnEvent("Click", (*) => this.Collapse())
                this.heading := this.gui.AddText("x30 y5 w130 h24 +0x100 BackgroundTrans", "Characters")
                this.dots := this.gui.AddText("x170 y5 w24 h24 Center +0x100 BackgroundTrans", "⋯")
                this.dots.OnEvent("Click", ObjBindMethod(this, "ShowMenu"))
                this.track := this.gui.AddText("x0 y0 w1 h1 +0x100", "")
                this.thumb := this.gui.AddText("x0 y0 w1 h1 +0x100", "")
                this.insertion := this.gui.AddText("x0 y0 w1 h1 BackgroundFF9900", "")
                this.insertion.Visible := false
            }
            g := this.gui, colors := this.Colors(), cols := Max(1, this.Get("Columns", 1)), cw := Max(50, this.Get("ColumnWidth", 220))
            width := cols*cw+16, height := Max(100, this.Get("Height", 340)), collapsed := this.Get("Collapsed", 0)
            if this.Get("Direction", "vertical") = "horizontal" {
                height := Max(height, 52+cols*32)
                width := Max(66, this.Get("Width", 0) ? this.Get("Width", 0) : cols*cw+16)
            }
            g.BackColor := colors["background"]
            g.SetFont("s10 c" colors["foreground"], "Segoe UI")
            this.arrow.SetFont("c" colors["foreground"]), this.heading.SetFont("c" colors["foreground"]), this.dots.SetFont("s14 c" colors["foreground"])
            this.arrow.Text := collapsed ? "▶" : "▼"
            this.heading.Text := width >= 140 ? "Characters · " this.items.Length : ""
            this.heading.Move(30, 5, Max(1, width-64), 24), this.dots.Move(width-28, 2, 24, 28)
            g.Opt(this.Get("LockPosition", 0) ? "-Resize" : "+Resize")
            WinSetAlwaysOnTop(this.Get("AlwaysOnTop", 1), g.Hwnd)
            horizontal := this.Get("Direction", "vertical") = "horizontal"
            visible := horizontal ? Max(1, Floor((width-16)/cw)) : Max(1, Floor((height-56)/32))
            this.layout := this.BuildLayout(this.items, visible)
            extent := 0
            for slot in this.layout
                extent := Max(extent, slot.along+1)
            this.maxOffset := Max(0, extent-visible), this.offset := Min(this.offset, this.maxOffset)
            this.visibleSpan := visible, this.contentWidth := width, this.contentHeight := height
            this.cells := Map(), this.placements := Map(), count := 0
            if !collapsed
                for slot in this.layout {
                    if slot.along < this.offset || slot.along >= this.offset+visible
                        continue
                    count += 1, entry := this.items[slot.index]
                    if count > this.pool.Length {
                        this.eventBorders.Push(g.AddText("x0 y0 w1 h1", ""))
                        this.pool.Push(g.AddText("x0 y0 w1 h28 +0x100 +0x200", ""))
                    }
                    cell := this.pool[count]
                    x := 8+(horizontal ? slot.along-this.offset : slot.lane)*cw
                    y := 36+(horizontal ? slot.lane : slot.along-this.offset)*32
                    if horizontal && y+28 > height-16
                        continue
                    cell.Move(x, y, cw-8, 28)
                    active := this.HasOwnProp("lastActive") && entry.hwnd = this.lastActive
                    event := this.EventState(entry.hwnd)
                    cell.Opt("Background" (active ? colors["selection"] : colors["background"]))
                    border := this.eventBorders[count]
                    border.Move(x-1, y-1, cw-6, 30)
                    border.Visible := event[2] != ""
                    if event[2] != ""
                        border.Opt("Background" event[2])
                    cell.SetFont((active ? "bold " : "norm ") "c" colors["foreground"])
                    label := entry.name
                    if entry.loggedOut && this.Get("ShowHwnd", 0)
                        label .= " [" Format("0x{:X}", entry.hwnd) "]"
                    if event[1] != ""
                        label .= " · " event[1]
                    cell.Text := label
                    cell.Visible := true, this.cells[cell.Hwnd] := slot.index
                    this.placements[cell.Hwnd] := {lane: slot.lane, along: slot.along, x: x, y: y, index: slot.index}
                }
            for i, cell in this.pool
                if i > count {
                    cell.Visible := false
                    this.eventBorders[i].Visible := false
                }
            this.track.Visible := !collapsed && this.maxOffset > 0
            this.thumb.Visible := !collapsed && this.maxOffset > 0
            this.insertion.Opt("Background" colors["accent"])
            this.track.Opt("Background" colors["selection"]), this.thumb.Opt("Background" colors["accent"])
            if horizontal {
                trackLength := width-16, thumbLength := Max(18, Round(trackLength*visible/Max(visible, extent)))
                this.track.Move(8, height-12, trackLength, 8)
                this.thumb.Move(8+Round((trackLength-thumbLength)*this.offset/Max(1, this.maxOffset)), height-12, thumbLength, 8)
            } else {
                trackLength := height-48, thumbLength := Max(18, Round(trackLength*visible/Max(visible, extent)))
                this.track.Move(width-9, 36, 7, trackLength)
                this.thumb.Move(width-9, 36+Round((trackLength-thumbLength)*this.offset/Max(1, this.maxOffset)), 7, thumbLength)
            }
            firstShow := !this.HasOwnProp("shownProfile") || this.shownProfile != this.Config["LastUsedProfile"]
            if firstShow {
                wx := this.Get("X", 40), wy := this.Get("Y", 80)
                this.shownProfile := this.Config["LastUsedProfile"]
            } else
                g.GetPos(&wx, &wy)
            DllCall("SetWindowPos", "Ptr", g.Hwnd, "Ptr", 0, "Int", wx, "Int", wy, "Int", width, "Int", collapsed ? 32 : height, "UInt", 0x54)
            WinSetTransparent(Round(Max(10, this.Get("Opacity", 100))*2.55), g.Hwnd)
        } finally this.rendering := false
    }
    Resize(guiObj, minMax, width, height) {
        if minMax = -1 || this.Get("Collapsed", 0) || (this.HasOwnProp("rendering") && this.rendering)
            return
        if this.Get("Direction", "vertical") = "horizontal" {
            if width != this.Get("Width", 0) || height != this.Get("Height", 340) {
                this.Put("Width", Max(66, width))
                this.Data["CharacterListHeight"] := Max(100, height)
            }
            return
        }
        cw := Max(50, Floor((width-16)/Max(1, this.Get("Columns", 1))))
        if cw != this.Get("ColumnWidth", 220) || height != this.Get("Height", 340) {
            this.Put("ColumnWidth", cw)
            this.Data["CharacterListHeight"] := Max(100, height)
        }
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
        this.offset := Min(this.maxOffset, Max(0, this.offset+((wParam >> 16 & 0x8000) ? 1 : -1))), this.signature := ""
        if this.dragging
            this.Render()
        else
            this.Refresh()
        return 0
    }
    BuildMenu() {
        popup := Menu()
        for key, label in Map("AlwaysOnTop", "Always on top", "LockPosition", "Lock position", "AutoSavePosition", "Auto save position") {
            popup.Add(label, ObjBindMethod(this, "Toggle", key))
            if this.Get(key, key != "LockPosition")
                popup.Check(label)
        }
        sortMenu := Menu()
        for key, label in Map("name", "Name A–Z", "descending", "Name Z–A", "oldest", "Oldest client first", "newest", "Newest client first", "manual", "Custom order") {
            sortMenu.Add(label, ObjBindMethod(this, "Choose", "Sort", key))
            if this.Get("Sort", "name") = key
                sortMenu.Check(label)
        }
        popup.Add("Sort order", sortMenu)
        opacityMenu := Menu()
        for percent in [10, 25, 50, 75, 90, 100] {
            opacityMenu.Add(percent "%", ObjBindMethod(this, "Choose", "Opacity", percent))
            if this.Get("Opacity", 100) = percent
                opacityMenu.Check(percent "%")
        }
        opacityMenu.Add("Custom…", ObjBindMethod(this, "CustomOpacity"))
        popup.Add("Opacity", opacityMenu)
        fillMenu := Menu()
        for direction in ["vertical", "horizontal"] {
            fillMenu.Add(direction, ObjBindMethod(this, "Choose", "Direction", direction))
            if this.Get("Direction", "vertical") = direction
                fillMenu.Check(direction)
        }
        popup.Add("Fill direction", fillMenu)
        popup.Add("Fill additional columns / rows", ObjBindMethod(this, "Toggle", "AutoFill"))
        if this.Get("AutoFill", 1)
            popup.Check("Fill additional columns / rows")
        return popup
    }
    ShowMenu(*) {
        popup := this.BuildMenu()
        this.menuOpen := true
        try popup.Show()
        finally this.menuOpen := false
    }
    Toggle(key, *) => this.Put(key, !this.Get(key, 0))
    Choose(key, value, *) => this.Put(key, value)
    CustomOpacity(*) {
        result := InputBox("Window opacity (10–100 percent)", "Character list opacity", "w260 h120", this.Get("Opacity", 100))
        if result.Result = "OK" && IsNumber(result.Value) && result.Value >= 10 && result.Value <= 100
            this.Put("Opacity", Round(result.Value))
    }
    BuildLayout(entries, visible) {
        result := [], used := Map(), lanes := Max(1, this.Get("Columns", 1)), saved := this.Get("Lanes", Map())
        span := Max(visible, Ceil(entries.Length/lanes))
        for i, entry in entries {
            lane := this.Get("AutoFill", 1) ? Min(lanes-1, Floor((i-1)/span)) : 0
            lane := Min(lanes-1, Max(0, saved.Get(entry.key, lane)))
            along := used.Get(lane, 0), used[lane] := along+1
            result.Push({index: i, lane: lane, along: along})
        }
        return result
    }
    NonClient(wParam, lParam, msg, hwnd) {
        if !this.gui || hwnd != this.gui.Hwnd
            return
        if msg = 0x83
            return 0
        if this.Get("LockPosition", 0)
            return 1
        rect := Buffer(16)
        DllCall("GetWindowRect", "Ptr", hwnd, "Ptr", rect)
        x := (lParam & 0xFFFF), y := (lParam >> 16 & 0xFFFF)
        x := x >= 32768 ? x-65536 : x, y := y >= 32768 ? y-65536 : y
        left := x-NumGet(rect, 0, "Int") < 4, right := NumGet(rect, 8, "Int")-x <= 4
        top := y-NumGet(rect, 4, "Int") < 4, bottom := NumGet(rect, 12, "Int")-y <= 4
        return top ? (left ? 13 : right ? 14 : 12) : bottom ? (left ? 16 : right ? 17 : 15) : left ? 10 : right ? 11 : 1
    }
    EventStart(event, hwnd) {
        if !this.Get("Enabled", 0)
            return
        if this.HasOwnProp("tick")
            SetTimer(this.tick, Max(30, Min(250, this.owner.flashBorderInterval)))
        this.events[hwnd] := {event: event, start: A_TickCount, expires: A_TickCount+this.owner.eventDisplayDuration}
        this.signature := ""
    }
    EventEnd(hwnd) {
        if this.events.Has(hwnd)
            this.events.Delete(hwnd)
        if !this.events.Count && this.HasOwnProp("tick")
            SetTimer(this.tick, 250)
        this.signature := ""
    }
    EventState(hwnd) {
        if !this.events.Has(hwnd)
            return ["", ""]
        item := this.events[hwnd]
        if A_TickCount >= item.expires {
            this.events.Delete(hwnd)
            if !this.events.Count && this.HasOwnProp("tick")
                SetTimer(this.tick, 250)
            return ["", ""]
        }
        focused := this.ActiveClient() = hwnd
        if focused && this.owner.supressForFocused
            return ["", ""]
        text := this.owner.showEventText ? this.owner.monitoredEventsTexts[item.event] : ""
        blink := this.owner.flashBorderEnabled && Mod(Floor((A_TickCount-item.start)/Max(50, this.owner.flashBorderInterval)), 2) = 0
        color := blink ? convertToHex(this.owner.monitoredEvents[item.event]["color"]) : ""
        return [text, color]
    }
    ScrollAt(x, y) {
        horizontal := this.Get("Direction", "vertical") = "horizontal"
        fraction := horizontal ? (x-8)/Max(1, this.contentWidth-16) : (y-36)/Max(1, this.contentHeight-48)
        this.offset := Round(Max(0, Min(1, fraction))*this.maxOffset)
        this.signature := ""
        this.Render()
    }
    DropAt(x, y) {
        horizontal := this.Get("Direction", "vertical") = "horizontal", cw := this.Get("ColumnWidth", 220)
        lane := Min(this.Get("Columns", 1)-1, Max(0, horizontal ? Floor((y-36)/32) : Floor((x-8)/cw)))
        along := (horizontal ? (x-8)/cw : (y-36)/32)+this.offset
        laneLength := 0
        for slot in this.layout
            if slot.lane = lane
                laneLength := Max(laneLength, slot.along+1)
        boundary := Min(laneLength, Max(0, Floor(along+0.5))), before := ""
        for slot in this.layout
            if slot.lane = lane && slot.along >= boundary {
                before := this.items[slot.index].key
                break
            }
        px := 8+(horizontal ? boundary-this.offset : lane)*cw
        py := 36+(horizontal ? lane : boundary-this.offset)*32
        this.insertion.Move(horizontal ? px-3 : px, horizontal ? py : py-2, horizontal ? 3 : cw-8, horizontal ? 28 : 3)
        this.insertion.Visible := true
        DllCall("SetWindowPos", "Ptr", this.insertion.Hwnd, "Ptr", 0, "Int", 0, "Int", 0, "Int", 0, "Int", 0, "UInt", 0x13)
        return {lane: lane, before: before}
    }
    ApplyDrop(index, drop) {
        entry := this.items[index]
        if drop.before = entry.key
            return
        assignments := Map()
        for slot in this.layout
            assignments[this.items[slot.index].key] := slot.lane
        assignments[entry.key] := drop.lane
        this.items.RemoveAt(index)
        destination := this.items.Length+1
        for i, item in this.items
            if item.key = drop.before {
                destination := i
                break
            }
        this.items.InsertAt(destination, entry), order := []
        for item in this.items
            order.Push(item.key)
        this.Put("Lanes", assignments), this.Put("AutoFill", 0), this.Put("Order", order), this.Put("Sort", "manual")
    }
    MouseDown(wParam, lParam, msg, hwnd) {
        if !this.gui
            return
        if hwnd = this.track.Hwnd || hwnd = this.thumb.Hwnd {
            this.dragging := true
            try {
                while GetKeyState("LButton", "P") {
                    MouseGetPos(&x, &y)
                    this.gui.GetClientPos(&gx, &gy)
                    this.ScrollAt(x-gx, y-gy)
                    Sleep(16)
                }
            } finally this.dragging := false
            return 0
        }
        if hwnd = this.heading.Hwnd || hwnd = this.gui.Hwnd {
            if !this.Get("LockPosition", 0)
                PostMessage(0xA1, 2,,, this.gui.Hwnd)
            return
        }
        if !this.cells.Has(hwnd)
            return
        index := this.cells[hwnd], entry := this.items[index]
        MouseGetPos(&startX, &startY)
        this.dragging := true, moved := false, drop := 0
        try {
            while GetKeyState("LButton", "P") {
                MouseGetPos(&x, &y, &window)
                if Abs(x-startX)+Abs(y-startY) >= 6 {
                    moved := true
                    this.gui.GetClientPos(&gx, &gy, &gw, &gh)
                    if x >= gx+4 && x < gx+gw-4 && y >= gy+32 && y < gy+gh-12 {
                        drop := this.DropAt(x-gx, y-gy)
                        ToolTip("Moving " entry.name " · insert at line")
                    } else {
                        drop := 0, this.insertion.Visible := false
                        ToolTip("Moving " entry.name " · choose a gap")
                    }
                }
                Sleep(16)
            }
            if !moved {
                try this.owner.ActivateEVEWindow(entry.hwnd)
            } else if drop
                this.ApplyDrop(index, drop)
        } finally {
            this.insertion.Visible := false
            ToolTip(), this.dragging := false, this.signature := ""
        }
        return 0
    }
}
