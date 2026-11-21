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
        this.owner := owner, this.gui := 0, this.signature := "", this.offset := 0, this.pool := [], this.dividers := [], this.systemPool := [], this.eventBorders := [], this.cells := Map(), this.dragging := false
        CharacterList.Migrate(owner._JSON)
        this.events := Map(), this.systemPaint := Map()
        this.systemDrawCallback := ObjBindMethod(this, "DrawSystemLabel")
        OnMessage(0x2B, this.systemDrawCallback)
        this.dragTimer := ObjBindMethod(this, "PollDrag")
        this.cursorCallback := ObjBindMethod(this, "ClientCursor")
        OnMessage(0x20, this.cursorCallback)
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
                entries.Push({hwnd: hwnd, name: name, created: created, title: title, loggedOut: loggedOut, key: loggedOut ? "hwnd:" hwnd : name})
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
            keys.Push([entry.key, entry.name, this.SystemLabel(entry), this.EventState(entry.hwnd)])
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
    SystemLabel(entry) {
        try {
            if this.owner.systemTrackingEnabled && !entry.loggedOut
                return this.owner.systemLogMonitor.readers[entry.HasOwnProp("title") ? entry.title : entry.name].system
        }
        return ""
    }
    DrawSystemLabel(wParam, lParam, msg, hwnd) {
        if !lParam
            return
        control := NumGet(lParam, A_PtrSize = 8 ? 24 : 20, "Ptr")
        if !this.systemPaint.Has(control)
            return
        paint := this.systemPaint[control]
        dc := NumGet(lParam, A_PtrSize = 8 ? 32 : 24, "Ptr")
        rect := lParam+(A_PtrSize = 8 ? 40 : 28)
        saved := DllCall("SaveDC", "Ptr", dc, "Int")
        brush := DllCall("CreateSolidBrush", "UInt", SettingsTheme.BGR(paint.background), "Ptr")
        try {
            ; Erase the full field before right-aligning the new string. Native
            ; right-aligned static text can leave trailing glyphs after scrolling.
            DllCall("FillRect", "Ptr", dc, "Ptr", rect, "Ptr", brush)
            font := SendMessage(0x31, 0, 0,, control)
            if font
                DllCall("SelectObject", "Ptr", dc, "Ptr", font)
            DllCall("SetBkMode", "Ptr", dc, "Int", 1)
            DllCall("SetTextColor", "Ptr", dc, "UInt", SettingsTheme.BGR(paint.foreground))
            DllCall("DrawTextW", "Ptr", dc, "Str", paint.text, "Int", -1, "Ptr", rect, "UInt", 0x8826)
        } finally {
            DllCall("DeleteObject", "Ptr", brush)
            DllCall("RestoreDC", "Ptr", dc, "Int", saved)
        }
        return true
    }
    Place(control, x, y, width, height) {
        control.GetPos(&oldX, &oldY, &oldWidth, &oldHeight)
        if x != oldX || y != oldY || width != oldWidth || height != oldHeight
            control.Move(x, y, width, height)
    }
    PaintText(control, text, background, font) {
        style := background "|" font
        if !control.HasOwnProp("listPaintStyle") || control.listPaintStyle != style {
            control.Opt("Background" background)
            control.SetFont(font)
            control.listPaintStyle := style
        }
        if control.Text != text
            control.Text := text
    }
    Render() {
        this.rendering := true
        try {
            if !this.gui {
                this.gui := Gui("-Caption -Border -Resize +ToolWindow +E0x02000000 -DPIScale", "EVE character list")
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
            wasVisible := DllCall("IsWindowVisible", "Ptr", g.Hwnd)
            SendMessage(0xB, 0, 0,, g.Hwnd) ; WM_SETREDRAW
            if g.BackColor != colors["background"]
                g.BackColor := colors["background"]
            g.SetFont("s10 c" colors["foreground"], "Segoe UI")
            this.PaintText(this.arrow, collapsed ? "▶" : "▼", colors["background"], "c" colors["foreground"])
            this.PaintText(this.heading, width >= 140 ? "Characters · " this.items.Length : "", colors["background"], "c" colors["foreground"])
            this.PaintText(this.dots, "⋯", colors["background"], "s14 c" colors["foreground"])
            this.heading.Move(30, 5, Max(1, width-64), 24), this.dots.Move(width-28, 2, 24, 28)
            ; Resizing is handled in client space; never enable the native sizing frame.
            if WinGetStyle(g.Hwnd) & 0xC40000
                g.Opt("-Caption -Border -Resize")
            topmost := this.Get("AlwaysOnTop", 1)
            if !!(WinGetExStyle(g.Hwnd) & 8) != !!topmost
                WinSetAlwaysOnTop(topmost, g.Hwnd)
            horizontal := this.Get("Direction", "vertical") = "horizontal"
            visible := horizontal ? Max(1, Floor((width-16)/cw)) : Max(1, Floor((height-44)/32))
            this.layout := this.BuildLayout(this.items, visible)
            extent := 0
            for slot in this.layout
                extent := Max(extent, slot.along+1)
            ; Saved height is a viewport limit, not a permanent empty hit area.
            height := horizontal ? 48+cols*32 : 44+Min(visible, Max(1, extent))*32
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
                        this.pool.Push(g.AddText("x0 y0 w1 h28 +0x100 +0x200 +0x4000", ""))
                        this.systemPool.Push(g.AddText("x0 y0 w1 h28 +0x10D", ""))
                    }
                    cell := this.pool[count]
                    x := 8+(horizontal ? slot.along-this.offset : slot.lane)*cw
                    y := 36+(horizontal ? slot.lane : slot.along-this.offset)*32
                    if horizontal && y+28 > height-8
                        continue
                    system := this.SystemLabel(entry), systemCell := this.systemPool[count]
                    systemWidth := system = "" ? 0 : Min(cw-28, Max(20, Round((cw-4)*0.4)))
                    this.Place(cell, x, y, cw-4-systemWidth, 28)
                    this.Place(systemCell, x+cw-4-systemWidth, y, Max(1, systemWidth), 28)
                    active := this.HasOwnProp("lastActive") && entry.hwnd = this.lastActive
                    event := this.EventState(entry.hwnd)
                    background := active ? colors["selection"] : colors["background"]
                    border := this.eventBorders[count]
                    this.Place(border, x-1, y-1, cw-2, 30)
                    border.Visible := event[2] != ""
                    if event[2] != ""
                        border.Opt("Background" event[2])
                    label := entry.name
                    if entry.loggedOut && this.Get("ShowHwnd", 0)
                        label .= " [" Format("0x{:X}", entry.hwnd) "]"
                    if event[1] != ""
                        label .= " · " event[1]
                    font := (active ? "bold " : "norm ") "c" colors["foreground"]
                    this.PaintText(cell, label, background, font)
                    this.systemPaint[systemCell.Hwnd] := {text: system, background: background, foreground: colors["foreground"]}
                    this.PaintText(systemCell, system, background, font)
                    systemCell.Visible := system != ""
                    cell.Visible := true, this.cells[cell.Hwnd] := slot.index
                    if system != ""
                        this.cells[systemCell.Hwnd] := slot.index
                    this.placements[cell.Hwnd] := {lane: slot.lane, along: slot.along, x: x, y: y, index: slot.index}
                }
            for i, cell in this.pool
                if i > count {
                    cell.Visible := false
                    this.eventBorders[i].Visible := false
                    this.systemPool[i].Visible := false
                }
            dividerCount := collapsed ? 0 : Max(0, (horizontal ? visible : cols)-1)
            Loop dividerCount {
                i := A_Index
                if i > this.dividers.Length
                    this.dividers.Push(g.AddText("x0 y0 w1 h1", ""))
                divider := this.dividers[i]
                this.Place(divider, 8+i*cw-2, 36, 1, Max(1, height-48))
                this.PaintText(divider, "", colors.Get("line", colors["selection"]), "norm")
                divider.Visible := true
            }
            for i, divider in this.dividers
                if i > dividerCount
                    divider.Visible := false
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
            targetHeight := collapsed ? 32 : height
            this.renderedWidth := width, this.renderedHeight := targetHeight
            g.GetPos(&currentX, &currentY, &currentWidth, &currentHeight)
            if wx != currentX || wy != currentY || width != currentWidth || targetHeight != currentHeight || !wasVisible
                DllCall("SetWindowPos", "Ptr", g.Hwnd, "Ptr", 0, "Int", wx, "Int", wy, "Int", width, "Int", targetHeight, "UInt", 0x54)
            opacity := Round(Max(10, this.Get("Opacity", 100))*2.55)
            if !this.HasOwnProp("paintedOpacity") || this.paintedOpacity != opacity {
                WinSetTransparent(opacity, g.Hwnd)
                this.paintedOpacity := opacity
            }
        } finally {
            if this.gui {
                SendMessage(0xB, 1, 0,, this.gui.Hwnd)
                DllCall("RedrawWindow", "Ptr", this.gui.Hwnd, "Ptr", 0, "Ptr", 0, "UInt", 0x185)
            }
            this.rendering := false
        }
    }
    Resize(guiObj, minMax, width, height) {
        if minMax = -1 || this.Get("Collapsed", 0) || (this.HasOwnProp("rendering") && this.rendering) || (this.HasOwnProp("renderedWidth") && width = this.renderedWidth && height = this.renderedHeight)
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
        nextOffset := Min(this.maxOffset, Max(0, this.offset+((wParam >> 16 & 0x8000) ? 1 : -1)))
        if nextOffset = this.offset
            return 0
        this.offset := nextOffset, this.signature := ""
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
        return msg = 0x83 ? 0 : 1
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
        nextOffset := Round(Max(0, Min(1, fraction))*this.maxOffset)
        if nextOffset = this.offset
            return
        this.offset := nextOffset, this.signature := ""
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
        this.insertion.Move(horizontal ? px-3 : px, horizontal ? py : py-2, horizontal ? 3 : cw-4, horizontal ? 28 : 3)
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
    EdgeAt(x, y, width, height) {
        if this.Get("LockPosition", 0)
            return ""
        return (x < 5 ? "L" : x >= width-5 ? "R" : "") (y < 5 ? "T" : y >= height-5 ? "B" : "")
    }
    ClientCursor(wParam, lParam, msg, hwnd) {
        if !this.gui || hwnd != this.gui.Hwnd || this.Get("LockPosition", 0)
            return
        point := Buffer(8)
        DllCall("GetCursorPos", "Ptr", point)
        this.gui.GetClientPos(&gx, &gy, &width, &height)
        edge := this.EdgeAt(NumGet(point, 0, "Int")-gx, NumGet(point, 4, "Int")-gy, width, height)
        if edge = ""
            return
        cursor := StrLen(edge) = 2 ? ((edge = "LT" || edge = "RB") ? 32642 : 32643) : (edge = "L" || edge = "R") ? 32644 : 32645
        DllCall("SetCursor", "Ptr", DllCall("LoadCursor", "Ptr", 0, "Ptr", cursor, "Ptr"))
        return 1
    }
    BeginDrag(mode, x, y, index := 0, edge := "") {
        this.gui.GetPos(&wx, &wy, &width, &height)
        this.dragState := {mode: mode, startX: x, startY: y, wx: wx, wy: wy, width: width, height: height, index: index, edge: edge, moved: false, drop: 0}
        this.dragging := true
        DllCall("SetCapture", "Ptr", this.gui.Hwnd)
        SetTimer(this.dragTimer, 16)
    }
    PollDrag(*) {
        if !this.dragging
            return
        if DllCall("GetCapture", "Ptr") != this.gui.Hwnd || (DllCall("GetAsyncKeyState", "Int", 0x1B, "Short") & 0x8000) {
            this.FinishDrag(true)
            return
        }
        point := Buffer(8)
        DllCall("GetCursorPos", "Ptr", point)
        try this.UpdateDrag(NumGet(point, 0, "Int"), NumGet(point, 4, "Int"), !!(DllCall("GetAsyncKeyState", "Int", 1, "Short") & 0x8000))
        catch as err {
            this.FinishDrag(true)
            throw err
        }
    }
    UpdateDrag(x, y, pressed) {
        state := this.dragState, dx := x-state.startX, dy := y-state.startY
        if Abs(dx)+Abs(dy) >= 6
            state.moved := true
        if state.mode = "row" && state.moved {
            this.gui.GetClientPos(&gx, &gy, &gw, &gh)
            if x >= gx+4 && x < gx+gw-4 && y >= gy+32 && y < gy+gh-12 {
                state.drop := this.DropAt(x-gx, y-gy)
                ToolTip("Moving " this.items[state.index].name " · insert at line")
            } else {
                state.drop := 0, this.insertion.Visible := false
                ToolTip("Moving " this.items[state.index].name " · choose a gap")
            }
        } else if state.mode = "scroll" {
            this.gui.GetClientPos(&gx, &gy)
            this.ScrollAt(x-gx, y-gy)
        } else if state.mode = "move" {
            DllCall("SetWindowPos", "Ptr", this.gui.Hwnd, "Ptr", 0, "Int", state.wx+dx, "Int", state.wy+dy, "Int", 0, "Int", 0, "UInt", 0x15)
        } else if state.mode = "resize" {
            width := Max(66, state.width+(InStr(state.edge, "L") ? -dx : InStr(state.edge, "R") ? dx : 0))
            height := Max(100, state.height+(InStr(state.edge, "T") ? -dy : InStr(state.edge, "B") ? dy : 0))
            wx := state.wx+(InStr(state.edge, "L") ? state.width-width : 0)
            wy := state.wy+(InStr(state.edge, "T") ? state.height-height : 0)
            DllCall("SetWindowPos", "Ptr", this.gui.Hwnd, "Ptr", 0, "Int", wx, "Int", wy, "Int", width, "Int", height, "UInt", 0x14)
        }
        if !pressed
            this.FinishDrag(false)
    }
    FinishDrag(cancel := false) {
        if !this.dragging
            return
        state := this.dragState
        SetTimer(this.dragTimer, 0)
        this.dragging := false
        if DllCall("GetCapture", "Ptr") = this.gui.Hwnd
            DllCall("ReleaseCapture")
        this.insertion.Visible := false
        ToolTip()
        if !cancel {
            if state.mode = "row" {
                if !state.moved {
                    try this.owner.ActivateEVEWindow(this.items[state.index].hwnd)
                } else if state.drop
                    this.ApplyDrop(state.index, state.drop)
            } else if state.mode = "move" || state.mode = "resize"
                this.SaveBounds(0, 0, 0, this.gui.Hwnd)
        }
        this.signature := ""
    }
    MouseDown(wParam, lParam, msg, hwnd) {
        if !this.gui || this.dragging
            return
        point := Buffer(8)
        DllCall("GetCursorPos", "Ptr", point)
        x := NumGet(point, 0, "Int"), y := NumGet(point, 4, "Int")
        if hwnd = this.track.Hwnd || hwnd = this.thumb.Hwnd {
            this.BeginDrag("scroll", x, y)
            return 0
        }
        if this.cells.Has(hwnd) {
            this.BeginDrag("row", x, y, this.cells[hwnd])
            return 0
        }
        if hwnd = this.heading.Hwnd || hwnd = this.gui.Hwnd {
            this.gui.GetClientPos(&gx, &gy, &width, &height)
            edge := this.EdgeAt(x-gx, y-gy, width, height)
            if edge != ""
                this.BeginDrag("resize", x, y, 0, edge)
            else if !this.Get("LockPosition", 0)
                this.BeginDrag("move", x, y)
            return 0
        }
    }
}
