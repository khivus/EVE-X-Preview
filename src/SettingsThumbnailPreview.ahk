; A separate renderer uses the real thumbnail windows and text layout, with
; sample data in place of a live client. It never joins the client's HWND maps.
class SettingsThumbnailPreview extends Main_Class {
    ; The isolated demo intentionally simulates an active client.
    EventSourceActive(hwnd) => true
    static SampleDuration := 1000
    static MouseMessages := [0x201, 0x204, 0x208] ; Left/right down, middle up.
    static Systems := ["Jita", "Dodixie", "Amarr", "Hek", "Tama", "Turnur", "Ahbazon"]
    static CharacterNames := ["Test Character", "Test Scout", "Test Logistics Wing Commander",
        "Test Mining Support Pilot", "Test Fleet Commander", "Test Explorer", "Test Capital Pilot"]
    ; Incoming/outgoing DPS and whole-number EM/Thermal/Kinetic/Explosive shares.
    static CombatSamples := [
        [2500, 1800, [25, 25, 25, 25]], [875, 2400, [60, 20, 15, 5]],
        [4250, 950, [0, 70, 30, 0]], [120, 5600, [10, 0, 20, 70]],
        [8100, 320, [0, 0, 100, 0]], [1650, 3750, [35, 15, 5, 45]],
        [3400, 6900, [0, 0, 0, 100]]
    ]

    __New(settings) {
        ; Deliberately skip Main_Class.__New(): no client/log polling or hotkeys.
        this.settings := settings
        this.closed := false
        this.snapshot := ""
        this.nextSettingsAt := 0
        this.sample := 0
        this.sampleHwnd := 0
        this.sampleStatusHwnd := 0
        this.sampleReady := false
        this.sampleStep := 0
        this.margins := Buffer(16, 0)
        this.DisabledChars := Map()
        this.QuickGroupChars := Map()
        this.HidedThumbs := Map()
        this.flashMethod := Map()
        this.gesture := 0
        this.timer := ObjBindMethod(this, "Tick")
        this.gestureTimer := ObjBindMethod(this, "GestureTick")
        this.mouseCallback := ObjBindMethod(this, "SampleMouseMessage")
        try {
            this.Refresh()
            if !IsObject(this.sample)
                throw ValueError("Complete the thumbnail position, size and text-size fields before opening the test thumbnail.")
            for message in SettingsThumbnailPreview.MouseMessages
                OnMessage(message, this.mouseCallback)
            SetTimer(this.timer, 100)
        } catch as err {
            this.Close()
            throw err
        }
    }

    ReadSettings() {
        app := this.settings
        groups := Map()
        for group, property in Map("Thumbnails Behavior", "ProfileThumbnailsBehavior",
            "Thumbnails Visuals", "ProfileThumbnailsVisuals", "Custom Colors", "ProfileCustomColors",
            "Thumbnails Interactions", "ProfileThumbnailsInteractions", "Game Logs Monitoring", "ProfileGameLogsMonitoring",
            "DPS Monitoring", "ProfileDPSMonitoring")
            groups[group] := app.Profiles[app.%property%][group]
        ; Monitored Events inherits independently, despite sharing JSON storage.
        monitoring := groups["Game Logs Monitoring"].Clone()
        monitoring["monitoredEvents"] := app.Profiles[app.ProfileMonitoredEvents]["Game Logs Monitoring"]["monitoredEvents"]
        groups["Game Logs Monitoring"] := monitoring
        ; These two colors are stored in the interactions profile's legacy group.
        interactions := app.Profiles[app.ProfileThumbnailsInteractions]["Hotkeys Settings"]
        groups["Hotkeys Settings"] := Map("dynamicGroupsColor", interactions["dynamicGroupsColor"], "QuickGroupColor", interactions["QuickGroupColor"])
        groups["__PreviewProfile"] := app.LastUsedProfile
        return JSON.Dump(groups)
    }

    Tick(*) {
        if this.closed
            return
        settings := this.settings
        try {
            if !DllCall("IsWindow", "Ptr", settings.S_Gui.Hwnd) {
                settings.StopTestThumbnail()
                return
            }
            this.Refresh()
        } catch as err {
            ProgramLog.Error(err, "Settings test thumbnail")
            settings.StopTestThumbnail()
        }
    }

    Refresh(now := A_TickCount) {
        if this.closed || IsObject(this.gesture)
            return
        if now >= this.nextSettingsAt {
            this.nextSettingsAt := now + 250
            snapshot := this.ReadSettings()
            if snapshot != this.snapshot {
                groups := JSON.Load(snapshot) ; Keep all preview changes in an independent copy.
                if this.ValidNumbers(groups) {
                    this.Rebuild(groups, now)
                    this.snapshot := snapshot
                }
            }
        }
        if !IsObject(this.sample)
            return
        if now >= this.nextSampleAt
            this.NextSample(now)
        else
            this.PaintSample(now)
    }

    ValidNumbers(groups) {
        position := groups["Thumbnails Behavior"]["ThumbnailStartLocation"]
        visuals := groups["Thumbnails Visuals"]
        for key in ["x", "y", "width", "height"]
            if !IsNumber(position[key]) || ((key = "width" || key = "height") && position[key] < 1)
                return false
        for key in ["ThumbnailTextSize", "ThumbnailOpacity", "ClientHighligtBorderthickness", "InactiveClientBorderthickness"]
            if !IsNumber(visuals.Get(key, 2))
                return false
        for key in ["x", "y"]
            if !IsNumber(visuals["ThumbnailTextMargins"][key])
                return false
        if visuals["ThumbnailTextSize"] <= 0
            return false
        return true
    }

    Rebuild(groups, now) {
        previous := IsObject(this.sample) ? this.samples[this.sampleIndex].id : "active"
        position := groups["Thumbnails Behavior"]["ThumbnailStartLocation"]
        defaults := [position["x"], position["y"], position["width"], position["height"]]
        profile := groups.Delete("__PreviewProfile")
        geometry := defaults
        if IsObject(this.sample) && this.previewProfile = profile && JSON.Dump(defaults) = JSON.Dump(this.defaultGeometry) {
            WinGetPos(&x, &y, &w, &h, this.sample["Window"].Hwnd)
            geometry := [x, y, w, h] ; Appearance edits retain the temporary move/resize.
        }
        this.previewProfile := profile
        this.defaultGeometry := defaults
        this.DisposeSample()
        groups["Thumbnail Visibility"] := Map()
        this._JSON := Map("LastUsedProfile", "Default", "_Profiles", Map("Default", groups))
        for property in ["ProfileThumbnailsBehavior", "ProfileThumbnailsVisuals", "ProfileCustomColors",
            "ProfileThumbnailsInteractions", "ProfileMonitoredEvents", "ProfileGameLogsMonitoring", "ProfileDPSMonitoring", "ProfileThumbnailVisibility"]
            this.%property% := "Default"
        this.NormalizeColors(groups)
        this.title := this.SampleTitle()
        this.sample := this.Create_Thumbnail(0, this.title)
        this.sampleHwnd := this.sample["Window"].Hwnd
        this.sampleStatusHwnd := this.sample["Window"]["SampleStatus"].Hwnd
        this.ThumbWindows.sample := this.sample
        this.sample["Window"].Opt("+Owner" this.settings.S_Gui.Hwnd)
        this.ThumbClickThrough(this.sample, true)
        this.BuildSampleInteractions()
        this.ApplyGeometry(geometry*)
        this.samples := this.BuildSamples()
        this.sampleIndex := 1
        for index, sample in this.samples
            if sample.id = previous
                this.sampleIndex := index
        this.StartSample(now)
        this.PaintSample(now)
        this.sampleReady := true
    }

    NormalizeColors(groups) {
        visuals := groups["Thumbnails Visuals"]
        for key, fallback in Map("ThumbnailTextColor", "FAC57A", "ThumbnailBackgroundColor", "57504E",
            "ClientHighligtColor", "E36A0D", "InactiveClientBorderColor", "8A8A8A")
            visuals[key] := ColorPicker.Normalize(visuals.Get(key, fallback), fallback)
        ; Invalid/incomplete manual colors affect only this sample, without dialogs.
        for key in ["TextColor", "Bordercolor", "IABordercolor"] {
            colors := groups["Custom Colors"]["cColors"]
            if !colors.Has(key)
                continue
            for index, color in colors[key]
                if color != ""
                    colors[key][index] := ColorPicker.Normalize(color)
        }
        visuals["ThumbnailOpacity"] := Min(100, Max(0, visuals["ThumbnailOpacity"]))
        for key in ["ClientHighligtBorderthickness", "InactiveClientBorderthickness"]
            visuals[key] := Max(0, Round(visuals.Get(key, 2)))
    }

    SampleTitle() {
        names := []
        if this.CustomColorsActive {
            for name in this.Profiles["Default"]["Custom Colors"]["cColors"]["CharNames"]
                if Trim(name) != ""
                    names.Push(name)
            if names.Length
                return names[Mod(this.sampleStep, names.Length) + 1]
        }
        for name in StrSplit(this.CustomThumbnailNames["Characters"], "`n", "`r")
            if Trim(name) != ""
                names.Push(this.AntiCleanTitle(Trim(name)))
        if names.Length
            return names[Mod(this.sampleStep, names.Length) + 1]
        return "EVE - " SettingsThumbnailPreview.CharacterNames[Mod(this.sampleStep, SettingsThumbnailPreview.CharacterNames.Length) + 1]
    }

    UpdateSampleData() {
        this.title := this.SampleTitle()
        this.sample["Window"].Title := this.title
        overlay := this.sample["TextOverlay"]
        overlay["OverlayText"].Text := this.GetThumbnailDisplayText(this.title)
        color := this.ThumbnailTextColor
        if this.CustomColorsActive {
            custom := this.CustomColorsGet[this.title]
            if custom["Char"] != "" && custom["Text"] != ""
                color := custom["Text"]
        }
        for name in ["OverlayText", "EventText", "DPSText", "DPSDamageText", "SystemText"]
            overlay[name].SetFont("c" color)
        overlay.SystemTextDefaultColor := color
        overlay.SystemTextColor := ""
        index := Mod(this.sampleStep, SettingsThumbnailPreview.Systems.Length) + 1
        system := this.systemTrackingEnabled ? SettingsThumbnailPreview.Systems[index] : ""
        this.updateThumbnailSystemText(system, "sample")
        ; Demo-only colors always distinguish successive systems, even with the
        ; live system-color option off. Never allocate colors in the live map.
        this.ApplyThumbnailSystemColor(overlay, system != "" ? SystemColorManager.Palette[index] : "")
        this.SampleDPS()
        this.RefreshThumbnailTextLayout(overlay)
    }

    BuildSamples() {
        duration := SettingsThumbnailPreview.SampleDuration
        samples := [{id: "active", label: "Active client", duration: duration}, {id: "inactive", label: "Inactive client", duration: duration}]
        for action, label in Map("DisableFromGroups", "Excluded from groups", "QuickGroup", "Quick Group") {
            enabled := false
            for key, value in this.ThumbnailsInteractions[action]
                enabled := enabled || value
            if enabled
                samples.Push({id: action, label: label, duration: duration})
        }
        if this.gameLogsMonitoringEnabled && (this.showEventText || this.flashBorderEnabled)
            for event, config in this.monitoredEvents
                if config["enabled"]
                    samples.Push({id: "event:" event, label: this.monitoredEventsTexts[event], event: event,
                        duration: IsNumber(this.eventDisplayDuration) ? Max(250, Min(duration, this.eventDisplayDuration)) : duration})
        return samples
    }

    SampleDPS() {
        if !this.gameLogsMonitoringEnabled {
            this.updateThumbnailDPSText("", "sample")
            return
        }
        config := this.dpsMonitoring.Clone()
        for key in ["incomingDPSThreshold", "outgoingDPSThreshold"]
            if !IsNumber(config[key])
                config[key] := 0
        meter := DPSMeter(config)
        index := Mod(this.sampleStep, SettingsThumbnailPreview.CombatSamples.Length) + 1
        data := SettingsThumbnailPreview.CombatSamples[index]
        multiplier := 1.25 + (index - 1) * 0.25
        incoming := Ceil(Max(data[1], config["incomingDPSThreshold"] * multiplier)) * meter.WindowSeconds
        outgoing := Ceil(Max(data[2], config["outgoingDPSThreshold"] * multiplier)) * meter.WindowSeconds
        weapons := ["EMP Smartbomb", "Plasma Smartbomb", "Graviton Smartbomb", "Proton Smartbomb"]
        for type, percentage in data[3]
            if percentage > 0
                meter.AddHit({second: 1, incoming: true, amount: incoming * percentage / 100,
                    line: "(combat) 2500 from Test NPC - " weapons[type]}, 1)
        meter.AddHit({second: 1, incoming: false, amount: outgoing, line: ""}, 1)
        this.updateThumbnailDPSText(meter.Text(1), "sample", meter.damageText)
    }

    NextSample(now := A_TickCount) {
        if this.closed || !IsObject(this.sample) || IsObject(this.gesture)
            return
        this.sampleIndex := Mod(this.sampleIndex, this.samples.Length) + 1
        this.StartSample(now, true)
        this.PaintSample(now)
    }

    StartSample(now, advanceData := false) {
        if advanceData
            this.sampleStep++
        this.sampleHidden := false
        this.sampleSince := now
        this.nextSampleAt := now + this.samples[this.sampleIndex].duration
        this.lastPaint := ""
        this.UpdateSampleData()
    }

    PaintSample(now) {
        state := this.samples[this.sampleIndex]
        interval := IsNumber(this.flashBorderInterval) ? Max(50, this.flashBorderInterval) : 300
        flashOn := Mod(Floor((now - this.sampleSince) / interval), 2) = 0
        paint := this.sampleIndex ":" (state.HasOwnProp("event") && this.flashBorderEnabled ? flashOn : 0) ":" this.sampleHidden
        if paint = this.lastPaint
            return
        this.lastPaint := paint
        if this.sampleHidden {
            for key in ["Window", "TextOverlay", "Border", "EventBorder"]
                if this.sample.Has(key)
                    this.sample[key].Hide()
            return
        }
        separateEvent := state.HasOwnProp("event") && this.SeparateEventBorder(this.sample)
        this.sample["Window"]["SampleStatus"].Text := "Test thumbnail`n" (separateEvent ? "Active + " : "") state.label
        this.updateThumbnailEventText(state.HasOwnProp("event") && this.showEventText ? state.label : "", "sample")
        this.sample["Window"].Show("NoActivate")
        this.sample["TextOverlay"].Show(this.ShowThumbnailTextOverlay ? "NoActivate" : "Hide")
        this.HideEventBorder("sample")
        this.clearBorder("sample", this.title)
        if state.id = "active"
            this.activateBorder("sample", this.title)
        else if state.id = "DisableFromGroups" || state.id = "QuickGroup"
            this.toggleColorBorder("sample", this.title, 1, state.id = "QuickGroup" ? this.QuickGroupColor : this.DisableFromGroupsColor)
        else if state.HasOwnProp("event") && this.flashBorderEnabled {
            if separateEvent {
                this.activateBorder("sample", this.title)
                this.SetEventBorder("sample", this.monitoredEvents[state.event]["color"], flashOn)
            } else if flashOn
                this.toggleColorBorder("sample", this.title, 1, this.monitoredEvents[state.event]["color"])
        }
    }

    AttachLivePreview(thumb, hwnd) {
        ; A lightweight sample surface: no game capture or DWM registration.
        window := thumb["Window"]
        WinGetClientPos(,, &width, &height, window.Hwnd)
        scale := A_ScreenDPI / 96
        status := window.Add("Text", Format("x0 y{} w{} h36 vSampleStatus +0x101 BackgroundTrans", Round(height / scale / 3), Round(width / scale)), "Test thumbnail")
        status.SetFont("s9 cC0C0C0", "Segoe UI")
        status.ToolTip := "Use configured thumbnail gestures to move/resize. Activate-click or middle-click cycles samples; close from Settings."
        window.OnEvent("Size", ObjBindMethod(this, "LayoutSampleSurface"))
    }

    LayoutSampleSurface(window, minMax, width, height) {
        if minMax != -1
            window["SampleStatus"].Move(0, Round(height / 3), width, 36)
    }

    BuildSampleInteractions() {
        this.sampleInteractions := Map()
        for action, config in this.ThumbnailsInteractions {
            combination := 0
            for key, bit in Map("lmb", 1, "rmb", 2, "shift", 4, "ctrl", 8)
                combination |= config[key] ? bit : 0
            if combination && !this.sampleInteractions.Has(combination)
                this.sampleInteractions[combination] := action
        }
    }

    SampleMouseMessage(wParam, lParam, message, hwnd) {
        if this.closed || !this.sampleReady || !IsObject(this.sample)
            return
        ; Global mouse callbacks also run while Destroy dispatches window messages.
        ; Filter using cached handles, without accessing a possibly destroyed Gui.
        if hwnd != this.sampleHwnd && hwnd != this.sampleStatusHwnd
            return
        if message = 0x208 {
            this.NextSample()
            return 0
        }
        flags := wParam & 15
        action := this.sampleInteractions.Get(flags, "")
        if action = "MoveThumbnail" || action = "MoveAllThumbnails" || action = "ResizeThumbnail" || action = "ResizeAllThumbnails"
            this.BeginGesture(action, flags)
        else if !IsObject(this.gesture) {
            if action = "ActivateThumbnail" || (action = "" && (flags = 1 || flags = 2))
                this.NextSample()
            else if action = "DisableFromGroups" || action = "QuickGroup" || action = "MinimizeClient" {
                target := action = "MinimizeClient" || this.samples[this.sampleIndex].id = action ? "inactive" : action
                for index, state in this.samples
                    if state.id = target {
                        this.sampleIndex := index
                        this.StartSample(A_TickCount, true)
                        this.PaintSample(A_TickCount)
                        break
                    }
            } else if action = "HideThumbnail" {
                this.sampleHidden := true
                this.nextSampleAt := A_TickCount + SettingsThumbnailPreview.SampleDuration
                this.PaintSample(A_TickCount) ; Automatically return at the next sample.
            } else if action = "CloseClient"
                this.settings.StopTestThumbnail()
        }
        return 0 ; Configured gestures take priority over label/control clicks.
    }

    PointerState() {
        CoordMode("Mouse", "Screen")
        MouseGetPos(&x, &y)
        buttons := (DllCall("GetAsyncKeyState", "Int", 1, "Short") & 0x8000 ? 1 : 0)
            | (DllCall("GetAsyncKeyState", "Int", 2, "Short") & 0x8000 ? 2 : 0)
        return {x: x, y: y, buttons: buttons}
    }

    BeginGesture(action, flags) {
        if this.LockThumbnailPositions
            return
        point := this.PointerState()
        WinGetPos(&x, &y, &w, &h, this.sample["Window"].Hwnd)
        this.gesture := {resize: InStr(action, "Resize") > 0, buttons: flags & 3,
            x: point.x, y: point.y, lastX: point.x, lastY: point.y, left: x, top: y, width: w, height: h}
        DllCall("SetCapture", "Ptr", this.sample["Window"].Hwnd, "Ptr")
        SetTimer(this.gestureTimer, 10)
    }

    GestureTick(*) {
        if this.closed || !IsObject(this.gesture)
            return
        try {
            point := this.PointerState()
            gesture := this.gesture
            if (point.buttons & gesture.buttons) != gesture.buttons {
                this.EndGesture()
                return
            }
            if point.x = gesture.lastX && point.y = gesture.lastY
                return
            gesture.lastX := point.x, gesture.lastY := point.y
            dx := point.x - gesture.x, dy := point.y - gesture.y
            if gesture.resize {
                minW := this.ThumbnailMinimumSize["width"], minH := this.ThumbnailMinimumSize["height"]
                width := Max(IsNumber(minW) ? Max(1, minW) : 50, gesture.width + dx)
                height := Max(IsNumber(minH) ? Max(1, minH) : 50, gesture.height + dy)
                this.ApplyGeometry(gesture.left, gesture.top, width, height)
            } else
                this.ApplyGeometry(gesture.left + dx, gesture.top + dy, gesture.width, gesture.height)
            this.PaintSample(A_TickCount)
        } catch as err {
            this.EndGesture()
            ProgramLog.Error(err, "Test thumbnail gesture")
        }
    }

    ApplyGeometry(x, y, width, height) {
        for key in ["Window", "TextOverlay", "Border", "EventBorder"]
            if this.sample.Has(key)
                WinMove(x, y, width, height, this.sample[key].Hwnd)
        this.lastPaint := ""
    }

    EndGesture() {
        SetTimer(this.gestureTimer, 0)
        this.ReleaseSampleCapture()
        this.gesture := 0
        this.nextSampleAt := A_TickCount + SettingsThumbnailPreview.SampleDuration
        this.nextSettingsAt := 0
    }

    ReleaseSampleCapture() {
        if this.sampleHwnd && DllCall("GetCapture", "Ptr") = this.sampleHwnd
            DllCall("ReleaseCapture")
    }

    ApplyThumbnailShadow(hwnd, generation) {
        ; A failed sample shadow must not pause the application's live previews.
        hr := this.ExtendThumbnailFrame(hwnd)
        if hr
            ProgramLog.Add("Test thumbnail shadow unavailable: " Format("0x{:08X}", hr & 0xFFFFFFFF))
    }

    DisposeSample() {
        this.sampleReady := false
        this.ReleaseSampleCapture()
        sample := this.sample
        ; Unpublish every old layer before native cleanup can re-enter callbacks.
        this.sample := 0
        this.sampleHwnd := 0
        this.sampleStatusHwnd := 0
        this.ThumbWindows := {}
        if IsObject(sample)
            for key in ["EventBorder", "Border", "TextOverlay", "Window"]
                try sample[key].Destroy()
    }

    Close() {
        this.closed := true
        this.gesture := 0
        this.ReleaseSampleCapture()
        if this.HasOwnProp("mouseCallback") {
            for message in SettingsThumbnailPreview.MouseMessages
                OnMessage(message, this.mouseCallback, 0)
            this.DeleteProp("mouseCallback")
        }
        if this.HasOwnProp("gestureTimer") {
            SetTimer(this.gestureTimer, 0)
            this.DeleteProp("gestureTimer")
        }
        if this.HasOwnProp("timer") {
            SetTimer(this.timer, 0)
            this.DeleteProp("timer") ; Release the bound callback's reference to this renderer.
        }
        this.DisposeSample()
        this.settings := 0
    }
}
