DPSSettings() {
    settings := JSON.Load(default_JSON)["_Profiles"]["Default"]["DPS Monitoring"]
    settings["incomingDPSEnabled"] := 1
    settings["outgoingDPSEnabled"] := 1
    return settings
}

DPSLine(second, damage, direction := "from", target := "Pilot[CORP](Ship)", weapon := "800mm Repeating Cannon II") {
    timestamp := FormatTime(DateAdd("19700101000000", second, "Seconds"), "yyyy.MM.dd HH:mm:ss")
    return "[ " timestamp " ] (combat) <color=0xff00ffff><b>" damage "</b> <color=0x77ffffff><font size=10>" direction "</font> <b><color=0xffffffff>" target "</b><font size=10><color=0x77ffffff> - " weapon " - Hits`r`n"
}

TestRunner.Register("Reported red incoming hits calculate DPS with damage percentages enabled", DPSReportedIncomingTest)
DPSReportedIncomingTest() {
    settings := DPSSettings()
    settings["showIncomingDPSResistances"] := 1
    meter := DPSMeter(settings)
    now := DateDiff("20260912025630", "19700101000000", "Seconds")
    hits := [[66, "Hammerhead I"], [119, "Inferno Heavy Assault Missile"], [18, "Modal Electron Particle Accelerator I"], [88, "425mm Medium Carbine Repeating Cannon I"], [68, "Heavy Anode Pulse Particle Stream I"], [21, "Modal Electron Particle Accelerator I"], [119, "Inferno Heavy Assault Missile"]]
    for index, hit in hits {
        line := RTrim(StrReplace(DPSLine(now - 7 + index, hit[1], "from", "Khivus[5TEAK](Gnosis)", hit[2]), "0xff00ffff", "0xffcc0000"), "`r`n")
        meter.AddLine(line, now)
    }
    AssertEqual(49.9, meter.Rates(now).incoming)
    AssertEqual(0, meter.Rates(now).outgoing)
    AssertArrayEqual([0, 304, 0, 0, 195], meter.Rates(now).types)
    AssertTrue(InStr(meter.Text(now), "In:") > 0)
    AssertEqual("% Th 61 ? 39", meter.damageText)
}

TestRunner.Register("DPS uses compact whole numbers and omits zero damage percentages", DPSCompactTextTest)
DPSCompactTextTest() {
    settings := DPSSettings()
    settings["showIncomingDPSResistances"] := 1
    meter := DPSMeter(settings)
    now := DPSMeter.CurrentSecond()
    meter.AddLine(DPSLine(now, 999, "from", "Pilot", "Hammerhead I"), now)
    meter.AddLine(DPSLine(now, 1, "from", "Pilot", "Mjolnir Torpedo"), now)
    meter.AddLine(DPSLine(now, 506, "to"), now)
    AssertEqual("DPS: In: 100 | Out: 51", meter.Text(now))
    AssertEqual("% Th 100", meter.damageText, "Types that round to zero must not reserve space.")
    AssertEqual(50.6, meter.Rates(now).outgoing, "Only the displayed rate is rounded.")
    AssertEqual("", meter.Text(now + 10))
    AssertEqual("", meter.damageText)
}

TestRunner.Register("DPS sums damage by log time and expires during idle periods", DPSWindowTest)
DPSWindowTest() {
    meter := DPSMeter(DPSSettings())
    now := DateDiff("20260912000005", "19700101000000", "Seconds")
    meter.AddLine(DPSLine(now - 9, 1000), now) ; Previous UTC date.
    meter.AddLine(DPSLine(now, 250.5), now)
    meter.AddLine(DPSLine(now, 500, "to"), now)
    meter.AddLine(DPSLine(now - 10, 9000), now)
    meter.AddLine(DPSLine(now + 1, 9000), now)
    AssertEqual(125.05, meter.Rates(now).incoming)
    AssertEqual(50, meter.Rates(now).outgoing)
    AssertEqual(25.05, meter.Rates(now + 1).incoming, "Oldest second expires at the window boundary.")
    AssertEqual("", meter.Text(now + 10), "Quiet DPS must clear without new log writes.")
    AssertEqual(0, meter.buckets.Count)
    Loop 1000
        meter.AddLine(DPSLine(now, 1), now)
    AssertEqual(1, meter.buckets.Count, "Hits in the same second share one bucket.")
    AssertEqual(100, meter.Rates(now).incoming)
}

TestRunner.Register("Configurable DPS window retains slow volleys and scales both rates", DPSAverageWindowTest)
DPSAverageWindowTest() {
    settings := DPSSettings()
    settings["dpsAverageSeconds"] := 20
    meter := DPSMeter(settings)
    serverSecond := DPSMeter.CurrentSecond() + 3
    meter.AddBatch([DPSLine(serverSecond, 2000), DPSLine(serverSecond, 1000, "to")], 100000)
    AssertEqual(100, meter.Rates(meter.LogSecond(117000)).incoming, "A volley remains visible between 17-second shots.")
    AssertEqual(50, meter.Rates(meter.LogSecond(117000)).outgoing)
    meter.AddBatch([DPSLine(serverSecond + 17, 2000), DPSLine(serverSecond + 17, 1000, "to")], 117000)
    AssertEqual(200, meter.Rates(meter.LogSecond(117000)).incoming)
    AssertEqual(100, meter.Rates(meter.LogSecond(117000)).outgoing)
    AssertEqual(100, meter.Rates(meter.LogSecond(120000)).incoming, "The first volley expires at the configured boundary.")
    AssertEqual("", meter.Text(meter.LogSecond(137000)))
    for invalid in [0, -1, 1.5, "invalid", 3601] {
        settings["dpsAverageSeconds"] := invalid
        AssertEqual(10, DPSMeter(settings).WindowSeconds)
    }
    settings.Delete("dpsAverageSeconds")
    AssertEqual(10, DPSMeter(settings).WindowSeconds, "Older settings retain the ten-second default.")
}

TestRunner.Register("Live DPS tolerates server clock offsets and expires using elapsed time", DPSLiveClockTest)
DPSLiveClockTest() {
    pcSecond := DPSMeter.CurrentSecond()
    for offset in [3, -30, 3600] {
        meter := DPSMeter(DPSSettings())
        serverSecond := pcSecond + offset
        meter.AddBatch([DPSLine(serverSecond - 11, 9000), DPSLine(serverSecond, 1000), DPSLine(serverSecond, 500, "to")], 100000)
        AssertEqual(serverSecond, meter.LogSecond(100000))
        AssertEqual(100, meter.Rates(meter.LogSecond(100000)).incoming, "Fresh hits must not depend on Windows UTC matching the server.")
        AssertEqual(50, meter.Rates(meter.LogSecond(100000)).outgoing)
        meter.AddBatch([DPSLine(serverSecond + 2, 200)], 102200)
        AssertEqual(120, meter.Rates(meter.LogSecond(102200)).incoming)
        AssertEqual(20, meter.Rates(meter.LogSecond(110000)).incoming, "Each log second retains its own expiry.")
        AssertEqual("", meter.Text(meter.LogSecond(112000)), "No writes are needed for idle expiry.")
        AssertEqual(0, meter.buckets.Count)
        meter.AddBatch([DPSLine(serverSecond, 9000)], 120000)
        AssertEqual("", meter.Text(meter.LogSecond(120000)), "An old delayed batch must not move the live clock backwards.")
    }
}

TestRunner.Register("Default live reader accepts the observed future log timestamps", DPSLiveReaderClockTest)
DPSLiveReaderClockTest() {
    oldHidden := A_DetectHiddenWindows
    DetectHiddenWindows(true)
    app := DPSLogFixture()
    try {
        serverSecond := DPSMeter.CurrentSecond() + 3
        path := app.WriteLog("20260912_000000_12345.txt", DPSLine(serverSecond, 9000))
        app.startLogMonitoring(app.character, 12345)
        app.monitorChanges(app.character)
        AssertEqual("", app.overlay["DPSText"].Text, "History is still skipped before clock alignment.")
        app.Append(path, DPSLine(serverSecond, 1000) DPSLine(serverSecond, 500, "to"))
        app.monitorChanges(app.character) ; No injected reference time: use the production path.
        meter := app.monitoredChars[app.character]["dps"]
        AssertEqual(100, meter.Rates(meter.LogSecond()).incoming)
        AssertEqual(50, meter.Rates(meter.LogSecond()).outgoing)
        AssertTrue(InStr(app.overlay["DPSText"].Text, "In:") > 0)
        AssertTrue(InStr(app.overlay["DPSText"].Text, "Out:") > 0)
        app.monitorChanges(app.character)
        AssertTrue(InStr(app.overlay["DPSText"].Text, "In:") > 0, "A quiet poll must use the same aligned clock.")
    } finally {
        app.Cleanup()
        DetectHiddenWindows(oldHidden)
    }
}

TestRunner.Register("Live reader tolerates empty reads and preserves unfinished lines", DPSEmptyReadTest)
DPSEmptyReadTest() {
    app := DPSLogFixture()
    try {
        path := app.WriteLog("20260912_000000_12345.txt", "")
        app.startLogMonitoring(app.character, 12345)
        app.ReadGameLogUpdates(app.character, true)
        line := RTrim(DPSLine(DPSMeter.CurrentSecond() + 3, 1000), "`r`n")
        app.Append(path, line)
        app.ReadGameLogUpdates(app.character, true)
        app.ReadGameLogUpdates(app.character, true)
        reader := app.monitoredChars[app.character]
        AssertEqual(line, reader["pending"])
        AssertEqual(0, reader["dps"].buckets.Count)
        app.Append(path, "`r`n")
        app.ReadGameLogUpdates(app.character, true)
        app.ReadGameLogUpdates(app.character, true)
        AssertEqual(100, reader["dps"].Rates(reader["dps"].LogSecond()).incoming)
        AssertEqual("", reader["pending"])
    } finally {
        app.Cleanup()
    }
}

TestRunner.Register("DPS ignores neuts repairs misses notifications and tackle", DPSDamageOnlyTest)
DPSDamageOnlyTest() {
    meter := DPSMeter(DPSSettings())
    now := DateDiff("20260912000005", "19700101000000", "Seconds")
    prefix := "[ 2026.09.12 00:00:05 ] (combat) "
    for payload in ["50 GJ energy neutralized from Pilot - Neutralizer", "Your weapon misses Pilot completely", "Pilot misses you completely", "100 remote shield boosted to Pilot", "Warp scramble attempt from Pilot to Other", "<b>50 GJ</b> energy neutralized from Pilot"]
        meter.AddLine(prefix payload, now)
    meter.AddLine(StrReplace(DPSLine(now, 500), "(combat)", "(notify)"), now)
    AssertEqual(0, meter.buckets.Count)
    meter.AddLine(prefix "150.5 from Pilot - Mjolnir Torpedo - Hits", now)
    AssertEqual(15.05, meter.Rates(now).incoming)
}

TestRunner.Register("DPS thresholds apply to each rate and respect direction switches", DPSThresholdTest)
DPSThresholdTest() {
    settings := DPSSettings()
    settings["incomingDPSThreshold"] := 100
    settings["outgoingDPSThreshold"] := 60
    meter := DPSMeter(settings)
    now := DPSMeter.CurrentSecond()
    meter.AddLine(DPSLine(now, 999), now)
    meter.AddLine(DPSLine(now, 500, "to"), now)
    AssertEqual("", meter.Text(now))
    meter.AddLine(DPSLine(now, 1), now)
    AssertTrue(InStr(meter.Text(now), "In:") > 0, "The threshold is inclusive and does not filter individual hits.")
    AssertFalse(InStr(meter.Text(now), "Out:") > 0)
    settings["incomingDPSEnabled"] := 0
    settings["outgoingDPSThreshold"] := 0
    meter := DPSMeter(settings)
    meter.AddLine(DPSLine(now, 9999), now)
    meter.AddLine(DPSLine(now, 500, "to"), now)
    AssertEqual(0, meter.Rates(now).incoming)
    AssertTrue(InStr(meter.Text(now), "Out:") > 0)
}

TestRunner.Register("Smartbomb families retain mixed damage and the named exception", DPSSmartbombFamilyTest)
DPSSmartbombFamilyTest() {
    settings := DPSSettings()
    settings["showIncomingDPSResistances"] := 1
    meter := DPSMeter(settings)
    now := DPSMeter.CurrentSecond()
    for index, family in ["EMP", "Plasma", "Graviton", "Proton"] {
        expected := [0, 0, 0, 0]
        expected[index] := 1
        for weapon in ["Large " family " Smartbomb II", "Faction Small " family " Smartbomb", "Modified Medium " StrLower(family) " smartbomb"]
            AssertArrayEqual(expected, meter.DamageMix(DPSLine(now, 100, "from", "Pilot", weapon)))
    }
    AssertArrayEqual([0.25, 0.25, 0.25, 0.25], meter.DamageMix(DPSLine(now, 100, "from", "Pilot", "Allecto's Modified Medium Multispectrum Smartbomb")))
    AssertArrayEqual([0, 1, 0, 0], meter.DamageMix(DPSLine(now, 100, "from", "Pilot", "Micro YF-12a Smartbomb")))
    AssertEqual(0, meter.DamageMix(DPSLine(now, 100, "from", "Large EMP Smartbomb I", "Unknown Weapon")))
    AssertEqual(0, meter.DamageMix(DPSLine(now, 100, "from", "Pilot", "EMP S")))
    AssertEqual(0, meter.DamageMix(DPSLine(now, 100, "from", "Pilot", "Unknown Smartbomb")))
}

TestRunner.Register("Missile damage families cover variants without matching attacker names", DPSMissileFamilyTest)
DPSMissileFamilyTest() {
    settings := DPSSettings()
    settings["showIncomingDPSResistances"] := 1
    meter := DPSMeter(settings)
    now := DPSMeter.CurrentSecond()
    for index, family in ["Mjolnir", "Inferno", "Scourge", "Nova"] {
        expected := [0, 0, 0, 0]
        expected[index] := 1
        for weapon in [family, family " Rocket", "Caldari Navy " family " Torpedo", family " Fury Heavy Missile", StrLower(family) " Precision Light Missile"]
            AssertArrayEqual(expected, meter.DamageMix(DPSLine(now, 100, "from", "Pilot", weapon)))
        AssertEqual(0, meter.DamageMix(DPSLine(now, 100, "from", family, "Unknown Weapon")))
        AssertEqual(0, meter.DamageMix(DPSLine(now, 100, "from", "Pilot", "Super" family)))
    }
}

TestRunner.Register("Incoming damage mix combines known ammo drones and unknown damage", DPSDamageMixTest)
DPSDamageMixTest() {
    settings := DPSSettings()
    settings["showIncomingDPSResistances"] := 1
    meter := DPSMeter(settings)
    now := DPSMeter.CurrentSecond()
    meter.AddLine(DPSLine(now, 400, "from", "Gecko", ""), now)
    meter.AddLine(DPSLine(now, 200, "from", "Pilot", "Mjolnir Torpedo"), now)
    meter.AddLine(DPSLine(now, 100, "from", "Hobgoblin II", ""), now)
    meter.AddLine(DPSLine(now, 300, "from", "Fake Gecko[CORP]", "800mm Repeating Cannon II"), now)
    rates := meter.Rates(now)
    AssertArrayEqual([300, 200, 100, 100, 300], rates.types)
    meter.Text(now)
    AssertEqual("% EM 30 Th 20 Ki 10 Ex 10 ? 30", meter.damageText)
    meter.Text(now + 10)
    AssertEqual("", meter.damageText)
    AssertTrue(meter.profiles = DPSMeter(settings).profiles, "Damage profiles are shared across characters.")
    AssertFalse(DPSMeter(DPSSettings()).HasOwnProp("profiles"), "The optional lookup is skipped when disabled.")
    AssertFalse(meter.profiles.Has("Antimatter Charge L"), "Turret ammo is outside damage-type lookup scope.")
    AssertArrayEqual([1, 0, 0, 0], meter.profiles["Templar I"].mix)
    AssertArrayEqual([0, 0, 1, 0], meter.profiles["Dragonfly I"].mix)
    AssertFalse(meter.profiles.Has("Large EMP Smartbomb I"))
    AssertArrayEqual([1, 0, 0, 0], meter.DamageMix(DPSLine(now, 100, "from", "Pilot", "Large EMP Smartbomb I")))
    AssertFalse(meter.profiles.Has("Nova Fury Heavy Missile"), "Missile variants do not need individual profiles.")
    AssertArrayEqual([0, 0, 0, 1], meter.DamageMix(DPSLine(now, 100, "from", "Pilot", "Nova Fury Heavy Missile")))
    AssertArrayEqual([1, 0, 0, 0], meter.DamageMix(DPSLine(now, 100, "from", "Templar I", "")))
    AssertArrayEqual([0, 0, 1, 0], meter.DamageMix(DPSLine(now, 100, "from", "Pilot", "Dragonfly I")))
    for name, item in meter.profiles {
        total := 0
        for fraction in item.mix {
            AssertTrue(fraction >= 0 && fraction <= 1, name)
            total += fraction
        }
        AssertTrue(Abs(total - 1) < 0.000001, name " must have a normalized mix.")
    }
}

class DPSLogFixture extends CombatEventsFixture {
    __New() {
        super.__New()
        this.ProfileThumbnailsBehavior := "Default"
        this.ProfileGameLogsMonitoring := "Default"
        this.EVEExe := "ahk_pid " DllCall("GetCurrentProcessId")
        this.gameLogsMonitoringEnabled := 1
        this.monitorOnlySelectedChars := 1
        this.charsToMonitor := []
        this.flashBorderEnabled := 0
        this.showEventText := 0
        this.supressForFocused := 0
        this.dpsMonitoring := DPSSettings()
        this.debugToolTipText := ""
        this.debugToolTipMethod := (*) => 0
        this.debugToolTipDelay := 60000
        this.Save_Settings_Delay_Timer := (*) => 0
        this.gameLogsMonitoring()
        SetTimer(this.monitorMethod, 0)
        this.directory := A_ScriptDir "\DPS_logs_fixture_" A_TickCount
        DirCreate(this.directory)
        this.gameLogsDirectory := this.directory
        this.files := []
        this.character := "EVE - DPS Test Pilot"
        this.client := Gui(, this.character)
        this.overlay := Gui("-Caption")
        this.overlay.MarginX := 5
        this.overlay.MarginY := 5
        this.overlay.SetFont("s12", "Gill Sans MT")
        this.AddThumbnailTextControls(this.overlay, this.character, 300, 220)
        this.ThumbWindows.%this.client.Hwnd% := Map("Window", {Title: this.character}, "TextOverlay", this.overlay)
        this.charsToMonitor := [this.character]
        this.events := []
    }
    WriteLog(name, content) {
        path := this.directory "\" name
        this.files.Push(path)
        FileAppend("----------------`r`nGamelog`r`nListener: DPS Test Pilot`r`nSession started: 2026.09.12 00:00:00`r`n" content, path, "UTF-8")
        return path
    }
    Append(path, text) {
        file := FileOpen(path, "a", "UTF-8")
        file.Write(text)
        file.Close()
    }
    handleEventActivation(character, *) {
        if this.monitoredChars[character].Get("event", "") != ""
            this.events.Push(this.monitoredChars[character]["event"])
    }
    Cleanup() {
        for character in this.monitoredChars.Clone()
            this.stopLogMonitoring(character)
        SetTimer(this.debugToolTipMethod, 0)
        SetTimer(this.Save_Settings_Delay_Timer, 0)
        SetTimer(this.monitorMethod, 0)
        for character, timer in this.shootingChars
            SetTimer(timer, 0)
        this.client.Destroy()
        this.overlay.Destroy()
        for path in this.files {
            if FileExist(path)
                FileDelete(path)
        }
        DirDelete(this.directory)
    }
}

TestRunner.Register("Live DPS reader shares settings buffers partial writes and handles truncation", DPSLiveReaderTest)
DPSLiveReaderTest() {
    oldHidden := A_DetectHiddenWindows
    DetectHiddenWindows(true)
    app := DPSLogFixture()
    now := DateDiff("20260912000005", "19700101000000", "Seconds")
    try {
        AssertTrue(app.monitoringInitialized, "DPS alone must initialize without event visuals.")
        AssertEqual(0, app.enabledMonitoredEvents.Count)
        AssertFalse(app.checkNPCs)
        path := app.WriteLog("20260912_000000_12345.txt", DPSLine(now, 9999))
        app.startLogMonitoring(app.character, 0)
        AssertTrue(app.monitoredChars.Has(app.character))
        app.monitorChanges(app.character, now)
        AssertEqual("", app.overlay["DPSText"].Text, "Startup must skip history.")
        app.Append(path, DPSLine(now, 1000) DPSLine(now, 500, "to"))
        app.Append(path, RTrim(DPSLine(now + 1, 400), "`r`n"))
        app.monitorChanges(app.character, now + 1)
        meter := app.monitoredChars[app.character]["dps"]
        AssertEqual(100, meter.Rates(now + 1).incoming)
        AssertEqual(50, meter.Rates(now + 1).outgoing)
        AssertTrue(InStr(app.overlay["DPSText"].Text, "In:") > 0)
        app.Append(path, "`r`n")
        app.monitorChanges(app.character, now + 1)
        AssertEqual(140, meter.Rates(now + 1).incoming)
        app.monitorChanges(app.character, now + 1)
        AssertEqual(140, meter.Rates(now + 1).incoming, "An unchanged file must not be counted twice.")
        app.monitorChanges(app.character, now + 11)
        AssertEqual("", app.overlay["DPSText"].Text)
        writer := FileOpen(path, "w", "UTF-8")
        writer.Write(DPSLine(now + 11, 200))
        writer.Close()
        app.monitorChanges(app.character, now + 11)
        AssertEqual(20, app.monitoredChars[app.character]["dps"].Rates(now + 11).incoming)
        app.charsToMonitor := []
        app.monitorChanges(app.character, now + 11)
        AssertFalse(app.monitoredChars.Has(app.character))
        AssertEqual("", app.overlay["DPSText"].Text)
        app.startLogMonitoring(app.character, 12345)
        AssertFalse(app.monitoredChars.Has(app.character), "Discovery must honor the selected character list too.")
    } finally {
        app.Cleanup()
        DetectHiddenWindows(oldHidden)
    }
}

TestRunner.Register("DPS preserves event processing and does not lose suppressed damage", DPSSharedEventReaderTest)
DPSSharedEventReaderTest() {
    app := DPSLogFixture()
    now := DPSMeter.CurrentSecond()
    try {
        path := app.WriteLog("20260912_000001_12345.txt", "")
        app.startLogMonitoring(app.character, 0)
        app.enabledMonitoredEvents["gateJumped"] := Map("needRegex", 0, "pattern", "Jumping from")
        app.Append(path, DPSLine(now, 1000) "[ now ] (notify) Jumping from Jita to Perimeter`r`n")
        app.ReadGameLogUpdates(app.character, false, now)
        AssertEqual("gateJumped", app.events[1])
        app.Append(path, DPSLine(now, 500) "[ now ] (notify) Jumping from Jita to Perimeter`r`n")
        app.ReadGameLogUpdates(app.character, true, now)
        AssertEqual(1, app.events.Length, "Suppressed events must not fire.")
        meter := app.monitoredChars[app.character]["dps"]
        AssertEqual(150, meter.Rates(now).incoming, "Suppression must not discard damage.")
        app.RefreshCharacterDPS(meter, app.client.Hwnd, now, true)
        AssertEqual("", app.overlay["DPSText"].Text)
        app.RefreshCharacterDPS(meter, app.client.Hwnd, now, false)
        AssertTrue(InStr(app.overlay["DPSText"].Text, "In:") > 0)
    } finally {
        app.Cleanup()
    }
}

TestRunner.Register("DPS display preserves names events and systems when resized", DPSLayoutTest)
DPSLayoutTest() {
    app := DPSLogFixture()
    try {
        overlay := app.overlay
        overlay.SystemTrackingEnabled := 1
        overlay["EventText"].Text := "Under attack"
        overlay["SystemText"].Text := "Jita"
        app.updateThumbnailDPSText("DPS: In: 150 | Out: 50", app.client.Hwnd, "% EM 25 Th 25 Ki 25 Ex 25")
        for size in [[300, 220], [250, 140], [100, 70], [8, 8]] {
            app.LayoutThumbnailText(overlay, 0, size[1], size[2])
            lastBottom := 0
            for name in ["OverlayText", "DPSText", "DPSDamageText", "EventText", "SystemText"] {
                overlay[name].GetPos(&x, &y, &w, &h)
                AssertTrue(y >= lastBottom, "Text controls must not overlap.")
                AssertTrue(x >= 0 && y >= 0 && x + w <= size[1] && y + h <= size[2], "Text must stay inside the thumbnail.")
                lastBottom := y + h
            }
        }
        app.updateThumbnailEventText("", app.client.Hwnd)
        AssertTrue(InStr(overlay["DPSText"].Text, "In:") > 0)
        app.SetThumbnailText[app.client.Hwnd] := "EVE - New Pilot"
        AssertEqual("", overlay["DPSText"].Text)
        AssertEqual("", overlay["DPSDamageText"].Text)
    } finally {
        app.Cleanup()
    }
}

TestRunner.Register("Incoming DPS remains visible at the reported compact thumbnail size", DPSCompactIncomingLayoutTest)
DPSCompactIncomingLayoutTest() {
    app := DPSLogFixture()
    overlay := Gui("-Caption")
    try {
        overlay.MarginX := 3
        overlay.MarginY := 3
        overlay.SetFont("s12", "Comic Sans")
        app.systemTrackingEnabled := 1
        app.AddThumbnailTextControls(overlay, "EVE - Myamlyach", 128, 80)
        overlay["DPSText"].Text := "DPS: In: 50"
        overlay["DPSDamageText"].Text := "% Th 61 ? 39"
        overlay["EventText"].Text := "Under Attack By Player"
        overlay["SystemText"].Text := "Jita"
        app.LayoutThumbnailText(overlay, 0, 128, 80)
        overlay["DPSText"].GetPos(,,, &dpsHeight)
        overlay["OverlayText"].GetPos(,,, &nameHeight)
        AssertTrue(dpsHeight > 0, "Incoming DPS hidden: measured name=" app.MeasureThumbnailNameHeight(overlay, 122) ", line=" overlay.EventTextHeight)
        AssertTrue(nameHeight >= app.MeasureThumbnailNameHeight(overlay, 122))
    } finally {
        overlay.Destroy()
        app.Cleanup()
    }
}

TestRunner.Register("DPS settings migrate inherit profiles and fit their panel", DPSSettingsIntegrationTest)
DPSSettingsIntegrationTest() {
    app := DPSLogFixture()
    try {
        defaults := JSON.Load(default_JSON)
        legacy := JSON.Load(default_JSON)
        legacy["_Profiles"]["Default"]["DPS Monitoring"] := Map()
        merged := JsonMergeNoOverwrite(defaults, legacy)
        AssertEqual(0, merged["_Profiles"]["Default"]["DPS Monitoring"]["incomingDPSEnabled"])
        AssertEqual(10, merged["_Profiles"]["Default"]["DPS Monitoring"]["dpsAverageSeconds"])
        app._JSON["_Profiles"]["Other"] := DeepClone(app._JSON["_Profiles"]["Default"])
        app.LastUsedProfile := "Other"
        app.ProfileOverride()
        app.dpsMonitoring["incomingDPSThreshold"] := 123.5
        AssertEqual(0, app._JSON["_Profiles"]["Default"]["DPS Monitoring"]["incomingDPSThreshold"])
        app.Global_Groups["DPS Monitoring"] := 1
        app.ProfileOverride()
        AssertEqual("Default", app.ProfileDPSMonitoring)
        AssertEqual(0, app.dpsMonitoring["incomingDPSThreshold"])
        app.SetState()
        app.MainFrame := Gui()
        app.MainFrame.Group := Map()
        try {
            app.DPSMonitoring_Ctrl()
            for control in app.MainFrame.Group["DPS Monitoring"] {
                control.GetPos(&x, &y, &w, &h)
                AssertTrue(x + w <= app.contentW - 10 && y + h <= app.guiHeight - 10, "DPS settings must fit the panel.")
                if control.Type = "Text" && RegExMatch(control.Text, ":$")
                    AssertTrue(x + w <= app.contentGap + app.offsetX, "Labels must not overlap inputs.")
            }
        } finally {
            app.MainFrame.Destroy()
        }
    } finally {
        app.Cleanup()
    }
}

TestRunner.Register("Game log discovery refreshes when a file is replaced at the same count", DPSLogReplacementTest)
DPSLogReplacementTest() {
    app := DPSLogFixture()
    now := DPSMeter.CurrentSecond()
    try {
        old := app.WriteLog("20260912_000000_12345.txt", "")
        app.startLogMonitoring(app.character, 12345)
        app.stopLogMonitoring(app.character)
        FileDelete(old)
        newer := app.WriteLog("20260912_000010_12345.txt", "")
        app.startLogMonitoring(app.character, 12345)
        AssertEqual(newer, app.monitoredChars[app.character]["fileName"])
        app.Append(newer, DPSLine(now, 300))
        app.ReadGameLogUpdates(app.character, true, now)
        AssertEqual(30, app.monitoredChars[app.character]["dps"].Rates(now).incoming)
    } finally {
        app.Cleanup()
    }
}

class DPSAlertOrderFixture extends DPSLogFixture {
    handleEventActivation(character, *) {
        this.dpsAtAlert := this.overlay["DPSText"].Text
        super.handleEventActivation(character)
    }
}

TestRunner.Register("Incoming DPS is drawn before batched alerts with first and last priority", DPSBeforeAlertsTest)
DPSBeforeAlertsTest() {
    app := DPSAlertOrderFixture()
    now := DPSMeter.CurrentSecond()
    try {
        path := app.WriteLog("20260912_000000_12345.txt", "")
        app.startLogMonitoring(app.character, 12345)
        app.enabledMonitoredEvents := Map("gateJumped", Map("needRegex", 0, "pattern", "Jumping from"), "warpDisrupted", Map("needRegex", 0, "pattern", "you!"))
        for priority in [0, 1] {
            app.lastEventPriority := priority
            app.events := []
            app.updateThumbnailDPSText("", app.client.Hwnd)
            batch := ""
            Loop 10
                batch .= DPSLine(now, 100)
            app.Append(path, batch "[ now ] (notify) Jumping from Jita to Perimeter`r`n[ now ] (combat) Warp scramble attempt to you!`r`n")
            app.ReadGameLogUpdates(app.character, false, now, app.client.Hwnd)
            AssertTrue(InStr(app.dpsAtAlert, "In:") > 0, "DPS must be drawn before any border/event callback runs.")
            AssertEqual(1, app.events.Length, "Only one alert visual update is needed per batch.")
            AssertEqual(priority ? "warpDisrupted" : "gateJumped", app.events[1])
        }
    } finally {
        app.Cleanup()
    }
}

TestRunner.Register("Flashing borders restore the caller's critical state", DPSBorderCriticalTest)
DPSBorderCriticalTest() {
    app := DPSLogFixture()
    previous := A_IsCritical
    try {
        Critical(false)
        app.flashBorder("Missing character")
        AssertEqual(0, A_IsCritical, "A synchronous flash must not leave log polling uninterruptible.")
        Critical(25)
        app.flashBorder("Missing character")
        AssertEqual(25, A_IsCritical, "An already-critical caller must keep its original state.")
    } finally {
        Critical(previous)
        app.Cleanup()
    }
}

TestRunner.Register("Damage percentages have no marker or unused event row below them", DPSPercentageSpacingTest)
DPSPercentageSpacingTest() {
    app := DPSLogFixture()
    settings := DPSSettings()
    settings["showIncomingDPSResistances"] := 1
    meter := DPSMeter(settings)
    now := DPSMeter.CurrentSecond()
    try {
        meter.AddLine(DPSLine(now, 200, "from", "Pilot", "Hammerhead I"), now)
        app.overlay.SystemTrackingEnabled := 1
        app.RefreshCharacterDPS(meter, app.client.Hwnd, now, false)
        AssertFalse(InStr(meter.damageText, "~") > 0)
        AssertFalse(RegExMatch(meter.damageText, "\R$") > 0)
        app.overlay["EventText"].GetPos(,,, &eventHeight)
        app.overlay["DPSDamageText"].GetPos(, &damageY,, &damageHeight)
        AssertEqual(app.overlay.DPSTextHeight, damageHeight, "Percentages must occupy exactly one line.")
        app.overlay["SystemText"].GetPos(, &systemY)
        AssertEqual(0, eventHeight)
        AssertEqual(systemY, damageY + damageHeight, "Percentages must sit directly above the system when no event is active.")
        app.updateThumbnailEventText("Under Attack By Player", app.client.Hwnd)
        app.overlay["EventText"].GetPos(,,, &eventHeight)
        AssertTrue(eventHeight > 0)
        app.updateThumbnailEventText("", app.client.Hwnd)
        app.overlay["EventText"].GetPos(,,, &eventHeight)
        AssertEqual(0, eventHeight, "Ending an event must release its line immediately.")
    } finally {
        app.Cleanup()
    }
}
