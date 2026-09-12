class CombatEventsFixture extends ThumbnailNamesFixture {
    __New() {
        super.__New()
        this._JSON := JSON.Load(default_JSON)
        this.ProfileMonitoredEvents := "Default"
        this.checkNPCs := true
        this.checkGeneralNPCs := true
        this.checkFactionNPCs := false
        this.checkOfficerNPCs := false
        this.checkCapitalNPCs := false
        this.generalNPCs := NPCDatabase.General()
        this.factionNPCs := NPCDatabase.Faction()
        this.officerNPCs := NPCDatabase.Officer()
        this.capitalNPCs := NPCDatabase.Capital()
        this.playerEngagmentEnabled := true
        this.anyNPCEngagmentEnabled := true
        this.monitoredChars := Map("Pilot", Map("event", ""))
    }

    Check(line) {
        this.monitoredChars["Pilot"]["event"] := ""
        this.processNPCCheck("Pilot", line)
        return this.monitoredChars["Pilot"]["event"]
    }
}

TestRunner.Register("Neutralization uses independent player and NPC options", CombatNeutralizationTest)
CombatNeutralizationTest() {
    app := CombatEventsFixture()
    player := '[ 2026.09.08 20:48:31 ] (combat) <color=0xffe57f7f><b>50 GJ</b><color=0x77ffffff><font size=10> energy neutralized </font><b><color=0xffffffff><color=0xFF40FFFF><b>Drekavac</b></color> <color=0xFF40FF40><b>Di9i9</b></color> <color=0xFFFFFF40>[B0MJ]</color></b><color=0x77ffffff><font size=10> - Small Gremlin Compact Energy Neutralizer</font>'
    npc := '[ 2026.09.08 20:09:16 ] (combat) <color=0xffe57f7f><b>285 GJ</b><color=0x77ffffff><font size=10> energy neutralized </font><b><color=0xffffffff><color=0xFF40FFFF><b>Guristas Pirates Stronghold</b></color> <color=0xFF40FF40><b>Guristas Pirates Stronghold</b></color> </b><color=0x77ffffff><font size=10> - Standup Heavy Energy Neutralizer I</font>'
    AssertEqual("", app.Check(player))
    AssertEqual("", app.Check(npc))
    app.monitoredEvents["underAttackByPlayer"]["includeNeutralization"] := 1
    AssertEqual("underAttackByPlayer", app.Check(player))
    AssertEqual("", app.Check(npc))
    app.monitoredEvents["underAttackByPlayer"]["includeNeutralization"] := 0
    app.monitoredEvents["underAttackByNPC"]["includeNeutralization"] := 1
    AssertEqual("", app.Check(player))
    AssertEqual("underAttackByNPC", app.Check(npc))
    app.anyNPCEngagmentEnabled := false
    AssertEqual("", app.Check(npc), "Neuts must not enable a disabled parent event.")
    app.monitoredEvents["underAttackByPlayer"]["includeNeutralization"] := 1
    app.playerEngagmentEnabled := false
    AssertEqual("", app.Check(player))
    app.playerEngagmentEnabled := true
    AssertEqual("", app.Check(StrReplace(player, "0xffe57f7f", "0xff7fffff")), "Outgoing neutralization must not trigger an attack event.")
    AssertEqual("underAttackByPlayer", app.Check(StrReplace(player, "0xffe57f7f", "0xFFE57F7F")))
    AssertEqual("", app.Check(StrReplace(player, "<color=0xffe57f7f>", "")), "Missing amount color must not be guessed as incoming.")
    AssertEqual("", app.Check(StrReplace(player, "0xffe57f7f", "0xffffffff")))
    app.anyNPCEngagmentEnabled := true
    AssertEqual("", app.Check(StrReplace(npc, "0xffe57f7f", "0xff7fffff")))
    AssertEqual("underAttackByNPC", app.Check(npc))
    app.monitoredChars["Pilot"]["event"] := "warpDisrupted"
    app.processNPCCheck("Pilot", player)
    AssertEqual("warpDisrupted", app.monitoredChars["Pilot"]["event"], "An existing event keeps its priority.")
}

TestRunner.Register("Player smartbombs are ignored while ordinary attacks still trigger", CombatSmartbombTest)
CombatSmartbombTest() {
    app := CombatEventsFixture()
    player := '22:44:30    Combat    209 from Excavator Lambda[WH40K](Proteus) - Caldari Navy Large Graviton Smartbomb - Hits'
    AssertEqual("", app.Check(player))
    AssertEqual("", app.Check(StrReplace(player, "Graviton Smartbomb", "Proton Smartbomb")))
    AssertEqual("underAttackByPlayer", app.Check(StrReplace(player, "Caldari Navy Large Graviton Smartbomb", "Heavy Pulse Laser II")))
    raw := '[ 2026.09.08 22:44:30 ] (combat) <color=0xffcc0000><b>209</b> <font size=10>from</font> <b><color=0xffffffff>Excavator Delta[WH40K](Proteus)</b> - Large Plasma Smartbomb - Hits'
    AssertEqual("", app.Check(raw))
    AssertEqual("underAttackByPlayer", app.Check(StrReplace(raw, "Large Plasma Smartbomb", "Heavy Pulse Laser II")))
    AssertEqual("", app.Check(StrReplace(raw, ">from<", ">to<")))
    AssertEqual("underAttackByNPC", app.Check(StrReplace(raw, "Excavator Delta[WH40K](Proteus)", "Guristas Pirates Stronghold")), "Smartbomb filtering applies only to players.")
    app.monitoredEvents["underAttackByPlayer"]["ignoreSmartbombDamage"] := 0
    AssertEqual("underAttackByPlayer", app.Check(player))
    AssertEqual("underAttackByPlayer", app.Check(raw))
    AssertEqual("", app.Check(StrReplace(raw, ">from<", ">to<")))
    app.monitoredEvents["underAttackByPlayer"]["ignoreSmartbombDamage"] := 1
    AssertEqual("", app.Check(raw))
    app.monitoredEvents["underAttackByPlayer"].Delete("ignoreSmartbombDamage")
    AssertEqual("", app.Check(raw), "Old settings must default to ignoring smartbomb damage.")
}

TestRunner.Register("Neutralization options survive settings merge and JSON roundtrip", CombatOptionsPersistenceTest)
CombatOptionsPersistenceTest() {
    app := CombatEventsFixture()
    app.monitoredEvents["underAttackByPlayer"]["includeNeutralization"] := 1
    app.monitoredEvents["underAttackByPlayer"]["ignoreSmartbombDamage"] := 0
    app.monitoredEvents["underAttackByNPC"].Delete("includeNeutralization")
    app._JSON := JsonMergeNoOverwrite(JSON.Load(default_JSON), JSON.Load(JSON.Dump(app._JSON)))
    AssertEqual(1, app.monitoredEvents["underAttackByPlayer"]["includeNeutralization"])
    AssertEqual(0, app.monitoredEvents["underAttackByNPC"]["includeNeutralization"])
    AssertEqual(0, app.monitoredEvents["underAttackByPlayer"]["ignoreSmartbombDamage"])
    app.monitoredEvents["underAttackByPlayer"].Delete("ignoreSmartbombDamage")
    app._JSON := JsonMergeNoOverwrite(JSON.Load(default_JSON), app._JSON)
    AssertEqual(1, app.monitoredEvents["underAttackByPlayer"]["ignoreSmartbombDamage"])
}

TestRunner.Register("Neutralization direction follows the amount color in paired logs", CombatNeutralizationDirectionTest)
CombatNeutralizationDirectionTest() {
    app := CombatEventsFixture()
    app.monitoredEvents["underAttackByPlayer"]["includeNeutralization"] := 1
    outgoing := '[ 2026.09.08 23:13:54 ] (combat) <color=0xff7fffff><b>180 GJ</b><color=0x77ffffff><font size=10> energy neutralized </font><b><color=0xffffffff><color=0xFF40FFFF><b>Damnation</b></color> <color=0xFF40FF40><b>Myamlyach</b></color> <color=0xFFFFFF40>[5TEAK]</color><color=0xFFFFFF40>[B B C]</color></b><color=0x77ffffff><font size=10> - Medium Energy Neutralizer II</font>'
    incoming := '[ 2026.09.08 23:13:54 ] (combat) <color=0xffe57f7f><b>180 GJ</b><color=0x77ffffff><font size=10> energy neutralized </font><b><color=0xffffffff><color=0xFF40FFFF><b>Brutix Navy Issue</b></color> <color=0xFF40FF40><b>Khivus</b></color> <color=0xFFFFFF40>[5TEAK]</color><color=0xFFFFFF40>[B B C]</color></b><color=0x77ffffff><font size=10> - Medium Energy Neutralizer II</font>'
    AssertEqual("", app.Check(outgoing))
    AssertEqual("underAttackByPlayer", app.Check(incoming))
    AssertEqual("", app.Check(StrReplace(outgoing, "<color=0x77ffffff>", "<color=0xffe57f7f>")), "Red elsewhere in the message is not an incoming amount.")
}

TestRunner.Register("Monitoring initializes with only attack events enabled", CombatOnlyMonitoringTest)
CombatOnlyMonitoringTest() {
    for attackEvent in ["underAttackByPlayer", "underAttackByNPC"] {
        app := CombatEventsFixture()
        app.ProfileGameLogsMonitoring := "Default"
        app.gameLogsDirectory := A_ScriptDir
        app.gameLogsMonitoringEnabled := 1
        app.showEventText := 1
        app.monitorOnlySelectedChars := 1
        app.charsToMonitor := []
        app.debugToolTipText := ""
        app.debugToolTipMethod := (*) => 0
        app.debugToolTipDelay := 60000
        app.monitoringInitialized := 0
        for event, settings in app.monitoredEvents
            settings["enabled"] := event = attackEvent
        try {
            app.gameLogsMonitoring()
            AssertTrue(app.monitoringInitialized, "An attack event alone must start monitoring.")
            AssertTrue(app.checkNPCs)
        } finally {
            SetTimer(app.debugToolTipMethod, 0)
            if app.HasOwnProp("monitorMethod")
                SetTimer(app.monitorMethod, 0)
        }
    }
}

TestRunner.Register("Neuts checkboxes fit beside the attack labels", CombatOptionsLayoutTest)
CombatOptionsLayoutTest() {
    app := CombatEventsFixture()
    app.SetState()
    app.MainFrame := Gui()
    app.MainFrame.Group := Map()
    try {
        app.MonitoredEvents_Ctrl()
        for event in ["underAttackByPlayer", "underAttackByNPC"] {
            app.MainFrame["N" event].GetPos(&nx, &ny, &nw, &nh)
            app.MainFrame["E" event].GetPos(&ex, &ey)
            AssertTrue(nx + nw <= ex && ny = ey, "Neuts must fit before the main event checkbox.")
            for control in app.MainFrame.Group["Monitored Events"] {
                if control.Text = app.monitoredEventsTexts[event] ":" {
                    control.GetPos(&lx, &ly, &lw)
                    AssertTrue(lx + lw <= nx && ly = ny, "Neuts must sit beside, without overlapping, its event label.")
                }
            }
        }
        for control in app.MainFrame.Group["Monitored Events"] {
            control.GetPos(&x, &y, &w, &h)
            AssertTrue(x + w <= app.contentW - 10 && y + h <= app.guiHeight - 10, "Event controls must fit the panel.")
        }
    } finally {
        app.MainFrame.Destroy()
    }
}
