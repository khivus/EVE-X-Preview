TestRunner.Register("NPC database matches complete names without player substring collisions", NPCExactNamesTest)
NPCExactNamesTest() {
    app := CombatEventsFixture()
    AssertTrue(app.generalNPCs.Count > 6000, "The full reference dataset must be bundled.")
    for name in ["Angel Warlord", "Blood Apostle", "Annihilator Alvum", "Niarja Myelen",
        "Blastgrip Tessera", "Autothysian Lancer", "Mordus Katana", "Zor", "[AIR] Incursus", Chr(0x2666) " Rifter"] {
        AssertEqual("npc", app.ClassifyTarget(name), name)
        AssertEqual("npc", app.ClassifyTarget(StrLower(name)), "Case-insensitive: " name)
        AssertEqual("player", app.ClassifyTarget(name " Fan"), "Suffix collision: " name)
        AssertEqual("player", app.ClassifyTarget("Pilot " name), "Prefix collision: " name)
        AssertEqual("player", app.ClassifyTarget(name "[CORP](Rifter)"), "Tagged player: " name)
    }
    AssertEqual("npc", app.ClassifyTarget("  Angel Warlord  "))
    for name in ["Guristas", "Navy", "Zorro", "Unknown Capsuleer", "", "Abyssal Overmind"]
        AssertEqual("player", app.ClassifyTarget(name), "A fragment is not an NPC name: " name)
}

TestRunner.Register("Special NPC maps retain priority and fall back to general classification", NPCSpecialPriorityTest)
NPCSpecialPriorityTest() {
    app := CombatEventsFixture()
    for item in [["faction", "Domination Warlord", "checkFactionNPCs"],
        ["officer", "Estamel Tharchon", "checkOfficerNPCs"],
        ["capital", "Guristas Dreadnought", "checkCapitalNPCs"]] {
        AssertEqual("npc", app.ClassifyTarget(item[2]))
        app.%item[3]% := true
        AssertEqual(item[1], app.ClassifyTarget(item[2]))
        AssertEqual("player", app.ClassifyTarget(item[2] "[CORP]"))
        app.checkGeneralNPCs := false
        AssertEqual(item[1], app.ClassifyTarget(item[2]))
        app.%item[3]% := false
        AssertEqual("player", app.ClassifyTarget(item[2]))
        app.checkGeneralNPCs := true
    }
    AssertEqual(60, app.factionNPCs.Count)
    AssertEqual(24, app.officerNPCs.Count)
    AssertEqual(18, app.capitalNPCs.Count)
}

TestRunner.Register("NPC maps are shared across monitoring instances", NPCSharedMapsTest)
NPCSharedMapsTest() {
    first := CombatEventsFixture()
    second := CombatEventsFixture()
    for property in ["generalNPCs", "factionNPCs", "officerNPCs", "capitalNPCs"]
        AssertEqual(ObjPtr(first.%property%), ObjPtr(second.%property%))
}

TestRunner.Register("Exact NPC names classify damage and missed shots in combat logs", NPCCombatNamesTest)
NPCCombatNamesTest() {
    app := CombatEventsFixture()
    AssertEqual("underAttackByNPC", app.Check("22:44:30 Combat 209 from Angel Warlord - 1400mm Artillery - Hits"))
    AssertEqual("underAttackByPlayer", app.Check("22:44:30 Combat 209 from Angel Warlord Fan - 1400mm Artillery - Hits"))
    AssertEqual("underAttackByNPC", app.Check('[ 2026.09.11 20:00:00 ] (combat) <b>Angel Warlord</b> misses you completely'))
    AssertEqual("underAttackByPlayer", app.Check('[ 2026.09.11 20:00:00 ] (combat) <b>Angel Warlord Fan</b> misses you completely'))
    app.checkOfficerNPCs := true
    AssertEqual("engagedWithOfficerNPC", app.Check('[ 2026.09.11 20:00:00 ] (combat) Your Artillery misses <b>Estamel Tharchon</b> completely'))
}
