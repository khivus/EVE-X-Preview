; Shared, lazily initialized NPC name maps. No file or network I/O during combat.
#Include "data/NPCGeneralNames.ahk"

class NPCDatabase {
    static Faction() {
        static names := NPCDatabase.ToMap([ ; NPCs to trigger engagedWithFactionBSNPC event
            "Domination Cherubim",
            "Domination Commander",
            "Domination General",
            "Domination Malakim",
            "Domination Nephilim",
            "Domination Saint",
            "Domination Seraphim",
            "Domination Throne",
            "Domination War General",
            "Domination Warlord",
            "Dark Blood Apostle",
            "Dark Blood Archbishop",
            "Dark Blood Archon",
            "Dark Blood Cardinal",
            "Dark Blood Harbinger",
            "Dark Blood Monsignor",
            "Dark Blood Oracle",
            "Dark Blood Patriarch",
            "Dark Blood Pope",
            "Dark Blood Prophet",
            "Dread Guristas Conquistador",
            "Dread Guristas Destroyer",
            "Dread Guristas Dismantler",
            "Dread Guristas Eliminator",
            "Dread Guristas Eradicator",
            "Dread Guristas Exterminator",
            "Dread Guristas Extinguisher",
            "Dread Guristas Massacrer",
            "Dread Guristas Obliterator",
            "Dread Guristas Usurper",
            "Sentient Alvus Controller",
            "Sentient Alvus Creator",
            "Sentient Alvus Queen",
            "Sentient Alvus Ruler",
            "Sentient Domination Alvus",
            "Sentient Matriarch Alvus",
            "Sentient Patriarch Alvus",
            "Sentient Spearhead Alvus",
            "Sentient Supreme Alvus Parasite",
            "Sentient Swarm Preserver Alvus",
            "True Sansha's Beast Lord",
            "True Sansha's Dark Lord",
            "True Sansha's Dread Lord",
            "True Sansha's Lord",
            "True Sansha's Mutant Lord",
            "True Sansha's Overlord",
            "True Sansha's Plague Lord",
            "True Sansha's Savage Lord",
            "True Sansha's Slave Lord",
            "True Sansha's Tyrant",
            "Shadow Serpentis Admiral",
            "Shadow Serpentis Baron",
            "Shadow Serpentis Commodore",
            "Shadow Serpentis Flotilla Admiral",
            "Shadow Serpentis Grand Admiral",
            "Shadow Serpentis High Admiral",
            "Shadow Serpentis Lord Admiral",
            "Shadow Serpentis Port Admiral",
            "Shadow Serpentis Rear Admiral",
            "Shadow Serpentis Vice Admiral"
        ])
        return names
    }

    static Officer() {
        static names := NPCDatabase.ToMap([ ; NPCs to trigger engagedWithOfficerNPC event
            "Gotan Kreiss",
            "Hakim Stormare",
            "Mizuro Cybon",
            "Tobias Kruzhor",
            "Ahremen Arkah",
            "Draclira Merlonne",
            "Raysere Giant",
            "Tairei Namazoth",
            "Estamel Tharchon",
            "Kaikka Peunato",
            "Thon Eney",
            "Vepas Minimala",
            "Unit D-34343",
            "Unit F-435454",
            "Unit P-343554",
            "Unit W-634",
            "Brokara Ryver",
            "Chelm Soran",
            "Selynne Mardakar",
            "Vizan Ankonin",
            "Brynn Jerdola",
            "Cormack Vaaja",
            "Setele Schellan",
            "Tuvan Orth"
        ])
        return names
    }

    static Capital() {
        static names := NPCDatabase.ToMap([ ; NPCs to trigger engagedWithCapitalNPC event
            "Domination Titan",
            "Dark Blood Titan",
            "Shadow Serpentis Titan",
            "Angel Dreadnought",
            "Domination Dreadnought",
            "Blood Dreadnought",
            "Dark Blood Dreadnought",
            "Dread Guristas Dreadnought",
            "Guristas Dreadnought",
            "Sansha's Dreadnought",
            "True Sansha's Dreadnought",
            "Serpentis Dreadnought",
            "Shadow Serpentis Dreadnought",
            "Infested Carrier",
            "Sentient Infested Carrier",
            "Sentient Infested Supercarrier",
            "True Sansha's Supercarrier",
            "Dread Guristas Titan"
        ])
        return names
    }

    static General() {
        static names := NPCDatabase.BuildGeneral()
        return names
    }

    static BuildGeneral() {
        names := Map()
        names.CaseSense := "Off"
        NPCGeneralNames.AddTo(names)
        ; FOB logs use faction + type name instead of the static diamond name.
        ; Build complete aliases once, retaining exact matching during combat.
        aliases := []
        for name in names {
            if SubStr(name, 1, 2) != Chr(0x2666) " "
                continue
            for prefix in ["Blood Raiders ", "Guristas "]
                aliases.Push(prefix SubStr(name, 3))
        }
        for name in aliases
            names[name] := true
        ; Preserve recognition when a special event is disabled, even if a name
        ; from the curated lists is absent from a future reference snapshot.
        for special in [this.Faction(), this.Officer(), this.Capital()]
            for name in special
                names[name] := true
        ; Display-name alias observed in combat logs (see Combat_Events_Test).
        names["Guristas Pirates Stronghold"] := true
        return names
    }

    static ToMap(list) {
        names := Map()
        for name in list
            names[name] := true
        return names
    }
}
