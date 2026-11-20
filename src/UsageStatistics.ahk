class UsageStatistics {
    static DefaultDirectory => EnvGet("LOCALAPPDATA") "\EVE-X-Preview\Statistics"
    static Counters := ["switchesRequested", "switchesConfirmed", "characterLaunches", "alertsTriggered", "maxCombo",
        "monitoredEvents", "monitoredDamageIn", "monitoredDamageOut", "monitoredSystemChanges"]
    ; Positions match the worker's archive summary, which contains counts only.
    static ArchiveGroups := [
        ["Archive coverage", 1, "Game logs indexed", 2, "Chat logs indexed"],
        ["Combat archive", 3, "Combat records", 4, "Damage taken", 5, "Damage dealt", 6, "Miss records"],
        ["Travel archive", 7, "Gate jump records", 8, "Local system change records", 10, "Undock records", 13, "Conduit jump records"],
        ["Activity archive", 9, "Chat message records", 11, "Fleet invitation records", 12, "Mining notification records"]]
    lastSwitch := -1
    combo := 0
    job := 0
    viewJob := 0
    viewRevision := ""
    progress := 0
    scansPaused := false
    status := ""
    archive := []

    static ForOwner(owner, directory := UsageStatistics.DefaultDirectory) {
        if owner.HasProp("usageStatistics") && owner.usageStatistics
            return owner.usageStatistics
        if !owner.HasOwnProp("statisticsBrowser")
            owner.statisticsBrowser := UsageStatistics(owner, directory)
        return owner.statisticsBrowser
    }

    __New(owner, directory := UsageStatistics.DefaultDirectory, live := false) {
        this.owner := owner
        this.allEvents := live && owner._JSON.Get("CollectAllStatisticsEvents", false)
        this.directory := directory
        this.path := directory "\totals.json"
        this.saved := Map()
        this.characterCounts := Map()
        this.savedCharacterCounts := Map()
        for key in UsageStatistics.Counters
            this.%key% := 0, this.saved[key] := 0
        this.totals := UsageStatistics.ReadTotals(this.path)
        this.archive := UsageStatistics.ReadArchive(this.directory)
        this.saveTimer := ObjBindMethod(this, "Flush")
        this.pollTimer := ObjBindMethod(this, "Poll")
        this.scanTimer := ObjBindMethod(this, "AutomaticScan")
        this.views := Map()
        this.viewFileInfo := Buffer(36, 0)
        this.exitHandler := ObjBindMethod(this, "Close")
        if live
            SetTimer(this.saveTimer, 60000, -10)
        OnExit(this.exitHandler)
        this.liveCutoff := A_Now
        this.liveSince := A_NowUTC
        if this.allEvents
            SetTimer(this.scanTimer, -60000, -10)
    }

    ConfirmSwitch(holdDelay, tick := A_TickCount, character := "") {
        this.switchesConfirmed += 1
        this.combo := this.lastSwitch < 0 || tick - this.lastSwitch > 5 * holdDelay ? 1 : this.combo + 1
        this.lastSwitch := tick
        this.maxCombo := Max(this.maxCombo, this.combo)
        this.Attribute("switchesConfirmed", character)
        this.Attribute("maxCombo", character, this.combo)
    }

    Attribute(key, character, amount := 1) {
        character := RegExReplace(Trim(character), "^EVE - ")
        if character = "" || character = "EVE" || character = "Char Screen"
            return
        if !this.characterCounts.Has(character)
            this.characterCounts[character] := Map()
        counts := this.characterCounts[character]
        counts[key] := key = "maxCombo" ? Max(counts.Get(key, 0), amount) : counts.Get(key, 0) + amount
    }

    CharacterLeaders() {
        counts := Map(), bests := Map(), leaders := Map()
        for character, values in this.totals.Get("characters", Map())
            counts[character] := values.Clone()
        for character, values in this.characterCounts {
            if !counts.Has(character)
                counts[character] := Map()
            saved := this.savedCharacterCounts.Get(character, Map())
            for key, value in values
                counts[character][key] := key = "maxCombo" ? Max(counts[character].Get(key, 0), value)
                    : counts[character].Get(key, 0) + value - saved.Get(key, 0)
        }
        for character, values in counts
            for key, value in values
                if value > bests.Get(key, 0)
                    bests[key] := value, leaders[key] := character
        return leaders
    }

    static ReadTotals(path) {
        values := Map()
        try values := JSON.Load(FileRead(path, "UTF-8"))
        return values is Map ? values : Map()
    }

    Combined(base, current, key) {
        savedValue := base.Get(key, 0)
        return key = "maxCombo" ? Max(savedValue, current) : savedValue + current - this.saved[key]
    }

    Flush(*) {
        current := Map(), dirty := false
        for key in UsageStatistics.Counters {
            current[key] := this.%key%
            dirty := dirty || current[key] != this.saved[key]
        }
        for character, values in this.characterCounts
            for key, value in values
                dirty := dirty || value != this.savedCharacterCounts.Get(character, Map()).Get(key, 0)
        if !dirty
            return
        mutex := DllCall("CreateMutexW", "Ptr", 0, "Int", 0, "Str", "Local\EVE-X-Preview.StatisticsTotals", "Ptr")
        if !mutex
            return
        locked := false
        try {
            locked := DllCall("WaitForSingleObject", "Ptr", mutex, "UInt", 0, "UInt") <= 0x80
            if !locked
                return
            DirCreate(this.directory)
            totals := UsageStatistics.ReadTotals(this.path)
            for key, value in current
                totals[key] := this.Combined(totals, value, key)
            characters := totals.Get("characters", Map())
            for character, values in this.characterCounts {
                if !characters.Has(character)
                    characters[character] := Map()
                saved := this.savedCharacterCounts.Get(character, Map())
                for key, value in values
                    characters[character][key] := key = "maxCombo" ? Max(characters[character].Get(key, 0), value)
                        : characters[character].Get(key, 0) + value - saved.Get(key, 0)
            }
            totals["characters"] := characters
            temporary := this.path "." DllCall("GetCurrentProcessId") ".tmp"
            if FileExist(temporary)
                FileDelete(temporary)
            FileAppend(JSON.Dump(totals, , "  "), temporary, "UTF-8-RAW")
            FileMove(temporary, this.path, true)
            this.totals := totals
            this.saved := current
            for character, values in this.characterCounts
                this.savedCharacterCounts[character] := values.Clone()
        } catch as err {
            ProgramLog.Add("Statistics save failed: " err.Message, "WARN")
        } finally {
            if locked
                DllCall("ReleaseMutex", "Ptr", mutex)
            DllCall("CloseHandle", "Ptr", mutex)
        }
    }

    static ReadArchive(directory) {
        try {
            values := []
            for value in StrSplit(Trim(FileRead(directory "\archive-summary.tsv", "UTF-8"), "`r`n"), "`t")
                values.Push(Number(value))
            return values
        } catch
            return []
    }

    AutomaticScan() {
        if this.scansPaused
            return
        this.StartScan(true)
        if !this.scansPaused
            SetTimer(this.scanTimer, -60000, -10)
    }

    StartScan(liveOnly := false, resolver?) {
        if this.job && this.job.busy
            return
        if !liveOnly {
            this.scansPaused := false
            if this.allEvents
                SetTimer(this.scanTimer, -60000, -10)
        }
        resolver := resolver ?? this.owner
        try {
            game := resolver.ResolveGameLogsDirectory()
            chat := resolver.ResolveChatLogsDirectory()
            if (!game || !DirExist(game)) && (!chat || !DirExist(chat)) {
                this.status := "No EVE log folders found."
                return
            }
            if !this.job
                this.job := LogScanJob("", 0)
            if this.viewJob && this.viewJob.busy && this.runningRevision = ""
                this.viewJob.StopProcess()
            request := (liveOnly ? "statistics-live" : "statistics") "`n" game "`n" chat "`n" this.directory
            if liveOnly
                request .= "`n" this.liveCutoff "`n" this.liveSince
            if this.job.Start(request) {
                this.liveScan := liveOnly
                this.scanStartedAt := A_Now
                this.status := "Starting the statistics worker…"
                this.progress := 0
                ProgramLog.Add("Statistics scan requested; scope=" (liveOnly ? "live" : "archive"))
                SetTimer(this.pollTimer, 1000, -10)
            } else
                this.status := "Could not start the archive scan. Try again shortly."
        } catch as err {
            this.status := "Could not start the archive scan. Try again shortly."
            ProgramLog.Add(err.Message, "WARN")
        }
    }

    StopScan() {
        this.scansPaused := true
        SetTimer(this.scanTimer, 0)
        if this.job && this.job.busy {
            try {
                if !FileExist(this.job.directory "\stop.txt")
                    FileAppend("", this.job.directory "\stop.txt", "UTF-8-RAW")
                this.status := "Stopping after the current batch…"
                ProgramLog.Add("Statistics scan stop requested")
            } catch as err {
                ProgramLog.Add("Statistics scan stop failed: " err.Message, "WARN")
            }
        } else {
            this.status := "Automatic log scans paused."
            ProgramLog.Add("Automatic statistics scans paused")
        }
    }

    Poll(*) {
        result := this.job.Poll()
        if !result
            return
        this.archive := UsageStatistics.ReadArchive(this.directory)
        lines := result["ok"] ? StrSplit(Trim(result["text"], "`r`n"), "`n", "`r") : []
        if lines.Length && this.status != "Stopping after the current batch…"
            this.status := lines[1]
        if lines.Length > 1 {
            fields := StrSplit(lines[2], "`t")
            if fields.Length >= 4 {
                finished := Number(fields[1]), total := Number(fields[2]), skipped := Number(fields[4])
                eta := finished > 0 ? Round(Max(0, total - finished) * Number(fields[3]) / finished / 1000) : -1
                    this.progress := Map("scanned", Max(0, finished - skipped), "total", total, "skipped", skipped, "eta", eta,
                        "phaseCount", fields.Length > 4 ? Number(fields[5]) : 0)
            }
        }
        if !this.job.busy {
            SetTimer(this.pollTimer, 0)
            this.status := lines.Length ? lines[1] : "Archive scan failed. Try starting it again."
            if this.liveScan && this.status = "Archive scan complete."
                this.liveCutoff := this.scanStartedAt
            ProgramLog.Add("Statistics scan ended; status=" this.status)
        }
    }

    FilteredView(character, range, refresh := false) {
        fileInfo := this.viewFileInfo
        available := DllCall("GetFileAttributesExW", "Str", this.directory "\dashboard.tsv", "Int", 0, "Ptr", fileInfo)
        revision := available ? String(NumGet(fileInfo, 20, "Int64")) : ""
        if revision != this.viewRevision {
            this.views.Clear()
            this.viewRevision := revision
        }
        key := character "`t" range (range = "all" || range = "session" ? "" : "`t" SubStr(A_NowUTC, 1, 8))
        if this.viewJob && this.viewJob.busy {
            result := this.viewJob.Poll()
            if !this.viewJob.busy && this.runningRevision = revision {
                if this.views.Count >= 12
                    this.views.Clear()
                this.views[this.runningView] := result && result["ok"] ? this.ParseView(result["text"]) : Map("error", "Could not read the statistics index.")
                if this.runningView = key
                    refresh := false
            }
        }
        if this.views.Has(key) && !refresh
            return this.views[key]
        if !available && this.job && this.job.busy
            return Map("deferred", true)
        if !this.viewJob
            this.viewJob := LogScanJob("", 120000)
        if !this.viewJob.busy {
            this.runningView := key
            this.runningRevision := revision
            if !this.viewJob.Start("statistics-view`n" this.directory "`n" character "`n" range)
                return Map("error", "Could not read the statistics index.")
        }
        return Map("pending", true)
    }

    ParseView(text) {
        result := Map("characters", [], "counts", [], "materials", [], "peaks", [0, 0], "alpha", [0, 0], "extraResidue", 0, "extra", Map(), "modules", [], "drones", [], "bursts", [], "systems", [], "leaders", Map())
        for line in StrSplit(text, "`n", "`r") {
            fields := StrSplit(line, "`t")
            if !fields.Length
                continue
            switch fields[1] {
                case "D":
                    loop 13
                        result["counts"].Push(Number(fields[A_Index + 1]))
                    result["peaks"] := [Number(fields[15]), Number(fields[16])]
                    result["extraResidue"] := Number(fields[17])
                    if fields.Length >= 20 {
                        result["alpha"] := [Number(fields[18]), Number(fields[19])]
                        result["combatNeedsRescan"] := !Number(fields[20])
                    }
                case "E": result["extendedNeedsRescan"] := !Number(fields[2])
                case "X": result["extra"][fields[2]] := Number(fields[3])
                case "U", "Y", "R", "B": result[fields[1] = "U" ? "modules" : fields[1] = "R" ? "drones" : fields[1] = "B" ? "bursts" : "systems"].Push(Map("name", fields[2], "count", Number(fields[3]), "character", fields[4]))
                case "T": result["leaders"][fields[2]] := fields[3]
                case "P": result["characters"].Push(fields[2])
                case "M": result["materials"].Push(Map("name", fields[2], "units", Number(fields[3]), "residue", Number(fields[4])))
            }
        }
        return result
    }

    static ProgressText(progress) {
        if !progress
            return ""
        if !progress["total"]
            return progress["phaseCount"] ? T("{1} entries checked.", progress["phaseCount"]) : ""
        scanned := progress["scanned"], total := progress["total"], percent := Floor(100 * scanned / total)
        text := T("{1} of {2} files scanned ({3}%). {4} remaining.", scanned, total, percent, Max(0, total - scanned))
        if progress["eta"] > 0 {
            eta := progress["eta"]
            text .= " " T("Estimated time remaining: {1} min {2} sec.", eta // 60, Mod(eta, 60))
        }
        if progress["skipped"]
            text .= " " T("Unreadable files: {1}.", progress["skipped"])
        return text
    }

    Close(*) {
        SetTimer(this.saveTimer, 0)
        SetTimer(this.scanTimer, 0)
        SetTimer(this.pollTimer, 0)
        if this.job
            this.job.Close()
        if this.viewJob
            this.viewJob.Close()
        this.Flush()
        OnExit(this.exitHandler, 0)
    }

    static Snapshot(owner, directory := UsageStatistics.DefaultDirectory, filters := 0) {
        categories := []
        if owner.HasProp("_JSON") && owner._JSON.Has("_Profiles")
            categories.Push(this.Category("Saved configuration", "Unique saved characters", this.SavedCharacters(owner._JSON)))
        live := owner.HasProp("usageStatistics") ? owner.usageStatistics : 0
        stats := filters ? this.ForOwner(owner, directory) : live
        totals := stats ? stats.totals : this.ReadTotals(directory "\totals.json")
        counts := stats ? stats.archive : this.ReadArchive(directory)
        view := Map()
        if stats && filters {
            view := stats.FilteredView(filters.Get("character", ""), filters.Get("range", "all"), filters.Get("refresh", false))
            counts := view.Get("counts", [])
        }
        if live || totals.Count {
            total := (key) => stats ? stats.Combined(totals, stats.%key%, key) : totals.Get(key, 0)
            categories.Push(this.Category("Client switching", "Switch requests", total("switchesRequested"),
                "Confirmed focus changes", total("switchesConfirmed"), "Longest switching combo", total("maxCombo")))
            categories.Push(this.Category("Thumbnails", "Character launches (thumbnails created)", total("characterLaunches")))
            categories.Push(this.Category("Application alerts", "Alerts triggered", total("alertsTriggered")))
            if (live && !live.allEvents) || total("monitoredEvents") || total("monitoredDamageIn")
                || total("monitoredDamageOut") || total("monitoredSystemChanges")
                categories.Push(this.Category("Live monitoring", "Enabled events observed", total("monitoredEvents"),
                    "Monitored damage taken", total("monitoredDamageIn"), "Monitored damage dealt", total("monitoredDamageOut"),
                    "Monitored local system changes", total("monitoredSystemChanges")))
        }
        if counts.Length = 13 {
            for group in this.ArchiveGroups {
                archive := []
                loop (group.Length - 1) // 2
                    archive.Push(group[2 * A_Index + 1], counts[group[2 * A_Index]])
                categories.Push(this.Category(group[1], archive*))
            }
        }
        if view.Has("peaks") {
            extra := view["extra"]
            categories.Push(this.Category("Logistics", "Armor repaired (given)", extra.Get("armorGiven", 0),
                "Armor repaired (received)", extra.Get("armorReceived", 0), "Shield boosted (given)", extra.Get("shieldGiven", 0),
                "Shield boosted (received)", extra.Get("shieldReceived", 0),
                "Hull repaired (given)", extra.Get("hullGiven", 0), "Hull repaired (received)", extra.Get("hullReceived", 0), "Logistics records", extra.Get("logisticsRecords", 0)))
            categories.Push(this.Category("Command and mining bursts", "Burst applications", extra.Get("burstApplications", 0),
                "Recipient applications", extra.Get("burstRecipients", 0), "Most members affected by one burst", extra.Get("burstMaxRecipients", 0)))
            categories.Push(this.Category("Fleet and waiting", "Fleet warps", extra.Get("fleetWarps", 0),
                "Regroups", extra.Get("regroups", 0), "Please wait messages", extra.Get("pleaseWait", 0)))
            categories.Push(this.Category("Combat peaks", "Peak incoming DPS (10 seconds)", Round(view["peaks"][1], 2),
                "Peak outgoing DPS (10 seconds)", Round(view["peaks"][2], 2),
                "Peak incoming alpha strike", view["alpha"][1], "Peak outgoing alpha strike", view["alpha"][2]))
            if filters && filters.Get("character", "") = "" {
                metrics := categories[-1]["metrics"]
                for index, key in ["peak1", "peak2", "alpha1", "alpha2"]
                    metrics[index]["character"] := view["leaders"].Get(key, "")
            }
            gas := 0, solid := 0, residue := 0
            for material in view["materials"] {
                if RegExMatch(material["name"], "i)^(Fullerite-)|\b(?:Cytoserocin|Mykoserocin)\b")
                    gas += material["units"]
                else
                    solid += material["units"]
                residue += material["residue"]
            }
            categories.Push(this.Category("Mining and gas harvesting", "Gas harvested (units)", gas, "Other materials mined (units)", solid,
                "Named residue (units)", residue, "Unattributed residue records (units)", view["extraResidue"]))
        }
        if filters && filters.Get("character", "") = "" {
            keys := Map("Gas harvested (units)", "gas", "Other materials mined (units)", "solid",
                "Named residue (units)", "residue", "Unattributed residue records (units)", "extraResidue")
            for group in this.ArchiveGroups
                loop (group.Length - 1) // 2
                    keys[group[2 * A_Index + 1]] := "count" group[2 * A_Index]
            for label, key in Map("Armor repaired (given)", "armorGiven", "Armor repaired (received)", "armorReceived",
                "Shield boosted (given)", "shieldGiven", "Shield boosted (received)", "shieldReceived",
                "Hull repaired (given)", "hullGiven", "Hull repaired (received)", "hullReceived", "Logistics records", "logisticsRecords",
                "Burst applications", "burstApplications", "Recipient applications", "burstRecipients",
                "Most members affected by one burst", "burstMaxRecipients", "Fleet warps", "fleetWarps", "Regroups", "regroups", "Please wait messages", "pleaseWait")
                keys[label] := key
            appLeaders := stats ? stats.CharacterLeaders() : Map()
            for label, key in Map("Switch requests", "switchesRequested", "Confirmed focus changes", "switchesConfirmed",
                "Longest switching combo", "maxCombo", "Character launches (thumbnails created)", "characterLaunches",
                "Alerts triggered", "alertsTriggered", "Enabled events observed", "monitoredEvents",
                "Monitored damage taken", "monitoredDamageIn", "Monitored damage dealt", "monitoredDamageOut",
                "Monitored local system changes", "monitoredSystemChanges") {
                keys[label] := "app:" key
            }
            for category in categories
                for metric in category["metrics"]
                    if keys.Has(metric["label"])
                        metric["character"] := InStr(keys[metric["label"]], "app:") = 1
                            ? appLeaders.Get(SubStr(keys[metric["label"]], 5), "")
                            : view.Get("leaders", Map()).Get(keys[metric["label"]], "")
            for material in view.Get("materials", [])
                material["character"] := view.Get("leaders", Map()).Get("material:" material["name"], "")
        }
        return Map("enabled", !!live, "available", !!stats, "scanning", !!(stats && stats.job && stats.job.busy),
            "scanStatus", stats ? stats.status : "", "categories", categories, "view", view,
            "progressText", stats ? this.ProgressText(stats.progress) : "", "progress", stats ? stats.progress : 0)
    }

    static Category(title, pairs*) {
        metrics := []
        loop pairs.Length // 2
            metrics.Push(Map("label", pairs[A_Index * 2 - 1], "value", pairs[A_Index * 2]))
        return Map("title", title, "metrics", metrics)
    }

    static SavedCharacters(config) {
        names := Map()
        names.CaseSense := "Off"
        add := (name) => this.AddCharacter(names, name)
        for , profile in config["_Profiles"] {
            for section in ["Thumbnail Positions", "Client Possitions"]
                for name in profile.Get(section, Map())
                    add(name)
            for kind in ["character-shortcuts", "character-colors", "custom-names", "exceptions", "hidden-characters", "monitored-characters"]
                for row in SettingsTables.Read(profile, kind)
                    add(row[1])
            for group in profile.Get("Hotkey Groups", Map())
                for row in SettingsTables.Read(profile, "group-members", group)
                    add(row[1])
            for name in profile["Game Logs Monitoring"].Get("charsIds", Map())
                add(name)
        }
        return names.Count
    }

    static AddCharacter(names, rawName) {
        name := Trim(SettingsTables.CleanName(rawName))
        if name != "" && name != "EVE" && name != "Char Screen"
            names[name] := true
    }
}
