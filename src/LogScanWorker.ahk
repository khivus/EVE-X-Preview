#Requires AutoHotkey v2.0
#SingleInstance Off
#NoTrayIcon
; This file is also embedded as an alternate script in the standalone EXE.
; Background mode lowers disk I/O and CPU priority as well as keeping UI isolated.
if !DllCall("SetPriorityClass", "Ptr", DllCall("GetCurrentProcess", "Ptr"), "UInt", 0x100000)
    try ProcessSetPriority("BelowNormal")
OnError(WorkerError)
if A_Args.Length != 2
    ExitApp(1)
jobDirectory := A_Args[1]
parentPid := Integer(A_Args[2])
parentWatch := (*) => (ProcessExist(parentPid) ? 0 : ExitApp())
SetTimer(parentWatch, 1000)
try {
    request := StrSplit(RTrim(FileRead(jobDirectory "\input.txt", "UTF-8"), "`r`n"), "`n", "`r")
    mode := request[1], directory := request[2]
    if mode = "game"
        output := ScanGame(directory)
    else if mode = "chat"
        output := ScanChat(directory, request, jobDirectory)
    else if mode = "statistics"
        output := StatisticsArchive.Run(directory, request[3], request[4], jobDirectory, request)
    else if mode = "statistics-live"
        output := StatisticsArchive.Run(directory, request[3], request[4], jobDirectory, request, request[5], request[6])
    else if mode = "statistics-view"
        output := StatisticsArchive.View(directory, request[3], request[4])
    else
        throw Error("Unknown log scan mode")
    PublishResult(jobDirectory, output)
    ExitApp()
} catch as err {
    WorkerError(err)
}

WorkerError(err, *) {
    if A_Args.Length
        try FileAppend(err.Message " at " err.File ":" err.Line, A_Args[1] "\error.txt", "UTF-8")
    ExitApp(1)
}

ScanGame(directory) {
    newest := Map()
    Loop Files, directory "\*.txt", "F" {
        if !RegExMatch(A_LoopFileName, "i)^(\d{8}_\d{6})_(\d+)\.txt$", &match)
            continue
        id := match[2], session := match[1]
        if !newest.Has(id) || StrCompare(session, newest[id].session) > 0
            newest[id] := {path: A_LoopFileFullPath, session: session}
    }
    output := ""
    for id, entry in newest {
        listener := ""
        try {
            file := FileOpen(entry.path, "r", "UTF-8")
            if file {
                try {
                    header := StrSplit(file.Read(4096), "`n", "`r")
                    if header.Length >= 4 && RegExMatch(header[3], "^\s*Listener:\s*(.+?)\s*$", &name)
                        && RegExMatch(header[4], "i)^\s*Session Started:\s*\d{4}\.\d{2}\.\d{2} \d{2}:\d{2}:\d{2}\s*$")
                        listener := name[1]
                } finally file.Close()
            }
        }
        output .= id "`t" entry.path "`t" entry.session "`t" listener "`n"
    }
    return output
}

ScanChat(directory, request, jobDirectory) {
    active := Map(), cache := Map(), newest := Map()
    for index, character in request
        if index > 2 && character != ""
            active[character] := true
    if FileExist(jobDirectory "\cache.txt") {
        for row in StrSplit(FileRead(jobDirectory "\cache.txt", "UTF-8"), "`n", "`r") {
            fields := StrSplit(row, "`t")
            if fields.Length = 4
                cache[fields[1]] := {stamp: fields[2], listener: fields[3], session: fields[4]}
        }
    }
    nextCache := "", candidates := ""
    Loop Files, directory "\Local_*.txt", "F" {
        if !RegExMatch(A_LoopFileName, "i)^Local_\d{8}_\d{6}_\d+\.txt$")
            continue
        path := A_LoopFileFullPath
        stamp := A_LoopFileTimeCreated "|" A_LoopFileTimeModified "|" A_LoopFileSize
        candidates .= A_LoopFileName "`t" path "`t" stamp "`n"
    }
    ; Current sessions are found first, but all headers still determine final ordering.
    for row in StrSplit(Sort(candidates, "R"), "`n", "`r") {
        fields := StrSplit(row, "`t")
        if fields.Length != 3
            continue
        path := fields[2], stamp := fields[3]
        if cache.Has(path) && cache[path].stamp = stamp
            header := cache[path]
        else {
            try header := ReadChatHeader(path)
            catch
                continue
            if header.listener = "" || header.session = ""
                continue ; Incomplete/locked files are retried on the next discovery.
        }
        nextCache .= path "`t" stamp "`t" header.listener "`t" header.session "`n"
        if active.Has(header.listener) && (!newest.Has(header.listener) || StrCompare(header.session, newest[header.listener].session) > 0) {
            newest[header.listener] := {path: path, session: header.session}
            PublishResult(jobDirectory, ChatResult(newest))
        }
    }
    FileAppend(nextCache, jobDirectory "\cache.tmp", "UTF-8")
    FileMove(jobDirectory "\cache.tmp", jobDirectory "\cache.txt", 1)
    return ChatResult(newest)
}

ChatResult(newest) {
    output := ""
    for character, entry in newest {
        try seed := ReadLatestSystem(entry.path)
        catch
            continue
        output .= character "`t" entry.path "`t" entry.session "`t" seed.system "`t" seed.cursor "`t" seed.identity "`n"
    }
    return output
}

ReadChatHeader(path) {
    header := {listener: "", session: ""}
    file := FileOpen(path, "r", "UTF-8")
    if !file
        return header
    try {
        ; Bound header I/O even if an unrelated/corrupt log has a giant first line.
        text := file.Read(8192)
        for line in StrSplit(text, "`n", "`r") {
            if RegExMatch(line, "^\s*Listener:\s*(.+?)\s*$", &match)
                header.listener := match[1]
            else if RegExMatch(line, "i)^\s*Session started:\s*(\d{4}\.\d{2}\.\d{2} \d{2}:\d{2}:\d{2})", &match)
                header.session := RegExReplace(match[1], "\D")
            if header.listener != "" && header.session != ""
                break
        }
    } finally file.Close()
    return header
}

ReadLatestSystem(path) {
    file := FileOpen(path, "r", "UTF-8")
    if !file
        throw Error("Chat log could not be opened")
    try {
        info := Buffer(52, 0)
        identity := ""
        if DllCall("GetFileInformationByHandle", "Ptr", file.Handle, "Ptr", info)
            identity := NumGet(info, 28, "UInt") ":" NumGet(info, 44, "UInt") ":" NumGet(info, 48, "UInt")
        cursor := file.Length, ending := cursor, prefix := "", system := ""
        utf16 := InStr(file.Encoding, "UTF-16")
        while ending > 0 {
            start := Max(0, ending - 65536)
            if utf16
                start -= Mod(start, 2)
            file.Seek(start)
            text := file.Read(utf16 ? 32768 : 65536) prefix
            ; Pick the last complete system announcement, never an unfinished line.
            lines := StrSplit(text, "`n", "`r")
            if SubStr(text, -1) != "`n"
                lines.Pop()
            for index, line in lines {
                if (index = 1 && start > 0) || StrLen(line) > 8192
                    continue
                if RegExMatch(line, "^\x{FEFF}?\s*\[\s*\d{4}\.\d{2}\.\d{2} \d{2}:\d{2}:\d{2}\s*\] EVE System > Channel changed to Local :\s*(.+?)\s*$", &match)
                    system := match[1]
            }
            if system != ""
                break
            prefix := SubStr(text, 1, 2048)
            ending := start
        }
        return {system: system, cursor: cursor, identity: identity}
    } finally file.Close()
}

PublishResult(jobDirectory, output) {
    static serial := 0
    for name in ["output.tmp", "progress.tmp"]
        if FileExist(jobDirectory "\" name)
            FileDelete(jobDirectory "\" name)
    FileAppend(output, jobDirectory "\output.tmp", "UTF-8")
    MoveWorkerFile(jobDirectory "\output.tmp", jobDirectory "\output.txt")
    FileAppend(String(++serial), jobDirectory "\progress.tmp", "UTF-8")
    MoveWorkerFile(jobDirectory "\progress.tmp", jobDirectory "\progress.txt")
}

MoveWorkerFile(source, destination) {
    loop 21 {
        if DllCall("MoveFileExW", "Str", source, "Str", destination, "UInt", 1)
            return A_Index - 1
        reason := A_LastError
        if !(reason = 32 || reason = 33 || reason = 5) || A_Index = 21
            throw OSError(reason)
        Sleep(100)
    }
}

; The low-priority worker is the only archive writer. The append journal
; checkpoints byte cursors and aggregate counters without retaining log text.
class StatisticsArchive {
    static MetricCount := 13
    liveSince := ""

    static Run(game, chat, storage, jobDirectory, request, cutoff := "", liveSince := "") {
        mutex := DllCall("CreateMutexW", "Ptr", 0, "Int", 0, "Str", "Local\EVE-X-Preview.StatisticsArchive", "Ptr")
        if !mutex
            throw OSError()
        result := DllCall("WaitForSingleObject", "Ptr", mutex, "UInt", 0, "UInt")
        try {
            if result != 0 && result != 0x80
                return "An archive scan is already running."
            DirCreate(storage)
            log := storage "\scan.log"
            try {
                if FileExist(log) && FileGetSize(log) >= 524288
                    FileMove(log, storage "\scan.previous.log", true)
            }
            this.Log(storage, "Started; scope=" (cutoff != "" ? "live" : request.Length > 4 ? "monitored" : "archive")
                "; gameFolder=" (game && DirExist(game) ? "found" : "unavailable")
                "; chatFolder=" (chat && DirExist(chat) ? "found" : "unavailable"))
            if cutoff != "" {
                PublishResult(jobDirectory, "Finding EVE log files…")
                candidates := []
                this.FindLiveFiles(game, "game", cutoff, candidates, jobDirectory)
                this.FindLiveFiles(chat, "chat", cutoff, candidates, jobDirectory)
                if FileExist(jobDirectory "\stop.txt")
                    return "Archive scan stopped. Start a scan to continue."
                if !candidates.Length {
                    this.Log(storage, "Ended; no changed live logs")
                    return "Archive scan complete."
                }
            }
            this.Log(storage, "Loading saved scan checkpoints")
            PublishResult(jobDirectory, "Loading saved scan checkpoints…")
            archive := StatisticsArchive(storage, jobDirectory)
            archive.liveSince := liveSince
            if archive.stopped {
                this.Log(storage, "Stopped while loading checkpoints")
                return archive.Report("Archive scan stopped. Start a scan to continue.")
            }
            if !FileExist(archive.dashboard) && archive.files.Count {
                archive.BuildDashboard(true)
                if archive.stopped
                    return archive.Report("Archive scan stopped. Start a scan to continue.")
            }
            archive.phaseCount := 0
            if cutoff != "" {
                archive.totalFiles := candidates.Length
                archive.scanStarted := A_TickCount
                archive.Phase("Scanning EVE logs…")
                for candidate in candidates {
                    if archive.ShouldStop()
                        break
                    archive.ProcessFile(candidate.path, candidate.stamp, candidate.kind)
                }
            } else if request.Length > 4 {
                archive.totalFiles := request.Length - 4
                archive.scanStarted := A_TickCount
                archive.Phase("Scanning EVE logs…")
                for index, row in request {
                    if index <= 4
                        continue
                    if archive.ShouldStop()
                        break
                    parts := StrSplit(row, "`t", , 2)
                    try {
                        stamp := FileGetTime(parts[2], "C") "|" FileGetTime(parts[2], "M") "|" FileGetSize(parts[2])
                        archive.ProcessFile(parts[2], stamp, parts[1])
                    } catch {
                        archive.skipped += 1
                        archive.finished += 1
                    }
                }
            } else {
                archive.Phase("Finding EVE log files…")
                archive.totalFiles := archive.CountFiles(game, "game") + archive.CountFiles(chat, "chat")
                archive.phaseCount := 0
                archive.scanStarted := A_TickCount
                archive.Phase("Scanning EVE logs…")
                archive.Scan(game, "game")
                archive.Scan(chat, "chat")
            }
            archive.Checkpoint()
            archive.CompactIfNeeded()
            if archive.changed || archive.NeedsDashboard()
                archive.BuildDashboard()
            this.Log(storage, "Ended; stopped=" archive.stopped "; checked=" archive.finished
                "; total=" archive.totalFiles "; unreadable=" archive.skipped "; elapsedMs=" (A_TickCount - archive.started))
            if archive.stopped
                return archive.Report("Archive scan stopped. Start a scan to continue.")
            return archive.Report(archive.skipped ? "Archive scan completed with unreadable files. Start a scan to retry." : "Archive scan complete.")
        } catch as err {
            this.Log(storage, "Failed: " err.Message " at line " err.Line)
            throw err
        } finally {
            if result = 0 || result = 0x80
                DllCall("ReleaseMutex", "Ptr", mutex)
            DllCall("CloseHandle", "Ptr", mutex)
        }
    }

    static Log(storage, message) {
        try FileAppend(FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss") "  " message "`n", storage "\scan.log", "UTF-8-RAW")
    }

    static FindLiveFiles(directory, kind, cutoff, candidates, jobDirectory) {
        if directory = "" || !DirExist(directory)
            return
        pattern := this.Pattern(kind)
        Loop Files, directory "\*.txt", "F" {
            if Mod(A_Index - 1, 256) = 0 && FileExist(jobDirectory "\stop.txt")
                break
            if A_LoopFileTimeModified >= cutoff && RegExMatch(A_LoopFileName, pattern)
                candidates.Push({path: A_LoopFileFullPath, kind: kind,
                    stamp: A_LoopFileTimeCreated "|" A_LoopFileTimeModified "|" A_LoopFileSize})
        }
    }

    __New(storage, jobDirectory := "", loadIndex := true) {
        this.storage := storage
        this.path := storage "\archive-index.tsv"
        this.summary := storage "\archive-summary.tsv"
        this.dashboard := storage "\dashboard.tsv"
        this.jobDirectory := jobDirectory
        if !loadIndex
            return
        this.files := StatisticsFileIndex()
        this.paths := StatisticsFileIndex(true)
        this.totals := this.EmptyMetrics()
        this.pending := ""
        this.changed := false
        this.lastCheck := this.started := A_TickCount
        this.totalFiles := this.finished := 0
        this.phaseCount := 0
        this.scanStarted := 0
        this.lastLog := this.started
        this.lines := 0
        this.stopped := false
        this.skipped := 0
        if FileExist(this.path) {
            file := this.OpenShared(this.path, 0x700)
            try {
                ; Release the handle before parsing so journal appends and
                ; replacements never wait for a large dashboard query.
                bytes := file.Length - file.Pos
                snapshot := Buffer(bytes + 1, 0)
                length := file.RawRead(snapshot, bytes)
            } finally file.Close()
            text := length ? StrGet(snapshot, length, "UTF-8") : ""
            snapshot := 0
            torn := text != "" && SubStr(text, -1) != "`n"
            if torn
                text := SubStr(text, 1, InStr(text, "`n", , -1))
            basic := 4 + StatisticsArchive.MetricCount
            Loop Parse, text, "`n", "`r" {
                if jobDirectory != "" && Mod(A_Index, 1024) = 0 {
                    this.phaseCount := A_Index
                    this.Phase("Loading saved scan checkpoints…", false)
                    if this.ShouldStop()
                        return
                }
                fields := StrSplit(A_LoopField, "`t")
                if fields.Length != basic && fields.Length != basic + 3 && fields.Length != basic + 4
                    continue
                values := []
                loop StatisticsArchive.MetricCount
                    values.Push(Number(fields[4 + A_Index]))
                entry := this.Replace(fields[1], fields[2], fields[3], Number(fields[4]), values, jobDirectory != "")
                if fields.Length > basic
                    entry.character := fields[basic + 1], entry.session := fields[basic + 2], entry.detail := fields[basic + 3]
                if fields.Length = basic + 4
                    entry.window := fields[basic + 4]
                this.lines += 1
            }
            if torn && jobDirectory != ""
                this.Compact()
        }
    }

    Phase(status, immediate := true) {
        now := A_TickCount
        if immediate || now - this.lastCheck >= 3000 {
            PublishResult(this.jobDirectory, this.Report(status))
            this.lastCheck := now
        }
        if immediate || now - this.lastLog >= 10000 {
            StatisticsArchive.Log(this.storage, status " entries=" this.phaseCount "; checked=" this.finished
                "; total=" this.totalFiles "; unreadable=" this.skipped)
            this.lastLog := now
        }
    }

    static Pattern(kind) => kind = "game" ? "i)^\d{8}_\d{6}_\d+\.txt$" : "i)^.+_\d{8}_\d{6}_\d+\.txt$"

    CountFiles(directory, kind) {
        count := 0
        if directory = "" || !DirExist(directory)
            return count
        pattern := StatisticsArchive.Pattern(kind)
        Loop Files, directory "\*.txt", "F" {
            if Mod(A_Index, 256) = 0 {
                this.Phase("Finding EVE log files…", false)
                if this.ShouldStop()
                    break
            }
            if RegExMatch(A_LoopFileName, pattern) {
                count += 1
                this.phaseCount += 1
            }
        }
        return count
    }

    Scan(directory, kind) {
        if this.stopped || directory = "" || !DirExist(directory)
            return
        pattern := StatisticsArchive.Pattern(kind)
        Loop Files, directory "\*.txt", "F" {
            if Mod(A_Index - 1, 64) = 0 && this.ShouldStop()
                break
            if !RegExMatch(A_LoopFileName, pattern)
                continue
            this.ProcessFile(A_LoopFileFullPath, A_LoopFileTimeCreated "|" A_LoopFileTimeModified "|" A_LoopFileSize, kind)
        }
    }

    ProcessFile(path, stamp, kind) {
        key := this.paths.Get(path, "")
        entry := key != "" ? this.files[key] : 0
        upgrade := entry && !this.liveSince && (this.CombatNeedsBackfill(entry) || this.ExtendedNeedsBackfill(entry))
        if !entry || (entry.stamp != stamp && (this.liveSince = "" || entry.stamp != "live:" stamp)) || !entry.HasOwnProp("detail") || upgrade {
            try this.ScanFile(path, stamp, kind)
            catch as err {
                this.skipped += 1
                if this.skipped <= 10
                    StatisticsArchive.Log(this.storage, "Skipped " kind " log; error=" err.Message)
            }
        }
        if !this.stopped
            this.finished += 1
        if A_TickCount - this.lastCheck >= 3000
            this.Checkpoint(true)
    }

    ExtendedNeedsBackfill(entry) {
        if !entry.HasOwnProp("detail")
            return true
        for row in StrSplit(entry.detail, "|")
            if row != "" && !InStr(row, ";V6")
                return true
        return false
    }

    CombatNeedsBackfill(entry) => (entry.values[4] || entry.values[5])
        && (!entry.HasOwnProp("window") || InStr(entry.window, "old:") = 1)

    ScanFile(path, stamp, kind) {
        file := FileOpen(path, "r", "UTF-8")
        if !file
            throw Error("Could not open log.")
        try {
            info := Buffer(52, 0)
            if !DllCall("GetFileInformationByHandle", "Ptr", file.Handle, "Ptr", info)
                throw OSError()
            key := Format("{}:{}:{}:{}:{}", kind, NumGet(info, 28, "UInt"), NumGet(info, 44, "UInt"), NumGet(info, 48, "UInt"), NumGet(info, 4, "Int64"))
            previous := this.files.Get(key, 0)
            partial := previous && InStr(previous.stamp, "live:") = 1
            upgrade := previous && (this.CombatNeedsBackfill(previous) || this.ExtendedNeedsBackfill(previous))
            backfill := (partial || upgrade) && this.liveSince = ""
            window := StatisticsDamageWindow(), legacyWindow := false
            if previous && !backfill && previous.HasOwnProp("detail")
                && previous.cursor >= file.Pos && file.Length >= previous.cursor {
                file.Seek(previous.cursor)
                values := previous.values.Clone()
                header := {listener: previous.character, session: previous.session}
                days := this.DecodeDays(previous.detail)
                legacyWindow := upgrade
                if previous.HasOwnProp("window")
                    window.Restore(previous.window)
            } else {
                partial := false
                values := this.EmptyMetrics()
                values[kind = "game" ? 1 : 2] := 1
                header := ReadChatHeader(path)
                if header.session = "" && RegExMatch(path, "(\d{8})_(\d{6})_(\d+)\.txt$", &name)
                    header.session := name[1] name[2]
                days := Map()
            }
            date := "", day := 0, dayStart := 0
            cursor := file.Pos, limit := file.Length
            while file.Pos < limit {
                line := file.ReadLine()
                if file.Pos <= cursor || file.Pos > limit || (file.Pos = limit && !this.EndedWithNewline(file))
                    break ; Only checkpoint complete records.
                cursor := file.Pos
                if RegExMatch(line, "^\x{FEFF}?\s*\[\s*(\d{4})\.(\d{2})\.(\d{2}) (\d{2}:\d{2}:\d{2})\s*\]", &record) {
                    nextDate := record[1] record[2] record[3]
                    if this.liveSince && (nextDate StrReplace(record[4], ":")) < this.liveSince
                        partial := true
                    else {
                        if nextDate != date {
                            date := nextDate
                            if kind = "game"
                                dayStart := DateDiff(date, "19700101", "Days") * 86400
                            if !days.Has(date)
                                days[date] := this.NewDay()
                            day := days[date]
                        }
                        this.CountLine(line, kind, day.values, day, dayStart, record[4], window)
                    }
                }
                if Mod(A_Index, 256) = 0 && (this.ShouldStop() || A_TickCount - this.lastCheck >= 3000) {
                    if !backfill
                        this.Record(key, path, partial ? "live:" : "", cursor, values, header, days, window, legacyWindow)
                    this.Checkpoint(true)
                    if this.stopped
                        return
                }
            }
            ; Mark unchanged incomplete tails as seen. If new bytes arrive the
            ; next scan resumes from the last complete record.
            this.Record(key, path, (partial ? "live:" : "") stamp, cursor, values, header, days, window, legacyWindow)
        } finally file.Close()
    }

    EndedWithNewline(file) {
        position := file.Pos
        bytes := InStr(file.Encoding, "UTF-16") ? 2 : 1
        if position < bytes
            return false
        file.Seek(position - bytes)
        newline := bytes = 2 ? file.ReadUShort() : file.ReadUChar()
        return newline = 10
    }

    CountLine(line, kind, values, day, dayStart, time, window) {
        this.CountExtended(line, kind, day)
        if kind = "game" {
            if InStr(line, "(combat)") {
                values[3] += 1
                if RegExMatch(line, "i)\(combat\)\s*(?:<[^>]*>)*([0-9]+(?:\.[0-9]+)?)(?:<[^>]*>|\s)*(from|to)(?=<|\s)", &hit) {
                    direction := hit[2] = "from" ? 1 : 2, damage := Number(hit[1])
                    values[direction + 3] += damage
                    second := dayStart + Number(SubStr(time, 1, 2)) * 3600
                        + Number(SubStr(time, 4, 2)) * 60 + Number(SubStr(time, 7, 2))
                    peer := this.HitPeer(line)
                    if peer != ""
                        this.Bump(day.contacts, (direction = 1 ? "damageIn:" : "damageOut:") peer, damage)
                    if damage > day.alpha[direction]
                        day.alpha[direction] := damage, day.alphaPeers[direction] := peer
                    rate := window.Add(second, direction, damage, peer)
                    if rate > day.peak[direction]
                        day.peak[direction] := rate, day.peakPeers[direction] := window.TopPeer(direction)
                }
                else if InStr(line, "misses") {
                    values[6] += 1
                    text := RegExReplace(line, "<[^>]*>")
                    if RegExMatch(text, "\(combat\) (.+?) misses you completely", &miss)
                        key := "missIn", peer := Trim(miss[1])
                    else if RegExMatch(text, "\(combat\) Your .+? misses (.+?) completely", &miss)
                        key := "missOut", peer := Trim(miss[1])
                    else
                        key := "", peer := ""
                    if key != "" {
                        this.Bump(day.extra, key)
                        this.Bump(day.contacts, key ":" peer)
                    }
                }
            } else if InStr(line, "(mining)")
                this.CountMining(line, day)
            else if InStr(line, "Jumping from") || InStr(line, "jumping to")
                values[7] += 1
            else if InStr(line, "Undocking from")
                values[10] += 1
            else if InStr(line, "wants you to join their fleet")
                values[11] += 1
            else if InStr(line, "deactivates due to the destruction") || InStr(line, "pale shadow of its former glory")
                || InStr(line, "has completed operations")
                values[12] += 1
            else if InStr(line, "A Conduit Field activated by")
                values[13] += 1
        } else if RegExMatch(line, "^\x{FEFF}?\s*\[\s*\d{4}\.\d{2}\.\d{2} \d{2}:\d{2}:\d{2}\s*\]\s*([^>]+)>\s*(.*)", &message) {
            if Trim(message[1]) != "EVE System"
                values[9] += 1
            else if InStr(message[2], "Channel changed to Local :") = 1
                values[8] += 1
        }
    }

    ShouldStop() {
        this.stopped := this.stopped || !!FileExist(this.jobDirectory "\stop.txt")
        return this.stopped
    }

    EmptyMetrics() => [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]

    Replace(key, path, stamp, cursor, values, aggregate := true) {
        old := this.files.Get(key, 0)
        if aggregate
            loop StatisticsArchive.MetricCount
                this.totals[A_Index] += values[A_Index] - (old ? old.values[A_Index] : 0)
        if old && this.paths.Get(old.path, "") = key
            this.paths.Delete(old.path)
        this.files[key] := {path: path, stamp: stamp, cursor: cursor, values: values}
        this.paths[path] := key
        return this.files[key]
    }

    Record(key, path, stamp, cursor, seen, header, days, window, legacyWindow) {
        values := this.EmptyMetrics()
        values[1] := seen[1], values[2] := seen[2]
        for , day in days
            loop StatisticsArchive.MetricCount - 2
                values[A_Index + 2] += day.values[A_Index + 2]
        entry := this.Replace(key, path, stamp, cursor, values)
        entry.character := StrReplace(header.listener, "`t"), entry.session := header.session, entry.detail := this.EncodeDays(days)
        entry.window := (legacyWindow ? "old:" : "") window.Serialize()
        this.pending .= this.Row(key, entry)
        this.changed := true
        this.lines += 1
    }

    Row(key, entry) => key "`t" entry.path "`t" entry.stamp "`t" entry.cursor "`t" this.Join(entry.values)
        . (entry.HasOwnProp("detail") ? "`t" entry.character "`t" entry.session "`t" entry.detail : "")
        . (entry.HasOwnProp("window") ? "`t" entry.window : "") "`n"

    Join(values) {
        row := ""
        for value in values
            row .= (A_Index = 1 ? "" : "`t") value
        return row
    }

    Checkpoint(progress := false) {
        if this.pending != "" {
            file := this.OpenShared(this.path, 0x702)
            try file.Write(this.pending)
            finally file.Close()
            this.pending := ""
        }
        this.WriteAtomic(this.summary, this.Join(this.totals))
        if progress
            this.Phase("Scanning EVE logs…", false)
        this.lastCheck := A_TickCount
    }

    Report(status) => status "`n" this.finished "`t" Max(this.finished, this.totalFiles) "`t"
        . (A_TickCount - (this.scanStarted ? this.scanStarted : this.started)) "`t" this.skipped "`t" this.phaseCount

    OpenShared(path, flags) {
        loop 21 {
            try {
                file := FileOpen(path, flags, "UTF-8-RAW")
                if !file
                    throw OSError()
                return file
            } catch OSError as err {
                if !(err.Number = 32 || err.Number = 33) || A_Index = 21
                    throw err
                Sleep(100)
            }
        }
    }

    NewDay() => {values: this.EmptyMetrics(), peak: [0, 0], alpha: [0, 0], combatIndexed: true, mining: Map(), extraResidue: 0, extra: Map(), modules: Map(), systems: Map(), bursts: Map(), contacts: Map(), peakPeers: ["", ""], alphaPeers: ["", ""], extendedIndexed: true}

    Bump(counts, key, amount := 1) => counts[key] := counts.Get(key, 0) + amount

    CountExtended(line, kind, day) {
        text := RegExReplace(line, "<[^>]*>")
        if kind = "chat" {
            if RegExMatch(text, "EVE System > Channel changed to Local :\s*(.+?)\s*$", &system)
                this.Bump(day.systems, system[1])
            return
        }
        if RegExMatch(text, "\(notify\) Your (.+?) has applied bonuses to ([\d,]+) fleet members?\.\s*$", &burst) && InStr(burst[1], "Burst") {
            recipients := Number(StrReplace(burst[2], ","))
            this.Bump(day.extra, "burstApplications")
            this.Bump(day.extra, "burstRecipients", recipients)
            day.extra["burstMaxRecipients"] := Max(day.extra.Get("burstMaxRecipients", 0), recipients)
            this.Bump(day.bursts, burst[1])
            this.Bump(day.modules, burst[1])
        }
        if RegExMatch(text, "i)\(combat\)\s*([\d,.]+) remote (armor repaired|shield boosted|hull repaired|structure repaired) (by|to) .+ - (.+)$", &repair) {
            direction := repair[3] = "by" ? "Received" : "Given"
            type := repair[2] = "armor repaired" ? "armor" : repair[2] = "shield boosted" ? "shield" : "hull"
            this.Bump(day.extra, type direction, Number(StrReplace(repair[1], ",")))
            this.Bump(day.extra, "logisticsRecords")
            this.Bump(day.modules, Trim(repair[4]))
        } else if InStr(text, "(combat)") && RegExMatch(text, ".* - (.+?) - (?:Hits|Smashes|Penetrates|Grazes|Glances Off|Wrecks|misses)", &module)
            this.Bump(day.modules, Trim(module[1]))
        else if RegExMatch(text, "External factors are preventing your (.+?) from responding", &module)
            this.Bump(day.modules, Trim(module[1]))
        if RegExMatch(text, "\(notify\) Following .+ in warp")
            this.Bump(day.extra, "fleetWarps")
        if RegExMatch(text, "\(notify\) Regrouping(?: to .+)?\s*$")
            this.Bump(day.extra, "regroups")
        if RegExMatch(text, "\(notify\) Please wait\.\.\.\s*$")
            this.Bump(day.extra, "pleaseWait")
    }

    HitPeer(line) {
        ; The outer bold name ends immediately before the weapon delimiter.
        if RegExMatch(line, "i)(?:from|to)(?:<[^>]*>|\s)*(.+?)</b><font size=\d+><color=0x[0-9a-f]+> - ", &peer)
            return Trim(RegExReplace(peer[1], "<[^>]*>"))
        text := RegExReplace(line, "<[^>]*>")
        if RegExMatch(text, "i)\(combat\)\s*[\d.]+ (?:from|to) (.+) - .+? - (?:Hits|Smashes|Penetrates|Grazes|Glances Off|Wrecks)", &peer)
            return Trim(peer[1])
        if RegExMatch(text, "i)\(combat\)\s*[\d.]+ (?:from|to) (.+)$", &peer)
            return Trim(peer[1])
        return ""
    }

    CountMining(line, day) {
        line := RegExReplace(line, "<[^>]*>")
        if RegExMatch(line, "i)\(mining\)\s*You mined ([\d,]+) units? of (.+?)(?: with a lost residue of ([\d,]+) units?)?\s*[.]?$", &mined) {
            material := Trim(mined[2])
            if !day.mining.Has(material)
                day.mining[material] := [0, 0]
            day.mining[material][1] += Number(StrReplace(mined[1], ","))
            if mined[3] != ""
                day.mining[material][2] += Number(StrReplace(mined[3], ","))
        } else if RegExMatch(line, "i)\(mining\)\s*Additional ([\d,]+) units? depleted from .+ as residue", &residue)
            day.extraResidue += Number(StrReplace(residue[1], ","))
    }

    Escape(material) {
        material := StrReplace(material, "%", "%25")
        return StrReplace(StrReplace(StrReplace(StrReplace(material, ",", "%2C"), "|", "%7C"), ";", "%3B"), "`t", "%09")
    }

    Unescape(material) => StrReplace(StrReplace(StrReplace(StrReplace(StrReplace(material, "%09", "`t"), "%3B", ";"), "%7C", "|"), "%2C", ","), "%25", "%")

    EncodeDays(days) {
        text := ""
        for date, day in days {
            row := date
            loop StatisticsArchive.MetricCount - 2
                row .= "," day.values[A_Index + 2]
            ; Keep the legacy tail columns reserved. The rolling window now
            ; lives on the file so it can continue across midnight and resumes.
            row .= "," day.peak[1] "," day.peak[2] ",,0,0," day.extraResidue
                . "," day.alpha[1] "," day.alpha[2] "," day.combatIndexed
            if day.extendedIndexed
                row .= ";V6"
            for field in ["extra", "modules", "systems", "bursts", "contacts"]
                for name, count in day.%field%
                    row .= ";@" field "," this.Escape(name) "," count
            loop 2
                row .= ";@peer,peak" A_Index "," this.Escape(day.peakPeers[A_Index]) ";@peer,alpha" A_Index "," this.Escape(day.alphaPeers[A_Index])
            for material, record in day.mining
                row .= ";" this.Escape(material) "," record[1] "," record[2]
            text .= (text = "" ? "" : "|") row
        }
        return text
    }

    DecodeDays(detail) {
        days := Map()
        if detail = ""
            return days
        for row in StrSplit(detail, "|") {
            parts := StrSplit(row, ";")
            values := StrSplit(parts[1], ",")
            day := this.NewDay()
            day.extendedIndexed := InStr(row, ";V6") != 0
            loop StatisticsArchive.MetricCount - 2
                day.values[A_Index + 2] := Number(values[A_Index + 1])
            offset := StatisticsArchive.MetricCount
            day.extraResidue := Number(values[offset + 5])
            if values.Length >= offset + 8 {
                day.peak := [Number(values[offset]), Number(values[offset + 1])]
                day.alpha := [Number(values[offset + 6]), Number(values[offset + 7])]
                day.combatIndexed := !!Number(values[offset + 8])
            } else
                day.combatIndexed := !(day.values[4] || day.values[5])
            days[values[1]] := day
            for index, part in parts {
                if index = 1
                    continue
                if part = "V2" || part = "V3" || part = "V4" || part = "V5" || part = "V6"
                    continue
                fields := StrSplit(part, ",")
                if SubStr(fields[1], 1, 1) = "@" {
                    field := SubStr(fields[1], 2)
                    if field = "peer" {
                        index := Number(SubStr(fields[2], -1))
                        day.%(InStr(fields[2], "peak") = 1 ? "peakPeers" : "alphaPeers")%[index] := this.Unescape(fields[3])
                        continue
                    }
                    if field = "extra" || field = "modules" || field = "systems" || field = "bursts" || field = "contacts"
                        day.%field%[this.Unescape(fields[2])] := Number(fields[3])
                    continue
                }
                day.mining[this.Unescape(fields[1])] := [Number(fields[2]), Number(fields[3])]
            }
        }
        return days
    }

    static View(storage, character, range) {
        if FileExist(storage "\dashboard.tsv")
            return this.ViewDashboard(storage, character, range)
        archive := StatisticsArchive(storage)
        names := Map(), recent := Map(), mining := Map(), combined := archive.NewDay(), byCharacter := Map()
        names.CaseSense := recent.CaseSense := "Off"
        totals := archive.EmptyMetrics(), peaks := [0, 0], alpha := [0, 0], residue := 0, combatIndexed := true
        offsets := Map("day", 0, "week", -6, "month", -29, "year", -364)
        since := offsets.Has(range) ? SubStr(DateAdd(A_NowUTC, offsets[range], "Days"), 1, 8) : ""
        for key, entry in archive.files {
            if !entry.HasOwnProp("detail")
                continue
            if entry.character != ""
                names[entry.character] := true
            if range = "session" && InStr(key, "game:") = 1 {
                existing := recent.Get(entry.character, 0)
                if !existing || entry.session > existing.session
                    recent[entry.character] := entry
            }
        }
        for key, entry in archive.files {
            if character != "" && (!entry.HasOwnProp("detail") || entry.character != character)
                continue
            if !entry.HasOwnProp("detail") {
                if range = "all" {
                    loop StatisticsArchive.MetricCount
                        totals[A_Index] += entry.values[A_Index]
                    combatIndexed := combatIndexed && !(entry.values[4] || entry.values[5])
                }
                continue
            }
            if range = "session" && (InStr(key, "game:") != 1 || recent.Get(entry.character, 0) != entry)
                continue
            covered := !since || SubStr(entry.session, 1, 8) >= since
            if covered || entry.detail != "" {
                for date, day in archive.DecodeDays(entry.detail) {
                    if since && date < since
                        continue
                    covered := true
                    archive.MergeDay(combined, day)
                    if !byCharacter.Has(entry.character)
                        byCharacter[entry.character] := archive.NewDay()
                    archive.MergeDay(byCharacter[entry.character], day)
                    loop StatisticsArchive.MetricCount - 2
                        totals[A_Index + 2] += day.values[A_Index + 2]
                    peaks[1] := Max(peaks[1], day.peak[1]), peaks[2] := Max(peaks[2], day.peak[2])
                    alpha[1] := Max(alpha[1], day.alpha[1]), alpha[2] := Max(alpha[2], day.alpha[2])
                    combatIndexed := combatIndexed && day.combatIndexed
                    residue += day.extraResidue
                    for material, amount in day.mining {
                        if !mining.Has(material)
                            mining[material] := [0, 0]
                        mining[material][1] += amount[1], mining[material][2] += amount[2]
                    }
                }
            }
            if covered {
                totals[1] += entry.values[1], totals[2] += entry.values[2]
                if !byCharacter.Has(entry.character)
                    byCharacter[entry.character] := archive.NewDay()
                byCharacter[entry.character].values[1] += entry.values[1]
                byCharacter[entry.character].values[2] += entry.values[2]
            }
        }
        return archive.FormatView(totals, peaks, residue, names, mining, alpha, combatIndexed, combined, byCharacter)
    }

    PickLeader(leaders, bests, key, character, value) {
        if value > bests.Get(key, 0) {
            bests[key] := value
            leaders[key] := character
        }
    }

    IsDrone(name) {
        ; Official ESI universe categories 18 (drones) and 87 (fighters), 2026-11-21.
        ; Names are embedded so archive filtering works offline and in compiled workers.
        static names := Map()
        if !names.Count {
            names.CaseSense := "Off"
            names["'Aergia' Hobgoblin SW-300"] := true
            names["'Augmented' Acolyte"] := true
            names["'Augmented' Berserker"] := true
            names["'Augmented' Hammerhead"] := true
            names["'Augmented' Hobgoblin"] := true
            names["'Augmented' Hornet"] := true
            names["'Augmented' Ice Harvesting Drone"] := true
            names["'Augmented' Infiltrator"] := true
            names["'Augmented' Mining Drone"] := true
            names["'Augmented' Ogre"] := true
            names["'Augmented' Praetor"] := true
            names["'Augmented' Valkyrie"] := true
            names["'Augmented' Vespa"] := true
            names["'Augmented' Warrior"] := true
            names["'Augmented' Wasp"] := true
            names["'Dunk' Salvage Drone"] := true
            names["'Excavator' Ice Harvesting Drone"] := true
            names["'Excavator' Mining Drone"] := true
            names["'Integrated' Acolyte"] := true
            names["'Integrated' Berserker"] := true
            names["'Integrated' Hammerhead"] := true
            names["'Integrated' Hobgoblin"] := true
            names["'Integrated' Hornet"] := true
            names["'Integrated' Infiltrator"] := true
            names["'Integrated' Ogre"] := true
            names["'Integrated' Praetor"] := true
            names["'Integrated' Valkyrie"] := true
            names["'Integrated' Vespa"] := true
            names["'Integrated' Warrior"] := true
            names["'Integrated' Wasp"] := true
            names["'Pluto' Ice Harvesting Drone"] := true
            names["'Pluto' Mining Drone"] := true
            names["'Subverted' JVN-UC49"] := true
            names["10mn webscramblifying Drone"] := true
            names["Able_PLACEHOLDER"] := true
            names["Academy Acolyte"] := true
            names["Academy Hobgoblin"] := true
            names["Academy Warrior"] := true
            names["Acolyte EV-300"] := true
            names["Acolyte I"] := true
            names["Acolyte II"] := true
            names["Acolyte TD-300"] := true
            names["Ametat I"] := true
            names["Ametat II"] := true
            names["Antaeus I"] := true
            names["Antaeus II"] := true
            names["Arabellata SW-900-I"] := true
            names["Aralez"] := true
            names["Atlas SD-900-I"] := true
            names["Baker_PLACEHOLDER"] := true
            names["Berserker I"] := true
            names["Berserker II"] := true
            names["Berserker SW-900"] := true
            names["Berserker TP-900"] := true
            names["Bouncer I"] := true
            names["Bouncer II"] := true
            names["Caldari Navy Dragonfly"] := true
            names["Caldari Navy Hornet"] := true
            names["Caldari Navy Locust"] := true
            names["Caldari Navy Scarab"] := true
            names["Caldari Navy Vespa"] := true
            names["Caldari Navy Warden"] := true
            names["Caldari Navy Wasp"] := true
            names["Cenobite I"] := true
            names["Cenobite II"] := true
            names["Charlie_PLACEHOLDER"] := true
            names["Civilian Hobgoblin"] := true
            names["Civilian Mining Drone"] := true
            names["Curator I"] := true
            names["Curator II"] := true
            names["Cyclops I"] := true
            names["Cyclops II"] := true
            names["Darter TP-600-I"] := true
            names["Dog_PLACEHOLDER"] := true
            names["Dragonfly I"] := true
            names["Dragonfly II"] := true
            names["Dromi I"] := true
            names["Dromi II"] := true
            names["Easy_PLACEHOLDER"] := true
            names["Einherji I"] := true
            names["Einherji II"] := true
            names["Equite I"] := true
            names["Equite II"] := true
            names["FA-14 Templar"] := true
            names["Federation Navy Firbolg"] := true
            names["Federation Navy Garde"] := true
            names["Federation Navy Hammerhead"] := true
            names["Federation Navy Hobgoblin"] := true
            names["Federation Navy Ogre"] := true
            names["Federation Navy Satyr"] := true
            names["Federation Navy Siren"] := true
            names["Fighter Uno"] := true
            names["Firbolg I"] := true
            names["Firbolg II"] := true
            names["Garde I"] := true
            names["Garde II"] := true
            names["Gecko"] := true
            names["Gram I"] := true
            names["Gram II"] := true
            names["Gungnir I"] := true
            names["Gungnir II"] := true
            names["Hammerhead I"] := true
            names["Hammerhead II"] := true
            names["Hammerhead SD-600"] := true
            names["Harvester Mining Drone"] := true
            names["Heavy Armor Maintenance Bot I"] := true
            names["Heavy Armor Maintenance Bot II"] := true
            names["Heavy Hull Maintenance Bot I"] := true
            names["Heavy Hull Maintenance Bot II"] := true
            names["Heavy Mutated Drone"] := true
            names["Heavy Shield Maintenance Bot I"] := true
            names["Heavy Shield Maintenance Bot II"] := true
            names["Hellhound I"] := true
            names["Hobgoblin I"] := true
            names["Hobgoblin II"] := true
            names["Hobgoblin SD-300"] := true
            names["Hornet EC-300"] := true
            names["Hornet I"] := true
            names["Hornet II"] := true
            names["Humboldt EC-900-I"] := true
            names["Huntsman SW-600-I"] := true
            names["Ice Harvesting Drone I"] := true
            names["Ice Harvesting Drone II"] := true
            names["Immaculate TD-600-I"] := true
            names["Imperial Navy Acolyte"] := true
            names["Imperial Navy Cenobite"] := true
            names["Imperial Navy Curator"] := true
            names["Imperial Navy Equite"] := true
            names["Imperial Navy Infiltrator"] := true
            names["Imperial Navy Praetor"] := true
            names["Imperial Navy Templar"] := true
            names["Infiltrator EV-600"] := true
            names["Infiltrator I"] := true
            names["Infiltrator II"] := true
            names["Infiltrator TD-600"] := true
            names["Inshore EC-300-I"] := true
            names["Light Armor Maintenance Bot I"] := true
            names["Light Armor Maintenance Bot II"] := true
            names["Light Hull Maintenance Bot I"] := true
            names["Light Hull Maintenance Bot II"] := true
            names["Light Mutated Drone"] := true
            names["Light Shield Maintenance Bot I"] := true
            names["Light Shield Maintenance Bot II"] := true
            names["Locust I"] := true
            names["Locust II"] := true
            names["Luna SD-600-I"] := true
            names["Malleus I"] := true
            names["Malleus II"] := true
            names["Mantis I"] := true
            names["Mantis II"] := true
            names["Medium Armor Maintenance Bot I"] := true
            names["Medium Armor Maintenance Bot II"] := true
            names["Medium Hull Maintenance Bot I"] := true
            names["Medium Hull Maintenance Bot II"] := true
            names["Medium Mutated Drone"] := true
            names["Medium Shield Maintenance Bot I"] := true
            names["Medium Shield Maintenance Bot II"] := true
            names["Meganeura TP-900-I"] := true
            names["Mining Drone - Elite"] := true
            names["Mining Drone - Improved"] := true
            names["Mining Drone I"] := true
            names["Mining Drone II"] := true
            names["Mosquito EV-600-I"] := true
            names["Mutated 'Excavator' Ice Harvesting Drone"] := true
            names["Mutated 'Excavator' Mining Drone"] := true
            names["Mutated Ice Harvesting Drone"] := true
            names["Mutated Mining Drone"] := true
            names["Nertic EC-600-I"] := true
            names["Ogre I"] := true
            names["Ogre II"] := true
            names["Ogre SD-900"] := true
            names["Orbweaver SW-300-I"] := true
            names["Praetor EV-900"] := true
            names["Praetor I"] := true
            names["Praetor II"] := true
            names["Praetor TD-900"] := true
            names["QA Salvage Drone"] := true
            names["Repair Drone"] := true
            names["Republic Fleet Berserker"] := true
            names["Republic Fleet Bouncer"] := true
            names["Republic Fleet Dromi"] := true
            names["Republic Fleet Einherji"] := true
            names["Republic Fleet Gram"] := true
            names["Republic Fleet Valkyrie"] := true
            names["Republic Fleet Warrior"] := true
            names["Salvage Drone I"] := true
            names["Salvage Drone II"] := true
            names["Satyr I"] := true
            names["Satyr II"] := true
            names["Scarab I"] := true
            names["Scarab II"] := true
            names["Sentry Mutated Drone"] := true
            names["Shadow"] := true
            names["Siren I"] := true
            names["Siren II"] := true
            names["Skimmer TP-300-I"] := true
            names["Standup Ametat I"] := true
            names["Standup Ametat II"] := true
            names["Standup Antaeus I"] := true
            names["Standup Antaeus II"] := true
            names["Standup Cenobite I"] := true
            names["Standup Cenobite II"] := true
            names["Standup Cyclops I"] := true
            names["Standup Cyclops II"] := true
            names["Standup Dragonfly I"] := true
            names["Standup Dragonfly II"] := true
            names["Standup Dromi I"] := true
            names["Standup Dromi II"] := true
            names["Standup Einherji I"] := true
            names["Standup Einherji II"] := true
            names["Standup Equite I"] := true
            names["Standup Equite II"] := true
            names["Standup Firbolg I"] := true
            names["Standup Firbolg II"] := true
            names["Standup Gram I"] := true
            names["Standup Gram II"] := true
            names["Standup Gungnir I"] := true
            names["Standup Gungnir II"] := true
            names["Standup Locust I"] := true
            names["Standup Locust II"] := true
            names["Standup Malleus I"] := true
            names["Standup Malleus II"] := true
            names["Standup Mantis I"] := true
            names["Standup Mantis II"] := true
            names["Standup Satyr I"] := true
            names["Standup Satyr II"] := true
            names["Standup Scarab I"] := true
            names["Standup Scarab II"] := true
            names["Standup Shadow"] := true
            names["Standup Siren I"] := true
            names["Standup Siren II"] := true
            names["Standup Templar I"] := true
            names["Standup Templar II"] := true
            names["Standup Termite I"] := true
            names["Standup Termite II"] := true
            names["Standup Tyrfing I"] := true
            names["Standup Tyrfing II"] := true
            names["Stellate TD-300-I"] := true
            names["Stigmella SD-300-I"] := true
            names["Tabanida EV-900-I"] := true
            names["Templar I"] := true
            names["Templar II"] := true
            names["Termite I"] := true
            names["Termite II"] := true
            names["Tick EV-300-I"] := true
            names["Torafugu TD-900-I"] := true
            names["Tutorial Attack Drone"] := true
            names["Tyrfing I"] := true
            names["Tyrfing II"] := true
            names["Unanchoring Drone"] := true
            names["Valkyrie I"] := true
            names["Valkyrie II"] := true
            names["Valkyrie SW-600"] := true
            names["Valkyrie TP-600"] := true
            names["Vespa EC-600"] := true
            names["Vespa I"] := true
            names["Vespa II"] := true
            names["Warden I"] := true
            names["Warden II"] := true
            names["Warp Scrambling Drone"] := true
            names["Warrior I"] := true
            names["Warrior II"] := true
            names["Warrior SW-300"] := true
            names["Warrior TP-300"] := true
            names["Wasp EC-900"] := true
            names["Wasp I"] := true
            names["Wasp II"] := true
        }
        return names.Has(name)
    }

    FormatView(totals, peaks, residue, names, mining, alpha, combatIndexed, combined := 0, byCharacter := 0) {
        result := "D`t" this.Join(totals) "`t" peaks[1] "`t" peaks[2] "`t" residue
            . "`t" alpha[1] "`t" alpha[2] "`t" combatIndexed "`n"
        for name in names
            result .= "P`t" name "`n"
        for material, amount in mining
            result .= "M`t" material "`t" amount[1] "`t" amount[2] "`n"
        if combined {
            result .= "E`t" combined.extendedIndexed "`n"
            for key, value in combined.extra
                result .= "X`t" key "`t" value "`n"
            for field in ["modules", "systems", "bursts"] {
                for name, count in combined.%field% {
                    leader := "", best := 0
                    for character, day in byCharacter {
                        value := day.%field%.Get(name, 0)
                        if value > best
                            best := value, leader := character
                    }
                    result .= (field = "modules" ? (this.IsDrone(name) ? "R" : "U") : field = "bursts" ? "B" : "Y") "`t" name "`t" count "`t" leader "`n"
                }
            }
            leaders := Map(), bests := Map()
            for character, day in byCharacter {
                if character = ""
                    continue
                loop StatisticsArchive.MetricCount
                    this.PickLeader(leaders, bests, "count" A_Index, character, day.values[A_Index])
                for key, value in day.extra
                    this.PickLeader(leaders, bests, key, character, value)
                gas := 0, solid := 0, residue := 0
                for material, amounts in day.mining {
                    if RegExMatch(material, "i)^(Fullerite-)|\b(?:Cytoserocin|Mykoserocin)\b")
                        gas += amounts[1]
                    else
                        solid += amounts[1]
                    residue += amounts[2]
                    this.PickLeader(leaders, bests, "material:" material, character, amounts[1])
                }
                for key, value in Map("gas", gas, "solid", solid, "residue", residue, "extraResidue", day.extraResidue)
                    this.PickLeader(leaders, bests, key, character, value)
            }
            for key in ["damageIn", "damageOut", "missIn", "missOut"] {
                best := 0, peer := ""
                for contact, value in combined.contacts
                    if InStr(contact, key ":") = 1 && value > best
                        best := value, peer := SubStr(contact, StrLen(key) + 2)
                result .= "C`t" key "`t" peer "`n"
            }
            for field in ["peak", "alpha"]
                loop 2 {
                    peers := field "Peers"
                    result .= "C`t" field A_Index "`t" combined.%peers%[A_Index] "`n"
                }
            for key, character in leaders
                result .= "T`t" key "`t" character "`n"
            for field in ["peak", "alpha"] {
                loop 2 {
                    index := A_Index, best := 0, leader := ""
                    for character, day in byCharacter
                        if day.%field%[index] > best
                            best := day.%field%[index], leader := character
                    result .= "T`t" field index "`t" leader "`n"
                }
            }
        }
        return result
    }

    ; The worker publishes a compact read model. Filter requests only read
    ; character/day buckets, never the per-file journal while a scan writes it.
    BuildDashboard(allowStop := false) {
        this.Phase("Preparing dashboard totals…")
        characters := Map(), latest := Map(), legacy := this.EmptyMetrics()
        characters.CaseSense := latest.CaseSense := "Off"
        for key, entry in this.files {
            if Mod(A_Index, 1024) = 0 {
                if allowStop && this.ShouldStop()
                    return
                this.Phase("Preparing dashboard totals…", false)
            }
            if !entry.HasOwnProp("detail") {
                loop StatisticsArchive.MetricCount
                    legacy[A_Index] += entry.values[A_Index]
                continue
            }
            name := entry.character
            if !characters.Has(name)
                characters[name] := Map()
            dates := characters[name]
            start := SubStr(entry.session, 1, 8)
            if start = ""
                start := "00000000"
            if !dates.Has(start)
                dates[start] := this.NewDay()
            dates[start].values[1] += entry.values[1]
            dates[start].values[2] += entry.values[2]
            for date, day in this.DecodeDays(entry.detail) {
                if !dates.Has(date)
                    dates[date] := day
                else
                    this.MergeDay(dates[date], day)
            }
            if InStr(key, "game:") = 1 && (!latest.Has(name) || entry.session > latest[name].session)
                latest[name] := entry
        }
        text := "L`t" this.Join(legacy) "`n"
        for name, dates in characters {
            for date, day in dates
                text .= this.DashboardRow("D", name, date, day)
            if latest.Has(name) {
                session := this.NewDay()
                session.values[1] := 1
                for , day in this.DecodeDays(latest[name].detail)
                    this.MergeDay(session, day)
                text .= this.DashboardRow("S", name, "00000000", session)
            }
        }
        this.WriteAtomic(this.dashboard, text)
        this.changed := false
    }

    NeedsDashboard() {
        info := Buffer(36, 0)
        if !DllCall("GetFileAttributesExW", "Str", this.dashboard, "Int", 0, "Ptr", info)
            return true
        published := NumGet(info, 20, "Int64")
        return DllCall("GetFileAttributesExW", "Str", this.path, "Int", 0, "Ptr", info)
            && NumGet(info, 20, "Int64") > published
    }

    DashboardRow(kind, name, date, day) => kind "`t" name "`t" date "`t" day.values[1] "`t" day.values[2]
        . "`t" this.EncodeDays(Map(date, day)) "`n"

    MergeDay(target, day) {
        loop StatisticsArchive.MetricCount
            target.values[A_Index] += day.values[A_Index]
        for field in ["peak", "alpha"]
            loop 2
                if day.%field%[A_Index] > target.%field%[A_Index] {
                    target.%field%[A_Index] := day.%field%[A_Index]
                    peers := field "Peers"
                    target.%peers%[A_Index] := day.%peers%[A_Index]
                }
        target.extendedIndexed := target.extendedIndexed && day.extendedIndexed
        target.combatIndexed := target.combatIndexed && day.combatIndexed
        target.extraResidue += day.extraResidue
        for field in ["extra", "modules", "systems", "bursts", "contacts"]
            for name, count in day.%field%
                if field = "extra" && name = "burstMaxRecipients"
                    target.extra[name] := Max(target.extra.Get(name, 0), count)
                else
                    this.Bump(target.%field%, name, count)
        for material, count in day.mining {
            if !target.mining.Has(material)
                target.mining[material] := [0, 0]
            target.mining[material][1] += count[1]
            target.mining[material][2] += count[2]
        }
    }

    static ViewDashboard(storage, character, range) {
        reader := StatisticsArchive(storage, "", false)
        names := Map(), combined := reader.NewDay(), byCharacter := Map()
        names.CaseSense := "Off"
        kind := range = "session" ? "S" : "D"
        offsets := Map("day", 0, "week", -6, "month", -29, "year", -364)
        since := offsets.Has(range) ? SubStr(DateAdd(A_NowUTC, offsets[range], "Days"), 1, 8) : ""
        file := reader.OpenShared(reader.dashboard, 0x700)
        try text := file.Read()
        finally file.Close()
        Loop Parse, text, "`n", "`r" {
            fields := StrSplit(A_LoopField, "`t", , 6)
            if fields.Length < 2
                continue
            if fields[1] = "L" {
                if range = "all" && character = "" {
                    legacy := StrSplit(SubStr(A_LoopField, 3), "`t")
                    loop StatisticsArchive.MetricCount
                        combined.values[A_Index] += Number(legacy[A_Index])
                    combined.combatIndexed := combined.combatIndexed && !(Number(legacy[4]) || Number(legacy[5]))
                }
                continue
            }
            if fields.Length != 6
                continue
            name := fields[2], date := fields[3]
            if name != ""
                names[name] := true
            if fields[1] != kind || (character != "" && name != character) || (since && date < since)
                continue
            day := reader.DecodeDays(fields[6])[date]
            day.values[1] := Number(fields[4])
            day.values[2] := Number(fields[5])
            reader.MergeDay(combined, day)
            if !byCharacter.Has(name)
                byCharacter[name] := reader.NewDay()
            reader.MergeDay(byCharacter[name], day)
        }
        return reader.FormatView(combined.values, combined.peak, combined.extraResidue, names, combined.mining,
            combined.alpha, combined.combatIndexed, combined, byCharacter)
    }

    CompactIfNeeded() {
        if this.lines > this.files.Count * 2 + 256
            this.Compact()
    }

    Compact() {
        text := ""
        for key, entry in this.files
            text .= this.Row(key, entry)
        this.WriteAtomic(this.path, text)
        this.lines := this.files.Count
    }

    WriteAtomic(path, text) {
        temporary := path ".tmp"
        if FileExist(temporary)
            FileDelete(temporary)
        FileAppend(text, temporary, "UTF-8-RAW")
        retries := MoveWorkerFile(temporary, path)
        if retries
            StatisticsArchive.Log(this.storage, "Recovered a busy checkpoint file; retries=" retries)
    }
}

class StatisticsDamageWindow {
    seconds := []
    totals := [0, 0]

    Add(second, direction, damage, peer := "") {
        buckets := this.seconds
        if buckets.Length && second < buckets[-1][1] {
            buckets.Length := 0
            this.totals := [0, 0]
        }
        while buckets.Length && buckets[1][1] <= second - 10 {
            expired := buckets.RemoveAt(1)
            this.totals[1] -= expired[2], this.totals[2] -= expired[3]
        }
        if !buckets.Length || buckets[-1][1] != second
            buckets.Push([second, 0, 0, Map(), Map()])
        if peer != ""
            buckets[-1][direction + 3][peer] := buckets[-1][direction + 3].Get(peer, 0) + damage
        buckets[-1][direction + 1] += damage
        this.totals[direction] += damage
        return this.totals[direction] / 10
    }

    TopPeer(direction) {
        totals := Map(), best := 0, peer := ""
        for bucket in this.seconds
            for name, amount in bucket[direction + 3] {
                totals[name] := totals.Get(name, 0) + amount
                if totals[name] > best
                    best := totals[name], peer := name
            }
        return peer
    }

    Serialize() {
        if !this.seconds.Length
            return "-"
        text := ""
        for bucket in this.seconds {
            text .= (text = "" ? "" : ";") bucket[1] "," bucket[2] "," bucket[3]
            loop 2 {
                peers := ""
                for name, damage in bucket[A_Index + 3] {
                    encoded := ""
                    loop parse name
                        encoded .= (encoded = "" ? "" : ".") Ord(A_LoopField)
                    peers .= (peers = "" ? "" : "/") encoded ":" damage
                }
                text .= "," peers
            }
        }
        return text
    }

    Restore(text) {
        if InStr(text, "old:") = 1
            text := SubStr(text, 5)
        if text = "-" || text = ""
            return
        for row in StrSplit(text, ";") {
            fields := StrSplit(row, ",")
            second := Number(fields[1])
            this.Add(second, 1, Number(fields[2]))
            this.Add(second, 2, Number(fields[3]))
            if fields.Length >= 5
                loop 2 {
                    direction := A_Index
                    for item in StrSplit(fields[direction + 3], "/") {
                        if item = ""
                            continue
                        parts := StrSplit(item, ":"), name := ""
                        for code in StrSplit(parts[1], ".")
                            name .= Chr(Number(code))
                        this.seconds[-1][direction + 3][name] := Number(parts[2])
                    }
                }
        }
    }
}

; Keep the maps small when an installation has hundreds of thousands of
; checkpoints. AutoHotkey maps insert into an ordered array of keys.
class StatisticsFileIndex {
    Count := 0

    __New(ignoreCase := false) {
        this.ignoreCase := ignoreCase
        this.buckets := []
        loop 256 {
            bucket := Map()
            if ignoreCase
                bucket.CaseSense := "Off"
            this.buckets.Push(bucket)
        }
    }

    Bucket(key) {
        hashKey := this.ignoreCase ? StrLower(key) : key
        hash := DllCall("ntdll\RtlComputeCrc32", "UInt", 0, "Ptr", StrPtr(hashKey), "UInt", StrLen(hashKey) * 2, "UInt")
        return this.buckets[(hash & 255) + 1]
    }

    __Item[key] {
        get => this.Bucket(key)[key]
        set {
            bucket := this.Bucket(key)
            oldCount := bucket.Count
            bucket[key] := value
            this.Count += bucket.Count - oldCount
        }
    }

    Get(key, fallback) => this.Bucket(key).Get(key, fallback)

    Delete(key) {
        this.Bucket(key).Delete(key)
        this.Count -= 1
    }

    __Enum(*) {
        index := 0, enumerator := 0
        return Next
        Next(&key, &value) {
            loop {
                if enumerator && enumerator(&key, &value)
                    return true
                index += 1
                if index > this.buckets.Length
                    return false
                enumerator := this.buckets[index].__Enum(2)
            }
        }
    }
}
