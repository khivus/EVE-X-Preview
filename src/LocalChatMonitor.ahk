; Local chat has a different header and lifecycle from combat logs. Keep its
; file cursors separate so event suppression never discards system changes.
class LocalChatMonitor {
    __New(owner) {
        this.owner := owner
        this.readers := Map()
        this.headers := Map()
        this.lastDiscovery := -5000
    }

    Poll(active) {
        removed := []
        for character, reader in this.readers {
            if !active.Has(character) || active[character] != reader.hwnd
                removed.Push(character)
        }
        for character in removed {
            this.owner.updateThumbnailSystemText("", this.readers[character].hwnd)
            this.readers[character].file.Close()
            this.readers.Delete(character)
        }

        ; Rescan even quiet logs: a new session can replace one that never grows.
        if A_TickCount - this.lastDiscovery >= 5000 {
            this.lastDiscovery := A_TickCount
            paths := this.FindLogs(this.owner.ResolveChatLogsDirectory(), active)
            for character, path in paths {
                if this.readers.Has(character) && this.readers[character].path = path
                    continue
                try file := FileOpen(path, "r", "UTF-8") ; BOM also detects UTF-16 chat logs.
                catch
                    continue
                if !file
                    continue
                if this.readers.Has(character)
                    this.readers[character].file.Close()
                this.readers[character] := {path: path, file: file, pending: "", system: "", hwnd: active[character]}
            }
        }

        failed := []
        for character, reader in this.readers {
            try {
                if reader.file.Length < reader.file.Pos {
                    reader.file.Close()
                    reader.file := FileOpen(reader.path, "r", "UTF-8")
                    reader.pending := ""
                    reader.system := ""
                }
                this.ReadUpdates(reader)
                this.owner.updateThumbnailSystemText(reader.system, reader.hwnd)
            } catch {
                try reader.file.Close()
                failed.Push(character)
                this.owner.updateThumbnailSystemText("", reader.hwnd)
            }
        }
        for character in failed
            this.readers.Delete(character)
    }

    FindLogs(directory, active) {
        result := Map()
        if !active.Count || !DirExist(directory)
            return result
        sessions := Map()
        seenHeaders := Map()
        Loop Files, directory "\Local_*.txt", "F" {
            if !RegExMatch(A_LoopFileName, "i)^Local_\d{8}_\d{6}_\d+\.txt$")
                continue
            try {
                path := A_LoopFileFullPath
                cacheKey := path "|" A_LoopFileTimeCreated
                if !this.headers.Has(cacheKey) {
                    header := this.ReadHeader(path)
                    if header.listener = "" || header.session = ""
                        continue ; Incomplete header: retry on the next scan.
                    this.headers[cacheKey] := header
                }
                header := this.headers[cacheKey]
                seenHeaders[cacheKey] := header
                character := this.owner.AntiCleanTitle(header.listener)
                ; Use the session header, not modification time: old chat
                ; activity and the filename's numeric suffix are not reliable.
                if active.Has(character) && (!sessions.Has(character) || header.session > sessions[character]) {
                    sessions[character] := header.session
                    result[character] := path
                }
            }
        }
        this.headers := seenHeaders
        return result
    }

    ReadHeader(path) {
        header := {listener: "", session: ""}
        file := FileOpen(path, "r", "UTF-8")
        if !file
            return header
        try {
            Loop 32 {
                if file.AtEOF
                    break
                line := file.ReadLine()
                if RegExMatch(line, "^\s*Listener:\s*(.+?)\s*$", &match)
                    header.listener := match[1]
                else if RegExMatch(line, "^\s*Session started:\s*(\d{4}\.\d{2}\.\d{2} \d{2}:\d{2}:\d{2})", &match)
                    header.session := RegExReplace(match[1], "\D")
                if header.listener != "" && header.session != ""
                    break
            }
        } finally {
            file.Close()
        }
        return header
    }

    ReadUpdates(reader) {
        if reader.file.AtEOF
            return
        lines := StrSplit(reader.pending reader.file.Read(), "`n", "`r")
        reader.pending := lines.Pop() ; Wait for a full line before changing systems.
        for line in lines {
            if RegExMatch(line, "^\x{FEFF}?\s*\[\s*\d{4}\.\d{2}\.\d{2} \d{2}:\d{2}:\d{2}\s*\] EVE System > Channel changed to Local :\s*(.+?)\s*$", &match)
                reader.system := match[1]
        }
    }

    Close(*) {
        for character, reader in this.readers
            reader.file.Close()
        this.readers.Clear()
    }
}
