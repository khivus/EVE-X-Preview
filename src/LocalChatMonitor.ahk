; Discovery and startup system lookup run in a worker; live reads stay bounded.
class LocalChatMonitor {
    __New(owner) {
        this.owner := owner
        this.readers := Map()
        this.lastDiscovery := -5000
        this.busy := false
    }

    Poll(active) {
        if this.busy
            return
        this.busy := true
        try this.PollReaders(active)
        finally this.busy := false
    }

    PollReaders(active) {
        for character, reader in this.readers.Clone() {
            if !active.Has(character) || active[character] != reader.hwnd {
                this.owner.updateThumbnailSystemText("", reader.hwnd)
                reader.file.Close()
                this.readers.Delete(character)
            }
        }
        if !active.Count {
            if this.HasOwnProp("discovery") {
                this.discovery.Close()
                this.DeleteProp("discovery")
            }
            return
        }
        paths := this.FindLogs(this.owner.ResolveChatLogsDirectory(), active)
        for character, path in paths {
            seed := this.discovery.byId[this.owner.CleanTitle(character)]
            previous := this.readers.Get(character, 0)
            if previous && previous.path = path && previous.generation = this.discovery.generation
                continue
            try file := FileOpen(path, "r", "UTF-8")
            catch
                continue
            if !file
                continue
            identity := this.owner.GameLogFileIdentity(file)
            if previous && previous.path = path && identity != "" && identity = previous.identity {
                if previous.HasOwnProp("needsSeed") && previous.needsSeed {
                    previous.system := seed.system
                    previous.needsSeed := false
                }
                previous.generation := this.discovery.generation
                file.Close()
                continue
            }
            if seed.identity != "" && identity != seed.identity {
                file.Close()
                this.lastDiscovery := -5000 ; Replaced after the worker's snapshot.
                continue
            }
            if previous
                previous.file.Close()
            reader := {path: path, file: file, pending: "", system: seed.system, hwnd: active[character],
                identity: identity, generation: this.discovery.generation, dropFragment: false}
            ; Re-read a small tail to retain lines completed after the worker's snapshot.
            this.SeekRecentTail(reader, Min(seed.cursor, file.Length))
            this.readers[character] := reader
        }
        for character, reader in this.readers.Clone() {
            try {
                if reader.file.Length < reader.file.Pos {
                    reader.file.Seek(0)
                    reader.pending := "", reader.system := "", reader.dropFragment := false
                }
                this.ReadUpdates(reader)
                this.owner.updateThumbnailSystemText(reader.system, reader.hwnd)
            } catch {
                try reader.file.Close()
                this.readers.Delete(character)
                this.owner.updateThumbnailSystemText("", reader.hwnd)
            }
        }
    }

    FindLogs(directory, active) {
        if !this.HasOwnProp("discovery") || this.directory != directory {
            if this.HasOwnProp("discovery")
                this.discovery.Close()
            this.directory := directory
            this.discovery := LogFileDiscovery(directory, "chat")
        }
        characters := ""
        for character in active
            characters .= this.owner.CleanTitle(character) "`n"
        characters := Sort(characters) ; Window Z-order changes must not restart a scan.
        this.discovery.Refresh(this.lastDiscovery < 0, characters)
        this.lastDiscovery := A_TickCount
        result := Map()
        for character in active {
            name := this.owner.CleanTitle(character)
            if this.discovery.byId.Has(name)
                result[character] := this.discovery.byId[name].path
        }
        return result
    }

    SeekRecentTail(reader, ending) {
        offset := Max(0, ending - 65536)
        if InStr(reader.file.Encoding, "UTF-16")
            offset -= Mod(offset, 2)
        reader.file.Seek(offset)
        reader.pending := ""
        reader.dropFragment := offset > 0
    }

    ReadUpdates(reader) {
        if reader.file.AtEOF
            return
        if reader.file.Length - reader.file.Pos > 262144 {
            this.SeekRecentTail(reader, reader.file.Length)
            reader.needsSeed := true
            this.lastDiscovery := -5000
        }
        text := reader.pending reader.file.Read(16384)
        lines := StrSplit(text, "`n", "`r")
        reader.pending := lines.Pop()
        for line in lines {
            if reader.dropFragment {
                reader.dropFragment := false
                continue
            }
            if RegExMatch(line, "^\x{FEFF}?\s*\[\s*\d{4}\.\d{2}\.\d{2} \d{2}:\d{2}:\d{2}\s*\] EVE System > Channel changed to Local :\s*(.+?)\s*$", &match)
                reader.system := match[1], reader.needsSeed := false
        }
        if StrLen(reader.pending) > 8192 {
            reader.pending := ""
            reader.dropFragment := true ; Discard oversized/non-system records in chunks.
        }
    }

    Close(*) {
        if this.HasOwnProp("discovery")
            this.discovery.Close()
        for character, reader in this.readers
            reader.file.Close()
        this.readers.Clear()
    }
}
