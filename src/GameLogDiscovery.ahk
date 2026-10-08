; Shared asynchronous discovery for game sessions and Local chat listeners.
class LogFileDiscovery {
    __New(directory, mode := "game") {
        fullPath := Buffer(65536, 0)
        DllCall("GetFullPathNameW", "Str", directory, "UInt", 32768, "Ptr", fullPath, "Ptr", 0)
        this.directory := StrGet(fullPath, "UTF-16")
        this.mode := mode
        this.lastRequest := ""
        this.retryHeaders := false
        this.files := []
        this.byId := Map()
        this.headers := Map()
        this.generation := 0
        this.lastScan := -60000
        this.busy := false
        this.watch := 0
        this.job := LogScanJob()
        this.OpenWatch()
    }

    OpenWatch() {
        if !DirExist(this.directory)
            return
        handle := DllCall("FindFirstChangeNotificationW", "Str", this.directory, "Int", false, "UInt", 1, "Ptr")
        if handle && handle != -1
            this.watch := handle
    }

    Refresh(force := false, characters := "") {
        request := this.mode "`n" this.directory "`n" characters
        if this.job.busy && this.mode = "chat" && (force || request != this.lastRequest)
            this.job.StopProcess()
        result := this.job.Poll()
        if result && result["ok"] {
            newest := Map(), sorted := []
            this.retryHeaders := false
            for row in StrSplit(result["text"], "`n", "`r") {
                fields := StrSplit(row, "`t")
                if fields.Length != (this.mode = "chat" ? 6 : 4)
                    continue
                entry := {path: fields[2], session: fields[3]}
                if this.mode = "chat" {
                    entry.system := fields[4]
                    entry.cursor := Integer(fields[5])
                    entry.identity := fields[6]
                }
                if this.mode = "game" {
                    entry.listener := fields[4]
                    if entry.listener = ""
                        this.retryHeaders := true
                }
                newest[fields[1]] := entry
                sorted.Push(entry)
            }
            if sorted.Length > 1
                Main_Class.Prototype.QuickSort(sorted, (a, b) => -StrCompare(a.session, b.session))
            files := [], headers := Map()
            for entry in sorted {
                files.Push(entry.path)
                if this.mode = "game" && entry.listener != ""
                    headers[entry.path] := Main_Class.Prototype.AntiCleanTitle(entry.listener)
            }
            this.headers := headers
            this.byId := newest
            this.files := files
            this.lastScan := A_TickCount
            this.generation++
        }
        this.busy := this.job.busy
        if this.busy
            return this.files ; Existing readers continue while the worker scans.
        dirty := false
        if this.watch {
            status := DllCall("WaitForSingleObject", "Ptr", this.watch, "UInt", 0, "UInt")
            if status = 0 {
                dirty := true
                if !DllCall("FindNextChangeNotification", "Ptr", this.watch)
                    this.CloseWatch()
            } else if status != 0x102
                this.CloseWatch()
        }
        interval := this.watch && !this.retryHeaders ? 60000 : 5000
        if force || request != this.lastRequest || !this.generation || dirty || A_TickCount - this.lastScan >= interval {
            if !this.watch
                this.OpenWatch()
            this.busy := this.job.Start(request)
            if this.busy
                this.lastRequest := request
        }
        return this.files
    }

    CloseWatch() {
        if this.watch {
            DllCall("FindCloseChangeNotification", "Ptr", this.watch)
            this.watch := 0
        }
    }

    Close(*) {
        this.CloseWatch()
        this.job.Close()
        this.busy := false
    }
}

class GameLogDiscovery extends LogFileDiscovery {
    __New(directory) {
        super.__New(directory, "game")
    }
}
