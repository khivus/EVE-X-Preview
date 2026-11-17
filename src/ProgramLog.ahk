; Bounded in-memory event history plus rotating, buffered on-disk diagnostics.
class ProgramLog {
    static Text := ""
    static Pending := ""
    static Directory := ""
    static LastErrorExport := ""
    static LastErrorTick := -60000
    static Busy := false
    static MaxChars := 262144
    static MemoryCooldownMs := 30000
    static MemoryRetryAt := 0
    static MemoryNoticePending := false
    static Activity := Map()
    static SlowOperationAt := Map()
    static ActivityAt := A_TickCount
    static DisplayNotificationAt := Map()

    static Count(event) {
        if this.IsMemoryBackoff()
            return
        this.Activity[event] := this.Activity.Get(event, 0) + 1
    }

    static SlowOperation(name, started) {
        if this.IsMemoryBackoff()
            return
        elapsed := A_TickCount - started
        if elapsed < 100 || A_TickCount - this.SlowOperationAt.Get(name, -10000) < 10000
            return
        this.SlowOperationAt[name] := A_TickCount
        this.Add("Slow operation: " name "; elapsed=" elapsed "ms", "WARN")
    }

    static ReportActivity(now := A_TickCount) {
        if this.IsMemoryBackoff()
            return
        if now - this.ActivityAt < 10000
            return
        elapsed := now - this.ActivityAt
        this.ActivityAt := now
        if !this.Activity.Count
            return
        counts := this.Activity
        this.Activity := Map()
        message := "Runtime activity over " elapsed "ms:"
        for event, count in counts
            message .= " " event "=" count ";"
        process := DllCall("GetCurrentProcess", "Ptr")
        message .= " GDI=" DllCall("GetGuiResources", "Ptr", process, "UInt", 0)
            . " USER=" DllCall("GetGuiResources", "Ptr", process, "UInt", 1)
        this.Add(message)
    }

    static DisplayChanged(wParam, lParam, msg, hwnd) {
        if this.IsMemoryBackoff()
            return
        ; Broadcasts reach each preview/overlay. Log the burst once, not per HWND.
        notification := msg ":" wParam ":" lParam
        if A_TickCount - this.DisplayNotificationAt.Get(notification, -1000) < 1000
            return
        if this.DisplayNotificationAt.Count >= 64
            this.DisplayNotificationAt.Clear()
        this.DisplayNotificationAt[notification] := A_TickCount
        this.Add("Windows display notification=" Format("0x{:X}", msg) "; bpp=" wParam
            "; size=" (lParam & 0xFFFF) "x" ((lParam >> 16) & 0xFFFF), "WARN")
        this.Flush()
    }

    static Init(directory := "") {
        this.Directory := directory != "" ? directory : EnvGet("LOCALAPPDATA") "\EVE-X-Preview\Logs"
        try DirCreate(this.Directory)
        this.Add("Session started; Windows=" A_OSVersion "; AHK=" A_AhkVersion "; bits=" (A_PtrSize * 8) "; compiled=" A_IsCompiled)
        try {
            version := A_IsCompiled ? FileGetVersion(A_ScriptFullPath) : (RegExMatch(FileRead(A_ScriptFullPath), "U_version = ([\d.]+)", &m) ? m[1] : "source")
            this.Add("Version=" version "; monitors=" MonitorGetCount() "; screen=" A_ScreenWidth "x" A_ScreenHeight "; DPI=" A_ScreenDPI)
        }
        this.Flush()
        this.FlushTimer := ObjBindMethod(this, "Flush")
        SetTimer(this.FlushTimer, 2000)
        this.ExitHandler := ObjBindMethod(this, "Exit")
        OnExit(this.ExitHandler)
    }

    static Add(message, level := "INFO") {
        if this.IsMemoryBackoff()
            return
        try {
            message := Trim(SubStr(message, -this.MaxChars), "`r`n ")
            if message = ""
                return
            if this.MemoryNoticePending {
                notice := "[WARN] Out of memory: diagnostic buffers were discarded; logging resumed after a 30-second pause.`r`n"
                this.Text := this.AppendTail(this.Text, notice)
                this.Pending := this.AppendTail(this.Pending, notice)
                this.MemoryNoticePending := false
            }
            line := FormatTime(, "yyyy-MM-dd HH:mm:ss") " +" A_TickCount " [" level "] " message "`r`n"
            this.Text := this.AppendTail(this.Text, line)
            this.Pending := this.AppendTail(this.Pending, line)
        } catch MemoryError {
            this.RecoverMemory()
        }
    }

    static AppendTail(buffer, addition) {
        ; Trim before concatenating, so even temporary joins stay within the cap.
        addition := SubStr(addition, -this.MaxChars)
        remaining := this.MaxChars - StrLen(addition)
        return (remaining > 0 ? SubStr(buffer, -remaining) : "") addition
    }

    static IsMemoryBackoff() => A_TickCount < this.MemoryRetryAt

    static RecoverMemory() {
        ; No stack formatting, exports, GUI or recursive logging while exhausted.
        this.MemoryRetryAt := A_TickCount + this.MemoryCooldownMs
        this.MemoryNoticePending := true
        this.Text := ""
        this.Pending := ""
        this.Activity.Clear()
        this.SlowOperationAt.Clear()
        this.DisplayNotificationAt.Clear()
    }

    static Flush(*) {
        if this.IsMemoryBackoff() || this.Busy
            return
        this.Busy := true
        started := A_TickCount
        try {
            this.ReportActivity()
            if this.IsMemoryBackoff() || this.Directory = "" || this.Pending = ""
                return
            path := this.Directory "\current.log"
            if FileExist(path) && FileGetSize(path) > 1048576
                FileMove(path, this.Directory "\previous.log", 1)
            batch := this.Pending
            this.Pending := ""
            try this.WriteBatch(batch, path)
            catch MemoryError as err {
                throw err
            }
            catch {
                this.Pending := this.AppendTail(batch, this.Pending)
            }
        } catch MemoryError {
            this.RecoverMemory()
        } catch {
            ; Diagnostics must never recursively fail the application.
        } finally {
            this.Busy := false
            try this.SlowOperation("Diagnostic log flush", started)
            catch MemoryError {
                this.RecoverMemory()
            }
        }
    }

    static WriteBatch(batch, path) => FileAppend(batch, path, "UTF-8")

    static Error(err, context := "Unhandled error") {
        if err is MemoryError {
            this.RecoverMemory()
            return
        }
        if this.IsMemoryBackoff()
            return
        try {
            details := context
            try details .= ": " err.Message "`n" err.File ":" err.Line " " err.What "`n" err.Stack
            catch MemoryError as memoryFailure {
                throw memoryFailure
            }
            catch {
            }
            this.Add(details, "ERROR")
            if this.IsMemoryBackoff()
                return
            this.Flush()
            if this.IsMemoryBackoff()
                return
            ; Preserve every error in current.log, but avoid exporting a file storm.
            if this.Directory != "" && A_TickCount - this.LastErrorTick >= 10000 {
                this.LastErrorTick := A_TickCount
                try this.LastErrorExport := this.Export(this.Directory "\error-" A_Now "-" A_TickCount ".log")
                catch MemoryError as memoryFailure {
                    throw memoryFailure
                }
                catch {
                }
                ; Keep at most ten automatic error snapshots.
                files := ""
                loop files this.Directory "\error-*.log"
                    files .= A_LoopFileName "`n"
                names := StrSplit(Trim(Sort(files), "`n"), "`n")
                loop Max(0, names.Length - 10)
                    try FileDelete(this.Directory "\" names[A_Index])
            }
        } catch MemoryError {
            this.RecoverMemory()
        }
    }

    static Export(path) {
        this.Flush()
        stream := FileOpen(path, "w", "UTF-8")
        if !stream
            throw Error("Cannot write diagnostic log",, path)
        try stream.Write(this.Text)
        finally stream.Close()
        return path
    }

    static Show(*) {
        this.ShowOwned(0)
    }

    static ShowOwned(owner) {
        viewer := this.CreateViewer()
        if owner {
            viewer.Opt("+Owner" owner)
            work := SettingsWindowSizing.WorkArea(owner)
            fit := SettingsWindowSizing.Fit(work.left, work.top, work.right, work.bottom, A_ScreenDPI)
            SettingsWindowSizing.ShowCentered(viewer, owner, Min(880, fit.width), Min(520, fit.height))
        } else
            viewer.Show()
    }

    static CreateViewer() {
        viewer := Gui("+Resize", "EVE-X-Preview - Program Log")
        try {
            viewer.SetFont("s9", "Consolas")
            ; Large text can fail during native control creation. Create an empty
            ; multiline edit first, then load the snapshot via WM_SETTEXT.
            logControl := viewer.Add("Edit", "vLog Multi ReadOnly -Wrap HScroll w860 h500")
            logControl.Value := this.Text
            viewer.OnEvent("Size", (g, state, w, h) => state != -1 ? g["Log"].Move(8, 8, Max(1, w - 16), Max(1, h - 16)) : 0)
            viewer.OnEvent("Close", (g, *) => g.Destroy())
            return viewer
        } catch as err {
            viewer.Destroy()
            throw err
        }
    }

    static ExportDialog(*) {
        path := FileSelect("S16", "EVE-X-Preview-" A_Now ".log", "Export program log", "Logs (*.log)")
        if path = ""
            return
        try {
            this.Export(path)
            MsgBox("Log exported to:`n" path)
        } catch as err {
            MsgBox("Could not export log:`n" err.Message)
        }
    }

    static Exit(reason, code) {
        this.Add("Session ending: " reason "; code=" code)
        this.Flush()
    }
}
