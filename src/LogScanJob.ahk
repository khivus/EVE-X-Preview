; Archive I/O runs outside the UI interpreter; the parent only polls small results.
class LogScanJob {
    static serial := 0
    __New(workerScript := "") {
        this.workerScript := workerScript
        this.directory := A_Temp "\EVE-X-Preview-LogScan-" DllCall("GetCurrentProcessId") "-" A_TickCount "-" (++LogScanJob.serial)
        DirCreate(this.directory)
        this.process := 0
        this.busy := false
        this.started := 0
        this.progress := ""
        this.failedAt := -10000
        this.exitHandler := ObjBindMethod(this, "Close")
        OnExit(this.exitHandler)
    }

    Start(request) {
        if this.busy || A_TickCount - this.failedAt < 10000
            return false
        try {
            for name in ["input.txt", "output.txt", "output.tmp", "error.txt", "cache.tmp", "progress.txt", "progress.tmp"]
                if FileExist(this.directory "\" name)
                    FileDelete(this.directory "\" name)
            FileAppend(request, this.directory "\input.txt", "UTF-8")
            runtime := A_IsCompiled ? A_ScriptFullPath : A_AhkPath
            if this.workerScript != ""
                script := this.workerScript
            else if A_IsCompiled
                script := "*LogScanWorker"
            else {
                SplitPath(A_LineFile, , &sourceDirectory)
                script := sourceDirectory "\LogScanWorker.ahk"
            }
            command := '"' runtime '" /script /ErrorStdOut "' script '" "' this.directory '" ' DllCall("GetCurrentProcessId")
            Run(command, , "Hide", &pid)
            this.process := DllCall("OpenProcess", "UInt", 0x100001, "Int", false, "UInt", pid, "Ptr")
            if !this.process
                throw OSError()
            this.progress := ""
            this.busy := true
            this.started := A_TickCount
            ProgramLog.Add("Log scan worker started; type=" StrSplit(request, "`n")[1])
            return true
        } catch as err {
            this.failedAt := A_TickCount
            ProgramLog.Add("Log scan worker launch failed: " err.Message, "WARN")
            return false
        }
    }

    Poll() {
        if !this.busy
            return 0
        status := DllCall("WaitForSingleObject", "Ptr", this.process, "UInt", 0, "UInt")
        if status = 0x102 && A_TickCount - this.started < 300000 {
            try {
                if FileExist(this.directory "\progress.txt") {
                    serial := FileRead(this.directory "\progress.txt", "UTF-8")
                    if serial != this.progress {
                        result := this.ReadResult()
                        this.progress := serial
                        return result ; Publish early matches while the scan continues.
                    }
                }
            }
            return 0
        }
        result := Map("ok", false, "text", "")
        try {
            if status = 0 && FileExist(this.directory "\output.txt") {
                result := this.ReadResult()
                ProgramLog.Add("Log scan worker completed; elapsed=" (A_TickCount - this.started) "ms")
            } else {
                detail := FileExist(this.directory "\error.txt") ? SubStr(FileRead(this.directory "\error.txt", "UTF-8"), 1, 2048) : "worker exited or timed out without a result"
                ProgramLog.Add("Log scan worker failed: " detail, "WARN")
                this.failedAt := A_TickCount
            }
        } catch as err {
            this.failedAt := A_TickCount
            ProgramLog.Add("Log scan result failed: " err.Message, "WARN")
        } finally {
            this.StopProcess()
        }
        return result
    }

    ReadResult() {
        if FileGetSize(this.directory "\output.txt") > 4194304
            throw Error("Log discovery result exceeded its size limit")
        return Map("ok", true, "text", FileRead(this.directory "\output.txt", "UTF-8"))
    }

    StopProcess() {
        if this.process {
            if DllCall("WaitForSingleObject", "Ptr", this.process, "UInt", 0, "UInt") = 0x102
                DllCall("TerminateProcess", "Ptr", this.process, "UInt", 1)
            DllCall("CloseHandle", "Ptr", this.process)
            this.process := 0
        }
        this.busy := false
    }

    Close(*) {
        this.StopProcess()
        for name in ["input.txt", "output.txt", "output.tmp", "error.txt", "cache.txt", "cache.tmp", "progress.txt", "progress.tmp"]
            try FileDelete(this.directory "\" name)
        try DirDelete(this.directory) ; Only remove our empty, private job directory.
        OnExit(this.exitHandler, 0)
    }
}
