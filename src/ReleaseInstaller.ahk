; Stage separate release EXEs and replace only installed components together.
; The embedded transaction preserves recovery without adding release artifacts.
class ReleaseInstaller {
    static RecoveryUpdater(path) {
        ; Inspect version resources instead of executing an unknown old updater.
        size := DllCall("version\GetFileVersionInfoSizeW", "Str", path, "Ptr", 0, "UInt")
        if !size
            return false
        data := Buffer(size)
        if !DllCall("version\GetFileVersionInfoW", "Str", path, "UInt", 0, "UInt", size, "Ptr", data)
            return false
        pointer := 0, length := 0
        if !DllCall("version\VerQueryValueW", "Ptr", data, "Str", "\VarFileInfo\Translation", "Ptr*", &pointer, "UInt*", &length) || length < 4
            return false
        language := Format("{:04x}{:04x}", NumGet(pointer, 0, "UShort"), NumGet(pointer, 2, "UShort"))
        if !DllCall("version\VerQueryValueW", "Ptr", data, "Str", "\StringFileInfo\" language "\FileDescription", "Ptr*", &pointer, "UInt*", &length)
            return false
        return StrGet(pointer, "UTF-16") = "EVE-X-Preview Recovery Updater (protocol 2)"
    }
    static StageUpdater(directory, destination, currentVersion) {
        adjacent := directory "\EVE-X-Preview-Updater.exe"
        if FileExist(adjacent) && this.RecoveryUpdater(adjacent) {
            FileCopy(adjacent, destination)
            return
        }
        if !RegExMatch(currentVersion, "^\d+(\.\d+){1,3}$")
            throw Error("Invalid installed version for updater recovery.")
        ; Never obtain an installer from a downgrade target: it may predate the
        ; recovery protocol. A missing compatible installer fails without launch.
        Download("https://github.com/khivus/EVE-X-Preview/releases/download/v" currentVersion "/EVE-X-Preview-Updater.exe", destination)
        if !this.RecoveryUpdater(destination)
            throw Error("Restore the matching EVE-X-Preview-Updater.exe beside the application before retrying this update.")
    }
    __New(mainPath, tag, sourceDirectory := "") {
        if !RegExMatch(tag, "^\d+(\.\d+){1,3}$")
            throw ValueError("Invalid release version.")
        SplitPath(mainPath,,, &extension)
        if StrLower(extension) != "exe"
            throw ValueError("Only compiled installations can be updated.")
        this.mainPath := mainPath, this.tag := tag
        SplitPath(mainPath,, &directory)
        this.directory := directory
        this.sourceDirectory := sourceDirectory
        this.includeSettings := !!FileExist(directory "\EVE-X-Preview-Settings.exe")
    }
    DownloadAsset(url, path) {
        if this.sourceDirectory != "" {
            SplitPath(path, &name)
            FileCopy(this.sourceDirectory "\" name, path)
        }
        else
            Download(url, path)
    }
    RunTransaction(engine, request) {
        executable := A_WinDir "\System32\WindowsPowerShell\v1.0\powershell.exe"
        return RunWait('"' executable '" -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "' engine '" -RequestPath "' request '"',, "Hide")
    }
    Transaction(workspace, action, status) {
        engine := workspace "\ExeUpdate.ps1"
        if A_IsCompiled {
            FileInstall "src\ExeUpdate.ps1", engine
        } else {
            SplitPath(A_LineFile,, &sourceDirectory)
            FileCopy(sourceDirectory "\ExeUpdate.ps1", engine)
        }
        request := workspace "\" action ".json"
        FileAppend(JSON.Dump(Map("action", action, "source", workspace, "main", this.mainPath, "version", this.tag)), request, "UTF-8")
        status.UpdateStatus(action = "recover" ? "Checking interrupted updates…" : "Verifying and installing release executables…")
        code := this.RunTransaction(engine, request)
        if !FileExist(request ".result")
            throw Error("The recovery installer did not return a result. Check Windows PowerShell availability and the updater log.")
        result := JSON.Load(FileRead(request ".result", "UTF-8"))
        ProgramLog.Add("Executable transaction: " result.Get("message", "No result message"))
        if code || !result.Get("ok", false)
            throw Error(result.Get("message", "Executable installation failed."))
        FileDelete(engine)
    }
    Recover(status) => this.Execute(status, true)
    Install(status) => this.Execute(status, false)
    Execute(status, recoverOnly) {
        workspace := A_Temp "\EVE-X-Preview-update-" ProcessExist() "-" A_TickCount "-" Random(10000, 99999)
        DirCreate(workspace)
        try {
            ; Recovery needs no downloads, and may restore a temporarily missing
            ; optional EXE. Decide which components exist only after it completes.
            this.Transaction(workspace, "recover", status)
            if recoverOnly
                return
            this.includeSettings := !!FileExist(this.directory "\EVE-X-Preview-Settings.exe")
            assets := ["EVE-X-Preview.exe"]
            if this.includeSettings
                assets.Push("EVE-X-Preview-Settings.exe")
            for asset in assets {
                status.UpdateStatus("Downloading " asset "…")
                this.DownloadAsset("https://github.com/khivus/EVE-X-Preview/releases/download/v" this.tag "/" asset, workspace "\" asset)
            }
            this.Transaction(workspace, "install", status)
        } finally {
            try DirDelete(workspace, true)
        }
    }
}
