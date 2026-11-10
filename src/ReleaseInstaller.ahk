; Download and verify a complete versioned package before replacing any component.
; The PowerShell transaction is embedded in the updater; no .NET installation is needed.
class ReleaseInstaller {
    static PackageUpdater(path) {
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
        return StrGet(pointer, "UTF-16") = "EVE-X-Preview Package Updater (protocol 1)"
    }
    static StageUpdater(directory, destination, currentVersion) {
        adjacent := directory "\EVE-X-Preview-Updater.exe"
        if FileExist(adjacent) && this.PackageUpdater(adjacent) {
            FileCopy(adjacent, destination)
            return
        }
        if !RegExMatch(currentVersion, "^\d+(\.\d+){1,3}$")
            throw Error("Invalid installed version for updater recovery.")
        ; Never obtain an installer from a downgrade target: it may predate the
        ; package protocol. A missing compatible installer fails without launch.
        Download("https://github.com/khivus/EVE-X-Preview/releases/download/v" currentVersion "/EVE-X-Preview-Updater.exe", destination)
        if !this.PackageUpdater(destination)
            throw Error("This installation needs its package-aware updater. Restore EVE-X-Preview-Updater.exe from the full portable package.")
    }
    __New(mainPath, tag, archivePath := "") {
        if !RegExMatch(tag, "^\d+(\.\d+){1,3}$")
            throw ValueError("Invalid release version.")
        SplitPath(mainPath,,, &extension)
        if StrLower(extension) != "exe"
            throw ValueError("Only compiled installations can be updated.")
        this.mainPath := mainPath, this.tag := tag
        SplitPath(mainPath,, &directory)
        this.directory := directory
        this.archivePath := archivePath
        this.includeSettings := !!FileExist(directory "\EVE-X-Preview-Settings.exe")
    }
    DownloadAsset(url, path) {
        if this.archivePath != ""
            FileCopy(this.archivePath (SubStr(url, -7) = ".sha256" ? ".sha256" : ""), path)
        else
            Download(url, path)
    }
    RunTransaction(engine, request) {
        executable := A_WinDir "\System32\WindowsPowerShell\v1.0\powershell.exe"
        return RunWait('"' executable '" -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "' engine '" -RequestPath "' request '"',, "Hide")
    }
    Install(status) {
        workspace := A_Temp "\EVE-X-Preview-update-" ProcessExist() "-" A_TickCount "-" Random(10000, 99999)
        DirCreate(workspace)
        try {
            asset := "EVE-X-Preview-win-x64.zip"
            url := "https://github.com/khivus/EVE-X-Preview/releases/download/v" this.tag "/" asset
            archive := workspace "\" asset, checksum := archive ".sha256"
            status.UpdateStatus("Downloading the complete release package…")
            this.DownloadAsset(url, archive)
            this.DownloadAsset(url ".sha256", checksum)
            if FileGetSize(checksum) > 1024 || !RegExMatch(Trim(FileRead(checksum, "UTF-8")), "i)^([a-f0-9]{64})(?:[ \t]+\*?[^\r\n]+)?$", &digest)
                throw Error("The release package checksum is missing or invalid. Install its full portable package manually.")
            engine := workspace "\PackageUpdate.ps1"
            if A_IsCompiled {
                FileInstall "src\PackageUpdate.ps1", engine
            } else {
                SplitPath(A_LineFile,, &sourceDirectory)
                FileCopy(sourceDirectory "\PackageUpdate.ps1", engine)
            }
            request := workspace "\request.json"
            FileAppend(JSON.Dump(Map("archive", archive, "main", this.mainPath, "version", this.tag,
                "includeSettings", this.includeSettings, "sha256", digest[1])), request, "UTF-8")
            status.UpdateStatus("Verifying and installing all package files…")
            exitCode := this.RunTransaction(engine, request)
            if !FileExist(request ".result")
                throw Error("The package installer did not return a result. Check Windows PowerShell availability and the updater log.")
            result := JSON.Load(FileRead(request ".result", "UTF-8"))
            ProgramLog.Add("Package transaction: " result.Get("message", "No result message"))
            if exitCode || !result.Get("ok", false)
                throw Error(result.Get("message", "Package installation failed."))
        } finally {
            try DirDelete(workspace, true)
        }
    }
}
