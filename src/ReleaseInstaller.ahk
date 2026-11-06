; Stage the complete release before replacing either installed component.
; An installation without the optional settings EXE keeps that choice.
class ReleaseInstaller {
    __New(mainPath, tag) {
        if !RegExMatch(tag, "^\d+(\.\d+){1,3}$")
            throw ValueError("Invalid release version.")
        SplitPath(mainPath,,, &extension)
        if StrLower(extension) != "exe"
            throw ValueError("Only compiled installations can be updated.")
        this.mainPath := mainPath, this.tag := tag
        SplitPath(mainPath,, &directory)
        this.directory := directory
        this.files := [{path: mainPath, asset: "EVE-X-Preview.exe"}]
        settings := directory "\EVE-X-Preview-Settings.exe"
        if FileExist(settings)
            this.files.Push({path: settings, asset: "EVE-X-Preview-Settings.exe"})
    }
    DownloadAsset(url, path) => Download(url, path)
    Verify(path) {
        if !FileExist(path) || VerCompare(FileGetVersion(path), this.tag) != 0
            throw Error("The downloaded executable does not match release v" this.tag ".")
    }
    Install(status) {
        replaced := [], backedUp := []
        try {
            for item in this.files {
                status.UpdateStatus("Downloading " item.asset "…")
                item.staged := item.path ".update", item.backup := item.path ".previous"
                this.DownloadAsset("https://github.com/khivus/EVE-X-Preview/releases/download/v" this.tag "/" item.asset, item.staged)
                this.Verify(item.staged)
            }
            for item in this.files {
                status.UpdateStatus("Installing " item.asset "…")
                if FileExist(item.path) {
                    FileMove(item.path, item.backup, true)
                    backedUp.Push(item)
                }
                FileMove(item.staged, item.path, true)
                replaced.Push(item)
            }
        } catch as err {
            for item in replaced
                try FileDelete(item.path)
            for item in backedUp
                FileMove(item.backup, item.path, true)
            throw err
        } finally {
            for item in this.files
                if item.HasOwnProp("staged") && FileExist(item.staged)
                    try FileDelete(item.staged)
        }
        for item in backedUp
            try FileDelete(item.backup)
    }
}
