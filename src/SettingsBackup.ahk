; Startup snapshots are independent of the editable draft and migration backups.
class SettingsBackup {
    static Root := ""
    static sequence := 0
    static SuggestedRepair := ""

    static PrepareFirstRun(installation, hadConfig) {
        root := this.Root != "" ? this.Root : EnvGet("LOCALAPPDATA") "\EVE-X-Preview\Backups"
        marker := root "\.recovery-guidance-seen"
        first := !hadConfig && !FileExist(marker) && !DirExist(this.Directory(installation))
        try {
            DirCreate(root)
            if !FileExist(marker)
                FileAppend(first ? "pending" : "seen", marker, "UTF-8")
        }
        return first
    }
    static TakeFirstRunNotice(config) {
        pending := config.Get("RecoveryGuidancePending", false)
        if config.Has("RecoveryGuidancePending")
            config.Delete("RecoveryGuidancePending")
        marker := (this.Root != "" ? this.Root : EnvGet("LOCALAPPDATA") "\EVE-X-Preview\Backups") "\.recovery-guidance-seen"
        try {
            if !pending || FileRead(marker,"UTF-8") != "pending"
                return false
            file := FileOpen(marker,"w","UTF-8")
            file.Write("seen")
            file.Close()
            return true
        }
        return false
    }

    static LatestValid(installation) {
        installation := SettingsBridge.NormalizeInstallation(installation)
        candidates := ""
        backupDirectory := this.Directory(installation)
        for directory in [backupDirectory, installation] {
            pattern := directory = backupDirectory ? "*.json" : "EVE-X-Preview*.json"
            loop files directory "\" pattern, "F" {
                if A_LoopFileFullPath = installation "\EVE-X-Preview.json" || InStr(A_LoopFileName, ".tmp")
                    continue
                try candidates .= FileGetTime(A_LoopFileFullPath, "M") "`t" A_LoopFileFullPath "`n"
            }
        }
        for item in StrSplit(RTrim(Sort(candidates, "R"), "`n"), "`n") {
            if item = ""
                continue
            path := SubStr(item, InStr(item, "`t") + 1)
            try {
                if FileGetSize(path) > 16777216
                    continue
                config := JSON.Load(FileRead(path, "UTF-8"))
                if config is Map && config.Has("_Profiles") && config["_Profiles"] is Map && config["_Profiles"].Has("Default") && config["_Profiles"]["Default"] is Map {
                    SettingsModel(config)
                    return path
                }
            }
        }
        return ""
    }

    static HelpText(installation) {
        candidate := this.SuggestedRepair
        return "If settings are lost or corrupted, check automatic snapshots and migration backups in:`n" this.Directory(installation)
            . "`n`nAlso check EVE-X-Preview*.json beside the app, especially in an older installation."
            . "`n`n" (candidate != "" ? "Latest readable saved copy suggested for repair:`n" candidate : "No readable saved copy was found automatically. Check older installation folders too.")
            . "`n`nTo repair: close EVE-X-Preview and Settings, keep a copy of the current config, copy a backup into the app folder as EVE-X-Preview.json, then reopen the app. Try an older snapshot or migration copy if needed."
            . "`n`nIf recovery still fails, report the problem to khivus on GitHub or Discord:`nhttps://github.com/khivus/EVE-X-Preview/issues/new`nhttps://discord.gg/d9EzgUXnT6"
    }

    static ShowGuidance(installation) {
        window := Gui(, "EVE-X-Preview — Settings recovery")
        window.SetFont("s10", "Segoe UI")
        window.AddText("w640", this.HelpText(installation))
        window.AddButton("w155", "Open settings backups").OnEvent("Click", (*) => this.Open(installation))
        window.AddButton("x+8 w140", "Open app folder").OnEvent("Click", (*) => Run('explorer.exe "' installation '"'))
        window.AddButton("x+8 w100", "GitHub").OnEvent("Click", (*) => Run("https://github.com/khivus/EVE-X-Preview/issues/new"))
        window.AddButton("x+8 w100", "Discord").OnEvent("Click", (*) => Run("https://discord.gg/d9EzgUXnT6"))
        window.AddButton("x+8 w70 Default", "Close").OnEvent("Click", (*) => window.Destroy())
        window.OnEvent("Close", (*) => window.Destroy())
        window.OnEvent("Escape", (*) => window.Destroy())
        window.Show()
        return window
    }

    static Directory(installation) => (this.Root != "" ? this.Root : EnvGet("LOCALAPPDATA") "\EVE-X-Preview\Backups") "\" SettingsBridge.InstallationId(installation)

    static ArchivePath(source, name) {
        SplitPath(SettingsBridge.NormalizeInstallation(source),, &installation)
        SplitPath(name, &leaf)
        directory := this.Directory(installation)
        DirCreate(directory)
        path := directory "\" leaf
        while FileExist(path)
            path := directory "\" (++this.sequence) "." leaf
        return path
    }

    static CollectMigrationBackups(installation) {
        loop files installation "\EVE-X-Preview.*.backup.json", "F" {
            try FileMove(A_LoopFileFullPath, this.ArchivePath(installation "\EVE-X-Preview.json", A_LoopFileName), false)
            catch as err
                ProgramLog.Add("Migration backup move skipped: " err.Message, "WARN")
        }
    }

    static Open(installation) {
        directory := this.Directory(installation)
        DirCreate(directory)
        Run('explorer.exe "' directory '"')
    }

    static Capture(source) {
        temporary := ""
        try {
            source := SettingsBridge.NormalizeInstallation(source)
            SplitPath(source,, &installation)
            directory := this.Directory(installation)
            DirCreate(directory)
            stamp := 0
            DllCall("GetSystemTimeAsFileTime", "Int64*", &stamp)
            destination := directory "\EVE-X-Preview." A_NowUTC "." Format("{:019}", stamp) "." DllCall("GetCurrentProcessId") "." Format("{:06}", ++this.sequence) ".json"
            temporary := destination ".tmp"
            FileCopy(source, temporary, false)
            FileMove(temporary, destination, false)
            files := ""
            loop files directory "\EVE-X-Preview.*.json", "F"
                if !InStr(A_LoopFileName, ".backup.json")
                    files .= A_LoopFileFullPath "`n"
            for index, path in StrSplit(RTrim(Sort(files, "R"), "`n"), "`n")
                if index > 5
                    FileDelete(path)
            return destination
        } catch as err {
            if temporary != "" && FileExist(temporary)
                try FileDelete(temporary)
            ProgramLog.Add("Startup settings backup skipped: " err.Message, "WARN")
            return ""
        }
    }
}
