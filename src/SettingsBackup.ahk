; Startup snapshots are independent of the editable draft and migration backups.
class SettingsBackup {
    static Root := ""
    static sequence := 0

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
