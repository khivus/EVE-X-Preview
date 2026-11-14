#Requires AutoHotkey v2.0
#Include "src/SettingsWindowSizing.ahk"
#Include "src/ProgramLog.ahk"
#Include "Lib/json.ahk"
#Include "src/ReleaseInstaller.ahk"

;@Ahk2Exe-Let U_version = 1.7.0.7
;@Ahk2Exe-SetVersion %U_version%
;@Ahk2Exe-SetFileVersion %U_version%
;@Ahk2Exe-SetCopyright gonzo83+khivus
;@Ahk2Exe-SetDescription EVE-X-Preview Recovery Updater (protocol 2)
;@Ahk2Exe-SetProductName EVE-X-Preview Updater
;@Ahk2Exe-SetMainIcon res/Icon.ico
;@Ahk2Exe-AddResource res/Icon.ico, 160
;@Ahk2Exe-AddResource res/Icon-Suspend.ico, 206
;@Ahk2Exe-AddResource res/Icon.ico, 207
;@Ahk2Exe-AddResource res/Icon-Suspend.ico, 208

VERSION := "1.7.0.7"
if !A_IsCompiled
    TraySetIcon(A_ScriptDir "\res\Icon.ico",, true)
ProgramLog.Init(EnvGet("LOCALAPPDATA") "\EVE-X-Preview\Logs\Updater")

class UpdaterStatus {
    __New() {
        this.Gui := Gui("+AlwaysOnTop -MaximizeBox -MinimizeBox", "Updater v" VERSION)
        this.ProgressBar := this.Gui.Add("Progress", "w340 h20")
        this.Text := this.Gui.Add("Text", "h160 w340", "Starting...`n")
        this.Gui.Show("Autosize")
    }
    UpdateStatus(value) {
        ProgramLog.Add("Updater: " value)
        ProgramLog.Flush()
        this.ProgressBar.Value := Min(100, this.ProgressBar.Value + 15)
        this.Text.Value .= value "`n"
    }
    Close() => this.Gui.Destroy()
}

UpdaterFullPath(path) {
    pathBuffer := Buffer(65536)
    length := DllCall("GetFullPathNameW", "Str", path, "UInt", 32768, "Ptr", pathBuffer, "Ptr", 0, "UInt")
    if !length || length >= 32768
        throw Error("Invalid updater path.")
    return StrGet(pathBuffer)
}
UpdaterArguments() {
    if !A_Args.Length
        return {mode: "latest", main: A_ScriptDir "\EVE-X-Preview.exe", version: "0.0", source: ""}
    if A_Args.Length = 2 && (A_Args[1] = "--recover" || A_Args[1] = "--latest")
        return {mode: A_Args[1] = "--recover" ? "recover" : "latest", main: UpdaterFullPath(A_Args[2]), version: "0.0", source: ""}
    if (A_Args.Length != 2 && A_Args.Length != 3) || !RegExMatch(A_Args[2], "^\d+(\.\d+){1,3}$")
        throw Error("Use <main EXE> <version> [local EXE directory], or --recover <main EXE>.")
    return {mode: "install", main: UpdaterFullPath(A_Args[1]), version: A_Args[2], source: A_Args.Length = 3 ? UpdaterFullPath(A_Args[3]) : ""}
}
UpdaterLatestVersion() {
    request := ComObject("WinHttp.WinHttpRequest.5.1")
    request.SetTimeouts(5000, 5000, 15000, 15000)
    request.Open("GET", "https://api.github.com/repos/khivus/EVE-X-Preview/releases/latest")
    request.SetRequestHeader("User-Agent", "EVE-X-Preview")
    request.SetRequestHeader("Accept", "application/vnd.github+json")
    request.Send()
    if request.Status != 200
        throw Error("GitHub release lookup failed; HTTP " request.Status ".")
    release := JSON.Load(request.ResponseText)
    tag := RegExReplace(release.Get("tag_name", ""), "i)^v")
    if release.Get("draft", true) || release.Get("prerelease", true) || !RegExMatch(tag, "^\d+(\.\d+){1,3}$")
        throw Error("No supported stable release was found.")
    return tag
}

status := 0, options := 0
try {
    options := UpdaterArguments()
    SplitPath(options.main,, &installation)
    ; An installed updater relocates before opening its own replacement lock.
    if A_IsCompiled && StrLower(A_ScriptDir) = StrLower(installation) {
        temporary := A_Temp "\EVE-X-Preview-installer-" ProcessExist() "-" A_TickCount
        DirCreate(temporary)
        executable := temporary "\EVE-X-Preview-Updater.exe"
        FileCopy(A_ScriptFullPath, executable)
        arguments := options.mode = "install" ? '"' options.main '" "' options.version '"' (options.source != "" ? ' "' options.source '"' : "")
            : '--' options.mode ' "' options.main '"'
        Run('"' executable '" ' arguments, temporary)
        ExitApp()
    }
    A_TrayMenu.Delete()
    A_TrayMenu.Add("Exit", (*) => ExitApp())
    status := UpdaterStatus()
    installer := ReleaseInstaller(options.main, options.version, options.source)
    if options.mode != "install"
        installer.Recover(status) ; Works before any network request, including offline.
    if options.mode = "latest"
        installer.tag := UpdaterLatestVersion()
    if options.mode != "recover"
        installer.Install(status)
    status.UpdateStatus("Done! Launching EVE-X-Preview…")
    Run('"' options.main '"', installation)
    ProgramLog.Flush()
    Sleep(1000)
} catch as err {
    ProgramLog.Error(err, "Executable update")
    MsgBox("The update could not be completed.`n" err.Message, "EVE-X-Preview update")
    if IsObject(options) {
        SplitPath(options.main,, &installation)
        if FileExist(options.main) && !DirExist(installation "\.EVE-X-Preview-update")
            try Run('"' options.main '"', installation)
    }
} finally {
    if IsObject(status)
        status.Close()
    ExitApp()
}
