#Requires AutoHotkey v2.0
#Include "src/SettingsWindowSizing.ahk"
#Include "src/ProgramLog.ahk"
#Include "src/ReleaseInstaller.ahk"

;@Ahk2Exe-Let U_version = 1.7.0.4
;@Ahk2Exe-SetVersion %U_version%
;@Ahk2Exe-SetFileVersion %U_version%
;@Ahk2Exe-SetCopyright gonzo83+khivus
;@Ahk2Exe-SetDescription EVE-X-Preview Updater
;@Ahk2Exe-SetProductName EVE-X-Preview Updater
;@Ahk2Exe-SetMainIcon res/Icon.ico
;@Ahk2Exe-AddResource res/Icon.ico, 160
;@Ahk2Exe-AddResource res/Icon-Suspend.ico, 206
;@Ahk2Exe-AddResource res/Icon.ico, 207
;@Ahk2Exe-AddResource res/Icon-Suspend.ico, 208

VERSION := "1.7.0.4"
if !A_IsCompiled
    TraySetIcon(A_ScriptDir "\res\Icon.ico",, true)
ProgramLog.Init(EnvGet("LOCALAPPDATA") "\EVE-X-Preview\Logs\Updater")
ProgramLog.Add("Standalone updater version=" VERSION)
OnError(UpdaterError)
UpdaterError(err, mode) {
    try ProgramLog.Error(err, "Unhandled updater " mode)
    return 0
}

class UpdaterStatus {
    __New() {
        this.Gui := Gui("+AlwaysOnTop -MaximizeBox -MinimizeBox", "Updater v" VERSION)
        this.ProgressBar := this.Gui.Add("Progress", "w300 h20")
        this.Text := this.Gui.Add("Text", "h100 w300", "Starting...`n")
        this.Gui.Show("Autosize")
    }

    UpdateStatus(text) {
        ProgramLog.Add("Updater: " text)
        ProgramLog.Flush()
        this.ProgressBar.Value += 20
        this.Text.Value .= text "`n"
    }

    Close() {
        this.Gui.Destroy()
    }
}

A_TrayMenu.Delete()
A_TrayMenu.Add("Exit", (*) => ExitApp())

status := UpdaterStatus()
status.UpdateStatus("Getting last version...")

try {
    oldScriptPath := A_Args[1]
    
    SplitPath(oldScriptPath, &name, &dir)
    if dir = ""
        oldScriptPath := A_ScriptDir "\" name

    newTag := A_Args[2]
    ProgramLog.Add("Updater arguments; destination=" oldScriptPath "; target=" newTag)
}
catch { ; If started without arguments we parse updates and get latest release version
    ProgramLog.Add("Updater started without complete arguments; looking up latest release")
    oldScriptPath := A_ScriptDir "\EVE-X-Preview.exe"

    try {
        ; Getting json of latest release
        apiUrl := "https://api.github.com/repos/khivus/EVE-X-Preview/releases"
        whr := ComObject("WinHttp.WinHttpRequest.5.1")
        whr.Open("GET", apiUrl)
        whr.SetRequestHeader("User-Agent", "AHK")
        whr.Send()
        whr.WaitForResponse()
        ProgramLog.Add("Updater release lookup; HTTP " whr.Status)
        if whr.Status != 200
            throw Error("GitHub release lookup HTTP " whr.Status)
        json_ans := whr.ResponseText
    }
    catch as err {
        ProgramLog.Error(err, "Updater release lookup")
        MsgBox("GitHub not available or no internet connection!") ; Probably no internet connection
        return
    }

    ; Finding tag of latest release
    pattern := '"tag_name":\s*"v?V?([^"]+)",[\s\S]*?"prerelease":\s*(true|false),'
    pos := 1
    latestReleaseTag := ""
    while matchPos := RegExMatch(json_ans, pattern, &match, pos) {
        tag := match[1]
        preRelease := match[2]

        if preRelease = "false" {
            latestReleaseTag := tag
            break
        }

        pos := matchPos + match.Len ; Advance past this match
    }

    if latestReleaseTag != ''
        newTag := latestReleaseTag
    else {
        ProgramLog.Add("Updater found no stable release; exiting", "WARN")
        ProgramLog.Flush()
        ExitApp
    }
}

SplitPath(oldScriptPath, &oldScriptName, &oldScriptDir)
SetWorkingDir(oldScriptDir)
ProgramLog.Add("Updater working directory=" oldScriptDir "; target=" newTag)
Sleep 500 ; Wait some time berfore doing anything

try {
    installer := ReleaseInstaller(oldScriptPath, newTag)
    ; The settings window is asked to close by the app before the handoff.
    ; Briefly allow Windows to release executable image handles.
    Sleep(500)
    installer.Install(status)
    status.UpdateStatus("Done! Launching " oldScriptName "!")
    Run('"' oldScriptPath '"', oldScriptDir)
    ProgramLog.Add("Updater relaunch requested successfully")
    ProgramLog.Flush()
    Sleep 1500
}
catch Error as e {
    ProgramLog.Error(e, "Installing app update")
    MsgBox("The update could not be installed. Existing files were retained or restored.`n" e.Message)
    if FileExist(oldScriptPath)
        try Run('"' oldScriptPath '"', oldScriptDir)
}
finally {
    status.Close()
    ExitApp
}
