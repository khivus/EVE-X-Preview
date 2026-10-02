#Requires AutoHotkey v2.0
#Include "src/ProgramLog.ahk"

VERSION := "1.5"
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

; Downloading file and running
newScriptName := "EVE-X-Preview-v" newTag ".exe"
oldNewScriptName := "EVE-X-Preview-Old.exe"

try {
    status.UpdateStatus("Renaming old version...")
    if FileExist(oldScriptName) { ; Renaming old script to different name
        FileMove(oldScriptName, oldNewScriptName, true)
        ProgramLog.Add("Updater backup created: " oldNewScriptName)
        if !FileExist(oldNewScriptName) ; If not renamed
            Throw Error("Error renaming old script!")
    }

    ; Checking if file of new verison exist and deleting if so
    if FileExist(newScriptName)
        FileDelete(newScriptName)

    exeUrl := "https://github.com/khivus/EVE-X-Preview/releases/download/v" newTag "/EVE-X-Preview.exe"
    status.UpdateStatus("Downloading new version...")
    ProgramLog.Add("Updater download URL=" exeUrl)
    ProgramLog.Flush()
    Download(exeUrl, newScriptName) ; Download file from GitHub

    ; Checking if new version downloaded
    if !FileExist(newScriptName)
        Throw Error("Error downloading new version!")
    ProgramLog.Add("Updater executable downloaded; bytes=" FileGetSize(newScriptName))

    status.UpdateStatus("Renaming new version...")
    FileMove(newScriptName, oldScriptName, true) ; Renaming new script to old name
    ProgramLog.Add("Updater executable replaced: " oldScriptName)

    status.UpdateStatus("Deleting old version...")
    if FileExist(oldNewScriptName) ; Deleting old script
        FileDelete(oldNewScriptName)
        
    status.UpdateStatus("Done! Launching " oldScriptName "!")
    Run(oldScriptName)
    ProgramLog.Add("Updater relaunch requested successfully")
    ProgramLog.Flush()
    Sleep 1500
}
catch Error as e {
    ProgramLog.Error(e, "Installing app update")
    MsgBox("An error occurred while trying to update the program:`n" e.Message "`nIf program stops working properly, redownload it from github!")
}
finally {
    status.Close()
    ExitApp
}
