#Requires AutoHotkey v2.0

VERSION := "1.4"

class UpdaterStatus {
    __New() {
        this.Gui := Gui("+AlwaysOnTop -MaximizeBox -MinimizeBox", "Updating")
        this.ProgressBar := this.Gui.Add("Progress", "w300 h20")
        this.Text := this.Gui.Add("Text", "h100 w300", "Starting...`n")
        this.Gui.Show("Autosize")
    }

    UpdateStatus(text) {
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
}
catch { ; If started without arguments we parse updates and get latest release version
    oldScriptPath := A_ScriptDir "\EVE-X-Preview.exe"

    try {
        ; Getting json of latest release
        apiUrl := "https://api.github.com/repos/khivus/EVE-X-Preview/releases"
        whr := ComObject("WinHttp.WinHttpRequest.5.1")
        whr.Open("GET", apiUrl)
        whr.SetRequestHeader("User-Agent", "AHK")
        whr.Send()
        whr.WaitForResponse()
        json_ans := whr.ResponseText
    }
    catch {
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
    else
        ExitApp
}

SplitPath(oldScriptPath, &oldScriptName, &oldScriptDir)
SetWorkingDir(oldScriptDir)
Sleep 500 ; Wait some time berfore doing anything

; Downloading file and running
newScriptName := "EVE-X-Preview-v" newTag ".exe"
oldNewScriptName := "EVE-X-Preview-Old.exe"

try {
    status.UpdateStatus("Renaming old version...")
    if FileExist(oldScriptName) { ; Renaming old script to different name
        FileMove(oldScriptName, oldNewScriptName, true)
        if !FileExist(oldNewScriptName) ; If not renamed
            Throw Error("Error renaming old script!")
    }

    ; Checking if file of new verison exist and deleting if so
    if FileExist(newScriptName)
        FileDelete(newScriptName)

    exeUrl := "https://github.com/khivus/EVE-X-Preview/releases/download/v" newTag "/EVE-X-Preview.exe"
    status.UpdateStatus("Downloading new version...")
    Download(exeUrl, newScriptName) ; Download file from GitHub

    ; Checking if new version downloaded
    if !FileExist(newScriptName)
        Throw Error("Error downloading new version!")

    status.UpdateStatus("Renaming new version...")
    FileMove(newScriptName, oldScriptName, true) ; Renaming new script to old name

    status.UpdateStatus("Deleting old version...")
    if FileExist(oldNewScriptName) ; Deleting old script
        FileDelete(oldNewScriptName)
        
    status.UpdateStatus("Done! Launching " oldScriptName "!")
    Run(oldScriptName)
    Sleep 1500
}
catch Error as e {
    MsgBox("An error occurred while trying to update the program:`n" e.Message "`nIf program stops working properly, redownload it from github!")
}
finally {
    status.Close()
    ExitApp
}