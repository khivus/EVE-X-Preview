#Requires AutoHotkey v2.0

VERSION := "1.0"

try {
    oldScriptName := A_Args[1]
    newTag := A_Args[2]
}
catch { ; If started without arguments we parse updates and get latest release version
    oldScriptName := "EVE-X-Preview.exe"

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

SetWorkingDir(A_ScriptDir)
Sleep 500 ; Wait some time berfore doing anything

; Downloading file and running
newScriptName := "EVE-X-Preview-v" newTag ".exe"
oldNewScriptName := "EVE-X-Preview-Old.exe"

try {
    if FileExist(oldScriptName) { ; Renaming old script to different name
        FileMove(oldScriptName, oldNewScriptName, true)
        if !FileExist(oldNewScriptName) ; If not renamed
            Throw Error("Error renaming old script!")
    }

    ; Checking if file of new verison exist and deleting if so
    if FileExist(newScriptName)
        FileDelete(newScriptName)

    exeUrl := "https://github.com/khivus/EVE-X-Preview/releases/download/v" newTag "/EVE-X-Preview.exe"
    Download(exeUrl, newScriptName) ; Download file from GitHub

    ; Checking if new version downloaded
    if !FileExist(newScriptName)
        Throw Error("Error downloading new version!")

    FileMove(newScriptName, oldScriptName, true) ; Renaming new script to old name

    if FileExist(oldNewScriptName) { ; Deleting old script
        FileDelete(oldNewScriptName)
    }
        
    Run(oldScriptName)
}
catch Error as e {
    MsgBox("An error occurred while trying to update the application:`n" e.Message "`nIF program stops working properly, redownload it from github!")
}
finally {
    ExitApp
}