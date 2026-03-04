#Requires AutoHotkey v2.0

#Include <DefaultJSON> ; The Default Settings Values
#Include <JSON>
#Include <LiveThumb>
; #Include <ScrollBar>
#Include <../src/Main_Class>
#Include <../src/ThumbWindow>
#Include <../src/TrayMenu>
#Include <../src/Propertys>
#Include <../src/Settings_Gui>

#SingleInstance Force
Persistent
ListLines False
KeyHistory 0

; CoordMode "Mouse", "Screen" ; to track Window Mouse possition while DragMoving the thumbnails
SetWinDelay -1
FileEncoding("UTF-8") ; Encoding for JSSON file

SetTitleMatchMode 3

A_MaxHotKeysPerInterval := 10000 

/*
TODO #########################
*/

;@Ahk2Exe-Let U_version = 1.5.0.1
;@Ahk2Exe-SetVersion %U_version%
;@Ahk2Exe-SetFileVersion %U_version%
;@Ahk2Exe-SetCopyright gonzo83+khivus
;@Ahk2Exe-SetDescription EVE-X-Preview
;@Ahk2Exe-SetProductName EVE-X-Preview
;@Ahk2Exe-ExeName EVE-X-Preview

;@Ahk2Exe-AddResource icon.ico, 160  ; Replaces 'H on blue'
;@Ahk2Exe-AddResource icon-suspend.ico, 206  ; Replaces 'S on green'
;@Ahk2Exe-AddResource icon.ico, 207  ; Replaces 'H on red'
;@Ahk2Exe-AddResource icon-suspend.ico, 208  ; Replaces 'S on red'

;@Ahk2Exe-SetMainIcon icon.ico

if !(A_IsCompiled)
    TraySetIcon("icon.ico",,true)

; Catch all unhandled Errors to prevent the Script from stopping 
OnError(Error_Handler)

Call := Main_Class()


; Improved settings loader
Load_JSON() {
    defaultPath := default_JSON  ; existing variable in your script pointing to default JSON content/path
    userPath    := "EVE-X-Preview.json"
    tmpPath     := "EVE-X-Preview.tmp.json"
    backupPath  := "EVE-X-Preview." . A_Now . ".bak.json"

    ; Load default JSON (throws on failure)
    DJSON := JSON.Load(defaultPath)

    ; If user file doesn't exist -> create it from default and return
    if !FileExist(userPath) {
        FileAppend(JSON.Dump(DJSON, , "    "), userPath)
        return JSON.Load(FileRead(userPath))
    }

    ; Try reading & merging
    try {
        UJSON := JSON.Load(FileRead(userPath))  ; may throw if corrupted

        ; If you keep a version key in JSON, run migrations before merging
        if (DJSON.Has("settings_version")) {
            dver := DJSON["settings_version"]
            uver := UJSON.Has("settings_version") ? UJSON["settings_version"] : ""
            if (uver != dver) {
                ; Migrate in-place (implement logic inside MigrateSettings)
                UJSON := MigrateSettings(UJSON, uver, dver)
                ; ensure we set new version so merge will not re-trigger
                UJSON["settings_version"] := dver
                ; Create a timestamped backup of the existing file before merge
                FileMove(userPath, backupPath, true)
            }
        }

        ; Merge: default -> user, but don't overwrite user's explicit values
        _JSON := JsonMergeNoOverwrite(DJSON, UJSON)

        ; Atomic write: write to tmp file then move/rename to target
        if FileExist(tmpPath)
            FileDelete(tmpPath)
        FileAppend(JSON.Dump(_JSON, , "    "), tmpPath)
        FileMove(tmpPath, userPath, true)

    } catch Error as e {
        MsgBox("Exception at " e.File ":" e.Line "`n" e.Message "`n" e.Extra)
        ; corrupted or other error: ask user and recreate from default if they agree
        value := MsgBox("The settings file is corrupted.`nDo you want to create a new one?`nOld one will be backed up.",, "YesNo")
        if (value = "No")
            ExitApp()

        ; backup the corrupted file (if not already moved)
        try FileMove(userPath, backupPath, true)
        catch

        FileAppend(JSON.Dump(DJSON, , "    "), userPath)
        _JSON := JSON.Load(FileRead(userPath))
        return _JSON
    }

    return _JSON
}

DeepClone(obj) {
    if !IsObject(obj)
        return obj

    ; Detect array

    if IsArrayLike(obj) {
        newArr := obj.Clone()   ; create same-type array

        while (newArr.Length > 0)
            newArr.Pop()

        for _, v in obj
            newArr.Push(DeepClone(v))

        return newArr
    }

    ; Create same-type object
    newObj := obj.Clone()
    newObj.Clear()

    for k, v in obj
        newObj[k] := DeepClone(v)

    return newObj
}

IsArrayLike(obj) {
    if !IsObject(obj)
        return false
    try {
        _ := obj.Length
        return true
    } catch {
        return false
    }
}

; Profile aware JsonMergeNoOverwrite
; Merge defaults (objDefault) into user's object (objUser) without overwriting user's explicit values.
JsonMergeNoOverwrite(objDefault, objUser) {
    ; If objUser is NOT an object, replace it entirely
    if !IsObject(objUser) {
        return DeepClone(objDefault)
    }

    for key, defVal in objDefault {

        if !objUser.Has(key) {
            ; if objUser is array, skip keyed assignment
            if IsArrayLike(objUser)
                continue
            else if IsObject(defVal)
                objUser[key] := DeepClone(defVal)
            else
                objUser[key] := defVal

            continue
        }

        userVal := objUser[key]

        ; If both are arrays -> do NOT treat as object recursion
        if IsArrayLike(defVal) && IsArrayLike(userVal)
            continue

        ; _Profiles is an object keyed by profile-name
        if (key = "_Profiles" && IsObject(defVal) && IsObject(userVal)) {
            ; find canonical default profile (if present)
            defaultProfile := defVal.Has("Default") ? defVal["Default"] : {}

            ; merge canonical profile into every existing user profile
            for profName, profObj in userVal {
                if IsObject(profObj) && IsObject(defaultProfile) {
                    userVal[profName] := JsonMergeNoOverwrite(defaultProfile, profObj)
                }
            }

            ; add any default-only profiles (copy whole profile objects)
            for defProfName, defProfObj in defVal {
                if !userVal.Has(defProfName) {
                    userVal[defProfName] := DeepClone(defProfObj)
                }
            }

            objUser[key] := DeepClone(userVal)
            continue
        }

        ; RECURSIVE OBJECT MERGE
        if IsObject(defVal) && IsObject(userVal) {
            objUser[key] := JsonMergeNoOverwrite(defVal, userVal)
            continue
        }

        ; TYPE MISMATCH: default object vs user primitive -> prefer default structure
        if IsObject(defVal) && !IsObject(userVal) {
            objUser[key] := DeepClone(defVal)
            continue
        }

        ; default primitive & user object -> keep user's object
        ; both primitives -> keep user's value (no overwrite)
    }

    return objUser
}

; Simple migration hook: implement any structural changes between versions here.
; - uver may be "", meaning user has no version key (old file).
; - Return the modified userObj.
MigrateSettings(userObj, uver, dver) {

    if !IsObject(userObj)
        return userObj

    if (uver == "1" || uver == "") && dver == "2" {

        userObj["new_anime"] := "What the"

        if !userObj.Has("global_Settings")
            throw Error("Missing global_Settings!")

        ; Global -> out
        g := userObj["global_Settings"]
        if g.Has("LastUsedProfile")
            userObj["LastUsedProfile"] := g["LastUsedProfile"]
        if g.Has("First_Start_After_Update")
            userObj["First_Start_After_Update"] := g["First_Start_After_Update"]
        if g.Has("ThisThat")
            userObj["ThisThat"] := g["ThisThat"]

        if !userObj.Has("_Profiles") || !IsObject(userObj["_Profiles"])
            throw Error("No profiles found!")

        for prof_name, prof_settings in userObj["_Profiles"] {

            if !IsObject(prof_settings) {
                ; if it's not an object, skip (unexpected)
                continue
            }

            ; Hotkeys -> Hotkeys : CharacterHotkeys
            if prof_settings.Has("Hotkeys") {
                tempHot := DeepClone(prof_settings["Hotkeys"]) ; Write hotkeys to temp location
                prof_settings["Hotkeys"] := Map() ; new object (fresh)
                prof_settings.Set("Hotkeys", Map("CharacterHotkeys", tempHot))
            } else {
                prof_settings["Hotkeys"] := Map()
                prof_settings["Hotkeys"]["CharacterHotkeys"] := []
            }

            ; Global -> Other
            prof_settings["Other"] := {}
            if g.Has("SwitchLangOnErr")
                prof_settings.Set("Other", Map("SwitchLangOnErr", g["SwitchLangOnErr"]))
            if g.Has("Check_Updates")
                prof_settings.Set("Other", Map("Check_Updates", g["Check_Updates"]))

            ; Global -> Client Settings
            if g.Has("Minimize_Delay")
                prof_settings.Set("Client Settings", Map("Minimize_Delay", g["Minimize_Delay"]))

            ; Global -> Thumbnail Settings
            ts := prof_settings["Thumbnail Settings"]
            if g.Has("ThumbnailStartLocation")
                ts["ThumbnailStartLocation"] := DeepClone(g["ThumbnailStartLocation"])
            if g.Has("ThumbnailBackgroundColor")
                ts["ThumbnailBackgroundColor"] := g["ThumbnailBackgroundColor"]
            if g.Has("ThumbnailSnap")
                ts["ThumbnailSnap"] := g["ThumbnailSnap"]
            if g.Has("ThumbnailSnap_Distance")
                ts["ThumbnailSnap_Distance"] := g["ThumbnailSnap_Distance"]
            if g.Has("ThumbnailMinimumSize")
                ts["ThumbnailMinimumSize"] := DeepClone(g["ThumbnailMinimumSize"])
            if g.Has("HideThumbForActiveWin")
                ts["HideThumbForActiveWin"] := g["HideThumbForActiveWin"]
            if g.Has("ShiftThumbsForLoginScreen")
                ts["ShiftThumbsForLoginScreen"] := g["ShiftThumbsForLoginScreen"]
            if g.Has("ShiftThumbsCollisionCheck")
                ts["ShiftThumbsCollisionCheck"] := g["ShiftThumbsCollisionCheck"]
            if g.Has("ShiftThumbsDirection")
                ts["ShiftThumbsDirection"] := g["ShiftThumbsDirection"]
            if g.Has("ShiftThumbHorizontalStep")
                ts["ShiftThumbHorizontalStep"] := g["ShiftThumbHorizontalStep"]
            if g.Has("ShiftThumbVerticalStep")
                ts["ShiftThumbVerticalStep"] := g["ShiftThumbVerticalStep"]
            if g.Has("PreserveThumbPosOnLogout")
                ts["PreserveThumbPosOnLogout"] := g["PreserveThumbPosOnLogout"]
            if g.Has("PreserveCharNameOnLogout")
                ts["PreserveCharNameOnLogout"] := g["PreserveCharNameOnLogout"]

            ; Global -> Hotkeys
            hs := prof_settings["Hotkeys"]
            if g.Has("Suspend_Hotkeys_Hotkey")
                hs["Suspend_Hotkeys_Hotkey"] := g["Suspend_Hotkeys_Hotkey"]
            if g.Has("Global_Hotkeys")
                hs["Global_Hotkeys"] := g["Global_Hotkeys"]
            if g.Has("Login_Screen_Cycle_Hotkey")
                hs["Login_Screen_Cycle_Hotkey"] := g["Login_Screen_Cycle_Hotkey"]
            if g.Has("LoginScreenCycleDirection")
                hs["LoginScreenCycleDirection"] := g["LoginScreenCycleDirection"]
            if g.Has("PreserveHotkeysOnLogout")
                hs["PreserveHotkeysOnLogout"] := g["PreserveHotkeysOnLogout"]
            if g.Has("Close_Active_EVE_Win_Hotkey")
                hs["Close_Active_EVE_Win_Hotkey"] := g["Close_Active_EVE_Win_Hotkey"]
            if g.Has("Close_All_EVE_Win_Hotkey")
                hs["Close_All_EVE_Win_Hotkey"] := g["Close_All_EVE_Win_Hotkey"]
            if g.Has("Reload_Program_Hotkey")
                hs["Reload_Program_Hotkey"] := g["Reload_Program_Hotkey"]
        }

        try
            userObj.Delete("global_Settings")
        catch
            throw Error("Error deleting global_Settings!")
    }

    return userObj
}

; Hanles unmanaged Errors
Error_Handler(Thrown, Mode) {
    ; There we try to get right layout of keyboard
    if Thrown.Message == "Invalid key name." {
        if !hwnd := WinActive("A") {
            try {
                WinActivate "ahk_class Shell_TrayWnd"
                WinActivate "ahk_class Button"
                hwnd := WinActive("A")
            }
        }

        try {
            Send "{LAlt down}{Shift down}"
            Sleep(10)
            Send "{Shift up}{LAlt up}"
        }
        Sleep(50)

        EN_US := 0x0409

        try {
            PostMessage(0x50, 0, EN_US, , hwnd)
        }
        Sleep(50)

        try {
            control := ControlGetFocus(hwnd)
            PostMessage(0x50, 0, EN_US, , control)
        }
        Sleep(50)

        Reload
    }

    MsgBox("Error: " Thrown.Message "`nIn file: " Thrown.File " at line: " Thrown.Line "`n" Thrown.Extra)
    
    return -1
}