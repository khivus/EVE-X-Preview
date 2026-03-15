#Requires AutoHotkey v2.0

#Include <DefaultJSON> ; The Default Settings Values
#Include <JSON>
#Include <LiveThumb>
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
FileEncoding("UTF-8") ; Encoding for JSON file

SetTitleMatchMode 3

A_MaxHotKeysPerInterval := 10000 

;@Ahk2Exe-Let U_version = 1.5.1.8
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
                ; Create a timestamped backup of the existing file before merge
                FileMove(userPath, backupPath, true)
                ; Migrate in-place (implement logic inside MigrateSettings)
                UJSON := MigrateSettings(UJSON, uver, dver)
                ; ensure we set new version so merge will not re-trigger
                UJSON["settings_version"] := dver
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
; Merge defaultb (objDefault) into user's object (objUser) without overwriting user's explicit values.
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

            ; add any default-only profiles (copy whole profile objectb)
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

    if uver == "" && Integer(dver) >= 1 {

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

            ; Global -> Other
            prof_settings["Other"] := Map()
            if g.Has("SwitchLangOnErr")
                prof_settings.Set("Other", Map("SwitchLangOnErr", g["SwitchLangOnErr"]))
            if g.Has("Check_Updates")
                prof_settings.Set("Other", Map("Check_Updates", g["Check_Updates"]))

            ; Global -> Client Settings
            if g.Has("Minimize_Delay")
                prof_settings.Set("Client Settings", Map("Minimize_Delay", g["Minimize_Delay"]))

            ; Global -> Thumbnails Behavior
            prof_settings["Thumbnails Behavior"] := Map()
            tb := prof_settings["Thumbnails Behavior"]
            if g.Has("ThumbnailStartLocation")
                tb["ThumbnailStartLocation"] := DeepClone(g["ThumbnailStartLocation"])
            if g.Has("ThumbnailSnap")
                tb["ThumbnailSnap"] := g["ThumbnailSnap"]
            if g.Has("ThumbnailSnap_Distance")
                tb["ThumbnailSnap_Distance"] := g["ThumbnailSnap_Distance"]
            if g.Has("HideThumbForActiveWin")
                tb["HideThumbForActiveWin"] := g["HideThumbForActiveWin"]
            if g.Has("ShiftThumbsForLoginScreen")
                tb["ShiftThumbsForLoginScreen"] := g["ShiftThumbsForLoginScreen"]
            if g.Has("ShiftThumbsCollisionCheck")
                tb["ShiftThumbsCollisionCheck"] := g["ShiftThumbsCollisionCheck"]
            if g.Has("ShiftThumbsDirection")
                tb["ShiftThumbsDirection"] := g["ShiftThumbsDirection"]
            if g.Has("ShiftThumbHorizontalStep")
                tb["ShiftThumbHorizontalStep"] := g["ShiftThumbHorizontalStep"]
            if g.Has("ShiftThumbVerticalStep")
                tb["ShiftThumbVerticalStep"] := g["ShiftThumbVerticalStep"]
            if g.Has("PreserveThumbPosOnLogout")
                tb["PreserveThumbPosOnLogout"] := g["PreserveThumbPosOnLogout"]
            if g.Has("PreserveCharNameOnLogout")
                tb["PreserveCharNameOnLogout"] := g["PreserveCharNameOnLogout"]

            ; Global -> Thumbnails Visuals
            prof_settings["Thumbnails Visuals"] := Map()
            tv := prof_settings["Thumbnails Visuals"]
            if g.Has("ThumbnailBackgroundColor")
                tv["ThumbnailBackgroundColor"] := g["ThumbnailBackgroundColor"]
            if g.Has("ThumbnailMinimumSize")
                tv["ThumbnailMinimumSize"] := DeepClone(g["ThumbnailMinimumSize"])

            ; Thumbnail Settings -> Thumbnails Behavior
            ts := prof_settings["Thumbnail Settings"]
            if ts.Has("HideThumbnailsOnLostFocus")
                tb["HideThumbnailsOnLostFocus"] := ts["HideThumbnailsOnLostFocus"]
            if ts.Has("ShowThumbnailsAlwaysOnTop")
                tb["ShowThumbnailsAlwaysOnTop"] := ts["ShowThumbnailsAlwaysOnTop"]

            ; Thumbnail Settings -> Thumbnails Visuals
            if ts.Has("ShowThumbnailTextOverlay")
                tv["ShowThumbnailTextOverlay"] := ts["ShowThumbnailTextOverlay"]
            if ts.Has("ThumbnailTextColor")
                tv["ThumbnailTextColor"] := ts["ThumbnailTextColor"]
            if ts.Has("ThumbnailTextSize")
                tv["ThumbnailTextSize"] := ts["ThumbnailTextSize"]
            if ts.Has("ThumbnailTextFont")
                tv["ThumbnailTextFont"] := ts["ThumbnailTextFont"]
            if ts.Has("ThumbnailTextMargins")
                tv["ThumbnailTextMargins"] := DeepClone(ts["ThumbnailTextMargins"])
            if ts.Has("ShowClientHighlightBorder")
                tv["ShowClientHighlightBorder"] := ts["ShowClientHighlightBorder"]
            if ts.Has("ClientHighligtColor")
                tv["ClientHighligtColor"] := ts["ClientHighligtColor"]
            if ts.Has("ClientHighligtBorderthickness")
                tv["ClientHighligtBorderthickness"] := ts["ClientHighligtBorderthickness"]
            if ts.Has("ThumbnailOpacity")
                tv["ThumbnailOpacity"] := ts["ThumbnailOpacity"]
            if ts.Has("ShowAllColoredBorders")
                tv["ShowAllColoredBorders"] := ts["ShowAllColoredBorders"]
            if ts.Has("InactiveClientBorderthickness")
                tv["InactiveClientBorderthickness"] := ts["InactiveClientBorderthickness"]
            if ts.Has("InactiveClientBorderColor")
                tv["InactiveClientBorderColor"] := ts["InactiveClientBorderColor"]

            ; Hotkeys -> Hotkeys Settings
            prof_settings["Hotkeys Settings"] := Map()
            hs := prof_settings["Hotkeys Settings"]

            if prof_settings.Has("Hotkeys")
                hs["CharacterHotkeys"] := DeepClone(prof_settings["Hotkeys"])

            ; Global -> Hotkeys Settings
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

            ; Exclude from Closing -> Client settings
            if prof_settings.has("Exclude from Closing") {
                efc := prof_settings["Exclude from Closing"]
                if efc.Has("ExcludeOnLoginScreen")
                    prof_settings["Client Settings"]["DontCloseOnLoginScreen"] := efc["ExcludeOnLoginScreen"] ; +Renamed
                if efc.Has("DontCloseClients")
                    prof_settings["Client Settings"]["DontCloseClients"] := DeepClone(efc["DontCloseClients"])
            }

            for group in ["Hotkeys", "Exclude from Closing", "Thumbnail Settings"]
                try
                    prof_settings.Delete(group)
        }

        try
            userObj.Delete("global_Settings")
    }

    if (uver = "" || Integer(uver) = 1) && Integer(dver) >= 2 {
        if !userObj.Has("_Profiles") || !IsObject(userObj["_Profiles"])
            throw Error("No profiles found!")

        for prof_name, prof_settings in userObj["_Profiles"] {
            if !IsObject(prof_settings) ; if it's not an object, skip (unexpected)
                continue

            ; Updating all characters names with prefix "EVE - "
            if prof_settings.Has("Hotkey Groups") {
                for k, v in prof_settings["Hotkey Groups"] {
                    temp := []
                    for i, c in v["Characters"]
                        temp.Push(AddEVE(c))
                    prof_settings["Hotkey Groups"][k]["Characters"] := DeepClone(temp)
                }
            }
            if prof_settings.Has("Hotkeys Settings") {
                temp := []
                for c, v in prof_settings["Hotkeys Settings"]["CharacterHotkeys"]
                    for c, h in v
                        temp.Push(Map(AddEVE(c), h))
                prof_settings["Hotkeys Settings"]["CharacterHotkeys"] := DeepClone(temp)
            }
            if prof_settings.Has("Thumbnail Visibility") {
                temp := Map()
                for c, v in prof_settings["Thumbnail Visibility"]
                    temp[AddEVE(c)] := v
                prof_settings["Thumbnail Visibility"] := DeepClone(temp)
            }
            if prof_settings.Has("Client Settings") {
                temp := []
                for i, c in prof_settings["Client Settings"]["DontCloseClients"]
                    temp.Push(AddEVE(c))
                prof_settings["Client Settings"]["DontCloseClients"] := DeepClone(temp)
                temp := []
                for i, c in prof_settings["Client Settings"]["Dont_Minimize_Clients"]
                    temp.Push(AddEVE(c))
                prof_settings["Client Settings"]["Dont_Minimize_Clients"] := DeepClone(temp)
            }
            if prof_settings.Has("Custom Colors") {
                temp := []
                for i, c in prof_settings["Custom Colors"]["cColors"]["CharNames"]
                    temp.Push(AddEVE(c))
                prof_settings["Custom Colors"]["cColors"]["CharNames"] := DeepClone(temp)
            }
            if prof_settings.Has("Client Possitions") {
                temp := Map()
                for c, v in prof_settings["Client Possitions"]
                    temp[AddEVE(c)] := DeepClone(v)
                prof_settings["Client Possitions"] := DeepClone(temp)
            }
            if prof_settings.Has("Thumbnail Positions") {
                temp := Map()
                for c, v in prof_settings["Thumbnail Positions"]
                    temp[AddEVE(c)] := DeepClone(v)
                prof_settings["Thumbnail Positions"] := DeepClone(temp)
            }
        }
    }

    return userObj

    AddEVE(name) {
        return "EVE - " . name
    }
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