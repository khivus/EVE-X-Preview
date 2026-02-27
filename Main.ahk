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
FileEncoding("UTF-8") ; Encoding for JSSON file

SetTitleMatchMode 3

A_MaxHotKeysPerInterval := 10000 

/*
TODO #########################
*/

;@Ahk2Exe-Let U_version = 1.4.1.2
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
            }
        }

        ; Merge: default -> user, but don't overwrite user's explicit values
        _JSON := JsonMergeNoOverwrite(DJSON, UJSON)

        ; Create a timestamped backup of the existing file before changing it
        ; FileMove(userPath, backupPath, true)  ; Used only for debug purposes

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
    ; Will be used later
    ; Add more conditionals for each version step as needed.
    ; By default do nothing and return userObj
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

    MsgBox("Error: " Thrown.Message "`n" Thrown.Extra)
    
    return -1
}