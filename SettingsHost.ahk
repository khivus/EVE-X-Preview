#Requires AutoHotkey v2.0
#NoTrayIcon
; Import canonical settings/migrations without Main_Class.__New or runtime work.
global __EVE_X_PREVIEW_TESTING__ := true
#Include "Main.ahk"
#SingleInstance Off
;@Ahk2Exe-Let U_version = 1.7.0.2
;@Ahk2Exe-SetVersion %U_version%
;@Ahk2Exe-SetFileVersion %U_version%
;@Ahk2Exe-SetDescription EVE-X-Preview canonical settings host
;@Ahk2Exe-SetProductName EVE-X-Preview

Persistent
SetWorkingDir(A_ScriptDir)
FileEncoding("UTF-8")
parentPid := 0
if A_Args.Length = 2 && A_Args[1] = "--parent" && RegExMatch(A_Args[2], "^[1-9][0-9]{0,9}$")
    parentPid := Integer(A_Args[2])
if !parentPid || !ProcessExist(parentPid)
    ExitApp(2)
try {
    if SettingsBridge.ProcessUser(parentPid) != SettingsBridge.ProcessUser(DllCall("GetCurrentProcessId"))
        ExitApp(2)
    ProgramLog.Init()
    hostOwner := SettingsHostOwner(parentPid)
} catch as err {
    try ProgramLog.Error(err, "Offline settings host initialization")
    ExitApp(2)
}

class SettingsHostOwner {
    __New(parentPid) {
        this.parentPid := parentPid
        this.parentHandle := DllCall("OpenProcess", "UInt", 0x100000, "Int", false, "UInt", parentPid, "Ptr")
        if !this.parentHandle
            throw OSError()
        this.instance := DllCall("CreateMutexW", "Ptr", 0, "Int", false, "Str", "Local\EVE-X-Preview.SettingsOffline/" SettingsBridge.InstallationId(A_ScriptDir), "Ptr")
        if !this.instance
            throw OSError()
        if A_LastError = 183
            ExitApp()
        this.settingsBridge := SettingsBridge(this, "offline")
        ; If a runtime is already starting, let the editor discover that endpoint.
        if this.settingsBridge.RuntimeWindow()
            ExitApp()
        lock := SettingsFileLock(A_ScriptDir)
        try {
            if this.settingsBridge.RuntimeWindow()
                ExitApp()
            this._JSON := Load_JSONUnlocked(false)
        } finally lock.Dispose()
        this.settingsBridge.ready := true
        this.timer := ObjBindMethod(this, "CheckParent")
        SetTimer(this.timer, 500)
    }
    CheckParent() {
        if DllCall("WaitForSingleObject", "Ptr", this.parentHandle, "UInt", 0, "UInt") != 0x102
            ExitApp()
        ; Keep the endpoint just long enough for the editor to receive
        ; runtime-started and reconnect using its opaque canonical snapshot.
        if this.settingsBridge.RuntimeWindow() {
            if !this.HasOwnProp("handoffStarted")
                this.handoffStarted := A_TickCount
            if A_TickCount - this.handoffStarted > 10000
                ExitApp()
        }
    }
    SaveJsonToFile(*) {
        lock := SettingsFileLock(A_ScriptDir)
        try {
            if this.settingsBridge.RuntimeWindow()
                throw SettingsServiceError("runtime-started", "EVE-X-Preview has started. Reconnect with your retained draft.")
            ; Detect changes made by another canonical host or a disk editor.
            current := JSON.Load(FileRead(A_ScriptDir "\EVE-X-Preview.json"))
            if IsObject(this.settingsBridge.session) && this.settingsBridge.session.model.Conflicts(current).Length
                throw SettingsServiceError("conflict", "The settings file changed while this offline draft was open.")
            if IsObject(this.settingsBridge.session)
                this._JSON := this.settingsBridge.session.model.Merge(current)
            serialized := JSON.Dump(this._JSON,, "    ")
            destination := A_ScriptDir "\EVE-X-Preview.json"
            temporary := destination ".settings-host.tmp"
            file := FileOpen(temporary, "w", "UTF-8-RAW")
            if !file
                throw OSError()
            try {
                file.Write(serialized)
                if !DllCall("FlushFileBuffers", "Ptr", file.Handle)
                    throw OSError()
            } finally file.Close()
            FileMove(temporary, destination, true)
        } finally lock.Dispose()
    }
}
