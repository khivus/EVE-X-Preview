#Requires AutoHotkey v2.0
#SingleInstance Off
#NoTrayIcon
; Real transport failure fixture, never a host-desktop development entry point.
global __EVE_X_PREVIEW_TESTING__ := true
#Include "..\Main.ahk"
if A_UserName != "WDAGUtilityAccount"
    ExitApp(2)
Persistent
SetWorkingDir(A_ScriptDir)
FileEncoding("UTF-8")
ProgramLog.Init()
fixtureOwner := ResponseRetryOwner()

class ResponseRetryOwner {
    saves := 0
    __New() {
        this._JSON := JSON.Load(default_JSON)
        this.settingsBridge := ResponseRetryBridge(this, "offline")
    }
    SaveJsonToFile(*) {
        this.saves++
        path := A_ScriptDir "\fixture-save-count.txt"
        if FileExist(path)
            FileDelete(path)
        FileAppend(this.saves, path)
    }
}
class ResponseRetryBridge extends SettingsBridge {
    failedOnce := false
    WriteResponse(id, response) {
        if !this.failedOnce && response.Get("ok", false) && response.Get("data", 0) is Map
            && response["data"].Get("applied", false) {
            this.failedOnce := true
            throw Error("Fixture response write failed after the successful canonical save.")
        }
        return super.WriteResponse(id, response)
    }
}
