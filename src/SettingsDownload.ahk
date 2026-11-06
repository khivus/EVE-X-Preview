; The optional editor follows the main release version. Downloads are staged
; beside the application and replaced only after the complete EXE is verified.
class SettingsDownload {
    busy := false
    status := ""
    __New(owner) => this.owner := owner
    Directory => this.owner.HasOwnProp("settingsBridge") ? this.owner.settingsBridge.installation : A_ScriptDir
    Path => this.Directory "\EVE-X-Preview-Settings.exe"
    Current(tag) {
        if tag = "" || !FileExist(this.Path)
            return false
        try return VerCompare(FileGetVersion(this.Path), tag) >= 0
        return false
    }
    Start(url, tag) {
        if this.busy || this.Current(tag)
            return
        if !RegExMatch(url, "^https://github\.com/khivus/EVE-X-Preview/releases/download/[^/]+/EVE-X-Preview-Settings\.exe$")
            throw Error("Check for updates first. No settings editor download is available yet.")
        if SettingsEditor(this.owner).InternalDraft()
            throw Error("Apply or discard the internal settings draft before downloading the editor.")
        this.busy := true, this.status := "Downloading settings editor…"
        this.stagedReady := false, this.closeRequestedAt := 0
        this.tag := tag, this.startedAt := A_TickCount
        try {
            this.request := ComObject("WinHttp.WinHttpRequest.5.1")
            this.request.SetTimeouts(5000, 5000, 15000, 120000)
            this.request.Open("GET", url, true)
            this.request.SetRequestHeader("User-Agent", "EVE-X-Preview")
            this.request.Send()
            this.timer := ObjBindMethod(this, "Poll")
            SetTimer(this.timer, 200)
        } catch as err
            this.Finish("Settings download failed: " err.Message)
    }
    Poll() {
        temporary := this.Path ".download"
        try {
            if A_TickCount - this.startedAt > 180000
                throw Error("Download timed out. Try again in About.")
            if !this.stagedReady {
                if !this.request.WaitForResponse(0)
                    return
                if this.request.Status != 200
                    throw Error("GitHub returned HTTP " this.request.Status ".")
                stream := ComObject("ADODB.Stream")
                try {
                    stream.Type := 1, stream.Open()
                    stream.Write(this.request.ResponseBody)
                    stream.SaveToFile(temporary, 2)
                } finally stream.Close()
                if VerCompare(FileGetVersion(temporary), this.tag) != 0
                    throw Error("The downloaded editor does not match the selected release.")
                this.stagedReady := true
            }
            if this.owner.HasOwnProp("settingsBridge") && this.owner.settingsBridge.HasExternal {
                bridge := this.owner.settingsBridge
                if bridge.session.Dirty
                    throw Error("Apply or discard the settings draft before replacing the editor.")
                if !this.closeRequestedAt {
                    this.closeRequestedAt := A_TickCount
                    this.status := "Download complete. Closing the previous settings editor…"
                    PostMessage(0x10, 0, 0,, "ahk_pid " bridge.ExternalPid)
                }
                if A_TickCount - this.closeRequestedAt > 10000
                    throw Error("Close the previous settings editor, then retry the download.")
                return
            }
            backup := this.Path ".previous"
            existed := !!FileExist(this.Path)
            if existed
                FileMove(this.Path, backup, true)
            try FileMove(temporary, this.Path, true)
            catch as err {
                if existed
                    FileMove(backup, this.Path, true)
                throw err
            }
            if existed
                try FileDelete(backup)
            this.Finish("Settings editor downloaded. Opening settings…")
            ; Return the bridge response before changing editor ownership.
            SetTimer(ObjBindMethod(this, "Launch"), -300)
        } catch as err {
            try FileDelete(temporary)
            this.Finish("Settings download failed: " err.Message)
        }
    }
    Launch() {
        editor := SettingsEditor(this.owner)
        if editor.InternalDraft()
            return
        editor.CloseCleanInternal()
        editor.Launch(this.Path)
    }
    Finish(message) {
        if this.HasOwnProp("timer")
            SetTimer(this.timer, 0)
        if this.HasOwnProp("request") {
            try this.request.Abort()
            this.DeleteProp("request")
        }
        this.busy := false, this.status := message
        ProgramLog.Add(message)
    }
}
