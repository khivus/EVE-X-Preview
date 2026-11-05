; Routes both tray entry points without moving any runtime settings into C#.
class SettingsEditor {
    static ProbeTimeout := 15000
    static HandshakeTimeout := 60000

    static Open(owner, *) => SettingsEditor(owner).OpenPreferred()
    static OpenInternal(owner, *) {
        if owner.HasOwnProp("settingsBridge") && IsObject(owner.settingsBridge)
            if !owner.settingsBridge.TryOpenInternal()
                return false
        owner.ShowInternalSettings()
        return true
    }

    __New(owner) => this.owner := owner
    Preference() {
        value := this.owner._JSON.Get("SettingsEditor", "automatic")
        return value = "modern" || value = "builtin" ? value : "automatic"
    }
    Candidate() => A_ScriptDir "\EVE-X-Preview-Settings.exe"
    Exists(path) => !!FileExist(path)
    BuiltIn() => SettingsEditor.OpenInternal(this.owner)

    OpenPreferred() {
        if this.Preference() = "builtin"
            return this.BuiltIn()
        if this.ExistingExternal()
            return true
        if this.InternalDraft()
            return false
        path := this.Candidate()
        if !this.Exists(path)
            return this.Fallback("The modern settings editor is not installed beside EVE-X-Preview.")
        try {
            if !this.Compatible(this.Probe(path))
                return this.Fallback("The modern settings editor uses an incompatible bridge or schema version.")
            this.CloseCleanInternal()
            this.Launch(path)
            return true
        } catch as err {
            return this.Fallback("The modern settings editor could not start: " err.Message)
        }
    }

    ExistingExternal() {
        if !this.owner.HasOwnProp("settingsBridge") || !IsObject(this.owner.settingsBridge)
            return false
        bridge := this.owner.settingsBridge
        if !bridge.HasExternal
            return false
        try WinActivate("ahk_pid " bridge.ExternalPid)
        return true
    }
    InternalDraft() {
        if !this.owner.HasOwnProp("settingsWorkbench") || this.owner.settingsWorkbench.closed
            return false
        editor := this.owner.settingsWorkbench
        if !editor.HasDraftChanges() && !editor.HasOwnProp("activeDialog") && !editor.HasOwnProp("closePrompt")
            return false
        editor.S_Gui.Show()
        try WinActivate(editor.S_Gui.Hwnd)
        editor.UpdateStatus("Apply or discard this draft before opening the modern settings editor.")
        return true
    }
    CloseCleanInternal() {
        if this.owner.HasOwnProp("settingsWorkbench") && !this.owner.settingsWorkbench.closed
            this.owner.settingsWorkbench.Dispose()
    }
    static Compatible(value) => value is Map
        && value.Get("application", "") = "EVE-X-Preview-Settings"
        && value.Get("bridgeVersion", 0) = 1 && value.Get("schemaVersion", 0) = 1
    Compatible(value) => SettingsEditor.Compatible(value)

    Probe(path) {
        output := A_Temp "\EVE-X-Preview-probe-" ProcessExist() "-" A_TickCount "-" Random(10000, 99999) ".json"
        pid := 0
        try {
            Run('"' path '" --probe "' output '"', A_ScriptDir, "Hide", &pid)
            deadline := A_TickCount + SettingsEditor.ProbeTimeout
            while ProcessExist(pid) && A_TickCount < deadline
                Sleep(15)
            if ProcessExist(pid) {
                try ProcessClose(pid)
                throw Error("Compatibility check timed out.")
            }
            if !FileExist(output) || FileGetSize(output) > 4096
                throw Error("Compatibility check did not return a valid response.")
            return JSON.Load(FileRead(output, "UTF-8"))
        } finally {
            if FileExist(output)
                FileDelete(output)
        }
    }
    Launch(path) {
        pid := 0
        Run('"' path '" --installation "' A_ScriptDir '"', A_ScriptDir,, &pid)
        this.pid := pid
        this.deadline := A_TickCount + SettingsEditor.HandshakeTimeout
        this.watch := ObjBindMethod(this, "WatchHandshake")
        this.owner.settingsEditorLaunch := this
        SetTimer(this.watch, 75)
    }
    WatchHandshake(*) {
        if this.ExistingExternal() {
            SetTimer(this.watch, 0)
            this.owner.settingsEditorStatus := "Modern settings editor connected."
            return
        }
        if A_TickCount >= this.deadline || !ProcessExist(this.pid) {
            SetTimer(this.watch, 0)
            ; Only the process started by this attempt is eligible for cleanup.
            ; Its temporary canonical host observes the parent and exits too.
            if ProcessExist(this.pid) {
                try WinClose("ahk_pid " this.pid)
                if ProcessExist(this.pid)
                    try ProcessClose(this.pid)
            }
            this.Fallback("The modern settings editor did not complete its startup handshake. Internal settings were opened.")
        }
    }
    Fallback(reason) {
        this.owner.settingsEditorStatus := reason
        ProgramLog.Add(reason, "WARN")
        return this.BuiltIn()
    }
}
