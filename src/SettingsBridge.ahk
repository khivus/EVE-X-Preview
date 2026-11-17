; Bounded local-only RPC. COPYDATA contains a request; the synchronous return
; acknowledges an atomically flushed response file, never a pointer or callback.
class SettingsBridge {
    static Magic := 0x45585031
    static MaxBytes := 4 * 1024 * 1024
    ExternalPid := 0
    ready := false
    session := 0
    senderPid := 0

    __New(owner, role := "runtime", installation := A_ScriptDir, transport := true) {
        this.owner := owner
        this.role := role
        this.installation := SettingsBridge.NormalizeInstallation(installation)
        this.installationId := SettingsBridge.InstallationId(installation)
        this.title := "EVE-X-Preview.SettingsBridge/1/" this.installationId "/" role
        this.responseDir := EnvGet("LOCALAPPDATA") "\EVE-X-Preview\IPC\" this.installationId
        this.cache := Map()
        this.cacheOrder := []
        this.cacheBytes := 0
        this.nextCleanup := A_TickCount + 60000
        this.ready := owner.HasOwnProp("_JSON")
        this.transport := transport
        if transport {
            this.window := Gui("+ToolWindow -Caption", this.title)
            handle := this.window.Hwnd ; Hidden top-level HWND, discoverable before Load_JSON.
            this.copyData := ObjBindMethod(this, "OnCopyData")
            OnMessage(0x4A, this.copyData)
            this.poll := ObjBindMethod(this, "CheckClient")
            SetTimer(this.poll, 1000)
            this.exitCallback := ObjBindMethod(this, "Dispose")
            OnExit(this.exitCallback)
            try {
                DirCreate(this.responseDir)
                this.Cleanup()
            } catch as err {
                ; A filesystem problem with optional IPC must not prevent the
                ; standalone runtime and maintained internal editor from starting.
                ProgramLog.Error(err, "Settings bridge response storage unavailable; use internal settings")
            }
        }
    }

    HasExternal => this.ExternalPid && ProcessExist(this.ExternalPid)

    static NormalizeInstallation(path) {
        size := DllCall("GetFullPathNameW", "Str", path, "UInt", 0, "Ptr", 0, "Ptr", 0, "UInt")
        if !size
            throw OSError()
        pathBuffer := Buffer(size * 2, 0)
        if !DllCall("GetFullPathNameW", "Str", path, "UInt", size, "Ptr", pathBuffer, "Ptr", 0, "UInt")
            throw OSError()
        return StrLower(RTrim(StrReplace(StrGet(pathBuffer), "/", "\"), "\/"))
    }

    static InstallationId(path) {
        value := this.NormalizeInstallation(path), data := Buffer(StrPut(value, "UTF-8"), 0)
        StrPut(value, data, "UTF-8")
        algorithm := 0, hash := 0
        if DllCall("bcrypt\BCryptOpenAlgorithmProvider", "Ptr*", &algorithm, "Str", "SHA256", "Ptr", 0, "UInt", 0, "UInt")
            throw Error("Cannot initialize settings installation identity.")
        try {
            if DllCall("bcrypt\BCryptCreateHash", "Ptr", algorithm, "Ptr*", &hash, "Ptr", 0, "UInt", 0, "Ptr", 0, "UInt", 0, "UInt", 0, "UInt")
                throw Error("Cannot create installation identity hash.")
            if DllCall("bcrypt\BCryptHashData", "Ptr", hash, "Ptr", data, "UInt", data.Size - 1, "UInt", 0, "UInt")
                throw Error("Cannot hash installation identity.")
            digest := Buffer(32, 0)
            if DllCall("bcrypt\BCryptFinishHash", "Ptr", hash, "Ptr", digest, "UInt", 32, "UInt", 0, "UInt")
                throw Error("Cannot finish installation identity.")
            result := ""
            loop 32
                result .= Format("{:02x}", NumGet(digest, A_Index - 1, "UChar"))
            return result
        } finally {
            if hash
                DllCall("bcrypt\BCryptDestroyHash", "Ptr", hash)
            DllCall("bcrypt\BCryptCloseAlgorithmProvider", "Ptr", algorithm, "UInt", 0)
        }
    }

    static NewId() {
        bytes := Buffer(16, 0)
        if DllCall("bcrypt\BCryptGenRandom", "Ptr", 0, "Ptr", bytes, "UInt", 16, "UInt", 2, "UInt")
            throw Error("Cannot create a secure settings session ID.")
        result := ""
        loop 16
            result .= Format("{:02x}", NumGet(bytes, A_Index - 1, "UChar"))
        return result
    }

    static ProcessUser(pid) {
        process := DllCall("OpenProcess", "UInt", 0x1000, "Int", false, "UInt", pid, "Ptr")
        if !process
            throw OSError()
        token := 0
        try {
            if !DllCall("advapi32\OpenProcessToken", "Ptr", process, "UInt", 8, "Ptr*", &token)
                throw OSError()
            size := 0
            DllCall("advapi32\GetTokenInformation", "Ptr", token, "Int", 1, "Ptr", 0, "UInt", 0, "UInt*", &size)
            if !size
                throw OSError()
            tokenBuffer := Buffer(size, 0)
            if !DllCall("advapi32\GetTokenInformation", "Ptr", token, "Int", 1, "Ptr", tokenBuffer, "UInt", size, "UInt*", &size)
                throw OSError()
            sid := NumGet(tokenBuffer, 0, "Ptr"), pointer := 0
            if !DllCall("advapi32\ConvertSidToStringSidW", "Ptr", sid, "Ptr*", &pointer)
                throw OSError()
            try return StrGet(pointer)
            finally DllCall("LocalFree", "Ptr", pointer)
        } finally {
            if token
                DllCall("CloseHandle", "Ptr", token)
            DllCall("CloseHandle", "Ptr", process)
        }
    }

    OnCopyData(wParam, lParam, message, hwnd) {
        if hwnd != this.window.Hwnd
            return
        try {
            if !wParam || !DllCall("IsWindow", "Ptr", wParam) || !lParam
                return 0
            pid := 0
            DllCall("GetWindowThreadProcessId", "Ptr", wParam, "UInt*", &pid)
            if !pid || SettingsBridge.ProcessUser(pid) != SettingsBridge.ProcessUser(DllCall("GetCurrentProcessId"))
                return 0
            if NumGet(lParam, 0, "UPtr") != SettingsBridge.Magic
                return 0
            size := NumGet(lParam, A_PtrSize, "UInt"), pointer := NumGet(lParam, A_PtrSize * 2, "Ptr")
            if size < 3 || size > SettingsBridge.MaxBytes || !pointer || NumGet(pointer, size - 1, "UChar") != 0
                return 0
            payload := StrGet(pointer, size - 1, "UTF-8")
            if StrPut(payload, "UTF-8") != size
                return 0
            request := JSON.Load(payload)
            if !(request is Map) || !RegExMatch(request.Get("id", ""), "^[a-f0-9]{32}$")
                return 0
            this.senderPid := pid
            try response := this.Handle(request)
            finally this.senderPid := 0
            this.WriteResponse(request["id"], response)
            ; Synchronous SendMessage returns only after the response is flushed.
            if this.HasOwnProp("restartPending") && this.restartPending {
                this.restartPending := false
                SetTimer(ObjBindMethod(this, "RestartRuntime"), -150)
            }
            return 1
        } catch as err {
            ProgramLog.Error(err, "Settings bridge transport; response unavailable")
            return 0
        }
    }

    Handle(request) {
        id := request is Map ? request.Get("id", "") : ""
        try {
            if !(request is Map) || !RegExMatch(id, "^[a-f0-9]{32}$")
                throw SettingsServiceError("invalid-request", "The request ID is invalid.")
            if request.Get("version", 0) != SettingsContract.BridgeVersion
                throw SettingsServiceError("incompatible", "The settings editor uses an incompatible bridge version.")
            method := request.Get("method", ""), args := request.Get("arguments", 0)
            if !(args is Map) || !(method is String) || StrLen(method) > 64 || !(request.Get("session", "") is String)
                throw SettingsServiceError("invalid-request", "The settings request is malformed.")
            fingerprint := JSON.Dump(request)
            if this.cache.Has(id) {
                cached := this.cache[id]
                if cached.fingerprint != fingerprint || cached.pid != this.senderPid
                    throw SettingsServiceError("invalid-request", "A request ID was reused for different settings data.")
                return cached.response
            }
            if method = "hello"
                data := Map("bridgeVersion", 1, "schemaVersion", 1, "role", this.role, "installation", this.installationId,
                    "ready", SettingsContract.Bool(this.ready), "pid", DllCall("GetCurrentProcessId"))
            else if method = "metadata"
                data := SettingsContract.Export()
            else {
                if !this.ready
                    throw SettingsServiceError("not-ready", "EVE-X-Preview is still loading settings. Please reconnect shortly.")
                if this.role = "offline" && this.RuntimeWindow()
                    throw SettingsServiceError("runtime-started", "EVE-X-Preview has started. Reconnect to the runtime with your retained draft.")
                if method = "begin"
                    data := this.Begin(args)
                else {
                    if !IsObject(this.session) || request.Get("session", "") != this.session.id || (this.senderPid && this.ExternalPid != this.senderPid)
                        throw SettingsServiceError("stale-session", "This settings session is no longer active. Reconnect with your retained draft.")
                    data := this.Dispatch(method, args)
                }
            }
            response := Map("version", 1, "id", id, "ok", JSON.true, "data", data)
            ; Freeze response snapshots: later edits must not mutate the opaque
            ; recovery state attached to a previously acknowledged request ID.
            bytes := StrPut(JSON.Dump(response), "UTF-8") + StrPut(fingerprint, "UTF-8")
            this.cache[id] := {fingerprint: fingerprint, pid: this.senderPid, response: SettingsContract.Wire(response), bytes: bytes}
            this.cacheOrder.Push(id)
            this.cacheBytes += bytes
            while this.cache.Count > 64 || (this.cacheBytes > SettingsBridge.MaxBytes * 4 && this.cache.Count > 1) {
                previousId := this.cacheOrder.RemoveAt(1)
                previous := this.cache.Delete(previousId)
                this.cacheBytes -= previous.bytes
            }
            return response
        } catch as err {
            code := err is SettingsServiceError ? err.code : "validation"
            return Map("version", 1, "id", id, "ok", JSON.false,
                "error", Map("code", code, "message", err.Message, "details", err is SettingsServiceError ? err.details : []))
        }
    }

    Begin(args) {
        pid := args.Get("clientPid", 0)
        if !IsInteger(pid) || pid <= 0 || !ProcessExist(pid) || (this.senderPid && pid != this.senderPid)
            throw SettingsServiceError("invalid-request", "The settings client process identity is invalid.")
        if this.owner.HasOwnProp("settingsWorkbench") && !this.owner.settingsWorkbench.closed {
            editor := this.owner.settingsWorkbench
            if editor.HasDraftChanges() || editor.HasOwnProp("activeDialog") || editor.HasOwnProp("closePrompt") {
                editor.S_Gui.Show()
                throw SettingsServiceError("busy", "Finish or discard the open internal settings draft before opening the modern editor.")
            }
            editor.Dispose()
        }
        if IsObject(this.session) {
            if this.ExternalPid = pid
                return this.session.Snapshot()
            if this.HasExternal {
                this.FocusExternal()
                throw SettingsServiceError("busy", "Another settings draft is already open. Finish that draft first.")
            }
            if this.session.Dirty {
                ; A dead client cannot release its draft. Resume the existing
                ; authoritative session rather than silently replacing it.
                this.ExternalPid := pid
                return this.session.Snapshot()
            }
        }
        nextSession := SettingsSession(this.owner, args.Get("recovery", 0), this.role = "runtime")
        if IsObject(this.session)
            this.session.Dispose()
        this.session := nextSession
        this.ExternalPid := pid
        return this.session.Snapshot()
    }

    Dispatch(method, args) {
        model := this.session.model
        switch method {
            case "snapshot": return this.session.Snapshot()
            case "client-state":
                this.session.clientDirty := !!args.Get("dirty", false)
            case "select-profile": model.SelectProfile(args.Get("name", ""))
            case "set":
                id := args.Get("id", "")
                if !SettingsSchema.Fields.Has(id)
                    throw ValueError("Unknown settings field.")
                if !args.Has("value") || IsObject(args["value"])
                    throw ValueError("Enter one scalar setting value.")
                model.Set(id, args["value"])
            case "profile":
                switch args.Get("action", "") {
                    case "add": model.AddProfile(args.Get("name", ""), args.Get("copyCurrent", true), args.Get("inheritance", model.CreationInheritance()))
                    case "rename": model.RenameProfile(args.Get("name", ""))
                    case "delete": model.DeleteProfile()
                    case "activate": model.ActivateProfile()
                    default: throw ValueError("Unknown profile action.")
                }
            case "inheritance": model.SetInheritance(args.Get("section", ""), args.Get("enabled", false))
            case "creation-inheritance":
                flags := args.Get("flags", 0)
                if !(flags is Map)
                    throw ValueError("Choose section inheritance flags.")
                model.SetCreationInheritance(flags)
            case "rows":
                return args.Has("rows") ? this.session.Rows(args.Get("kind", ""), args.Get("group", ""), args["rows"])
                    : this.session.Rows(args.Get("kind", ""), args.Get("group", ""))
            case "group": return this.session.Group(args)
            case "event": return args.Has("values") ? this.session.Event(args.Get("id", ""), args["values"]) : this.session.Event(args.Get("id", ""))
            case "sound": return this.session.Sound(args)
            case "restore-draft": this.session.Restore(args["draft"])
            case "undo-category":
                section := args.Get("section", "")
                SettingsSchema.FindSection(section)
                original := SettingsModel(model.original)
                if section = "profiles" {
                    restored := Map()
                    for name, profile in original.Profiles {
                        currentName := model.Profiles.Has(name) ? name : ""
                        for renamed, originName in model.profileOrigins
                            if originName = name && model.Profiles.Has(renamed)
                                currentName := renamed
                        restored[name] := SettingsModel.Clone(currentName != "" ? model.Profiles[currentName] : profile)
                        restored[name]["ActivationHotkey"] := profile.Get("ActivationHotkey", "")
                    }
                    selected := model.profileOrigins.Get(model.profile, model.profile)
                    model.config["_Profiles"] := restored
                    model.config["LastUsedProfile"] := original.config["LastUsedProfile"]
                    model.profileOrigins := Map()
                    for name in restored
                        model.profileOrigins[name] := name
                    model.profile := restored.Has(selected) ? selected : original.config["LastUsedProfile"]
                }
                origin := model.profileOrigins.Get(model.profile, model.profile)
                if !original.Profiles.Has(origin) {
                    baseline := args.Get("baseline", 0)
                    if !(baseline is Map) || !baseline.Has("_Profiles") || !baseline["_Profiles"].Has(model.profile)
                        throw ValueError("The category's initial profile values are unavailable.")
                    original := SettingsModel(baseline)
                    origin := model.profile
                }
                if section = "inheritance" {
                    model.Profiles[model.profile]["SettingsInheritance"] := SettingsModel.Clone(original.Profiles[origin]["SettingsInheritance"])
                    model.config["NewProfileInheritance"] := SettingsModel.Clone(original.config.Get("NewProfileInheritance", Map()))
                } else if SettingsModel.Units().Has(section) && model.Profiles[model.profile].Has("SettingsInheritance")
                    model.Profiles[model.profile]["SettingsInheritance"][section] := original.Profiles[origin]["SettingsInheritance"].Get(section, 0)
                for item in SettingsModel.Units().Get(section, [])
                    SettingsModel.Put(model.Profiles[model.profile], item.path, SettingsModel.Clone(SettingsModel.At(original.Profiles[origin], item.path, "")))
                for id, field in SettingsSchema.Fields
                    if field.section = section {
                        if field.global
                            SettingsModel.Put(model.config, field.path, SettingsModel.Clone(original.Get(id)))
                        else
                            SettingsModel.Put(model.Profiles[model.profile], field.path, SettingsModel.Clone(SettingsModel.At(original.Profiles[origin], field.path, "")))
                    }
                if section = "sounds" {
                    for key in ["SoundLibrary", "SoundVolumes", "MaxSimultaneousSounds", "SoundRepeatTimeout"]
                        model.config[key] := SettingsModel.Clone(original.config[key])
                    currentEvents := model.Profiles[model.profile]["Game Logs Monitoring"]["monitoredEvents"]
                    originalEvents := original.Profiles[origin]["Game Logs Monitoring"]["monitoredEvents"]
                    for id, event in currentEvents
                        if originalEvents.Has(id) {
                            event["sound"] := originalEvents[id].Get("sound", "")
                            event["soundVolume"] := originalEvents[id].Get("soundVolume", 100)
                        }
                }
            case "undo": model.UndoLast()
            case "validate": return this.session.Validate()
            case "apply":
                ; Check handoff under the same lock as save; runtime advertises
                ; before waiting to load, so independent writers cannot race.
                lock := SettingsFileLock(this.installation)
                try {
                    if this.role = "offline" && this.RuntimeWindow()
                        throw SettingsServiceError("runtime-started", "EVE-X-Preview has started. Reconnect and recover this draft before applying.")
                    result := this.session.Apply()
                    this.restartPending := this.role = "runtime"
                    return result
                } finally lock.Dispose()
            case "discard":
                pid := this.ExternalPid
                this.session.Dispose()
                this.session := SettingsSession(this.owner, 0, this.role = "runtime")
                this.ExternalPid := pid
            case "command": return this.session.Command(args.Get("name", ""), args.Get("arguments", Map()))
            case "end":
                if this.session.Dirty && !args.Get("discard", false)
                    throw SettingsServiceError("busy", "Apply or explicitly discard your pending settings before closing.")
                this.session.Dispose()
                this.session := 0
                this.ExternalPid := 0
                return Map("closed", JSON.true)
            default: throw SettingsServiceError("unknown-method", "Unknown settings bridge method: " method)
        }
        return this.session.Snapshot()
    }

    WriteResponse(id, response) {
        DirCreate(this.responseDir)
        serialized := JSON.Dump(SettingsContract.JsonFlags(response))
        if StrPut(serialized, "UTF-8") > SettingsBridge.MaxBytes {
            response := Map("version", 1, "id", id, "ok", JSON.false,
                "error", Map("code", "invalid-request", "message", "The settings response exceeds the 4 MiB bridge limit. Reduce the draft size before reconnecting.", "details", []))
            serialized := JSON.Dump(response)
        }
        final := this.responseDir "\" id ".json", temporary := final "." DllCall("GetCurrentProcessId") ".tmp"
        file := FileOpen(temporary, "w", "UTF-8-RAW")
        if !file
            throw OSError()
        try {
            file.Write(serialized)
            if !DllCall("FlushFileBuffers", "Ptr", file.Handle)
                throw OSError()
        } finally file.Close()
        FileMove(temporary, final, true)
    }

    Cleanup() {
        loop files this.responseDir "\*", "F" {
            try {
                if DateDiff(A_Now, A_LoopFileTimeModified, "Seconds") > 600
                    FileDelete(A_LoopFileFullPath)
            }
        }
    }

    RuntimeWindow() => DllCall("FindWindowW", "Ptr", 0, "Str", "EVE-X-Preview.SettingsBridge/1/" this.installationId "/runtime", "Ptr")

    CheckClient() {
        if A_TickCount >= this.nextCleanup {
            this.nextCleanup := A_TickCount + 60000
            try this.Cleanup()
        }
        if this.ExternalPid && !ProcessExist(this.ExternalPid) {
            if IsObject(this.session) && this.session.Dirty {
                ; Preserve orphaned canonical state in memory. A subsequent client
                ; must explicitly recover/discard it rather than replace it.
                this.orphaned := this.session.Recovery()
            } else if IsObject(this.session) {
                this.session.Dispose()
                this.session := 0
            }
            this.ExternalPid := 0
        }
    }

    FocusExternal() {
        if !this.HasExternal
            return false
        for hwnd in WinGetList("ahk_pid " this.ExternalPid) {
            if !DllCall("IsWindowVisible", "Ptr", hwnd)
                continue
            try {
                WinRestore(hwnd)
                WinActivate(hwnd)
                return true
            }
        }
        return false
    }

    TryOpenInternal() {
        this.CheckClient()
        if IsObject(this.session) && this.session.Dirty {
            if this.HasExternal {
                this.FocusExternal()
                return false
            }
            choice := this.ConfirmOrphanRecovery()
            if choice = "Cancel"
                return false
            if choice = "Yes" {
                this.owner.settingsRecoveredModel := this.session.model
                this.owner.settingsEditorStatus := "The disconnected modern editor's accepted settings draft was recovered in internal settings."
            }
            this.session.Dispose()
            this.session := 0
            this.ExternalPid := 0
            return true
        }
        if this.HasExternal {
            ; A clean client receives a normal close request and releases its lease.
            for hwnd in WinGetList("ahk_pid " this.ExternalPid)
                if DllCall("IsWindowVisible", "Ptr", hwnd)
                    PostMessage(0x10, 0, 0,, hwnd)
            deadline := A_TickCount + 1500
            while this.HasExternal && IsObject(this.session) && A_TickCount < deadline
                Sleep(25)
            this.CheckClient()
            if IsObject(this.session) && this.HasExternal {
                this.FocusExternal()
                ProgramLog.Add("Finish the modern settings dialog before opening internal settings.", "WARN")
                return false
            }
        }
        return true
    }

    CloseExternalForUpdate(timeout := 10000) {
        this.CheckClient()
        if IsObject(this.session) && this.session.Dirty {
            this.FocusExternal()
            return false
        }
        ; End releases the lease before Avalonia finishes shutting down. Keep
        ; waiting for this exact PID even after ExternalPid becomes zero.
        pid := this.ExternalPid
        deadline := A_TickCount + timeout
        nextClose := 0
        while pid && ProcessExist(pid) {
            if IsObject(this.session) && this.session.Dirty
                return false
            if A_TickCount >= deadline
                return false
            if A_TickCount >= nextClose {
                for hwnd in WinGetList("ahk_pid " pid)
                    if DllCall("IsWindowVisible", "Ptr", hwnd)
                        PostMessage(0x10, 0, 0,, hwnd)
                nextClose := A_TickCount + 250
            }
            Sleep(25)
        }
        this.CheckClient()
        return true
    }

    ConfirmOrphanRecovery() {
        return MsgBox("The modern settings editor closed while a draft was pending.`n`nYes: recover its accepted settings, undo history and profile changes in internal settings.`nNo: discard that draft and open the latest saved settings.`nCancel: keep the draft for a later modern-editor reconnect.`n`nUnfinished text held only by the closed editor cannot be recovered. No saved settings change until Apply.",
            "Recover disconnected settings draft", "YesNoCancel Icon?")
    }

    GuardRuntimeAction() {
        this.CheckClient()
        if !IsObject(this.session) || !this.session.Dirty
            return true
        this.FocusExternal()
        ProgramLog.Add("Runtime action postponed: apply or discard the open modern settings draft first.", "WARN")
        return false
    }

    RestartRuntime() {
        ; The response has already been atomically flushed and acknowledged.
        Reload()
    }

    Dispose(*) {
        if this.HasOwnProp("disposed") && this.disposed
            return
        this.disposed := true
        if IsObject(this.session)
            this.session.Dispose()
        if this.transport {
            SetTimer(this.poll, 0)
            OnMessage(0x4A, this.copyData, 0)
            this.window.Destroy()
            OnExit(this.exitCallback, 0)
        }
    }
}
