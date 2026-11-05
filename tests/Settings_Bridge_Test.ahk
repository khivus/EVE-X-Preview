TestRunner.Register("Settings bridge installation identity is independent of case and trailing slashes", SettingsBridgeIdentityTest)
TestRunner.Register("Settings category undo preserves other categories and supports new profiles", SettingsCategoryUndoTest)
SettingsCategoryUndoTest() {
    bridge := SettingsBridge(SettingsServiceFixture(JSON.Load(default_JSON)), "runtime", A_ScriptDir, false)
    try {
        bridge.Begin(Map("clientPid", ProcessExist()))
        model := bridge.session.model
        originalSize := model.Get("ThumbnailStartLocationwidth")
        model.Set("ThumbnailTextSize", 23)
        model.Set("ThumbnailStartLocationwidth", originalSize + 100)
        bridge.Dispatch("undo-category", Map("section", "layout"))
        AssertEqual(originalSize, model.Get("ThumbnailStartLocationwidth"))
        AssertEqual(23, model.Get("ThumbnailTextSize"))
        model.AddProfile("New category fixture")
        initial := SettingsModel.Clone(model.config)
        model.Set("ThumbnailStartLocationwidth", originalSize + 200)
        bridge.Dispatch("undo-category", Map("section", "layout", "baseline", initial))
        AssertEqual(originalSize, model.Get("ThumbnailStartLocationwidth"))
        AssertEqual(23, model.Get("ThumbnailTextSize"))
        before := SettingsModel.Clone(bridge.session.Recovery())
        model.Set("ThumbnailTextSize", 27)
        bridge.Dispatch("restore-draft", Map("draft", before))
        AssertEqual(23, bridge.session.model.Get("ThumbnailTextSize"))
        AssertTrue(bridge.session.model.Profiles.Has("New category fixture"))
        model := bridge.session.model
        model.Set("MasterSoundVolume", 175)
        model.config["SoundLibrary"].Push("fixture.wav")
        bridge.Dispatch("undo-category", Map("section", "sounds", "baseline", initial))
        AssertEqual(100, model.Get("MasterSoundVolume"))
        AssertEqual(0, model.config["SoundLibrary"].Length)
        model.Group("events", "Game Logs Monitoring")["monitoredEvents"]["stoppedShooting"]["soundVolume"] := 180
        model.Set("SoundRepeatTimeout", 2500)
        bridge.Dispatch("undo-category", Map("section", "sounds", "baseline", initial))
        AssertEqual(2000, model.Get("SoundRepeatTimeout"))
        AssertEqual(100, model.Group("events", "Game Logs Monitoring")["monitoredEvents"]["stoppedShooting"]["soundVolume"])
        bridge.Dispatch("undo-category", Map("section", "profiles"))
        AssertFalse(model.Profiles.Has("New category fixture"))
        AssertEqual(23, model.Profiles["Default"]["Thumbnails Visuals"]["ThumbnailTextSize"])
    } finally bridge.Dispose()
}
SettingsBridgeIdentityTest() {
    first := SettingsBridge.InstallationId(A_ScriptDir)
    AssertTrue(RegExMatch(first, "^[a-f0-9]{64}$"))
    AssertEqual(first, SettingsBridge.InstallationId(StrUpper(A_ScriptDir) "\\"))
    AssertNotEqual(first, SettingsBridge.InstallationId(A_ScriptDir "\\other"))
}

TestRunner.Register("Settings bridge validates envelopes lease identity and pending client drafts", SettingsBridgeEnvelopeTest)
SettingsBridgeEnvelopeTest() {
    bridge := SettingsBridge(SettingsServiceFixture(JSON.Load(default_JSON)), "runtime", A_ScriptDir, false)
    AssertEqual("incompatible", bridge.Handle(Map("version", 2, "id", Format("{:032x}", 1), "method", "hello", "session", "", "arguments", Map()))["error"]["code"])
    hello := bridge.Handle(Map("version", 1, "id", Format("{:032x}", 2), "method", "hello", "session", "", "arguments", Map()))
    AssertTrue(hello["ok"])
    begin := bridge.Handle(Map("version", 1, "id", Format("{:032x}", 3), "method", "begin", "session", "", "arguments", Map("clientPid", ProcessExist())))
    AssertTrue(begin["ok"])
    session := begin["data"]["session"]
    response := bridge.Handle(Map("version", 1, "id", Format("{:032x}", 4), "method", "client-state", "session", session, "arguments", Map("dirty", true)))
    AssertTrue(response["ok"])
    AssertTrue(bridge.session.Dirty)
    response := bridge.Handle(Map("version", 1, "id", Format("{:032x}", 5), "method", "end", "session", session, "arguments", Map("discard", false)))
    AssertEqual("busy", response["error"]["code"])
    response := bridge.Handle(Map("version", 1, "id", Format("{:032x}", 6), "method", "snapshot", "session", "stale", "arguments", Map()))
    AssertEqual("stale-session", response["error"]["code"])
    bridge.Dispose()
}

TestRunner.Register("Settings bridge guards unfinished and orphaned drafts until explicit release", SettingsBridgeLeaseGuardTest)
SettingsBridgeLeaseGuardTest() {
    bridge := SettingsBridge(SettingsServiceFixture(JSON.Load(default_JSON)), "runtime", A_ScriptDir, false)
    try {
        snapshot := bridge.Begin(Map("clientPid", ProcessExist()))
        bridge.session.clientDirty := true
        AssertFalse(bridge.GuardRuntimeAction())
        AssertFalse(bridge.TryOpenInternal())
        session := bridge.session.id
        bridge.ExternalPid := 0
        recovered := bridge.Begin(Map("clientPid", ProcessExist()))
        AssertEqual(session, recovered["session"])
        AssertTrue(recovered["clientDirty"])
        AssertThrows(() => bridge.Dispatch("end", Map("discard", false)), "pending")
        AssertTrue(bridge.Dispatch("end", Map("discard", true))["closed"])
        AssertTrue(bridge.GuardRuntimeAction())
        AssertTrue(bridge.TryOpenInternal())
    } finally bridge.Dispose()
}

TestRunner.Register("Settings bridge orphan recovery can transfer canonical model to internal editor", SettingsBridgeOrphanRecoveryTest)
SettingsBridgeOrphanRecoveryTest() {
    owner := SettingsWorkbenchOwnerFixture()
    bridge := SettingsOrphanBridgeFixture(owner, "runtime", A_ScriptDir, false)
    try {
        bridge.Begin(Map("clientPid", ProcessExist()))
        bridge.session.model.Set("SettingsTheme", "dark")
        savedModel := bridge.session.model
        bridge.ExternalPid := 0
        bridge.choice := "Cancel"
        AssertFalse(bridge.TryOpenInternal())
        AssertTrue(bridge.session.Dirty)
        bridge.choice := "Yes"
        AssertTrue(bridge.TryOpenInternal())
        AssertTrue(owner.settingsRecoveredModel = savedModel)
        app := SettingsWorkbenchFixture(owner)
        try {
            AssertEqual("dark", app.model.Get("SettingsTheme"))
            AssertTrue(app.model.Dirty)
            AssertFalse(owner.HasOwnProp("settingsRecoveredModel"))
        } finally app.Dispose()
    } finally bridge.Dispose()
}

class SettingsOrphanBridgeFixture extends SettingsBridge {
    choice := "Cancel"
    ConfirmOrphanRecovery() => this.choice
}

TestRunner.Register("Settings bridge real hidden COPYDATA uses UTF8 atomic responses and rejects malformed payloads", SettingsBridgeCopyDataTest)
SettingsBridgeCopyDataTest() {
    owner := SettingsServiceFixture(JSON.Load(default_JSON))
    bridge := SettingsBridge(owner, "offline", A_Temp "\bridge-" SettingsBridge.NewId())
    owner.settingsBridge := bridge
    sender := Gui("+ToolWindow", "Settings bridge test sender")
    senderHwnd := sender.Hwnd
    try {
        request := Map("version", 1, "id", SettingsBridge.NewId(), "method", "begin", "session", "", "arguments", Map("clientPid", ProcessExist()))
        response := SettingsBridgeTestSend(bridge, senderHwnd, JSON.Dump(request))
        AssertTrue(response["ok"])
        session := response["data"]["session"]
        request := Map("version", 1, "id", SettingsBridge.NewId(), "method", "profile", "session", session, "arguments", Map("action", "add", "name", "Fleet α"))
        response := SettingsBridgeTestSend(bridge, senderHwnd, JSON.Dump(request))
        AssertEqual("Fleet α", response["data"]["profile"])
        repeated := SettingsBridgeTestSend(bridge, senderHwnd, JSON.Dump(request))
        AssertDeepEqual(response, repeated)
        AssertEqual(2, bridge.session.model.Profiles.Count)
        originalDir := bridge.responseDir
        blockedPath := originalDir "\blocked-" SettingsBridge.NewId()
        FileAppend("occupied", blockedPath)
        bridge.responseDir := blockedPath
        request := Map("version", 1, "id", SettingsBridge.NewId(), "method", "profile", "session", session, "arguments", Map("action", "add", "name", "Recovered β"))
        AssertThrows(() => SettingsBridgeTestSend(bridge, senderHwnd, JSON.Dump(request)), "rejected")
        bridge.responseDir := originalDir
        FileDelete(blockedPath)
        response := SettingsBridgeTestSend(bridge, senderHwnd, JSON.Dump(request))
        AssertEqual("Recovered β", response["data"]["profile"])
        AssertEqual(3, bridge.session.model.Profiles.Count)
        AssertThrows(() => SettingsBridgeTestSend(bridge, senderHwnd, "{broken"), "rejected")
        request := Map("version", 1, "id", SettingsBridge.NewId(), "method", "begin", "session", "", "arguments", Map("clientPid", 4))
        response := SettingsBridgeTestSend(bridge, senderHwnd, JSON.Dump(request))
        AssertEqual("invalid-request", response["error"]["code"])
    } finally {
        sender.Destroy()
        bridge.Dispose()
    }
}

SettingsBridgeTestSend(bridge, senderHwnd, jsonText) {
    payloadBuffer := Buffer(StrPut(jsonText, "UTF-8"), 0)
    StrPut(jsonText, payloadBuffer, "UTF-8")
    copyData := Buffer(A_PtrSize * 3, 0)
    NumPut("UPtr", SettingsBridge.Magic, copyData, 0)
    NumPut("UInt", payloadBuffer.Size, copyData, A_PtrSize)
    NumPut("Ptr", payloadBuffer.Ptr, copyData, A_PtrSize * 2)
    accepted := SendMessage(0x4A, senderHwnd, copyData.Ptr,, bridge.window.Hwnd)
    if !accepted
        throw Error("Settings bridge rejected the transport payload.")
    request := JSON.Load(jsonText)
    path := bridge.responseDir "\" request["id"] ".json"
    response := JSON.Load(FileRead(path, "UTF-8"))
    FileDelete(path)
    return response
}
