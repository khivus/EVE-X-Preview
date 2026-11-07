TestRunner.Register("Settings contract exports every scalar and canonical specialized workflow", SettingsContractMetadataTest)
TestRunner.Register("Settings canonical accepts the first exported installed font and keeps update labels canonical", SettingsFontAndVersionMetadataTest)
TestRunner.Register("Settings canonical preview commands expose the current pause state and matching action", SettingsPreviewPauseStateTest)
SettingsPreviewPauseStateTest() {
    owner := SettingsPauseStateFixture(JSON.Load(default_JSON))
    session := SettingsSession(owner, 0, true)
    try {
        status := session.Command("live-preview-status")
        AssertFalse(status["paused"])
        AssertEqual("Pause live previews", status["label"])
        status := session.Command("toggle-live-previews")
        AssertTrue(status["paused"])
        AssertEqual("Resume live previews", status["label"])
        AssertFalse(session.Command("toggle-live-previews")["paused"])
    } finally session.Dispose()
}
class SettingsPauseStateFixture extends SettingsServiceFixture {
    livePreviewsPaused := false
    ToggleLivePreviews() => this.livePreviewsPaused := !this.livePreviewsPaused
}
TestRunner.Register("Settings canonical defaults use an installed font and only migrate the missing legacy factory font", SettingsFontDefaultMigrationTest)
SettingsFontDefaultMigrationTest() {
    config := JSON.Load(default_JSON)
    AssertEqual("Segoe UI", config["_Profiles"]["Default"]["Thumbnails Visuals"]["ThumbnailTextFont"])
    AssertTrue(SettingsFonts.Available("Segoe UI"))
    config["_Profiles"]["Default"]["Thumbnails Visuals"]["ThumbnailTextFont"] := "Gill Sans MT"
    config["_Profiles"]["Custom"] := SettingsModel.Clone(config["_Profiles"]["Default"])
    config["_Profiles"]["Custom"]["Thumbnails Visuals"]["ThumbnailTextFont"] := "My missing custom font"
    SettingsFonts.NormalizeLegacyDefaults(config, ["Arial", "Segoe UI"])
    AssertEqual("Segoe UI", config["_Profiles"]["Default"]["Thumbnails Visuals"]["ThumbnailTextFont"])
    AssertEqual("My missing custom font", config["_Profiles"]["Custom"]["Thumbnails Visuals"]["ThumbnailTextFont"])
    config["_Profiles"]["Default"]["Thumbnails Visuals"]["ThumbnailTextFont"] := "Gill Sans MT"
    SettingsFonts.NormalizeLegacyDefaults(config, ["Gill Sans MT", "Segoe UI"])
    AssertEqual("Gill Sans MT", config["_Profiles"]["Default"]["Thumbnails Visuals"]["ThumbnailTextFont"])
    SettingsFonts.NormalizeLegacyDefaults(config, ["Arial"])
    AssertEqual("Arial", config["_Profiles"]["Default"]["Thumbnails Visuals"]["ThumbnailTextFont"])
    config["_Profiles"]["Default"]["Thumbnails Visuals"]["ThumbnailTextFont"] := "Gill Sans MT"
    folder := A_Temp "\EveX-font-migration-" SettingsBridge.NewId(), previousDirectory := A_WorkingDir
    DirCreate(folder)
    try {
        SetWorkingDir(folder)
        FileAppend(JSON.Dump(config), "EVE-X-Preview.json", "UTF-8")
        loaded := Load_JSONUnlocked(false)
        expected := SettingsFonts.Available("Gill Sans MT") ? "Gill Sans MT" : SettingsFonts.DefaultName()
        AssertEqual(expected, loaded["_Profiles"]["Default"]["Thumbnails Visuals"]["ThumbnailTextFont"])
        AssertEqual(expected, JSON.Load(FileRead("EVE-X-Preview.json"))["_Profiles"]["Default"]["Thumbnails Visuals"]["ThumbnailTextFont"])
        AssertEqual("My missing custom font", loaded["_Profiles"]["Custom"]["Thumbnails Visuals"]["ThumbnailTextFont"])
    } finally {
        SetWorkingDir(previousDirectory)
        DirDelete(folder, true)
    }
}
SettingsFontAndVersionMetadataTest() {
    names := SettingsContract.Export()["fonts"]
    AssertTrue(names.Length > 0)
    session := SettingsSession(SettingsServiceFixture(JSON.Load(default_JSON)), 0, false)
    try {
        session.model.Set("ThumbnailTextFont", names[1])
        AssertEqual(names[1], session.model.Get("ThumbnailTextFont"))
        for issue in session.Validate()["issues"]
            AssertFalse(issue["id"] = "ThumbnailTextFont", issue["message"])
        status := session.Command("update-status")
        AssertEqual("Install stable version (check first)", status["stableLabel"])
        AssertFalse(status["stableAvailable"])
        facade := session.Facade()
        facade.updateChecker.releaseTag := "0.0.1"
        status := session.Command("update-status")
        AssertTrue(InStr(status["stableLabel"], "Downgrade to stable v0.0.1"))
        AssertTrue(status["stableAvailable"])
    } finally session.Dispose()
}
TestRunner.Register("Settings companion preview opens offline and follows validated draft changes", SettingsCompanionPreviewTest)
SettingsCompanionPreviewTest() {
    session := SettingsSession(SettingsServiceFixture(JSON.Load(default_JSON)), 0, false)
    try {
        session.Command("test-thumbnail")
        facade := session.Facade()
        preview := facade.settingsThumbnail
        AssertTrue(IsObject(preview))
        AssertTrue(DllCall("IsWindowVisible", "Ptr", preview.sample["Window"].Hwnd), "The demo must have a visible owner and appear on screen.")
        session.model.Set("ThumbnailTextSize", 23)
        session.Snapshot()
        preview.Refresh(A_TickCount + 1000)
        AssertEqual(23, preview.ThumbnailTextSize)
        session.Command("stop-test-thumbnail")
        AssertFalse(IsObject(facade.settingsThumbnail))
    } finally session.Dispose()
}
SettingsContractMetadataTest() {
    contract := SettingsContract.Export()
    AssertEqual(1, contract["bridgeVersion"])
    AssertEqual(1, contract["schemaVersion"])
    AssertEqual(SettingsSchema.Fields.Count, contract["fields"].Length)
    ids := Map()
    for field in contract["fields"] {
        ids[field["id"]] := true
        AssertTrue(field.Has("default"))
    }
    AssertTrue(ids.Has("SettingsEditor"))
    AssertEqual("automatic", SettingsModel(JSON.Load(default_JSON)).Get("SettingsEditor"))
    AssertEqual(10, contract["tables"].Length)
    AssertEqual(SettingsEventTable.Names.Count, contract["events"].Length)
    AssertTrue(contract["commands"].Length >= 10)
}

TestRunner.Register("Settings canonical session recovery preserves rename origins unknown values and undo", SettingsContractRecoveryTest)
SettingsContractRecoveryTest() {
    config := JSON.Load(default_JSON)
    config["_Profiles"]["Fleet"] := SettingsModel.Clone(config["_Profiles"]["Default"])
    config["LastUsedProfile"] := "Fleet"
    config["future-extension"] := Map("nested", ["untouched", 42])
    owner := SettingsServiceFixture(config)
    session := SettingsSession(owner)
    session.model.RenameProfile("Logistics")
    session.model.Set("ThumbnailTextSize", 18)
    recovered := SettingsSession(owner, session.Recovery())
    owner._JSON["_Profiles"]["Fleet"]["Thumbnail Positions"]["EVE - Late"] := [1, 2, 3, 4]
    AssertEqual("Logistics", recovered.model.profile)
    AssertEqual("Fleet", recovered.model.profileOrigins["Logistics"])
    AssertDeepEqual([1, 2, 3, 4], recovered.model.Merge(owner._JSON)["_Profiles"]["Logistics"]["Thumbnail Positions"]["EVE - Late"])
    AssertDeepEqual(config["future-extension"], recovered.model.config["future-extension"])
}

TestRunner.Register("Settings canonical apply detects conflicts and rolls back failed saves", SettingsContractApplyTest)
SettingsContractApplyTest() {
    owner := SettingsServiceFixture(JSON.Load(default_JSON))
    session := SettingsSession(owner)
    session.model.Set("ThumbnailTextSize", 18)
    owner._JSON["_Profiles"]["Default"]["Thumbnails Visuals"]["ThumbnailTextSize"] := 19
    AssertFalse(session.Validate()["canApply"])
    AssertTrue(session.Validate()["conflicts"].Length > 0)
    AssertThrows(() => session.Apply(), "changed")
    owner._JSON["_Profiles"]["Default"]["Thumbnails Visuals"]["ThumbnailTextSize"] := 12
    owner._JSON["_Profiles"]["Default"]["Thumbnail Positions"]["EVE - Late"] := [1, 2, 3, 4]
    before := JSON.Dump(owner._JSON)
    owner.failSave := true
    AssertThrows(() => session.Apply(), "fixture save failure")
    AssertEqual(before, JSON.Dump(owner._JSON))
    AssertTrue(session.model.Dirty)
    owner.failSave := false
    AssertTrue(session.Apply()["applied"])
    AssertEqual(18, owner._JSON["_Profiles"]["Default"]["Thumbnails Visuals"]["ThumbnailTextSize"])
    AssertDeepEqual([1, 2, 3, 4], owner._JSON["_Profiles"]["Default"]["Thumbnail Positions"]["EVE - Late"])
}

TestRunner.Register("Settings canonical specialized operations preserve unknown event values and group undo", SettingsContractSpecializedTest)
SettingsContractSpecializedTest() {
    owner := SettingsServiceFixture(JSON.Load(default_JSON))
    session := SettingsSession(owner)
    session.model.Group("events", "Game Logs Monitoring")["monitoredEvents"]["underAttackByPlayer"]["future"] := 42
    session.Event("underAttackByPlayer", Map("enabled", 1, "includeMisses", 0, "color", "#ff0010"))
    event := session.Event("underAttackByPlayer")
    AssertEqual(42, event["values"]["future"])
    AssertEqual("FF0010", event["values"]["color"])
    session.Group(Map("kind", "characters", "action", "add", "name", "Fleet"))
    session.Group(Map("kind", "characters", "action", "shortcut", "name", "Fleet", "key", "ForwardsHotkey", "value", "^F2"))
    session.Rows("group-members", "Fleet", [["Pilot α"]])
    key := ""
    for shortcut in session.Rows("group-members", "Fleet")["groupShortcuts"]
        if shortcut["key"] = "ForwardsHotkey"
            key := shortcut["value"]
    AssertEqual("^F2", key)
    session.Group(Map("kind", "characters", "action", "delete", "name", "Fleet"))
    AssertTrue(session.model.UndoLast())
    AssertDeepEqual([["Pilot α"]], session.model.Rows("group-members", "Fleet"))
}

class SettingsServiceFixture {
    failSave := false
    __New(config) => this._JSON := SettingsModel.Clone(config)
    SaveJsonToFile(*) {
        if this.failSave
            throw Error("fixture save failure")
    }
}

TestRunner.Register("Settings canonical recovery rejects cross-installation drafts", SettingsContractRecoveryIsolationTest)
SettingsContractRecoveryIsolationTest() {
    owner := SettingsServiceFixture(JSON.Load(default_JSON))
    state := SettingsSession(owner).Recovery()
    state["installation"] := Format("{:064x}", 1)
    AssertThrows(() => SettingsSession(owner, state), "another")
}

TestRunner.Register("Settings canonical table replacement preserves unknown extension keys", SettingsContractUnknownTableTest)
SettingsContractUnknownTableTest() {
    model := SettingsModel(JSON.Load(default_JSON))
    profile := model.Profiles["Default"]
    profile["Thumbnails Visuals"]["CustomThumbnailNames"] := Map("Characters", "", "Names", "", "future", 42)
    profile["Custom Colors"]["cColors"]["future"] := 43
    profile["Non-EVE Applications"]["NonEVEHotkeys"]["future"] := 44
    profile["Thumbnails Interactions"]["ActivateThumbnail"]["future"] := 45
    model.SaveRows("custom-names", [["Pilot", "Scout"]])
    model.SaveRows("character-colors", [["Pilot", "FFFFFF", "000000", "FF0000"]])
    model.SaveRows("app-shortcuts", [["notepad.exe", "", "^F2"]])
    model.SaveRows("gestures", model.Rows("gestures"))
    AssertEqual(42, profile["Thumbnails Visuals"]["CustomThumbnailNames"]["future"])
    AssertEqual(43, profile["Custom Colors"]["cColors"]["future"])
    AssertEqual(44, profile["Non-EVE Applications"]["NonEVEHotkeys"]["future"])
    AssertEqual(45, profile["Thumbnails Interactions"]["ActivateThumbnail"]["future"])
}

TestRunner.Register("Settings canonical unrelated edits tolerate unchanged unavailable legacy fonts and folders", SettingsContractLegacyValidationTest)
SettingsContractLegacyValidationTest() {
    config := JSON.Load(default_JSON)
    config["_Profiles"]["Default"]["Thumbnails Visuals"]["ThumbnailTextFont"] := "Missing fixture font"
    config["_Profiles"]["Default"]["Game Logs Monitoring"]["gameLogsDirectory"] := A_Temp "\missing-" SettingsBridge.NewId()
    owner := SettingsServiceFixture(config)
    session := SettingsSession(owner)
    session.model.Set("SettingsTheme", "dark")
    validation := session.Validate()
    AssertTrue(validation["canApply"])
    AssertTrue(validation["issues"].Length >= 2)
    AssertTrue(session.Apply()["applied"])
    AssertEqual("dark", owner._JSON["SettingsTheme"])
    AssertThrows(() => session.model.Set("gameLogsDirectory", A_Temp "\another-missing-" SettingsBridge.NewId()), "existing folder")
    state := session.Recovery()
    state["config"]["_Profiles"]["Default"]["Game Logs Monitoring"]["gameLogsDirectory"] := A_Temp "\invalid-recovered-" SettingsBridge.NewId()
    recovered := SettingsSession(owner, state)
    AssertFalse(recovered.Validate()["canApply"])
}

TestRunner.Register("Settings canonical internal event dialog and bridge share validation and optional references", SettingsContractInternalEventParityTest)
SettingsContractInternalEventParityTest() {
    owner := SettingsWorkbenchOwnerFixture()
    missing := A_Temp "\missing-event-" SettingsBridge.NewId() ".wav"
    event := owner._JSON["_Profiles"]["Default"]["Game Logs Monitoring"]["monitoredEvents"]["underAttackByPlayer"]
    event["sound"] := missing, event["future-event-option"] := Map("keep", 42)
    app := SettingsWorkbenchFixture(owner)
    bridgeSession := SettingsSession(owner)
    try {
        app.Navigate("events")
        editor := SettingsEventEditor(app.eventTable, "underAttackByPlayer", 2)
        editor.Show()
        AssertEqual(3, editor.eventInputs.Count)
        editor.color.Value := "rgb(1, 2, 3)"
        editor.volume.Value := "145"
        editor.neuts.Value := 1, editor.misses.Value := 0, editor.smartbomb.Value := 0
        editor.Save()
        AssertFalse(app.HasOwnProp("activeDialog"), app.HasOwnProp("activeDialog") ? editor.error.Value : "")
        bridgeSession.Event("underAttackByPlayer", Map("color", "rgb(1, 2, 3)", "soundVolume", "145", "includeNeutralization", 1, "includeMisses", 0, "ignoreSmartbombDamage", 0))
        expected := bridgeSession.model.Group("events", "Game Logs Monitoring")["monitoredEvents"]["underAttackByPlayer"]
        actual := app.model.Group("events", "Game Logs Monitoring")["monitoredEvents"]["underAttackByPlayer"]
        AssertDeepEqual(expected, actual)
        AssertEqual(missing, actual["sound"])
        AssertEqual(42, actual["future-event-option"]["keep"])
        before := JSON.Dump(actual)
        editor := SettingsEventEditor(app.eventTable, "underAttackByPlayer", 2)
        editor.Show()
        editor.volume.Value := "301"
        editor.Save()
        AssertTrue(app.HasOwnProp("activeDialog"))
        AssertEqual(before, JSON.Dump(app.model.Group("events", "Game Logs Monitoring")["monitoredEvents"]["underAttackByPlayer"]))
        AssertThrows(() => bridgeSession.Event("underAttackByPlayer", Map("soundVolume", "301")), "0 to 300")
    } finally app.Dispose()
}
TestRunner.Register("Settings change counts track distinct values and return to zero after undo", SettingsChangeCountTest)
SettingsChangeCountTest() {
    model := SettingsModel(JSON.Load(default_JSON))
    AssertEqual(0, model.ChangeCount)
    original := model.Get("ThumbnailOpacity")
    model.Set("ThumbnailOpacity", original - 1)
    AssertEqual(1, model.ChangeCount)
    model.Set("ThumbnailOpacity", original - 2)
    AssertEqual(1, model.ChangeCount, "Repeated edits to one setting count once.")
    model.Set("ThumbnailTextSize", model.Get("ThumbnailTextSize") + 1)
    AssertEqual(2, model.ChangeCount)
    model.Set("ThumbnailOpacity", original)
    AssertEqual(1, model.ChangeCount)
    AssertEqual("character-shortcuts", HotkeyBindings.SectionFor("character:2"))
    AssertEqual("character-groups", HotkeyBindings.SectionFor("group:Fleet:ForwardsHotkey"))
    AssertEqual("profiles", HotkeyBindings.SectionFor("profile:Mining"))
    AssertEqual(0, JSON.Load(default_JSON)["_Profiles"]["Default"]["Tray Menu Settings"]["TrayMenuShortcuts"]["Pause Live Previews"])
    AssertTrue(SettingsSchema.Fields.Has("TrayCommand:Pause Live Previews"), "Both editors receive the optional tray command from the schema.")
}
