TestRunner.Register("Settings editor automatic and modern use compatible companion with reliable internal fallback", SettingsEditorRoutingTest)
SettingsEditorRoutingTest() {
    for preference in ["automatic", "modern"] {
        editor := SettingsEditorRoutingFixture(preference)
        AssertTrue(editor.OpenPreferred())
        AssertEqual(1, editor.launches)
        AssertEqual(0, editor.internal)
        editor := SettingsEditorRoutingFixture(preference)
        editor.available := false
        AssertTrue(editor.OpenPreferred())
        AssertEqual(1, editor.internal)
        AssertEqual(0, editor.launches)
        AssertTrue(editor.owner.settingsEditorStatus != "")
        editor := SettingsEditorRoutingFixture(preference)
        editor.probeResult["schemaVersion"] := 999
        AssertTrue(editor.OpenPreferred())
        AssertEqual(1, editor.internal)
        AssertEqual(0, editor.launches)
        editor := SettingsEditorRoutingFixture(preference)
        editor.failLaunch := true
        AssertTrue(editor.OpenPreferred())
        AssertEqual(1, editor.internal)
    }
}
TestRunner.Register("Settings editor built-in preference bypasses companion probes and reuses existing external editor", SettingsEditorBuiltInTest)
SettingsEditorBuiltInTest() {
    editor := SettingsEditorRoutingFixture("builtin")
    AssertTrue(editor.OpenPreferred())
    AssertEqual(0, editor.probes)
    AssertEqual(1, editor.internal)
    editor := SettingsEditorRoutingFixture("modern")
    editor.external := true
    AssertTrue(editor.OpenPreferred())
    AssertEqual(0, editor.launches)
    AssertEqual(0, editor.probes)
    editor := SettingsEditorRoutingFixture("modern")
    editor.draft := true
    AssertFalse(editor.OpenPreferred())
    AssertEqual(0, editor.launches)
    AssertEqual(0, editor.internal)
}
TestRunner.Register("Settings editor compatibility requires explicit application bridge and schema handshake", SettingsEditorCompatibilityTest)
SettingsEditorCompatibilityTest() {
    AssertFalse(SettingsEditor.Compatible(Map()))
    AssertFalse(SettingsEditor.Compatible(Map("application", "Wrong", "bridgeVersion", 1, "schemaVersion", 1)))
    AssertFalse(SettingsEditor.Compatible(Map("application", "EVE-X-Preview-Settings", "bridgeVersion", 2, "schemaVersion", 1)))
    AssertTrue(SettingsEditor.Compatible(Map("application", "EVE-X-Preview-Settings", "bridgeVersion", 1, "schemaVersion", 1)))
}
class SettingsEditorRoutingFixture extends SettingsEditor {
    available := true
    external := false
    draft := false
    launches := 0
    internal := 0
    probes := 0
    failLaunch := false
    __New(preference) {
        super.__New({_JSON: Map("SettingsEditor", preference)})
        this.probeResult := Map("application", "EVE-X-Preview-Settings", "bridgeVersion", 1, "schemaVersion", 1)
    }
    Exists(path) => this.available
    ExistingExternal() => this.external
    InternalDraft() => this.draft
    CloseCleanInternal() {
        return
    }
    BuiltIn() {
        this.internal++
        return true
    }
    Probe(path) {
        this.probes++
        return this.probeResult
    }
    Launch(path) {
        if this.failLaunch
            throw Error("Fixture launch failure")
        this.launches++
    }
}
