TestRunner.Register("Settings custom themes round trip through the canonical global selector", SettingsCustomThemesRoundTripTest)
SettingsCustomThemesRoundTripTest() {
    model := SettingsModel(JSON.Load(default_JSON))
    for id in ["matrix", "kitsune", "synthwave", "cyberpunk"] {
        model.Set("SettingsTheme", id)
        AssertEqual(id, model.Get("SettingsTheme"))
        AssertEqual(id, model.Merge(JSON.Load(default_JSON))["SettingsTheme"])
    }
    AssertTrue(SettingsSchema.Fields["SettingsTheme"].global)
    AssertThrows(() => model.Set("SettingsTheme", "unknown-theme"))
}
TestRunner.Register("Settings custom themes share palette metadata between both editors", SettingsCustomThemesMetadataTest)
SettingsCustomThemesMetadataTest() {
    metadata := SettingsContract.Export()
    names := Map()
    for theme in metadata["themes"] {
        names[theme["id"]] := theme["label"]
        for key in ["background", "sidebar", "input", "foreground", "muted", "line", "accent", "selection"]
            AssertTrue(RegExMatch(theme[key], "^[0-9A-F]{6}$"))
    }
    AssertEqual("The Matrix", names["matrix"])
    AssertEqual("Kitsune", names["kitsune"])
    AssertEqual("Synthwave", names["synthwave"])
    AssertEqual("Cyberpunk", names["cyberpunk"])
}
TestRunner.Register("Settings custom themes repaint between dark palettes without changing thumbnail colors or drafts", SettingsCustomThemesNativeTest)
SettingsCustomThemesNativeTest() {
    app := SettingsWorkbenchFixture()
    try {
        thumbnailColor := app.model.Get("ThumbnailTextColor")
        app.Navigate("interface")
        for id in ["matrix", "kitsune", "synthwave", "cyberpunk", "matrix"] {
            AssertTrue(app.theme.SetMode(id), "Different dark palettes must repaint immediately.")
            AssertTrue(app.theme.dark)
            AssertEqual(SettingsPalettes.Items[id]["background"], app.theme.background)
            AssertEqual(SettingsPalettes.Items[id]["input"], app.theme.input)
            AssertEqual(thumbnailColor, app.model.Get("ThumbnailTextColor"))
            AssertFalse(app.model.Dirty)
        }
        AssertEqual("041008", app.theme.background)
        app.theme.SetMode("light")
        AssertFalse(app.theme.dark)
    } finally app.Dispose()
}
