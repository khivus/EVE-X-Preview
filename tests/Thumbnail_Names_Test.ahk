; Exercise display names without starting EVE monitoring or loading user settings.
class ThumbnailNamesFixture extends Main_Class {
    __New() {
        this._JSON := Map("_Profiles", Map("Default", Map("Thumbnails Visuals", Map()), "Other", Map("Thumbnails Visuals", Map())))
        this.ProfileThumbnailsVisuals := "Default"
        this.ProfileGameLogsMonitoring := "Default"
        this._JSON["_Profiles"]["Default"]["Game Logs Monitoring"] := Map()
    }
}

TestRunner.Register("Thumbnail names preserve blank and unmatched rows", ThumbnailNamesRowsTest)
ThumbnailNamesRowsTest() {
    app := ThumbnailNamesFixture()
    AssertEqual("Alice", app.GetThumbnailDisplayText("EVE - Alice"))
    app.CustomThumbnailNames := Map("Characters", "Alice`n`nBob`nCarol`nDave", "Names", "Scout & 1`nUnused`n`nMiner")
    AssertEqual("Scout & 1", app.GetThumbnailDisplayText("EVE - Alice"))
    AssertEqual("Bob", app.GetThumbnailDisplayText("EVE - Bob"))
    AssertEqual("Miner", app.GetThumbnailDisplayText("EVE - Carol"))
    AssertEqual("Dave", app.GetThumbnailDisplayText("EVE - Dave"))
    AssertEqual("Alice Two", app.GetThumbnailDisplayText("EVE - Alice Two"))
    AssertEqual("", app.GetThumbnailDisplayText("EVE"))
}

TestRunner.Register("Thumbnail aliases preserve event text and literal replacements", ThumbnailNamesEventsTest)
ThumbnailNamesEventsTest() {
    app := ThumbnailNamesFixture()
    app.CustomThumbnailNames := Map("Characters", " EVE - Alice `r`nBob", "Names", "EVE - Scout & <1>`r`nSupport")
    AssertEqual("EVE - Scout & <1>`nUnder attack`nAlice", app.GetThumbnailDisplayText("EVE - Alice`nUnder attack`nAlice"))
    AssertEqual("Support", app.GetThumbnailDisplayText("EVE - bob"))
}

TestRunner.Register("Thumbnail aliases survive JSON and stay scoped to visuals profile", ThumbnailNamesProfilesTest)
ThumbnailNamesProfilesTest() {
    app := ThumbnailNamesFixture()
    app.CustomThumbnailNames := Map("Characters", "Alice`n`nBob", "Names", "Scout`n`nSupport")
    app._JSON := JSON.Load(JSON.Dump(app._JSON))
    AssertEqual("Scout", app.GetThumbnailDisplayText("EVE - Alice"))
    AssertEqual("Support", app.GetThumbnailDisplayText("EVE - Bob"))
    app.ProfileThumbnailsVisuals := "Other"
    AssertEqual("Alice", app.GetThumbnailDisplayText("EVE - Alice"))
    app.CustomThumbnailNames := Map("Characters", "Alice", "Names", "Miner")
    AssertEqual("Miner", app.GetThumbnailDisplayText("EVE - Alice"))
    app.ProfileThumbnailsVisuals := "Default"
    AssertEqual("Scout", app.GetThumbnailDisplayText("EVE - Alice"))
}

TestRunner.Register("Thumbnail title updates retain original window identity", ThumbnailNamesUpdatesTest)
ThumbnailNamesUpdatesTest() {
    app := ThumbnailNamesFixture()
    app.CustomThumbnailNames := Map("Characters", "Alice", "Names", "Scout")
    overlay := {Text: ""}
    window := {Title: "EVE"}
    app.ThumbWindows.test := Map("TextOverlay", Map("OverlayText", overlay, "EventText", {Text: ""}, "SystemText", {Text: ""}), "Window", window)
    app.SetThumbnailText["test"] := "EVE - Alice"
    AssertEqual("Scout", overlay.Text)
    AssertEqual("EVE - Alice", window.Title)
    app.updateThumbnailText("EVE - Alice`nUnder attack", "test")
    AssertEqual("Scout`nUnder attack", overlay.Text)
    app.updateThumbnailText("EVE - Alice", "test")
    AssertEqual("Scout", overlay.Text)
    AssertEqual("EVE - Alice", window.Title)
}
