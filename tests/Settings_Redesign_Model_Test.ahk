TestRunner.Register("Settings draft isolates edits and merges unrelated runtime state", SettingsDraftMergeTest)
SettingsDraftMergeTest() {
    config := JSON.Load(default_JSON)
    model := SettingsModel(config)
    AssertFalse(model.Dirty)
    model.Set("ThumbnailStartLocationwidth", "320")
    AssertEqual(250, config["_Profiles"]["Default"]["Thumbnails Behavior"]["ThumbnailStartLocation"]["width"])
    config["_Profiles"]["Default"]["Thumbnail Positions"]["EVE - New Pilot"] := [50, 60, 70, 80]
    config["_Profiles"]["Default"]["Game Logs Monitoring"]["charsIds"]["New Pilot"] := 42
    merged := model.Merge(config)
    AssertEqual(320, merged["_Profiles"]["Default"]["Thumbnails Behavior"]["ThumbnailStartLocation"]["width"])
    AssertDeepEqual([50, 60, 70, 80], merged["_Profiles"]["Default"]["Thumbnail Positions"]["EVE - New Pilot"])
    AssertEqual(42, merged["_Profiles"]["Default"]["Game Logs Monitoring"]["charsIds"]["New Pilot"])
    AssertTrue(model.Dirty)
    config["_Profiles"]["Fleet"] := SettingsModel.Clone(config["_Profiles"]["Default"])
    config["LastUsedProfile"] := "Fleet"
    renamed := SettingsModel(config)
    renamed.RenameProfile("Logistics")
    renamed.Set("ThumbnailTextSize", 18)
    config["_Profiles"]["Fleet"]["Thumbnail Positions"]["EVE - Late Pilot"] := [1, 2, 3, 4]
    config["_Profiles"]["Fleet"]["Game Logs Monitoring"]["charsIds"]["Late Pilot"] := 99
    merged := renamed.Merge(config)
    AssertFalse(merged["_Profiles"].Has("Fleet"))
    AssertEqual("Logistics", merged["LastUsedProfile"])
    AssertDeepEqual([1, 2, 3, 4], merged["_Profiles"]["Logistics"]["Thumbnail Positions"]["EVE - Late Pilot"])
    AssertEqual(99, merged["_Profiles"]["Logistics"]["Game Logs Monitoring"]["charsIds"]["Late Pilot"])
    AssertEqual(18, merged["_Profiles"]["Logistics"]["Thumbnails Visuals"]["ThumbnailTextSize"])
}

TestRunner.Register("Settings validation rejects instead of silently clamping and preserves accepted values", SettingsDraftValidationTest)
SettingsDraftValidationTest() {
    model := SettingsModel(JSON.Load(default_JSON))
    AssertThrows(() => model.Set("ThumbnailPollingInterval", 5000), "50 to 1000")
    AssertEqual(50, model.Get("ThumbnailPollingInterval"))
    AssertThrows(() => model.Set("Minimizeclients_Delay", 0), "50 to 60000")
    AssertThrows(() => model.Set("ThumbnailStartLocationwidth", "12px"), "whole number")
    AssertEqual(-1200, model.Set("ThumbnailStartLocationx", "-1200"))
    AssertEqual("FF0010", model.Set("ThumbnailTextColor", "rgb(255, 0, 16)"))
    AssertThrows(() => model.Set("ThumbnailTextColor", "999, 0, 0"), "0 and 255")
    AssertThrows(() => model.Set("ThumbnailTextColor", "not a color"), "RGB")
    AssertEqual(12.5, model.Set("incomingDPSThreshold", "12.5"))
    AssertThrows(() => model.Set("dpsAverageSeconds", "1.5"), "whole number")
    AssertThrows(() => model.Set("ProfileActivationHotkey", "NoSuchKey123"), "Unknown key")
    AssertEqual("^!F23", model.Set("ProfileActivationHotkey", "^!F23"))
    AssertTrue(InStr(SettingsValue.HotkeyText("^!F23"), "F23"))
}

TestRunner.Register("Explicit inheritance preserves legacy effective values and changes only one profile", SettingsDraftInheritanceTest)
SettingsDraftInheritanceTest() {
    config := JSON.Load(default_JSON)
    config["_Profiles"]["Second"] := SettingsModel.Clone(config["_Profiles"]["Default"])
    config["_Profiles"]["Third"] := SettingsModel.Clone(config["_Profiles"]["Default"])
    config["_Profiles"]["Default"]["Other"]["Global_Groups"]["Thumbnails Behavior"] := 1
    config["_Profiles"]["Default"]["Other"]["Global_Groups"]["Thumbnails Visuals"] := 1
    config["_Profiles"]["Second"]["Other"]["Global_Groups"]["Thumbnails Visuals"] := 0
    config["_Profiles"]["Second"]["Other"]["Global_Groups"]["Thumbnails Behavior"] := 0
    config["_Profiles"]["Third"]["Other"]["Global_Groups"]["Thumbnails Behavior"] := 1
    model := SettingsModel(config)
    model.SelectProfile("Second")
    AssertEqual("Default", model.Source("layout"))
    AssertThrows(() => model.Set("ThumbnailStartLocationwidth", 300), "Customize")
    model.SetInheritance("layout", false)
    AssertEqual("Second", model.Source("layout"))
    AssertEqual("Default", model.Source("placement"))
    AssertEqual(300, model.Set("ThumbnailStartLocationwidth", 300))
    model.SelectProfile("Third")
    AssertEqual("Third", model.Source("Thumbnails Behavior"))
    AssertEqual(250, model.Get("ThumbnailStartLocationwidth"))
    AssertEqual("Default", model.config["LastUsedProfile"], "Editing a profile must not activate it.")
}

TestRunner.Register("Settings migration preserves profile-local shortcuts groups positions and effective focus policies", SettingsLegacySourcesTest)
SettingsLegacySourcesTest() {
    config := JSON.Load(default_JSON)
    primary := config["_Profiles"]["Default"]
    config["_Profiles"]["Fleet"] := SettingsModel.Clone(primary)
    localProfile := config["_Profiles"]["Fleet"]
    for group in ["Hotkeys Settings", "Hotkey Groups", "Game Logs Monitoring"] {
        primary["Other"]["Global_Groups"][group] := 1
        localProfile["Other"]["Global_Groups"][group] := 0
    }
    primary["Game Logs Monitoring"]["supressForFocused"] := 1
    localProfile["Game Logs Monitoring"]["supressForFocused"] := 0
    primary["Hotkeys Settings"]["CharacterHotkeys"] := [Map("EVE - Default Pilot", "^F1")]
    localProfile["Hotkeys Settings"]["CharacterHotkeys"] := [Map("EVE - Fleet Pilot", "!F2")]
    primary["Hotkey Groups"] := Map("Default group", Map("Characters", ["EVE - Default Pilot"], "ForwardsHotkey", "", "BackwardsHotkey", "", "FirstCharHotkey", ""))
    localProfile["Hotkey Groups"] := Map("Fleet group", Map("Characters", ["EVE - Fleet Pilot"], "ForwardsHotkey", "", "BackwardsHotkey", "", "FirstCharHotkey", ""))
    config["LastUsedProfile"] := "Fleet"
    model := SettingsModel(config)
    AssertEqual("Default", model.Source("Hotkeys Settings"))
    AssertDeepEqual([["Fleet Pilot", "!F2"]], model.Rows("character-shortcuts"))
    AssertTrue(model.Groups("characters").Has("Fleet group"))
    AssertEqual(1, model.Get("suppressFocusedEventSounds"))
    AssertEqual(1, model.Get("hideDPSForFocused"))
    runtime := CombatEventsFixture()
    runtime._JSON := model.config
    runtime.ProfileOverride()
    AssertEqual("!F2", runtime._Hotkeys["EVE - Fleet Pilot"])
    AssertTrue(runtime.Hotkey_Groups.Has("Fleet group"))
    AssertTrue(runtime.SuppressFocusedEventSounds)
    AssertTrue(runtime.HideDPSForFocused)
    model.SetInheritance("Character shortcuts", true)
    model.SetInheritance("Character groups", true)
    runtime.ProfileOverride()
    AssertDeepEqual([["Default Pilot", "^F1"]], model.Rows("character-shortcuts"))
    AssertEqual("^F1", runtime._Hotkeys["EVE - Default Pilot"])
    AssertTrue(runtime.Hotkey_Groups.Has("Default group"))
    model.Set("suppressFocusedEventSounds", 0)
    model.Set("hideDPSForFocused", 0)
    AssertFalse(runtime.SuppressFocusedEventSounds)
    AssertFalse(runtime.HideDPSForFocused)
    AssertEqual(1, primary["Game Logs Monitoring"]["supressForFocused"], "Migration must not modify the input configuration.")
}

TestRunner.Register("Profile and group names are protected and removed data can be restored", SettingsDraftNamesTest)
SettingsDraftNamesTest() {
    model := SettingsModel(JSON.Load(default_JSON))
    AssertThrows(() => model.AddProfile("   "), "Enter a name")
    model.AddProfile("Fleet")
    AssertThrows(() => model.AddProfile("fleet"), "already exists")
    AssertEqual("", model.Get("ProfileActivationHotkey"))
    model.AddGroup("characters", "Logistics")
    AssertThrows(() => model.AddGroup("characters", "logistics"), "already exists")
    model.DeleteGroup("characters", "Logistics")
    AssertFalse(model.Groups("characters").Has("Logistics"))
    model.UndoLast()
    AssertTrue(model.Groups("characters").Has("Logistics"))
    model.DeleteProfile()
    AssertFalse(model.Profiles.Has("Fleet"))
    model.UndoLast()
    AssertTrue(model.Profiles.Has("Fleet"))
    AssertTrue(model.Groups("characters").Has("Logistics"))
}

TestRunner.Register("Row editors keep names colors titles and shortcuts paired through reorder and deletion", SettingsRowsRoundTripTest)
SettingsRowsRoundTripTest() {
    model := SettingsModel(JSON.Load(default_JSON))
    for pair in [
        ["character-shortcuts", [["Alice", "^F23"], ["Bob", "!F24"]]],
        ["character-colors", [["Alice", "FF0000", "0000FF", "FFFFFF"], ["Bob", "00FF00", "FFFF00", "123456"]]],
        ["custom-names", [["Alice", "Scout"], ["Bob", ""]]],
        ["app-shortcuts", [["alpha.exe", "First window", "^F23"], ["beta.exe", "", "!F24"]]],
        ["exceptions", [["Alice", 1, 0], ["Bob", 0, 1]]]
    ] {
        kind := pair[1], rows := pair[2]
        model.SaveRows(kind, rows)
        AssertDeepEqual(rows, model.Rows(kind))
        rows := [rows[2], rows[1]]
        model.SaveRows(kind, rows)
        if kind = "exceptions" {
            matched := Map()
            for row in model.Rows(kind)
                matched[row[1]] := [row[2], row[3]]
            AssertDeepEqual(Map("Alice", [1, 0], "Bob", [0, 1]), matched)
        } else
            AssertDeepEqual(rows, model.Rows(kind))
        rows.RemoveAt(1)
        model.SaveRows(kind, rows)
        AssertDeepEqual(rows, model.Rows(kind))
    }
    model.AddGroup("apps", "Tools")
    model.SaveRows("app-group-members", [["alpha.exe", ""], ["beta.exe", "Beta"]], "Tools")
    AssertDeepEqual([["alpha.exe", ""], ["beta.exe", "Beta"]], model.Rows("app-group-members", "Tools"))
    AssertThrows(() => model.SaveRows("app-shortcuts", [["alpha.exe", "", ""], ["ALPHA.exe", "", ""]]), "duplicates")
}

TestRunner.Register("Mouse gestures reject duplicates and search finds reorganized controls", SettingsGestureSearchTest)
SettingsGestureSearchTest() {
    model := SettingsModel(JSON.Load(default_JSON))
    rows := model.Rows("gestures")
    rows[2][2] := rows[1][2]
    AssertThrows(() => model.SaveRows("gestures", rows), "same gesture")
    result := SettingsSchema.Search("opacity")
    AssertEqual(1, result.Length)
    AssertEqual("appearance", result[1].id)
    AssertEqual(0, SettingsSchema.Search("does-not-exist-42").Length)
    AssertEqual(7, SettingsSchema.Pages.Length)
}
