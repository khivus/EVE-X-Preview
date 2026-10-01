class ProfileSwitchFixture extends Main_Class {
    __New() {
        This._JSON := Map("LastUsedProfile", "First", "_Profiles", Map("First", Map(), "Second", Map()))
        This.layoutSwitches := 0
    }

    TrySwitchKeyboardToEnglish() {
        This.layoutSwitches++
        return true
    }
}

TestRunner.Register("Profile save replaces settings with complete selected profile", ProfileSaveTest)
ProfileSaveTest() {
    dir := A_Temp "\EVE-X-Preview-profile-test-" A_TickCount
    path := dir "\settings.json"
    DirCreate(dir)
    try {
        fixture := ProfileSwitchFixture()
        fixture.SaveJsonToFile(path)
        AssertEqual("First", JSON.Load(FileRead(path))["LastUsedProfile"])
        fixture.LastUsedProfile := "Second"
        fixture.SaveJsonToFile(path)
        AssertEqual("Second", JSON.Load(FileRead(path))["LastUsedProfile"])
        AssertFalse(FileExist(path ".save.tmp"), "Completed save must leave no temporary settings file.")
    } finally {
        if FileExist(path)
            FileDelete(path)
        if FileExist(path ".save.tmp")
            FileDelete(path ".save.tmp")
        if FileExist(dir)
            DirDelete(dir)
    }
}

TestRunner.Register("Invalid layout-sensitive hotkey does not abort profile startup", ProfileHotkeyRetryTest)
ProfileHotkeyRetryTest() {
    fixture := ProfileSwitchFixture()
    result := fixture.RegisterHotkeyWithLanguageRetry("!NotARealKeyName", (*) => 0, "P1", "Test hotkey")
    AssertFalse(result, "An invalid hotkey must be skipped after one retry.")
    AssertEqual(1, fixture.layoutSwitches)
}
