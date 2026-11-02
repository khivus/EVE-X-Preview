TestRunner.Register("Release packaging updater version and application suspend resources match the main package", ReleaseUpdaterMetadataTest)
ReleaseUpdaterMetadataTest() {
    source := FileRead(A_ScriptDir "\..\Updater.ahk", "UTF-8")
    AssertTrue(RegExMatch(source, 'VERSION := "([\d.]+)"', &displayed))
    AssertTrue(RegExMatch(source, "Ahk2Exe-Let U_version = ([\d.]+)", &compiled))
    AssertEqual(displayed[1], compiled[1])
    for directive in ["SetVersion %U_version%", "SetFileVersion %U_version%", "SetMainIcon res/Icon.ico",
        "AddResource res/Icon.ico, 160", "AddResource res/Icon-Suspend.ico, 206",
        "AddResource res/Icon.ico, 207", "AddResource res/Icon-Suspend.ico, 208"]
        AssertTrue(InStr(source, "Ahk2Exe-" directive), "Updater package needs " directive)
    AssertTrue(InStr(source, 'ProgramLog.Init(EnvGet("LOCALAPPDATA") "\EVE-X-Preview\Logs\Updater")'))
}
