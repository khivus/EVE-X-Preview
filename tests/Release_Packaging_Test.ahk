TestRunner.Register("Release packaging updater version and application suspend resources match the main package", ReleaseUpdaterMetadataTest)
ReleaseUpdaterMetadataTest() {
    source := FileRead(A_ScriptDir "\..\Updater.ahk", "UTF-8")
    versionSource := FileRead(A_ScriptDir "\..\Version.ahk", "UTF-8")
    AssertTrue(RegExMatch(versionSource, 'ReleaseVersion := "([\d.]+)"', &release))
    AssertEqual(release[1], UpdateChecker.CurrentVersion())
    AssertTrue(InStr(source, 'VERSION := ReleaseVersion'))
    for entry in ["Main.ahk", "Updater.ahk"]
        AssertTrue(InStr(FileRead(A_ScriptDir "\..\" entry), '#Include "Version.ahk"'))
    for directive in ["SetVersion %U_version%", "SetFileVersion %U_version%"]
        AssertTrue(InStr(versionSource, "Ahk2Exe-" directive))
    for directive in ["SetMainIcon res/Icon.ico",
        "AddResource res/Icon.ico, 160", "AddResource res/Icon-Suspend.ico, 206",
        "AddResource res/Icon.ico, 207", "AddResource res/Icon-Suspend.ico, 208"]
        AssertTrue(InStr(source, "Ahk2Exe-" directive), "Updater package needs " directive)
    AssertTrue(InStr(source, 'ProgramLog.Init(EnvGet("LOCALAPPDATA") "\EVE-X-Preview\Logs\Updater")'))
}
