#Requires AutoHotkey v2.0
#NoTrayIcon
#Warn All, StdOut
#Include "../Lib/json.ahk"
#Include "../src/SettingsWindowSizing.ahk"
#Include "../src/ProgramLog.ahk"
#Include "../src/ReleaseInstaller.ahk"
if EnvGet("USERNAME") != "WDAGUtilityAccount"
    ExitApp(2)
try {
    directory := A_Args[1], destination := A_Args[2]
    if !ReleaseInstaller.RecoveryUpdater(directory "\EVE-X-Preview-Updater.exe")
        throw Error("The compiled recovery updater protocol marker was not recognized.")
    if ReleaseInstaller.RecoveryUpdater(A_AhkPath)
        throw Error("An unrelated executable was accepted as a package updater.")
    ; A deliberately unavailable release must not be contacted when the current
    ; installation contains a compatible updater, including when downgrading.
    ReleaseInstaller.StageUpdater(directory, destination, "99.99.99.99")
    if !ReleaseInstaller.RecoveryUpdater(destination)
        throw Error("The installed updater was not staged for the handoff.")
    FileAppend("PASS: compatible installed updater used independently of the target release`n", A_Args[3])
    missing := A_Temp "\evex-missing-updater-" SettingsBridgeId()
    DirCreate(missing)
    try {
        ProtocolDownloadFixture.source := directory "\EVE-X-Preview-Updater.exe"
        ProtocolDownloadFixture.StageUpdater(missing, missing "\download.exe", "1.7.0.5")
        if ProtocolDownloadFixture.url != "https://github.com/khivus/EVE-X-Preview/releases/latest/download/EVE-X-Preview-Updater.exe"
            throw Error("Updater bootstrap still depends on the installed build's removed release.")
        if !ReleaseInstaller.RecoveryUpdater(missing "\download.exe")
            throw Error("Latest stable recovery updater was not staged.")
        FileAppend("PASS: missing updater bootstraps from stable independently of a removed installed release`n", A_Args[3])
        ProtocolDownloadFixture.source := A_AhkPath
        rejected := false
        try ProtocolDownloadFixture.StageUpdater(missing, missing "\invalid.exe", "1.7.0.5")
        catch as err
            rejected := InStr(err.Message, "compatible EVE-X-Preview-Updater.exe")
        if !rejected
            throw Error("Downloaded executable without recovery protocol was accepted.")
        FileAppend("PASS: incompatible downloaded updater rejected before execution`n", A_Args[3])
    } finally DirDelete(missing, true)
    ExitApp(0)
} catch as err {
    FileAppend("FAIL: " err.Message "`n", A_Args[3])
    ExitApp(1)
}

SettingsBridgeId() => ProcessExist() "-" A_TickCount
class ProtocolDownloadFixture extends ReleaseInstaller {
    static source := ""
    static url := ""
    static DownloadUpdater(url, destination) {
        this.url := url
        FileCopy(this.source, destination)
    }
}
