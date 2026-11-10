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
    ExitApp(0)
} catch as err {
    FileAppend("FAIL: " err.Message "`n", A_Args[3])
    ExitApp(1)
}
