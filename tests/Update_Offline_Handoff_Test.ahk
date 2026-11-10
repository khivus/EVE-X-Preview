#Requires AutoHotkey v2.0
#NoTrayIcon
#Warn All, StdOut
#Include "../SettingsHost.ahk"
if EnvGet("USERNAME") != "WDAGUtilityAccount"
    ExitApp(2)
; Exercise the actual host lifetime check with a signalled handle, just as when
; closing the modern editor ends the offline host's parent process.
output := hostOwner.installation "\handoff-result.txt"
FileAppend("Reached pending updater handoff.`n", output)
hostOwner.settingsUpdateHandoff := true
DllCall("CloseHandle", "Ptr", hostOwner.parentHandle)
hostOwner.parentHandle := DllCall("CreateEventW", "Ptr", 0, "Int", true, "Int", true, "Ptr", 0, "Ptr")
hostOwner.CheckParent()
FileAppend("PASS: offline host survives editor exit until updater handoff finishes.`n", output)
ExitApp()
