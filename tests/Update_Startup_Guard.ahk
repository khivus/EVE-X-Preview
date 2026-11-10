#Requires AutoHotkey v2.0
#NoTrayIcon
#Warn All, StdOut
if EnvGet("USERNAME") != "WDAGUtilityAccount"
    ExitApp(2)
try {
    Run('"' A_Args[1] '"',, "Hide", &pid)
    process := DllCall("OpenProcess", "UInt", 0x100000 | 0x1000, "Int", false, "UInt", pid, "Ptr")
    try {
        window := WinWait("EVE-X-Preview update ahk_pid " pid,, 10)
        if !window
            throw Error("The application did not show its incomplete-update notice.")
        ControlClick("Button1", "ahk_id " window)
        if ProcessWaitClose(pid, 5)
            throw Error("The blocked application did not exit.")
        code := 0
        if !DllCall("GetExitCodeProcess", "Ptr", process, "UInt*", &code) || code != 1
            throw Error("The blocked application returned an unexpected exit code.")
    } finally DllCall("CloseHandle", "Ptr", process)
    FileAppend("PASS: compiled main refused an unfinished package transaction`n", A_Args[2])
    ExitApp(0)
} catch as err {
    FileAppend("FAIL: " err.Message "`n", A_Args[2])
    ExitApp(1)
}
