#Requires AutoHotkey v2.0
#NoTrayIcon
#Warn All, StdOut
if EnvGet("USERNAME") != "WDAGUtilityAccount"
    ExitApp(2)
pid := 0
process := 0
result := 1
try {
    Run('"' A_Args[1] '"',, "Hide", &pid)
    ; Keep the launched process handle for both the exit assertion and cleanup.
    process := DllCall("OpenProcess", "UInt", 0x100000 | 0x1000 | 0x1, "Int", false, "UInt", pid, "Ptr")
    if !process
        throw OSError()
    deadline := A_TickCount + 10000
    window := WinWait("EVE-X-Preview update ahk_pid " pid,, 10)
    if !window
        throw Error("The application did not show its incomplete-update notice.")
    button := 0
    loop {
        try button := ControlGetHwnd("Button1", "ahk_id " window)
        if button && DllCall("IsWindowVisible", "Ptr", button) && DllCall("IsWindowEnabled", "Ptr", button)
            break
        if A_TickCount >= deadline
            throw Error("The incomplete-update notice was not ready for acknowledgement.")
        Sleep(10)
    }
    ; Send the button's native BN_CLICKED command to this process's notice.
    ; Mouse messages can be lost during dialog activation or initialization.
    buttonId := DllCall("GetDlgCtrlID", "Ptr", button, "Int")
    if buttonId <= 0
        throw Error("The incomplete-update notice has no acknowledgement command.")
    PostMessage(0x111, buttonId, button,, "ahk_id " window)
    if DllCall("WaitForSingleObject", "Ptr", process, "UInt", 5000, "UInt") != 0
        throw Error("The blocked application did not exit.")
    code := 0
    if !DllCall("GetExitCodeProcess", "Ptr", process, "UInt*", &code) || code != 1
        throw Error("The blocked application returned an unexpected exit code.")
    result := 0
} catch as err {
    FileAppend("FAIL: " err.Message "`n", A_Args[2])
} finally {
    if process {
        ; A failed assertion must not leave a modal child behind: PowerShell's
        ; Start-Process -Wait otherwise keeps waiting for that descendant.
        if DllCall("WaitForSingleObject", "Ptr", process, "UInt", 0, "UInt") = 258 {
            if !DllCall("TerminateProcess", "Ptr", process, "UInt", 2)
                || DllCall("WaitForSingleObject", "Ptr", process, "UInt", 5000, "UInt") != 0 {
                FileAppend("FAIL: Could not clean up the launched startup-guard fixture.`n", A_Args[2])
                result := 1
            }
        }
        DllCall("CloseHandle", "Ptr", process)
    } else if pid && ProcessExist(pid) {
        ProcessClose(pid)
    }
}
if result = 0
    FileAppend("PASS: compiled main refused an unfinished package transaction`n", A_Args[2])
ExitApp(result)
