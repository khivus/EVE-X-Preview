#Requires AutoHotkey v2.0
; Standalone receiver regression. Run only through Test-RapidSwitchGuest.ps1.
global root := RegExReplace(A_ScriptDir, "\\tests$")
if EnvGet("USERNAME") != "WDAGUtilityAccount" || InStr(root, "C:\Checks\") != 1
    ExitApp(2)
global output := A_Args.Length ? A_Args[1] : "", pid := 0
if InStr(output, "C:\Output\") != 1 || FileExist(output "\driver.csv")
    ExitApp(2)
DirCreate(output)
OnExit(Cleanup)
OnError((err, *) => (FileAppend(err.Message "`n", output "\error.txt"), ExitApp(1)))

Key(vk, up := false, tag := 0) {
    packet := Buffer(40, 0)
    NumPut("UInt", 1, packet, 0)
    NumPut("UShort", vk, packet, 8)
    NumPut("UInt", up ? 2 : 0, packet, 12)
    NumPut("UPtr", tag, packet, 24)
    if DllCall("SendInput", "UInt", 1, "Ptr", packet, "Int", 40) != 1
        throw Error("SendInput failed")
}
Cleanup(*) {
    Key(0x46, true), Key(0xA0, true)
    if pid && ProcessExist(pid)
        ProcessClose(pid)
}

; No EVE-X-Preview process: this isolates the observer from window switching.
Run('"' root '\exefile.exe" 1 "' output '" --slow-log',,, &pid)
hwnd := WinWait("EVE - Rapid Pilot 01",, 10)
if !hwnd
    throw Error("Receiver did not start")
WinActivate(hwnd)
if !WinWaitActive(hwnd,, 3)
    throw Error("Receiver focus failed")
Sleep(300)
FileAppend("case,trial,token,expected`n", output "\driver.csv")
loop 32 {
    if !WinActive(hwnd)
        throw Error("Receiver lost focus; foreground=" WinGetTitle("A"))
    token := 0x4D000000 + A_Index
    ; Back-to-back Shift/F key-downs exposed the old OnMessage observer's
    ; reentrancy limit when its Shift logging callback was still running.
    Key(0xA0), Key(0x46, false, token)
    DllCall("Sleep", "UInt", 10)
    Key(0x46, true, token), Key(0xA0, true)
    FileAppend("receiver-slow-log," A_Index "," token ",1`n", output "\driver.csv")
    Sleep(150)
}
Sleep(1000) ; Allow queued events to drain through the deliberately slow logger.
ExitApp(0)
