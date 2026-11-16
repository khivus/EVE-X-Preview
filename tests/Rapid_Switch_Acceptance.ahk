#Requires AutoHotkey v2.0
; Standalone desktop acceptance, deliberately excluded from TestRunner.
global __EVE_X_PREVIEW_TESTING__ := true
#Include "..\Main.ahk"

global root := RegExReplace(A_ScriptDir, "\\tests$")
if EnvGet("USERNAME") != "WDAGUtilityAccount" || InStr(root, "C:\Checks\") != 1
    ExitApp(2)
global output := A_Args.Length ? A_Args[1] : ""
if InStr(output, "C:\Output\") != 1 || FileExist(output "\driver.csv")
    ExitApp(2)
global clientCount := A_Args.Length >= 2 ? Integer(A_Args[2]) : 16
if clientCount < 12 || clientCount > 32
    ExitApp(2)
global clients := [], titles := [], appPid := 0, sequence := 0, cleaned := false
DirCreate(output)
SetTitleMatchMode(3)
SetWinDelay(-1)
OnError(Failed)
OnExit(Cleanup)
DllCall("winmm\timeBeginPeriod", "UInt", 1)

Failed(err, *) {
    FileAppend(err.Message " at " err.Line "`n", output "\error.txt")
    ExitApp(1)
}
Cleanup(*) {
    global cleaned
    if cleaned
        return
    cleaned := true
    for vk in [0x46, 0x7C, 0xA0]
        Key(vk, true)
    if appPid && ProcessExist(appPid)
        ProcessClose(appPid)
    for client in clients
        if ProcessExist(client.pid)
            ProcessClose(client.pid)
    DllCall("winmm\timeEndPeriod", "UInt", 1)
}
Now() => DllCall("GetTickCount64", "Int64")
Pause(ms) {
    if ms > 0
        DllCall("Sleep", "UInt", ms)
}
Key(vk, up := false, tag := 0) {
    packet := Buffer(40, 0) ; x64 INPUT / KEYBDINPUT
    NumPut("UInt", 1, packet, 0)
    NumPut("UShort", vk, packet, 8)
    NumPut("UInt", up ? 2 : 0, packet, 12)
    NumPut("UPtr", tag, packet, 24)
    if DllCall("SendInput", "UInt", 1, "Ptr", packet, "Int", 40) != 1
        throw Error("SendInput failed")
}
ActiveIndex() {
    active := DllCall("GetForegroundWindow", "Ptr")
    for index, client in clients
        if client.hwnd = active
            return index
    return 0
}
StartApp(keep) {
    global appPid
    Key(0xA0, true)
    if appPid && ProcessExist(appPid)
        ProcessClose(appPid)
    config := JSON.Load(default_JSON)
    config["SettingsEditor"] := "builtin"
    config["SettingsDownloadSuggested"] := 1
    profile := config["_Profiles"]["Default"]
    hotkeys := profile["Hotkeys Settings"]
    hotkeys["GroupsHoldDelay"] := 30
    hotkeys["KeepGroupsPositions"] := keep
    hotkeys["PreserveHotkeysOnLogout"] := 1
    hotkeys["MaxActiveWindowRetries"] := 3
    hotkeys["ActiveWindowRetryInterval"] := 25
    profile["Thumbnails Behavior"]["ThumbnailPollingInterval"] := 50
    for name in ["SlowThumbnailCreation", "AutoPauseOnDwmFailure", "SwitchLangOnErr"]
        profile["Other"][name] := 0
    profile["Hotkey Groups"] := Map("Rapid", Map("Characters", titles,
        "ForwardsHotkey", "*F13", "BackwardsHotkey", "", "FirstCharHotkey", ""))
    stream := FileOpen(root "\EVE-X-Preview.json", "w", "UTF-8")
    stream.Write(JSON.Dump(config)), stream.Close()
    FileCopy(root "\EVE-X-Preview.json", output "\settings-keep" keep ".json", true)
    Run('"' root '\EVE-X-Preview.exe"', root,, &appPid)
    Sleep(5000)
    ; Prime the remembered cursor at member 1 without assuming startup focus.
    WinActivate(clients[clientCount].hwnd)
    WinWaitActive(clients[clientCount].hwnd,, 3)
    Key(0x7C), Pause(1), Key(0x7C, true)
    if !WinWaitActive(clients[1].hwnd,, 3)
        throw Error("Initial group switch did not reach member 1")
    Sleep(200)
}
ManualChecks(keep) {
    StartApp(keep)
    CoordMode("Mouse", "Screen")
    Key(0xA0)
    loop clientCount * 2 {
        chosen := Mod(A_Index * 5 + 2, clientCount) + 1
        hwnd := clients[chosen].hwnd
        WinSetAlwaysOnTop(1, hwnd)
        Sleep(50) ; Let the topmost z-order update before the actual mouse input.
        WinGetPos(&x, &y,,, hwnd)
        ; Click the title bar, away from the app's preview grid and controls.
        DllCall("SetCursorPos", "Int", x + 200, "Int", y + 15)
        DllCall("mouse_event", "UInt", 2, "UInt", 0, "UInt", 0, "UInt", 0, "UPtr", 0)
        DllCall("mouse_event", "UInt", 4, "UInt", 0, "UInt", 0, "UInt", 0, "UPtr", 0)
        if !WinWaitActive(hwnd,, 2)
            throw Error("Manual click did not focus the selected client")
        expected := Mod(keep ? A_Index : chosen, clientCount) + 1
        Key(0x7C), Pause(1), Key(0x7C, true)
        Pause(200)
        actual := ActiveIndex()
        FileAppend(keep "," chosen "," expected "," actual "`n", output "\manual.csv")
        WinSetAlwaysOnTop(0, hwnd)
        if actual != expected
            throw Error("Manual selection violated the configured group position behavior")
        Pause(50)
    }
    Key(0xA0, true)
}

try {
    for vk in [0x5B, 0x5C, 0xA2, 0xA3, 0xA4, 0xA5, 0xA0, 0xA1]
        Key(vk, true)
    loop clientCount {
        title := "EVE - Rapid Pilot " Format("{:02}", A_Index)
        Run('"' root '\exefile.exe" /ErrorStdOut "' A_ScriptDir '\fixtures\RapidSwitchClient.ahk" ' A_Index ' "' output '"',,, &pid)
        clients.Push({hwnd: 0, pid: pid})
        if !WinWait(title,, 10)
            throw Error("Synthetic client did not start")
        clients[A_Index].hwnd := WinExist(title)
        titles.Push(title)
        FileAppend(A_Index "," WinExist(title) "`n", output "\windows.csv")
    }
    FileAppend("case,trial,start,keydown,finish,before,expected,atF,after,hold,gap,period,heldShift,token`n", output "\driver.csv")
    FileAppend("keep,clicked,expected,actual`n", output "\manual.csv")
    ; The 20 ms sequence intentionally exceeds the configured 30 ms hold
    ; threshold. Record its behavior, but do not promise lossless input there.
    scenarios := [
        {name: "held-short", gap: 10, period: 100, duration: 10, held: true},
        {name: "held-long", gap: 10, period: 150, duration: 40, held: true},
        {name: "held-30ms", gap: 30, period: 100, duration: 10, held: true},
        {name: "released", gap: 10, period: 100, duration: 10, held: false},
        {name: "stress-zero", gap: 0, period: 100, duration: 10, held: true},
        {name: "stress-below-hold", gap: 0, period: 20, duration: 10, held: true}]
    for keep in [1, 0] {
        for scenario in scenarios {
            if !keep && scenario.name != "held-short"
                continue
            StartApp(keep)
            label := "keep" keep "-" scenario.name
            FileAppend(label "`n", output "\progress.txt")
            if scenario.held
                Key(0xA0)
            loop clientCount * 4 {
                start := Now(), before := ActiveIndex()
                expected := Mod(A_Index, clientCount) + 1
                Key(0x7C), Pause(1), Key(0x7C, true)
                Pause(scenario.gap)
                if !scenario.held
                    Key(0xA0)
                atF := ActiveIndex(), down := Now(), token := 0x4D000000 + (++sequence)
                Key(0x46, false, token), Pause(scenario.duration), Key(0x46, true, token)
                if !scenario.held
                    Key(0xA0, true)
                Pause(Max(0, scenario.period - (Now() - start)))
                FileAppend(label "," A_Index "," start "," down "," Now() "," before "," expected "," atF "," ActiveIndex()
                    ",30," scenario.gap "," scenario.period "," scenario.held "," token "`n", output "\driver.csv")
            }
            Key(0xA0, true), Sleep(300)
        }
        ManualChecks(keep)
    }
    FileAppend("completed`n", output "\progress.txt")
} finally Cleanup()
ExitApp(0)
