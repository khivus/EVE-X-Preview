#Requires AutoHotkey v2.0
#SingleInstance Off
if EnvGet("USERNAME") != "WDAGUtilityAccount"
    ExitApp(2)
global clientId := Integer(A_Args[1])
global logPath := A_Args[2] "\client-" clientId ".csv"
global panel := Gui(, "EVE - Rapid Pilot " Format("{:02}", clientId))
panel.BackColor := "182838"
panel.SetFont("s18 cFFFFFF", "Segoe UI")
panel.AddText("w650 h400", "Input receiver " clientId)
panel.Show("NA x150 y100 w850 h600")
OnMessage(0x100, KeyEvent)
OnMessage(0x101, KeyEvent)
OnMessage(0x104, KeyEvent)
OnMessage(0x105, KeyEvent)
OnMessage(0x6, Activation)
KeyEvent(w, l, msg, hwnd) {
    if w != 0x46 && w != 0x10 && w != 0x7C
        return
    LogEvent(msg ":" w)
}
Activation(w, l, msg, hwnd) => LogEvent("activate:" (w & 65535))
LogEvent(event) {
    shifted := !!(DllCall("GetKeyState", "Int", 0x10, "Short") & 0x8000)
    asyncShift := !!(DllCall("GetAsyncKeyState", "Int", 0x10, "Short") & 0x8000)
    extra := DllCall("GetMessageExtraInfo", "UPtr")
    FileAppend(DllCall("GetTickCount64", "Int64") "," clientId "," event "," shifted "," asyncShift "," extra "`n", logPath)
}
Persistent
