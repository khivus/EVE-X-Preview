#Requires AutoHotkey v2.0
#NoTrayIcon
if EnvGet("USERNAME") != "WDAGUtilityAccount"
    ExitApp(2)
mode := EnvGet("EVEX_GUARD_FIXTURE_MODE")
FileAppend(ProcessExist(), EnvGet("EVEX_GUARD_FIXTURE_PID"))
if mode = "no-notice" {
    Persistent()
} else {
    notice := Gui(, "EVE-X-Preview update")
    notice.AddText(, "Controlled startup notice for the compiled guard test.")
    button := notice.AddButton("Disabled Default", "OK")
    button.OnEvent("Click", Acknowledge)
    notice.Show()
    if mode != "never-ready"
        SetTimer(() => button.Enabled := true, -1500)
}

Acknowledge(*) {
    global mode, notice, button
    if !button.Enabled
        return
    if mode = "hang-after-ack" {
        notice.Hide()
        Persistent()
    } else
        ExitApp(mode = "wrong-exit" ? 0 : 1)
}
