; Command-hotkey receipts have their own tooltip slot and one resettable timer.
; Window switching and thumbnail mouse gestures never use this dispatcher.
class HotkeyFeedback {
    static Show(message) {
        if !this.HasOwnProp("clearTimer")
            this.clearTimer := (*) => ToolTip(,,,18)
        SetTimer(this.clearTimer, 0)
        ToolTip(message,,,18)
        SetTimer(this.clearTimer, -1500)
    }
    static Hide() {
        if this.HasOwnProp("clearTimer")
            SetTimer(this.clearTimer, 0)
        ToolTip(,,,18)
    }
    static Run(app, id, *) {
        try {
            switch id {
                case "Suspend_Hotkeys_Hotkey":
                    app.Suspend_Hotkeys()
                    message := app.activationHotkeysSuspended ? "EVE switching hotkeys paused" : "EVE switching hotkeys resumed"
                case "Suspend_All_Hotkeys_Hotkey":
                    app.Suspend_All_Hotkeys()
                    message := A_IsSuspended ? "All hotkeys paused. Press the same shortcut to resume." : app.activationHotkeysSuspended ? "All-hotkey pause cleared. EVE switching remains paused." : "All hotkeys resumed"
                case "HideThumbnailsHotkey":
                    app.ShowHideThumbnails()
                    message := app.HideThumbnails ? "Thumbnails hidden" : "Thumbnails shown"
                case "ClickThroughHotkey":
                    app.Toggle_ClickThrough()
                    message := app.ClickThroughActive ? "Thumbnail click-through enabled" : "Thumbnail click-through disabled"
                case "ToggleLivePreviewsHotkey":
                    previous := app.livePreviewsPaused
                    app.ToggleLivePreviews()
                    message := previous = app.livePreviewsPaused ? "Live previews are busy; try again shortly." : app.livePreviewsPaused ? "Live previews paused" : "Live previews resumed"
                case "Close_Active_EVE_Win_Hotkey":
                    message := app.CloseActiveEVEWin() ? "Close requested for the focused EVE client" : "No focused EVE client to close"
                case "Close_All_EVE_Win_Hotkey":
                    app.HotkeyFeedback("Closing eligible EVE clients...")
                    count := app.CloseAllEVEWindows()
                    message := count ? "Close requested for " count " EVE client" (count = 1 ? "" : "s") : "No eligible EVE clients to close"
                case "Reload_Program_Hotkey", "Exit_Program_Hotkey":
                    action := id = "Exit_Program_Hotkey" ? "exit" : "reload"
                    if !app.RequestRuntimeAction(action, true)
                        app.HotkeyFeedback((action = "exit" ? "Exit" : "Reload") " blocked: finish, apply or discard Settings first.")
                    return
                default: return
            }
            app.HotkeyFeedback(message)
        } catch as err {
            app.HotkeyFeedback(HotkeyBindings.Settings.Get(id, id) " failed: " err.Message)
            ProgramLog.Error(err, "Command hotkey " id)
        }
    }
}
