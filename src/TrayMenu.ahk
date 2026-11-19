

Class TrayMenu extends Settings_Gui {
    usageStatistics := 0
    TrayMenuObj := A_TrayMenu
    Saved_overTray := 0
    Tray_Profile_scwitch := 0

    TrayMenu() {
        Profiles_Submenu := Menu()

        for k in This.Profiles {
            If (k = This.LastUsedProfile) {
                Profiles_Submenu.Add(This.LastUsedProfile, (name, *) => This.ActivateProfile(name, "Tray"))
                Profiles_Submenu.Check(This.LastUsedProfile)
            }
            Profiles_Submenu.Add(k, (name, *) => This.ActivateProfile(name, "Tray"))
        }

        TrayMenu := This.TrayMenuObj
        TrayMenu.Delete() ; Delete the Default TrayMenu Items

        TrayMenu.Add(T("Open"), MenuHandler.Bind("Open"))
        TrayMenu.Add(T("Open internal settings"), (*) => SettingsEditor.OpenInternal(This))
        TrayMenu.Add() ; Seperator
        TrayMenu.Add(T("Profiles"), Profiles_Submenu)
        TrayMenu.Add() ; Seperator

        if This.TrayMenuShortcuts["Suspend Hotkeys"] {
            TrayMenu.Add(T("Pause EVE switching hotkeys"), ObjBindMethod(This, "Suspend_Hotkeys"))
            if This.activationHotkeysSuspended
                TrayMenu.Check(T("Pause EVE switching hotkeys"))
        }
        if This.TrayMenuShortcuts.Get("Suspend All Hotkeys", 1) {
            TrayMenu.Add(T("Pause all hotkeys"), ObjBindMethod(This, "Suspend_All_Hotkeys"))
            if A_IsSuspended
                TrayMenu.Check(T("Pause all hotkeys"))
        }

        if This.TrayMenuShortcuts["Hide Thumbnails"] {
            TrayMenu.Add(T("Hide Thumbnails"), MenuHandler.Bind("Hide Thumbnails"))
            if This.HideThumbnails
                TrayMenu.check(T("Hide Thumbnails"))
            else
                TrayMenu.Uncheck(T("Hide Thumbnails"))
        }

        if This.TrayMenuShortcuts["Show Thumbnails Always on Top"] {
            TrayMenu.Add(T("Show Thumbnails Always on Top"), MenuHandler.Bind("Show Thumbnails Always on Top"))
            if This.ShowThumbnailsAlwaysOnTop
                TrayMenu.check(T("Show Thumbnails Always on Top"))
            else
                TrayMenu.Uncheck(T("Show Thumbnails Always on Top"))
        }

        if This.TrayMenuShortcuts["Click Through Thumbnails"] {
            TrayMenu.Add(T("Click Through Thumbnails"), MenuHandler.Bind("Click Through Thumbnails"))
            if This.ClickThroughActive
                TrayMenu.check(T("Click Through Thumbnails"))
            else
                TrayMenu.Uncheck(T("Click Through Thumbnails"))
        }

        if This.TrayMenuShortcuts["Minimize Inactive Clients"] {
            TrayMenu.Add(T("Minimize Inactive Clients"), MenuHandler.Bind("Minimize Inactive Clients"))
            if This.MinimizeInactiveClients
                TrayMenu.check(T("Minimize Inactive Clients"))
            else
                TrayMenu.Uncheck(T("Minimize Inactive Clients"))
        }

        if This.TrayMenuShortcuts["Suspend Hotkeys"] || This.TrayMenuShortcuts["Hide Thumbnails"] || This.TrayMenuShortcuts["Minimize Inactive Clients"] || This.TrayMenuShortcuts["Click Through Thumbnails"]
            TrayMenu.Add() ; Seperator

        if This.TrayMenuShortcuts["Don't Close Active Client"] {
            TrayMenu.Add(T("Don't Close Active Client"), MenuHandler.Bind("Don't Close Active Client"))
            if This.dontCloseActiveClient
                TrayMenu.check(T("Don't Close Active Client"))
            else
                TrayMenu.Uncheck(T("Don't Close Active Client"))
        }

        if This.TrayMenuShortcuts["Close all EVE Clients"]
            TrayMenu.Add(T("Close all EVE Clients"), (*) => This.CloseAllEVEWindows())

        if This.TrayMenuShortcuts["Close all EVE Clients"] || This.TrayMenuShortcuts["Don't Close Active Client"]
            TrayMenu.Add() ; Seperator

        if This.TrayMenuShortcuts["Restore Client Positions"] {
            TrayMenu.Add(T("Restore Client Positions"), MenuHandler.Bind("Restore Client Positions"))
            if (This.TrackClientPossitions)
                TrayMenu.check(T("Restore Client Positions"))
            else
                TrayMenu.Uncheck(T("Restore Client Positions"))
        }

        if This.TrayMenuShortcuts["Save Client Positions"]
            TrayMenu.Add(T("Save Client Positions"), (*) => This.Client_Possitions())

        if This.TrayMenuShortcuts["Restore Client Positions"] || This.TrayMenuShortcuts["Save Client Positions"]
            TrayMenu.Add() ; Seperator

        if This.TrayMenuShortcuts["Auto Save Thumbnail Positions"] {
            TrayMenu.Add(T("Auto Save Thumbnail Positions"), MenuHandler.Bind("Auto Save Thumbnail Positions"))
            if (This.AutoSaveThumbnailPositions)
                TrayMenu.check(T("Auto Save Thumbnail Positions"))
            else
                TrayMenu.Uncheck(T("Auto Save Thumbnail Positions"))
        }

        if This.TrayMenuShortcuts["Save Thumbnail Positions"]
            TrayMenu.Add(T("Save Thumbnail Positions"), MenuHandler.Bind("Save Thumbnail Positions"))

        if This.TrayMenuShortcuts["Auto Save Thumbnail Positions"] || This.TrayMenuShortcuts["Save Thumbnail Positions"]
            TrayMenu.Add() ; Seperator

        if This.TrayMenuShortcuts.Get("Lock Thumbnail Positions", 0) {
            TrayMenu.Add(T("Lock Thumbnail Positions"), ObjBindMethod(This, "ToggleLockThumbnailPositions"))
            if This.LockThumbnailPositions
                TrayMenu.Check(T("Lock Thumbnail Positions"))
        }
        if This.TrayMenuShortcuts.Get("Anonymize Thumbnails", 0) {
            TrayMenu.Add(T("Anonymize Thumbnails"), ObjBindMethod(This, "ToggleAnonymizeThumbnails"))
            if This.AnonymizeThumbnails
                TrayMenu.Check(T("Anonymize Thumbnails"))
        }

        if This.TrayMenuShortcuts.Get("Pause Live Previews", 0) {
            This.trayPreviewLabel := T(This.livePreviewsPaused ? "Resume Live Previews" : "Pause Live Previews")
            TrayMenu.Add(This.trayPreviewLabel, ObjBindMethod(This, "ToggleLivePreviews"))
        } else if This.HasOwnProp("trayPreviewLabel")
            This.DeleteProp("trayPreviewLabel")
        TrayMenu.Add(T("Reload"), (*) => This.RequestRuntimeAction("reload"))
        TrayMenu.Add(T("Exit"), (*) => This.RequestRuntimeAction("exit"))
        TrayMenu.Default := T("Open")

        MenuHandler(ItemName, *) {
            if (ItemName = "Auto Save Thumbnail Positions") {
                This.AutoSaveThumbnailPositions := !This.AutoSaveThumbnailPositions
                if This.AutoSaveThumbnailPositions ; if turned on save thumbnails
                    This.Save_ThumbnailPossitions()
                TrayMenu.ToggleCheck(T("Auto Save Thumbnail Positions"))
                SetTimer(This.Save_Settings_Delay_Timer, -200)
            }
            Else if (ItemName = "Save Thumbnail Positions") {
                ; Saved Thumbnail Positions only if the Saved button is used on the Traymenu
                This.Save_ThumbnailPossitions()
            }
            Else if (ItemName = "Restore Client Positions") {
                This.TrackClientPossitions := !This.TrackClientPossitions
                TrayMenu.ToggleCheck(T("Restore Client Positions"))
                SetTimer(This.Save_Settings_Delay_Timer, -200)
            }
            Else if (ItemName = "Hide Thumbnails") {
                This.ShowHideThumbnails()
                if This.HideThumbnails
                    TrayMenu.check(T("Hide Thumbnails"))
                else
                    TrayMenu.Uncheck(T("Hide Thumbnails"))
            }
            Else if (ItemName = "Show Thumbnails Always on Top") {
                This.ShowThumbnailsAlwaysOnTop := !This.ShowThumbnailsAlwaysOnTop
                TrayMenu.ToggleCheck(T("Show Thumbnails Always on Top"))
                SetTimer(This.Save_Settings_Delay_Timer, -200)
                Sleep(300)
                This.RequestRuntimeAction("reload")
            }
            Else if (ItemName = "Click Through Thumbnails") {
                This.Toggle_ClickThrough()
                if This.ClickThroughActive
                    TrayMenu.check(T("Click Through Thumbnails"))
                else
                    TrayMenu.Uncheck(T("Click Through Thumbnails"))
            }
            Else if (ItemName = "Don't Close Active Client") {
                This.dontCloseActiveClient := !This.dontCloseActiveClient
                TrayMenu.ToggleCheck(T("Don't Close Active Client"))
                SetTimer(This.Save_Settings_Delay_Timer, -200)
            }
            Else if (ItemName = "Minimize Inactive Clients") {
                This.MinimizeInactiveClients := !This.MinimizeInactiveClients
                TrayMenu.ToggleCheck(T("Minimize Inactive Clients"))
                SetTimer(This.Save_Settings_Delay_Timer, -200)
            }
            Else if (ItemName = "Open") {
                This.ShowSettings()
            }
        }
    }

    ActivateProfile(profileName, source := "Hotkey", *) {
        if !This.Profiles.Has(profileName)
            return false
        if This.HasOwnProp("settingsBridge") && !This.settingsBridge.GuardRuntimeAction()
            return false
        if This.HasOwnProp("settingsWorkbench") && !This.settingsWorkbench.closed {
            editor := This.settingsWorkbench
            if editor.HasOwnProp("activeDialog") || editor.HasOwnProp("closePrompt") {
                editor.UpdateStatus("Finish the current settings dialog before changing profiles.")
                return false
            }
            if editor.model.Profiles.Has(profileName) {
                editor.SelectProfile(profileName)
                editor.model.ActivateProfile()
                editor.Changed()
                editor.Navigate("profiles")
                editor.S_Gui.Show()
                editor.UpdateStatus("Profile activation is staged. Apply & restart activates it and your other draft changes.")
            } else
                editor.UpdateStatus("That profile was removed from the draft. Undo the removal or discard the draft first.")
            return false
        }
        ; Keep selection, save and reload together while login timers are active.
        previousCritical := A_IsCritical
        Critical
        try {
            SetTimer(This.Save_Settings_Delay_Timer, 0)
            ProgramLog.Add(source " profile switch requested; from=" This.LastUsedProfile "; to=" profileName)
            This.LastUsedProfile := profileName
            if source = "Hotkey"
                This._JSON["HotkeyFeedbackOnStartup"] := T("Profile activated: {1}", profileName)
            This.SaveJsonToFile()
            This.ReloadForProfile()
        } finally {
            Critical(previousCritical)
        }
        return true
    }

    ReloadForProfile() {
        Reload()
    }

    RequestRuntimeAction(action, hotkey := false) {
        if hotkey && This.HasOwnProp("pendingCommandHotkey") && This.pendingCommandHotkey {
            This.HotkeyFeedback(T("A reload or exit is already requested."))
            return true
        }
        if This.HasOwnProp("settingsBridge") && !This.settingsBridge.GuardRuntimeAction()
            return false
        if This.HasOwnProp("settingsWorkbench") && !This.settingsWorkbench.closed {
            editor := This.settingsWorkbench
            if editor.HasOwnProp("activeDialog") || editor.HasOwnProp("closePrompt") {
                editor.UpdateStatus("Finish the current settings dialog before exiting or reloading.")
                return false
            }
            if editor.HasDraftChanges() {
                editor.S_Gui.Show()
                editor.UpdateStatus(action = "exit" ? "Apply or discard your settings draft before exiting." : "Apply or discard your settings draft before reloading.")
                return false
            }
            editor.Dispose()
        }
        if This.HasOwnProp("Save_Settings_Delay_Timer")
            SetTimer(This.Save_Settings_Delay_Timer, 0)
        This.SaveJsonToFile()
        if hotkey {
            This.pendingCommandHotkey := true
            This.HotkeyFeedback(T(action = "exit" ? "Exit requested" : "Reload requested"))
            SetTimer(ObjBindMethod(This, "CompleteCommandHotkey", action), -650)
        } else
            This.PerformRuntimeAction(action)
        return true
    }

    CompleteCommandHotkey(action) {
        This.pendingCommandHotkey := false
        if !This.RequestRuntimeAction(action)
            This.HotkeyFeedback(T("Command cancelled: finish, apply or discard Settings first."))
    }

    PerformRuntimeAction(action) {
        if action = "exit"
            ExitApp()
        else
            Reload()
    }

    ShowHideThumbnails() {
        This.HideThumbnails := !This.HideThumbnails
        if This.HideThumbnails
            This.ShowHideAllThumbnails("Hide")
        else
            This.ShowHideAllThumbnails("Show")
        This.BorderActive := 0
        SetTimer(This.Save_Settings_Delay_Timer, -200)
    }

    CloseAllEVEWindows(*) {
        requested := 0
        try {
            list := WinGetList(This.EVEExe)
            GroupAdd("EVE", "Ahk_Exe exefile.exe")
            for hwnd in list {
                WinTitle := WinGetTitle(hwnd)

                if This.dontCloseActiveClient && WinActive("ahk_id " hwnd)
                    continue

                if This.DontCloseWIn(WinTitle, hwnd)
                    continue

                PostMessage 0x0112, 0xF060, , , hwnd
                requested++
                Sleep(50)
            }
        }
        return requested
    }
}
