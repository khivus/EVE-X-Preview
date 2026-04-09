

Class TrayMenu extends Settings_Gui {
    TrayMenuObj := A_TrayMenu
    Saved_overTray := 0
    Tray_Profile_scwitch := 0

    TrayMenu() {
        Profiles_Submenu := Menu()

        for k in This.Profiles {
            If (k = This.LastUsedProfile) {
                Profiles_Submenu.Add(This.LastUsedProfile, MenuHandler)
                Profiles_Submenu.Check(This.LastUsedProfile)
            }
            Profiles_Submenu.Add(k, MenuHandler)
        }

        TrayMenu := This.TrayMenuObj
        TrayMenu.Delete() ; Delete the Default TrayMenu Items

        TrayMenu.Add("Open", MenuHandler)
        TrayMenu.Add() ; Seperator
        TrayMenu.Add("Profiles", Profiles_Submenu)
        TrayMenu.Add() ; Seperator

        if This.TrayMenuShortcuts["Suspend Hotkeys"]
            TrayMenu.Add("Suspend Hotkeys", MenuHandler)

        if This.TrayMenuShortcuts["Hide Thumbnails"] {
            TrayMenu.Add("Hide Thumbnails", MenuHandler)
            if This.HideThumbnails
                TrayMenu.check("Hide Thumbnails")
            else
                TrayMenu.Uncheck("Hide Thumbnails")
        }

        if This.TrayMenuShortcuts["Show Thumbnails Always on Top"] {
            TrayMenu.Add("Show Thumbnails Always on Top", MenuHandler)
            if This.ShowThumbnailsAlwaysOnTop
                TrayMenu.check("Show Thumbnails Always on Top")
            else
                TrayMenu.Uncheck("Show Thumbnails Always on Top")
        }

        if This.TrayMenuShortcuts["Click Through Thumbnails"] {
            TrayMenu.Add("Click Through Thumbnails", MenuHandler)
            if This.ClickThroughActive
                TrayMenu.check("Click Through Thumbnails")
            else
                TrayMenu.Uncheck("Click Through Thumbnails")
        }

        if This.TrayMenuShortcuts["Minimize Inactive Clients"] {
            TrayMenu.Add("Minimize Inactive Clients", MenuHandler)
            if This.MinimizeInactiveClients
                TrayMenu.check("Minimize Inactive Clients")
            else
                TrayMenu.Uncheck("Minimize Inactive Clients")
        }

        if This.TrayMenuShortcuts["Suspend Hotkeys"] || This.TrayMenuShortcuts["Hide Thumbnails"] || This.TrayMenuShortcuts["Minimize Inactive Clients"] || This.TrayMenuShortcuts["Click Through Thumbnails"]
            TrayMenu.Add() ; Seperator

        if This.TrayMenuShortcuts["Don't Close Active Client"] {
            TrayMenu.Add("Don't Close Active Client", MenuHandler)
            if This.dontCloseActiveClient
                TrayMenu.check("Don't Close Active Client")
            else
                TrayMenu.Uncheck("Don't Close Active Client")
        }

        if This.TrayMenuShortcuts["Close all EVE Clients"]
            TrayMenu.Add("Close all EVE Clients", (*) => This.CloseAllEVEWindows())

        if This.TrayMenuShortcuts["Close all EVE Clients"] || This.TrayMenuShortcuts["Don't Close Active Client"]
            TrayMenu.Add() ; Seperator

        if This.TrayMenuShortcuts["Restore Client Positions"] {
            TrayMenu.Add("Restore Client Positions", MenuHandler)
            if (This.TrackClientPossitions)
                TrayMenu.check("Restore Client Positions")
            else
                TrayMenu.Uncheck("Restore Client Positions")
        }

        if This.TrayMenuShortcuts["Save Client Positions"]
            TrayMenu.Add("Save Client Positions", (*) => This.Client_Possitions())

        if This.TrayMenuShortcuts["Restore Client Positions"] || This.TrayMenuShortcuts["Save Client Positions"]
            TrayMenu.Add() ; Seperator

        if This.TrayMenuShortcuts["Auto Save Thumbnail Positions"] {
            TrayMenu.Add("Auto Save Thumbnail Positions", MenuHandler)
            if (This.AutoSaveThumbnailPositions)
                TrayMenu.check("Auto Save Thumbnail Positions")
            else
                TrayMenu.Uncheck("Auto Save Thumbnail Positions")
        }

        if This.TrayMenuShortcuts["Save Thumbnail Positions"]
            TrayMenu.Add("Save Thumbnail Positions", MenuHandler)

        if This.TrayMenuShortcuts["Auto Save Thumbnail Positions"] || This.TrayMenuShortcuts["Save Thumbnail Positions"]
            TrayMenu.Add() ; Seperator

        TrayMenu.Add("Reload", (*) => Reload())
        TrayMenu.Add("Exit", (*) => ExitApp())
        TrayMenu.Default := "Open"

        MenuHandler(ItemName, ItemPos, MyMenu) {
            If (ItemName = "Exit")
                ExitApp
            Else if (ItemName = "Auto Save Thumbnail Positions") {
                This.AutoSaveThumbnailPositions := !This.AutoSaveThumbnailPositions
                if This.AutoSaveThumbnailPositions ; if turned on save thumbnails
                    This.Save_ThumbnailPossitions
                TrayMenu.ToggleCheck("Auto Save Thumbnail Positions")
                SetTimer(This.Save_Settings_Delay_Timer, -200)
            }
            Else if (ItemName = "Save Thumbnail Positions") {
                ; Saved Thumbnail Positions only if the Saved button is used on the Traymenu
                This.Save_ThumbnailPossitions
            }
            Else if (ItemName = "Restore Client Positions") {
                This.TrackClientPossitions := !This.TrackClientPossitions
                TrayMenu.ToggleCheck("Restore Client Positions")
                SetTimer(This.Save_Settings_Delay_Timer, -200)
            }
            Else if (ItemName = "Hide Thumbnails") {
                This.ShowHideThumbnails()
                if This.HideThumbnails
                    TrayMenu.check("Hide Thumbnails")
                else
                    TrayMenu.Uncheck("Hide Thumbnails")
            }
            Else if (ItemName = "Show Thumbnails Always on Top") {
                This.ShowThumbnailsAlwaysOnTop := !This.ShowThumbnailsAlwaysOnTop
                TrayMenu.ToggleCheck("Show Thumbnails Always on Top")
                SetTimer(This.Save_Settings_Delay_Timer, -200)
                Sleep(300)
                Reload()
            }
            Else if (ItemName = "Click Through Thumbnails") {
                This.Toggle_ClickThrough()
                if This.ClickThroughActive
                    TrayMenu.check("Click Through Thumbnails")
                else
                    TrayMenu.Uncheck("Click Through Thumbnails")
            }
            Else if (ItemName = "Don't Close Active Client") {
                This.dontCloseActiveClient := !This.dontCloseActiveClient
                TrayMenu.ToggleCheck("Don't Close Active Client")
                SetTimer(This.Save_Settings_Delay_Timer, -200)
            }
            Else if (ItemName = "Minimize Inactive Clients") {
                This.MinimizeInactiveClients := !This.MinimizeInactiveClients
                TrayMenu.ToggleCheck("Minimize Inactive Clients")
                SetTimer(This.Save_Settings_Delay_Timer, -200)
            }
            Else if (This.Profiles.Has(ItemName)) {
                ; Change the lastUsedProfile to the Profile name, save it to Json file and reload the script with the new Settings
                This.LastUsedProfile := ItemName
                This.SaveJsonToFile()
                Sleep(500)
                Reload()
            }
            Else if (ItemName = "Open") {
                if WinExist("EVE-X-Preview - Settings") {
                    WinActivate("EVE-X-Preview - Settings")
                    Return
                }
                This.MainGui()
            }
            Else If (ItemName = "Suspend Hotkeys") {
                Suspend(-1)
                if A_IsSuspended
                    TrayMenu.check("Suspend Hotkeys")
                else
                    TrayMenu.Uncheck("Suspend Hotkeys")
            }
        }
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
        try {
            list := WinGetList("Ahk_Exe exefile.exe")
            GroupAdd("EVE", "Ahk_Exe exefile.exe")
            for hwnd in list {
                WinTitle := WinGetTitle(hwnd)

                if This.dontCloseActiveClient && WinActive("ahk_id " hwnd)
                    continue

                if This.DontCloseWIn(WinTitle)
                    continue

                PostMessage 0x0112, 0xF060, , , hwnd
                Sleep(50)
            }
        }
    }
}
