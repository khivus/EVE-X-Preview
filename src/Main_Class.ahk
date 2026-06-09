Class Main_Class extends ThumbWindow {
    Static  WM_DESTROY := 0x02,
            WM_SIZE := 0x05,
            WM_NCCALCSIZE := 0x83,
            WM_NCHITTEST := 0x84,
            WM_NCLBUTTONDOWN := 0xA1,
            WM_SYSKEYDOWN := 0x104,
            WM_SYSKEYUP := 0x105,
            WM_MOUSEMOVE := 0x200,
            WM_LBUTTONDOWN := 0x201,
            WM_LBUTTONUP := 0x0202,
            WM_RBUTTONDOWN := 0x0204,
            WM_RBUTTONUP := 0x0205,
            WM_KEYDOWN := 0x0100,
            WM_MOVE := 0x03,
            WM_MOUSELEAVE := 0x02A2

    ;! This key is for the internal Hotkey to bring the Window in forgeround 
    ;! it is possible this key needs to be changed if Windows updates and changes the unused virtual keys 
    static virtualKey := "vk0xE8"

    LISTENERS := [
        Main_Class.WM_LBUTTONDOWN,
        Main_Class.WM_RBUTTONDOWN
        ;Main_Class.WM_SIZE,
        ;Main_Class.WM_MOVE
    ]


    EVEExe := "ahk_Exe exefile.exe"
    
    ; Values for WM_NCHITTEST
    ; Size from the invisible edge for resizing    
    border_size := 4
    HT_VALUES := [[13, 12, 14], [10, 1, 11], [16, 15, 17]]

    ;### Predifining Arrays and Maps #########
    EventHooks := Map() 
    ThumbWindows := {}
    ThumbHwnd_EvEHwnd := Map()

    __New() { 

        This._JSON := Load_JSON()
        This.default_JSON := JSON.Load(default_JSON)
       
        This.ProfileOverride()
        This.TrayMenu()
        This.MinimizeDelay := This.Minimizeclients_Delay    
        
        ;Hotkey to trigger by the script to get permissions t bring a Window in foreground
        ;Register all posible modifire combinations 
        prefixArr := ["","^","!", "#", "+", "+^", "+#", "+!", "^#", "^!","#!", "^+!", "^+#", "^#!", "+!#","^+#!"]
        for index, prefix in prefixArr
            Hotkey(  prefix . Main_Class.virtualKey, ObjBindMethod(This, "ActivateForgroundWindow"), "S P1")

        This.Save_Settings_Delay_Timer := ObjBindMethod(This, "SaveJsonToFile")

        if This.First_Start_After_Update { ; Display message after succsessful update
            Sleep 500 ; Let updater time to close
            SetWorkingDir(A_ScriptDir)
            updaterExe := "EVE-X-Preview-Updater.exe" ; Delete updater exe if still exists
            if FileExist(updaterExe)
                FileDelete(updaterExe)

            This.First_Start_After_Update := 0
            SetTimer(This.Save_Settings_Delay_Timer, -200)

            Version := FileGetVersion(A_ScriptName)
            MsgBox("EVE-X-Preview succsessfully updated to version " Version)
        }

        ; Register Hotkey for Puase Hotkeys if the user has is Set
        if (This.Suspend_Hotkeys_Hotkey != "") {
            HotIf (*) => This.anyWinExists()
            if !This.SwitchLangOnErr {
                try
                    Hotkey This.Suspend_Hotkeys_Hotkey, ( * ) => This.Suspend_Hotkeys(), "S1"
                catch ValueError as e
                    MsgBox(e.Message ": --> " e.Extra " <-- in: Global Settings -> Suspend Hotkeys-Hotkey" )
            }
            else
                Hotkey This.Suspend_Hotkeys_Hotkey, ( * ) => This.Suspend_Hotkeys(), "S1"
        }

        ; Register Hotkey for Hide Thumbnails if the user has is Set
        if (This.HideThumbnailsHotkey != "") {
            HotIf (*) => This.anyWinExists()
            if !This.SwitchLangOnErr {
                try
                    Hotkey This.HideThumbnailsHotkey, ( * ) => This.ShowHideThumbnails(), "S1"
                catch ValueError as e
                    MsgBox(e.Message ": --> " e.Extra " <-- in: Global Settings -> Hide Thumbnails Hotkey" )
            }
            else
                Hotkey This.HideThumbnailsHotkey, ( * ) => This.ShowHideThumbnails(), "S1"
        }

        ; Register Hotkey for Click Through if the user has is Set
        if (This.ClickThroughHotkey != "") {
            HotIf (*) => This.anyWinExists()
            if !This.SwitchLangOnErr {
                try
                    Hotkey This.ClickThroughHotkey, ( * ) => This.Toggle_ClickThrough(), "S1"
                catch ValueError as e
                    MsgBox(e.Message ": --> " e.Extra " <-- in: Global Settings -> Click Through Thumbnails Hotkey" )
            }
            else
                Hotkey This.ClickThroughHotkey, ( * ) => This.Toggle_ClickThrough(), "S1"
        }

        ; Register Hotkey for Login Screen Cycle Hotkey if user set
        if (This.Login_Screen_Cycle_Hotkey != "") {
            if This.Global_Hotkeys
                HotIf (*) => WinExist(This.EVEExe)
            else
                HotIf (*) => WinActive(This.EVEExe)
            
            if !This.SwitchLangOnErr {
                try
                    Hotkey(This.Login_Screen_Cycle_Hotkey, ObjBindMethod(This, "Cycle_Login_Windows"),"P1" )
                catch ValueError as e
                    MsgBox(e.Message ": --> " e.Extra " <-- in Login Screen Cycle Hotkey")
            }
            else
                Hotkey(This.Login_Screen_Cycle_Hotkey, ObjBindMethod(This, "Cycle_Login_Windows"),"P1" )
        }

        ; Register Hotkey for Close Active EVE Window Hotkey if user set
        if (This.Close_Active_EVE_Win_Hotkey != "") {
            HotIf (*) => WinExist(This.EVEExe)
            if !This.SwitchLangOnErr {
                try
                    Hotkey(This.Close_Active_EVE_Win_Hotkey, ObjBindMethod(This, "CloseActiveEVEWin"),"P1" )
                catch ValueError as e
                    MsgBox(e.Message ": --> " e.Extra " <-- in Close Active EVE Window Hotkey")
            }
            else
                Hotkey(This.Close_Active_EVE_Win_Hotkey, ObjBindMethod(This, "CloseActiveEVEWin"),"P1" )
        }

        ; Register Hotkey for Close All EVE Windows Hotkey if user set
        if (This.Close_All_EVE_Win_Hotkey != "") {
            HotIf (*) => WinExist(This.EVEExe)
            if !This.SwitchLangOnErr {
                try
                    Hotkey(This.Close_All_EVE_Win_Hotkey, ObjBindMethod(This, "CloseAllEVEWindows"),"P1" )
                catch ValueError as e
                    MsgBox(e.Message ": --> " e.Extra " <-- in Close All EVE Windows Hotkey")
            }
            else
                Hotkey(This.Close_All_EVE_Win_Hotkey, ObjBindMethod(This, "CloseAllEVEWindows"),"P1" )
        }

        ; Register Hotkey for Reload EVE-X-Preview Hotkey if user set
        if (This.Reload_Program_Hotkey != "") {
            HotIf (*) => This.anyWinExists()
            if !This.SwitchLangOnErr {
                try
                    Hotkey This.Reload_Program_Hotkey, ( * ) => Reload(), "S1"
                catch ValueError as e
                    MsgBox(e.Message ": --> " e.Extra " <-- in Reload EVE-X-Preview Hotkey")
            }
            else
                Hotkey This.Reload_Program_Hotkey, ( * ) => Reload(), "S1"
        }

        HotIf() ; Reset hotif

        ; Profiling
        This.ProfActive := false
        ProfEnabled := false
        if ProfEnabled
            This.StartProfiling()

        ; Map of ignored chars in hotkey groups
        This.ignoredChars := Map()

        ; Inited monitoring check
        This.monitoringInitialized := 0

        ; Initiate the variable to store the last active thumbnail hwnd
        This.LastActiveThumbHwnd := 0

        ; Resets the position of Shifting thubmnails
        This.allLoginClosed := false

        ; Skips collision check and thumb move if thumb "touch" screen edge
        This.skipShiftThumbs := false

        ; List for NonEVE Apps to manage thumbnails
        This.CreateNonEVEAppsList()
        This.allTrackedApps := Map("exefile.exe", true)
        for NonEVEapp in This.NonEVEAppsList {
            This.allTrackedApps[NonEVEapp["exe"]] := true
        }

        ; The Timer property for Asycn Minimizing.
        this.timer := ObjBindMethod(this, "EVEMinimize")
        
        ;margins for DwmExtendFrameIntoClientArea. higher values extends the shadow
        This.margins := Buffer(16, 0)
        NumPut("Int", 0, This.margins)
        
        ;Register all messages wich are inside LISTENERS
        for i, message in this.LISTENERS
            OnMessage(message, ObjBindMethod(This, "_OnMessage"))

        ;Property for the delay to hide Thumbnails if not client is in foreground and user has set Hide on lost Focus
        This.CheckforActiveWindow := ObjBindMethod(This, "HideOnLostFocusTimer")

        ;The Main Timer who checks for new EVE Windows or closes Windows 
        SetTimer(ObjBindMethod(This, "HandleMainTimer"), 50)
        
        ;Timer property to remove Thumbnails for closed EVE windows 
        This.DestroyThumbnails := ObjBindMethod(This, "EvEWindowDestroy")
        This.DestroyThumbnailsToggle := 1
        
        ;Register the Hotkeys for cycle groups 
        This.Register_Hotkey_Groups()
        This.RegisterNonEVEGroups()
        This.BorderActive := 0

        ; Trigger logs monitoring
        This.gameLogsMonitoring()

        return This
    }


    ; Profiling for optimization testing
    StartProfiling(minutes := 1) {
        This.TickCount := 0
        This.TotalTime := 0
        This.MaxTime := 0

        This.ProfStart := A_TickCount
        This.ProfActive := true

        SetTimer(ObjBindMethod(This, "StopProfiling"), -(minutes * 60000))
    }

    
    StopProfiling() {
        This.ProfActive := false

        elapsedS := (A_TickCount - This.ProfStart) / 1000
        avg := This.TotalTime / This.TickCount
        rate := This.TickCount / This.TotalTime

        MsgBox(
            "Profiling results:`n`n"
            "Duration: " Round(elapsedS, 1) " s`n"
            "Total Time: " This.TotalTime " ms`n"
            "Ticks: " This.TickCount "`n"
            "Rate: " Round(rate, 2) "/ms`n"
            "Avg time: " Round(avg, 2) " ms`n"
            "Max spike: " This.MaxTime " ms"
        )
    }


    HandleMainTimer() {
        ; if This.ProfActive ; Profiling
        static __t1 := 0
        __t0 := A_TickCount
        if __t0 - __t1 > 1000 { ; Check every second to avoid lag and not check for new windows on every tick
            This.UpdateActiveNonEVEApps()
            __t1 := __t0
        }

        static HideShowToggle := 0, LastActiveHWND := 0, WinList := [], activeExe := ""

        try
            WinList := WinGetList(This.EVEExe)
        Catch 
            return

        WinList.Push(This.ActiveNonEVEApps*)
        ; If any EVE Window exist
        if WinList.Length {
            try { ; Another++ attempt to fix error on thumbnail destruction
                ;Check if a window exist without Thumbnail and if the user is in Character selection screen or not
                for index, hwnd in WinList {
                    WinList.%hwnd% := { Title: WinGetTitle(hwnd) }

                    if !This.ThumbWindows.HasProp(hwnd) {
                        This.EVE_WIN_Created(hwnd, WinList.%hwnd%.title)
                        if (!This.HideThumbnailsOnLostFocus) {
                            This.ShowThumb(hwnd, "Show")
                        }                      
                        HideShowToggle := 1
                    }
                    else { ; This change improved performance by ~15-17%
                        ; Writes character name to OldTitle if PreserveHotkeysOnLogout enabled
                        if This.PreserveHotkeysOnLogout && This.ThumbWindows.%hwnd%["Window"].Title != "EVE" && This.ThumbWindows.%hwnd%["Window"].Title != "Char Screen" {
                            This.ThumbWindows.%hwnd%["Window"].OldTitle := This.ThumbWindows.%hwnd%["Window"].Title
                        }

                        ; If PreserveThumbPosOnLogout is false we move thumbnail after logout to default position
                        if !This.PreserveThumbPosOnLogout && This.ThumbWindows.%hwnd%["Window"].Title != "Char Screen" && This.ThumbWindows.%hwnd%["Window"].Title != WinList.%hwnd%.Title && WinList.%hwnd%.Title == "EVE" {
                            if This.ShiftThumbsForLoginScreen
                                This.ShiftThumbs(hwnd)
                            else
                                This.ThumbMove( This.ThumbnailStartLocation["x"],
                                                This.ThumbnailStartLocation["y"],
                                                This.ThumbnailStartLocation["width"],
                                                This.ThumbnailStartLocation["height"],
                                                This.ThumbWindows.%hwnd%)
                        }

                        if This.monitoringInitialized && WinList.%hwnd%.Title != "EVE" && !This.monitoredChars.Has(WinList.%hwnd%.Title) && !This.waitingMonitoringChars.Has(WinList.%hwnd%.Title) { ; Check if for some reason char isn't monitored and add it to monitoring list
                            ; ToolTip "Adding new -> " WinList.%hwnd%.Title
                            ; SetTimer () => ToolTip(), -1000
                            This.waitingMonitoringChars[WinList.%hwnd%.Title] := ObjBindMethod(This, "startLogMonitoring", WinList.%hwnd%.Title, This.charsIds.Has(WinList.%hwnd%.Title) ? This.charsIds[WinList.%hwnd%.Title] : 0)
                            SetTimer(This.waitingMonitoringChars[WinList.%hwnd%.Title], -5000) ; Wait 5 sec to initialize file
                        }
                    
                        ; if in Character selection screen 
                        if (This.ThumbWindows.%hwnd%["Window"].Title != WinList.%hwnd%.Title && WinList.%hwnd%.Title = "EVE" && This.PreserveCharNameOnLogout) {
                            if This.monitoringInitialized && This.monitoredChars.Has(This.ThumbWindows.%hwnd%["Window"].Title)
                                This.stopLogMonitoring(This.ThumbWindows.%hwnd%["Window"].Title)
                            This.ThumbWindows.%hwnd%["Window"].Title := "Char Screen"
                        }
                        else if (This.ThumbWindows.%hwnd%["Window"].Title != WinList.%hwnd%.Title) {
                            if This.monitoringInitialized && This.monitoredChars.Has(This.ThumbWindows.%hwnd%["Window"].Title)
                                This.stopLogMonitoring(This.ThumbWindows.%hwnd%["Window"].Title)
                            This.EVENameChange(hwnd, WinList.%hwnd%.Title)
                        }
                    }
                }
            }
            catch
                This.debugToolTip("Error in Main Timer EVE Window Check")

            try {
                ;if HideThumbnailsOnLostFocus is selectet check if a eve window is still in foreground, runs a timer once with a delay to prevent stuck thumbnails
                Ahwnd := WinExist("A")
                if Ahwnd != LastActiveHWND
                    activeExe := WinGetProcessName("A")
                activeWinTracked := 0
                if This.allTrackedApps.Has(activeExe) {
                    activeWinTracked := 1
                }

                if This.HideThumbnailsOnLostFocus && !HideShowToggle && !activeWinTracked {
                    SetTimer(This.CheckforActiveWindow, -500)
                    HideShowToggle := 1
                }
                else if activeWinTracked {
                    if This.HideThumbForActiveWin && !HideShowToggle {
                        This.ShowThumb(Ahwnd, "Hide")

                        if LastActiveHWND && LastActiveHWND != Ahwnd {
                            This.ShowThumb(LastActiveHWND, "Show")
                            This.UpdateThumb_AfterActivation(, Ahwnd)
                        }
                        
                    }
                    else if HideShowToggle {
                        for EVEHWND in This.ThumbWindows.OwnProps() {
                            This.ShowThumb(EVEHWND, "Show")
                        }
                        HideShowToggle := 0
                        This.BorderActive := 0
                    }
                    ; sets the Border to the active window thumbnail 
                    else if Ahwnd != This.BorderActive {
                        ; Shows the Thumbnail on top of other thumbnails
                        if This.ShowThumbnailsAlwaysOnTop
                            WinSetAlwaysOnTop(1,This.ThumbWindows.%Ahwnd%["Window"].Hwnd )
                        
                        This.ShowActiveBorder(Ahwnd)
                        This.UpdateThumb_AfterActivation(, Ahwnd)
                        This.BorderActive := Ahwnd
                    }
                }
                LastActiveHWND := Ahwnd
            }
            catch
                This.debugToolTip("Error in Main Timer Active Window Check")
        }

        ; Check if a Thumbnail exist without EVE Window. if so destroy the Thumbnail and free memory
        if ( This.DestroyThumbnailsToggle ) {
            try {
                for k, v in This.ThumbWindows.Clone().OwnProps() {
                    if !Winlist.HasProp(k) {
                        SetTimer(This.DestroyThumbnails, -500)
                        This.DestroyThumbnailsToggle := 0
                    }
                }
                if !WinList.Length
                    This.allLoginClosed := true
                else {
                    for EVEHWND in This.ThumbWindows.OwnProps() {
                        if This.ThumbWindows.%EVEHWND%["Window"].Title == "EVE" {
                            This.allLoginClosed := false
                            break
                        }
                        This.allLoginClosed := true
                    }
                }
            }
        }

        if This.ProfActive { ; Profiling
            elapsed := A_TickCount - __t0

            This.TickCount++
            This.TotalTime += elapsed

            if elapsed > This.MaxTime
                This.MaxTime := elapsed
        }
    }
    

    ; The function for the timer which gets started if no EVE window is in focus 
    HideOnLostFocusTimer() {
        Try {
            ForegroundPName := WinGetProcessName("A")
            if (ForegroundPName = "exefile.exe") {
                if (DllCall("IsIconic", "UInt", WinActive("ahk_exe exefile.exe"))) {
                    for EVEHWND in This.ThumbWindows.OwnProps() {
                        This.ShowThumb(EVEHWND, "Hide")
                    }
                }
            }
            else if (ForegroundPName != "exefile.exe") {
                for EVEHWND in This.ThumbWindows.OwnProps() {
                    This.ShowThumb(EVEHWND, "Hide")
                }
            }
        }
    }

    anyWinExists() {
        try
            if WinGetList(This.EVEExe).Length != 0 || This.ActiveNonEVEApps.Length != 0
                return true

        return false
    }

    ;Register set Hotkeys by the user in settings
    RegisterHotkeys(title, EvE_hwnd) {
        if !This._Hotkeys[title]
            return

        This.debugToolTip("Registering hotkeys for " title)
    
        if This.Global_Hotkeys
            HotIf (*) => WinExist(title " " This.EVEExe) && !WinActive("EVE-X-Preview - Settings")
        else    
            HotIf (*) => WinExist(title " " This.EVEExe) && WinActive(This.EVEExe) && !WinActive("EVE-X-Preview - Settings")
    
        if !This.SwitchLangOnErr {
            try
                Hotkey This._Hotkeys[title], (*) => This.ActivateEVEWindow(,,title, 1), "P1"
            catch ValueError as e
                MsgBox(e.Message ": --> " e.Extra " <-- in Hotkey Settings - " This.LastUsedProfile " Hotkeys")
        } 
        else
            Hotkey This._Hotkeys[title], (*) => This.ActivateEVEWindow(,,title, 1), "P1"
    }    

    ;Register the Hotkeys for cycle Groups if any set
    Register_Hotkey_Groups() {
        keys := Map()
        This.HotkGroups := []
        This.HotkGroupsInds := []
        index := 1

        if !IsObject(This.Hotkey_Groups) || This.Hotkey_Groups.Count = 0
            return

        This.debugToolTip("Registering Hotkey Groups")

        for k, v in This.Hotkey_Groups {
            This.HotkGroups.Push(v["Characters"])
            This.HotkGroupsInds.Push(0)

            if This.Global_Hotkeys
                method := "OnWinExist"
            else
                method := "OnWinActive"

            if v["ForwardsHotkey"] != ""
                keys["ForwardsHotkey"] := v["ForwardsHotkey"]
            if v["BackwardsHotkey"] != ""
                keys["BackwardsHotkey"] := v["BackwardsHotkey"]

            HotIf ObjBindMethod(This, method, This.HotkGroups[index])
            for direction, key in keys
                if !This.SwitchLangOnErr {
                    try
                        Hotkey(key, ObjBindMethod(This, "Cycle_Hotkey_Groups", index, direction), "P1")
                    catch ValueError as e
                        MsgBox(e.Message ": --> " e.Extra " <-- in Hotkeys Groups - " This.LastUsedProfile " - " k "  - " direction)
                }
                else
                    Hotkey(key, ObjBindMethod(This, "Cycle_Hotkey_Groups", index, direction), "P1")

            index += 1
            keys := Map()
        }
    }

    ; The method to make it possible to cycle throw the EVE Windows. Used with the Hotkey Groups
    Cycle_Hotkey_Groups(ArrInd, direction, *) {
        static tick, prevTick := 0
        tick := A_TickCount
        if tick - prevTick < This.GroupsHoldDelay
            return
        prevTick := tick

        arr := This.HotkGroups[ArrInd]
        ; index := This.HotkGroupsInds[ArrInd]
        length := arr.Length

        ; if !This.KeepGroupsPositions {
        ;     try {
        ;         loop 2 {
        ;             title := This.PreserveHotkeysOnLogout ? This.ThumbWindows.%WinExist("A")%["Window"].OldTitle : WinGetTitle("A")
        ;             index := IsActiveWinInGroup(title, arr)

        ;             if index = -1 && A_Index = 1
        ;                 Sleep 50 ; Delay so if game lags it checks 2 times and not return to 1 win in middle of a cycle group
        ;             else if index = -1
        ;                 index += 1
        ;             else
        ;                 break
        ;         }
        ;     }
        ;     catch
        ;         index := 0
        ; }

        if This.KeepGroupsPositions {
            ; Trust stored index always — no re-sync ever
            index := This.HotkGroupsInds[ArrInd]
        } else {
            ; Derive from active window, reset to start if not found
            currentIndex := This._GetCurrentGroupIndex(arr)
            index := currentIndex = -1 ? 0 : currentIndex
        }

        index := DirectionHandler(direction, index, length)
    
        if !This.OnWinExist(arr)
            return

        Loop length { ; Using loop len instead of while to avoid infinite while
            hwndEVE := WinExist(arr[index] " ahk_exe exefile.exe")
            if hwndEVE && !This.ignoredChars.Has(hwndEVE)
                break
            else if hwndTemp := This.hasMathcingOldTitle(arr[index]) {
                if !This.ignoredChars.Has(hwndTemp) {
                    hwndEVE := hwndTemp
                    break
                }
            }

            index := DirectionHandler(direction, index, length)
        }

        if This.ignoredChars.Has(hwndEVE) ; Last check if loop leaked window there
            return

        try
            This.ActivateEVEWindow(hwndEVE,,)

        ; Only persist index when KeepGroupsPositions is on
        if This.KeepGroupsPositions
            This.HotkGroupsInds[ArrInd] := index

        ; Get index by specified direction
        DirectionHandler(direction, index, length) {
            if (direction == "ForwardsHotkey") {
                index += 1
                if index > length
                    index := 1
            }
            else if (direction == "BackwardsHotkey") {
                index -= 1
                if index <= 0
                    index := length
            }
            return index
        }

        ; IsActiveWinInGroup(title, arr) {
        ;     for ind, names in arr {
        ;         if names = title
        ;             return ind
        ;     }
        ;     return -1
        ; }
    }

    ; Tries to find active window in group, with brief retry on lag
    _GetCurrentGroupIndex(arr, maxAttempts := 3, interval := 25) {
        loop maxAttempts {
            try {
                if This.PreserveHotkeysOnLogout
                    title := This.ThumbWindows.%WinExist("A")%["Window"].OldTitle
                else
                    title := WinGetTitle("A")

                for ind, name in arr {
                    if name = title
                        return ind
                }
            }
            ; Window not in group yet — only retry if lag is plausible
            ; (i.e. we just activated an EVE window and it hasn't focused yet)
            if A_Index < maxAttempts
                Sleep interval
        }
        return -1
    }

    ; SetHotkGroupInd(hwnd) {
    ;     if This.ThumbHwnd_EvEHwnd.Has(hwnd) {
    ;         hwnd := WinExist(This.ThumbHwnd_EvEHwnd[hwnd])
    ;         title := WinGetTitle("Ahk_id " Hwnd)
    ;     }
    ;     else
    ;         return
        
    ;     for groupId, arr in This.HotkGroups {
    ;         for i, charName in arr {
    ;             if charName != title
    ;                 continue

    ;             This.HotkGroupsInds[groupId] := i
    ;             break
    ;         }
    ;     }
    ; }

    hitThis() {
        if not This.ThisThat
            return

        static counter := 0
        static hit_score := 0
        static last_tick := A_TickCount
        tick := A_TickCount
        delay := 5 * This.GroupsHoldDelay

        if tick - last_tick > delay
            hit_score := 0

        hit_score++
        counter++

        x := A_ScreenWidth
        y := A_ScreenHeight

        CoordMode "ToolTip", "Screen"
        ToolTip("Combo " hit_score "`nTotal " counter, x, y)
        SetTimer(() => ToolTip(), -(delay))

        last_tick := tick
    }


    ; Checks for OldTitle = title in all login screen windows
    ; returns hwnd if found
    hasMathcingOldTitle(title) {
        if !This.PreserveHotkeysOnLogout
            return

        loginHWNDs := WinGetList("EVE ahk_exe exefile.exe")

        for hwnd in loginHWNDs {
            if This.ThumbWindows.HasProp(hwnd) && This.ThumbWindows.%hwnd%["Window"].OldTitle = title
                return hwnd
        }
        return
    }

    ; Cycle windows on Character selection screen
    Cycle_Login_Windows(*) {
        static tick, prevTick := 0
        tick := A_TickCount
        if tick - prevTick < This.GroupsHoldDelay
            return
        prevTick := tick

        LoginWins := []
        loginHWNDs := WinGetList("EVE ahk_exe exefile.exe")

        for hwnd in loginHWNDs {
            if This.ThumbWindows.HasProp(hwnd) && This.ThumbWindows.%hwnd%["Window"].OldTitle != "EVE" && This.PreserveHotkeysOnLogout
                continue

            PID := WinGetPID(hwnd)

            if This.LoginScreenCycleDirection
                CreationTime := This.GetProcessCreationTime(PID)[3]
            else
                CreationTime := This.GetProcessCreationTime(PID)[2]

            LoginWins.Push(Map("hwnd", hwnd, "CreationTime", CreationTime))
        }

        if !LoginWins.Length
            return

        if LoginWins.Length = 1 {
            if This.ignoredChars.Has(LoginWins[1]["hwnd"])
                return

            try
                This.ActivateEVEWindow(LoginWins[1]["hwnd"],,)
            return
        }

        if (LoginWins.Length > 1 ) {
            LoginWins := This.CustomSort(LoginWins, "CreationTime")
        }

        currentIndex := 0
        currentHWND := WinExist("A")
        for i, Win in LoginWins {
            if currentHWND == Win["hwnd"] {
                currentIndex := i
                break
            }
        }

        currentIndex += 1
        if currentIndex > LoginWins.Length
            currentIndex := 1

        Loop LoginWins.Length {
            if !This.ignoredChars.Has(LoginWins[currentIndex]["hwnd"])
                break

            currentIndex += 1
            if currentIndex > LoginWins.Length
                currentIndex := 1
        }

        if This.ignoredChars.Has(LoginWins[currentIndex]["hwnd"]) ; Last check if loop leaked window there
            return

        try
            This.ActivateEVEWindow(LoginWins[currentIndex]["hwnd"],,)
    }

    ; Close Active EVE Client 
    CloseActiveEVEWin(*) {
        if WinActive("ahk_exe exefile.exe")
		    WinClose("A")
    }

     ; To Check if atleast One Win stil Exist in the Array for the cycle groups hotkeys
    OnWinExist(Arr, *) {
        for index, Name in Arr {
            ; If ( WinExist("EVE - " Name " Ahk_Exe exefile.exe") && !WinActive("EVE-X-Preview - Settings") && !This.PreserveHotkeysOnLogout) {
            If ( WinExist(Name " Ahk_Exe exefile.exe") && !WinActive("EVE-X-Preview - Settings") ) {
                return true
            }
            else if This.hasMathcingOldTitle(Name) && !WinActive("EVE-X-Preview - Settings") {
                return true
            }
        }
        return false
    }

    OnWinActive(Arr, *) {        
        If (This.OnWinExist(Arr) && WinActive("Ahk_exe exefile.exe")) {
            return true
        }        
        return false
    }

    ;## Updates the Thumbnail in the GUI after Activation
    ;## Do not Update thumbnails from minimized windows or this will leed in no picture for the Thumbnail
    UpdateThumb_AfterActivation(event?, hwnd?) {
        if !IsSet(hwnd)
            return

        MinMax := -1
        try MinMax := WinGetMinMax("ahk_id " hwnd)

        if (This.ThumbWindows.HasProp(hwnd)) {
            if MinMax != -1 {
                This.Update_Thumb(false, This.ThumbWindows.%hwnd%["Window"].Hwnd)
            }
        }
    }

    ;This function updates the Thumbnails and hotkeys if the user switches Charakters in the character selection screen 
    EVENameChange(hwnd, title) {
        This.debugToolTip("Name changed for: " title)
        if (This.ThumbWindows.HasProp(hwnd)) {
            This.SetThumbnailText[hwnd] := title
            ; moves the Window to the saved positions if any stored, a bit of sleep is usfull to give the window time to move before creating the thumbnail
            This.RestoreClientPossitions(hwnd, title)

            if (This.ThumbnailPositions.Has(title)) {
                if This.AutoSaveThumbnailPositions ; Positions autosave
                    This.Save_ThumbnailPossitions
                This.EvEWindowDestroy(hwnd, title)
                This.EVE_WIN_Created(hwnd,title)
                rect := This.ThumbnailPositions[title]  
                This.ShowThumb(hwnd, "Hide")              
                This.ThumbMove( rect["x"],
                                rect["y"],
                                rect["width"],
                                rect["height"],
                                This.ThumbWindows.%hwnd% )

                This.BorderSize(This.ThumbWindows.%hwnd%["Window"].Hwnd, This.ThumbWindows.%hwnd%["Border"].Hwnd) 
                This.Update_Thumb(false)
                if (!This.HideThumbnailsOnLostFocus || WinActive(This.EVEExe)) {
                    for k, v in This.ThumbWindows.OwnProps()
                        This.ShowThumb(k, "Show")
                }
            }
            if This.dynamicGroupsEnabled && This.ignoredChars.Has(hwnd)
                This.toggleColorBorder(hwnd, title)
            This.BorderActive := 0
            This.RegisterHotkeys(title, hwnd)
        }
    }

    ; Toggle click through mode on all thumbnail windows
    Toggle_ClickThrough(*) {
        This.ClickThroughActive := !This.ClickThroughActive

        for hwnd, thumbObj in This.ThumbWindows.OwnProps() {
            This.ThumbClickThrough(thumbObj, 1)
        }

        SetTimer(This.Save_Settings_Delay_Timer, -200)

        ; Show a brief tooltip indicating the state
        ToolTip(This.ClickThroughActive ? "Click Through Thumbnails - ON" : "Click Through Thumbnails - OFF")
        SetTimer () => ToolTip(), -1500
    }

    ThumbClickThrough(thumbObj, forced?) { ; Thanks to @CJKondur
        if !(This.ClickThroughActive || IsSet(forced))
            return

        WS_EX_TRANSPARENT := 0x20
        WS_EX_LAYERED := 0x80000
        GWL_EXSTYLE := -20
        try {
            thumbHwnd := thumbObj["Window"].Hwnd
            exStyle := DllCall("GetWindowLongPtr", "Ptr", thumbHwnd, "Int", GWL_EXSTYLE, "Ptr")
            if (This.ClickThroughActive) {
                exStyle := exStyle | WS_EX_TRANSPARENT | WS_EX_LAYERED
            } else {
                exStyle := exStyle & ~WS_EX_TRANSPARENT
            }
            DllCall("SetWindowLongPtr", "Ptr", thumbHwnd, "Int", GWL_EXSTYLE, "Ptr", exStyle)
        }
    }

    ;#### Gets Called after receiveing a mesage from the Listeners
    ;#### Handels Window Border, Resize, Activation 
    _OnMessage(wparam, lparam, msg, hwnd) {
        If (This.ThumbHwnd_EvEHwnd.Has(hwnd)  ) {            
            ; Move the Window with right mouse button 
            If (msg == Main_Class.WM_RBUTTONDOWN) {
                while (GetKeyState("RButton")) {
                    if !(GetKeyState("LButton")) {
                        try {
                            This.Mouse_DragMove(wparam, lparam, msg, hwnd)
                            This.Window_Snap(hwnd, This.ThumbWindows)
                        }
                    }
                    else {
                        try
                            This.Mouse_ResizeThumb(wparam, lparam, msg, hwnd)

                    }
                }

                if This.AutoSaveThumbnailPositions ; Positions autosave
                    This.Save_ThumbnailPossitions
                
                return 0
            }

            ; Wparam -  9 Ctrl+Lclick
            ;           5 Shift+Lclick
            ;           13 Shift+ctrl+click
            Else If (msg == Main_Class.WM_LBUTTONDOWN) {
                ;Activates the EVE Window by clicking on the Thumbnail 
                if (wparam = 1) {
                    try { ; Probably fix for bug
                        if !(WinActive(This.ThumbHwnd_EvEHwnd[hwnd])) {
                            This.ActivateEVEWindow(hwnd)
                        }
                    }
                }
                ; Ctrl+Lbutton, Minimizes the Window on whose thumbnail the user clicks
                else if (wparam = 9) { 
                    ; Minimize
                    if (!GetKeyState("RButton"))
                        PostMessage 0x0112, 0xF020, , , This.ThumbHwnd_EvEHwnd[hwnd]
                } ; Shift+Click and feature enabled
                else if wparam = 5 && This.dynamicGroupsEnabled {
                    hwndEVE := This.ThumbHwnd_EvEHwnd[hwnd]
                    title := This.ThumbWindows.%hwndEVE%["Window"].Title
                    if !This.ignoredChars.Has(hwndEVE) {
                        This.ignoredChars[hwndEVE] := true
                        This.toggleColorBorder(hwndEVE, title)
                    }
                    else {
                        This.ignoredChars.Delete(hwndEVE)
                        This.toggleColorBorder(hwndEVE, title, false)
                        if WinActive("ahk_id " hwndEVE)
                            This.BorderActive := 0
                    }
                }
                return 0
            }   
        }
    }

    ; Creates a new thumbnail if a new window got created
    EVE_WIN_Created(Win_Hwnd, Win_Title) {
        This.debugToolTip("Creating thumbnail for " Win_Title)
        ; Moves the Window to the saved possition if any are stored 
        This.RestoreClientPossitions(Win_Hwnd, Win_Title)        
        
        ;Creates the Thumbnail and stores the EVE Hwnd in the array
        If This.ThumbWindows.HasProp(Win_Hwnd)
            return

        This.ThumbWindows.%Win_Hwnd% := This.Create_Thumbnail(Win_Hwnd, Win_Title)
        This.ThumbHwnd_EvEHwnd[This.ThumbWindows.%Win_Hwnd%["Window"].Hwnd] := Win_Hwnd
        This.ThumbWindows.%Win_Hwnd%["Window"].OldTitle := "EVE"

        ;if the User is in character selection screen
        if (This.ThumbWindows.%Win_Hwnd%["Window"].Title = "EVE") {
            This.SetThumbnailText[Win_Hwnd] := Win_Title
            This.ShiftThumbs(Win_Hwnd)
            return
        }  
        ;if the user loged in into a Character then move the Thumbnail to the right possition 
        else If (This.ThumbnailPositions.Has(Win_Title)) {
            This.SetThumbnailText[Win_Hwnd] := Win_Title
            rect := This.ThumbnailPositions[Win_Title]
            This.ThumbMove( rect["x"],
                            rect["y"],
                            rect["width"],
                            rect["height"],
                            This.ThumbWindows.%Win_Hwnd% )

            This.BorderSize(This.ThumbWindows.%Win_Hwnd%["Window"].Hwnd, This.ThumbWindows.%Win_Hwnd%["Border"].Hwnd)
            This.Update_Thumb(false)
            If ((This.HideThumbnailsOnLostFocus && WinActive(This.EVEExe)) || (!This.HideThumbnailsOnLostFocus)) {
                for k, v in This.ThumbWindows.OwnProps()
                    This.ShowThumb(k, "Show")
            }
        }
        else {
            This.SetThumbnailText[Win_Hwnd] := Win_Title  ; No saved position, still set the label
        }

        This.ThumbClickThrough(This.ThumbWindows.%Win_Hwnd%) ; If click through active, enable for new thumbnails
        This.RegisterNonEVEHotkeys()
        This.RegisterHotkeys(Win_Title, Win_Hwnd)
    }

    ; if ShiftThumbsForLoginScreen enabled we try to shift thumbnail using user settings
    ShiftThumbs(Win_Hwnd) {
        if !This.ShiftThumbsForLoginScreen || This.skipShiftThumbs
            return

        static nextPosX := This.ThumbnailStartLocation["x"]
        static nextPosY := This.ThumbnailStartLocation["y"]
        step_x := This.ShiftThumbHorizontalStep
        step_y := This.ShiftThumbVerticalStep

        if step_x = 0 || !IsInteger(step_x)
            step_x := This.ThumbnailStartLocation["width"]
        if step_y = 0 || !IsInteger(step_y)
            step_y := This.ThumbnailStartLocation["height"]
        
        switch This.ShiftThumbsDirection {
            case 2 || 7:
                step_y := -step_y
            case 3 || 6:
                step_x := -step_x
            case 4 || 8:
                step_x := -step_x
                step_y := -step_y
        }

        ; if ShiftThumbsCollisionCheck enabled checks position of all thumbnails and tries to avoid collision
        ; if we have collision check enabled we can reset nextPos so we start checking from beginning
        if This.ShiftThumbsCollisionCheck {
            nextPosX := This.ThumbnailStartLocation["x"]
            nextPosY := This.ThumbnailStartLocation["y"]
            Collision := This.CheckCollisions(nextPosX, nextPosY, This.ThumbnailStartLocation["width"], This.ThumbnailStartLocation["height"], This.ThumbWindows.%Win_Hwnd%["Window"].Hwnd)
        }
        ; if all login windows are closed we reset the position to start from beginning
        else if This.allLoginClosed {
            nextPosX := This.ThumbnailStartLocation["x"]
            nextPosY := This.ThumbnailStartLocation["y"]
            This.allLoginClosed := false
            Collision := 0
        }
        else
            Collision := 1

        while Collision {
            ; Horizontal -> Vertical
            if This.ShiftThumbsDirection <= 4 {
                nextPosX += step_x
                if nextPosX + This.ThumbnailStartLocation["width"] > A_ScreenWidth || nextPosX < 0 {
                    nextPosX := This.ThumbnailStartLocation["x"]
                    nextPosY += step_y
                    ; if end of screen reached, return to the default position
                    if nextPosY + This.ThumbnailStartLocation["height"] > A_ScreenHeight || nextPosY < 0 {
                        nextPosX := This.ThumbnailStartLocation["x"]
                        nextPosY := This.ThumbnailStartLocation["y"]
                        MsgBox("Thumbnail shifting reached end of screen! Returning to default position. Try change thumbnail default position, size, shift direction or step.")
                        This.skipShiftThumbs := true
                        break
                    }
                }
            }
            ; Vertical -> Horizontal
            else {
                nextPosY += step_y
                if nextPosY + This.ThumbnailStartLocation["height"] > A_ScreenHeight || nextPosY < 0 {
                    nextPosY := This.ThumbnailStartLocation["y"]
                    nextPosX += step_x
                    ; if end of screen reached, return to the default position
                    if nextPosX + This.ThumbnailStartLocation["width"] > A_ScreenWidth || nextPosX < 0 {
                        nextPosX := This.ThumbnailStartLocation["x"]
                        nextPosY := This.ThumbnailStartLocation["y"]
                        MsgBox("Thumbnail shifting reached end of screen! Returning to default position. Try change thumbnail default position, size, shift direction or step.")
                        This.skipShiftThumbs := true
                        break
                    }
                }
            }

            Collision := This.CheckCollisions(nextPosX, nextPosY, This.ThumbnailStartLocation["width"], This.ThumbnailStartLocation["height"], This.ThumbWindows.%Win_Hwnd%["Window"].Hwnd)
        }

        This.ThumbMove( nextPosX,
                        nextPosY,
                        This.ThumbnailStartLocation["width"],
                        This.ThumbnailStartLocation["height"],
                        This.ThumbWindows.%Win_Hwnd%)

    }

    ; Checks collisions for the new thumbnail position
    CheckCollisions(x1, y1, w1, h1, ThumbHwnd) {
        if !This.ShiftThumbsCollisionCheck
            return

        for EvEHwnd, ThumbObj in This.ThumbWindows.OwnProps() {
            for Name, Obj in ThumbObj {
                if (Name = "Window") {
                    WinGetPos(&x2, &y2, &w2, &h2, Obj.Hwnd)

                    if (x1 < x2 + w2) && (x2 < x1 + w1) && (y1 < y2 + h2) && (y2 < y1 + h1) && (Obj.Hwnd != ThumbHwnd) {
                        return 1
                    }
                }
            }
        }
        return
    }

    ;if a EVE Window got closed this destroyes the Thumbnail and frees the memory.
    EvEWindowDestroy(hwnd?, WinTitle?) {
        if IsSet(WinTitle)
            This.debugToolTip("Destroying thumbnail for " WinTitle)
        else if IsSet(hwnd)
            This.debugToolTip("Destroying thumbnail for hwnd " hwnd)
         else
            This.debugToolTip("Destroying thumbnail for unknown window")

        if (IsSet(hwnd) && This.ThumbWindows.HasProp(hwnd)) {
            for k, v in This.ThumbWindows.%hwnd% {
                if (k = "Thumbnail")
                    continue
                v.Destroy()
                ;This.ThumbWindows.%Win_Hwnd%.Delete()
            }
            This.ThumbWindows.DeleteProp(hwnd)
            if This.monitoringInitialized && IsSet(WinTitle) && This.monitoredChars.Has(WinTitle)
                This.stopLogMonitoring(WinTitle)
            This.DestroyThumbnailsToggle := 1
            Return
        }
        ;If a EVE Windows get destroyed 
        for Win_Hwnd,v in This.ThumbWindows.Clone().OwnProps() {
            if (!WinExist("Ahk_Id " Win_Hwnd)) {
                title := This.ThumbWindows.%Win_Hwnd%["Window"].Title
                for k, v in This.ThumbWindows.Clone().%Win_Hwnd% {
                    if (k = "Thumbnail")
                        continue
                    v.Destroy()
                }
                This.ThumbWindows.DeleteProp(Win_Hwnd)
                if This.monitoringInitialized && This.monitoredChars.Has(title)
                    This.stopLogMonitoring(title)
            }
        }
        This.DestroyThumbnailsToggle := 1
    }
    
    ActivateEVEWindow(hwnd?, ThisHotkey?, title?, direct?) {   
        ; If the user clicks the Thumbnail then hwnd stores the Thumbnail Hwnd. Here the Hwnd gets changed to the contiguous EVE window hwnd
        if (IsSet(hwnd) && This.ThumbHwnd_EvEHwnd.Has(hwnd)) {
            hwnd := WinExist(This.ThumbHwnd_EvEHwnd[hwnd])
            title := WinGetTitle("Ahk_id " hwnd)
        }
        ;if the user presses the Hotkey 
        Else if (IsSet(title)) {
            hwnd := WinExist(title " Ahk_exe exefile.exe")
        }
        if !IsSet(hwnd) || !hwnd {
            This.debugToolTip("ActivateEVEWindow: hwnd not found for " title)
            return
        }
        ;return when the user tries to bring a window to foreground which is already in foreground 
        if (WinActive("Ahk_id " hwnd))
            Return

        If (DllCall("IsIconic", "UInt", hwnd)) {
            if This.AlwaysMaximize || (This.TrackClientPossitions && This.ClientPossitions.Has(title) && This.ClientPossitions[title]["IsMaximized"]) {
                ; Maximize
                This.ShowWindowAsync(hwnd, 3)
            }
            else {
                ; Restore
                This.ShowWindowAsync(hwnd)          
            }
        }
        Else {    
            ; Use the virtual key to trigger the internal Hotkey.        
            This.ActivateHwnd := hwnd
            SendEvent("{Blind}{" Main_Class.virtualKey "}")            
        }

        ; if IsSet(direct) && direct { ; Might need later
        ; }

        This.hitThis()

        ;Sets the timer to minimize client if the user enable this.
        if (This.MinimizeInactiveClients) {
            This.wHwnd := hwnd
            SetTimer(This.timer, -This.MinimizeDelay)
        }
    }

    ;The function for the Internal Hotkey to bring a not minimized window in foreground 
    ActivateForgroundWindow(*) {
        ; 2 attempts for brining the window in foreground 
        try {
            if !(DllCall("SetForegroundWindow", "UInt", This.ActivateHwnd)) {
                DllCall("SetForegroundWindow", "UInt", This.ActivateHwnd)
            }

                ;If the user has selected to always maximize. this prevents wrong sized windows on heavy load.
            if (This.AlwaysMaximize && WinGetMinMax("ahk_id " This.ActivateHwnd) = 0) || ( This.TrackClientPossitions && This.ClientPossitions[WinGetTitle("Ahk_id " This.ActivateHwnd)]["IsMaximized"] && WinGetMinMax("ahk_id " This.ActivateHwnd) = 0 )
                This.ShowWindowAsync(This.ActivateHwnd, 3)
        }       
        Return 
    }

    ; Minimize All windows after Activting one with the exception of Titels in the DontMinimize Wintitels
    ; gets called by the timer to run async
    EVEMinimize() {
        for EveHwnd, GuiObj in This.ThumbWindows.OwnProps() {
            ThumbHwnd := GuiObj["Window"].Hwnd
            try
                WinTitle := WinGetTitle("Ahk_Id " EveHwnd)
            catch
                continue

            ; if (EveHwnd = This.wHwnd || Dont_Minimze_Enum(EveHwnd, WinTitle) || WinTitle == "EVE" || WinTitle = "")
            if (EveHwnd = This.wHwnd || Dont_Minimze_Enum(EveHwnd, WinTitle))
                continue
            else {
                ; Just to make sure its not minimizeing the active Window
                if !(EveHwnd = WinExist("A")) {
                    This.ShowWindowAsync(EveHwnd, 11)                    
                }
            }
        }
        ;to check which names are in the list that should not be minimized
        Dont_Minimze_Enum(hwnd, EVEwinTitle) {
            WinTitle := EVEwinTitle
            if !(WinTitle = "EVE") {
                for k in This.Dont_Minimize_Clients {
                    value := k
                    if value == WinTitle
                        return 1
                }
                return 0
            }
        }
    }

    ; Function t move the Thumbnails into the saved positions from the user
    ThumbMove(x := "", y := "", Width := "", Height := "", GuiObj := "") {
        if GuiObj = ""
            return
        for Names, Obj in GuiObj {
            if (Names = "Thumbnail")
                continue
            WinMove(x, y, Width, Height, Obj.Hwnd)
        }
    }

    ;Saves the possitions of all Windows and stores
    Client_Possitions() {
        IDs := WinGetList("Ahk_Exe " This.EVEExe)
        for k, v in IDs {
            Title := WinGetTitle("Ahk_id " v)
            if !(Title = "EVE") {
                ;If Minimzed then restore before saving the coords
                if (DllCall("IsIconic", "UInt", v)) {
                    This.ShowWindowAsync(v)
                    ;wait for getting Active for maximum of 2 seconds
                    if (WinWaitActive("Ahk_Id " v, , 2)) {
                        Sleep(200)
                        WinGetPos(&X, &Y, &Width, &Height, "Ahk_Id " v)
                        ;If the Window is Maximized
                        if (DllCall("IsZoomed", "UInt", v)) {
                            This.ClientPossitions[Title] := [X, Y, Width, Height, 1]
                        }
                        else {
                            This.ClientPossitions[Title] := [X, Y, Width, Height, 0]
                        }

                    }
                }
                ;If the Window is not Minimized
                else {
                    WinGetPos(&X, &Y, &Width, &Height, "Ahk_Id " v)
                    ;is the window Maximized?
                    if (DllCall("IsZoomed", "UInt", v)) {
                        This.ClientPossitions[Title] := [X, Y, Width, Height, 1]
                    }
                    else
                        This.ClientPossitions[Title] := [X, Y, Width, Height, 0]
                }
            }
        }
        SetTimer(This.Save_Settings_Delay_Timer, -200)
    }

    ;Restore the clients to the saved positions 
    RestoreClientPossitions(hwnd, title) {              
        if (This.TrackClientPossitions) {
            if ( This.TrackClientPossitions && This.ClientPossitions[title] ) {  
                if (DllCall("IsIconic", "UInt", hwnd) && This.ClientPossitions[title]["IsMaximized"] || DllCall("IsZoomed", "UInt", hwnd) && This.ClientPossitions[title]["IsMaximized"])  {
                    This.SetWindowPlacement(hwnd,This.ClientPossitions[title]["x"], This.ClientPossitions[title]["y"],
                    This.ClientPossitions[title]["width"], This.ClientPossitions[title]["height"], 9 )
                    This.ShowWindowAsync(hwnd, 3)
                    Return 
                }
                else if (DllCall("IsIconic", "UInt", hwnd) && !This.ClientPossitions[title]["IsMaximized"] || DllCall("IsZoomed", "UInt", hwnd) && !This.ClientPossitions[title]["IsMaximized"])  {
                    This.SetWindowPlacement(hwnd,This.ClientPossitions[title]["x"], This.ClientPossitions[title]["y"],
                    This.ClientPossitions[title]["width"], This.ClientPossitions[title]["height"], 9 )
                    This.ShowWindowAsync(hwnd, 4)
                    Return 
                }
                else if ( This.ClientPossitions[title]["IsMaximized"]) {
                    This.SetWindowPlacement(hwnd,This.ClientPossitions[title]["x"], This.ClientPossitions[title]["y"],
                    This.ClientPossitions[title]["width"], This.ClientPossitions[title]["height"] )
                    This.ShowWindowAsync(hwnd, 3)                    
                    Return 
                }    
                else if ( !This.ClientPossitions[title]["IsMaximized"]) {
                    This.SetWindowPlacement(hwnd,This.ClientPossitions[title]["x"], This.ClientPossitions[title]["y"],
                    This.ClientPossitions[title]["width"], This.ClientPossitions[title]["height"], 4 )
                    This.ShowWindowAsync(hwnd, 4) 
                    Return 
                }                  
            }
        }
    }

    CreateNonEVEAppsList() {
        This.NonEVEAppsList := []
        apps := This.NonEVEHotkeys
        for i, _ in apps["exe"]
            if !AppIsInArray(apps["exe"][i], apps["title"][i])
                This.NonEVEAppsList.Push(Map("exe", apps["exe"][i], "title", apps["title"][i]))
        
        for _, group in This.NonEVEGroups
            for i, exe in group["exe"]
                if !AppIsInArray(exe, group["title"][i])
                    This.NonEVEAppsList.Push(Map("exe", exe, "title", group["title"][i]))

        AppIsInArray(exe, title) {
            for app in This.NonEVEAppsList
                if app["exe"] == exe && app["title"] == title
                    return true
            return false
        }
    }

    UpdateActiveNonEVEApps() {
        This.ActiveNonEVEApps := [] ; List of hwnds
        for app in This.NonEVEAppsList {
            criteria := "ahk_exe " . app["exe"]
            if app["title"] != ""
                criteria := app["title"] . " " . criteria
            if WinExist(criteria) {
                hwnd := WinGetID(criteria)
                This.ActiveNonEVEApps.Push(hwnd)
            }
        }
    }

    RegisterNonEVEGroups() {
        This.NonEVEGroupsL := []
        This.NonEVEGroupsInds := []
        directions := Map()
        index := 1

        This.debugToolTip("Registering Non-EVE Groups")

        for _, group in This.NonEVEGroups {
            gr := group
            This.NonEVEGroupsL.Push(gr)
            This.NonEVEGroupsInds.Push(0)

            if gr.has("fkey") && gr["fkey"] != ""
                directions["Forward"] := gr["fkey"]
            if gr.has("bkey") && gr["bkey"] != ""
                directions["Backward"] := gr["bkey"]

            for direction in directions {
                HotIf ObjBindMethod(This, "AtLeastOneWinExist_", gr)
                if !This.SwitchLangOnErr {
                    try
                        Hotkey(directions[direction], ObjBindMethod(This, "CycleNonEVEGroups", index, direction), "P1")
                    catch ValueError as e
                        MsgBox(e.Message " --> " e.Extra " <-- in Non-EVE Applications - " This.LastUsedProfile " Non-EVE Hotkey Groups")
                }
                else
                    Hotkey(directions[direction], ObjBindMethod(This, "CycleNonEVEGroups", index, direction), "P1")
            }
            index += 1
            directions := Map()
        }
    }

    AtLeastOneWinExist_(group, *) {
        for i, exec in group["exe"] {
            if !WinExist("ahk_exe " exec) || WinActive("EVE-X-Preview - Settings")
                continue
            if group["title"][i] != ""
                return WinExist(group["title"][i])
        
            return true
        }
        return false
    }

    CycleNonEVEGroups(groupIndex, direction, *) {
        static tick, prevTick := 0
        tick := A_TickCount
        if tick - prevTick < This.GroupsHoldDelay
            return
        prevTick := tick

        group := This.NonEVEGroupsL[groupIndex]
        index := This.NonEVEGroupsInds[groupIndex]
        length := group["exe"].Length

        if !This.KeepGroupsPositions {
            try {
                exec := WinGetProcessName("A")
                title := WinGetTitle("A")
                index := IsActiveWinInGroup(exec, title, group)
            }
            catch
                index := 0
        }

        index := DirectionHandler(direction, index, length)
    
        if !This.AtLeastOneWinExist_(group)
            return

        loop length {
            hwnd := This.OnWinExist_(group["exe"][index], group["title"][index])
            if hwnd && !This.ignoredChars.Has(hwnd)
                break

            index := DirectionHandler(direction, index, length)
        }

        if This.ignoredChars.Has(hwnd) ; Last check if loop leaked window there
            return

        try
            This.ActivateNonEVE(group["exe"][index], group["title"][index])

        This.NonEVEGroupsInds[groupIndex] := index

        ; Get new index by specified direction
        DirectionHandler(direction, index, length) {
            if (direction == "Forward") {
                index += 1
                if index > length
                    index := 1
            }
            else if (direction == "Backward") {
                index -= 1
                if (index <= 0)
                    index := length
            }
            return index
        }

        IsActiveWinInGroup(exe, title, group) {
            for i, exec in group["exe"] {
                if exec = exe && ((group["title"][i] != "" && group["title"][i] = title  || group["title"][i])== "")
                    return i
            }
            return 0
        }
    }

    RegisterNonEVEHotkeys() {
        This.debugToolTip("Registering Non-EVE Hotkeys")
        apps := This.NonEVEHotkeys
        for i, _ in apps["exe"] {
            if !apps["hotkey"].Has(i) || apps["hotkey"][i] = ""
                continue
            HotIf ObjBindMethod(This, "OnWinExist_", apps["exe"][i], apps["title"][i])
            if !This.SwitchLangOnErr {
                try
                    Hotkey(apps["hotkey"][i], ObjBindMethod(This, "ActivateNonEVE", apps["exe"][i], apps["title"][i], 1), "P1")
                catch ValueError as e
                    MsgBox(e.Message " --> " e.Extra " <-- in Non-EVE Applications - " This.LastUsedProfile " - Non-EVE Hotkeys")
            }
            else
                Hotkey(apps["hotkey"][i], ObjBindMethod(This, "ActivateNonEVE", apps["exe"][i], apps["title"][i], 1), "P1")
        }
    }

    OnWinExist_(exe, title, *) {
        if !WinExist("ahk_exe " exe) || WinActive("EVE-X-Preview - Settings")
            return false
    
        if title != ""
            return WinExist(title " ahk_exe " exe)
    
        return WinExist("ahk_exe " exe)
    }

    ActivateNonEVE(exe, title, direct?, *) {
        criteria := "ahk_exe " exe
        if title != ""
            criteria := title . " " . criteria
    
        ; if IsSet(direct) && direct

        hwnd := WinExist(criteria)
        if !hwnd || WinActive("Ahk_id " hwnd)
            return

        if (DllCall("IsIconic", "UInt", hwnd)) {
            This.ShowWindowAsync(hwnd) ; Restore
        }
        else { ; Use the virtual key to trigger the internal Hotkey.
            This.ActivateHwnd := hwnd
            SendEvent("{Blind}{" Main_Class.virtualKey "}")
        }

        This.hitThis()

        ;Sets the timer to minimize client if the user enable this.
        if This.MinimizeInactiveClients {
            This.wHwnd := hwnd
            SetTimer(This.timer, -This.MinimizeDelay)
        }
    }
    
    ;*WinApi Functions
    ;Gets the normal possition from the Windows. Not to use for Maximized Windows 
    GetWindowPlacement(hwnd) {
        DllCall("User32.dll\GetWindowPlacement", "Ptr", hwnd, "Ptr", WP := Buffer(44))
        Lo := NumGet(WP, 28, "Int")        ; X coordinate of the upper-left corner of the window in its original restored state
        To := NumGet(WP, 32, "Int")        ; Y coordinate of the upper-left corner of the window in its original restored state
        Wo := NumGet(WP, 36, "Int") - Lo   ; Width of the window in its original restored state
        Ho := NumGet(WP, 40, "Int") - To   ; Height of the window in its original restored state

        CMD := NumGet(WP, 8, "Int") ; ShowCMD
        flags := NumGet(WP, 4, "Int")  ; flags
        MinX := NumGet(WP, 12, "Int")
        MinY := NumGet(WP, 16, "Int")
        MaxX := NumGet(WP, 20, "Int")
        MaxY := NumGet(WP, 24, "Int")
        WP := ""

        return { X: Lo, Y: to, W: Wo, H: Ho , cmd: CMD, flags: flags, MinX: MinX, MinY: MinY, MaxX: MaxX, MaxY: MaxY }
    }

    ;Moves the window to the given possition immediately
    SetWindowPlacement(hwnd:="", X:="", Y:="", W:="", H:="", action := 9) {
        ;hwnd := hwnd = "" ? WinExist("A") : hwnd
        DllCall("User32.dll\GetWindowPlacement", "Ptr", hwnd, "Ptr", WP := Buffer(44))
        Lo := NumGet(WP, 28, "Int")        ; X coordinate of the upper-left corner of the window in its original restored state
        To := NumGet(WP, 32, "Int")        ; Y coordinate of the upper-left corner of the window in its original restored state
        Wo := NumGet(WP, 36, "Int") - Lo   ; Width of the window in its original restored state
        Ho := NumGet(WP, 40, "Int") - To   ; Height of the window in its original restored state
        L := X = "" ? Lo : X               ; X coordinate of the upper-left corner of the window in its new restored state
        T := Y = "" ? To : Y               ; Y coordinate of the upper-left corner of the window in its new restored state
        R := L + (W = "" ? Wo : W)         ; X coordinate of the bottom-right corner of the window in its new restored state
        B := T + (H = "" ? Ho : H)         ; Y coordinate of the bottom-right corner of the window in its new restored state

        NumPut("UInt",action,WP,8)
        NumPut("UInt",L,WP,28)
        NumPut("UInt",T,WP,32)
        NumPut("UInt",R,WP,36)
        NumPut("UInt",B,WP,40)
        
        Return DllCall("User32.dll\SetWindowPlacement", "Ptr", hwnd, "Ptr", WP)
    }


    ShowWindowAsync(hWnd, nCmdShow := 9) {
        DllCall("ShowWindowAsync", "UInt", hWnd, "UInt", nCmdShow)
    }
    GetActiveWindow() {
        Return DllCall("GetActiveWindow", "Ptr")
    }
    SetActiveWindow(hWnd) {
        Return DllCall("SetActiveWindow", "Ptr", hWnd)
    }
    SetFocus(hWnd) {
        Return DllCall("SetFocus", "Ptr", hWnd)
    }
    SetWindowPos(hWnd, x, y, w, h, hWndInsertAfter := 0, uFlags := 0x0020) {
        ; SWP_FRAMECHANGED 0x0020
        ; SWP_SHOWWINDOW 0x40
        Return DllCall("SetWindowPos", "Ptr", hWnd, "Ptr", hWndInsertAfter, "Int", x, "Int", y, "Int", w, "Int", h, "UInt", uFlags)
    }

    ;removes "EVE" from the Titel and leaves only the Character names
    CleanTitle(title) { ; More optimized than regex
        len := StrLen(title)
        if len >= 6 && SubStr(title, 1, 6) == "EVE - "
            return SubStr(title, 7)
        if title == "EVE"
            return ""
        return title
    }

    ; adds "EVE" to the titel
    AntiCleanTitle(title) {
        if title = "" || title = "EVE"
            return "EVE"
        len := StrLen(title)
        if len >= 6 && SubStr(title, 1, 6) == "EVE - "
            return title
        return "EVE - " title
    }

    SaveJsonToFile() {
        FileDelete("EVE-X-Preview.json")
        FileAppend(JSON.Dump(This._JSON, , "    "), "EVE-X-Preview.json")
    }

    ; Thanks to SKAN
    GetProcessCreationTime(PID) {
        Local  hProcess, T1601 := 0,  ExitCode := 0
            ,  CT := 0,  XT := 0,  KT := 0,  UT := 0           ;  PROCESS_QUERY_LIMITED_INFORMATION := 0x1000

        If ! ( hProcess := DllCall("Kernel32\OpenProcess", "uint",0x1000, "uint",0, "uint",PID, "ptr") )
            Return [0, 0, 0, 0, 0]

        DllCall("Kernel32\GetSystemTimeAsFileTime", "int64p",&T1601)
        , DllCall("Kernel32\GetProcessTimes", "ptr",hProcess, "int64p",&CT, "int64p",&XT, "int64p",&KT, "int64p",&UT)
        , DllCall("Kernel32\GetExitCodeProcess", "ptr",hProcess, "ptrp",&ExitCode)
        , DllCall("Kernel32\CloseHandle", "ptr",hProcess)          

        Return [ Round((KT / 10000000) + (UT / 10000000), 7)  ;  CPU Time (in seconds)
            , Round((T1601 - CT) / 10000000, 7 )           ;  Running  time: Seconds elapsed since creation time
            , Round(CT / 10000000, 7)                      ;  Creation time: Seconds elapsed since 1-Jan-1601 (UTC)
            , Round(XT / 10000000, 7)                      ;  Exit time:     Seconds elapsed since 1-Jan-1601 (UTC)
            , ExitCode ]                                   ;  will be 259 (STILL_ACTIVE) for running process
    }

    ; Bubble sort my beloved
    CustomSort(arr, sortBy, ascending := true) {
        n := arr.Length
        if (n < 2)
            return arr
    
        loop n - 1 {
            swapped := false
            for j, _ in arr {
                if (j >= n)
                    break
                if (ascending) {
                    if (arr[j][sortBy] > arr[j + 1][sortBy]) {
                        tmp := arr[j]
                        arr[j] := arr[j + 1]
                        arr[j + 1] := tmp
                        swapped := true
                    }
                } else {
                    if (arr[j][sortBy] < arr[j + 1][sortBy]) {
                        tmp := arr[j]
                        arr[j] := arr[j + 1]
                        arr[j + 1] := tmp
                        swapped := true
                    }
                }
            }
            if !swapped
                break
        }
        return arr
    }

    QuickSort(arr, comp, low := 1, high := "") {
        if (high = "")
            high := arr.Length
        if (low >= high)
            return

        mid := Floor((low + high)/2)
        pivot := arr[mid]

        i := low
        j := high

        while (i <= j)
        {
            while (comp(arr[i], pivot) < 0)
                i++
            while (comp(arr[j], pivot) > 0)
                j--
            if (i <= j)
            {
                tmp := arr[i]
                arr[i] := arr[j]
                arr[j] := tmp
                i++
                j--
            }
        }

        if (low < j)
            This.QuickSort(arr, comp, low, j)
        if (i < high)
            This.QuickSort(arr, comp, i, high)
    }

    DontCloseWIn(WinTitle) {
        if !(WinTitle = "EVE") {
            for k in This.DontCloseClients {
                value := k
                if value == WinTitle
                    return 1
            }
        }
        else if (WinTitle = "EVE" && This.DontCloseOnLoginScreen) {
            return 1
        }
        return 0
    }

    getFilesList() {
        if !DirExist(This.gameLogsDirectory) {
            MsgBox "Game logs folder not found!`nUsually it's in USERNAME\Documents\EVE\logs\Gamelogs."
            This.gameLogsDirectory := DirSelect() ; Select directory
            SetTimer(This.Save_Settings_Delay_Timer, -200)
            if This.gameLogsDirectory = "" ; Cancelled
                return
        }

        ; Optimized files count check
        static filesCount := 0
        static oldFilesList := []
        newFilesCount := 0

        files := []
        Loop Files, This.gameLogsDirectory "\*.*" {
            files.Push({name: A_LoopFileName, time: A_LoopFileTimeModified})
            newFilesCount++
        }
        if newFilesCount = filesCount
            return oldFilesList
        else
            filesCount := newFilesCount

        ; comparator: return <0 if a < b, 0 if equal, >0 if a > b
        ; for descending (newest first) we invert the usual order
        comp := (a, b) => (a.time > b.time) ? -1 : (a.time < b.time) ? 1 : 0

        ; in-place quicksort
        This.QuickSort(files, comp)
        
        ; build newline string or process in order
        fileList := []
        for file in files {
            fileList.Push(This.gameLogsDirectory "\" file.name)
        }
        oldFilesList := fileList
        return fileList
    }

    gameLogsMonitoring() {
        ; filename has structure
        ; YYYYMMDD_XXXXXX_CCCCCCCCCC.txt where
        ; First is date like: YYYYMMDD
        ; Second is some session? ID contains 6 digits: XXXXXX
        ; Third is character ID 10 digits: CCCCCCCCCC
        ; If character didn't logged in there is structure like:
        ; YYYYMMDD_XXXXXX.txt with same first 2 line explained before

        if This.gameLogsDirectory = "" {
            This.gameLogsDirectory := "C:\Users\" A_UserName "\Documents\EVE\logs\Gamelogs"
            SetTimer(This.Save_Settings_Delay_Timer, -200)
        }

        This.monitoredChars := Map()
        This.waitingMonitoringChars := Map()
        This.shootingChars := Map()
        This.flashMethod := Map()
        This.eventMethods := Map()

        if !This.gameLogsMonitoringEnabled || (!This.flashBorderEnabled && !This.showEventText) ; When events displaying disabled, don't initiate monitoring
            return

        ; Thanks to @CJKondur to having this list
        ; Comprehensive list of all EVE Online NPC naming prefixes.
        ; Used by PVE mode to filter NPC damage from attack alerts.
        ; CCP blocks players from using faction names in character creation.
        This.generalNPCPatterns := [
            ; --- Pirate Factions ---
            "Guristas",
            "Sansha", "Sansha's",
            "Blood Raider",
            "Angel Cartel",
            "Serpentis",
            "Mordu's Legion", "Mordu's",
            ; --- Pirate Named Variants (Faction-specific hull prefixes) ---
            ; Angel Cartel
            "Gistii", "Gistum", "Gistior", "Gistatis", "Gist",
            ; Blood Raiders
            "Corpii", "Corpum", "Corpior", "Corpatis", "Corpus",
            ; Guristas
            "Pithi", "Pithum", "Pithior", "Pithatis", "Pith",
            ; Sansha's Nation
            "Centii", "Centum", "Centior", "Centatis", "Centus",
            ; Serpentis
            "Coreli", "Corelum", "Corelior", "Corelatis", "Core ",
            ; --- Empire Factions ---
            "Amarr Navy", "Amarr",
            "Caldari Navy", "Caldari",
            "Gallente Navy", "Gallente",
            "Minmatar Fleet", "Minmatar",
            "Imperial Navy",
            "State",
            "Federation Navy", "Federation",
            "Republic Fleet", "Republic",
            "CONCORD",
            ; --- Rogue Drones ---
            "Rogue",
            ; Drone hull suffixes used as prefixes in some contexts
            "Infester", "Render", "Raider", "Strain",
            "Decimator", "Sunder", "Nuker",
            "Predator", "Hunter", "Destructor",
            ; --- Sleepers ---
            "Sleepless", "Awakened", "Emergent",
            ; --- Triglavian ---
            "Starving", "Renewing", "Blinding",
            "Harrowing", "Ghosting", "Tangling",
            "Raznaborg", "Vedmak", "Vila",
            "Zorya",
            ; --- Drifter ---
            "Artemis", "Apollo", "Hikanta", "Drifter",
            "Tyrannos",
            ; --- EDENCOM ---
            "EDENCOM",
            ; --- Triglavian Invasion NPCs ---
            "Anchoring", "Liminal",
            ; --- Sentry Guns & Structures ---
            "Sentry", "Sentry Gun",
            "Territorial",
            ; --- FOB / Diamond NPCs ---
            "Forward Operating",
            ; --- Mercenary NPCs ---
            "Mercenary",
            ; --- Thukker ---
            "Thukker",
            ; --- Sisters of EVE ---
            "Sisters of",
            ; --- ORE ---
            "ORE",
            ; --- Faction Warfare NPCs ---
            "Navy",
            ; NPC name suffixes (for rogue drones: "Infester Alvi", etc.)
            ; Drone name suffixes (these appear as full names)
            "Alvi", 
            "Alvus", 
            "Alvatis", 
            "Alvior"
        ]

        factionNPCs := [ ; NPCs to trigger engagedWithFactionBSNPC event
            "Domination Cherubim",
            "Domination Commander",
            "Domination General",
            "Domination Malakim",
            "Domination Nephilim",
            "Domination Saint",
            "Domination Seraphim",
            "Domination Throne",
            "Domination War General",
            "Domination Warlord",
            "Dark Blood Apostle",
            "Dark Blood Archbishop",
            "Dark Blood Archon",
            "Dark Blood Cardinal",
            "Dark Blood Harbinger",
            "Dark Blood Monsignor",
            "Dark Blood Oracle",
            "Dark Blood Patriarch",
            "Dark Blood Pope",
            "Dark Blood Prophet",
            "Dread Guristas Conquistador",
            "Dread Guristas Destroyer",
            "Dread Guristas Dismantler",
            "Dread Guristas Eliminator",
            "Dread Guristas Eradicator",
            "Dread Guristas Exterminator",
            "Dread Guristas Extinguisher",
            "Dread Guristas Massacrer",
            "Dread Guristas Obliterator",
            "Dread Guristas Usurper",
            "Sentient Alvus Controller",
            "Sentient Alvus Creator",
            "Sentient Alvus Queen",
            "Sentient Alvus Ruler",
            "Sentient Domination Alvus",
            "Sentient Matriarch Alvus",
            "Sentient Patriarch Alvus",
            "Sentient Spearhead Alvus",
            "Sentient Supreme Alvus Parasite",
            "Sentient Swarm Preserver Alvus",
            "True Sansha's Beast Lord",
            "True Sansha's Dark Lord",
            "True Sansha's Dread Lord",
            "True Sansha's Lord",
            "True Sansha's Mutant Lord",
            "True Sansha's Overlord",
            "True Sansha's Plague Lord",
            "True Sansha's Savage Lord",
            "True Sansha's Slave Lord",
            "True Sansha's Tyrant",
            "Shadow Serpentis Admiral",
            "Shadow Serpentis Baron",
            "Shadow Serpentis Commodore",
            "Shadow Serpentis Flotilla Admiral",
            "Shadow Serpentis Grand Admiral",
            "Shadow Serpentis High Admiral",
            "Shadow Serpentis Lord Admiral",
            "Shadow Serpentis Port Admiral",
            "Shadow Serpentis Rear Admiral",
            "Shadow Serpentis Vice Admiral"
        ]
        This.factionNPCs := Map()
        for npc in factionNPCs
            This.factionNPCs[npc] := true

        officerNPCs := [ ; NPCs to trigger engagedWithOfficerNPC event
            "Gotan Kreiss",
            "Hakim Stormare",
            "Mizuro Cybon",
            "Tobias Kruzhor",
            "Ahremen Arkah",
            "Draclira Merlonne",
            "Raysere Giant",
            "Tairei Namazoth",
            "Estamel Tharchon",
            "Kaikka Peunato",
            "Thon Eney",
            "Vepas Minimala",
            "Unit D-34343",
            "Unit F-435454",
            "Unit P-343554",
            "Unit W-634",
            "Brokara Ryver",
            "Chelm Soran",
            "Selynne Mardakar",
            "Vizan Ankonin",
            "Brynn Jerdola",
            "Cormack Vaaja",
            "Setele Schellan",
            "Tuvan Orth"
        ]
        This.officerNPCs := Map()
        for npc in officerNPCs
            This.officerNPCs[npc] := true

        capitalNPCs := [ ; NPCs to trigger engagedWithCapitalNPC event
            "Domination Titan",
            "Dark Blood Titan",
            "Shadow Serpentis Titan",
            "Angel Dreadnought",
            "Domination Dreadnought",
            "Blood Dreadnought",
            "Dark Blood Dreadnought",
            "Dread Guristas Dreadnought",
            "Guristas Dreadnought",
            "Sansha's Dreadnought",
            "True Sansha's Dreadnought",
            "Serpentis Dreadnought",
            "Shadow Serpentis Dreadnought",
            "Infested Carrier",
            "Sentient Infested Carrier",
            "Sentient Infested Supercarrier",
            "True Sansha's Supercarrier",
            "Dread Guristas Titan"
        ]
        This.capitalNPCs := Map()
        for npc in capitalNPCs
            This.capitalNPCs[npc] := true

        eventPatterns := Map(
            ; "underAttackByPlayer", Map("pattern", "", "needRegex", 1, "checkNPC", 1),
            ; "underAttackByNPC", Map("pattern", "", "needRegex", 0, "checkNPC", 1),
            ; "engagedWithFactionBSNPC", Map("pattern", "", "needRegex", 0, "checkNPC", 1),
            ; "engagedWithOfficerNPC", Map("pattern", "", "needRegex", 0, "checkNPC", 1),
            ; "engagedWithCapitalNPC", Map("pattern", "", "needRegex", 0, "checkNPC", 1),
            "warpDisrupted", Map("pattern", "you!", "needRegex", 0),
            "fleetInvited", Map("pattern", "wants you to join their fleet", "needRegex", 0),
            "fleetWarped", Map("pattern", "Following .+? in warp", "needRegex", 1), ; Just following don't works because "following reasons"
            "fleetRegrouped", Map("pattern", "Regrouping", "needRegex", 0),
            "decloaked", Map("pattern", "cloak deactivates", "needRegex", 0),
            "convoRequest", Map("pattern", "is inviting you to a conversation", "needRegex", 0),
            "conduited", Map("pattern", "A Conduit Field activated by", "needRegex", 0),
            "gateJumped", Map("pattern", "(Jumping from|jumping to)", "needRegex", 1),
            "crystalBroke",Map("pattern", "deactivates due to the destruction", "needRegex", 0),
            "miningStopped", Map("pattern", "pale shadow of its former glory", "needRegex", 0),
            "miningBayIsFull", Map("pattern", "has completed operations", "needRegex", 0),
            "stoppedShooting", Map("pattern", "123456789abcdefg", "needRegex", 0) ; Placeholder pattern
        )

        This.checkFactionNPCs := false
        This.checkOfficerNPCs := false
        This.checkCapitalNPCs := false
        This.checkGeneralNPCs := false
        This.anyNPCEngagmentEnabled := false
        This.playerEngagmentEnabled := false

        This.checkedNPCs := Map()
        This.enabledMonitoredEvents := Map() ; To check only enabled events
        for event, v in This.monitoredEvents {
            if v["enabled"] {
                if event = "engagedWithFactionBSNPC"
                    This.checkFactionNPCs := true
                else if event = "engagedWithOfficerNPC"
                    This.checkOfficerNPCs := true
                else if event = "engagedWithCapitalNPC"
                    This.checkCapitalNPCs := true
                else if event = "underAttackByNPC" {
                    This.checkGeneralNPCs := true
                    This.anyNPCEngagmentEnabled := true
                } else if event = "underAttackByPlayer" {
                    This.checkGeneralNPCs := true
                    This.playerEngagmentEnabled := true
                } else {
                    This.enabledMonitoredEvents[event] := This.monitoredEvents[event].Clone()
                    This.enabledMonitoredEvents[event]["pattern"] := eventPatterns[event]["pattern"]
                    This.enabledMonitoredEvents[event]["needRegex"] := eventPatterns[event]["needRegex"]
                    ; This.enabledMonitoredEvents[event]["checkNPC"] := eventPatterns[event]["checkNPC"]
                }
            }
        }
        if !This.enabledMonitoredEvents.Count ; If dont enabled
            return

        This.checkNPCs := false
        if This.checkFactionNPCs || This.checkOfficerNPCs || This.checkCapitalNPCs || This.checkGeneralNPCs
            This.checkNPCs := true

        activeCharsToMonitor := Map() ; Must be active/logged in and in list of monitored
        if This.monitorOnlySelectedChars {
            for char in This.charsToMonitor {
                if WinExist(char) {
                    for charName, id in This.charsIds {
                        if char = charName {
                            activeCharsToMonitor[char] := id
                            break
                        }
                    }
                    activeCharsToMonitor[char] := 0
                }
            }
        }
        else {
            try
                WinList := WinGetList(This.EVEExe)
            catch
                WinList := []
    
            if WinList.Length {
                for index, hwnd in WinList {
                    title := WinGetTitle(hwnd)
                    if title = "EVE" ; In character selection screen
                        continue
    
                    for charName, id in This.charsIds {
                        if title = charName {
                            activeCharsToMonitor[title] := id
                            break
                        }
                    }
                    activeCharsToMonitor[title] := 0
                }
            }
        }

        timer := -5000 ; Dynamic timer so we don't lag the PC if there is a lot of chars
        for charName, charId in activeCharsToMonitor {
            This.waitingMonitoringChars[charName] := ObjBindMethod(This, "startLogMonitoring", charName, charId)
            SetTimer(This.waitingMonitoringChars[charName], timer) ; Wait 5s to initialize file if program started in the middle of login
            timer -= 1000 ; Adding offset for each char to avoid lags if there is a lot of them
        }

        This.debugToolTip("Initialized log monitoring for chars:`n" This.StrJoin("`n", activeCharsToMonitor) "`n`nWaiting for chars to update logs...")
        This.monitoringInitialized := 1
        This.monitorMethod := ObjBindMethod(This, "monitorAllChars")
        SetTimer(This.monitorMethod, This.monitoringInterval) ; Poll every specified by user time
    }

    getCharNameFromFile(fileName) {
        file := FileOpen(fileName, "r", "UTF-8")        
        if !file
            return
        Loop 3 {
            line := file.ReadLine()
            if A_Index = 3
                charName := Trim(StrReplace(line, "Listener: "))
        }
        file.Close()
        return This.AntiCleanTitle(charName)
    }

    startLogMonitoring(charName, charId) {
        if !This.gameLogsMonitoringEnabled
            return

        filesListSorted := This.getFilesList()
        if !filesListSorted
            return

        foundFile := 0
        for fileName in filesListSorted { ; Finding char names in headers

            fileNameData := StrSplit(fileName, "_")
            if !fileNameData.Has(3) ; in character selection screen
                continue
            fileCharId := StrReplace(fileNameData[3], ".txt") ; removing .txt in the end

            if charId != fileCharId {
                fileCharName := This.getCharNameFromFile(fileName)
                if !fileCharName || fileCharName != charName
                    continue
            }

            if charId = 0 {
                This.charsIds[charName] := fileCharId ; adding to config to speedup process
                SetTimer(This.Save_Settings_Delay_Timer, -200)
            }
            
            foundFile := fileName
            break
        }

        if !foundFile
            return

        file := FileOpen(foundFile, "r", "UTF-8")
        if !file
            return
        size := file.Length
        file.Seek(-1, 2)
        file.ReadLine()

        This.monitoredChars[charName] := Map("fileName", foundFile, "file", file, "size", size, "launchTime", A_TickCount, "fileUpdated", false)

        if This.waitingMonitoringChars.Has(charName)
            This.waitingMonitoringChars.Delete(charName)

        This.debugToolTip("Started monitoring: " . charName . "`n`nAll monitored chars:`n" . This.StrJoin("`n", This.monitoredChars))
    }

    stopLogMonitoring(charName) {
        if !This.monitoredChars.Has(charName)
            return
        This.monitoredChars[charName]["file"].Close() ; Closing file
        This.monitoredChars.Delete(charName)

        This.debugToolTip("Stopped monitoring " charName)
    }

    monitorAllChars() {
        for charName in This.monitoredChars
            This.monitorChanges(charName)
    }

    monitorChanges(charName) {
        fileObj := This.monitoredChars[charName]["file"]
        if !This.monitoredChars.Has(charName) || !IsObject(fileObj) || !WinExist(charName " ahk_exe exefile.exe") {
            This.stopLogMonitoring(charName)
            return
        }

        tick := A_TickCount
        size := This.monitoredChars[charName]["file"].Length
        if size = This.monitoredChars[charName]["size"] {
            if !This.monitoredChars[charName]["fileUpdated"] && tick - This.monitoredChars[charName]["launchTime"] > 30000 { ; if there is no changes for some time, try find new file
                This.stopLogMonitoring(charName)
            }
            return
        }
        This.monitoredChars[charName]["size"] := size
        This.monitoredChars[charName]["fileUpdated"] := True

        if (This.supressForFocused || This.HideThumbForActiveWin) && WinActive(charName " ahk_exe exefile.exe") { ; Skipping event check if thumb active and supressForFocused or HideThumbForActiveWin enabled
            while !This.monitoredChars[charName]["file"].AtEOF
                This.monitoredChars[charName]["file"].ReadLine()
            return
        }

        while !This.monitoredChars[charName]["file"].AtEOF {
            line := This.monitoredChars[charName]["file"].ReadLine()

            if line = ""
                continue

            This.monitoredChars[charName]["event"] := ""
            for e, v in This.enabledMonitoredEvents {
                if This.monitoredChars[charName]["event"] != ""
                    break
                if e = "stoppedShooting" && (RegExMatch(line, "\s(\d+):(\d+):(\d+)\s\].+?<color=0xff00ffff><b>\d+</b> <color=0x77ffffff><font size=\d+>to</font> <b><color=0xffffffff>", &m) || RegExMatch(line, "\s(\d+):(\d+):(\d+)\s\].+?Your .+? misses .+? completely", &m)) {
                    if This.shootingChars.Has(charName) {
                        SetTimer(This.shootingChars[charName], 0)
                        This.shootingChars.Delete(charName)
                    }

                    This.shootingChars[charName] := ObjBindMethod(This, "handleEventActivation", charName, 1)
                    SetTimer(This.shootingChars[charName], -This.shootingInterval)

                    This.monitoredChars[charName]["event"] := e
                }
                else if (!v["needRegex"] && InStr(line, v["pattern"])) || (v["needRegex"] && RegExMatch(line, v["pattern"])) {
                    This.monitoredChars[charName]["event"] := e
                }
            }
            This.processNPCCheck(charName, line) ; checking npcs if any event enabled

            This.handleEventActivation(charName)
        }
    }

    handleEventActivation(charName, shooting := 0) {
        if !This.monitoredChars.Has(charName) || This.HideThumbnails || (This.HideThumbForActiveWin && WinActive(charName " ahk_exe exefile.exe")) || !WinExist(charName " ahk_exe exefile.exe")
            return
        if !shooting && (This.monitoredChars[charName]["event"] = "" || This.monitoredChars[charName]["event"] = "stoppedShooting")
            return
        else if shooting {
            event := "stoppedShooting"
            if This.shootingChars.Has(charName)
                This.shootingChars.Delete(charName)
        }
        else
            event := This.monitoredChars[charName]["event"]

        hwnd := WinGetID(charName " ahk_exe exefile.exe")

        exists := This.eventMethods.Has(charName)
        if !This.lastEventPriority && exists 
            return
        else if This.lastEventPriority && exists ; Clearing last event 
            This.endEvent(charName, hwnd)

        This.debugToolTip("handleEventActivation: event " event " triggered for: " charName)

        if This.flashBorderEnabled { ; Flashing border if enabled
            try {
                This.flashMethod[charName] := Map()
                This.flashMethod[charName]["isOn"] := false
                This.flashMethod[charName]["event"] := event
                This.flashMethod[charName]["flashMethod"] := ObjBindMethod(This, "flashBorder", charName)
                This.flashBorder(charName)
                SetTimer(This.flashMethod[charName]["flashMethod"], This.flashBorderInterval)
            }
        }
        if This.showEventText {
            This.updateThumbnailText(charName . "`n" This.monitoredEventsTexts[event], hwnd)
        }
        This.eventMethods[charName] := ObjBindMethod(This, "endEvent", charName, hwnd)
        SetTimer(This.eventMethods[charName], -This.eventDisplayDuration)
    }

    endEvent(charName, hwnd) {
        if !This.eventMethods.Has(charName)
            return
        SetTimer(This.eventMethods[charName], 0)
        This.eventMethods.Delete(charName)

        This.debugToolTip("endEvent: event ended for: " charName)

        if This.flashBorderEnabled && This.flashMethod.Has(charName) {
            flashRef := This.flashMethod[charName]["flashMethod"]
            This.flashMethod.Delete(charName)
            SetTimer(flashRef, 0)
            This.clearBorder(hwnd, charName)
            ; This.ThumbWindows.%hwnd%["Border"].Show("Hide")

            if This.LastActiveThumbHwnd = hwnd
                This.BorderActive := 0

            if This.ignoredChars.Has(hwnd)
                This.toggleColorBorder(hwnd, charName)
        }
        if This.showEventText {
            This.updateThumbnailText(charName, hwnd)
        }
    }

    updateThumbnailText(title, hwnd) {
        This.ThumbWindows.%hwnd%["TextOverlay"]["OverlayText"].Text := This.CleanTitle(title)
    }

    processNPCCheck(charName, line) {
        if !This.checkNPCs || (This.monitoredChars[charName]["event"] != "" && This.monitoredChars[charName]["event"] != "stoppedShooting")
            return

        if RegExMatch(line, "<b>\d+</b>.*?>(from|to)<.*?<b><[^>]*>([^<]+)</b>", &m) { ; Getting from or to damage is dealt and target
            fromOrTo := m[1]
            target := m[2]
        } else if RegExMatch(line, "(combat) (.+?) misses you completely", &m) { ; Missed you
            fromOrTo := "from"
            target := m[1]
        } else if RegExMatch(line, "Your .+? misses (.+?) completely", &m) { ; You missed target
            fromOrTo := "to"
            target := m[1]
        } else
            return

        kind := This.ClassifyTarget(target)
        event := ""

        switch kind {
            case "player":
                if This.playerEngagmentEnabled && fromOrTo = "from"
                    event := "underAttackByPlayer"

            case "npc":
                if This.anyNPCEngagmentEnabled && fromOrTo = "from"
                    event := "underAttackByNPC"

            case "faction":
                if This.checkFactionNPCs
                    event := "engagedWithFactionBSNPC"

            case "officer":
                if This.checkOfficerNPCs
                    event := "engagedWithOfficerNPC"

            case "capital":
                if This.checkCapitalNPCs
                    event := "engagedWithCapitalNPC"

            default:
                return
        }

        if event = ""
            return

        This.monitoredChars[charName]["event"] := event
    }

    isExact(set, target) {
        return set.Has(target)
    }

    isGeneralNPC(target) {
        for pat in This.generalNPCPatterns {
            if InStr(target, pat)
                return true
        }
        return false
    }
    
    ClassifyTarget(target) {
        ; Early player catch. All players not in NPC have alli and/or corp tag
        if InStr(target, "[")
            return "player"

        if This.checkFactionNPCs && This.isExact(This.factionNPCs, target)
            return "faction"
    
        if This.checkOfficerNPCs && This.isExact(This.officerNPCs, target)
            return "officer"
    
        if This.checkCapitalNPCs && This.isExact(This.capitalNPCs, target)
            return "capital"
    
        if This.checkGeneralNPCs && This.isGeneralNPC(target)
            return "npc"

        return "player"
    }

    debugToolTip(text?) {
        if not IsSet(text) || !This.DebugMode
            return

        x := A_ScreenWidth
        y := A_ScreenHeight

        CoordMode "ToolTip", "Screen"
        ToolTip(text, x, y)
        SetTimer(() => ToolTip(), -3000)
    }

    StrJoin(delimiter, arr) {
        result := ""
        for k, v in arr
            result .= k . delimiter
        return SubStr(result, 1, -StrLen(delimiter)) ; Remove last delimiter
    }
}

