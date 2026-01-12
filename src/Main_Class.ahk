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


    EVEExe := "Ahk_Exe exefile.exe"
    
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
       
        This.TrayMenu()   
        This.MinimizeDelay := This.Minimizeclients_Delay    
        
        ;Hotkey to trigger by the script to get permissions t bring a Window in foreground
        ;Register all posible modifire combinations 
        prefixArr := ["","^","!", "#", "+", "+^", "+#", "+!", "^#", "^!","#!", "^+!", "^+#", "^#!", "+!#","^+#!"]
        for index, prefix in prefixArr
            Hotkey(  prefix . Main_Class.virtualKey, ObjBindMethod(This, "ActivateForgroundWindow"), "S P1")

        This.Updates_Checker() ; Check app updates if setting checked

        if This.First_Start_After_Update { ; Display message after succsessful update
            Version := FileGetVersion("EVE-X-Preview.exe")
            MsgBox("EVE-X-Preview succsessfully updated to version " Version)
            This.First_Start_After_Update := 0
            This.SaveJsonToFile()
        }

        ; Register Hotkey for Puase Hotkeys if the user has is Set
        if (This.Suspend_Hotkeys_Hotkey != "") {
            HotIf (*) => WinExist(This.EVEExe)
            if !This.SwitchLangOnErr {
                try {
                    Hotkey This.Suspend_Hotkeys_Hotkey, ( * ) => This.Suspend_Hotkeys(), "S1"
                }
                catch ValueError as e {
                    MsgBox(e.Message ": --> " e.Extra " <-- in: Global Settings -> Suspend Hotkeys-Hotkey" )
                }
            }
            else {
                Hotkey This.Suspend_Hotkeys_Hotkey, ( * ) => This.Suspend_Hotkeys(), "S1"
            }
        }

        ; Register Hotkey for Login Screen Cycle Hotkey if user set
        if (This.Login_Screen_Cycle_Hotkey != "") {
            if !This.SwitchLangOnErr {
                try {
                    Hotkey(This.Login_Screen_Cycle_Hotkey, ObjBindMethod(This, "Cycle_Login_Windows"),"P1" )
                }
                catch ValueError as e {
                    MsgBox(e.Message ": --> " e.Extra " <-- in Login Screen Cycle Hotkey")
                }
            }
            else {
                Hotkey(This.Login_Screen_Cycle_Hotkey, ObjBindMethod(This, "Cycle_Login_Windows"),"P1" )
            }
        }

        ; Register Hotkey for Close Active EVE Window Hotkey if user set
        if (This.Close_Active_EVE_Win_Hotkey != "") {
            if !This.SwitchLangOnErr {
                try {
                    Hotkey(This.Close_Active_EVE_Win_Hotkey, ObjBindMethod(This, "CloseActiveEVEWin"),"P1" )
                }
                catch ValueError as e {
                    MsgBox(e.Message ": --> " e.Extra " <-- in Close Active EVE Window Hotkey")
                }
            }
            else {
                Hotkey(This.Close_Active_EVE_Win_Hotkey, ObjBindMethod(This, "CloseActiveEVEWin"),"P1" )
            }
        }

        ; Register Hotkey for Close All EVE Windows Hotkey if user set
        if (This.Close_All_EVE_Win_Hotkey != "") {
            if !This.SwitchLangOnErr {
                try {
                    Hotkey(This.Close_All_EVE_Win_Hotkey, ObjBindMethod(This, "CloseAllEVEWindows"),"P1" )
                }
                catch ValueError as e {
                    MsgBox(e.Message ": --> " e.Extra " <-- in Close All EVE Windows Hotkey")
                }
            }
            else {
                Hotkey(This.Close_All_EVE_Win_Hotkey, ObjBindMethod(This, "CloseAllEVEWindows"),"P1" )
            }
        }

        ; Register Hotkey for Reload EVE-X-Preview Hotkey if user set
        if (This.Reload_Program_Hotkey != "") {
            if !This.SwitchLangOnErr {
                try {
                    Hotkey(This.Reload_Program_Hotkey, ObjBindMethod(This, "ReloadProgram"),"P1" )
                }
                catch ValueError as e {
                    MsgBox(e.Message ": --> " e.Extra " <-- in Reload EVE-X-Preview Hotkey")
                }
            }
            else {
                Hotkey(This.Reload_Program_Hotkey, ObjBindMethod(This, "ReloadProgram"),"P1" )
            }
        }

        ; Profiling
        This.ProfActive := false
        ProfEnabled := false
        if ProfEnabled
            This.StartProfiling()

        ; Resets the position of Shifting thubmnails
        This.allLoginClosed := false

        ; Skips collision check and thumb move if thumb "touch" screen edge
        This.skipShiftThumbs := false

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
        This.Save_Settings_Delay_Timer := ObjBindMethod(This, "SaveJsonToFile")
        ;Timer property to remove Thumbnails for closed EVE windows 
        This.DestroyThumbnails := ObjBindMethod(This, "EvEWindowDestroy")
        This.DestroyThumbnailsToggle := 1
        
        ;Register the Hotkeys for cycle groups 
        This.Register_Hotkey_Groups()
        This.BorderActive := 0

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
        if This.ProfActive ; Profiling
            __t0 := A_TickCount

        static HideShowToggle := 0, LastActiveHWND := 0, WinList := []

        try
            WinList := WinGetList(This.EVEExe)
        Catch 
            return

        ; If any EVE Window exist
        if WinList.Length {
            ;Check if a window exist without Thumbnail and if the user is in Character selection screen or not
            for index, hwnd in WinList {
                WinList.%hwnd% := { Title: This.CleanTitle(WinGetTitle(hwnd)) }

                if !This.ThumbWindows.HasProp(hwnd) {
                    This.EVE_WIN_Created(hwnd, WinList.%hwnd%.title)
                    if (!This.HideThumbnailsOnLostFocus) {
                        This.ShowThumb(hwnd, "Show")
                    }                      
                    HideShowToggle := 1
                }
                else { ; This change improved performance by ~15-17%
                    ; Writes character name to OldTitle if PreserveHotkeysOnLogout enabled
                    if This.PreserveHotkeysOnLogout && This.ThumbWindows.%hwnd%["Window"].Title != "" && This.ThumbWindows.%hwnd%["Window"].Title != "Char Screen" {
                        This.ThumbWindows.%hwnd%["Window"].OldTitle := This.ThumbWindows.%hwnd%["Window"].Title
                    }

                    ; If PreserveThumbPosOnLogout is false we move thumbnail after logout to default position
                    if !This.PreserveThumbPosOnLogout && This.ThumbWindows.%hwnd%["Window"].Title != "Char Screen" && This.ThumbWindows.%hwnd%["Window"].Title != WinList.%hwnd%.Title && WinList.%hwnd%.Title == "" {
                        if This.ShiftThumbsForLoginScreen
                            This.ShiftThumbs(hwnd)
                        else
                            This.ThumbMove( This.ThumbnailStartLocation["x"],
                                            This.ThumbnailStartLocation["y"],
                                            This.ThumbnailStartLocation["width"],
                                            This.ThumbnailStartLocation["height"],
                                            This.ThumbWindows.%hwnd%)
                    }

                    ; if in Character selection screen 
                    if (This.ThumbWindows.%hwnd%["Window"].Title != WinList.%hwnd%.Title && WinList.%hwnd%.Title = "" && This.PreserveCharNameOnLogout) {
                        This.ThumbWindows.%hwnd%["Window"].Title := "Char Screen"
                        if (This.ThumbWindows.%hwnd%["Window"].Title == "Char Screen" && WinList.%hwnd%.Title != "") {
                            This.EVENameChange(hwnd, WinList.%hwnd%.Title)
                        }
                    }
                    else if (This.ThumbWindows.%hwnd%["Window"].Title != WinList.%hwnd%.Title) {
                        This.EVENameChange(hwnd, WinList.%hwnd%.Title)
                    }
                }
            }

            try {
                ;if HideThumbnailsOnLostFocus is selectet check if a eve window is still in foreground, runs a timer once with a delay to prevent stuck thumbnails
                ActiveProcessName := WinGetProcessName("A")
                CallResponse := DllCall("IsIconic","UInt", WinActive("ahk_exe exefile.exe")) ; ~16-19% improvment in performance

                if ((CallResponse || ActiveProcessName != "exefile.exe") && !HideShowToggle && This.HideThumbnailsOnLostFocus) {
                    SetTimer(This.CheckforActiveWindow, -500)
                    HideShowToggle := 1
                }
                else if ( ActiveProcessName == "exefile.exe" && !CallResponse) {
                    Ahwnd := WinExist("A")
                    if This.HideThumbForActiveWin && !HideShowToggle {
                        This.ShowThumb(Ahwnd, "Hide")

                        if LastActiveHWND && LastActiveHWND != Ahwnd
                            This.ShowThumb(LastActiveHWND, "Show")
                            This.UpdateThumb_AfterActivation(, Ahwnd)
                        LastActiveHWND := Ahwnd
                    }
                    else {
                        if HideShowToggle {
                            for EVEHWND in This.ThumbWindows.OwnProps() {
                                This.ShowThumb(EVEHWND, "Show")
                            }
                            HideShowToggle := 0
                            This.BorderActive := 0
                        }
                        
                        ; sets the Border to the active window thumbnail 
                        else if (Ahwnd != This.BorderActive) {
                            ;Shows the Thumbnail on top of other thumbnails
                            if (This.ShowThumbnailsAlwaysOnTop)
                                WinSetAlwaysOnTop(1,This.ThumbWindows.%Ahwnd%["Window"].Hwnd )
                            
                            This.ShowActiveBorder(Ahwnd)
                            This.UpdateThumb_AfterActivation(, Ahwnd)
                            This.BorderActive := Ahwnd
                        }
                    }
                }
            }
        }

        ; Check if a Thumbnail exist without EVE Window. if so destroy the Thumbnail and free memory
        if ( This.DestroyThumbnailsToggle ) {
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
                    if This.ThumbWindows.%EVEHWND%["Window"].Title == "" {
                        This.allLoginClosed := false
                        break
                    }
                    This.allLoginClosed := true
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
    

    Updates_Checker() {
        if !This.Check_Updates
            return

        try {
            ; Getting .exe version
            Version := FileGetVersion("EVE-X-Preview.exe")

            ; Getting json of latest release
            apiUrl := "https://api.github.com/repos/khivus/EVE-X-Preview/releases/latest"
            whr := ComObject("WinHttp.WinHttpRequest.5.1")
            whr.Open("GET", apiUrl)
            whr.SetRequestHeader("User-Agent", "AHK")
            whr.Send()
            whr.WaitForResponse()
            json_ans := whr.ResponseText

            ; Finding tag of latest release
            tag := RegExReplace(json_ans, '.*"tag_name":\s*"([^"]+)".*', "$1")
            tag := StrReplace(tag, "v")

            if tag == Version
                return ; No update available

            ; Finding download link
            if RegExMatch(json_ans, '"browser_download_url"\s*:\s*"([^"]+\.exe)"', &m)
                exeUrl := m[1]

            Message := "New version " tag " available!`n" . 
                "Current version:" Version "`n`n" . 
                "`"Cancel`" or `"X`" button will disable updates!`n`n" . 
                "Do you want automatically download and install the update now?`n" . 
                "Please don`'t touch the application until the update is complete."

            result := MsgBox(Message, "EVE-X-Preview update", "YesNoCancel")

            if result = "Yes" {
                ; Downloading file and running
                scriptDir := A_ScriptDir
                scriptName := A_ScriptName
                scriptPath := (scriptDir "\" scriptName)
                newPath := scriptDir "\EVE-X-Preview_new.exe"

                ; Checking if files exist and deleting if so
                if FileExist("EVE-X-Preview_new.exe")
                    FileDelete("EVE-X-Preview_new.exe")
                
                Download(exeUrl, newPath) ; Download file from GitHub

                ; Checking if all variables are correct
                if !FileExist(scriptPath)
                    Throw Error("Could not get application path!")
                if !FileExist(newPath)
                    Throw Error("Error downloading new version!")
                if !scriptName
                    Throw Error("Could not get application name!")
                
                batContent := '@echo off' . "`r`n" .
                    'ping 127.0.0.1 -n 6 >nul' . "`r`n" .
                    'del /f /q "' . scriptPath . '"' . "`r`n" .
                    'if exist "' . newPath . '" (' . "`r`n" .
                    '  ren "' . newPath . '" "' . scriptName . '"' . "`r`n" .
                    '  start "" "' . scriptPath . '"' . "`r`n" .
                    ')' . "`r`n" .
                    'del "' . scriptDir '\update.bat"'

                ; Removing old update.bat if exist
                if FileExist("update.bat")
                    FileDelete("update.bat")

                FileAppend(batContent, "update.bat") ; Creating update.bat

                if !FileExist("update.bat")
                    Throw Error("Could not create update.bat!")

                This.First_Start_After_Update := 1 ; For showing update message
                This.SaveJsonToFile()

                Run("update.bat", scriptDir, "hide") ; Running update.bat
                ExitApp()
            }
            else if result = "Cancel" {
                This.Check_Updates := 0
                This.SaveJsonToFile()
            }
        }
        catch ValueError as e {
            MsgBox("An error occurred while trying to update the application:`n" e.Message "`n" e.Extra "`nUpdate checker is disabled.")
            This.Check_Updates := 0
            This.SaveJsonToFile()
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

    ;Register set Hotkeys by the user in settings
    RegisterHotkeys(title, EvE_hwnd) {  
        static registerGroups := 0
        ;if the user has set Hotkeys in Options 
        if (This._Hotkeys[title]) {  
            ;if the user has selected Global Hotkey. This means the Hotkey will alsways trigger as long at least 1 EVE Window exist.
            ;if a Window does not Exist which was assigned to the hotkey the hotkey will be dissabled until the Window exist again
            if(This.Global_Hotkeys) {
                HotIf (*) => WinExist(This.EVEExe) && WinExist("EVE - " title ) && !WinActive("EVE-X-Preview - Settings")
                if !This.SwitchLangOnErr {
                    try {
                        Hotkey This._Hotkeys[title], (*) => This.ActivateEVEWindow(,,title), "P1"
                    }
                    catch ValueError as e {
                        MsgBox(e.Message ": --> " e.Extra " <-- in Profile Settings - " This.LastUsedProfile " Hotkeys" )
                    }
                }
                else {
                    Hotkey This._Hotkeys[title], (*) => This.ActivateEVEWindow(,,title), "P1"
                }
            }
            ;if the user has selected (Win Active) the hotkeys will only trigger if at least 1 EVE Window is Active and in Focus
            ;This makes it possible to still use all keys outside from EVE 
            else {
                HotIf (*) => WinExist("EVE - " title ) && WinActive(This.EVEExe)
                if !This.SwitchLangOnErr {
                    try {
                        Hotkey This._Hotkeys[title], (*) => This.ActivateEVEWindow(,,title),"P1"
                    }
                    catch ValueError as e {
                        MsgBox(e.Message ": --> " e.Extra " <-- in Profile Settings - " This.LastUsedProfile " Hotkeys" )
                    }
                }
                else {
                    Hotkey This._Hotkeys[title], (*) => This.ActivateEVEWindow(,,title),"P1"
                }
            }
        }
    }    

    ;Register the Hotkeys for cycle Groups if any set
    Register_Hotkey_Groups() {
        static Fkey := "", BKey := "", Arr := []
        if (IsObject(This.Hotkey_Groups) && This.Hotkey_Groups.Count != 0) {
            for k, v in This.Hotkey_Groups {
                ;If any EVE Window Exist and at least 1 character matches the the list from the group windows
                if(This.Global_Hotkeys) {
                    if( v["ForwardsHotkey"] != "" ) {                        
                        Fkey := v["ForwardsHotkey"], Arr := v["Characters"]
                        HotIf ObjBindMethod(This, "OnWinExist", Arr)
                        if !This.SwitchLangOnErr {
                            try {
                                Hotkey( v["ForwardsHotkey"], ObjBindMethod(This, "Cycle_Hotkey_Groups",Arr,"ForwardsHotkey"), "P1")
                            }
                            catch ValueError as e {
                                MsgBox(e.Message ": --> " e.Extra " <-- in Profile Settings - " This.LastUsedProfile " - Hotkey Groups - " k "  - Forwards Hotkey" )
                            }
                        }
                        else {
                            Hotkey( v["ForwardsHotkey"], ObjBindMethod(This, "Cycle_Hotkey_Groups",Arr,"ForwardsHotkey"), "P1")
                        }
                    }
                    if( v["BackwardsHotkey"] != "" ) {
                        Fkey := v["BackwardsHotkey"], Arr := v["Characters"]
                        HotIf ObjBindMethod(This, "OnWinExist", Arr)
                        if !This.SwitchLangOnErr {
                            try {
                                Hotkey( v["BackwardsHotkey"], ObjBindMethod(This, "Cycle_Hotkey_Groups",Arr,"BackwardsHotkey"), "P1")   
                            }
                            catch ValueError as e {
                                MsgBox(e.Message ": --> " e.Extra " <-- in Profile Settings - " This.LastUsedProfile " Hotkey Groups - " k " - Backwards Hotkey" )
                            }
                        }
                        else {
                            Hotkey( v["BackwardsHotkey"], ObjBindMethod(This, "Cycle_Hotkey_Groups",Arr,"BackwardsHotkey"), "P1")   
                        }
                    }  
                }  
                ;If any EVE Window is Active
                else {
                    if( v["ForwardsHotkey"] != "" ) {
                        Fkey := v["ForwardsHotkey"], Arr := v["Characters"]
                        HotIf ObjBindMethod(This, "OnWinActive", Arr)
                        if !This.SwitchLangOnErr {
                            try {
                                Hotkey( v["ForwardsHotkey"], ObjBindMethod(This, "Cycle_Hotkey_Groups",Arr,"ForwardsHotkey"), "P1")
                            }
                            catch ValueError as e {
                                MsgBox(e.Message ": --> " e.Extra " <-- in Profile Settings - " This.LastUsedProfile " - Hotkey Groups - " k "  - Forwards Hotkey" )
                            }
                        }
                        else {
                            Hotkey( v["ForwardsHotkey"], ObjBindMethod(This, "Cycle_Hotkey_Groups",Arr,"ForwardsHotkey"), "P1")
                        }
                    }
                    if( v["BackwardsHotkey"] != "" ) {
                        Fkey := v["BackwardsHotkey"], Arr := v["Characters"]
                        HotIf ObjBindMethod(This, "OnWinActive", Arr)
                        if !This.SwitchLangOnErr {
                            try {
                                Hotkey( v["BackwardsHotkey"], ObjBindMethod(This, "Cycle_Hotkey_Groups",Arr,"BackwardsHotkey"), "P1")   
                            }
                            catch ValueError as e {
                                MsgBox(e.Message ": --> " e.Extra " <-- in Profile Settings - " This.LastUsedProfile " Hotkey Groups - " k " - Backwards Hotkey" )
                            } 
                        }
                        else {
                            Hotkey( v["BackwardsHotkey"], ObjBindMethod(This, "Cycle_Hotkey_Groups",Arr,"BackwardsHotkey"), "P1")   
                        }
                    }  
                }             
            }
        }
    }

    ; The method to make it possible to cycle throw the EVE Windows. Used with the Hotkey Groups
     Cycle_Hotkey_Groups(Arr, direction,*) {
        static Index := 0
        HWND := 0
        activateByHWND := 0
        length := Arr.Length

        if (direction == "ForwardsHotkey") {
            try {
                if !This.PreserveHotkeysOnLogout
                    Index := (n := IsActiveWinInGroup(This.CleanTitle(WinGetTitle("A")), Arr)) ? n+1 : 1

                else {
                    AHWND := WinExist("A")
                    if WinGetProcessName(AHWND) == "exefile.exe"
                        Index := (n := IsActiveWinInGroup(This.ThumbWindows.%AHWND%["Window"].OldTitle, Arr)) ? n+1 : 1
                    else
                        Index := 1
                }
            }
        
            if (Index > length)
                Index := 1

            if (This.OnWinExist(Arr)) {
                try {
                    if !(WinExist("EVE - " Arr[Index])) {
                        while (!(WinExist("EVE - " Arr[Index]))) {
                            if HWND := This.hasMathcingOldTitle(Arr[Index]) {
                                activateByHWND := 1
                                break
                            }
                            index += 1
                            if (Index > length)
                                Index := 1
                        }
                    }

                    if !activateByHWND
                        This.ActivateEVEWindow(,,Arr[Index])

                    else
                        This.ActivateEVEWindow(HWND,,)
                }
            }
        }

        else if (direction == "BackwardsHotkey") {
            try {
                if !This.PreserveHotkeysOnLogout
                    Index := (n := IsActiveWinInGroup(This.CleanTitle(WinGetTitle("A")), Arr)) ? n-1 : length

                else {
                    AHWND := WinExist("A")
                    if WinGetProcessName(AHWND) == "exefile.exe"
                        Index := (n := IsActiveWinInGroup(This.ThumbWindows.%AHWND%["Window"].OldTitle, Arr)) ? n-1 : length
                    else 
                        Index := length
                }
            }
            
            if (Index <= 0)
                Index := length

            if (This.OnWinExist(Arr)) {
                try {
                    if !(WinExist("EVE - " Arr[Index])) {
                        while (!(WinExist("EVE - " Arr[Index]))) {
                            if HWND := This.hasMathcingOldTitle(Arr[Index]) {
                                activateByHWND := 1
                                break
                            }
                            Index -= 1
                            if (Index <= 0)
                                Index := length
                        }
                    }
                    if !activateByHWND
                        This.ActivateEVEWindow(,,Arr[Index])

                    else
                        This.ActivateEVEWindow(HWND,,)
                }
            }
        }

        IsActiveWinInGroup(Title, Arr) {
            for index, names in Arr {
                if names = Title
                    return index
            }
            return false
        }
    }

    ; Checks for OldTitle == CleanTitle in all login screen windows
    ; return HWND if found
    hasMathcingOldTitle(CleanTitle) {
        if !This.PreserveHotkeysOnLogout
            return

        loginHWNDs := WinGetList("EVE")

        for hwnd in loginHWNDs {
            if This.ThumbWindows.%hwnd%["Window"].OldTitle == CleanTitle
                return hwnd
        }
        return
    }

    ; Cycle windows on Character selection screen
    Cycle_Login_Windows(*) {
        LoginWins := []
        currentHWND := WinExist("A")
        loginHWNDs := WinGetList("EVE")

        for hwnd in loginHWNDs {
            if This.ThumbWindows.%hwnd%["Window"].OldTitle && This.PreserveHotkeysOnLogout
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

        if LoginWins.Length == 1 {
            This.ActivateEVEWindow(LoginWins[1]["hwnd"],,)
            return
        }

        if (LoginWins.Length > 1 ) {
            LoginWins := This.CustomSort(LoginWins, "CreationTime")
        }

        for i, Win in LoginWins {
            if currentHWND == Win["hwnd"] {
                currentIndex := i
                break
            }
            else
                currentIndex := 0
        }

        currentIndex += 1
        if currentIndex > LoginWins.Length
            currentIndex := 1

        This.ActivateEVEWindow(LoginWins[currentIndex]["hwnd"],,)
    }

    ; Close Active EVE Client 
    CloseActiveEVEWin(*) {
        if WinActive("ahk_exe exefile.exe")
		    WinClose("A")
    }

    ; Reload program
    ReloadProgram(*) {
        Reload
    }

     ; To Check if atleast One Win stil Exist in the Array for the cycle groups hotkeys
    OnWinExist(Arr, *) {
        for index, Name in Arr {
            ; If ( WinExist("EVE - " Name " Ahk_Exe exefile.exe") && !WinActive("EVE-X-Preview - Settings") && !This.PreserveHotkeysOnLogout) {
            If ( WinExist("EVE - " Name " Ahk_Exe exefile.exe") && !WinActive("EVE-X-Preview - Settings") ) {
                return true
            }
            else if This.PreserveHotkeysOnLogout && This.hasMathcingOldTitle(Name) && !WinActive("EVE-X-Preview - Settings") {
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
        MinMax := -1
        try MinMax := WinGetMinMax("ahk_id " hwnd)

        if (This.ThumbWindows.HasProp(hwnd)) {
            if !(MinMax == -1) {
                This.Update_Thumb(false, This.ThumbWindows.%hwnd%["Window"].Hwnd)
            }
        }
    }

    ;This function updates the Thumbnails and hotkeys if the user switches Charakters in the character selection screen 
    EVENameChange(hwnd, title) {
        if (This.ThumbWindows.HasProp(hwnd)) {
            This.SetThumbnailText[hwnd] := title
            ; moves the Window to the saved positions if any stored, a bit of sleep is usfull to give the window time to move before creating the thumbnail
            This.RestoreClientPossitions(hwnd, title)

            ; if (title = "") {
            ;     This.EvEWindowDestroy(hwnd, title)
            ;     This.EVE_WIN_Created(hwnd,title)
            ; }

            If (This.ThumbnailPositions.Has(title)) {
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
                This.Update_Thumb(true)
                If ( This.HideThumbnailsOnLostFocus && WinActive(This.EVEExe) || !This.HideThumbnailsOnLostFocus && !WinActive(This.EVEExe) || !This.HideThumbnailsOnLostFocus && WinActive(This.EVEExe)) {
                    for k, v in This.ThumbWindows.OwnProps()
                        This.ShowThumb(k, "Show")
                } 
            }
            This.BorderActive := 0
            This.RegisterHotkeys(title, hwnd)
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
                            ;sleep 1
                            This.Mouse_DragMove(wparam, lparam, msg, hwnd)
                            This.Window_Snap(hwnd, This.ThumbWindows)
                        }
                        else
                            This.Mouse_ResizeThumb(wparam, lparam, msg, hwnd)
                    }                    
                return 0
            }

            ; Wparam -  9 Ctrl+Lclick
            ;           5 Shift+Lclick
            ;           13 Shift+ctrl+click
            Else If (msg == Main_Class.WM_LBUTTONDOWN) {
                ;Activates the EVE Window by clicking on the Thumbnail 
                if (wparam = 1) {
                    if !(WinActive(This.ThumbHwnd_EvEHwnd[hwnd]))
                        This.ActivateEVEWindow(hwnd)
                }
                ; Ctrl+Lbutton, Minimizes the Window on whose thumbnail the user clicks
                else if (wparam = 9) { 
                    ; Minimize
                    if (!GetKeyState("RButton"))
                        PostMessage 0x0112, 0xF020, , , This.ThumbHwnd_EvEHwnd[hwnd]
                }
                return 0
            }   
        }
    }

    ; Creates a new thumbnail if a new window got created
    EVE_WIN_Created(Win_Hwnd, Win_Title) {
        ; Moves the Window to the saved possition if any are stored 
        This.RestoreClientPossitions(Win_Hwnd, Win_Title)        
        
        ;Creates the Thumbnail and stores the EVE Hwnd in the array
        If !(This.ThumbWindows.HasProp(Win_Hwnd)) {       
            This.ThumbWindows.%Win_Hwnd% := This.Create_Thumbnail(Win_Hwnd, Win_Title)
            This.ThumbHwnd_EvEHwnd[This.ThumbWindows.%Win_Hwnd%["Window"].Hwnd] := Win_Hwnd
            This.ThumbWindows.%Win_Hwnd%["Window"].OldTitle := ""

            ;if the User is in character selection screen show the window always 
            if (This.ThumbWindows.%Win_Hwnd%["Window"].Title = "") {
                This.SetThumbnailText[Win_Hwnd] := Win_Title
                This.ShiftThumbs(Win_Hwnd)
                ;if the Title is just "EVE" that means it is in the Charakter selection screen
                ;in this case show always the Thumbnail 
                ; This.ShowThumb(Win_Hwnd, "Show")
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
                This.Update_Thumb(true)
                If ( This.HideThumbnailsOnLostFocus && WinActive(This.EVEExe) || !This.HideThumbnailsOnLostFocus && !WinActive(This.EVEExe) || !This.HideThumbnailsOnLostFocus && WinActive(This.EVEExe)) {
                    for k, v in This.ThumbWindows.OwnProps()
                        This.ShowThumb(k, "Show")
                }
            }
            This.RegisterHotkeys(Win_Title, Win_Hwnd)
        }
    }

    ; if ShiftThumbsForLoginScreen enabled we try to shift thumbnail using user settings
    ShiftThumbs(Win_Hwnd) {
        if !This.ShiftThumbsForLoginScreen || This.skipShiftThumbs
            return

        static nextPosX := This.ThumbnailStartLocation["x"]
        static nextPosY := This.ThumbnailStartLocation["y"]
        step_x := This.ShiftThumbHorizontalStep
        step_y := This.ShiftThumbVerticalStep

        if step_x == 0
            step_x := This.ThumbnailStartLocation["width"]
        if step_y == 0
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
        if (IsSet(hwnd) && This.ThumbWindows.HasProp(hwnd)) {
            for k, v in This.ThumbWindows.Clone().%hwnd% {
                if (K = "Thumbnail")
                    continue
                v.Destroy()
                ;This.ThumbWindows.%Win_Hwnd%.Delete()
            }
            This.ThumbWindows.DeleteProp(hwnd)
            Return
        }
        ;If a EVE Windows get destroyed 
        for Win_Hwnd,v in This.ThumbWindows.Clone().OwnProps() {
            if (!WinExist("Ahk_Id " Win_Hwnd)) {
                for k,v in This.ThumbWindows.Clone().%Win_Hwnd% {
                    if (K = "Thumbnail")
                        continue
                    v.Destroy()
                }
                This.ThumbWindows.DeleteProp(Win_Hwnd)        
            }
        }
        This.DestroyThumbnailsToggle := 1
    }
    
    ActivateEVEWindow(hwnd?,ThisHotkey?, title?) {   
        ; If the user clicks the Thumbnail then hwnd stores the Thumbnail Hwnd. Here the Hwnd gets changed to the contiguous EVE window hwnd
        if (IsSet(hwnd) && This.ThumbHwnd_EvEHwnd.Has(hwnd)) {
            hwnd := WinExist(This.ThumbHwnd_EvEHwnd[hwnd])
            title := This.CleanTitle(WinGetTitle("Ahk_id " Hwnd))
        }
        ;if the user presses the Hotkey 
        Else if (IsSet(title)) {
            title := "EVE - " title
            hwnd := WinExist(title " Ahk_exe exefile.exe")
        }
        ;return when the user tries to bring a window to foreground which is already in foreground 
        if (WinActive("Ahk_id " hwnd))
            Return

        If (DllCall("IsIconic", "UInt", hwnd)) {
            if (This.AlwaysMaximize)  || ( This.TrackClientPossitions && This.ClientPossitions[This.CleanTitle(title)]["IsMaximized"] ) {
                ; ; Maximize
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
            if (This.AlwaysMaximize && WinGetMinMax("ahk_id " This.ActivateHwnd) = 0) || ( This.TrackClientPossitions && This.ClientPossitions[This.CleanTitle(WinGetTitle("Ahk_id " This.ActivateHwnd))]["IsMaximized"] && WinGetMinMax("ahk_id " This.ActivateHwnd) = 0 )
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
            WinTitle := This.CleanTitle(EVEwinTitle)
            if !(WinTitle = "") {
                for k in This.Dont_Minimize_Clients {
                    value := This.CleanTitle(k)
                    if value == WinTitle
                        return 1
                }
                return 0
            }
        }
    }

    ; Function t move the Thumbnails into the saved positions from the user
    ThumbMove(x := "", y := "", Width := "", Height := "", GuiObj := "") {
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
            Title := This.CleanTitle(WinGetTitle("Ahk_id " v))
            if !(Title = "") {
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
    CleanTitle(title) {
        Return RegExReplace(title, "^(?i)eve(?:\s*-\s*)?\b", "")
        ;RegExReplace(title, "(?i)eve\s*-\s*", "")

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
}

