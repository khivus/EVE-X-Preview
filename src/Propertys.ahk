

class Propertys extends TrayMenu {


    ;######################
    ;## Script Propertys

    SetThumbnailText[hwnd, *] {
        set {
            if This.ThumbWindows.HasProp(hwnd) {
                newtext := Value
                This.ThumbWindows.%hwnd%["TextOverlay"]["OverlayText"].Text := This.CleanTitle(newtext)
                This.ThumbWindows.%hwnd%["Window"].Title := newtext
            }
        }
    }

    Profiles => This._JSON["_Profiles"]


    ;######################
    ; General settings

    First_Start_After_Update {
        get => This._JSON["First_Start_After_Update"]
        set => This._JSON["First_Start_After_Update"] := value
    }

    ThisThat {
        get => This._JSON["ThisThat"]
        set => This._JSON["ThisThat"] := value
    }

    DebugMode {
        get => This._JSON["DebugMode"]
        set => This._JSON["DebugMode"] := value
    }

    LastUsedProfile {
        get => This._JSON["LastUsedProfile"]
        set => This._JSON["LastUsedProfile"] := value
    }

    AntiGlobalGroups {
        get => This._JSON["_Profiles"][This.LastUsedProfile]["Other"]["Global_Groups"]
        set => This._JSON["_Profiles"][This.LastUsedProfile]["Other"]["Global_Groups"] := value
    }

    Global_Groups {
        get => This._JSON["_Profiles"]["Default"]["Other"]["Global_Groups"]
        set => This._JSON["_Profiles"]["Default"]["Other"]["Global_Groups"] := value
    }

    ProfileOverride() {
        This.ComboGroups := Map()
        for ag, av in This.AntiGlobalGroups {
            for g, v in This.Global_Groups {
                if g == ag {
                    if v
                        This.ComboGroups[g] := av ? 0 : 1
                    else
                        This.ComboGroups[g] := 0

                    break
                }
            }
        }

        This.ProfileHotkeysSettings := This.ComboGroups["Hotkeys Settings"] ? "Default" : This.LastUsedProfile
        This.ProfileThumbnailsBehavior := This.ComboGroups["Thumbnails Behavior"] ? "Default" : This.LastUsedProfile
        This.ProfileThumbnailsVisuals := This.ComboGroups["Thumbnails Visuals"] ? "Default" : This.LastUsedProfile
        This.ProfileThumbnailVisibility := This.ComboGroups["Thumbnail Visibility"] ? "Default" : This.LastUsedProfile
        This.ProfileClientSettings := This.ComboGroups["Client Settings"] ? "Default" : This.LastUsedProfile
        This.ProfileCustomColors := This.ComboGroups["Custom Colors"] ? "Default" : This.LastUsedProfile
        This.ProfileGameLogsMonitoring := This.ComboGroups["Game Logs Monitoring"] ? "Default" : This.LastUsedProfile
        This.ProfileMonitoredEvents := This.ComboGroups["Monitored Events"] ? "Default" : This.LastUsedProfile
        This.ProfileTrayMenuSettings := This.ComboGroups["Tray Menu Settings"] ? "Default" : This.LastUsedProfile
        This.ProfileOther := This.ComboGroups["Other"] ? "Default" : This.LastUsedProfile
        This.ProfileHotkeysGroups := This.ComboGroups["Hotkey Groups"] ? "Default" : This.LastUsedProfile
        This.ProfileNonEVEApplications := This.ComboGroups["Non-EVE Applications"] ? "Default" : This.LastUsedProfile
    }



    ;########################
    ;## Profile ThumbnailSettings

    ThumbnailStartLocation[key] {
        get => This._JSON["_Profiles"][This.ProfileThumbnailsBehavior]["Thumbnails Behavior"]["ThumbnailStartLocation"][key]
        set => This._JSON["_Profiles"][This.ProfileThumbnailsBehavior]["Thumbnails Behavior"]["ThumbnailStartLocation"][key] := value
    }

    AutoSaveThumbnailPositions {
        get => This._JSON["_Profiles"][This.ProfileThumbnailsBehavior]["Thumbnails Behavior"]["AutoSaveThumbnailPositions"]
        set => This._JSON["_Profiles"][This.ProfileThumbnailsBehavior]["Thumbnails Behavior"]["AutoSaveThumbnailPositions"] := value
    }

    ThumbnailBackgroundColor {
        get => convertToHex(This._JSON["_Profiles"][This.ProfileThumbnailsVisuals]["Thumbnails Visuals"]["ThumbnailBackgroundColor"])
        set => This._JSON["_Profiles"][This.ProfileThumbnailsVisuals]["Thumbnails Visuals"]["ThumbnailBackgroundColor"] := convertToHex(value)
    }

    ThumbnailSnap[*] {
        get => This._JSON["_Profiles"][This.ProfileThumbnailsBehavior]["Thumbnails Behavior"]["ThumbnailSnap"]
        set => This._JSON["_Profiles"][This.ProfileThumbnailsBehavior]["Thumbnails Behavior"]["ThumbnailSnap"] := Value
    }

    ThumbnailSnap_Distance {
        get => This._JSON["_Profiles"][This.ProfileThumbnailsBehavior]["Thumbnails Behavior"]["ThumbnailSnap_Distance"]
        set => This._JSON["_Profiles"][This.ProfileThumbnailsBehavior]["Thumbnails Behavior"]["ThumbnailSnap_Distance"] := (value ? value : "20")
    }

    ThumbnailMinimumSize[key] {
        get => This._JSON["_Profiles"][This.ProfileThumbnailsVisuals]["Thumbnails Visuals"]["ThumbnailMinimumSize"][key]
        set => This._JSON["_Profiles"][This.ProfileThumbnailsVisuals]["Thumbnails Visuals"]["ThumbnailMinimumSize"][key] := value
    }

    PreserveThumbPosOnLogout {
        get => This._JSON["_Profiles"][This.ProfileThumbnailsBehavior]["Thumbnails Behavior"]["PreserveThumbPosOnLogout"]
        set => This._JSON["_Profiles"][This.ProfileThumbnailsBehavior]["Thumbnails Behavior"]["PreserveThumbPosOnLogout"] := value
    }

    PreserveCharNameOnLogout {
        get => This._JSON["_Profiles"][This.ProfileThumbnailsBehavior]["Thumbnails Behavior"]["PreserveCharNameOnLogout"]
        set => This._JSON["_Profiles"][This.ProfileThumbnailsBehavior]["Thumbnails Behavior"]["PreserveCharNameOnLogout"] := value
    }

    HideThumbForActiveWin {
        get => This._JSON["_Profiles"][This.ProfileThumbnailsBehavior]["Thumbnails Behavior"]["HideThumbForActiveWin"]
        set => This._JSON["_Profiles"][This.ProfileThumbnailsBehavior]["Thumbnails Behavior"]["HideThumbForActiveWin"] := value
    }

    ShiftThumbsForLoginScreen {
        get => This._JSON["_Profiles"][This.ProfileThumbnailsBehavior]["Thumbnails Behavior"]["ShiftThumbsForLoginScreen"]
        set => This._JSON["_Profiles"][This.ProfileThumbnailsBehavior]["Thumbnails Behavior"]["ShiftThumbsForLoginScreen"] := value
    }

    ShiftThumbsCollisionCheck {
        get => This._JSON["_Profiles"][This.ProfileThumbnailsBehavior]["Thumbnails Behavior"]["ShiftThumbsCollisionCheck"]
        set => This._JSON["_Profiles"][This.ProfileThumbnailsBehavior]["Thumbnails Behavior"]["ShiftThumbsCollisionCheck"] := value
    }

    ShiftThumbsDirection {
        get => This._JSON["_Profiles"][This.ProfileThumbnailsBehavior]["Thumbnails Behavior"]["ShiftThumbsDirection"]
        set => This._JSON["_Profiles"][This.ProfileThumbnailsBehavior]["Thumbnails Behavior"]["ShiftThumbsDirection"] := value
    }

    ShiftThumbHorizontalStep {
        get => This._JSON["_Profiles"][This.ProfileThumbnailsBehavior]["Thumbnails Behavior"]["ShiftThumbHorizontalStep"]
        set => This._JSON["_Profiles"][This.ProfileThumbnailsBehavior]["Thumbnails Behavior"]["ShiftThumbHorizontalStep"] := value
    }

    ShiftThumbVerticalStep {
        get => This._JSON["_Profiles"][This.ProfileThumbnailsBehavior]["Thumbnails Behavior"]["ShiftThumbVerticalStep"]
        set => This._JSON["_Profiles"][This.ProfileThumbnailsBehavior]["Thumbnails Behavior"]["ShiftThumbVerticalStep"] := value
    }

    HideThumbnails {
        get => This._JSON["_Profiles"][This.ProfileThumbnailsBehavior]["Thumbnails Behavior"]["HideThumbnails"]
        set => This._JSON["_Profiles"][This.ProfileThumbnailsBehavior]["Thumbnails Behavior"]["HideThumbnails"] := value
    }
    
    ClickThroughActive {
        get => This._JSON["_Profiles"][This.ProfileThumbnailsBehavior]["Thumbnails Behavior"]["ClickThroughActive"]
        set => This._JSON["_Profiles"][This.ProfileThumbnailsBehavior]["Thumbnails Behavior"]["ClickThroughActive"] := value
    }

    ShowAllColoredBorders {
        get => This._JSON["_Profiles"][This.ProfileThumbnailsVisuals]["Thumbnails Visuals"]["ShowAllColoredBorders"]
        set => This._JSON["_Profiles"][This.ProfileThumbnailsVisuals]["Thumbnails Visuals"]["ShowAllColoredBorders"] := value
    }

    Thumbnail_visibility[key?] {
        get {
            return This._JSON["_Profiles"][This.ProfileThumbnailVisibility]["Thumbnail Visibility"]
        }
        set {
            if (IsObject(value)) {
                This._JSON["_Profiles"][This.ProfileThumbnailVisibility]["Thumbnail Visibility"] := value
            }
            This.Save_Settings()

        }
    }


    HideThumbnailsOnLostFocus {
        get => This._JSON["_Profiles"][This.ProfileThumbnailsBehavior]["Thumbnails Behavior"]["HideThumbnailsOnLostFocus"]
        set => This._JSON["_Profiles"][This.ProfileThumbnailsBehavior]["Thumbnails Behavior"]["HideThumbnailsOnLostFocus"] := value
    }
    ShowThumbnailsAlwaysOnTop {
        get => This._JSON["_Profiles"][This.ProfileThumbnailsBehavior]["Thumbnails Behavior"]["ShowThumbnailsAlwaysOnTop"]
        set => This._JSON["_Profiles"][This.ProfileThumbnailsBehavior]["Thumbnails Behavior"]["ShowThumbnailsAlwaysOnTop"] := value
    }

    ThumbnailOpacity {
        get {
            percentage := This._JSON["_Profiles"][This.ProfileThumbnailsVisuals]["Thumbnails Visuals"]["ThumbnailOpacity"]
            return Round((percentage < 0 ? 0 : percentage > 100 ? 100 : percentage) * 2.55)
        }
        set {
            This._JSON["_Profiles"][This.ProfileThumbnailsVisuals]["Thumbnails Visuals"]["ThumbnailOpacity"] := Value
        }
    }

    ClientHighligtBorderthickness {
        get => This._JSON["_Profiles"][This.ProfileThumbnailsVisuals]["Thumbnails Visuals"]["ClientHighligtBorderthickness"]
        set => This._JSON["_Profiles"][This.ProfileThumbnailsVisuals]["Thumbnails Visuals"]["ClientHighligtBorderthickness"] := (Trim(value, "`n ") <= 0 ? 1 : Trim(value, "`n "))
    }

    ClientHighligtColor {
        get => convertToHex(This._JSON["_Profiles"][This.ProfileThumbnailsVisuals]["Thumbnails Visuals"]["ClientHighligtColor"])
        set => This._JSON["_Profiles"][This.ProfileThumbnailsVisuals]["Thumbnails Visuals"]["ClientHighligtColor"] := convertToHex(Trim(value, "`n "))
    }
    ShowClientHighlightBorder {
        get => This._JSON["_Profiles"][This.ProfileThumbnailsVisuals]["Thumbnails Visuals"]["ShowClientHighlightBorder"]
        set => This._JSON["_Profiles"][This.ProfileThumbnailsVisuals]["Thumbnails Visuals"]["ShowClientHighlightBorder"] := value
    }
    ThumbnailTextFont {
        get => This._JSON["_Profiles"][This.ProfileThumbnailsVisuals]["Thumbnails Visuals"]["ThumbnailTextFont"]
        set => This._JSON["_Profiles"][This.ProfileThumbnailsVisuals]["Thumbnails Visuals"]["ThumbnailTextFont"] := Trim(value, "`n ")
    }
    ThumbnailTextSize {
        get => This._JSON["_Profiles"][This.ProfileThumbnailsVisuals]["Thumbnails Visuals"]["ThumbnailTextSize"]
        set => This._JSON["_Profiles"][This.ProfileThumbnailsVisuals]["Thumbnails Visuals"]["ThumbnailTextSize"] := Trim(value, "`n ")
    }

    ThumbnailTextColor {
        get => convertToHex(This._JSON["_Profiles"][This.ProfileThumbnailsVisuals]["Thumbnails Visuals"]["ThumbnailTextColor"])
        set => This._JSON["_Profiles"][This.ProfileThumbnailsVisuals]["Thumbnails Visuals"]["ThumbnailTextColor"] := convertToHex(Trim(value, "`n "))
    }
    ShowThumbnailTextOverlay {
        get => This._JSON["_Profiles"][This.ProfileThumbnailsVisuals]["Thumbnails Visuals"]["ShowThumbnailTextOverlay"]
        set => This._JSON["_Profiles"][This.ProfileThumbnailsVisuals]["Thumbnails Visuals"]["ShowThumbnailTextOverlay"] := value
    }
    ThumbnailTextMargins[var] {
        get => This._JSON["_Profiles"][This.ProfileThumbnailsVisuals]["Thumbnails Visuals"]["ThumbnailTextMargins"][var]
        set => This._JSON["_Profiles"][This.ProfileThumbnailsVisuals]["Thumbnails Visuals"]["ThumbnailTextMargins"][var] := Trim(value, "`n ")
    }
    InactiveClientBorderthickness {
        get {
            if ( !This._JSON["_Profiles"][This.ProfileThumbnailsVisuals]["Thumbnails Visuals"].Has("InactiveClientBorderthickness") ) 
                This._JSON["_Profiles"][This.ProfileThumbnailsVisuals]["Thumbnails Visuals"]["InactiveClientBorderthickness"] := "2"
            return This._JSON["_Profiles"][This.ProfileThumbnailsVisuals]["Thumbnails Visuals"]["InactiveClientBorderthickness"]
        } 
        set => This._JSON["_Profiles"][This.ProfileThumbnailsVisuals]["Thumbnails Visuals"]["InactiveClientBorderthickness"] := (Trim(value, "`n ") <= 0 ? 1 : Trim(value, "`n "))
    }
    InactiveClientBorderColor {
        get {
            if ( !This._JSON["_Profiles"][This.ProfileThumbnailsVisuals]["Thumbnails Visuals"].Has("InactiveClientBorderColor") )
                This._JSON["_Profiles"][This.ProfileThumbnailsVisuals]["Thumbnails Visuals"]["InactiveClientBorderColor"] := "#8A8A8A"

             return convertToHex(This._JSON["_Profiles"][This.ProfileThumbnailsVisuals]["Thumbnails Visuals"]["InactiveClientBorderColor"])
        }
        set => This._JSON["_Profiles"][This.ProfileThumbnailsVisuals]["Thumbnails Visuals"]["InactiveClientBorderColor"] := convertToHex(Trim(value, "`n "))
    }


    ;########################
    ;## Profile ClientSettings

    DontCloseOnLoginScreen {
        get => This._JSON["_Profiles"][This.ProfileClientSettings]["Client Settings"]["DontCloseOnLoginScreen"]
        set => This._JSON["_Profiles"][This.ProfileClientSettings]["Client Settings"]["DontCloseOnLoginScreen"] := value
    }

    dontCloseActiveClient {
        get => This._JSON["_Profiles"][This.ProfileClientSettings]["Client Settings"]["dontCloseActiveClient"]
        set => This._JSON["_Profiles"][This.ProfileClientSettings]["Client Settings"]["dontCloseActiveClient"] := value
    }

    DontCloseClients {
        get => This._JSON["_Profiles"][This.ProfileClientSettings]["Client Settings"]["DontCloseClients"]
        set {
            This._JSON["_Profiles"][This.ProfileClientSettings]["Client Settings"]["DontCloseClients"] := []

            For index, Client in StrSplit(Value, ["`n", ","]) {
                if (Client = "")
                    continue
                This._JSON["_Profiles"][This.ProfileClientSettings]["Client Settings"]["DontCloseClients"].Push(This.AntiCleanTitle(Trim(Client, "`n ")))
            }
        }
    }

    CustomColorsGet[CName?] {
        get {
            name := "", nameIndex := 0, ctext := "", cBorder := "", cIABorder := ""
            for index, names in This._JSON["_Profiles"][This.ProfileCustomColors]["Custom Colors"]["cColors"]["CharNames"] {
                if (names = CName) {
                    nameIndex := index
                    name := names
                    break
                }
                else
                    nameIndex := index

            }
            if (nameIndex) {
                if (This._JSON["_Profiles"][This.ProfileCustomColors]["Custom Colors"]["cColors"]["Bordercolor"].Length >= nameIndex) {
                    cBorder := This._JSON["_Profiles"][This.ProfileCustomColors]["Custom Colors"]["cColors"]["Bordercolor"][nameIndex]
                }
                if (This._JSON["_Profiles"][This.ProfileCustomColors]["Custom Colors"]["cColors"]["TextColor"].Length >= nameIndex)
                    ctext := This._JSON["_Profiles"][This.ProfileCustomColors]["Custom Colors"]["cColors"]["TextColor"][nameIndex]
                if (This._JSON["_Profiles"][This.ProfileCustomColors]["Custom Colors"]["cColors"]["IABordercolor"].Length >= nameIndex)
                    cIABorder := This._JSON["_Profiles"][This.ProfileCustomColors]["Custom Colors"]["cColors"]["IABordercolor"][nameIndex]
            }
            return Map("Char", name, "Border", cBorder, "Text", ctext, "IABorder", cIABorder)
        }
    }


    IndexcChars => This._JSON["_Profiles"][This.ProfileCustomColors]["Custom Colors"]["cColors"]["CharNames"].Length
    IndexcBorder => This._JSON["_Profiles"][This.ProfileCustomColors]["Custom Colors"]["cColors"]["Bordercolor"].Length
    IndexcText => This._JSON["_Profiles"][This.ProfileCustomColors]["Custom Colors"]["cColors"]["TextColor"].Length
    IndexcIABorders => This._JSON["_Profiles"][This.ProfileCustomColors]["Custom Colors"]["cColors"]["IABordercolor"].Length

    CustomColors_AllCharNames {
        get {
            names := ""
            for k, v in This._JSON["_Profiles"][This.ProfileCustomColors]["Custom Colors"]["cColors"]["CharNames"] {
                if (A_Index < This._JSON["_Profiles"][This.ProfileCustomColors]["Custom Colors"]["cColors"]["CharNames"].Length)
                    names .= k ": " This.CleanTitle(v) "`n"
                else
                    names .= k ": " This.CleanTitle(v)
            }
            return names
        }
        set {
            tempvar := []
            ListChars := StrSplit(value, "`n")
            for k, v in ListChars {
                chars := This.AntiCleanTitle(RegExReplace(Trim(v, "`n "), ".*:\s*", ""))
                tempvar.Push(chars)
            }
            This._JSON["_Profiles"][This.ProfileCustomColors]["Custom Colors"]["cColors"]["CharNames"] := tempvar
        }
    }
    CustomColors_AllBColors {
        get {
            names := ""
            for k, v in This._JSON["_Profiles"][This.ProfileCustomColors]["Custom Colors"]["cColors"]["Bordercolor"] {
                if (A_Index < This._JSON["_Profiles"][This.ProfileCustomColors]["Custom Colors"]["cColors"]["Bordercolor"].Length)
                    names .= k ": " v "`n"
                else
                    names .= k ": " v
            }
            return names
        }
        set {
            tempvar := []
            ListChars := StrSplit(value, "`n")
            for k, v in ListChars {
                chars := RegExReplace(Trim(v, "`n "), ".*:\s*", "")
                tempvar.Push(convertToHex(chars))
            }
            This._JSON["_Profiles"][This.ProfileCustomColors]["Custom Colors"]["cColors"]["Bordercolor"] := tempvar
        }
    }
    CustomColors_AllTColors {
        get {
            names := ""
            for k, v in This._JSON["_Profiles"][This.ProfileCustomColors]["Custom Colors"]["cColors"]["TextColor"] {
                if (A_Index < This._JSON["_Profiles"][This.ProfileCustomColors]["Custom Colors"]["cColors"]["TextColor"].Length)
                    names .= k ": " v "`n"
                else
                    names .= k ": " v
            }
            return names
        }
        set {
            tempvar := []
            ListChars := StrSplit(value, "`n")
            for k, v in ListChars {
                chars := RegExReplace(Trim(v, "`n "), ".*:\s*", "")
                tempvar.Push(convertToHex(chars))
            }
            This._JSON["_Profiles"][This.ProfileCustomColors]["Custom Colors"]["cColors"]["TextColor"] := tempvar
        }
    }

    CustomColors_IABorder_Colors {
        get {
            names := ""
            if (!This._JSON["_Profiles"][This.ProfileCustomColors]["Custom Colors"]["cColors"].Has("IABordercolor")) {
                This._JSON["_Profiles"][This.ProfileCustomColors]["Custom Colors"]["cColors"]["IABordercolor"] := ["FFFFFF"]
                SetTimer(This.Save_Settings_Delay_Timer, -200)
            }
            for k, v in This._JSON["_Profiles"][This.ProfileCustomColors]["Custom Colors"]["cColors"]["IABordercolor"] {
                if (A_Index < This._JSON["_Profiles"][This.ProfileCustomColors]["Custom Colors"]["cColors"]["IABordercolor"].Length)
                    names .= k ": " v "`n"
                else
                    names .= k ": " v
            }
            return names
        }
        set {
            tempvar := []
            ListChars := StrSplit(value, "`n")
            for k, v in ListChars {
                chars := RegExReplace(Trim(v, "`n "), ".*:\s*", "")
                tempvar.Push(convertToHex(chars))
            }
            This._JSON["_Profiles"][This.ProfileCustomColors]["Custom Colors"]["cColors"]["IABordercolor"] := tempvar
        }
    }


    CustomColorsActive {
        get => This._JSON["_Profiles"][This.ProfileCustomColors]["Custom Colors"]["cColorActive"]
        set => This._JSON["_Profiles"][This.ProfileCustomColors]["Custom Colors"]["cColorActive"] := Value
    }

    Minimizeclients_Delay {
        get => This._JSON["_Profiles"][This.ProfileClientSettings]["Client Settings"]["Minimize_Delay"]
        set => This._JSON["_Profiles"][This.ProfileClientSettings]["Client Settings"]["Minimize_Delay"] := (value < 50 ? "50" : value)
    }
    MinimizeInactiveClients {
        get => This._JSON["_Profiles"][This.ProfileClientSettings]["Client Settings"]["MinimizeInactiveClients"]
        set => This._JSON["_Profiles"][This.ProfileClientSettings]["Client Settings"]["MinimizeInactiveClients"] := value
    }
    AlwaysMaximize {
        get => This._JSON["_Profiles"][This.ProfileClientSettings]["Client Settings"]["AlwaysMaximize"]
        set => This._JSON["_Profiles"][This.ProfileClientSettings]["Client Settings"]["AlwaysMaximize"] := value
    }
    TrackClientPossitions {
        get => This._JSON["_Profiles"][This.ProfileClientSettings]["Client Settings"]["TrackClientPossitions"]
        set => This._JSON["_Profiles"][This.ProfileClientSettings]["Client Settings"]["TrackClientPossitions"] := value
    }
    Dont_Minimize_Clients {
        get => This._JSON["_Profiles"][This.ProfileClientSettings]["Client Settings"]["Dont_Minimize_Clients"]
        set {
            This._JSON["_Profiles"][This.ProfileClientSettings]["Client Settings"]["Dont_Minimize_Clients"] := []

            For index, Client in StrSplit(Value, ["`n", ","]) {
                if (Client = "")
                    continue
                This._JSON["_Profiles"][This.ProfileClientSettings]["Client Settings"]["Dont_Minimize_Clients"].Push(This.AntiCleanTitle(Trim(Client, "`n ")))
            }
        }
    }

    ThumbnailPositions[wTitle?] {
        get {
            if (IsSet(wTitle))
                return This._JSON["_Profiles"][This.LastUsedProfile]["Thumbnail Positions"][wTitle]
            return This._JSON["_Profiles"][This.LastUsedProfile]["Thumbnail Positions"]
        }
        set {
            form := ["x", "y", "width", "height"]

            if !(This._JSON["_Profiles"][This.LastUsedProfile]["Thumbnail Positions"].Has(wTitle))
                This._JSON["_Profiles"][This.LastUsedProfile]["Thumbnail Positions"][wTitle] := Map()

            for v in form {
                This._JSON["_Profiles"][This.LastUsedProfile]["Thumbnail Positions"][wTitle][v] := value[A_Index]
            }
            SetTimer(This.Save_Settings_Delay_Timer, -200)
        }

    }

    ClientPossitions[wTitle] {
        get {
            if (This._JSON["_Profiles"][This.LastUsedProfile]["Client Possitions"].Has(wTitle))
                return This._JSON["_Profiles"][This.LastUsedProfile]["Client Possitions"][wTitle]
            else
                return 0
        }
        set {
            form := ["x", "y", "width", "height", "IsMaximized"]
            if !(This._JSON["_Profiles"][This.LastUsedProfile]["Client Possitions"].Has(wTitle))
                This._JSON["_Profiles"][This.LastUsedProfile]["Client Possitions"][wTitle] := Map()
            for v in form {
                This._JSON["_Profiles"][This.LastUsedProfile]["Client Possitions"][wTitle][v] := value[A_Index]
            }

        }
    }

    ;#########
    ; Other
    SwitchLangOnErr {
        get => This._JSON["_Profiles"][This.ProfileOther]["Other"]["SwitchLangOnErr"]
        set => This._JSON["_Profiles"][This.ProfileOther]["Other"]["SwitchLangOnErr"] := value
    }

    ;########################
    ;## Profile Hotkeys

    Suspend_Hotkeys_Hotkey {
        get => This._JSON["_Profiles"][This.ProfileHotkeysSettings]["Hotkeys Settings"]["Suspend_Hotkeys_Hotkey"]
        set => This._JSON["_Profiles"][This.ProfileHotkeysSettings]["Hotkeys Settings"]["Suspend_Hotkeys_Hotkey"] := value
    }
    
    Global_Hotkeys {
        get => This._JSON["_Profiles"][This.ProfileHotkeysSettings]["Hotkeys Settings"]["Global_Hotkeys"]
        set => This._JSON["_Profiles"][This.ProfileHotkeysSettings]["Hotkeys Settings"]["Global_Hotkeys"] := value
    }
    
    Login_Screen_Cycle_Hotkey {
        get => This._JSON["_Profiles"][This.ProfileHotkeysSettings]["Hotkeys Settings"]["Login_Screen_Cycle_Hotkey"]
        set => This._JSON["_Profiles"][This.ProfileHotkeysSettings]["Hotkeys Settings"]["Login_Screen_Cycle_Hotkey"] := value
    }

    LoginScreenCycleDirection[*] {
        get => This._JSON["_Profiles"][This.ProfileHotkeysSettings]["Hotkeys Settings"]["LoginScreenCycleDirection"]
        set => This._JSON["_Profiles"][This.ProfileHotkeysSettings]["Hotkeys Settings"]["LoginScreenCycleDirection"] := Value
    }

    PreserveHotkeysOnLogout {
        get => This._JSON["_Profiles"][This.ProfileHotkeysGroups]["Hotkeys Settings"]["PreserveHotkeysOnLogout"]
        set => This._JSON["_Profiles"][This.ProfileHotkeysGroups]["Hotkeys Settings"]["PreserveHotkeysOnLogout"] := value
    }

    KeepGroupsPositions {
        get => This._JSON["_Profiles"][This.ProfileHotkeysGroups]["Hotkeys Settings"]["KeepGroupsPositions"]
        set => This._JSON["_Profiles"][This.ProfileHotkeysGroups]["Hotkeys Settings"]["KeepGroupsPositions"] := value
    }

    dynamicGroupsEnabled {
        get => This._JSON["_Profiles"][This.ProfileHotkeysGroups]["Hotkeys Settings"]["dynamicGroupsEnabled"]
        set => This._JSON["_Profiles"][This.ProfileHotkeysGroups]["Hotkeys Settings"]["dynamicGroupsEnabled"] := value
    }

    dynamicGroupsColor {
        get => convertToHex(This._JSON["_Profiles"][This.ProfileHotkeysGroups]["Hotkeys Settings"]["dynamicGroupsColor"])
        set => This._JSON["_Profiles"][This.ProfileHotkeysGroups]["Hotkeys Settings"]["dynamicGroupsColor"] := convertToHex(value)
    }

    Close_Active_EVE_Win_Hotkey {
        get => This._JSON["_Profiles"][This.ProfileHotkeysSettings]["Hotkeys Settings"]["Close_Active_EVE_Win_Hotkey"]
        set => This._JSON["_Profiles"][This.ProfileHotkeysSettings]["Hotkeys Settings"]["Close_Active_EVE_Win_Hotkey"] := value
    }

    Close_All_EVE_Win_Hotkey {
        get => This._JSON["_Profiles"][This.ProfileHotkeysSettings]["Hotkeys Settings"]["Close_All_EVE_Win_Hotkey"]
        set => This._JSON["_Profiles"][This.ProfileHotkeysSettings]["Hotkeys Settings"]["Close_All_EVE_Win_Hotkey"] := value
    }

    Reload_Program_Hotkey {
        get => This._JSON["_Profiles"][This.ProfileHotkeysSettings]["Hotkeys Settings"]["Reload_Program_Hotkey"]
        set => This._JSON["_Profiles"][This.ProfileHotkeysSettings]["Hotkeys Settings"]["Reload_Program_Hotkey"] := value
    }

    GroupsHoldDelay {
        get => This._JSON["_Profiles"][This.ProfileHotkeysGroups]["Hotkeys Settings"]["GroupsHoldDelay"]
        set => This._JSON["_Profiles"][This.ProfileHotkeysGroups]["Hotkeys Settings"]["GroupsHoldDelay"] := value
    }

    HideThumbnailsHotkey {
        get => This._JSON["_Profiles"][This.ProfileHotkeysSettings]["Hotkeys Settings"]["HideThumbnailsHotkey"]
        set => This._JSON["_Profiles"][This.ProfileHotkeysSettings]["Hotkeys Settings"]["HideThumbnailsHotkey"] := value
    }

    ClickThroughHotkey {
        get => This._JSON["_Profiles"][This.ProfileHotkeysSettings]["Hotkeys Settings"]["ClickThroughHotkey"]
        set => This._JSON["_Profiles"][This.ProfileHotkeysSettings]["Hotkeys Settings"]["ClickThroughHotkey"] := value
    }

    Hotkey_Groups[key?] {
        get {
            if (IsSet(key)) {
                return This._JSON["_Profiles"][This.LastUsedProfile]["Hotkey Groups"][key]
            }
            else
                return This._JSON["_Profiles"][This.LastUsedProfile]["Hotkey Groups"]
        }
        set {
            This._JSON["_Profiles"][This.LastUsedProfile]["Hotkey Groups"][Key] := Map("Characters", value, "ForwardsHotkey", "", "BackwardsHotkey", "")
        }
    }

    _Hotkeys[key?] {
        get {
            if (IsSet(Key)) {
                loop This._JSON["_Profiles"][This.LastUsedProfile]["Hotkeys Settings"]["CharacterHotkeys"].Length {
                    if (This._JSON["_Profiles"][This.LastUsedProfile]["Hotkeys Settings"]["CharacterHotkeys"][A_Index].Has(key)) {
                        return This._JSON["_Profiles"][This.LastUsedProfile]["Hotkeys Settings"]["CharacterHotkeys"][A_Index][key]
                    }
                }
                return 0
            }
            if !(IsSet(Key))
                return This._JSON["_Profiles"][This.LastUsedProfile]["Hotkeys Settings"]["CharacterHotkeys"]
        }
        set => This._JSON["_Profiles"][This.LastUsedProfile]["Hotkeys Settings"]["CharacterHotkeys"] := Value
    }

    ; Tray Menu Settings ##############################

    TrayMenuShortcuts {
        get => This._JSON["_Profiles"][This.ProfileTrayMenuSettings]["Tray Menu Settings"]["TrayMenuShortcuts"]
        set => This._JSON["_Profiles"][This.ProfileTrayMenuSettings]["Tray Menu Settings"]["TrayMenuShortcuts"] := value
    }

    ; Non-EVE Applications ############################
    
    NonEVEGroups {
        get {
            temp := This._JSON["_Profiles"][This.ProfileNonEVEApplications]["Non-EVE Applications"]["NonEVEGroups"]
            res := Map()
            for n, group in temp {
                if group["exe"] = []
                    continue
                for i, _ in group["exe"]
                    if !group["title"].Has(i)
                        group["title"].Push("") ; filling in array with empty titles
                res[n] := group
            }
            return res
        }
        set => This._JSON["_Profiles"][This.ProfileNonEVEApplications]["Non-EVE Applications"]["NonEVEGroups"] := value
    }

    GetNonEVEGroupsList() {
        res := []
        groups := This.NonEVEGroups
        for n, _ in groups
            res.Push(n)
        return res
    }

    NonEVEHotkeys {
        get {
            data := This._JSON["_Profiles"][This.ProfileNonEVEApplications]["Non-EVE Applications"]["NonEVEHotkeys"]
            if data["exe"] = [] || data["exe"] = [""]
                return Map("exe", [], "title", [], "hotkey", [])
            for i, _ in data["exe"] {
                if !data["title"].Has(i)
                    data["title"].Push("") ; filling in array with empty titles
                if !data["hotkey"].Has(i)
                    data["hotkey"].Push("") ; filling in array with empty hotkeys
            }
            return data
        }
        set => This._JSON["_Profiles"][This.ProfileNonEVEApplications]["Non-EVE Applications"]["NonEVEHotkeys"] := value
    }

    ; Game Logs Monitoring ##############################################

    monitoredEventsTexts := Map(
        "stoppedShooting", "Stopped Shooting",
        "underAttackByPlayer", "Under Attack By Player",
        "underAttackByNPC", "Under Attack By NPC",
        "engagedWithFactionBSNPC", "Engaged With Faction BS NPC",
        "engagedWithOfficerNPC", "Engaged With Officer NPC",
        "engagedWithCapitalNPC", "Engaged With Capital NPC",
        "warpDisrupted", "Warp Disrupted",
        "decloaked", "Decloaked",
        "gateJumped", "Gate Jumped",
        "convoRequest", "Convo Request",
        "fleetInvited", "Fleet Invited",
        "fleetWarped", "Fleet Warped",
        "fleetRegrouped", "Fleet Regrouped",
        "conduited", "Conduited",
        "crystalBroke", "Crystal Broke",
        "miningStopped", "Mining Stopped",
        "miningBayIsFull", "Mining Bay Is Full"
    )

    gameLogsMonitoringEnabled {
        get => This._JSON["_Profiles"][This.ProfileGameLogsMonitoring]["Game Logs Monitoring"]["gameLogsMonitoringEnabled"]
        set => This._JSON["_Profiles"][This.ProfileGameLogsMonitoring]["Game Logs Monitoring"]["gameLogsMonitoringEnabled"] := value
    }

    monitoringInterval {
        get => This._JSON["_Profiles"][This.ProfileGameLogsMonitoring]["Game Logs Monitoring"]["monitoringInterval"]
        set => This._JSON["_Profiles"][This.ProfileGameLogsMonitoring]["Game Logs Monitoring"]["monitoringInterval"] := value
    }

    gameLogsDirectory {
        get => This._JSON["_Profiles"][This.ProfileGameLogsMonitoring]["Game Logs Monitoring"]["gameLogsDirectory"]
        set => This._JSON["_Profiles"][This.ProfileGameLogsMonitoring]["Game Logs Monitoring"]["gameLogsDirectory"] := value
    }

    charsIds { ; This is only for default profile because we want to save as much char ids as possible
        get => This._JSON["_Profiles"]["Default"]["Game Logs Monitoring"]["charsIds"]
        set => This._JSON["_Profiles"]["Default"]["Game Logs Monitoring"]["charsIds"] := value
    }

    monitorOnlySelectedChars {
        get => This._JSON["_Profiles"][This.ProfileGameLogsMonitoring]["Game Logs Monitoring"]["monitorOnlySelectedChars"]
        set => This._JSON["_Profiles"][This.ProfileGameLogsMonitoring]["Game Logs Monitoring"]["monitorOnlySelectedChars"] := value
    }

    charsToMonitor {
        get => This._JSON["_Profiles"][This.ProfileGameLogsMonitoring]["Game Logs Monitoring"]["charsToMonitor"]
        set => This._JSON["_Profiles"][This.ProfileGameLogsMonitoring]["Game Logs Monitoring"]["charsToMonitor"] := value
    }

    lastEventPriority {
        get => This._JSON["_Profiles"][This.ProfileGameLogsMonitoring]["Game Logs Monitoring"]["lastEventPriority"]
        set => This._JSON["_Profiles"][This.ProfileGameLogsMonitoring]["Game Logs Monitoring"]["lastEventPriority"] := value
    }

    supressForFocused {
        get => This._JSON["_Profiles"][This.ProfileGameLogsMonitoring]["Game Logs Monitoring"]["supressForFocused"]
        set => This._JSON["_Profiles"][This.ProfileGameLogsMonitoring]["Game Logs Monitoring"]["supressForFocused"] := value
    }

    showEventText {
        get => This._JSON["_Profiles"][This.ProfileGameLogsMonitoring]["Game Logs Monitoring"]["showEventText"]
        set => This._JSON["_Profiles"][This.ProfileGameLogsMonitoring]["Game Logs Monitoring"]["showEventText"] := value
    }

    flashBorderEnabled {
        get => This._JSON["_Profiles"][This.ProfileGameLogsMonitoring]["Game Logs Monitoring"]["flashBorderEnabled"]
        set => This._JSON["_Profiles"][This.ProfileGameLogsMonitoring]["Game Logs Monitoring"]["flashBorderEnabled"] := value
    }

    stopDisplayingOnSwitch {
        get => This._JSON["_Profiles"][This.ProfileGameLogsMonitoring]["Game Logs Monitoring"]["stopDisplayingOnSwitch"]
        set => This._JSON["_Profiles"][This.ProfileGameLogsMonitoring]["Game Logs Monitoring"]["stopDisplayingOnSwitch"] := value
    }

    eventDisplayDuration {
        get => This._JSON["_Profiles"][This.ProfileGameLogsMonitoring]["Game Logs Monitoring"]["eventDisplayDuration"]
        set => This._JSON["_Profiles"][This.ProfileGameLogsMonitoring]["Game Logs Monitoring"]["eventDisplayDuration"] := value
    }

    flashBorderInterval {
        get => This._JSON["_Profiles"][This.ProfileGameLogsMonitoring]["Game Logs Monitoring"]["flashBorderInterval"]
        set => This._JSON["_Profiles"][This.ProfileGameLogsMonitoring]["Game Logs Monitoring"]["flashBorderInterval"] := value
    }

    monitoredEvents {
        get => This._JSON["_Profiles"][This.ProfileMonitoredEvents]["Game Logs Monitoring"]["monitoredEvents"]
        set => This._JSON["_Profiles"][This.ProfileMonitoredEvents]["Game Logs Monitoring"]["monitoredEvents"] := value
    }

    shootingInterval {
        get => This._JSON["_Profiles"][This.ProfileMonitoredEvents]["Game Logs Monitoring"]["shootingInterval"]
        set => This._JSON["_Profiles"][This.ProfileMonitoredEvents]["Game Logs Monitoring"]["shootingInterval"] := value
    }

    _Hotkey_Delete(*) {
        if (This.LV_Item) {
            try {
                HKey_Char_Name := This.LV.GetText(This.LV_Item)
                if (This._Hotkeys.Has(HKey_Char_Name)) {
                    This._Hotkeys.Delete(HKey_Char_Name)
                    This.LV.Delete(This.LV_Item)

                    ;This.Save_Settings()
                }
            }
        }
    }

    _Hotkey_Add(*) {
        Obj := InputBox("Enter the Char Name", "Add New Char", "w200 h90")
        if (Obj.Result = "OK") {
            This._Hotkeys[Trim(Obj.Value, " ")] := ""
            This.LV.Add(, Trim(Obj.Value, " "))

            ;This.Save_Settings()
        }
    }

    _Hotkey_Edit(*) {
        if (This.LV_Item) {
            HKey_Char_Key := This.LV.GetText(This.LV_Item, 2), HKey_Char_Name := This.LV.GetText(This.LV_Item)
            if (This._Hotkeys.Has(HKey_Char_Name)) {
                Obj := InputBox(HKey_Char_Key, "Edit Hotkey for -> " HKey_Char_Name, "w250 h100")
                if (Obj.Result = "OK") {
                    This._Hotkeys[HKey_Char_Name] := Trim(Obj.Value, " ")
                    This.LV.Modify(This.LV_Item, , , Trim(Obj.Value, " "))
                    This.LV.Modify(This.LV_Item, "+Focus +Select")

                    ;This.Save_Settings()
                }
            }
        }
    }


    _Tv_LVSelectedRow(GuiCtrlObj, Item, Checked) {
        Obj := Map()
        if (GuiCtrlObj == This.Tv_LV) {
            loop {
                RowNumber := This.Tv_LV.GetNext(A_Index - 1, "Checked")
                if not RowNumber  ; The above returned zero, so there are no more selected rows.
                    break

                Obj[This.Tv_LV.GetText(RowNumber)] := 1
                This.Thumbnail_visibility[This.Tv_LV.GetText(RowNumber)] := 1
                ;MsgBox(GuiCtrlObj.value)
            }
            This.Thumbnail_visibility := Obj
            SetTimer(This.Save_Settings_Delay_Timer, -200)
            This.NeedRestart := 1
            ;This.LV_Item := Item
            ; ddd := GuiCtrlObj.GetText(Item)
            ; ToolTip(Item ", " ddd " -, " Checked)
        }
    }


    _LVSelectedRow(GuiCtrlObj, Item, Selected) {
        if (GuiCtrlObj == This.LV && Selected) {
            This.LV_Item := Item
            ddd := GuiCtrlObj.GetText(Item)
            ;ToolTip(Item ", " ddd " -, " Selected)
        }
    }


    ;######################
    ;## Methods


    Suspend_Hotkeys(*) {
        static state := 0
        ToolTip()
        state := !state
        state ? ToolTip("Hotkeys disabled") : ToolTip("Hotkeys enabled")
        Suspend(-1)

        SetTimer((*) => ToolTip(), -1500)
    }

    Delete_Profile(*) {
        if (This.SelectProfile_DDL.Text = "Default") {
            MsgBox("You cannot delete the default settings")
            Return
        }

        if (This.SelectProfile_DDL.Text = This.LastUsedProfile) {
            This.LastUsedProfile := "Default"
        }

        This._JSON["_Profiles"].Delete(This.SelectProfile_DDL.Text)

        if (This.LastUsedProfile = "" || !This.Profiles.Has(This.LastUsedProfile))
            This.LastUsedProfile := "Default"

        This.ProfileOverride()
        ; FileDelete("EVE-X-Preview.json")
        ; FileAppend(JSON.Dump(This._JSON, , "    "), "EVE-X-Preview.json")
        SetTimer(This.Save_Settings_Delay_Timer, -200)

        ;Index := This.SelectProfile_DDL.Value
        This.SelectProfile_DDL.Delete(This.SelectProfile_DDL.Value)
        This.SelectProfile_DDL.Redraw()

        for k, v in This.MainFrame {
            v.Enabled := 0
        }

        ;This.S_Gui.Show("AutoSize")
    }

    Create_Profile(*) {
        Obj := InputBox("Enter a Profile Name", "Create New Profile", "w200 h90")
        if (Obj.Result != "OK" || Obj.Result = "")
            return
        if (This.Profiles.Has(Obj.value)) {
            MsgBox("A profile with this name already exists")
            return
        }
        if !(This.LastUsedProfile = "Default") {
            Result := MsgBox("Do you want to use the current settings for the new profile?", , "YesNo")
        }
        else
            Result := "No"

        if Result = "Yes"
            This._JSON["_Profiles"][Obj.value] := JSON.Load(FileRead("EVE-X-Preview.json"))["_Profiles"][This.LastUsedProfile]
        else if Result = "No"
            This._JSON["_Profiles"][Obj.value] := This.default_JSON["_Profiles"]["Default"]
        else
            Return 0

        FileDelete("EVE-X-Preview.json")
        FileAppend(JSON.Dump(This._JSON, , "    "), "EVE-X-Preview.json")
        This.SelectProfile_DDL.Delete()
        This.SelectProfile_DDL.Add(This.Profiles_to_Array())
        ControlChooseString(Obj.value, This.SelectProfile_DDL, "EVE-X-Preview - Settings")
        This.LastUsedProfile := Obj.value
        Return
    }

    Rename_Profile(*) {
        if This.LastUsedProfile == "Default" {
            MsgBox("Can't rename Default profile")
            return
        }

        Obj := InputBox("Enter a Profile Name", "Rename Profile", "w200 h90", This.LastUsedProfile)

        if (Obj.Result != "OK" || Obj.Result = "")
            return
        if (This.Profiles.Has(Obj.value)) {
            MsgBox("A profile with this name already exists")
            return
        }

        This._JSON["_Profiles"][Obj.value] := JSON.Load(FileRead("EVE-X-Preview.json"))["_Profiles"][This.LastUsedProfile]
        This._JSON["_Profiles"].Delete(This.LastUsedProfile)

        This.SelectProfile_DDL.Delete()
        This.SelectProfile_DDL.Add(This.Profiles_to_Array())
        This.LastUsedProfile := Obj.value
        This.ProfileOverride()
        ControlChooseString(Obj.value, This.SelectProfile_DDL, "EVE-X-Preview - Settings")
        Return
    }

    Save_ThumbnailPossitions() {
        for EvEHwnd, GuiObj in This.ThumbWindows.OwnProps() {
            for Names, Obj in GuiObj {
                if (Names = "Window" && Obj.Title = "" || Obj.Title = "EVE")
                    continue
                Else if (Names = "Window") {
                    WinGetPos(&wX, &wY, &wWidth, &wHeight, Obj.Hwnd)
                    This.ThumbnailPositions[Obj.Title] := [wX, wY, wWidth, wHeight]
                }
            }
        }
    }

    ;### Stores the Thumbnail Size and Possitions in the Json file
    Save_Settings() {
        for EvEHwnd, GuiObj in This.ThumbWindows.OwnProps() {
            for Names, Obj in GuiObj {
                if (Names = "Window" && Obj.Title = "" || Obj.Title = "EVE")
                    continue
                Else if (Names = "Window") {
                    WinGetPos(&wX, &wY, &wWidth, &wHeight, Obj.Hwnd)
                    This.ThumbnailPositions[Obj.Title] := [wX, wY, wWidth, wHeight]
                }
            }
        }
        SetTimer(This.Save_Settings_Delay_Timer, -200)
    }

    Update_All_Thumbnails() {
        thumbW := This.ThumbnailStartLocation["width"]
        thumbH := This.ThumbnailStartLocation["height"]

        for pName, pValue in This._JSON["_Profiles"] {
            for charName, v in pValue["Thumbnail Positions"] {
                v["width"] := thumbW
                v["height"] := thumbH
            }
        }

        SetTimer(This.Save_Settings_Delay_Timer, -200)
    }

    ImportNamesFromThumbs(EditField) {
        text := EditField.Value
        EditField.Value := ""
        charList := ""
        for EvEHwnd, ThumbObj in This.ThumbWindows.OwnProps() {
            for k, v in ThumbObj {
                if k = "Window" {
                    if v.Title == "EVE" || v.Title == "Char Screen"
                        continue
                    charList .= This.CleanTitle(v.Title) . "`n"
                }
            }
        }
        sorted := Sort(charList)
        text .= sorted
        ControlSendText(text, , EditField.Hwnd)
    }
}


;########################
;## Functions

; Add_New_Profile() {
;     return
; }

convertToHex(rgbString) {
    rgbString := Trim(rgbString, "`n ")
    ; Check if the string corresponds to the decimal value format (e.g. "255, 255, 255" or "rgb(255, 255, 255)")
    if (RegExMatch(rgbString, "^\s*(rgb\s*\(?)?\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)?\s*$", &matches)) {
        red := matches[2], green := matches[3], blue := matches[4]

        ; covert decimal to hex
        hexValue := Format("{:02X}{:02X}{:02X}", red, green, blue)
        return hexValue
    }

    ; Check whether the string corresponds to the hexadecimal value format (e.g "#FFFFFF" or "0xFFFFFF")
    if (RegExMatch(rgbString, "^\s*(#|0x)?([0-9A-Fa-f]{6})\s*$", &matches)) {
        hexValue := matches[2]
        hexValue := StrLower(hexValue)
        return hexValue
    }
    ;  If no match was found or the string is already in hexadecimal value format, return directly
    return rgbString
}
