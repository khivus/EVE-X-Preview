

class Propertys extends TrayMenu {


    ;######################
    ;## Script Propertys

    SetThumbnailText[hwnd, *] {
        set {
            previousCritical := A_IsCritical
            Critical
            try {
                if This.ThumbWindows.HasProp(hwnd) {
                    newtext := Value
                    if This.ThumbWindows.%hwnd%["Window"].Title != newtext {
                        This.HideEventBorder(hwnd)
                        This.updateThumbnailEventText("", hwnd)
                        This.updateThumbnailSystemText("", hwnd)
                        This.updateThumbnailDPSText("", hwnd)
                    }
                    This.updateThumbnailText(newtext, hwnd)
                    This.ThumbWindows.%hwnd%["Window"].Title := newtext
                }
            } finally {
                Critical(previousCritical)
            }
        }
    }

    Profiles => This._JSON["_Profiles"]

    ; Profile activation belongs to the profile itself, independent of group overrides.
    ProfileActivationHotkey {
        get => This.Profiles[This.LastUsedProfile].Get("ActivationHotkey", "")
        set => This.Profiles[This.LastUsedProfile]["ActivationHotkey"] := Trim(value, "`r`n `t")
    }


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

    SectionProfile(section, legacyGroup) {
        aliases := Map("Thumbnails Visuals", "ProfileThumbnailsVisuals", "Client Settings", "ProfileClientSettings",
            "Non-EVE Applications", "ProfileNonEVEApplications", "Game Logs Monitoring", "ProfileGameLogsMonitoring")
        property := aliases.Get(legacyGroup, "")
        if !This._JSON.Has("LastUsedProfile")
            return property != "" && This.HasOwnProp(property) ? This.%property% : "Default"
        profile := This.Profiles[This.LastUsedProfile]
        if profile.Get("SettingsSectionsVersion", 0) >= 1
            return This.LastUsedProfile != "Default" && profile["SettingsInheritance"].Get(section, 0) ? "Default" : This.LastUsedProfile
        if This.HasOwnProp("ComboGroups")
            return This.ComboGroups.Get(legacyGroup, 0) ? "Default" : This.LastUsedProfile
        return property != "" && This.HasOwnProp(property) ? This.%property% : This.LastUsedProfile
    }
    FieldProfile(id, legacyProfile) {
        if !This._JSON.Has("LastUsedProfile") || This.Profiles[This.LastUsedProfile].Get("SettingsSectionsVersion", 0) < 1
            return legacyProfile
        field := SettingsSchema.Fields[id]
        return This.SectionProfile(field.group, field.legacyGroup)
    }

    ProfileOverride() {
        This.ComboGroups := Map()
        if This.Profiles[This.LastUsedProfile].Has("SettingsInheritance") {
            for group in SettingsModel.LegacyGroups
                This.ComboGroups[group] := This.LastUsedProfile != "Default" && This.Profiles[This.LastUsedProfile]["SettingsInheritance"].Get(group, 0)
        } else {
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
        }

        This.ProfileFocusedSounds := This.SectionProfile("alert-behavior", "Focused-client sounds")
        This.ProfileFocusedDPS := This.SectionProfile("dps", "Focused-client DPS")
        This.ProfileCharacterShortcuts := This.SectionProfile("character-shortcuts", "Character shortcuts")
        This.ProfileCharacterGroups := This.SectionProfile("character-groups", "Character groups")
        This.ProfileHotkeysSettings := This.SectionProfile("shortcuts", "Hotkeys Settings")
        This.ProfileThumbnailsBehavior := This.SectionProfile("layout", "Thumbnails Behavior")
        This.ProfileThumbnailsInteractions := This.SectionProfile("mouse-actions", "Thumbnails Interactions")
        This.ProfileThumbnailsVisuals := This.SectionProfile("appearance", "Thumbnails Visuals")
        This.ProfileThumbnailVisibility := This.SectionProfile("hidden-characters", "Thumbnail Visibility")
        This.ProfileClientSettings := This.SectionProfile("window-behavior", "Client Settings")
        This.ProfileCustomColors := This.SectionProfile("character-appearance", "Custom Colors")
        This.ProfileGameLogsMonitoring := This.SectionProfile("sources", "Game Logs Monitoring")
        This.ProfileMonitoredEvents := This.SectionProfile("events", "Monitored Events")
        This.ProfileDPSMonitoring := This.SectionProfile("dps", "DPS Monitoring")
        This.ProfileTrayMenuSettings := This.SectionProfile("tray", "Tray Menu Settings")
        This.ProfileOther := This.SectionProfile("performance", "Other")
        This.ProfileHotkeysGroups := This.SectionProfile("switching-behavior", "Hotkey Groups")
        This.ProfileNonEVEApplications := This.SectionProfile("app-groups", "Non-EVE Applications")
    }


    ; ### Profile Thumbnails Interactions

    ThumbnailsInteractions {
        get => This._JSON["_Profiles"][This.ProfileThumbnailsInteractions]["Thumbnails Interactions"]
        set => This._JSON["_Profiles"][This.ProfileThumbnailsInteractions]["Thumbnails Interactions"] := value
    }
    
    DisableFromGroupsColor {
        get => convertToHex(This._JSON["_Profiles"][This.FieldProfile("DisableFromGroupsColor", This.ProfileThumbnailsInteractions)]["Hotkeys Settings"]["dynamicGroupsColor"])
        set => This._JSON["_Profiles"][This.FieldProfile("DisableFromGroupsColor", This.ProfileThumbnailsInteractions)]["Hotkeys Settings"]["dynamicGroupsColor"] := convertToHex(value)
    }

    QuickGroupColor {
        get => convertToHex(This._JSON["_Profiles"][This.FieldProfile("QuickGroupColor", This.ProfileThumbnailsInteractions)]["Hotkeys Settings"]["QuickGroupColor"])
        set => This._JSON["_Profiles"][This.FieldProfile("QuickGroupColor", This.ProfileThumbnailsInteractions)]["Hotkeys Settings"]["QuickGroupColor"] := convertToHex(value)
    }

    QuickGroupHotkey {
        get => This._JSON["_Profiles"][This.FieldProfile("QuickGroupHotkey", This.ProfileThumbnailsInteractions)]["Hotkeys Settings"]["QuickGroupHotkey"]
        set => This._JSON["_Profiles"][This.FieldProfile("QuickGroupHotkey", This.ProfileThumbnailsInteractions)]["Hotkeys Settings"]["QuickGroupHotkey"] := value
    }

    QuickGroupSortOrder {
        get => This._JSON["_Profiles"][This.FieldProfile("QuickGroupSortOrder", This.ProfileThumbnailsInteractions)]["Hotkeys Settings"]["QuickGroupSortOrder"]
        set => This._JSON["_Profiles"][This.FieldProfile("QuickGroupSortOrder", This.ProfileThumbnailsInteractions)]["Hotkeys Settings"]["QuickGroupSortOrder"] := value
    }

    QuickGroupIgnoredInOtherGroups {
        get => This._JSON["_Profiles"][This.FieldProfile("QuickGroupIgnoredInOtherGroups", This.ProfileThumbnailsInteractions)]["Hotkeys Settings"]["QuickGroupIgnoredInOtherGroups"]
        set => This._JSON["_Profiles"][This.FieldProfile("QuickGroupIgnoredInOtherGroups", This.ProfileThumbnailsInteractions)]["Hotkeys Settings"]["QuickGroupIgnoredInOtherGroups"] := value
    }

    QuickGroupResetsPosition {
        get => This._JSON["_Profiles"][This.FieldProfile("QuickGroupResetsPosition", This.ProfileThumbnailsInteractions)]["Hotkeys Settings"]["QuickGroupResetsPosition"]
        set => This._JSON["_Profiles"][This.FieldProfile("QuickGroupResetsPosition", This.ProfileThumbnailsInteractions)]["Hotkeys Settings"]["QuickGroupResetsPosition"] := value
    }

    DontCloseDisabledClients {
        get => This._JSON["_Profiles"][This.FieldProfile("DontCloseDisabledClients", This.ProfileThumbnailsInteractions)]["Hotkeys Settings"]["DontCloseDisabledClients"]
        set => This._JSON["_Profiles"][This.FieldProfile("DontCloseDisabledClients", This.ProfileThumbnailsInteractions)]["Hotkeys Settings"]["DontCloseDisabledClients"] := value
    }

    PreserveQuickGroupOnLogout {
        get => This._JSON["_Profiles"][This.FieldProfile("PreserveQuickGroupOnLogout", This.ProfileThumbnailsInteractions)]["Hotkeys Settings"].Get("PreserveQuickGroupOnLogout", 1)
        set => This._JSON["_Profiles"][This.FieldProfile("PreserveQuickGroupOnLogout", This.ProfileThumbnailsInteractions)]["Hotkeys Settings"]["PreserveQuickGroupOnLogout"] := value
    }

    PreserveDisabledOnLogout {
        get => This._JSON["_Profiles"][This.FieldProfile("PreserveDisabledOnLogout", This.ProfileThumbnailsInteractions)]["Hotkeys Settings"].Get("PreserveDisabledOnLogout", 1)
        set => This._JSON["_Profiles"][This.FieldProfile("PreserveDisabledOnLogout", This.ProfileThumbnailsInteractions)]["Hotkeys Settings"]["PreserveDisabledOnLogout"] := value
    }

    DontCloseQuickGroupClients {
        get => This._JSON["_Profiles"][This.FieldProfile("DontCloseQuickGroupClients", This.ProfileThumbnailsInteractions)]["Hotkeys Settings"]["DontCloseQuickGroupClients"]
        set => This._JSON["_Profiles"][This.FieldProfile("DontCloseQuickGroupClients", This.ProfileThumbnailsInteractions)]["Hotkeys Settings"]["DontCloseQuickGroupClients"] := value
    }


    ;########################
    ;## Profile ThumbnailSettings

    ThumbnailStartLocation[key] {
        get => This._JSON["_Profiles"][This.ProfileThumbnailsBehavior]["Thumbnails Behavior"]["ThumbnailStartLocation"][key]
        set => This._JSON["_Profiles"][This.ProfileThumbnailsBehavior]["Thumbnails Behavior"]["ThumbnailStartLocation"][key] := value
    }

    ThumbnailToolsSettings {
        get {
            if !This.HasOwnProp("ProfileThumbnailsBehavior")
                return Map()
            return This.Profiles[This.ProfileThumbnailsBehavior].Get("Thumbnails Behavior", Map())
        }
    }

    LockThumbnailPositions {
        get => This.ThumbnailToolsSettings.Get("LockThumbnailPositions", 0)
        set => This.ThumbnailToolsSettings["LockThumbnailPositions"] := !!value
    }

    AnonymizeThumbnails {
        get => This.Profiles[This.FieldProfile("AnonymizeThumbnails", This.ProfileThumbnailsVisuals)].Get("Thumbnails Visuals", Map()).Get("AnonymizeThumbnails", 0)
        set => This.Profiles[This.FieldProfile("AnonymizeThumbnails", This.ProfileThumbnailsVisuals)]["Thumbnails Visuals"]["AnonymizeThumbnails"] := !!value
    }

    AnonymousNames {
        get => This.Profiles[This.FieldProfile("AnonymousNames", This.ProfileThumbnailsVisuals)].Get("Thumbnails Visuals", Map()).Get("AnonymousNames", "Aurora`nComet`nNova`nOrbit`nRaven`nZenith")
        set => This.Profiles[This.FieldProfile("AnonymousNames", This.ProfileThumbnailsVisuals)]["Thumbnails Visuals"]["AnonymousNames"] := value
    }

    ThumbnailPollingInterval {
        get {
            interval := This.Profiles[This.FieldProfile("ThumbnailPollingInterval", This.ProfileThumbnailsBehavior)]["Thumbnails Behavior"].Get("ThumbnailPollingInterval", 50)
            return IsNumber(interval) ? Max(50, Min(1000, Round(interval))) : 50
        }
        set => This.Profiles[This.FieldProfile("ThumbnailPollingInterval", This.ProfileThumbnailsBehavior)]["Thumbnails Behavior"]["ThumbnailPollingInterval"] := value
    }

    AutoWildcardWindowHotkeys {
        get => This.HasOwnProp("ProfileHotkeysGroups") ? This.Profiles[This.FieldProfile("AutoWildcardWindowHotkeys", This.ProfileHotkeysGroups)]["Hotkeys Settings"].Get("AutoWildcardWindowHotkeys", 1) : 1
        set => This.Profiles[This.FieldProfile("AutoWildcardWindowHotkeys", This.ProfileHotkeysGroups)]["Hotkeys Settings"]["AutoWildcardWindowHotkeys"] := !!value
    }

    AutoSaveThumbnailPositions {
        get => This._JSON["_Profiles"][This.FieldProfile("AutoSaveThumbnailPositions", This.ProfileThumbnailsBehavior)]["Thumbnails Behavior"]["AutoSaveThumbnailPositions"]
        set => This._JSON["_Profiles"][This.FieldProfile("AutoSaveThumbnailPositions", This.ProfileThumbnailsBehavior)]["Thumbnails Behavior"]["AutoSaveThumbnailPositions"] := value
    }

    ThumbnailBackgroundColor {
        get => convertToHex(This._JSON["_Profiles"][This.FieldProfile("ThumbnailBackgroundColor", This.ProfileThumbnailsVisuals)]["Thumbnails Visuals"]["ThumbnailBackgroundColor"])
        set => This._JSON["_Profiles"][This.FieldProfile("ThumbnailBackgroundColor", This.ProfileThumbnailsVisuals)]["Thumbnails Visuals"]["ThumbnailBackgroundColor"] := convertToHex(value)
    }

    ThumbnailSnap[*] {
        get => This._JSON["_Profiles"][This.FieldProfile("ThumbnailSnap", This.ProfileThumbnailsBehavior)]["Thumbnails Behavior"]["ThumbnailSnap"]
        set => This._JSON["_Profiles"][This.FieldProfile("ThumbnailSnap", This.ProfileThumbnailsBehavior)]["Thumbnails Behavior"]["ThumbnailSnap"] := Value
    }

    ThumbnailSnap_Distance {
        get => This._JSON["_Profiles"][This.FieldProfile("ThumbnailSnap_Distance", This.ProfileThumbnailsBehavior)]["Thumbnails Behavior"]["ThumbnailSnap_Distance"]
        set => This._JSON["_Profiles"][This.FieldProfile("ThumbnailSnap_Distance", This.ProfileThumbnailsBehavior)]["Thumbnails Behavior"]["ThumbnailSnap_Distance"] := (value ? value : "20")
    }

    ThumbnailMinimumSize[key] {
        get => This._JSON["_Profiles"][This.SectionProfile("layout", "Thumbnails Visuals")]["Thumbnails Visuals"]["ThumbnailMinimumSize"][key]
        set => This._JSON["_Profiles"][This.SectionProfile("layout", "Thumbnails Visuals")]["Thumbnails Visuals"]["ThumbnailMinimumSize"][key] := value
    }

    PreserveThumbPosOnLogout {
        get => This._JSON["_Profiles"][This.FieldProfile("PreserveThumbPosOnLogout", This.ProfileThumbnailsBehavior)]["Thumbnails Behavior"]["PreserveThumbPosOnLogout"]
        set => This._JSON["_Profiles"][This.FieldProfile("PreserveThumbPosOnLogout", This.ProfileThumbnailsBehavior)]["Thumbnails Behavior"]["PreserveThumbPosOnLogout"] := value
    }

    PreserveCharNameOnLogout {
        get => This._JSON["_Profiles"][This.FieldProfile("PreserveCharNameOnLogout", This.ProfileThumbnailsBehavior)]["Thumbnails Behavior"]["PreserveCharNameOnLogout"]
        set => This._JSON["_Profiles"][This.FieldProfile("PreserveCharNameOnLogout", This.ProfileThumbnailsBehavior)]["Thumbnails Behavior"]["PreserveCharNameOnLogout"] := value
    }

    HideThumbForActiveWin {
        get => This._JSON["_Profiles"][This.FieldProfile("HideThumbForActiveWin", This.ProfileThumbnailsBehavior)]["Thumbnails Behavior"]["HideThumbForActiveWin"]
        set => This._JSON["_Profiles"][This.FieldProfile("HideThumbForActiveWin", This.ProfileThumbnailsBehavior)]["Thumbnails Behavior"]["HideThumbForActiveWin"] := value
    }

    ShiftThumbsForLoginScreen {
        get => This._JSON["_Profiles"][This.FieldProfile("ShiftThumbsForLoginScreen", This.ProfileThumbnailsBehavior)]["Thumbnails Behavior"]["ShiftThumbsForLoginScreen"]
        set => This._JSON["_Profiles"][This.FieldProfile("ShiftThumbsForLoginScreen", This.ProfileThumbnailsBehavior)]["Thumbnails Behavior"]["ShiftThumbsForLoginScreen"] := value
    }

    ShiftThumbsForLoggedInCharacters {
        get => This._JSON["_Profiles"][This.FieldProfile("ShiftThumbsForLoggedInCharacters", This.ProfileThumbnailsBehavior)]["Thumbnails Behavior"].Get("ShiftThumbsForLoggedInCharacters", 1)
        set => This._JSON["_Profiles"][This.FieldProfile("ShiftThumbsForLoggedInCharacters", This.ProfileThumbnailsBehavior)]["Thumbnails Behavior"]["ShiftThumbsForLoggedInCharacters"] := value
    }

    ShiftThumbsCollisionCheck {
        get => This._JSON["_Profiles"][This.FieldProfile("ShiftThumbsCollisionCheck", This.ProfileThumbnailsBehavior)]["Thumbnails Behavior"]["ShiftThumbsCollisionCheck"]
        set => This._JSON["_Profiles"][This.FieldProfile("ShiftThumbsCollisionCheck", This.ProfileThumbnailsBehavior)]["Thumbnails Behavior"]["ShiftThumbsCollisionCheck"] := value
    }

    ShiftThumbsDirection {
        get => This._JSON["_Profiles"][This.FieldProfile("ShiftThumbsDirection", This.ProfileThumbnailsBehavior)]["Thumbnails Behavior"]["ShiftThumbsDirection"]
        set => This._JSON["_Profiles"][This.FieldProfile("ShiftThumbsDirection", This.ProfileThumbnailsBehavior)]["Thumbnails Behavior"]["ShiftThumbsDirection"] := value
    }

    ShiftThumbHorizontalStep {
        get => This._JSON["_Profiles"][This.FieldProfile("ShiftThumbHorizontalStep", This.ProfileThumbnailsBehavior)]["Thumbnails Behavior"]["ShiftThumbHorizontalStep"]
        set => This._JSON["_Profiles"][This.FieldProfile("ShiftThumbHorizontalStep", This.ProfileThumbnailsBehavior)]["Thumbnails Behavior"]["ShiftThumbHorizontalStep"] := value
    }

    ShiftThumbVerticalStep {
        get => This._JSON["_Profiles"][This.FieldProfile("ShiftThumbVerticalStep", This.ProfileThumbnailsBehavior)]["Thumbnails Behavior"]["ShiftThumbVerticalStep"]
        set => This._JSON["_Profiles"][This.FieldProfile("ShiftThumbVerticalStep", This.ProfileThumbnailsBehavior)]["Thumbnails Behavior"]["ShiftThumbVerticalStep"] := value
    }

    HideThumbnails {
        get => This._JSON["_Profiles"][This.FieldProfile("HideThumbnails", This.ProfileThumbnailsBehavior)]["Thumbnails Behavior"]["HideThumbnails"]
        set => This._JSON["_Profiles"][This.FieldProfile("HideThumbnails", This.ProfileThumbnailsBehavior)]["Thumbnails Behavior"]["HideThumbnails"] := value
    }
    
    ClickThroughActive {
        get => This._JSON["_Profiles"][This.FieldProfile("ClickThroughActive", This.ProfileThumbnailsBehavior)]["Thumbnails Behavior"]["ClickThroughActive"]
        set => This._JSON["_Profiles"][This.FieldProfile("ClickThroughActive", This.ProfileThumbnailsBehavior)]["Thumbnails Behavior"]["ClickThroughActive"] := value
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
        get => This._JSON["_Profiles"][This.FieldProfile("HideThumbnailsOnLostFocus", This.ProfileThumbnailsBehavior)]["Thumbnails Behavior"]["HideThumbnailsOnLostFocus"]
        set => This._JSON["_Profiles"][This.FieldProfile("HideThumbnailsOnLostFocus", This.ProfileThumbnailsBehavior)]["Thumbnails Behavior"]["HideThumbnailsOnLostFocus"] := value
    }
    ExcludeThumbnailsFromCapture {
        get => This._JSON["_Profiles"][This.FieldProfile("ExcludeThumbnailsFromCapture", This.ProfileThumbnailsBehavior)]["Thumbnails Behavior"].Get("ExcludeThumbnailsFromCapture", 0)
        set => This._JSON["_Profiles"][This.FieldProfile("ExcludeThumbnailsFromCapture", This.ProfileThumbnailsBehavior)]["Thumbnails Behavior"]["ExcludeThumbnailsFromCapture"] := value
    }

    ShowThumbnailsAlwaysOnTop {
        get => This._JSON["_Profiles"][This.FieldProfile("ShowThumbnailsAlwaysOnTop", This.ProfileThumbnailsBehavior)]["Thumbnails Behavior"]["ShowThumbnailsAlwaysOnTop"]
        set => This._JSON["_Profiles"][This.FieldProfile("ShowThumbnailsAlwaysOnTop", This.ProfileThumbnailsBehavior)]["Thumbnails Behavior"]["ShowThumbnailsAlwaysOnTop"] := value
    }

    ThumbnailOpacity {
        get {
            percentage := This._JSON["_Profiles"][This.FieldProfile("ThumbnailOpacity", This.ProfileThumbnailsVisuals)]["Thumbnails Visuals"]["ThumbnailOpacity"]
            return Round((percentage < 0 ? 0 : percentage > 100 ? 100 : percentage) * 2.55)
        }
        set {
            This._JSON["_Profiles"][This.FieldProfile("ThumbnailOpacity", This.ProfileThumbnailsVisuals)]["Thumbnails Visuals"]["ThumbnailOpacity"] := Value
        }
    }

    ClientHighligtBorderthickness {
        get => This._JSON["_Profiles"][This.FieldProfile("ClientHighligtBorderthickness", This.ProfileThumbnailsVisuals)]["Thumbnails Visuals"]["ClientHighligtBorderthickness"]
        set => This._JSON["_Profiles"][This.FieldProfile("ClientHighligtBorderthickness", This.ProfileThumbnailsVisuals)]["Thumbnails Visuals"]["ClientHighligtBorderthickness"] := (Trim(value, "`n ") <= 0 ? 1 : Trim(value, "`n "))
    }

    ClientHighligtColor {
        get => convertToHex(This._JSON["_Profiles"][This.FieldProfile("ClientHighligtColor", This.ProfileThumbnailsVisuals)]["Thumbnails Visuals"]["ClientHighligtColor"])
        set => This._JSON["_Profiles"][This.FieldProfile("ClientHighligtColor", This.ProfileThumbnailsVisuals)]["Thumbnails Visuals"]["ClientHighligtColor"] := convertToHex(Trim(value, "`n "))
    }
    ShowClientHighlightBorder {
        get => This._JSON["_Profiles"][This.FieldProfile("ShowClientHighlightBorder", This.ProfileThumbnailsVisuals)]["Thumbnails Visuals"]["ShowClientHighlightBorder"]
        set => This._JSON["_Profiles"][This.FieldProfile("ShowClientHighlightBorder", This.ProfileThumbnailsVisuals)]["Thumbnails Visuals"]["ShowClientHighlightBorder"] := value
    }

    EventActiveBorderMode {
        get {
            mode := This.Profiles[This.FieldProfile("EventActiveBorderMode", This.ProfileThumbnailsVisuals)]["Thumbnails Visuals"].Get("EventActiveBorderMode", 1)
            return mode = 2 ? 2 : (mode = 3 ? 3 : 1)
        }
        set => This.Profiles[This.FieldProfile("EventActiveBorderMode", This.ProfileThumbnailsVisuals)]["Thumbnails Visuals"]["EventActiveBorderMode"] := value = 2 ? 2 : (value = 3 ? 3 : 1)
    }
    ThumbnailTextFont {
        get => This._JSON["_Profiles"][This.FieldProfile("ThumbnailTextFont", This.ProfileThumbnailsVisuals)]["Thumbnails Visuals"]["ThumbnailTextFont"]
        set => This._JSON["_Profiles"][This.FieldProfile("ThumbnailTextFont", This.ProfileThumbnailsVisuals)]["Thumbnails Visuals"]["ThumbnailTextFont"] := Trim(value, "`n ")
    }
    CustomThumbnailNames {
        get {
            visuals := This._JSON["_Profiles"][This.SectionProfile("labels", "Thumbnails Visuals")]["Thumbnails Visuals"]
            if !visuals.Has("CustomThumbnailNames")
                visuals["CustomThumbnailNames"] := Map("Characters", "", "Names", "")
            return visuals["CustomThumbnailNames"]
        }
        set => This._JSON["_Profiles"][This.SectionProfile("labels", "Thumbnails Visuals")]["Thumbnails Visuals"]["CustomThumbnailNames"] := value
    }

    ThumbnailTextSize {
        get => This._JSON["_Profiles"][This.FieldProfile("ThumbnailTextSize", This.ProfileThumbnailsVisuals)]["Thumbnails Visuals"]["ThumbnailTextSize"]
        set => This._JSON["_Profiles"][This.FieldProfile("ThumbnailTextSize", This.ProfileThumbnailsVisuals)]["Thumbnails Visuals"]["ThumbnailTextSize"] := Trim(value, "`n ")
    }

    ThumbnailTextColor {
        get => convertToHex(This._JSON["_Profiles"][This.FieldProfile("ThumbnailTextColor", This.ProfileThumbnailsVisuals)]["Thumbnails Visuals"]["ThumbnailTextColor"])
        set => This._JSON["_Profiles"][This.FieldProfile("ThumbnailTextColor", This.ProfileThumbnailsVisuals)]["Thumbnails Visuals"]["ThumbnailTextColor"] := convertToHex(Trim(value, "`n "))
    }
    ShowThumbnailTextOverlay {
        get => This._JSON["_Profiles"][This.FieldProfile("ShowThumbnailTextOverlay", This.ProfileThumbnailsVisuals)]["Thumbnails Visuals"]["ShowThumbnailTextOverlay"]
        set => This._JSON["_Profiles"][This.FieldProfile("ShowThumbnailTextOverlay", This.ProfileThumbnailsVisuals)]["Thumbnails Visuals"]["ShowThumbnailTextOverlay"] := value
    }
    ThumbnailTextMargins[var] {
        get => This._JSON["_Profiles"][This.ProfileThumbnailsVisuals]["Thumbnails Visuals"]["ThumbnailTextMargins"][var]
        set => This._JSON["_Profiles"][This.ProfileThumbnailsVisuals]["Thumbnails Visuals"]["ThumbnailTextMargins"][var] := Trim(value, "`n ")
    }
    InactiveClientBorderthickness {
        get {
            if ( !This._JSON["_Profiles"][This.FieldProfile("InactiveClientBorderthickness", This.ProfileThumbnailsVisuals)]["Thumbnails Visuals"].Has("InactiveClientBorderthickness") )
                This._JSON["_Profiles"][This.FieldProfile("InactiveClientBorderthickness", This.ProfileThumbnailsVisuals)]["Thumbnails Visuals"]["InactiveClientBorderthickness"] := "2"
            return This._JSON["_Profiles"][This.FieldProfile("InactiveClientBorderthickness", This.ProfileThumbnailsVisuals)]["Thumbnails Visuals"]["InactiveClientBorderthickness"]
        } 
        set => This._JSON["_Profiles"][This.FieldProfile("InactiveClientBorderthickness", This.ProfileThumbnailsVisuals)]["Thumbnails Visuals"]["InactiveClientBorderthickness"] := (Trim(value, "`n ") <= 0 ? 1 : Trim(value, "`n "))
    }
    InactiveClientBorderColor {
        get {
            if ( !This._JSON["_Profiles"][This.FieldProfile("InactiveClientBorderColor", This.ProfileThumbnailsVisuals)]["Thumbnails Visuals"].Has("InactiveClientBorderColor") )
                This._JSON["_Profiles"][This.FieldProfile("InactiveClientBorderColor", This.ProfileThumbnailsVisuals)]["Thumbnails Visuals"]["InactiveClientBorderColor"] := "#8A8A8A"

             return convertToHex(This._JSON["_Profiles"][This.FieldProfile("InactiveClientBorderColor", This.ProfileThumbnailsVisuals)]["Thumbnails Visuals"]["InactiveClientBorderColor"])
        }
        set => This._JSON["_Profiles"][This.FieldProfile("InactiveClientBorderColor", This.ProfileThumbnailsVisuals)]["Thumbnails Visuals"]["InactiveClientBorderColor"] := convertToHex(Trim(value, "`n "))
    }


    ;########################
    ;## Profile ClientSettings

    DontCloseOnLoginScreen {
        get => This._JSON["_Profiles"][This.FieldProfile("DontCloseOnLoginScreen", This.ProfileClientSettings)]["Client Settings"]["DontCloseOnLoginScreen"]
        set => This._JSON["_Profiles"][This.FieldProfile("DontCloseOnLoginScreen", This.ProfileClientSettings)]["Client Settings"]["DontCloseOnLoginScreen"] := value
    }

    dontCloseActiveClient {
        get => This._JSON["_Profiles"][This.FieldProfile("dontCloseActiveClient", This.ProfileClientSettings)]["Client Settings"]["dontCloseActiveClient"]
        set => This._JSON["_Profiles"][This.FieldProfile("dontCloseActiveClient", This.ProfileClientSettings)]["Client Settings"]["dontCloseActiveClient"] := value
    }

    DontCloseClients {
        get => This._JSON["_Profiles"][This.SectionProfile("close-protection", "Client Settings")]["Client Settings"]["DontCloseClients"]
        set {
            This._JSON["_Profiles"][This.SectionProfile("close-protection", "Client Settings")]["Client Settings"]["DontCloseClients"] := []

            For index, Client in StrSplit(Value, ["`n", ","]) {
                if (Client = "")
                    continue
                This._JSON["_Profiles"][This.SectionProfile("close-protection", "Client Settings")]["Client Settings"]["DontCloseClients"].Push(This.AntiCleanTitle(Trim(Client, "`n ")))
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
        get => This._JSON["_Profiles"][This.FieldProfile("Minimizeclients_Delay", This.ProfileClientSettings)]["Client Settings"]["Minimize_Delay"]
        set => This._JSON["_Profiles"][This.FieldProfile("Minimizeclients_Delay", This.ProfileClientSettings)]["Client Settings"]["Minimize_Delay"] := (value < 50 ? "50" : value)
    }
    MinimizeInactiveClients {
        get => This._JSON["_Profiles"][This.FieldProfile("MinimizeInactiveClients", This.ProfileClientSettings)]["Client Settings"]["MinimizeInactiveClients"]
        set => This._JSON["_Profiles"][This.FieldProfile("MinimizeInactiveClients", This.ProfileClientSettings)]["Client Settings"]["MinimizeInactiveClients"] := value
    }
    AlwaysMaximize {
        get => This._JSON["_Profiles"][This.FieldProfile("AlwaysMaximize", This.ProfileClientSettings)]["Client Settings"]["AlwaysMaximize"]
        set => This._JSON["_Profiles"][This.FieldProfile("AlwaysMaximize", This.ProfileClientSettings)]["Client Settings"]["AlwaysMaximize"] := value
    }
    TrackClientPossitions {
        get => This._JSON["_Profiles"][This.FieldProfile("TrackClientPossitions", This.ProfileClientSettings)]["Client Settings"]["TrackClientPossitions"]
        set => This._JSON["_Profiles"][This.FieldProfile("TrackClientPossitions", This.ProfileClientSettings)]["Client Settings"]["TrackClientPossitions"] := value
    }
    Dont_Minimize_Clients {
        get => This._JSON["_Profiles"][This.SectionProfile("close-protection", "Client Settings")]["Client Settings"]["Dont_Minimize_Clients"]
        set {
            This._JSON["_Profiles"][This.SectionProfile("close-protection", "Client Settings")]["Client Settings"]["Dont_Minimize_Clients"] := []

            For index, Client in StrSplit(Value, ["`n", ","]) {
                if (Client = "")
                    continue
                This._JSON["_Profiles"][This.SectionProfile("close-protection", "Client Settings")]["Client Settings"]["Dont_Minimize_Clients"].Push(This.AntiCleanTitle(Trim(Client, "`n ")))
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
        get => This._JSON["_Profiles"][This.FieldProfile("SwitchLangOnErr", This.ProfileOther)]["Other"]["SwitchLangOnErr"]
        set => This._JSON["_Profiles"][This.FieldProfile("SwitchLangOnErr", This.ProfileOther)]["Other"]["SwitchLangOnErr"] := value
    }

    ;########################
    ;## Profile Hotkeys

    Suspend_Hotkeys_Hotkey {
        get => This._JSON["_Profiles"][This.FieldProfile("Suspend_Hotkeys_Hotkey", This.ProfileHotkeysSettings)]["Hotkeys Settings"]["Suspend_Hotkeys_Hotkey"]
        set => This._JSON["_Profiles"][This.FieldProfile("Suspend_Hotkeys_Hotkey", This.ProfileHotkeysSettings)]["Hotkeys Settings"]["Suspend_Hotkeys_Hotkey"] := value
    }
    
    Global_Hotkeys {
        get => This._JSON["_Profiles"][This.ProfileHotkeysSettings]["Hotkeys Settings"]["Global_Hotkeys"]
        set => This._JSON["_Profiles"][This.ProfileHotkeysSettings]["Hotkeys Settings"]["Global_Hotkeys"] := value
    }

    SwitchToPreviousWindow_Hotkey {
        get => This._JSON["_Profiles"][This.FieldProfile("SwitchToPreviousWindow_Hotkey", This.ProfileHotkeysSettings)]["Hotkeys Settings"]["SwitchToPreviousWindow_Hotkey"]
        set => This._JSON["_Profiles"][This.FieldProfile("SwitchToPreviousWindow_Hotkey", This.ProfileHotkeysSettings)]["Hotkeys Settings"]["SwitchToPreviousWindow_Hotkey"] := value
    }

    CycleEveryLoggedIn_Hotkey {
        get => This._JSON["_Profiles"][This.FieldProfile("CycleEveryLoggedIn_Hotkey", This.ProfileHotkeysSettings)]["Hotkeys Settings"]["CycleEveryLoggedIn_Hotkey"]
        set => This._JSON["_Profiles"][This.FieldProfile("CycleEveryLoggedIn_Hotkey", This.ProfileHotkeysSettings)]["Hotkeys Settings"]["CycleEveryLoggedIn_Hotkey"] := value
    }
    
    Login_Screen_Cycle_Hotkey {
        get => This._JSON["_Profiles"][This.FieldProfile("Login_Screen_Cycle_Hotkey", This.ProfileHotkeysSettings)]["Hotkeys Settings"]["Login_Screen_Cycle_Hotkey"]
        set => This._JSON["_Profiles"][This.FieldProfile("Login_Screen_Cycle_Hotkey", This.ProfileHotkeysSettings)]["Hotkeys Settings"]["Login_Screen_Cycle_Hotkey"] := value
    }

    LoginScreenCycleDirection[*] {
        get => This._JSON["_Profiles"][This.FieldProfile("LoginScreenCycleDirection", This.ProfileHotkeysSettings)]["Hotkeys Settings"]["LoginScreenCycleDirection"]
        set => This._JSON["_Profiles"][This.FieldProfile("LoginScreenCycleDirection", This.ProfileHotkeysSettings)]["Hotkeys Settings"]["LoginScreenCycleDirection"] := Value
    }

    PreserveHotkeysOnLogout {
        get => This._JSON["_Profiles"][This.FieldProfile("PreserveHotkeysOnLogout", This.ProfileHotkeysGroups)]["Hotkeys Settings"]["PreserveHotkeysOnLogout"]
        set => This._JSON["_Profiles"][This.FieldProfile("PreserveHotkeysOnLogout", This.ProfileHotkeysGroups)]["Hotkeys Settings"]["PreserveHotkeysOnLogout"] := value
    }

    KeepGroupsPositions {
        get => This._JSON["_Profiles"][This.FieldProfile("KeepGroupsPositions", This.ProfileHotkeysGroups)]["Hotkeys Settings"]["KeepGroupsPositions"]
        set => This._JSON["_Profiles"][This.FieldProfile("KeepGroupsPositions", This.ProfileHotkeysGroups)]["Hotkeys Settings"]["KeepGroupsPositions"] := value
    }

    Close_Active_EVE_Win_Hotkey {
        get => This._JSON["_Profiles"][This.FieldProfile("Close_Active_EVE_Win_Hotkey", This.ProfileHotkeysSettings)]["Hotkeys Settings"]["Close_Active_EVE_Win_Hotkey"]
        set => This._JSON["_Profiles"][This.FieldProfile("Close_Active_EVE_Win_Hotkey", This.ProfileHotkeysSettings)]["Hotkeys Settings"]["Close_Active_EVE_Win_Hotkey"] := value
    }

    Close_All_EVE_Win_Hotkey {
        get => This._JSON["_Profiles"][This.FieldProfile("Close_All_EVE_Win_Hotkey", This.ProfileHotkeysSettings)]["Hotkeys Settings"]["Close_All_EVE_Win_Hotkey"]
        set => This._JSON["_Profiles"][This.FieldProfile("Close_All_EVE_Win_Hotkey", This.ProfileHotkeysSettings)]["Hotkeys Settings"]["Close_All_EVE_Win_Hotkey"] := value
    }

    Reload_Program_Hotkey {
        get => This._JSON["_Profiles"][This.FieldProfile("Reload_Program_Hotkey", This.ProfileHotkeysSettings)]["Hotkeys Settings"]["Reload_Program_Hotkey"]
        set => This._JSON["_Profiles"][This.FieldProfile("Reload_Program_Hotkey", This.ProfileHotkeysSettings)]["Hotkeys Settings"]["Reload_Program_Hotkey"] := value
    }

    ToggleLivePreviewsHotkey {
        get => This._JSON["_Profiles"][This.FieldProfile("ToggleLivePreviewsHotkey", This.ProfileHotkeysSettings)]["Hotkeys Settings"]["ToggleLivePreviewsHotkey"]
        set => This._JSON["_Profiles"][This.FieldProfile("ToggleLivePreviewsHotkey", This.ProfileHotkeysSettings)]["Hotkeys Settings"]["ToggleLivePreviewsHotkey"] := value
    }

    Exit_Program_Hotkey {
        get => This._JSON["_Profiles"][This.FieldProfile("Exit_Program_Hotkey", This.ProfileHotkeysSettings)]["Hotkeys Settings"]["Exit_Program_Hotkey"]
        set => This._JSON["_Profiles"][This.FieldProfile("Exit_Program_Hotkey", This.ProfileHotkeysSettings)]["Hotkeys Settings"]["Exit_Program_Hotkey"] := value
    }

    GroupsHoldDelay {
        get => This._JSON["_Profiles"][This.FieldProfile("GroupsHoldDelay", This.ProfileHotkeysGroups)]["Hotkeys Settings"]["GroupsHoldDelay"]
        set => This._JSON["_Profiles"][This.FieldProfile("GroupsHoldDelay", This.ProfileHotkeysGroups)]["Hotkeys Settings"]["GroupsHoldDelay"] := value
    }
    
    MaxActiveWindowRetries {
        get => This._JSON["_Profiles"][This.FieldProfile("MaxActiveWindowRetries", This.ProfileHotkeysGroups)]["Hotkeys Settings"]["MaxActiveWindowRetries"]
        set => This._JSON["_Profiles"][This.FieldProfile("MaxActiveWindowRetries", This.ProfileHotkeysGroups)]["Hotkeys Settings"]["MaxActiveWindowRetries"] := value
    }
    
    ActiveWindowRetryInterval {
        get => This._JSON["_Profiles"][This.FieldProfile("ActiveWindowRetryInterval", This.ProfileHotkeysGroups)]["Hotkeys Settings"]["ActiveWindowRetryInterval"]
        set => This._JSON["_Profiles"][This.FieldProfile("ActiveWindowRetryInterval", This.ProfileHotkeysGroups)]["Hotkeys Settings"]["ActiveWindowRetryInterval"] := value
    }

    HideThumbnailsHotkey {
        get => This._JSON["_Profiles"][This.FieldProfile("HideThumbnailsHotkey", This.ProfileHotkeysSettings)]["Hotkeys Settings"]["HideThumbnailsHotkey"]
        set => This._JSON["_Profiles"][This.FieldProfile("HideThumbnailsHotkey", This.ProfileHotkeysSettings)]["Hotkeys Settings"]["HideThumbnailsHotkey"] := value
    }

    ClickThroughHotkey {
        get => This._JSON["_Profiles"][This.FieldProfile("ClickThroughHotkey", This.ProfileHotkeysSettings)]["Hotkeys Settings"]["ClickThroughHotkey"]
        set => This._JSON["_Profiles"][This.FieldProfile("ClickThroughHotkey", This.ProfileHotkeysSettings)]["Hotkeys Settings"]["ClickThroughHotkey"] := value
    }

    CharacterGroupsProfile => This.HasOwnProp("ProfileCharacterGroups") ? This.ProfileCharacterGroups : This.LastUsedProfile
    CharacterShortcutsProfile => This.HasOwnProp("ProfileCharacterShortcuts") ? This.ProfileCharacterShortcuts : This.LastUsedProfile

    Hotkey_Groups[key?] {
        get {
            if IsSet(key)
                return This._JSON["_Profiles"][This.CharacterGroupsProfile]["Hotkey Groups"][key]
            else
                return This._JSON["_Profiles"][This.CharacterGroupsProfile]["Hotkey Groups"]
        }
        set {
            This._JSON["_Profiles"][This.CharacterGroupsProfile]["Hotkey Groups"][Key] := Map("Characters", value, "ForwardsHotkey", "", "BackwardsHotkey", "", "FirstCharHotkey", "")
        }
    }

    _Hotkeys[key?] {
        get {
            if (IsSet(Key)) {
                loop This._JSON["_Profiles"][This.CharacterShortcutsProfile]["Hotkeys Settings"]["CharacterHotkeys"].Length {
                    if (This._JSON["_Profiles"][This.CharacterShortcutsProfile]["Hotkeys Settings"]["CharacterHotkeys"][A_Index].Has(key)) {
                        return This._JSON["_Profiles"][This.CharacterShortcutsProfile]["Hotkeys Settings"]["CharacterHotkeys"][A_Index][key]
                    }
                }
                return 0
            }
            if !(IsSet(Key))
                return This._JSON["_Profiles"][This.CharacterShortcutsProfile]["Hotkeys Settings"]["CharacterHotkeys"]
        }
        set => This._JSON["_Profiles"][This.CharacterShortcutsProfile]["Hotkeys Settings"]["CharacterHotkeys"] := Value
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
            data := This._JSON["_Profiles"][This.SectionProfile("app-shortcuts", "Non-EVE Applications")]["Non-EVE Applications"]["NonEVEHotkeys"]
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
        set => This._JSON["_Profiles"][This.SectionProfile("app-shortcuts", "Non-EVE Applications")]["Non-EVE Applications"]["NonEVEHotkeys"] := value
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
        "undockedFromNPCStation", "Undocked from NPC Station",
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
        get => This._JSON["_Profiles"][This.FieldProfile("gameLogsMonitoringEnabled", This.ProfileGameLogsMonitoring)]["Game Logs Monitoring"]["gameLogsMonitoringEnabled"]
        set => This._JSON["_Profiles"][This.FieldProfile("gameLogsMonitoringEnabled", This.ProfileGameLogsMonitoring)]["Game Logs Monitoring"]["gameLogsMonitoringEnabled"] := value
    }

    systemTrackingEnabled {
        get => This._JSON["_Profiles"][This.FieldProfile("systemTrackingEnabled", This.ProfileGameLogsMonitoring)]["Game Logs Monitoring"].Get("systemTrackingEnabled", 0)
        set => This._JSON["_Profiles"][This.FieldProfile("systemTrackingEnabled", This.ProfileGameLogsMonitoring)]["Game Logs Monitoring"]["systemTrackingEnabled"] := value
    }

    systemColorsEnabled {
        get => This._JSON["_Profiles"][This.FieldProfile("systemColorsEnabled", This.ProfileGameLogsMonitoring)]["Game Logs Monitoring"].Get("systemColorsEnabled", 0)
        set {
            This._JSON["_Profiles"][This.FieldProfile("systemColorsEnabled", This.ProfileGameLogsMonitoring)]["Game Logs Monitoring"]["systemColorsEnabled"] := value
            This.RefreshThumbnailSystemColors()
        }
    }

    SlowThumbnailCreation {
        get => This._JSON["_Profiles"][This.FieldProfile("SlowThumbnailCreation", This.ProfileOther)]["Other"].Get("SlowThumbnailCreation", 0)
        set => This._JSON["_Profiles"][This.FieldProfile("SlowThumbnailCreation", This.ProfileOther)]["Other"]["SlowThumbnailCreation"] := value
    }

    AutoPauseOnDwmFailure {
        get => This._JSON["_Profiles"][This.FieldProfile("AutoPauseOnDwmFailure", This.ProfileOther)]["Other"].Get("AutoPauseOnDwmFailure", 0)
        set => This._JSON["_Profiles"][This.FieldProfile("AutoPauseOnDwmFailure", This.ProfileOther)]["Other"]["AutoPauseOnDwmFailure"] := value
    }

    chatLogsDirectory {
        get => This._JSON["_Profiles"][This.FieldProfile("chatLogsDirectory", This.ProfileGameLogsMonitoring)]["Game Logs Monitoring"].Get("chatLogsDirectory", "")
        set => This._JSON["_Profiles"][This.FieldProfile("chatLogsDirectory", This.ProfileGameLogsMonitoring)]["Game Logs Monitoring"]["chatLogsDirectory"] := value
    }

    monitoringInterval {
        get => This._JSON["_Profiles"][This.FieldProfile("monitoringInterval", This.ProfileGameLogsMonitoring)]["Game Logs Monitoring"]["monitoringInterval"]
        set => This._JSON["_Profiles"][This.FieldProfile("monitoringInterval", This.ProfileGameLogsMonitoring)]["Game Logs Monitoring"]["monitoringInterval"] := value
    }

    gameLogsDirectory {
        get => This._JSON["_Profiles"][This.FieldProfile("gameLogsDirectory", This.ProfileGameLogsMonitoring)]["Game Logs Monitoring"]["gameLogsDirectory"]
        set => This._JSON["_Profiles"][This.FieldProfile("gameLogsDirectory", This.ProfileGameLogsMonitoring)]["Game Logs Monitoring"]["gameLogsDirectory"] := value
    }

    charsIds { ; This is only for default profile because we want to save as much char ids as possible
        get => This._JSON["_Profiles"]["Default"]["Game Logs Monitoring"]["charsIds"]
        set => This._JSON["_Profiles"]["Default"]["Game Logs Monitoring"]["charsIds"] := value
    }

    monitorOnlySelectedChars {
        get => This._JSON["_Profiles"][This.FieldProfile("monitorOnlySelectedChars", This.ProfileGameLogsMonitoring)]["Game Logs Monitoring"]["monitorOnlySelectedChars"]
        set => This._JSON["_Profiles"][This.FieldProfile("monitorOnlySelectedChars", This.ProfileGameLogsMonitoring)]["Game Logs Monitoring"]["monitorOnlySelectedChars"] := value
    }

    charsToMonitor {
        get => This._JSON["_Profiles"][This.SectionProfile("monitored-characters", "Game Logs Monitoring")]["Game Logs Monitoring"]["charsToMonitor"]
        set => This._JSON["_Profiles"][This.SectionProfile("monitored-characters", "Game Logs Monitoring")]["Game Logs Monitoring"]["charsToMonitor"] := value
    }

    lastEventPriority {
        get => This._JSON["_Profiles"][This.FieldProfile("lastEventPriority", This.ProfileGameLogsMonitoring)]["Game Logs Monitoring"]["lastEventPriority"]
        set => This._JSON["_Profiles"][This.FieldProfile("lastEventPriority", This.ProfileGameLogsMonitoring)]["Game Logs Monitoring"]["lastEventPriority"] := value
    }

    supressForFocused {
        get => This._JSON["_Profiles"][This.FieldProfile("supressForFocused", This.ProfileGameLogsMonitoring)]["Game Logs Monitoring"]["supressForFocused"]
        set => This._JSON["_Profiles"][This.FieldProfile("supressForFocused", This.ProfileGameLogsMonitoring)]["Game Logs Monitoring"]["supressForFocused"] := value
    }

    FocusedSoundsProfile => This.HasOwnProp("ProfileFocusedSounds") ? This.ProfileFocusedSounds : This.ProfileGameLogsMonitoring
    FocusedDPSProfile => This.HasOwnProp("ProfileFocusedDPS") ? This.ProfileFocusedDPS : This.ProfileDPSMonitoring

    SuppressFocusedEventSounds {
        get => This._JSON["_Profiles"][This.FocusedSoundsProfile]["Game Logs Monitoring"].Get("suppressFocusedEventSounds", This.supressForFocused || This.HideThumbForActiveWin)
        set => This._JSON["_Profiles"][This.FocusedSoundsProfile]["Game Logs Monitoring"]["suppressFocusedEventSounds"] := !!value
    }

    HideDPSForFocused {
        get => This._JSON["_Profiles"][This.FocusedDPSProfile]["DPS Monitoring"].Get("hideDPSForFocused", This.supressForFocused)
        set => This._JSON["_Profiles"][This.FocusedDPSProfile]["DPS Monitoring"]["hideDPSForFocused"] := !!value
    }

    showEventText {
        get => This._JSON["_Profiles"][This.FieldProfile("showEventText", This.ProfileGameLogsMonitoring)]["Game Logs Monitoring"]["showEventText"]
        set => This._JSON["_Profiles"][This.FieldProfile("showEventText", This.ProfileGameLogsMonitoring)]["Game Logs Monitoring"]["showEventText"] := value
    }

    flashBorderEnabled {
        get => This._JSON["_Profiles"][This.FieldProfile("flashBorderEnabled", This.ProfileGameLogsMonitoring)]["Game Logs Monitoring"]["flashBorderEnabled"]
        set => This._JSON["_Profiles"][This.FieldProfile("flashBorderEnabled", This.ProfileGameLogsMonitoring)]["Game Logs Monitoring"]["flashBorderEnabled"] := value
    }

    stopDisplayingOnSwitch {
        get => This._JSON["_Profiles"][This.FieldProfile("stopDisplayingOnSwitch", This.ProfileGameLogsMonitoring)]["Game Logs Monitoring"]["stopDisplayingOnSwitch"]
        set => This._JSON["_Profiles"][This.FieldProfile("stopDisplayingOnSwitch", This.ProfileGameLogsMonitoring)]["Game Logs Monitoring"]["stopDisplayingOnSwitch"] := value
    }

    eventDisplayDuration {
        get => This._JSON["_Profiles"][This.FieldProfile("eventDisplayDuration", This.ProfileGameLogsMonitoring)]["Game Logs Monitoring"]["eventDisplayDuration"]
        set => This._JSON["_Profiles"][This.FieldProfile("eventDisplayDuration", This.ProfileGameLogsMonitoring)]["Game Logs Monitoring"]["eventDisplayDuration"] := value
    }

    flashBorderInterval {
        get => This._JSON["_Profiles"][This.FieldProfile("flashBorderInterval", This.ProfileGameLogsMonitoring)]["Game Logs Monitoring"]["flashBorderInterval"]
        set => This._JSON["_Profiles"][This.FieldProfile("flashBorderInterval", This.ProfileGameLogsMonitoring)]["Game Logs Monitoring"]["flashBorderInterval"] := value
    }

    monitoredEvents {
        get => This._JSON["_Profiles"][This.ProfileMonitoredEvents]["Game Logs Monitoring"]["monitoredEvents"]
        set => This._JSON["_Profiles"][This.ProfileMonitoredEvents]["Game Logs Monitoring"]["monitoredEvents"] := value
    }

    shootingInterval {
        get => This._JSON["_Profiles"][This.FieldProfile("shootingInterval", This.ProfileMonitoredEvents)]["Game Logs Monitoring"]["shootingInterval"]
        set => This._JSON["_Profiles"][This.FieldProfile("shootingInterval", This.ProfileMonitoredEvents)]["Game Logs Monitoring"]["shootingInterval"] := value
    }

    dpsMonitoring {
        get => This._JSON["_Profiles"][This.ProfileDPSMonitoring]["DPS Monitoring"]
        set => This._JSON["_Profiles"][This.ProfileDPSMonitoring]["DPS Monitoring"] := value
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
        
        Result := MsgBox("Do you want to use the current settings for the new profile?", "Copy Settings?", "YesNo")

        if Result = "Yes" {
            This._JSON["_Profiles"][Obj.value] := DeepClone(This.Profiles[This.LastUsedProfile])

            if This.LastUsedProfile = "Default" ; Reset global groups from default profile
                for k, v in This._JSON["_Profiles"][Obj.value]["Other"]["Global_Groups"]
                    This._JSON["_Profiles"][Obj.value]["Other"]["Global_Groups"][k] := 0
        }
        else if Result = "No"
            This._JSON["_Profiles"][Obj.value] := DeepClone(This.default_JSON["_Profiles"]["Default"])
        else
            Return 0

        ; Copy settings without assigning the source profile's shortcut to the new profile.
        This.Profiles[Obj.value]["ActivationHotkey"] := ""

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

        This._JSON["_Profiles"][Obj.value] := This.Profiles[This.LastUsedProfile]
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
        eveWins := WinGetList("ahk_exe exefile.exe")

        for hwnd in eveWins {
            title := WinGetTitle("ahk_id " hwnd)
            if title != "EVE"
                charList .= This.CleanTitle(title) "`n"
        }

        sorted := Sort(charList)
        text .= sorted
        ControlSendText(text,, EditField.Hwnd)
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
