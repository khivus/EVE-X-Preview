; Labels and validation are independent of the legacy JSON keys. Reorganizing
; pages never moves a user's stored values or changes their profile ownership.
class SettingsSchema {
    static Fields := Map()
    static Sections := []
    static Pages := ["Thumbnails", "Switching", "EVE windows", "Monitoring & alerts", "Profiles", "Other applications", "Application"]

    static __New() {
        this.Section("Thumbnails", "appearance", "Appearance", "Thumbnails Visuals")
        this.Field("ShowThumbnailTextOverlay", "Show text on thumbnails", "bool")
        this.Field("ThumbnailTextFont", "Label font", "font",,,, "Choose an installed font. The test thumbnail previews your changes.")
        this.Field("ThumbnailTextSize", "Label font size", "integer", 1, 200, "pt")
        this.Field("ThumbnailTextColor", "Label color", "color")
        this.Field("ThumbnailTextMarginsx", "Horizontal label inset", "integer", 0, 16384, "px",, ["Thumbnails Visuals", "ThumbnailTextMargins", "x"])
        this.Field("ThumbnailTextMarginsy", "Vertical label inset", "integer", 0, 16384, "px",, ["Thumbnails Visuals", "ThumbnailTextMargins", "y"])
        this.Field("ThumbnailOpacity", "Thumbnail opacity", "slider", 0, 100, "%", "0% makes thumbnails invisible. Use the tray menu or Settings to recover.")
        this.Field("ThumbnailBackgroundColor", "Background color", "color")
        this.Field("ShowClientHighlightBorder", "Highlight the focused client's thumbnail", "bool")
        this.Field("ClientHighligtColor", "Focused-client border color", "color")
        this.Field("ClientHighligtBorderthickness", "Active and event border thickness", "integer", 1, 100, "px")
        this.Field("EventActiveBorderMode", "Active and event border layout", "choice",,,, "The diagram below shows which border is inside. The test thumbnail uses your actual colors.")
        this.Choices("EventActiveBorderMode", "Event inside, active outside", 1, "Active inside, event outside", 2, "Combined border", 3)
        this.Field("ShowAllBorders", "Show inactive and group borders", "bool",,,,, ["Thumbnails Visuals", "ShowAllColoredBorders"])
        this.Field("InactiveClientBorderColor", "Inactive border color", "color")
        this.Field("InactiveClientBorderthickness", "Inactive border thickness", "integer", 1, 100, "px")
        this.Depends("InactiveClientBorderColor", "ShowAllBorders")
        this.Depends("InactiveClientBorderthickness", "ShowAllBorders")
        for id in ["ThumbnailTextFont", "ThumbnailTextSize", "ThumbnailTextColor", "ThumbnailTextMarginsx", "ThumbnailTextMarginsy"]
            this.Depends(id, "ShowThumbnailTextOverlay")

        this.Section("Thumbnails", "labels", "Labels", "Thumbnails Visuals", "Custom names and anonymous labels only change the thumbnail text.")
        this.Field("AnonymizeThumbnails", "Use anonymous thumbnail labels", "bool")
        this.Field("AnonymousNames", "Random label pool", "multiline",,,, "One name per line. Empty or exhausted pools use Pilot plus a number.")
        this.Depends("AnonymousNames", "AnonymizeThumbnails")
        this.Section("Thumbnails", "character-appearance", "Character appearance", "Custom Colors", "Each row keeps a character and all of its colors together.")
        this.Field("Ccoloractive", "Use character-specific colors", "bool",,,,, ["Custom Colors", "cColorActive"])

        this.Section("Thumbnails", "layout", "Layout", "Thumbnails Behavior")
        this.Field("AutoSaveThumbnailPositions", "Remember thumbnail positions and sizes automatically", "bool")
        this.Field("LockThumbnailPositions", "Lock thumbnail positions and sizes", "bool")
        this.Field("ThumbnailStartLocationx", "Default horizontal screen position", "integer", -131072, 131072, "px", "Negative coordinates support monitors to the left.", ["Thumbnails Behavior", "ThumbnailStartLocation", "x"])
        this.Field("ThumbnailStartLocationy", "Default vertical screen position", "integer", -131072, 131072, "px", "Negative coordinates support monitors above.", ["Thumbnails Behavior", "ThumbnailStartLocation", "y"])
        this.Field("ThumbnailStartLocationwidth", "Default thumbnail width", "integer", 1, 16384, "px",, ["Thumbnails Behavior", "ThumbnailStartLocation", "width"])
        this.Field("ThumbnailStartLocationheight", "Default thumbnail height", "integer", 1, 16384, "px",, ["Thumbnails Behavior", "ThumbnailStartLocation", "height"])
        this.Field("ThumbnailMinimumSizewidth", "Minimum thumbnail width", "integer", 1, 16384, "px",, ["Thumbnails Visuals", "ThumbnailMinimumSize", "width"], "Thumbnails Visuals")
        this.Field("ThumbnailMinimumSizeheight", "Minimum thumbnail height", "integer", 1, 16384, "px",, ["Thumbnails Visuals", "ThumbnailMinimumSize", "height"], "Thumbnails Visuals")
        this.Field("ThumbnailSnap", "Snap to nearby thumbnail corners", "bool")
        this.Field("ThumbnailSnap_Distance", "Snapping distance", "integer", 1, 1000, "px")
        this.Depends("ThumbnailSnap_Distance", "ThumbnailSnap")
        this.Field("PreserveThumbPosOnLogout", "Keep thumbnail placement after logout", "bool")

        this.Section("Thumbnails", "placement", "Automatic placement", "Thumbnails Behavior")
        this.Field("ShiftThumbsForLoginScreen", "Arrange login-screen thumbnails automatically", "bool")
        this.Field("ShiftThumbsForLoggedInCharacters", "Arrange logged-in thumbnails without a saved position", "bool")
        this.Field("ShiftThumbsCollisionCheck", "Avoid occupied thumbnail positions", "bool")
        this.Field("ShiftThumbsDirection", "Placement direction", "choice")
        this.Choices("ShiftThumbsDirection", "Left to right; top to bottom", 1, "Left to right; bottom to top", 2,
            "Right to left; top to bottom", 3, "Right to left; bottom to top", 4,
            "Top to bottom; left to right", 5, "Top to bottom; right to left", 6,
            "Bottom to top; left to right", 7, "Bottom to top; right to left", 8)
        this.Field("ShiftThumbHorizontalStep", "Horizontal placement step", "integer", 0, 16384, "px", "0 = use each thumbnail's width.")
        this.Field("ShiftThumbVerticalStep", "Vertical placement step", "integer", 0, 16384, "px", "0 = use each thumbnail's height.")

        this.Section("Thumbnails", "visibility", "Visibility", "Thumbnails Behavior")
        this.Field("HideThumbnails", "Hide all thumbnails", "bool")
        this.Field("HideThumbnailsOnLostFocus", "Hide thumbnails while other applications are focused", "bool")
        this.Field("HideThumbForActiveWin", "Hide the focused client's thumbnail", "bool",,,, "This controls visibility. Focused-client sound preferences are in Alert behavior.")
        this.Field("ShowThumbnailsAlwaysOnTop", "Keep thumbnails above other windows", "bool")
        this.Field("ExcludeThumbnailsFromCapture", "Exclude thumbnails from screenshots and screen sharing", "bool",,,, "Windows capture exclusion depends on the recording application.")
        this.Field("PreserveCharNameOnLogout", "Keep the last character's label after logout", "bool")
        this.Section("Thumbnails", "hidden-characters", "Hidden characters", "Thumbnail Visibility", "Characters listed here have no visible thumbnail. Other characters remain visible.")
        this.Section("Thumbnails", "mouse-actions", "Mouse actions", "Thumbnails Interactions", "Choose a complete gesture for each action. Duplicate gestures are reported before Apply.")
        this.Field("ClickThroughActive", "Pass mouse clicks through thumbnails", "bool",,,, "Configure a click-through shortcut under Switching, or turn this off again in Settings.", ["Thumbnails Behavior", "ClickThroughActive"], "Thumbnails Behavior")

        this.Section("Switching", "shortcuts", "Keyboard shortcuts", "Hotkeys Settings", "Record a shortcut, clear it, or edit AutoHotkey syntax with Advanced. Conflicts are shown below.")
        for id, label in HotkeyBindings.Settings
            this.Field(id, label, "hotkey")
        this.Field("Hotkey_Scoope", "EVE switching shortcuts are available", "choice",,,,, ["Hotkeys Settings", "Global_Hotkeys"])
        this.Choices("Hotkey_Scoope", "Whenever an EVE window is open", 1, "Only while EVE is focused", 0)
        this.Field("LoginScreenCycleDirection", "Login-screen switching order", "choice",,,, "Order login-screen windows by when their EVE clients were started.")
        this.Choices("LoginScreenCycleDirection", "Oldest client first", 1, "Newest client first", 0)
        this.Field("AutoWildcardWindowHotkeys", "Allow extra modifier keys while using switching shortcuts", "bool",,,, "Equivalent to an automatic * prefix; saved shortcut text is preserved.", ["Hotkeys Settings", "AutoWildcardWindowHotkeys"], "Hotkey Groups")
        this.Section("Switching", "character-shortcuts", "Character shortcuts", "Character shortcuts", "Each row pairs a character with its shortcut.")
        this.Section("Switching", "character-groups", "Character groups", "Character groups", "Members are switched in row order. Use Move up and Move down to reorder them.")
        this.Section("Switching", "quick-group", "Quick Group", "Thumbnails Interactions", "Use the Quick Group mouse action to add or remove characters.", "Hotkeys Settings")
        this.Field("QuickGroupHotkey", "Cycle the Quick Group", "hotkey")
        this.Field("QuickGroupSortOrder", "Switching order", "choice")
        this.Choices("QuickGroupSortOrder", "Order characters were added", "Added order", "Character name", "Name")
        this.Field("QuickGroupColor", "Quick Group border color", "color")
        this.Field("QuickGroupIgnoredInOtherGroups", "Skip Quick Group members in other groups", "bool")
        this.Field("QuickGroupResetsPosition", "Restart other groups after using the Quick Group", "bool")
        this.Field("PreserveQuickGroupOnLogout", "Keep Quick Group membership after logout", "bool")
        this.Field("PreserveDisabledOnLogout", "Keep characters excluded from groups after logout", "bool")
        this.Field("DisableFromGroupsColor", "Excluded-character border color", "color",,,,, ["Hotkeys Settings", "dynamicGroupsColor"])
        this.Section("Switching", "switching-behavior", "Switching behavior", "Hotkey Groups",, "Hotkeys Settings")
        this.Field("PreserveHotkeysOnLogout", "Keep character shortcuts linked to clients after logout", "bool")
        this.Field("KeepGroupsPositions", "Remember the switching position in each group", "bool")
        this.Field("GroupsHoldDelay", "Minimum interval while holding a group shortcut", "integer", 75, 10000, "ms")
        this.Section("Switching", "advanced-switching", "Advanced switching", "Hotkey Groups", "Focus retry settings are intended for troubleshooting delayed window activation.", "Hotkeys Settings")
        this.Field("MaxActiveWindowRetries", "Maximum focus checks", "integer", 1, 100, "checks", "1 = do not wait for focus confirmation.")
        this.Field("ActiveWindowRetryInterval", "Time between focus checks", "integer", 1, 1000, "ms")

        this.Section("EVE windows", "window-behavior", "Activation and placement", "Client Settings")
        this.Field("MinimizeInactiveClients", "Minimize other EVE clients after switching", "bool")
        this.Field("Minimizeclients_Delay", "Delay before minimizing", "integer", 50, 60000, "ms",, ["Client Settings", "Minimize_Delay"])
        this.Depends("Minimizeclients_Delay", "MinimizeInactiveClients")
        this.Field("AlwaysMaximize", "Maximize EVE clients when activated", "bool")
        this.Field("TrackClientPossitions", "Restore saved EVE window positions and sizes", "bool")
        this.Section("EVE windows", "close-protection", "Close protection", "Client Settings", "These exceptions protect clients from Close all EVE windows.")
        this.Field("DontCloseOnLoginScreen", "Protect clients at the login screen", "bool")
        this.Field("dontCloseActiveClient", "Protect the focused client", "bool")
        this.Field("DontCloseDisabledClients", "Protect characters excluded from groups", "bool",,,,, ["Hotkeys Settings", "DontCloseDisabledClients"], "Thumbnails Interactions")
        this.Field("DontCloseQuickGroupClients", "Protect Quick Group members", "bool",,,,, ["Hotkeys Settings", "DontCloseQuickGroupClients"], "Thumbnails Interactions")


        this.Section("Monitoring & alerts", "sources", "Log sources", "Game Logs Monitoring", "Automatic detection is used when a folder field is blank. Monitoring changes are applied with the rest of the draft.")
        this.Field("gameLogsMonitoringEnabled", "Read game logs for alerts and DPS", "bool")
        this.Field("gameLogsDirectory", "Game-log folder", "directory",,,, "Leave blank for automatic detection.")
        this.Field("chatLogsDirectory", "Chat-log folder", "directory",,,, "Used by solar-system tracking; leave blank for automatic detection.")
        this.Field("monitoringInterval", "Time between log scans", "integer", 50, 60000, "ms", "Longer intervals reduce scanning and delay alerts.")
        this.Field("monitorOnlySelectedChars", "Monitor only selected characters", "bool")
        this.Section("Monitoring & alerts", "monitored-characters", "Monitored characters", "Game Logs Monitoring", "This list filters alerts and DPS when Monitor only selected characters is enabled.")
        this.Section("Monitoring & alerts", "systems", "Solar systems", "Game Logs Monitoring")
        this.Field("systemTrackingEnabled", "Show each character's solar system from Local chat logs", "bool")
        this.Field("systemColorsEnabled", "Use a consistent color for each solar system", "bool")
        this.Depends("systemColorsEnabled", "systemTrackingEnabled")
        this.Section("Monitoring & alerts", "alert-behavior", "Alert behavior", "Game Logs Monitoring")
        this.Field("lastEventPriority", "When another event arrives", "choice")
        this.Choices("lastEventPriority", "Replace the current alert with the latest event", 1, "Keep the first alert until it clears", 0)
        this.Field("supressForFocused", "Hide visual alerts for the focused client", "bool",,,, "Sound and DPS preferences are independent.")
        this.Field("suppressFocusedEventSounds", "Mute event sounds for the focused client", "bool",,,, "This does not hide visual alerts or DPS.", ["Game Logs Monitoring", "suppressFocusedEventSounds"], "Focused-client sounds")
        this.Field("showEventText", "Show event text on thumbnails", "bool")
        this.Field("flashBorderEnabled", "Flash the border during an alert", "bool")
        this.Field("flashBorderInterval", "Border flashing interval", "integer", 50, 60000, "ms")
        this.Depends("flashBorderInterval", "flashBorderEnabled")
        this.Field("stopDisplayingOnSwitch", "Clear visual alerts when switching to their client", "bool")
        this.Field("eventDisplayDuration", "Alert duration", "integer", 100, 3600000, "ms")
        this.Section("Monitoring & alerts", "events", "Event alerts", "Monitored Events", "Select an event to edit its conditions, color, sound and volume. Events are grouped by purpose.", "Game Logs Monitoring")
        this.Field("shootingInterval", "Shooting idle interval", "integer", 100, 3600000, "ms", "Used only by the Shooting idle alert.")
        this.Section("Monitoring & alerts", "dps", "DPS", "DPS Monitoring")
        this.Field("incomingDPSEnabled", "Show incoming DPS", "bool")
        this.Field("incomingDPSThreshold", "Minimum incoming DPS to display", "number", 0, 1000000000000, "DPS", "0 = show any incoming DPS.")
        this.Field("showIncomingDPSResistances", "Show estimated incoming damage mix", "bool",,,, "Percentages estimate EM, thermal, kinetic and explosive damage. ? marks unknown damage.")
        this.Depends("incomingDPSThreshold", "incomingDPSEnabled")
        this.Depends("showIncomingDPSResistances", "incomingDPSEnabled")
        this.Field("outgoingDPSEnabled", "Show outgoing DPS", "bool")
        this.Field("outgoingDPSThreshold", "Minimum outgoing DPS to display", "number", 0, 1000000000000, "DPS", "0 = show any outgoing DPS.")
        this.Depends("outgoingDPSThreshold", "outgoingDPSEnabled")
        this.Field("dpsAverageSeconds", "DPS averaging period", "integer", 1, 3600, "seconds", "Longer periods smooth short bursts.")
        this.Field("hideDPSForFocused", "Hide DPS for the focused client", "bool",,,, "Keep counting damage while hiding DPS on the client currently in focus.", ["DPS Monitoring", "hideDPSForFocused"], "Focused-client DPS")
        this.Section("Monitoring & alerts", "sounds", "Sound library and playback", "", "Shared settings affect every profile. Event assignments and event volume belong to the event profile.")
        this.Field("MasterSoundVolume", "Master event volume", "slider", 0, 300, "%", "Above 100% boosts the original sound.", ["MasterSoundVolume"], "", true)
        this.Section("Monitoring & alerts", "sound-playback", "Advanced sound playback", "", "Playback limits are shared across all profiles. Most users can keep the defaults.")
        this.Field("MaxSimultaneousSounds", "Maximum simultaneous event sounds", "integer", 1, 16, "sounds", "A new event sound replaces the oldest at this limit.", ["MaxSimultaneousSounds"], "", true)
        this.Field("SoundRepeatTimeout", "Minimum time before repeating the same sound", "integer", 0, 3600000, "ms", "Repeated files share this limit across events and characters. 0 allows repeats; manual previews ignore this limit.", ["SoundRepeatTimeout"], "", true)

        this.Section("Profiles", "profiles", "Profile management", "", "Selecting a profile edits its draft. Use Activate on Apply to change the running profile.")
        this.Field("ProfileActivationHotkey", "Activate this profile with a shortcut", "hotkey",,,,, ["ActivationHotkey"])
        this.Section("Profiles", "inheritance", "Section inheritance", "", "Inherited sections always show the effective Default values. Choosing Customize keeps this profile's stored values.")
        this.Section("Other applications", "app-groups", "Application groups", "Non-EVE Applications", "Pick running windows or enter exact executable names. Each row includes its optional title.")
        this.Section("Other applications", "app-shortcuts", "Application shortcuts", "Non-EVE Applications", "Each row contains the executable, optional window title and shortcut.")
        this.Section("Application", "interface", "Settings appearance", "", "Theme changes preview immediately. Apply saves the choice for future settings windows.")
        this.Field("SettingsTheme", "Settings theme", "choice",,,, "Choose a light or dark interface, or follow the Windows app theme.", ["SettingsTheme"], "", true)
        this.Choices("SettingsTheme", "Use Windows setting", "system", "Light", "light", "Dark", "dark")
        for id in SettingsPalettes.Order
            this.Fields["SettingsTheme"].choices.Push({label: SettingsPalettes.Items[id]["label"], value: id})
        this.Field("SettingsEditor", "Settings editor", "choice",,,, "Automatic prefers the compatible modern editor beside the application. The internal editor is always available from the tray.", ["SettingsEditor"], "", true)
        this.Choices("SettingsEditor", "Automatic", "automatic", "Modern (C#)", "modern", "Built-in (AHK)", "builtin")
        this.Section("Application", "tray", "Tray menu commands", "Tray Menu Settings", "Checked commands appear in the tray menu. These checks do not turn the features themselves on.")
        ; Ordinary metadata is generated from the canonical default command list.
        for name, visible in JSON.Load(default_JSON)["_Profiles"]["Default"]["Tray Menu Settings"]["TrayMenuShortcuts"]
            this.Field("TrayCommand:" name, name, "bool",,,, "Show this command in the tray menu.", ["Tray Menu Settings", "TrayMenuShortcuts", name])
        this.Section("Application", "performance", "Performance and compatibility", "Other")
        this.Field("ThumbnailPollingInterval", "Time between window and focus checks", "integer", 50, 1000, "ms", "Higher values reduce window checks and delay client and focus updates.", ["Thumbnails Behavior", "ThumbnailPollingInterval"], "Thumbnails Behavior")
        this.Field("SlowThumbnailCreation", "Stagger preview creation at startup", "bool",,,, "Wait 3 seconds, then create ready previews gradually to reduce startup GPU pressure.")
        this.Field("AutoPauseOnDwmFailure", "Pause previews if Windows composition fails", "bool",,,, "Resume previews with the tray command or preview shortcut.")
        this.Field("SwitchLangOnErr", "Retry failed shortcuts using an English keyboard layout", "bool",,,, "Compatibility fallback for shortcut registration; this does not change the interface language.")
        this.Section("Application", "diagnostics", "Diagnostics", "", "The following commands run immediately. They do not apply or discard your settings draft.")
        this.Field("DebugMode", "Show additional debugging information", "bool",,,, "Show diagnostic tooltips for app operations. File logging stays active when this is off.", ["DebugMode"], "", true)
        this.Field("ThisThat", "Show switching combo and total counters", "bool",,,, "Briefly show your switching streak and total. Hidden while debugging information is enabled.", ["ThisThat"], "", true)
        this.Section("Application", "about", "About and updates", "", "Version details, help, release notes and explicitly labelled installers.")
        this.FindSection("sounds").navTitle := "Sounds"
        this.FindSection("sound-playback").navTitle := "Advanced playback"
        this.FindSection("performance").navTitle := "Advanced"
        ; Storage paths remain compatible; inheritance follows the visible subsections.
        for section in this.Sections {
            section.legacyGroup := section.group
            section.group := section.group = "" ? "" : section.id
            for id in section.fields {
                field := this.Fields[id]
                field.legacyGroup := field.group
                if !field.global && field.group != ""
                    field.group := section.id
            }
        }
        ; Reuse the Log sources fields and their inheritance source on this page.
        this.FindSection("monitored-characters").fields.Push("gameLogsMonitoringEnabled", "monitorOnlySelectedChars")

    }

    static Section(page, id, title, group, help := "", storage := "") {
        this.currentSection := {page: page, id: id, title: title, navTitle: title, group: group, help: help, fields: []}
        this.Sections.Push(this.currentSection)
        this.blockStorage := storage != "" ? storage : group
    }

    static Field(id, label, kind, min := 0, max := 0, unit := "", help := "", path?, group?, shared := false) {
        section := this.currentSection
        field := {id: id, label: label, kind: kind, min: min, max: max, unit: unit, help: help,
            path: IsSet(path) ? path : [this.blockStorage, id], group: IsSet(group) ? group : section.group,
            global: shared, section: section.id, page: section.page, dependency: "", choices: []}
        if help = ""
            field.help := SettingsTooltips.Description(id)
        this.Fields[id] := field
        section.fields.Push(id)
    }

    static Choices(id, options*) {
        choices := []
        loop options.Length // 2
            choices.Push({label: options[A_Index * 2 - 1], value: options[A_Index * 2]})
        this.Fields[id].choices := choices
    }

    static Depends(id, parent) => this.Fields[id].dependency := parent

    static FindSection(id) {
        for section in this.Sections
            if section.id = id
                return section
        throw ValueError("Unknown settings page: " id)
    }

    static Search(query) {
        words := StrSplit(StrLower(Trim(query)), " ")
        result := []
        for section in this.Sections {
            haystack := StrLower(section.page " " section.title " " section.help)
            keywords := Map("events", "combat player NPC attack neuts neutralization Nosferatu misses smartbomb warp fleet mining crystal conversation idle sound volume border",
                "character-shortcuts", "character pilot key bindings assignments", "character-groups", "character pilot group members order next previous first",
                "character-appearance", "character pilot custom active inactive border text colors", "custom-names", "character names replacement aliases",
                "labels", "character names replacement aliases anonymous anonymization random names", "app-groups", "process executable exe title window application groups",
                "app-shortcuts", "process executable exe title window shortcut", "mouse-actions", "gesture click drag resize move minimize close Quick Group exclude disable",
                "close-protection", "character exceptions close protection never minimize", "tray", "tray menu show hide commands",
                "sounds", "audio WAV MP3 library file volume gain batch assignments timeout simultaneous", "about", "version upgrade downgrade stable prerelease help changelog")
            if keywords.Has(section.id)
                haystack .= " " StrLower(keywords[section.id])
            for id in section.fields {
                field := this.Fields[id]
                haystack .= " " StrLower(field.label " " field.help " " id)
            }
            match := true
            for word in words
                if word != "" && !InStr(haystack, word)
                    match := false
            if match
                result.Push(section)
        }
        return result
    }
}
