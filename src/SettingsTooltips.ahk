; Short explanations shared by fields, their labels and color previews.
class SettingsTooltips {
    static Fields := Map(
        ; Profiles and client windows.
        "SelectedProfile", "Choose the profile whose settings you want to edit.",
        "ProfileActivationHotkey", "Switch to this profile with a shortcut. Leave blank to disable it.",
        "TestThumbnailButton", "Preview appearance and sample alerts. Click again to close.",
        "MinimizeInactiveClients", "Minimize EVE clients when you switch away from them.",
        "AlwaysMaximize", "Maximize an EVE client whenever you activate it.",
        "TrackClientPossitions", "Restore saved EVE window positions and sizes.",
        "Minimizeclients_Delay", "Wait this many milliseconds before minimizing an inactive client.",
        "DontCloseOnLoginScreen", "Protect login-screen windows from the close-all action.",
        "dontCloseActiveClient", "Protect the focused EVE client from the close-all action.",
        "Dont_Minimize_Clients", "Never auto-minimize these characters. Enter one name per line.",
        "DontCloseClients", "Protect these characters from close-all. Enter one name per line.",
        ; Per-character colors.
        "Ccoloractive", "Use the per-character colors listed below instead of the default colors.",
        "Cchars", "Characters to customize, one per line. Colors on the same line belong to that character.",
        "CBorderColor", "Active border colors, one per character line. Use a color name or RGB hex code.",
        "CTextColor", "Label colors, one per character line. Use a color name or RGB hex code.",
        "IABorderColor", "Inactive border colors, one per character line. Use a color name or RGB hex code.",
        ; Hotkey groups and shortcuts.
        "AutoWildcardWindowHotkeys", "Let switching shortcuts work while extra modifier keys are held. Saved shortcuts stay unchanged.",
        "PreserveHotkeysOnLogout", "Keep a character's switching shortcuts linked to its client after logout.",
        "KeepGroupsPositions", "Remember your place in each group when switching between groups.",
        "GroupsHoldDelay", "Minimum time between group switches while holding a shortcut. Minimum: 30 ms.",
        "MaxActiveWindowRetries", "Focus checks after a recent group switch. Set to 1 to skip waiting.",
        "ActiveWindowRetryInterval", "Milliseconds between focus checks after a group switch.",
        "HotkeyGroupDDL", "Choose the character group to edit.",
        "ForwardsKey", "Activate the next available character in this group's list.",
        "BackwardsdKey", "Activate the previous available character in this group's list.",
        "FirstCharHotkey", "Activate the first available character in this group's list.",
        "HKCharlist", "Group members in switching order, one character name per line.",
        "ImpNamesBtn", "Add character names from currently running EVE clients.",
        "Suspend_Hotkeys_Hotkey", "Pause or resume EVE character activation and cycling shortcuts. Other application commands and non-EVE shortcuts remain available.",
        "Suspend_All_Hotkeys_Hotkey", "Pause or resume every application hotkey. Only this resume shortcut remains available while all hotkeys are paused; the tray menu still works.",
        "HideThumbnailsHotkey", "Toggle thumbnail visibility with this shortcut.",
        "ToggleLivePreviewsHotkey", "Pause or resume live previews, releasing preview resources while paused.",
        "ClickThroughHotkey", "Toggle whether mouse clicks pass through thumbnails.",
        "Hotkey_Scoope", "Choose whether switching shortcuts work everywhere or only while EVE is focused.",
        "SwitchToPreviousWindow_Hotkey", "Switch back to the previously active window.",
        "CycleEveryLoggedIn_Hotkey", "Cycle through all logged-in EVE character windows.",
        "Login_Screen_Cycle_Hotkey", "Cycle through EVE clients at the login screen.",
        "LoginScreenCycleDirectionForwards", "Cycle login screens from the oldest client to the newest.",
        "LoginScreenCycleDirectionBackwards", "Cycle login screens from the newest client to the oldest.",
        "Close_Active_EVE_Win_Hotkey", "Close the currently focused EVE client.",
        "Close_All_EVE_Win_Hotkey", "Close all EVE clients except those protected by Client Settings.",
        "Reload_Program_Hotkey", "Restart EVE-X-Preview and reload its settings.",
        "Exit_Program_Hotkey", "Exit EVE-X-Preview with this shortcut.",
        "HotkeyCharList", "Characters with individual shortcuts, one name per line.",
        "HotkeyList", "Shortcuts for the characters on matching lines. Use the keyboard button to record one.",
        ; Thumbnail gestures and quick groups.
        "DisableFromGroupsColor", "Border color for characters excluded from group cycling.",
        "DontCloseDisabledClients", "Protect characters disabled from groups from the close-all action.",
        "PreserveDisabledOnLogout", "Keep a character excluded from groups after logout and relogin.",
        "QuickGroupHotkey", "Cycle through characters added to the Quick Group with the thumbnail gesture.",
        "QuickGroupSortOrder", "Cycle Quick Group members in added order or alphabetically.",
        "QuickGroupColor", "Border color marking Quick Group members.",
        "QuickGroupIgnoredInOtherGroups", "Skip Quick Group members when cycling other groups.",
        "QuickGroupResetsPosition", "Restart other groups from their first member after using the Quick Group.",
        "DontCloseQuickGroupClients", "Protect Quick Group members from the close-all action.",
        "PreserveQuickGroupOnLogout", "Keep Quick Group membership after a character logs out and back in.",
        ; Thumbnail behavior and placement.
        "AutoSaveThumbnailPositions", "Automatically save thumbnail positions and sizes after moving or resizing them.",
        "HideThumbnails", "Hide all thumbnail windows until visibility is turned back on.",
        "HideThumbnailsOnLostFocus", "Hide thumbnails while another application is focused.",
        "ShowThumbnailsAlwaysOnTop", "Keep thumbnails above other windows.",
        "ExcludeThumbnailsFromCapture", "Ask Windows to exclude thumbnails from screenshots and screen sharing.",
        "ClickThroughActive", "Send mouse clicks to the windows underneath thumbnails.",
        "ThumbnailPollingInterval", "Time between window and focus checks: 50-1000 ms. Higher values update more slowly.",
        "ThumbnailStartLocationx", "Default horizontal screen position in pixels. Negative values support monitors to the left.",
        "ThumbnailStartLocationy", "Default vertical screen position in pixels. Negative values support monitors above.",
        "ThumbnailStartLocationwidth", "Default thumbnail width in pixels, used for new or reset thumbnails.",
        "ThumbnailStartLocationheight", "Default thumbnail height in pixels, used for new or reset thumbnails.",
        "LockThumbnailPositions", "Prevent thumbnail move and resize gestures from changing the layout.",
        "ThumbnailSnap", "Snap moved thumbnails to nearby thumbnail corners.",
        "ThumbnailSnap_Distance", "How close thumbnail corners must be before snapping, in pixels.",
        "HideThumbForActiveWin", "Hide the focused client's thumbnail. Its event sound alerts are also suppressed.",
        "ShiftThumbsForLoginScreen", "Place login-screen thumbnails in a sequence from the default position.",
        "ShiftThumbsForLoggedInCharacters", "Automatically place logged-in thumbnails that have no saved position.",
        "ShiftThumbsCollisionCheck", "Skip occupied positions when automatically placing thumbnails.",
        "ShiftThumbsDirection", "Choose the direction and row order for automatic thumbnail placement.",
        "ShiftThumbHorizontalStep", "Horizontal spacing between automatic positions. 0 uses the thumbnail width.",
        "ShiftThumbVerticalStep", "Vertical spacing between automatic positions. 0 uses the thumbnail height.",
        "PreserveCharNameOnLogout", "Keep the last character's label visible on its login-screen thumbnail.",
        "PreserveThumbPosOnLogout", "Keep the thumbnail in its current position when its character logs out.",
        ; Thumbnail appearance.
        "ShowThumbnailTextOverlay", "Show character names and other text on thumbnails.",
        "ThumbnailTextColor", "Default label color. Use a color name or RGB hex code such as FFFFFF.",
        "ThumbnailTextSize", "Font size of thumbnail labels, in points.",
        "ThumbnailTextFont", "Installed font used for thumbnail labels, such as Arial.",
        "ThumbnailTextMarginsx", "Horizontal inset of the label from the thumbnail edge, in pixels.",
        "ThumbnailTextMarginsy", "Vertical inset of the label from the thumbnail edge, in pixels.",
        "ClientHighligtColor", "Default border color for the focused EVE client.",
        "ClientHighligtBorderthickness", "Width of the active and event borders, in pixels.",
        "ShowClientHighlightBorder", "Mark the focused client's thumbnail with a colored border.",
        "EventActiveBorderMode", "Choose which border is inside: event or active. Combined draws them in the same border.",
        "ThumbnailOpacity", "Thumbnail opacity from 0 to 100%. Lower values are more transparent.",
        "ShowAllBorders", "Show inactive and group borders as well as active and event borders.",
        "InactiveClientBorderthickness", "Width of inactive thumbnail borders, in pixels.",
        "InactiveClientBorderColor", "Default border color for clients that are not focused.",
        "ThumbnailBackgroundColor", "Background color behind the live preview and label.",
        "ThumbnailMinimumSizewidth", "Smallest thumbnail width allowed when resizing, in pixels.",
        "ThumbnailMinimumSizeheight", "Smallest thumbnail height allowed when resizing, in pixels.",
        "ThumbnailNameCharacters", "Original character names, one per line, matched to displayed names on the right.",
        "ThumbnailNameValues", "Replacement labels for the matching character lines. Leave a line blank to keep its original name.",
        "AnonymizeThumbnails", "Replace character names with random labels from the list below.",
        "AnonymousNames", "Random labels, one per line. Empty or exhausted lists use Pilot plus a number.",
        "AnonymousNamesLabel", "Random labels, one per line. Empty or exhausted lists use Pilot plus a number.",
        "ReshuffleAnonymousNames", "Assign a fresh set of random labels to the current thumbnails.",
        "Visibility_List", "Check a character to hide its thumbnail. Uncheck it to show it again.",
        ; Game-log monitoring.
        "gameLogsMonitoringEnabled", "Read EVE game logs to detect the enabled events.",
        "systemTrackingEnabled", "Read Local chat logs to show each character's current solar system.",
        "systemColorsEnabled", "Give each solar system a consistent text color on thumbnails.",
        "selectChatLogsDirectory", "Choose the folder containing EVE chat log files.",
        "chatLogsDirectory", "EVE chat-log folder. Leave blank to find it automatically.",
        "selectGameLogsDirectory", "Choose the folder containing EVE game log files.",
        "gameLogsDirectory", "EVE game-log folder. Leave blank to find it automatically.",
        "monitoringInterval", "Milliseconds between log scans. Longer intervals delay alerts and reduce scanning.",
        "lastEventPriority", "Checked: show the newest event. Unchecked: keep the first event until it clears.",
        "supressForFocused", "Suppress event alerts, including sounds, for the focused client.",
        "showEventText", "Show the detected event's name on its thumbnail.",
        "flashBorderEnabled", "Flash the thumbnail border while an event alert is displayed.",
        "stopDisplayingOnSwitch", "Clear a client's displayed alert when you switch to that client.",
        "eventDisplayDuration", "How long an event alert stays visible, in milliseconds.",
        "flashBorderInterval", "Time between event-border flashes, in milliseconds.",
        "monitorOnlySelectedChars", "Monitor only the characters listed below.",
        "charsToMonitor", "Characters to monitor, one name per line, when the filter above is enabled.",
        "shootingInterval", "Alert after this many milliseconds without an outgoing shot.",
        "IgnorePlayerSmartbombDamage", "Ignore player smartbomb hits when detecting player-attack alerts.",
        ; DPS and other applications.
        "incomingDPSEnabled", "Show average damage per second received, calculated from game logs.",
        "incomingDPSThreshold", "Show incoming DPS only at or above this value. 0 shows any incoming DPS.",
        "showIncomingDPSResistances", "Show estimated incoming damage percentages by type. ? means unknown damage.",
        "outgoingDPSEnabled", "Show average damage per second dealt, calculated from game logs.",
        "outgoingDPSThreshold", "Show outgoing DPS only at or above this value. 0 shows any outgoing DPS.",
        "dpsAverageSeconds", "Seconds used to average DPS, from 1 to 3600. Longer windows smooth short bursts.",
        "NonEVEGroupsDDL", "Choose the application group to edit.",
        "NonEVEForwardsKey", "Activate the next available application in this group's list.",
        "NonEVEBackwardsdKey", "Activate the previous available application in this group's list.",
        "NonEVEProcessGroups", "Group members in switching order. Enter one exact executable name per line.",
        "NonEVETitlesGroups", "Optional window titles for the executable names on matching lines.",
        "NonEVEProcessHtks", "Applications with individual shortcuts. Enter one exact executable name per line.",
        "NonEVETitlesHtks", "Optional window titles to match the applications on the same lines.",
        "NonEVEHotkeysHtks", "Shortcuts for the applications on matching lines. Use the keyboard button to record one.",
        ; Other settings and About actions.
        "SlowThumbnailCreation", "Wait 3 seconds at startup, then create previews gradually to reduce GPU load spikes.",
        "AutoPauseOnDwmFailure", "Pause previews on Windows composition errors. Resume in About or with your preview shortcut.",
        "SwitchLangOnErr", "Retry failed shortcut registration using the English keyboard layout.",
        "UpdateGls", "Apply the profile override choices above and restart EVE-X-Preview.",
        "UpdateThumbnails", "Reset all thumbnails to the default size from Thumbnails Behavior.",
        "checkUpdatesBtn", "Check GitHub for the latest stable and pre-release versions.",
        "showChangelogBtn", "Read the changelog included with this version.",
        "updateToReleaseBtn", "Download and install the stable version shown above.",
        "updateToPreReleaseBtn", "Download and install the pre-release version shown above.",
        "helpBtn", "Open the project documentation in your browser.",
        "reportBugBtn", "Open GitHub to report a problem.",
        "showProgramLogBtn", "View recent application events and errors.",
        "exportProgramLogBtn", "Save a diagnostic snapshot to help investigate a problem.",
        "pausePreviewsBtn", "Pause or resume all live previews. Pausing releases their Windows preview resources.",
        "debugModeBtn", "Show extra debugging information while troubleshooting.",
        "funnyBtn", "Show combo and total counters when switching clients.",
        ; Event-sound dialog.
        "MasterVolume", "Volume for all event sounds across profiles: 0-300%. Above 100% boosts the sound.",
        "MasterVolumeText", "Master volume as an exact percentage, from 0 to 300.",
        "BrowseSound", "Add a sound file to the shared library. Check events and Apply to assign it.",
        "TestSound", "Play the selected library sound as a preview.",
        "RemoveSound", "Remove the library entry. Existing assignments and the original file are kept.",
        "SoundLibrary", "Select a saved sound to preview it or assign it to events.",
        "SelectedSound", "Full path of the selected library sound.",
        "SoundVolume", "Volume for this sound file across events: 0-300%. Multiplies the master volume.",
        "SoundVolumeText", "Selected sound's volume as an exact percentage, from 0 to 300.",
        "AssignedSound", "Sound currently assigned to the event used to open this dialog.",
        "NoSound", "Remove the sound assignment from the opened event.",
        "TestAssigned", "Preview the opened event's current sound and volume settings.",
        "SelectAllEvents", "Check every event for assignment to the selected sound.",
        "ClearEventChecks", "Uncheck every event. Apply to remove assignments of the selected sound.",
        "SoundEvents", "Check events to use the selected sound. Select a row to adjust that event's volume.",
        "EventVolumeLabel", "Volume multiplier for the selected event, independent of its assigned sound.",
        "EventVolume", "Volume for the selected event: 0-300%. Multiplies master and sound volume.",
        "EventVolumeText", "Selected event's volume as an exact percentage, from 0 to 300.",
        "ApplySounds", "Save the checked event assignments and remove unchecked assignments of this sound.",
        "MaxSimultaneousSounds", "Maximum concurrent event sounds across profiles: 1-16. New sounds replace the oldest.",
        "SoundRepeatTimeout", "Minimum milliseconds between repeats of the same file. 0 allows repeats; previews ignore it."
    )

    static Actions := Map(
        "ActivateThumbnail", "Activate the client shown in this thumbnail.",
        "MoveThumbnail", "Drag to move this thumbnail.",
        "MoveAllThumbnails", "Drag to move all thumbnails together.",
        "ResizeThumbnail", "Drag to resize this thumbnail.",
        "ResizeAllThumbnails", "Drag to resize all thumbnails together.",
        "HideThumbnail", "Hide this thumbnail.",
        "MinimizeClient", "Minimize the client shown in this thumbnail.",
        "CloseClient", "Close the client shown in this thumbnail.",
        "DisableFromGroups", "Toggle whether this character is skipped when cycling groups.",
        "QuickGroup", "Add or remove this character from the Quick Group."
    )

    static Events := Map(
        "stoppedShooting", "Alert when no outgoing shot is logged for the configured idle interval.",
        "underAttackByPlayer", "Alert when a player attacks this character.",
        "underAttackByNPC", "Alert when an NPC attacks this character.",
        "engagedWithFactionBSNPC", "Alert when combat logs identify a faction battleship NPC.",
        "engagedWithOfficerNPC", "Alert when combat logs identify an officer NPC.",
        "engagedWithCapitalNPC", "Alert when combat logs identify a capital NPC.",
        "warpDisrupted", "Alert when the logs report warp disruption.",
        "decloaked", "Alert when the logs report the character being decloaked.",
        "gateJumped", "Alert when the character jumps through a gate.",
        "undockedFromNPCStation", "Alert when the logs report undocking from an NPC station.",
        "convoRequest", "Alert when the character receives a conversation request.",
        "fleetInvited", "Alert when the character receives a fleet invitation.",
        "fleetWarped", "Alert when the logs report a fleet warp.",
        "fleetRegrouped", "Alert when the logs report a fleet regroup command.",
        "conduited", "Alert when the logs report a conduit jump.",
        "crystalBroke", "Alert when a mining crystal breaks.",
        "miningStopped", "Alert when the logs report mining has stopped.",
        "miningBayIsFull", "Alert when the logs report a full mining hold."
    )

    static Description(name) {
        if this.Fields.Has(name)
            return this.Fields[name]
        if SubStr(name, 1, 7) = "Preview"
            return "Click to choose a color. " this.Description(SubStr(name, 8))
        if RegExMatch(name, "^(lmb|rmb|shift|ctrl)_(.+)$", &match) && this.Actions.Has(match[2]) {
            key := Map("lmb", "left mouse button", "rmb", "right mouse button", "shift", "Shift key", "ctrl", "Ctrl key")[match[1]]
            return this.Actions[match[2]] " Require the " key " for this gesture."
        }
        if SubStr(name, 1, 2) = "TM"
            return "Show or hide the '" StrReplace(SubStr(name, 3), "_", " ") "' item in the tray menu."
        if SubStr(name, 1, 6) = "Global"
            return "Default profile: share this category with other profiles. Other profiles: use this profile's own settings. Apply with Update."
        if RegExMatch(name, "^(E|C|N|M)(.+)$", &match) && this.Events.Has(match[2]) {
            switch match[1] {
                case "E": return this.Events[match[2]]
                case "C": return "Border color for this event. Use a color name or RGB hex code."
                case "N": return "Also count incoming energy neutralization and Nosferatu drains as attacks."
                case "M": return "Also trigger this attack alert when the incoming shot misses."
            }
        }
        return ""
    }

    static Apply(gui, sidebar := false) {
        labels := []
        for control in gui {
            text := control.Type = "Text" || control.Type = "Button" ? control.Text : ""
            description := this.Description(control.Name)
            if description = "" && control.Type = "Button" {
                switch text {
                    case "Add": description := sidebar ? "Create a settings profile, optionally copying an existing profile." : "Create a new window-switching group."
                    case "Rename": description := "Rename the selected settings profile."
                    case "Delete": description := sidebar ? "Delete the selected settings profile." : "Delete the selected window-switching group."
                    case "Import from Launched": description := "Add character names from currently running EVE clients."
                    case "Done": description := "Close the sound dialog. Changes already applied are saved."
                    default:
                        if sidebar
                            description := "Open the " text " settings category."
                }
            }
            if description != "" {
                this.Set(control, description)
                for label in labels {
                    if !label.HasProp("ToolTip")
                        this.Set(label, description)
                }
            }
            if control.Type = "Text" {
                action := StrReplace(text, " ")
                if this.Actions.Has(action)
                    this.Set(control, this.Actions[action] " Check the buttons and keys required for this gesture.")
                else if text = "Add/Delete Groups:"
                    this.Set(control, "Create or remove an ordered group of windows.")
                else if text = "LMB" || text = "RMB" || text = "Shift" || text = "Ctrl"
                    this.Set(control, "Check this column to require " text " for an action's mouse gesture.")
                ; Only setting labels, not headings or explanatory paragraphs.
                if control.Name = "" && (SubStr(text, -1) = ":" || text = "Select profile"
                    || text = "Master volume" || text = "Sound volume" || text = "Simultaneous sounds" || text = "Repeat timeout (ms)")
                    labels.Push(control)
            } else
                labels := []
        }
    }

    static Set(control, description) {
        control.ToolTip := description
        ; Static labels need SS_NOTIFY to receive mouse messages themselves.
        if control.Type = "Text"
            control.Opt("+0x100")
    }
}
