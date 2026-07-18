default_JSON := "
(
{
    "settings_version": "3",
    "LastUsedProfile": "Default",
    "First_Start_After_Update": 0,
    "ThisThat" : 0,
    "DebugMode": 0,
    "_Profiles": {
        "Default": {
            "Client Settings": {
                "MinimizeInactiveClients": 0,
                "AlwaysMaximize": 0,
                "TrackClientPossitions": 0,
                "Dont_Minimize_Clients": [
                    "Example Name1",
                    "Example Name2",
                    "Example Name3"
                ],
                "Minimize_Delay": 100,
                "DontCloseOnLoginScreen": 0,
                "dontCloseActiveClient": 0,
                "DontCloseClients" : [
                    "Example Name1",
                    "Example Name2",
                    "Example Name3"
                ]
            },
            "Thumbnails Behavior": {
                "HideThumbnailsOnLostFocus": 0,
                "ShowThumbnailsAlwaysOnTop": 1,
                "ThumbnailStartLocation": {
                    "x": 20,
                    "y": 20,
                    "width": 250,
                    "height": 140
                },
                "ThumbnailSnap": 1,
                "ThumbnailSnap_Distance": 20,
                "HideThumbForActiveWin": 0,
                "ShiftThumbsForLoginScreen": 1,
                "ShiftThumbsCollisionCheck": 1,
                "ShiftThumbsDirection": 1,
                "ShiftThumbHorizontalStep": 0,
                "ShiftThumbVerticalStep": 0,
                "PreserveThumbPosOnLogout": 1,
                "PreserveCharNameOnLogout": 0,
                "AutoSaveThumbnailPositions": 0,
                "HideThumbnails": 0,
                "ClickThroughActive": 0
            },
            "Thumbnails Interactions": {
                "ActivateThumbnail": {"lmb": 1, "rmb": 0, "shift": 0, "ctrl": 0},
                "MoveThumbnail": {"lmb": 0, "rmb": 1, "shift": 0, "ctrl": 0},
                "ResizeThumbnail": {"lmb": 1, "rmb": 1, "shift": 0, "ctrl": 1},
                "MoveAllThumbnails": {"lmb": 0, "rmb": 1, "shift": 0, "ctrl": 1},
                "ResizeAllThumbnails": {"lmb": 1, "rmb": 1, "shift": 0, "ctrl": 0},
                "HideThumbnail": {"lmb": 0, "rmb": 0, "shift": 0, "ctrl": 0},
                "MinimizeClient": {"lmb": 1, "rmb": 0, "shift": 0, "ctrl": 1},
                "CloseClient": {"lmb": 0, "rmb": 0, "shift": 0, "ctrl": 0},
                "DisableFromGroups": {"lmb": 1, "rmb": 0, "shift": 1, "ctrl": 0},
                "QuickGroup": {"lmb": 0, "rmb": 1, "shift": 1, "ctrl": 0}
            },
            "Thumbnails Visuals": {
                "ShowThumbnailTextOverlay": 1,
                "ThumbnailTextColor":"FAC57A",
                "ThumbnailTextSize": 12,
                "ThumbnailTextFont": "Gill Sans MT",
                "ThumbnailTextMargins": {
                    "x": 5,
                    "y": 5
                },
                "ShowClientHighlightBorder": 1,
                "ClientHighligtColor": "E36A0D",
                "ClientHighligtBorderthickness": 4,
                "ThumbnailOpacity": 80,
                "ShowAllColoredBorders": 0,
                "InactiveClientBorderthickness": 2,
                "InactiveClientBorderColor": "8A8A8A",
                "ThumbnailBackgroundColor":"57504E",
                "ThumbnailMinimumSize": {
                    "width": 50,
                    "height": 50
                }
            },
            "Hotkeys Settings": {
                "Suspend_Hotkeys_Hotkey": "",
                "Global_Hotkeys": 1,
                "SwitchToPreviousWindow_Hotkey": "",
                "CycleEveryLoggedIn_Hotkey": "",
                "Login_Screen_Cycle_Hotkey": "",
                "LoginScreenCycleDirection": 1,
                "PreserveHotkeysOnLogout": 0,
                "KeepGroupsPositions": 0,
                "Close_Active_EVE_Win_Hotkey": "",
                "Close_All_EVE_Win_Hotkey": "",
                "Reload_Program_Hotkey": "",
                "CharacterHotkeys": [
                    {"Example Name1":"1"},
                    {"Example Name2":"ctrl & 1"},
                    {"Example Name3":"Xbutton1 & 1"},
                    {"Example Name4":"^XButton1 & 1"}
                ],
                "GroupsHoldDelay": 100,
                "MaxActiveWindowRetries": 3,
                "ActiveWindowRetryInterval": 25,
                "HideThumbnailsHotkey": "",
                "ClickThroughHotkey": "",
                "dynamicGroupsColor": "ff0000",
                "QuickGroupColor": "72efdd",
                "QuickGroupHotkey": "",
                "QuickGroupIgnoredInOtherGroups": 1,
                "QuickGroupResetsPosition": 1,
                "DontCloseDisabledClients": 0,
                "DontCloseQuickGroupClients": 0
            },
            "Thumbnail Positions": {},
            "Client Possitions": {},
            "Thumbnail Visibility":{},
            "Hotkey Groups":{},
            "Custom Colors":{                
                "cColorActive": "0",
                "cColors": {
                    "CharNames": ["Example Char"],
                    "TextColor": ["FFFFFF"],
                    "Bordercolor":["FFFFFF"],                
                    "IABordercolor":["FFFFFF"]                    
                }
            },
            "Other": {
                "SwitchLangOnErr": 0,
                "Global_Groups": {
                    "Client Settings": 0,
                    "Thumbnails Interactions": 0,
                    "Thumbnails Behavior": 0,
                    "Thumbnails Visuals": 0,
                    "Hotkeys Settings": 0,
                    "Thumbnail Visibility": 0,
                    "Custom Colors": 0,
                    "Game Logs Monitoring": 0,
                    "Monitored Events": 0,
                    "Tray Menu Settings": 0,
                    "Other": 0,
                    "Hotkey Groups": 0,
                    "Non-EVE Applications": 0
                }
            },
            "Tray Menu Settings": {
                "TrayMenuShortcuts": {
                    "Suspend Hotkeys": 1,
                    "Hide Thumbnails": 1,
                    "Minimize Inactive Clients": 0,
                    "Close all EVE Clients": 1,
                    "Restore Client Positions": 1,
                    "Save Client Positions": 1,
                    "Auto Save Thumbnail Positions": 1,
                    "Save Thumbnail Positions": 1,
                    "Click Through Thumbnails": 0,
                    "Show Thumbnails Always on Top": 0,
                    "Don't Close Active Client": 0
                }
            },
            "Non-EVE Applications": {
                "NonEVEGroups": {},
                "NonEVEHotkeys": {
                "exe": ["App1.exe"], 
                "title": ["Titel"], 
                "hotkey": ["*^!Tab"]
                }
            },
            "Game Logs Monitoring": {
                "gameLogsMonitoringEnabled": 0,
                "monitoringInterval": 1000,
                "gameLogsDirectory": "",
                "charsIds": {},
                "monitorOnlySelectedChars": 0,
                "charsToMonitor": [],
                "lastEventPriority": 1,
                "supressForFocused": 1, 
                "showEventText": 0,
                "flashBorderEnabled": 1,
                "stopDisplayingOnSwitch": 1,
                "eventDisplayDuration": 2000,
                "flashBorderInterval": 300,
                "shootingInterval": 10000,
                "monitoredEvents": {
                    "underAttackByPlayer": {"enabled": 0, "color": "ff0000"},
                    "underAttackByNPC": {"enabled": 0, "color": "ff8800"},
                    "engagedWithFactionBSNPC": {"enabled": 0, "color": "11470d"},
                    "engagedWithOfficerNPC": {"enabled": 0, "color": "340e73"},
                    "engagedWithCapitalNPC": {"enabled": 0, "color": "ffd700"},
                    "warpDisrupted": {"enabled": 0, "color": "ff0000"},
                    "fleetInvited": {"enabled": 0, "color": "2196f3"},
                    "fleetWarped": {"enabled": 0, "color": "ffeb3b"},
                    "fleetRegrouped": {"enabled": 0, "color": "ffeb3b"},
                    "decloaked": {"enabled": 0, "color": "ff0000"},
                    "convoRequest": {"enabled": 0, "color": "2196f3"},
                    "conduited": {"enabled": 0, "color": "ffeb3b"},
                    "gateJumped": {"enabled": 0, "color": "2196f3"},
                    "crystalBroke": {"enabled": 0, "color": "63f321"},
                    "miningStopped": {"enabled": 0, "color": "63f321"},
                    "miningBayIsFull": {"enabled": 0, "color": "63f321"},
                    "stoppedShooting": {"enabled": 0, "color": "ff8800"},
                    "undockedFromNPCStation": {"enabled": 0, "color": "2196f3"}
                }
            }
        }
    }
} 
)"
