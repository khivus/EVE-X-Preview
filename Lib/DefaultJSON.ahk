default_JSON := "
(
{
    "settings_version": "2",
    "LastUsedProfile": "Default",
    "First_Start_After_Update": 0,
    "ThisThat" : 0,
    "_Profiles": {
        "Default": {
            "Client Settings": {
                "MinimizeInactiveClients": false,
                "AlwaysMaximize": false,
                "TrackClientPossitions": false,
                "Dont_Minimize_Clients": [
                    "Example Name1",
                    "Example Name2",
                    "Example Name3"
                ],
                "Minimize_Delay": 100,
                "DontCloseOnLoginScreen": 0,
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
                "ThumbnailSnap": true,
                "ThumbnailSnap_Distance": 20,
                "HideThumbForActiveWin": 0,
                "ShiftThumbsForLoginScreen": 1,
                "ShiftThumbsCollisionCheck": 1,
                "ShiftThumbsDirection": 1,
                "ShiftThumbHorizontalStep": 0,
                "ShiftThumbVerticalStep": 0,
                "PreserveThumbPosOnLogout": 1,
                "PreserveCharNameOnLogout": 0,
                "AutoSaveThumbnailPositions": 1,
                "HideThumbnails": 0,
                "ClickThroughActive": 0
            },
            "Thumbnails Visuals": {
                "ShowThumbnailTextOverlay": 1,
                "ThumbnailTextColor":"#FAC57A",
                "ThumbnailTextSize": 12,
                "ThumbnailTextFont": "Gill Sans MT",
                "ThumbnailTextMargins": {
                    "x": 10,
                    "y": 5
                },
                "ShowClientHighlightBorder": 1,
                "ClientHighligtColor": "#E36A0D",
                "ClientHighligtBorderthickness": 4,
                "ThumbnailOpacity": 80,
                "ShowAllColoredBorders": 0,
                "InactiveClientBorderthickness": 2,
                "InactiveClientBorderColor": "#8A8A8A",
                "ThumbnailBackgroundColor":"#57504E",
                "ThumbnailMinimumSize": {
                    "width": 50,
                    "height": 50
                }
            },
            "Hotkeys Settings": {
                "Suspend_Hotkeys_Hotkey": "",
                "Global_Hotkeys": 1,
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
                "HideThumbnailsHotkey": "",
                "ClickThroughHotkey": ""
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
                "Check_Updates": 1,
                "Global_Groups": {
                    "Client Settings": 0,
                    "Thumbnails Behavior": 0,
                    "Thumbnails Visuals": 0,
                    "Hotkeys Settings": 0,
                    "Thumbnail Visibility": 0,
                    "Custom Colors": 0,
                    "Game Logs Monitoring": 0,
                    "Monitored Events": 0,
                    "Tray Menu Settings": 0,
                    "Other": 0
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
                    "Show Thumbnails Always on Top": 0
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
                "flashBorder": 1,
                "flashBorderUntilSwitched": 1,
                "flashBorderDuration": 2000,
                "flashBorderInterval": 300,
                "shootingInterval": 10000,
                "monitoredEvents": {
                    "underAttackByPlayer": {"enabled": 0, "color": "ff0000"},
                    "underAttackByNPC": {"enabled": 0, "color": "ffeb3b"},
                    "engagedWithFactionBSNPC": {"enabled": 0, "color": "11470d"},
                    "underAttackByOfficerNPC": {"enabled": 0, "color": "340e73"},
                    "damagedCapitalNPC": {"enabled": 0, "color": "ffd700"},
                    "warpDisrupted": {"enabled": 0, "color": "ff0000"},
                    "fleetInvited": {"enabled": 0, "color": "2196f3"},
                    "fleetWarped": {"enabled": 0, "color": "ffeb3b"},
                    "fleetRegrouped": {"enabled": 0, "color": "ffeb3b"},
                    "decloaked": {"enabled": 0, "color": "ff0000"},
                    "convoRequest": {"enabled": 0, "color": "2196f3"},
                    "conduited": {"enabled": 0, "color": "ffeb3b"},
                    "gateJumped": {"enabled": 0, "color": "2196f3"},
                    "crystalBroke": {"enabled": 0, "color": "ffeb3b"},
                    "miningStopped": {"enabled": 0, "color": "63f321"},
                    "miningBayIsFull": {"enabled": 0, "color": "63f321"},
                    "stoppedShooting": {"enabled": 0, "color": "ffeb3b"}
                }
            }
        }
    }
} 
)"
