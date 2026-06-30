Class Settings_Gui {

    MainGui() {
        ;if settings got chnaged which require a restart to apply
        This.NeedRestart := 0

        SetControlDelay(-1)
        This.S_Gui := Gui("+OwnDialogs -MinimizeBox -Resize -MaximizeBox SysMenu")
        This.S_Gui.Title := "EVE-X-Preview - Settings"

        This.SetState()
        
        This.CreateSidebar()
        This.CreateMainFrame()

        This.ClientSettings_Ctrl()
        This.CustomColors_Ctrl()
        This.HotkeyGroups_Ctrl()
        This.HotkeysSettings_Ctrl()
        This.ThumbnailsBehavior_Ctrl()
        This.ThumbnailsVisuals_Ctrl()
        This.ThumbnailVisibility_Ctrl()
        This.GameLogsMonitoring_Ctrl()
        This.MonitoredEvents_Ctrl()
        This.NonEVEApps_Ctrl()
        This.TrayMenuSettings_Ctrl()
        This.Other_Ctrl()
        This.About_Ctrl()

        This._Button_Load()
        ControlClick(This.firstBtn)

        This.S_Gui.OnEvent("Close", (*) => GuiDestroy())

        GuiDestroy(*) {
            This.S_Gui.Destroy()
            if (This.NeedRestart)
                Reload()
        }

        This.S_Gui.Show(Format("w{} h{} Center", This.guiWidth, This.guiHeight))
        This.Sidebar.Show(Format("x{} y{} w{} h{}", 0, 0, This.sidebarW, This.guiHeight))
        This.MainFrame.Show(Format("x{} y{} w{} h{}", This.sidebarW, 0, This.contentW - 10, This.guiHeight - 10))
    }

    SetState() {
        This.profilesGroups := [
            "Hotkey Groups",
            "Hotkeys Settings",
            "Thumbnails Behavior",
            "Thumbnails Visuals",
            "Thumbnail Visibility",
            "Client Settings",
            "Custom Colors",
            "Game Logs Monitoring",
            "Monitored Events",
            "Non-EVE Applications",
            "Tray Menu Settings",
            "Other",
            "About"
        ]

        This.baseGrid := 8
        This.contentGap := 16

        This.guiWidth := 700
        This.guiHeight := 160 + (This.profilesGroups.Length * 42) ; Dynamic height size! Depends on buttons count

        This.btnH := 32
        This.objH := 20

        This.sidebarW := 240
        This.sidebarInnerW := This.sidebarW - This.contentGap * 2
        This.sidebar3BtnW := (This.sidebarInnerW - This.contentGap) / 3

        This.contentW := This.guiWidth - This.sidebarW

        This.lGap := 24
        This.xlGap := 32
        This.editW := 160
        This.captBtnW := 60
        This.captBtnH := 26
        ; This.editHtkW := This.editW - This.captBtnW - This.baseGrid
        This.editHtkW := 160
        This.editExW := 197
        This.editEx2W := 250
        This.editH := 200
        This.editOffset := 3
        This.offsetX := 250
        This.editExH := 300
        This.sepW := This.contentW - This.contentGap * 2 + 2
        This.cPreviewSize := 24
        This.editC := This.editW - This.cPreviewSize - This.baseGrid

        This.latestReleaseTag := ""
        This.latestPreReleaseTag := ""
    }

    CreateSidebar() {
        This.Sidebar := Gui("+Parent" This.S_Gui.Hwnd " -Caption -Border")
        This.Sidebar.BackColor := "0xf0f0f0"
        This.Sidebar.SetFont("s11 w400")

        This.Sidebar.Add("Text", Format("x{} y{}", This.contentGap, This.contentGap), "Select profile")
        
        This.SelectProfile_DDL := This.Sidebar.Add("DDL", Format("xp yp+{} w{}", This.contentGap, This.sidebarInnerW) " Section vSelectedProfile", This.Profiles_to_Array())

        btnAdd := This.Sidebar.Add("Button", Format("xp-1 ys+{} w{} h{}", This.objH + This.baseGrid + This.editOffset, This.sidebar3BtnW, This.btnH), "Add")
        btnRename := This.Sidebar.Add("Button", Format("xp+{} yp w{} h{}", This.sidebar3BtnW + This.baseGrid + 1, This.sidebar3BtnW, This.btnH), "Rename")
        btnDelete := This.Sidebar.Add("Button", Format("xp+{} yp w{} h{}", This.sidebar3BtnW + This.baseGrid + 1, This.sidebar3BtnW, This.btnH), "Delete")

        This.Sidebar.Add("Text", Format("xs yp+{}", This.btnH + This.contentGap), "Select a settings group")

        for index, btn_text in This.profilesGroups {
            if index == 1 {
                btn := This.Sidebar.Add("Button", Format("xp yp+{} w{} h{}", This.contentGap, This.sidebarInnerW, This.btnH), btn_text)
                This.firstBtn := btn
            }
            else
                btn := This.Sidebar.Add("Button", Format("xp yp+{} w{} h{}", This.btnH + This.baseGrid, This.sidebarInnerW, This.btnH), btn_text)
            btn.OnEvent("Click", (Obj, *) => This.SettingsGroup_Handler(Obj))
        }

        This.SelectProfile_DDL.Choose(This.LastUsedProfile)
        This.SelectProfile_DDL.OnEvent("Change", (obj,*) => This._Button_Load(Obj))
        btnAdd.OnEvent("Click", ObjBindMethod(This, "Create_Profile"))
        btnRename.OnEvent("Click", ObjBindMethod(This, "Rename_Profile"))
        btnDelete.OnEvent("Click", ObjBindMethod(This, "Delete_Profile"))
    }

    CreateMainFrame() {
        This.MainFrame := Gui("+Parent" This.S_Gui.Hwnd " -Caption -Border")
        This.MainFrame.BackColor := "0xf7f7f7"

        ; Create map group which holds the GUI objects for the settings groups
        This.MainFrame.Group := Map()
    }

    SettingsGroup_Handler(Obj) {
        for k, v in This.MainFrame.Group {
            if k = Obj.Text {
                for _, ob in v
                    ob.Visible := 1
            }
            else {
                for _, ob in v
                    ob.Visible := 0
            }                
        }
        This.S_Gui.Show()
    }

    createHtkCaptureBtn() {
        return This.MainFrame.Add("Button", Format("xp+{} yp-{} w{} h{}", This.editHtkW + This.baseGrid, 1, This.captBtnW + 1, This.captBtnH), "Capture")
    }

    ClientSettings_Ctrl(visible?) {
        This.MainFrame.Group["Client Settings"] := [], ClientSettings := []
        
        This.MainFrame.SetFont("s12 w700 q5")
        ClientSettings.Push This.MainFrame.Add("Text", Format("x{} y{}", This.contentGap, This.contentGap), "Client Settings")
        This.MainFrame.SetFont("s11 w400")
        ClientSettings.Push This.MainFrame.Add("Text", Format("xp yp+{} w{} h2 +0x10", This.lGap, This.sepW))

        ClientSettings.Push This.MainFrame.Add("Text", Format("xp yp+{} Section", This.lGap), "Minimize Inactive Clients:")
        ClientSettings.Push This.MainFrame.Add("CheckBox", Format("xs+{} yp", This.offsetX) " vMinimizeInactiveClients Checked" This.MinimizeInactiveClients, "On/Off")

        ClientSettings.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Always Maximize Clients:")
        ClientSettings.Push This.MainFrame.Add("CheckBox", Format("xs+{} yp", This.offsetX) " vAlwaysMaximize Checked" This.AlwaysMaximize, "On/Off")

        ClientSettings.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "EVE Window Minimize Delay (ms):")
        ClientSettings.Push This.MainFrame.Add("Edit", Format("xs+{} yp-{} w{}", This.offsetX, This.editOffset, This.editW) " vMinimizeclients_Delay", This.Minimizeclients_Delay)

        ClientSettings.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Don't Close Clients on Login Screen:")
        ClientSettings.Push This.MainFrame.Add("CheckBox", Format("xp+{} yp", This.offsetX) " vDontCloseOnLoginScreen Checked" This.DontCloseOnLoginScreen, "On/Off")

        ClientSettings.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Don't Close Active Client:")
        ClientSettings.Push This.MainFrame.Add("CheckBox", Format("xp+{} yp", This.offsetX) " vdontCloseActiveClient", "On/Off")

        ClientSettings.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Don't Minimize Clients:")
        ClientSettings.Push This.MainFrame.Add("Edit", Format("xs yp+{} w{} h{}", This.contentGap, This.editExW, This.editExH) " vDont_Minimize_Clients -Wrap", This.Dont_Minimize_List())
        ImpBtn1 := This.MainFrame.Add("Button", Format("xp yp+{} w{}", This.editExH + This.baseGrid, This.editExW), "Import from Launched")

        ClientSettings.Push This.MainFrame.Add("Text", Format("xs+{} ys Section", This.editExW + This.contentGap), "Don't Close Clients:")
        ClientSettings.Push This.MainFrame.Add("Edit", Format("xp yp+{} w{} h{}", This.contentGap, This.editExW, This.editExH) " vDontCloseClients -Wrap", This.DontCloseList())
        ImpBtn2 := This.MainFrame.Add("Button", Format("xp yp+{} w{}", This.editExH + This.baseGrid, This.editExW), "Import from Launched")

        This.MainFrame["MinimizeInactiveClients"].OnEvent("Click", (obj, *) => cSettings_EventHandler(obj))
        This.MainFrame["AlwaysMaximize"].OnEvent("Click", (obj, *) => cSettings_EventHandler(obj))
        This.MainFrame["Minimizeclients_Delay"].OnEvent("Change", (obj, *) => cSettings_EventHandler(obj))
        This.MainFrame["Dont_Minimize_Clients"].OnEvent("Change", (obj, *) => cSettings_EventHandler(obj))
        ClientSettings.Push ImpBtn1
        ImpBtn1.OnEvent("Click", (*) => This.ImportNamesFromThumbs(This.MainFrame["Dont_Minimize_Clients"]))
        This.MainFrame["DontCloseOnLoginScreen"].OnEvent("Click", (obj, *) => cSettings_EventHandler(obj))
        This.MainFrame["dontCloseActiveClient"].OnEvent("Click", (obj, *) => cSettings_EventHandler(obj))
        This.MainFrame["DontCloseClients"].OnEvent("Change", (obj, *) => cSettings_EventHandler(obj))
        ClientSettings.Push ImpBtn2
        ImpBtn2.OnEvent("Click", (*) => This.ImportNamesFromThumbs(This.MainFrame["DontCloseClients"]))

        ;Pulls the GUI Object into the Map
        This.MainFrame.Group["Client Settings"] := ClientSettings

        for k, v in This.MainFrame.Group["Client Settings"]
            v.Visible := 0

        cSettings_EventHandler(obj) {
            if (obj.name = "MinimizeInactiveClients") {
                This.MinimizeInactiveClients := obj.value
            }
            else if (obj.name = "AlwaysMaximize") {
                This.AlwaysMaximize := obj.value
            }
            else if (obj.name = "TrackClientPossitions") {
                This.TrackClientPossitions := obj.value
            }
            else if (obj.name = "Dont_Minimize_Clients") {
                This.Dont_Minimize_Clients := obj.value
            }
            else if (obj.name = "Minimizeclients_Delay") {
                This.Minimizeclients_Delay := obj.value
            }
            else if (obj.name = "DontCloseOnLoginScreen") {
                This.DontCloseOnLoginScreen := obj.value
            }
            else if (obj.name = "dontCloseActiveClient") {
                This.dontCloseActiveClient := obj.value
            }
            else if (obj.name = "DontCloseClients") {
                This.DontCloseClients := obj.value
            }
            This.NeedRestart := 1
            SetTimer(This.Save_Settings_Delay_Timer, -200)
        }
    }

    ; User defined colors per Client
    CustomColors_Ctrl() {
        This.MainFrame.Group["Custom Colors"] := [], CustomColors := []

        This.MainFrame.SetFont("s12 w700 q5")
        CustomColors.Push This.MainFrame.Add("Text", Format("x{} y{}", This.contentGap, This.contentGap), "Custom Colors")
        This.MainFrame.SetFont("s11 w400")
        CustomColors.Push This.MainFrame.Add("Text", Format("xp yp+{} w{} h2 +0x10", This.lGap, This.sepW))

        CustomColors.Push This.MainFrame.Add("Text", Format("xp yp+{} Section", This.lGap), "Custom Colors Active:")
        CustomColors.Push This.MainFrame.Add("CheckBox", Format("xs+{} yp", This.offsetX) " vCcoloractive Checked" This.CustomColorsActive, "On/Off")

        CustomColors.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Character Name:")
        CustomColors.Push This.MainFrame.Add("Edit", Format("xp ys+{} w{} h{}", This.contentGap, This.editExW, This.editH) " -Wrap vCchars", This.CustomColors_AllCharNames)

        ImpBtn := This.MainFrame.Add("Button", Format("xp yp+{} w{}", This.editH + This.baseGrid, This.editExW), "Import from Launched")

        CustomColors.Push This.MainFrame.Add("Text", Format("xs+{} ys Section", This.editExW + This.contentGap), "Active Border Color:")
        CustomColors.Push This.MainFrame.Add("Edit", Format("xp ys+{} w{} h{}", This.contentGap, This.editExW, This.editH) " -Wrap vCBorderColor", This.CustomColors_AllBColors)

        CustomColors.Push This.MainFrame.Add("Text", Format("x{} ys+{} Section", This.contentGap, This.editH + This.xlGap + This.btnH), "Text Color:")
        CustomColors.Push This.MainFrame.Add("Edit", Format("xp ys+{} w{} h{}", This.contentGap, This.editExW, This.editH) " -Wrap vCTextColor", This.CustomColors_AllTColors)

        CustomColors.Push This.MainFrame.Add("Text", Format("xs+{} ys Section", This.editExW + This.contentGap), "Inactive Border Color:")
        CustomColors.Push This.MainFrame.Add("Edit", Format("xp ys+{} w{} h{}", This.contentGap, This.editExW, This.editH) " -Wrap vIABorderColor", This.CustomColors_IABorder_Colors)

        This.MainFrame["Ccoloractive"].OnEvent("Click", (obj, *) => Cclors_Eventhandler(obj))
        This.MainFrame["Cchars"].OnEvent("Change", (obj, *) => Cclors_Eventhandler(obj))
        This.MainFrame["CBorderColor"].OnEvent("Change", (obj, *) => Cclors_Eventhandler(obj))
        This.MainFrame["CTextColor"].OnEvent("Change", (obj, *) => Cclors_Eventhandler(obj))
        This.MainFrame["IABorderColor"].OnEvent("Change", (obj, *) => Cclors_Eventhandler(obj))
        CustomColors.Push ImpBtn
        ImpBtn.OnEvent("Click", (*) => This.ImportNamesFromThumbs(This.MainFrame["Cchars"]))

        This.MainFrame.Group["Custom Colors"] := CustomColors
        for k, v in This.MainFrame.Group["Custom Colors"]
            v.Visible := 0

        Cclors_Eventhandler(obj) {
            if (obj.Name = "Ccoloractive") {
                This.CustomColorsActive := obj.value
            }
            else if (obj.Name = "Cchars") {
                indexOld := This.IndexcChars
                This.CustomColors_AllCharNames := obj.value
                if (indexOld < This.IndexcChars) {
                    obj.value := This.CustomColors_AllCharNames
                    ControlSend("^{End}", obj.Hwnd)
                }
            }
            else if (obj.Name = "CBorderColor") {
                indexOld := This.IndexcBorder
                This.CustomColors_AllBColors := obj.value
                if (indexOld < This.IndexcBorder) {
                    obj.value := This.CustomColors_AllBColors
                    ControlSend("^{End}", obj.Hwnd)
                }
            }
            else if (obj.Name = "CTextColor") {
                indexOld := This.IndexcText
                This.CustomColors_AllTColors := obj.value
                if (indexOld < This.IndexcText) {
                    obj.value := This.CustomColors_AllTColors
                    ControlSend("^{End}", obj.Hwnd)
                }
            }            
            else if (obj.Name = "IABorderColor") {
                indexOld := This.IndexcText
                This.CustomColors_IABorder_Colors := obj.value
                if (indexOld < This.IndexcText) {
                    obj.value := This.CustomColors_IABorder_Colors
                    ControlSend("^{End}", obj.Hwnd)
                }
            }            
            This.NeedRestart := 1
            SetTimer(This.Save_Settings_Delay_Timer, -200)
        }
    }

    HotkeyGroups_Ctrl() {
        This.MainFrame.Group["Hotkey Groups"] := [], Hotkey_Groups := []

        btnW := ((This.editW - This.baseGrid) / 2) + 1
        btnEditH := 26

        This.MainFrame.SetFont("s12 w700 q5")
        Hotkey_Groups.Push This.MainFrame.Add("Text", Format("x{} y{}", This.contentGap, This.contentGap), "Hotkey Groups")
        This.MainFrame.SetFont("s11 w400")
        Hotkey_Groups.Push This.MainFrame.Add("Text", Format("xp yp+{} w{} h2 +0x10", This.lGap, This.sepW))

        Hotkey_Groups.Push This.MainFrame.Add("Text", Format("xp yp+{} Section", This.lGap), "Preserve Hotkeys on Logout:")
        Hotkey_Groups.Push This.MainFrame.Add("CheckBox", Format("xp+{} yp", This.offsetX) " vPreserveHotkeysOnLogout Checked" This.PreserveHotkeysOnLogout, "On/Off")

        Hotkey_Groups.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Keep Groups Positions:")
        Hotkey_Groups.Push This.MainFrame.Add("CheckBox", Format("xp+{} yp", This.offsetX) " vKeepGroupsPositions Checked" This.KeepGroupsPositions, "On/Off")

        Hotkey_Groups.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Disable Char in Group (Shift+Click):")
        Hotkey_Groups.Push This.MainFrame.Add("CheckBox", Format("xp+{} yp", This.offsetX) " vdynamicGroupsEnabled", "On/Off")

        Hotkey_Groups.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Disabled Thumbnail Color (Hex/RGB):")
        Hotkey_Groups.Push This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{}", This.offsetX, This.editOffset, This.editC) " vdynamicGroupsColor -Wrap")
        Hotkey_Groups.Push This.MainFrame.Add("Text", Format("xp+{} yp w{} h{}", This.editC + This.baseGrid, This.cPreviewSize, This.cPreviewSize) " vPreviewdynamicGroupsColor Border")

        Hotkey_Groups.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Groups Hold Delay (ms):")
        Hotkey_Groups.Push This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{}", This.offsetX, This.editOffset, 50) " vGroupsHoldDelay")
        Hotkey_Groups.Push This.MainFrame.Add("Text", Format("xp+{} yp+{}", 50 + This.baseGrid, This.editOffset), "Minimum = 75")

        Hotkey_Groups.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Add/Delete Groups:")
        addBtn := This.MainFrame.Add("Button", Format("xp+{} yp-{} w{} h{}", This.offsetX - 1, 5, btnW, btnEditH), "Add")
        Hotkey_Groups.Push addBtn
        DeleteButton := This.MainFrame.Add("Button", Format("xp+{} yp w{} h{}", btnW + This.baseGrid, btnW, btnEditH), "Delete")
        Hotkey_Groups.Push DeleteButton

        Hotkey_Groups.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Select Group:")
        ddl := This.MainFrame.Add("DDL", Format("xp+{} yp-{} w{}", This.offsetX, 6, This.editW) " vHotkeyGroupDDL", This.GetGroupList())
        Hotkey_Groups.Push ddl

        Hotkey_Groups.Push This.MainFrame.Add("Text", Format("xs yp+{} Section", This.xlGap), "Forwards Hotkey:")
        HKForwards := This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{}", This.offsetX, This.editOffset, This.editHtkW) " vForwardsKey")
        Hotkey_Groups.Push HKForwards
        ; Hotkey_Groups.Push This.MainFrame.Add("Button", Format("xp+{} yp-{} w{} h{}", This.editHtkW + This.baseGrid, 1, This.captBtnW + 1, This.captBtnH) " vcapGrHtkBtn1", "Capture")

        Hotkey_Groups.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Backwards Hotkey:")
        HKBackwards := This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{}", This.offsetX, This.editOffset, This.editHtkW) " vBackwardsdKey")
        Hotkey_Groups.Push HKBackwards
        ; Hotkey_Groups.Push This.MainFrame.Add("Button", Format("xp+{} yp-{} w{} h{}", This.editHtkW + This.baseGrid, 1, This.captBtnW + 1, This.captBtnH) " vcapGrHtkBtn2", "Capture")

        Hotkey_Groups.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Characters List:")
        EditBox := This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{} h{}", This.offsetX - (This.editEx2W - This.editW), This.editOffset, This.editEx2W, This.editExH) " -Wrap vHKCharlist")
        Hotkey_Groups.Push EditBox
        Hotkey_Groups.Push This.MainFrame.Add("Button", Format("xp yp+{} w{}", This.editExH + This.baseGrid, This.editEx2W) " vImpNamesBtn", "Import from Launched")

        This.MainFrame["PreserveHotkeysOnLogout"].OnEvent("Click", (obj, *) => EventHandler(obj))
        This.MainFrame["KeepGroupsPositions"].OnEvent("Click", (obj, *) => EventHandler(obj))
        This.MainFrame["dynamicGroupsEnabled"].OnEvent("Click", (obj, *) => EventHandler(obj))
        This.MainFrame["dynamicGroupsColor"].OnEvent("Change", (obj, *) => EventHandler(obj))
        This.MainFrame["GroupsHoldDelay"].OnEvent("Change", (obj, *) => EventHandler(obj))
        This.MainFrame["HotkeyGroupDDL"].OnEvent("Change", (*) => SetEditText(ddl, EditBox, HKForwards, HKBackwards,))
        addBtn.OnEvent("Click", (*) => CreateNewGroup(ddl, HKForwards, HKBackwards, EditBox))
        DeleteButton.OnEvent("Click", (*) => Delete_Group(ddl, HKForwards, HKBackwards, EditBox))
        This.MainFrame["ForwardsKey"].OnEvent("Change", (obj, *) => SaveHKGroupList(obj))
        This.MainFrame["BackwardsdKey"].OnEvent("Change", (obj, *) => SaveHKGroupList(obj))
        This.MainFrame["HKCharlist"].OnEvent("Change", (obj, *) => SaveHKGroupList(obj))
        This.MainFrame["ImpNamesBtn"].OnEvent("Click", (obj, *) => This.ImportNamesFromThumbs(EditBox))
        ; This.MainFrame["capGrHtkBtn1"].OnEvent("Click", (obj, *) => This.hotkeyCapture(HKForwards))
        ; This.MainFrame["capGrHtkBtn2"].OnEvent("Click", (obj, *) => This.hotkeyCapture(HKBackwards))

        This.MainFrame.Group["Hotkey Groups"] := Hotkey_Groups
        for k, v in This.MainFrame.Group["Hotkey Groups"]
            v.Visible := 0

        EventHandler(obj) {
            if (obj.name = "PreserveHotkeysOnLogout") {
                This.PreserveHotkeysOnLogout := obj.value
            }
            else if (obj.name = "KeepGroupsPositions") {
                This.KeepGroupsPositions := obj.value
            }
            else if (obj.name = "dynamicGroupsEnabled") {
                This.dynamicGroupsEnabled := obj.value
            }
            else if (obj.name = "dynamicGroupsColor") {
                This.dynamicGroupsColor := obj.value
                This.RedrawColorPreview(obj)
            }
            else if (obj.name = "GroupsHoldDelay") {
                delay := Integer(obj.value)
                if delay < 75 {
                    This.GroupsHoldDelay := 75
                }
                else
                    This.GroupsHoldDelay := delay
            }
            This.NeedRestart := 1
            SetTimer(This.Save_Settings_Delay_Timer, -200)
        }
        CreateNewGroup(ddlObj, ForwardHKObj, BackwardHKObj, EditObj) {
            ArrayIndex := 0
            Obj := InputBox("Enter a Groupname", "Create New Group", "w200 h90")
            if (Obj.Result != "OK")
                return
            This.Hotkey_Groups[Obj.value] := []
            ddlObj.Delete()
            ddlObj.Add(This.GetGroupList())
            for k in This.Hotkey_Groups {
                if k = Obj.value {
                    ArrayIndex := A_Index
                    break
                }
            }
            EditObj.value := "", ForwardHKObj.value := "", BackwardHKObj.value := ""
            This.enableCtrlsInGroupsSettings()
            ddlObj.Choose(ArrayIndex)
            This.NeedRestart := 1
            SetTimer(This.Save_Settings_Delay_Timer, -200)
        }

        Delete_Group(ddlObj, ForwardHKObj, BackwardHKObj, EditObj) {
            if (ddlObj.Text != "" && This.Hotkey_Groups.Has(ddlObj.Text))
                This.Hotkey_Groups.Delete(ddlObj.Text)

            ddlObj.Delete()
            ddlObj.Add(This.GetGroupList())
            ForwardHKObj.value := "", BackwardHKObj.value := "", EditObj.value := ""
            This.enableCtrlsInGroupsSettings(0)
            This.NeedRestart := 1
            SetTimer(This.Save_Settings_Delay_Timer, -200)
        }

        SetEditText(ddlObj, EditObj, ForwardHKObj, BackwardHKObj) {
            text := ""
            if (ddlObj.Text != "" && This.Hotkey_Groups.Has(ddlObj.Text)) {
                for index, Names in This.Hotkey_Groups[ddlObj.Text]["Characters"] {
                    text .= This.CleanTitle(Names) "`n"
                }
                EditObj.value := text
                ForwardHKObj.value := This.Hotkey_Groups[ddlObj.Text]["ForwardsHotkey"]
                BackwardHKObj.value := This.Hotkey_Groups[ddlObj.Text]["BackwardsHotkey"]
                This.enableCtrlsInGroupsSettings()
            }
        }

        SaveHKGroupList(obj) {
            if (obj.Name = "HKCharlist" && ddl.Text != "") {
                Arr := []
                for k, v in StrSplit(obj.value, "`n") {
                    char := Trim(v, "`n ")
                    if char = ""
                        continue
                    Arr.Push(This.AntiCleanTitle(char))
                }
                This.Hotkey_Groups[ddl.Text]["Characters"] := Arr
            }
            else if (obj.Name = "ForwardsKey" && ddl.Text != "") {
                This.Hotkey_Groups[ddl.Text]["ForwardsHotkey"] := Trim(obj.value, "`n ")
            }
            else if (obj.Name = "BackwardsdKey" && ddl.Text != "") {
                This.Hotkey_Groups[ddl.Text]["BackwardsHotkey"] := Trim(obj.value, "`n ")
            }
            This.NeedRestart := 1
            SetTimer(This.Save_Settings_Delay_Timer, -200)
        }
    }

    ; Enable or disable controls in groups settings
    enableCtrlsInGroupsSettings(enable := 1) {
        This.MainFrame["ForwardsKey"].Enabled := enable
        This.MainFrame["BackwardsdKey"].Enabled := enable
        This.MainFrame["HKCharlist"].Enabled := enable
        This.MainFrame["ImpNamesBtn"].Enabled := enable
        ; This.MainFrame["capGrHtkBtn1"].Enabled := enable
        ; This.MainFrame["capGrHtkBtn2"].Enabled := enable
    }

    HotkeysSettings_Ctrl() {
        This.MainFrame.Group["Hotkeys Settings"] := [], HotkeysSettings := []

        Charlist := "", Hklist := ""
        for index, value in This._Hotkeys {
            for name, hotkey in value {
                Charlist .= This.CleanTitle(name) "`n"
                Hklist .= hotkey "`n"
            }
        }

        This.MainFrame.SetFont("s12 w700 q5")
        HotkeysSettings.Push This.MainFrame.Add("Text", Format("x{} y{}", This.contentGap, This.contentGap), "Hotkeys Settings")
        This.MainFrame.SetFont("s11 w400")
        HotkeysSettings.Push This.MainFrame.Add("Text", Format("xp yp+{} w{} h2 +0x10", This.lGap, This.sepW))

        HotkeysSettings.Push This.MainFrame.Add("Text", Format("xp yp+{} Section", This.lGap), "Suspend All Hotkeys - Hotkey:")
        HotkeysSettings.Push This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{}", This.offsetX, This.editOffset, This.editHtkW) " vSuspend_Hotkeys_Hotkey", This.Suspend_Hotkeys_Hotkey)
        ; captBtn1 := This.createHtkCaptureBtn()
        ; HotkeysSettings.Push captBtn1

        HotkeysSettings.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Hide Thumbnails - Hotkey:")
        HotkeysSettings.Push This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{}", This.offsetX, This.editOffset, This.editHtkW) " vHideThumbnailsHotkey", This.HideThumbnailsHotkey)
        ; captBtn2 := This.createHtkCaptureBtn()
        ; HotkeysSettings.Push captBtn2

        HotkeysSettings.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Click Through Thumbnails - Hotkey:")
        HotkeysSettings.Push This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{}", This.offsetX, This.editOffset, This.editHtkW) " vClickThroughHotkey", This.ClickThroughHotkey)
        ; captBtn3 := This.createHtkCaptureBtn()
        ; HotkeysSettings.Push captBtn3

        HotkeysSettings.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Hotkey Activation Scope:")
        HotkeysSettings.Push This.MainFrame.Add("DDL", Format("xp+{} yp-{} w{}", This.offsetX, This.editOffset, This.editW) " vTTT vHotkey_Scoope Choose" (This.Global_Hotkeys ? 1 : 2), ["Global", "If an EVE window is Active"])

        HotkeysSettings.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Cycle Login Screens - Hotkey:")
        HotkeysSettings.Push This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{}", This.offsetX, This.editOffset, This.editHtkW) " vLogin_Screen_Cycle_Hotkey", This.Login_Screen_Cycle_Hotkey)
        ; captBtn4 := This.createHtkCaptureBtn()
        ; HotkeysSettings.Push captBtn4

        HotkeysSettings.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Login Screen Cycle Direction:")
        HotkeysSettings.Push This.MainFrame.Add("Radio", Format("xp+{} yp", This.offsetX + 1) " vLoginScreenCycleDirectionForwards Checked" This.LoginScreenCycleDirection, "Old->New")
        HotkeysSettings.Push This.MainFrame.Add("Radio", Format("xp+{} yp", 83) " vLoginScreenCycleDirectionBackwards Checked" (This.LoginScreenCycleDirection ? 0 : 1), "New->Old")

        HotkeysSettings.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Close Active EVE Window - Hotkey:")
        HotkeysSettings.Push This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{}", This.offsetX, This.editOffset, This.editHtkW) " vClose_Active_EVE_Win_Hotkey", This.Close_Active_EVE_Win_Hotkey)
        ; captBtn5 := This.createHtkCaptureBtn()
        ; HotkeysSettings.Push captBtn5

        HotkeysSettings.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Close All EVE Windows - Hotkey:")
        HotkeysSettings.Push This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{}", This.offsetX, This.editOffset, This.editHtkW) " vClose_All_EVE_Win_Hotkey", This.Close_All_EVE_Win_Hotkey)
        ; captBtn6 := This.createHtkCaptureBtn()
        ; HotkeysSettings.Push captBtn6

        HotkeysSettings.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Reload EVE-X-Preview - Hotkey:")
        HotkeysSettings.Push This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{}", This.offsetX, This.editOffset, This.editHtkW) " vReload_Program_Hotkey", This.Reload_Program_Hotkey)
        ; captBtn7 := This.createHtkCaptureBtn()
        ; HotkeysSettings.Push captBtn7

        HotkeysSettings.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap - 3), "Character Name:")
        HKCharList := This.MainFrame.Add("Edit", Format("xp yp+{} w{} h{}", This.contentGap, This.editExW, This.editH) " -Wrap vHotkeyCharList", Charlist)
        HotkeysSettings.Push HKCharList

        ImpBtn := This.MainFrame.Add("Button", Format("xp yp+{} w{}", This.editH + This.baseGrid, This.editExW), "Import from Launched")
        HotkeysSettings.Push ImpBtn

        HotkeysSettings.Push This.MainFrame.Add("Text", Format("xs+{} ys Section", This.editExW + This.contentGap), "Hotkeys:")
        HKKeylist := This.MainFrame.Add("Edit", Format("xp yp+{} w{} h{}", This.contentGap, This.editExW, This.editH) " -Wrap vHotkeyList", Hklist)
        HotkeysSettings.Push HKKeylist

        This.MainFrame["Suspend_Hotkeys_Hotkey"].OnEvent("Change", (obj, *) => cHotkeys_EventHandler(obj))
        This.MainFrame["HideThumbnailsHotkey"].OnEvent("Change", (obj, *) => cHotkeys_EventHandler(obj))
        This.MainFrame["ClickThroughHotkey"].OnEvent("Change", (obj, *) => cHotkeys_EventHandler(obj))
        This.MainFrame["Hotkey_Scoope"].OnEvent("Change", (obj, *) => cHotkeys_EventHandler(obj))
        This.MainFrame["Login_Screen_Cycle_Hotkey"].OnEvent("Change", (obj, *) => cHotkeys_EventHandler(obj))
        This.MainFrame["LoginScreenCycleDirectionForwards"].OnEvent("Click", (obj, *) => cHotkeys_EventHandler(obj))
        This.MainFrame["LoginScreenCycleDirectionBackwards"].OnEvent("Click", (obj, *) => cHotkeys_EventHandler(obj))
        This.MainFrame["Close_Active_EVE_Win_Hotkey"].OnEvent("Change", (obj, *) => cHotkeys_EventHandler(obj))
        This.MainFrame["Close_All_EVE_Win_Hotkey"].OnEvent("Change", (obj, *) => cHotkeys_EventHandler(obj))
        This.MainFrame["Reload_Program_Hotkey"].OnEvent("Change", (obj, *) => cHotkeys_EventHandler(obj))
        HKCharList.OnEvent("Change", (obj, *) => EventHandler(obj))
        HKKeylist.OnEvent("Change", (obj, *) => EventHandler(obj))
        ImpBtn.OnEvent("Click", (*) => This.ImportNamesFromThumbs(HKCharList))
        ; captBtn1.OnEvent("Click", (obj, *) => This.hotkeyCapture(This.MainFrame["Suspend_Hotkeys_Hotkey"]))
        ; captBtn2.OnEvent("Click", (obj, *) => This.hotkeyCapture(This.MainFrame["HideThumbnailsHotkey"]))
        ; captBtn3.OnEvent("Click", (obj, *) => This.hotkeyCapture(This.MainFrame["ClickThroughHotkey"]))
        ; captBtn4.OnEvent("Click", (obj, *) => This.hotkeyCapture(This.MainFrame["Login_Screen_Cycle_Hotkey"]))
        ; captBtn5.OnEvent("Click", (obj, *) => This.hotkeyCapture(This.MainFrame["Close_Active_EVE_Win_Hotkey"]))
        ; captBtn6.OnEvent("Click", (obj, *) => This.hotkeyCapture(This.MainFrame["Close_All_EVE_Win_Hotkey"]))
        ; captBtn7.OnEvent("Click", (obj, *) => This.hotkeyCapture(This.MainFrame["Reload_Program_Hotkey"]))

        This.MainFrame.Group["Hotkeys Settings"] := HotkeysSettings
        for k, v in This.MainFrame.Group["Hotkeys Settings"]
            v.Visible := 0

        cHotkeys_EventHandler(obj) {
            if (obj.name = "Suspend_Hotkeys_Hotkey") {
                This.Suspend_Hotkeys_Hotkey := Trim(obj.value, "`n ")
            }
            else if (obj.name = "HideThumbnailsHotkey") {
                This.HideThumbnailsHotkey := Trim(obj.value, "`n ")
            }
            else if (obj.name = "ClickThroughHotkey") {
                This.ClickThroughHotkey := Trim(obj.value, "`n ")
            }
            else if (obj.name = "Hotkey_Scoope") {
                This.Global_Hotkeys := (obj.value = 1 ? 1 : 0)
            }
            else if (obj.name = "Login_Screen_Cycle_Hotkey") {
                This.Login_Screen_Cycle_Hotkey := Trim(obj.value, "`n ")
            }
            else if (obj.name = "LoginScreenCycleDirectionForwards") {
                This.LoginScreenCycleDirection := 1
            }
            else if (obj.name = "LoginScreenCycleDirectionBackwards") {
                This.LoginScreenCycleDirection := 0
            }
            else if (obj.name = "Close_Active_EVE_Win_Hotkey") {
                This.Close_Active_EVE_Win_Hotkey := Trim(obj.value, "`n ")
            }
            else if (obj.name = "Close_All_EVE_Win_Hotkey") {
                This.Close_All_EVE_Win_Hotkey := Trim(obj.value, "`n ")
            }
            else if (obj.name = "Reload_Program_Hotkey") {
                This.Reload_Program_Hotkey := Trim(obj.value, "`n ")
            }
            This.NeedRestart := 1
            SetTimer(This.Save_Settings_Delay_Timer, -200)
        }

        ;Parse All hotkeys to a Array on value change
        EventHandler(obj) {
            tempvar := []
            ListChars := StrSplit(This.MainFrame["HotkeyCharList"].value, "`n"), ListHotkeys := StrSplit(This.MainFrame["HotkeyList"].value, "`n")
            for _, v in ListChars {
                chars := "", keys := ""
                if (A_Index <= ListChars.Length) {
                    chars := Trim(ListChars[A_Index], "`n ")
                }
                if (A_Index <= ListHotkeys.Length) {
                    keys := Trim(ListHotkeys[A_Index], "`n ")
                }
                if (A_Index > ListHotkeys.Length) {
                    keys := ""
                }
                if (chars = "")
                    continue
                tempvar.Push Map(This.AntiCleanTitle(chars), keys)
            }
            this._Hotkeys := tempvar
            This.NeedRestart := 1
            SetTimer(This.Save_Settings_Delay_Timer, -200)
        }
    }

    ThumbnailsBehavior_Ctrl() {
        This.MainFrame.Group["Thumbnails Behavior"] := [], arr := []

        This.directions := Map(
            1, "left -> right, top -> bottom",
            2, "left -> right, bottom -> top",
            3, "right -> left, top -> bottom",
            4, "right -> left, bottom -> top",
            5, "top -> bottom, left -> right",
            6, "top -> bottom, right -> left",
            7, "bottom -> top, left -> right",
            8, "bottom -> top, right -> left"
        )

        ddl_options := []
        for key, value in This.directions {
            ddl_options.Push(value)
        }

        smallEditW := 35
        mediumEditW := 50

        This.MainFrame.SetFont("s12 w700 q5")
        arr.Push This.MainFrame.Add("Text", Format("x{} y{}", This.contentGap, This.contentGap), "Thumbnails Behavior")
        This.MainFrame.SetFont("s11 w400")
        arr.Push This.MainFrame.Add("Text", Format("xp yp+{} w{} h2 +0x10", This.lGap, This.sepW))

        arr.Push This.MainFrame.Add("Text", Format("xp yp+{} Section", This.lGap), "Auto Save Thumbnail Positions:")
        arr.Push This.MainFrame.Add("CheckBox", Format("xp+{} yp", This.offsetX) " vAutoSaveThumbnailPositions Checked" This.AutoSaveThumbnailPositions, "On/Off")

        arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Hide Thumbnails:")
        arr.Push This.MainFrame.Add("CheckBox", Format("xp+{} yp", This.offsetX) " vHideThumbnails Checked" This.HideThumbnails, "On/Off")

        arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Hide Thumbnails on Lost Focus:")
        arr.Push This.MainFrame.Add("CheckBox", Format("xp+{} yp", This.offsetX) " vHideThumbnailsOnLostFocus Checked" This.HideThumbnailsOnLostFocus, "On/Off")

        arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Show Thumbnails Always on Top:")
        arr.Push This.MainFrame.Add("CheckBox", Format("xp+{} yp", This.offsetX) " vShowThumbnailsAlwaysOnTop Checked" This.ShowThumbnailsAlwaysOnTop, "On/Off")

        arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Click Through Thumbnails:")
        arr.Push This.MainFrame.Add("CheckBox", Format("xp+{} yp", This.offsetX) " vClickThroughActive Checked" This.ClickThroughActive, "On/Off")

        arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Default Thumbnail Position (px):")
        arr.Push This.MainFrame.Add("Text", Format("xp+{} yp", This.offsetX), "x:")
        arr.Push This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{}", 16, This.editOffset, mediumEditW) " vThumbnailStartLocationx", This.ThumbnailStartLocation["x"])
        arr.Push This.MainFrame.Add("Text", Format("xp+{} yp+{}", mediumEditW + This.baseGrid + 20, This.editOffset), "y:")
        arr.Push This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{}", 16, This.editOffset, mediumEditW) " vThumbnailStartLocationy", This.ThumbnailStartLocation["y"])

        arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Default Thumbnail Size (px):")
        arr.Push This.MainFrame.Add("Text", Format("xp+{} yp", This.offsetX), "width:")
        arr.Push This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{}", 37, This.editOffset, smallEditW) " vThumbnailStartLocationwidth", This.ThumbnailStartLocation["width"])
        arr.Push This.MainFrame.Add("Text", Format("xp+{} yp+{}", smallEditW + This.baseGrid + 1, This.editOffset), "height:")
        arr.Push This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{}", 44, This.editOffset, smallEditW) " vThumbnailStartLocationheight", This.ThumbnailStartLocation["height"])

        arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Enable Thumbnail Snap:")
        arr.Push This.MainFrame.Add("CheckBox", Format("xp+{} yp", This.offsetX) " vThumbnailSnap Checked" This.ThumbnailSnap, "On/Off")

        arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Thumbnail Snap Distance (px):")
        arr.Push This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{}", This.offsetX, This.editOffset, This.editW) " vThumbnailSnap_Distance", This.ThumbnailSnap_Distance)

        arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Hide Thumbnail for Active Window:")
        arr.Push This.MainFrame.Add("CheckBox", Format("xp+{} yp", This.offsetX) " vHideThumbForActiveWin Checked" This.HideThumbForActiveWin, "On/Off")

        arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Shift Thumbnails on Login Screen:")
        arr.Push This.MainFrame.Add("CheckBox", Format("xp+{} yp", This.offsetX) " vShiftThumbsForLoginScreen Checked" This.ShiftThumbsForLoginScreen, "On/Off")

        arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Enable Thumbnail Collision Avoidance:")
        arr.Push This.MainFrame.Add("CheckBox", Format("xp+{} yp", This.offsetX) " vShiftThumbsCollisionCheck Checked" This.ShiftThumbsCollisionCheck, "On/Off")

        arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Thumbnail Shift Direction:")
        arr.Push This.MainFrame.Add("DDL", Format("xp+{} yp-{} w{}", This.offsetX, This.editOffset, This.editW) " vShiftThumbsDirection Choose" . Integer(This.ShiftThumbsDirection), ddl_options)

        arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Thumbnail Shift Horizontal Step (px):")
        arr.Push This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{}", This.offsetX, This.editOffset, mediumEditW) " vShiftThumbHorizontalStep 0", This.ShiftThumbHorizontalStep)
        arr.Push This.MainFrame.Add("Text", Format("xp+{} yp+{}", mediumEditW + This.baseGrid, This.editOffset), "0 = Width")

        arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Thumbnail Shift Vertical Step (px):")
        arr.Push This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{}", This.offsetX, This.editOffset, mediumEditW) " vShiftThumbVerticalStep 0", This.ShiftThumbVerticalStep)
        arr.Push This.MainFrame.Add("Text", Format("xp+{} yp+{}", mediumEditW + This.baseGrid, This.editOffset), "0 = Height")

        arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Preserve Character Name on Logout:")
        arr.Push This.MainFrame.Add("CheckBox", Format("xp+{} yp", This.offsetX) " vPreserveCharNameOnLogout Checked" This.PreserveCharNameOnLogout, "On/Off")

        arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Preserve Thumbnail Position on Logout:")
        arr.Push This.MainFrame.Add("CheckBox", Format("xp+{} yp", This.offsetX) " vPreserveThumbPosOnLogout Checked" This.PreserveThumbPosOnLogout, "On/Off")

        This.MainFrame["AutoSaveThumbnailPositions"].OnEvent("Click", (obj, *) => EventHandler(obj))
        This.MainFrame["HideThumbnails"].OnEvent("Click", (obj, *) => EventHandler(obj))
        This.MainFrame["HideThumbnailsOnLostFocus"].OnEvent("Click", (obj, *) => EventHandler(obj))
        This.MainFrame["ShowThumbnailsAlwaysOnTop"].OnEvent("Click", (obj, *) => EventHandler(obj))
        This.MainFrame["ClickThroughActive"].OnEvent("Click", (obj, *) => EventHandler(obj))
        This.MainFrame["ThumbnailStartLocationx"].OnEvent("Change", (obj, *) => EventHandler(obj))
        This.MainFrame["ThumbnailStartLocationy"].OnEvent("Change", (obj, *) => EventHandler(obj))
        This.MainFrame["ThumbnailStartLocationwidth"].OnEvent("Change", (obj, *) => EventHandler(obj))
        This.MainFrame["ThumbnailStartLocationheight"].OnEvent("Change", (obj, *) => EventHandler(obj))
        This.MainFrame["ThumbnailSnap"].OnEvent("Click", (obj, *) => EventHandler(obj))
        This.MainFrame["ThumbnailSnap_Distance"].OnEvent("Change", (obj, *) => EventHandler(obj))
        This.MainFrame["HideThumbForActiveWin"].OnEvent("Click", (obj, *) => EventHandler(obj))
        This.MainFrame["ShiftThumbsForLoginScreen"].OnEvent("Click", (obj, *) => EventHandler(obj))
        This.MainFrame["ShiftThumbsCollisionCheck"].OnEvent("Click", (obj, *) => EventHandler(obj))
        This.MainFrame["ShiftThumbsDirection"].OnEvent("Change", (obj, *) => EventHandler(obj))
        This.MainFrame["ShiftThumbHorizontalStep"].OnEvent("Change", (obj, *) => EventHandler(obj))
        This.MainFrame["ShiftThumbVerticalStep"].OnEvent("Change", (obj, *) => EventHandler(obj))
        This.MainFrame["PreserveCharNameOnLogout"].OnEvent("Click", (obj, *) => EventHandler(obj))
        This.MainFrame["PreserveThumbPosOnLogout"].OnEvent("Click", (obj, *) => EventHandler(obj))

        This.MainFrame.Group["Thumbnails Behavior"] := arr
        for k, v in This.MainFrame.Group["Thumbnails Behavior"] {
            v.Visible := 0
        }

        EventHandler(obj) {
            if (obj.name = "AutoSaveThumbnailPositions") {
                This.AutoSaveThumbnailPositions := obj.value
            }
            else if (obj.name = "HideThumbnails") {
                This.HideThumbnails := obj.value
            }
            else if (obj.name = "HideThumbnailsOnLostFocus") {
                This.HideThumbnailsOnLostFocus := obj.value
            }
            else if (obj.name = "ShowThumbnailsAlwaysOnTop") {
                This.ShowThumbnailsAlwaysOnTop := obj.value
            }
            else if (obj.name = "ClickThroughActive") {
                This.ClickThroughActive := obj.value
            }
            else if (obj.name = "ThumbnailStartLocationx") {
                This.ThumbnailStartLocation["x"] := obj.value
            }
            else if (obj.name = "ThumbnailStartLocationy") {
                This.ThumbnailStartLocation["y"] := obj.value
            }
            else if (obj.name = "ThumbnailStartLocationwidth") {
                This.ThumbnailStartLocation["width"] := obj.value
            }
            else if (obj.name = "ThumbnailStartLocationheight") {
                This.ThumbnailStartLocation["height"] := obj.value
            }
            else if (obj.name = "ThumbnailSnap") {
                This.ThumbnailSnap := obj.value
            }
            else if (obj.name = "ThumbnailSnap_Distance") {
                This.ThumbnailSnap_Distance := obj.value
            }
            else if (obj.name = "PreserveThumbPosOnLogout") {
                This.PreserveThumbPosOnLogout := obj.value
            }
            else if (obj.name = "PreserveCharNameOnLogout") {
                This.PreserveCharNameOnLogout := obj.value
            }
            else if (obj.name = "HideThumbForActiveWin") {
                This.HideThumbForActiveWin := obj.value
            }
            else if (obj.name = "ShiftThumbsForLoginScreen") {
                This.ShiftThumbsForLoginScreen := obj.value
            }
            else if (obj.name = "ShiftThumbsCollisionCheck") {
                This.ShiftThumbsCollisionCheck := obj.value
            }
            else if (obj.name = "ShiftThumbsDirection") {
                This.ShiftThumbsDirection := obj.Value
            }
            else if (obj.name = "ShiftThumbHorizontalStep") {
                This.ShiftThumbHorizontalStep := obj.value
            }
            else if (obj.name = "ShiftThumbVerticalStep") {
                This.ShiftThumbVerticalStep := obj.value
            }
            This.NeedRestart := 1
            SetTimer(This.Save_Settings_Delay_Timer, -200)
        }
    }

    ThumbnailsVisuals_Ctrl() {
        This.MainFrame.Group["Thumbnails Visuals"] := [], arr := []

        smallEditW := 35
        mediumEditW := 50

        This.MainFrame.SetFont("s12 w700 q5")
        arr.Push This.MainFrame.Add("Text", Format("x{} y{}", This.contentGap, This.contentGap), "Thumbnails Visuals")
        This.MainFrame.SetFont("s11 w400")
        arr.Push This.MainFrame.Add("Text", Format("xp yp+{} w{} h2 +0x10", This.lGap, This.sepW))

        arr.Push This.MainFrame.Add("Text", Format("xp yp+{} Section", This.lGap), "Show Thumbnail Text Overlay:")
        arr.Push This.MainFrame.Add("CheckBox", Format("xp+{} yp", This.offsetX) " vShowThumbnailTextOverlay Checked" This.ShowThumbnailTextOverlay, "On/Off")

        arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Thumbnail Text Color (Hex/RGB):")
        arr.Push This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{}", This.offsetX, This.editOffset, This.editC) " vThumbnailTextColor -Wrap", This.ThumbnailTextColor)
        arr.Push This.MainFrame.Add("Text", Format("xp+{} yp w{} h{} Background{}", This.editC + This.baseGrid, This.cPreviewSize, This.cPreviewSize, This.ThumbnailTextColor) " vPreviewThumbnailTextColor Border")

        arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Thumbnail Text Size:")
        arr.Push This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{}", This.offsetX, This.editOffset, This.editW) " vThumbnailTextSize -Wrap", This.ThumbnailTextSize)

        arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Thumbnail Text Font:")
        arr.Push This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{}", This.offsetX, This.editOffset, This.editW) " vThumbnailTextFont -Wrap", This.ThumbnailTextFont)

        arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Thumbnail Text Margins (px):")
        arr.Push This.MainFrame.Add("Text", Format("xp+{} yp", This.offsetX), "width:")
        arr.Push This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{}", 37, This.editOffset, smallEditW) " vThumbnailTextMarginsx -Wrap", This.ThumbnailTextMargins["x"])
        arr.Push This.MainFrame.Add("Text", Format("xp+{} yp+{}", smallEditW + This.baseGrid + 1, This.editOffset), "height:")
        arr.Push This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{}", 44, This.editOffset, smallEditW) " vThumbnailTextMarginsy -Wrap", This.ThumbnailTextMargins["y"])

        arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Client Highligt Color (Hex/RGB):")
        arr.Push This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{}", This.offsetX, This.editOffset, This.editC) " vClientHighligtColor -Wrap", This.ClientHighligtColor)
        arr.Push This.MainFrame.Add("Text", Format("xp+{} yp w{} h{} Background{}", This.editC + This.baseGrid, This.cPreviewSize, This.cPreviewSize, This.ClientHighligtColor) " vPreviewClientHighligtColor Border")

        arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Client Highligt Border Thickness (px):")
        arr.Push This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{}", This.offsetX, This.editOffset, This.editW) " vClientHighligtBorderthickness -Wrap", This.ClientHighligtBorderthickness)

        arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Show Client Highlight Border:")
        arr.Push This.MainFrame.Add("CheckBox", Format("xp+{} yp", This.offsetX) " vShowClientHighlightBorder Checked" This.ShowClientHighlightBorder, "On/Off")

        arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Thumbnail Opacity (%):")
        arr.Push This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{}", This.offsetX, This.editOffset, This.editW) " vThumbnailOpacity -Wrap", IntegerToPercentage(This.ThumbnailOpacity))

        arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Show All Borders:")
        arr.Push This.MainFrame.Add("CheckBox", Format("xp+{} yp", This.offsetX) " vShowAllBorders Checked" This.ShowAllColoredBorders, "On/Off")

        arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Inactive Client Border Thickness (px):")
        arr.Push This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{}", This.offsetX, This.editOffset, This.editW) " vInactiveClientBorderthickness -Wrap", This.InactiveClientBorderthickness)

        arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Inactive Client Border Color (Hex/RGB):")
        arr.Push This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{}", This.offsetX, This.editOffset, This.editC) " vInactiveClientBorderColor -Wrap", This.InactiveClientBorderColor)
        arr.Push This.MainFrame.Add("Text", Format("xp+{} yp w{} h{} Background{}", This.editC + This.baseGrid, This.cPreviewSize, This.cPreviewSize, This.InactiveClientBorderColor) " vPreviewInactiveClientBorderColor Border")

        arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Thumbnail Background Color (Hex/RGB):")
        arr.Push This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{}", This.offsetX, This.editOffset, This.editC) " vThumbnailBackgroundColor", This.ThumbnailBackgroundColor)
        arr.Push This.MainFrame.Add("Text", Format("xp+{} yp w{} h{} Background{}", This.editC + This.baseGrid, This.cPreviewSize, This.cPreviewSize, This.ThumbnailBackgroundColor) " vPreviewThumbnailBackgroundColor Border")

        arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Minimum Thumbnail Size (px):")
        arr.Push This.MainFrame.Add("Text", Format("xp+{} yp", This.offsetX), "width:")
        arr.Push This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{}", 37, This.editOffset, smallEditW) " vThumbnailMinimumSizewidth", This.ThumbnailMinimumSize["width"])
        arr.Push This.MainFrame.Add("Text", Format("xp+{} yp+{}", smallEditW + This.baseGrid + 1, This.editOffset), "height:")
        arr.Push This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{}", 44, This.editOffset, smallEditW) " vThumbnailMinimumSizeheight", This.ThumbnailMinimumSize["height"])

        This.MainFrame["ShowThumbnailTextOverlay"].OnEvent("Click", (obj, *) => EventHandler(obj))
        This.MainFrame["ThumbnailTextColor"].OnEvent("Change", (obj, *) => EventHandler(obj))
        This.MainFrame["ThumbnailTextSize"].OnEvent("Change", (obj, *) => EventHandler(obj))
        This.MainFrame["ThumbnailTextFont"].OnEvent("Change", (obj, *) => EventHandler(obj))
        This.MainFrame["ThumbnailTextMarginsx"].OnEvent("Change", (obj, *) => EventHandler(obj))
        This.MainFrame["ThumbnailTextMarginsy"].OnEvent("Change", (obj, *) => EventHandler(obj))
        This.MainFrame["ClientHighligtColor"].OnEvent("Change", (obj, *) => EventHandler(obj))
        This.MainFrame["ClientHighligtBorderthickness"].OnEvent("Change", (obj, *) => EventHandler(obj))
        This.MainFrame["ShowClientHighlightBorder"].OnEvent("Click", (obj, *) => EventHandler(obj))
        This.MainFrame["ThumbnailOpacity"].OnEvent("Change", (obj, *) => EventHandler(obj))
        This.MainFrame["ShowAllBorders"].OnEvent("Click", (obj, *) => EventHandler(obj))
        This.MainFrame["InactiveClientBorderthickness"].OnEvent("Change", (obj, *) => EventHandler(obj))
        This.MainFrame["InactiveClientBorderColor"].OnEvent("Change", (obj, *) => EventHandler(obj))
        This.MainFrame["ThumbnailBackgroundColor"].OnEvent("Change", (obj, *) => EventHandler(obj))
        This.MainFrame["ThumbnailMinimumSizewidth"].OnEvent("Change", (obj, *) => EventHandler(obj))
        This.MainFrame["ThumbnailMinimumSizeheight"].OnEvent("Change", (obj, *) => EventHandler(obj))

        This.MainFrame.Group["Thumbnails Visuals"] := arr
        for k, v in This.MainFrame.Group["Thumbnails Visuals"] {
            v.Visible := 0
        }

        EventHandler(obj) {
            if (obj.name = "ShowThumbnailTextOverlay") {
                This.ShowThumbnailTextOverlay := obj.value
            }
            else if (obj.name = "ThumbnailTextColor") {
                This.ThumbnailTextColor := obj.value
                This.RedrawColorPreview(obj)
            }
            else if (obj.name = "ThumbnailTextSize") {
                This.ThumbnailTextSize := obj.value
            }
            else if (obj.name = "ThumbnailTextFont") {
                This.ThumbnailTextFont := obj.value
            }
            else if (obj.name = "ThumbnailTextMarginsx") {
                This.ThumbnailTextMargins["x"] := obj.value
            }
            else if (obj.name = "ThumbnailTextMarginsy") {
                This.ThumbnailTextMargins["y"] := obj.value
            }
            else if (obj.name = "ClientHighligtColor") {
                This.ClientHighligtColor := obj.value
                This.RedrawColorPreview(obj)
            }
            else if (obj.name = "ClientHighligtBorderthickness") {
                This.ClientHighligtBorderthickness := obj.value
            }
            else if (obj.name = "ShowClientHighlightBorder") {
                This.ShowClientHighlightBorder := obj.value
            }
            else if (obj.name = "ThumbnailOpacity") {
                This.ThumbnailOpacity := obj.value
            }
            else if (obj.Name = "ShowAllBorders") {
                This.ShowAllColoredBorders := obj.value
                This.MainFrame["InactiveClientBorderthickness"].Enabled := This.ShowAllColoredBorders
                This.MainFrame["InactiveClientBorderColor"].Enabled := This.ShowAllColoredBorders
            }
            else if (obj.Name = "InactiveClientBorderColor") {
                This.InactiveClientBorderColor := obj.value
                This.RedrawColorPreview(obj)
            }
            else if (obj.Name = "InactiveClientBorderthickness") {
                This.InactiveClientBorderthickness := obj.value
            }
            else if (obj.name = "ThumbnailBackgroundColor") {
                This.ThumbnailBackgroundColor := obj.value
                This.RedrawColorPreview(obj)
            }
            else if (obj.name = "ThumbnailMinimumSizewidth") {
                This.ThumbnailMinimumSize["width"] := obj.value
            }
            else if (obj.name = "ThumbnailMinimumSizeheight") {
                This.ThumbnailMinimumSize["height"] := obj.value
            }
            This.NeedRestart := 1
            SetTimer(This.Save_Settings_Delay_Timer, -200)
        }
    }

    RedrawColorPreview(obj) {
        try
            This.MainFrame["Preview" obj.name].Opt("Background" obj.value)
        This.MainFrame["Preview" obj.name].Redraw()
    }

    ThumbnailVisibility_Ctrl() {
        This.MainFrame.Group["Thumbnail Visibility"] := [], Thumbnail_visibility := []

        This.MainFrame.SetFont("s12 w700 q5")
        Thumbnail_visibility.Push This.MainFrame.Add("Text", Format("x{} y{}", This.contentGap, This.contentGap), "Thumbnail Visibility")
        This.MainFrame.SetFont("s11 w400")
        Thumbnail_visibility.Push This.MainFrame.Add("Text", Format("xp yp+{} w{} h2 +0x10", This.lGap, This.sepW))

        Thumbnail_visibility.Push This.MainFrame.Add("Text", Format("xp yp+{} Section", This.lGap), "Select Any Client to Hide The Thumbnail:")
        This.Tv_LV := This.MainFrame.Add("ListView", Format("xp yp+{} w{}", This.contentGap, This.editEx2W) " r20 Checked -LV0x10 -Multi -Sort vVisibility_List", ["Client Name"])
        Thumbnail_visibility.Push This.Tv_LV

        for k, v in This.compare_openclients_with_list() {
            if (k != "EVE" || v != "") {
                if This.Thumbnail_visibility.Has(v)
                    This.Tv_LV.Add("Check", v,)
                else
                    This.Tv_LV.Add("", v,)
            }
        }

        This.Tv_LV.ModifyCol(1, This.editEx2W) ; , This.Tv_LV.ModifyCol(2, 115)
        This.Tv_LV.OnEvent("ItemCheck", ObjBindMethod(This, "_Tv_LVSelectedRow"))

        This.MainFrame.Group["Thumbnail Visibility"] := Thumbnail_visibility
        for k, v in This.MainFrame.Group["Thumbnail Visibility"]
            v.Visible := 0
    }

    GameLogsMonitoring_Ctrl() {
        arr := []

        monitoredChars := ""
        for char in This.charsToMonitor {
            monitoredChars .= This.CleanTitle(char) "`n"
        }

        This.MainFrame.SetFont("s12 w700 q5")
        arr.Push This.MainFrame.Add("Text", Format("x{} y{}", This.contentGap, This.contentGap), "Game Logs Monitoring")
        This.MainFrame.SetFont("s11 w400")
        arr.Push This.MainFrame.Add("Text", Format("xp yp+{} w{} h2 +0x10", This.lGap, This.sepW))

        arr.Push This.MainFrame.Add("Text", Format("xp yp+{} Section", This.lGap), "This feature does not violate the EULA or ToS, but it is neither endorsed")
        arr.Push This.MainFrame.Add("Text", Format("xp yp+{} Section", This.contentGap), "nor supported by CCP. Use at your own risk. Performance impact")
        arr.Push This.MainFrame.Add("Text", Format("xp yp+{} Section", This.contentGap), "increases with the number of characters tracked simultaneously.")

        arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Game Logs Monitoring Enabled:")
        arr.Push This.MainFrame.Add("CheckBox", Format("xp+{} yp", This.offsetX) " vgameLogsMonitoringEnabled", "On/Off")

        arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Game Logs Directory:")
        arr.Push This.MainFrame.Add("Button", Format("xp+{} yp-{} w{} h{}", This.offsetX - 60 - This.baseGrid, 5, 60, 26) " vselectGameLogsDirectory", "Select")
        arr.Push This.MainFrame.Add("Edit", Format("xp+{} yp+{} w{}", 60 + This.baseGrid, 1, This.editW) " vgameLogsDirectory")

        arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Monitoring Interval (ms):")
        arr.Push This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{}", This.offsetX, This.editOffset, This.editW) " vmonitoringInterval")

        arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Event Priority if Several Incoming:")
        arr.Push This.MainFrame.Add("CheckBox", Format("xp+{} yp", This.offsetX) " vlastEventPriority", "Last/First")

        arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Supress for Focused Window:")
        arr.Push This.MainFrame.Add("CheckBox", Format("xp+{} yp", This.offsetX) " vsupressForFocused", "On/Off")

        arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Show Event Text:")
        arr.Push This.MainFrame.Add("CheckBox", Format("xp+{} yp", This.offsetX) " vshowEventText", "On/Off")

        arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Flash Border:")
        arr.Push This.MainFrame.Add("CheckBox", Format("xp+{} yp", This.offsetX) " vflashBorderEnabled", "On/Off")

        arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Stop Displaying Event on Switch:")
        arr.Push This.MainFrame.Add("CheckBox", Format("xp+{} yp", This.offsetX) " vstopDisplayingOnSwitch", "On/Off")
        
        arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Event Display Duration (ms):")
        arr.Push This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{}", This.offsetX, This.editOffset, This.editW) " veventDisplayDuration")

        arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Flash Border Interval (ms):")
        arr.Push This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{}", This.offsetX, This.editOffset, This.editW) " vflashBorderInterval")

        arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Monitor Only Selected Characters:")
        arr.Push This.MainFrame.Add("CheckBox", Format("xp+{} yp", This.offsetX) " vmonitorOnlySelectedChars", "On/Off")

        arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Characters to Monitor:")
        arr.Push This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{} h{}", This.offsetX - (This.editEx2W - This.editW), This.editOffset, This.editEx2W, This.editH - 40) " -Wrap vcharsToMonitor", monitoredChars)
        ImpBtn := This.MainFrame.Add("Button", Format("xp yp+{} w{}", This.editH - 40 + This.baseGrid, This.editEx2W), "Import from Launched")
        
        arr.Push ImpBtn

        This.MainFrame["gameLogsMonitoringEnabled"].OnEvent("Click", (obj, *) => EventHandler(obj))
        This.MainFrame["selectGameLogsDirectory"].OnEvent("Click", (obj, *) => EventHandler(obj))
        This.MainFrame["gameLogsDirectory"].OnEvent("Change", (obj, *) => EventHandler(obj))
        This.MainFrame["monitoringInterval"].OnEvent("Change", (obj, *) => EventHandler(obj))
        This.MainFrame["lastEventPriority"].OnEvent("Click", (obj, *) => EventHandler(obj))
        This.MainFrame["supressForFocused"].OnEvent("Click", (obj, *) => EventHandler(obj))
        This.MainFrame["showEventText"].OnEvent("Click", (obj, *) => EventHandler(obj))
        This.MainFrame["flashBorderEnabled"].OnEvent("Click", (obj, *) => EventHandler(obj))
        This.MainFrame["stopDisplayingOnSwitch"].OnEvent("Click", (obj, *) => EventHandler(obj))
        This.MainFrame["eventDisplayDuration"].OnEvent("Change", (obj, *) => EventHandler(obj))
        This.MainFrame["flashBorderInterval"].OnEvent("Change", (obj, *) => EventHandler(obj))
        This.MainFrame["monitorOnlySelectedChars"].OnEvent("Click", (obj, *) => EventHandler(obj))
        This.MainFrame["charsToMonitor"].OnEvent("Change", (obj, *) => EventHandler(obj))
        ImpBtn.OnEvent("Click", (*) => This.ImportNamesFromThumbs(This.MainFrame["charsToMonitor"]))

        EventHandler(obj) {
            if obj.name = "gameLogsMonitoringEnabled"
                This.gameLogsMonitoringEnabled := obj.value
            else if obj.name = "gameLogsDirectory"
                This.gameLogsDirectory := obj.value
            else if obj.name = "selectGameLogsDirectory" {
                dir := DirSelect(,, "Select Gamelogs folder.") ; Select directory
                if dir = "" ; Cancelled
                    return
                This.gameLogsDirectory := dir
                This.MainFrame["gameLogsDirectory"].Value := dir
            }
            else if obj.name = "monitoringInterval"
                This.monitoringInterval := obj.value
            else if obj.name = "lastEventPriority"
                This.lastEventPriority := obj.value
            else if obj.name = "supressForFocused"
                This.supressForFocused := obj.value
            else if obj.name = "showEventText"
                This.showEventText := obj.value
            else if obj.name = "flashBorderEnabled"
                This.flashBorderEnabled := obj.value
            else if obj.name = "stopDisplayingOnSwitch"
                This.stopDisplayingOnSwitch := obj.value
            else if obj.name = "eventDisplayDuration"
                This.eventDisplayDuration := obj.value
            else if obj.name = "flashBorderInterval"
                This.flashBorderInterval := obj.value
            else if obj.name = "monitorOnlySelectedChars"
                This.monitorOnlySelectedChars := obj.value
            else if obj.name = "charsToMonitor" {
                Arr := []
                for k, v in StrSplit(obj.value, "`n") {
                    char := Trim(v, "`n ")
                    if char = ""
                        continue
                    Arr.Push(This.AntiCleanTitle(char))
                }
                This.charsToMonitor := Arr
            }

            This.NeedRestart := 1
            SetTimer(This.Save_Settings_Delay_Timer, -200)
        }

        This.MainFrame.Group["Game Logs Monitoring"] := arr
        for k, v in This.MainFrame.Group["Game Logs Monitoring"]
            v.Visible := 0
    }

    MonitoredEvents_Ctrl() {
        arr := []

        monitoredEventsOrder := [ ; This added because Map() for some reason sorts itself
            "stoppedShooting",
            "underAttackByPlayer",
            "underAttackByNPC",
            "engagedWithFactionBSNPC",
            "engagedWithOfficerNPC",
            "engagedWithCapitalNPC",
            "warpDisrupted",
            "decloaked",
            "gateJumped",
            "convoRequest",
            "fleetInvited",
            "fleetWarped",
            "fleetRegrouped",
            "conduited",
            "crystalBroke",
            "miningStopped",
            "miningBayIsFull"
        ]

        editCW := 60

        This.MainFrame.SetFont("s12 w700 q5")
        arr.Push This.MainFrame.Add("Text", Format("x{} y{}", This.contentGap, This.contentGap), "Monitored Events")
        This.MainFrame.SetFont("s11 w400")
        arr.Push This.MainFrame.Add("Text", Format("xp yp+{} w{} h2 +0x10", This.lGap, This.sepW))

        arr.Push This.MainFrame.Add("Text", Format("xp yp+{} Section", This.lGap), "Enabling more options decreases performance. The effect increases")
        arr.Push This.MainFrame.Add("Text", Format("xp yp+{} Section", This.contentGap), "with the number of characters currently being tracked.")

        for event in monitoredEventsOrder {
            arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), This.monitoredEventsTexts[event] . (event = "stoppedShooting" ? " (Interval ms):" : ":"))

            if event = "stoppedShooting" {
                arr.Push This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{}", This.offsetX - 50 - This.baseGrid, This.editOffset, 50) " vshootingInterval")
                This.MainFrame["shootingInterval"].OnEvent("Change", (obj, *) => EventHandler(obj))
            }

            arr.Push This.MainFrame.Add("CheckBox", Format("xs+{} ys", This.offsetX) " vE" event, "On/Off")
            arr.Push This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{}", 60 + This.baseGrid, This.editOffset, editCW) " vC" event)
            arr.Push This.MainFrame.Add("Text", Format("xp+{} yp w{} h{}", editCW + This.baseGrid, This.cPreviewSize, This.cPreviewSize,) " vPreviewC" event " Border")
            
            This.MainFrame["E" event].OnEvent("Click", (obj, *) => EventHandler(obj))
            This.MainFrame["C" event].OnEvent("Change", (obj, *) => EventHandler(obj))
        }

        EventHandler(obj) {
            object := SubStr(obj.name, 1, 1)
            event := SubStr(obj.name, 2)
            if object = "E"
                This.monitoredEvents[event]["enabled"] := obj.value
            else if object = "C" {
                This.monitoredEvents[event]["color"] := obj.value
                This.RedrawColorPreview(obj)
            }
            else if obj.name = "shootingInterval" {
                This.shootingInterval := obj.value
            }

            This.NeedRestart := 1
            SetTimer(This.Save_Settings_Delay_Timer, -200)
        }

        This.MainFrame.Group["Monitored Events"] := arr
        for k, v in This.MainFrame.Group["Monitored Events"]
            v.Visible := 0
    }

    NonEVEApps_Ctrl() {
        This.MainFrame.Group["Non-EVE Applications"] := [], arr := []

        btnW := (This.editW - This.baseGrid) / 2
        btnEditH := 26
        editCustomW := 131
        editCustomH := 190

        This.MainFrame.SetFont("s12 w700 q5")
        arr.Push This.MainFrame.Add("Text", Format("x{} y{}", This.contentGap, This.contentGap), "Non-EVE Applications")
        This.MainFrame.SetFont("s11 w400")
        arr.Push This.MainFrame.Add("Text", Format("xp yp+{} w{} h2 +0x10", This.lGap, This.sepW))

        arr.Push This.MainFrame.Add("Text", Format("xp yp+{} Section", This.lGap), "Enter exact .exe name (optional: window title for better matching).`nThumbnail will appear, and assigned hotkey can activate the app.")

        arr.Push This.MainFrame.Add("Text", Format("xp yp+{} Section", This.lGap * 2), "Add/Delete Groups:")
        addBtn := This.MainFrame.Add("Button", Format("xp+{} yp-{} w{} h{}", This.offsetX, 5, btnW, btnEditH), "Add")
        delBtn := This.MainFrame.Add("Button", Format("xp+{} yp w{} h{}", btnW + This.baseGrid, btnW, btnEditH), "Delete")

        arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Select Group:")
        ddl := This.MainFrame.Add("DDL", Format("xp+{} yp-{} w{}", This.offsetX, 6, This.editW) " vNonEVEGroupsDDL", This.GetNonEVEGroupsList())

        arr.Push This.MainFrame.Add("Text", Format("xs yp+{} Section", This.xlGap), "Forwards Hotkey:")
        HKForwards := This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{}", This.offsetX, This.editOffset, This.editHtkW) " vNonEVEForwardsKey")
        ; arr.Push This.MainFrame.Add("Button", Format("xp+{} yp-{} w{} h{}", This.editHtkW + This.baseGrid, 1, This.captBtnW + 1, This.captBtnH) " vcapNonEVEGrHtkBtn1", "Capture")

        arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Backwards Hotkey:")
        HKBackwards := This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{}", This.offsetX, This.editOffset, This.editHtkW) " vNonEVEBackwardsdKey")
        ; arr.Push This.MainFrame.Add("Button", Format("xp+{} yp-{} w{} h{}", This.editHtkW + This.baseGrid, 1, This.captBtnW + 1, This.captBtnH) " vcapNonEVEGrHtkBtn2", "Capture")

        arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Process Names (.exe):")
        EditExe := This.MainFrame.Add("Edit", Format("xp yp+{} w{} h{}", This.contentGap, This.editExW, editCustomH) " -Wrap vNonEVEProcessGroups")

        arr.Push This.MainFrame.Add("Text", Format("xs+{} ys", This.editExW + This.contentGap), "Titles (optional):")
        EditTitle := This.MainFrame.Add("Edit", Format("xp yp+{} w{} h{}", This.contentGap, This.editExW, editCustomH) " -Wrap vNonEVETitlesGroups")

        arr.Push This.MainFrame.Add("Text", Format("xs yp+{} w{} h2 +0x10", editCustomH + This.contentGap, This.sepW))

        arr.Push This.MainFrame.Add("Text", Format("xp yp+{} Section", This.contentGap), "Process Names (.exe):")
        arr.Push This.MainFrame.Add("Edit", Format("xp yp+{} w{} h{}", This.contentGap, editCustomW, editCustomH) " -Wrap vNonEVEProcessHtks")

        arr.Push This.MainFrame.Add("Text", Format("xs+{} ys", editCustomW + This.baseGrid + 1), "Titles (optional):")
        arr.Push This.MainFrame.Add("Edit", Format("xp yp+{} w{} h{}", This.contentGap, editCustomW, editCustomH) " -Wrap vNonEVETitlesHtks")

        arr.Push This.MainFrame.Add("Text", Format("xs+{} ys", (editCustomW + This.baseGrid) * 2 + 1), "Hotkeys:")
        arr.Push This.MainFrame.Add("Edit", Format("xp yp+{} w{} h{}", This.contentGap, editCustomW, editCustomH) " -Wrap vNonEVEHotkeysHtks")

        arr.Push ddl
        arr.Push addBtn
        arr.Push delBtn
        arr.Push HKForwards
        arr.Push HKBackwards
        arr.Push EditExe
        arr.Push EditTitle

        This.MainFrame["NonEVEGroupsDDL"].OnEvent("Change", (*) => SetEditText(ddl, EditExe, EditTitle, HKForwards, HKBackwards))
        addBtn.OnEvent("Click", (*) => CreateNewGroup(ddl, HKForwards, HKBackwards, EditExe, EditTitle))
        delBtn.OnEvent("Click", (*) => Delete_Group(ddl, HKForwards, HKBackwards, EditExe, EditTitle))
        This.MainFrame["NonEVEForwardsKey"].OnEvent("Change", (obj, *) => SaveHKGroupList(obj))
        This.MainFrame["NonEVEBackwardsdKey"].OnEvent("Change", (obj, *) => SaveHKGroupList(obj))
        This.MainFrame["NonEVEProcessGroups"].OnEvent("Change", (obj, *) => SaveHKGroupList(obj))
        This.MainFrame["NonEVETitlesGroups"].OnEvent("Change", (obj, *) => SaveHKGroupList(obj))
        This.MainFrame["NonEVEProcessHtks"].OnEvent("Change", (obj, *) => SaveHKList(obj))
        This.MainFrame["NonEVETitlesHtks"].OnEvent("Change", (obj, *) => SaveHKList(obj))
        This.MainFrame["NonEVEHotkeysHtks"].OnEvent("Change", (obj, *) => SaveHKList(obj))
        ; This.MainFrame["capNonEVEGrHtkBtn1"].OnEvent("Click", (obj, *) => This.hotkeyCapture(This.MainFrame["NonEVEForwardsKey"]))
        ; This.MainFrame["capNonEVEGrHtkBtn2"].OnEvent("Click", (obj, *) => This.hotkeyCapture(This.MainFrame["NonEVEBackwardsdKey"]))

        This.MainFrame.Group["Non-EVE Applications"] := arr
        for k, v in This.MainFrame.Group["Non-EVE Applications"]
            v.Visible := 0

        CreateNewGroup(ddlObj, ForwardHKObj, BackwardHKObj, EditObj1, EditObj2) {
            ArrayIndex := 0
            Obj := InputBox("Enter a Groupname", "Create New Group", "w200 h90")
            if (Obj.Result != "OK")
                return

            temp := This.NonEVEGroups
            temp[Obj.value] := Map("exe", [], "title", [], "fkey", "", "bkey", "")
            This.NonEVEGroups := temp

            ddlObj.Delete()
            ddlObj.Add(This.GetNonEVEGroupsList())
            for k in This.NonEVEGroups {
                if k = Obj.value {
                    ArrayIndex := A_Index
                    break
                }
            }

            EditObj1.value := "", EditObj2.value := "", ForwardHKObj.value := "", BackwardHKObj.value := ""
            This.enableCtrlsInNonEVEGroupsSettings()

            ddlObj.Choose(ArrayIndex)
            This.NeedRestart := 1
            SetTimer(This.Save_Settings_Delay_Timer, -200)
        }

        Delete_Group(ddlObj, ForwardHKObj, BackwardHKObj, EditObj1, EditObj2) {
            if (ddlObj.Text != "" && This.NonEVEGroups.Has(ddlObj.Text)) {
                temp := This.NonEVEGroups
                temp.Delete(ddlObj.Text) ; Properly deletes obj
                This.NonEVEGroups := temp
            }

            ddlObj.Delete()
            ddlObj.Add(This.GetNonEVEGroupsList())
            ForwardHKObj.value := "", BackwardHKObj.value := "", EditObj1.value := "", EditObj2.value := ""
            This.enableCtrlsInNonEVEGroupsSettings(0)

            This.NeedRestart := 1
            SetTimer(This.Save_Settings_Delay_Timer, -200)
        }

        SetEditText(ddlObj, EditObj1, EditObj2, ForwardHKObj?, BackwardHKObj?) {
            if (ddlObj.Text != "" && This.NonEVEGroups.Has(ddlObj.Text)) {
                text1 := ""
                for Names in This.NonEVEGroups[ddlObj.Text]["exe"] {
                    text1 .= Names "`n"
                }
                text2 := ""
                for Names in This.NonEVEGroups[ddlObj.Text]["title"] {
                    text2 .= Names "`n"
                }
                EditObj1.value := text1, EditObj2.value := text2
                ForwardHKObj.value := This.NonEVEGroups[ddlObj.Text]["fkey"]
                BackwardHKObj.value := This.NonEVEGroups[ddlObj.Text]["bkey"]
                This.enableCtrlsInNonEVEGroupsSettings()
            }
        }

        SaveHKGroupList(obj) {
            if ddl.Text == ""
                return

            if obj.Name = "NonEVEProcessGroups" {
                Arr := []
                for k, v in StrSplit(obj.value, "`n") {
                    execs := Trim(v, "`n ")
                    if (execs = "")
                        continue
                    Arr.Push(execs)
                }
                This.NonEVEGroups[ddl.Text]["exe"] := Arr
            }
            else if obj.Name = "NonEVETitlesGroups" {
                Arr := []
                for k, v in StrSplit(obj.value, "`n") {
                    titles := Trim(v, "`n ")
                    Arr.Push(titles)
                }
                This.NonEVEGroups[ddl.Text]["title"] := Arr
            }
            else if obj.Name = "NonEVEForwardsKey" {
                This.NonEVEGroups[ddl.Text]["fkey"] := Trim(obj.value, "`n ")
            }
            else if obj.Name = "NonEVEBackwardsdKey" {
                This.NonEVEGroups[ddl.Text]["bkey"] := Trim(obj.value, "`n ")
            }
            This.NeedRestart := 1
            SetTimer(This.Save_Settings_Delay_Timer, -200)
        }

        SaveHKList(obj) {
            if obj.Name = "NonEVEProcessHtks" {
                what := "exe"
            }
            else if obj.Name = "NonEVETitlesHtks" {
                what := "title"
            }
            else if obj.Name = "NonEVEHotkeysHtks" {
                what := "hotkey"
            }

            arr := []
            for i, v in StrSplit(obj.value, "`n") {
                data := Trim(v, "`n ")
                if data = "" && what != "title" && what != "hotkey"
                    continue

                arr.Push(data)
            }

            This.NonEVEHotkeys[what] := arr

            This.NeedRestart := 1
            SetTimer(This.Save_Settings_Delay_Timer, -200)
        }
    }

    enableCtrlsInNonEVEGroupsSettings(enable := 1) {
        This.MainFrame["NonEVEForwardsKey"].Enabled := enable
        This.MainFrame["NonEVEBackwardsdKey"].Enabled := enable
        This.MainFrame["NonEVEProcessGroups"].Enabled := enable
        This.MainFrame["NonEVETitlesGroups"].Enabled := enable
        ; This.MainFrame["capNonEVEGrHtkBtn1"].Enabled := enable
        ; This.MainFrame["capNonEVEGrHtkBtn2"].Enabled := enable
    }

    Other_Ctrl() {
        This.MainFrame.Group["Other"] := [], Other := []

        This.GlobalGroupsOrder := [
            "Hotkey Groups",
            "Hotkeys Settings",
            "Thumbnails Behavior",
            "Thumbnails Visuals",
            "Thumbnail Visibility",
            "Client Settings",
            "Custom Colors",
            "Game Logs Monitoring",
            "Non-EVE Applications",
            "Monitored Events",
            "Tray Menu Settings",
            "Other"
        ]

        This.MainFrame.SetFont("s12 w700 q5")
        Other.Push This.MainFrame.Add("Text", Format("x{} y{}", This.contentGap, This.contentGap), "Other")
        This.MainFrame.SetFont("s11 w400")
        Other.Push This.MainFrame.Add("Text", Format("xp yp+{} w{} h2 +0x10", This.lGap, This.sepW))

        Other.Push This.MainFrame.Add("Text", Format("xp yp+{} Section", This.lGap), "Switch Language to English on Error:")
        Other.Push This.MainFrame.Add("CheckBox", Format("xp+{} yp", This.offsetX) " vSwitchLangOnErr Checked" This.SwitchLangOnErr, "On/Off")

        Other.Push This.MainFrame.Add("Text", Format("xs ys+{} w{} h2 +0x10", This.xlGap, This.sepW))
        Other.Push This.MainFrame.Add("Text", Format("xp yp+{} Section", This.contentGap) " vOvrLabel", "PlaceholderPlaceholderPlaceholderPlaceholderPlaceholderPlaceholder") ; If too short message will be displaying only part of text
        
        for group in This.GlobalGroupsOrder {
            group_ := StrReplace(group, A_Space, "_")
            Other.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), group ":")
            Other.Push This.MainFrame.Add("CheckBox", Format("xp+{} yp", This.offsetX) " vGlobal" group_ " Checked" This.Global_Groups[group], "On/Off")
        }

        Other.Push This.MainFrame.Add("Button", Format("xs ys+{} Section", This.xlGap) " vUpdateGls", "Update")

        Other.Push This.MainFrame.Add("Text", Format("xs ys+{} w{} h2 +0x10", This.lGap + This.objH, This.sepW))
        Other.Push This.MainFrame.Add("Text", Format("xp yp+{} Section", This.contentGap), "Update All Thumbnails Size to Default Size.")
        Other.Push This.MainFrame.Add("Button", Format("xs ys+{} Section", This.xlGap) " vUpdateThumbnails", "Update All Thumbnails")

        This.MainFrame["SwitchLangOnErr"].OnEvent("Click", (obj, *) => cOther_EventHandler(obj))
        This.MainFrame["UpdateGls"].OnEvent("Click", (obj, *) => cOther_EventHandler(obj))
        This.MainFrame["UpdateThumbnails"].OnEvent("Click", (obj, *) => cOther_EventHandler(obj))

        cOther_EventHandler(obj) {
            need_reload := 0

            if (obj.name = "SwitchLangOnErr") {
                This.SwitchLangOnErr := obj.value
            }
            else if (obj.name = "UpdateGls") {
                for group in This.GlobalGroupsOrder {
                    group_ := StrReplace(group, A_Space, "_")
                    if This.AntiGlobalGroups[group] != This.MainFrame["Global" group_].value {
                        This.AntiGlobalGroups[group] := This.MainFrame["Global" group_].value
                        need_reload := 1
                    }
                }
            }
            else if (obj.name = "UpdateThumbnails") {
                This.Update_All_Thumbnails()
                need_reload := 1
            }
            This.NeedRestart := 1
            SetTimer(This.Save_Settings_Delay_Timer, -200)
            if need_reload {
                Sleep(250)
                Reload
            }
        }

        This.MainFrame.Group["Other"] := Other
        for k, v in This.MainFrame.Group["Other"]
            v.Visible := 0
    }

    TrayMenuSettings_Ctrl() {
        arr := []

        This.TrayMenuShortcutsOrder := [
            "Suspend Hotkeys",
            "Hide Thumbnails",
            "Show Thumbnails Always on Top",
            "Click Through Thumbnails",
            "Minimize Inactive Clients",
            "Don't Close Active Client",
            "Close all EVE Clients",
            "Restore Client Positions",
            "Save Client Positions",
            "Auto Save Thumbnail Positions",
            "Save Thumbnail Positions"
        ]

        This.MainFrame.SetFont("s12 w700 q5")
        arr.Push This.MainFrame.Add("Text", Format("x{} y{}", This.contentGap, This.contentGap), "Tray Menu Settings")
        This.MainFrame.SetFont("s11 w400")
        arr.Push This.MainFrame.Add("Text", Format("xp yp+{} w{} h2 +0x10", This.lGap, This.sepW))

        arr.Push This.MainFrame.Add("Text", Format("xp yp+{} Section", This.lGap), "Show/Hide Tray Menu Items:")

        for TMI in This.TrayMenuShortcutsOrder {
            TMI_ := StrReplace(TMI, A_Space, "_")
            arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), TMI ":")
            arr.Push This.MainFrame.Add("CheckBox", Format("xp+{} yp", This.offsetX) " vTM" TMI_ " Checked" This.TrayMenuShortcuts[TMI], "On/Off")
            This.MainFrame["TM" TMI_].OnEvent("Click", (obj, *) => EventHandler(obj))
        }

        EventHandler(obj) {
            temp := StrReplace(obj.name, "TM")
            TMI := StrReplace(temp, "_", A_Space)
            This.TrayMenuShortcuts[TMI] := obj.value
            This.NeedRestart := 1
            SetTimer(This.Save_Settings_Delay_Timer, -200)
        }

        This.MainFrame.Group["Tray Menu Settings"] := arr
        for k, v in This.MainFrame.Group["Tray Menu Settings"]
            v.Visible := 0
    }

    About_Ctrl() {
        arr := []

        try
            This.programVersion := FileGetVersion(A_ScriptName)
        catch
            This.programVersion := "1.0.0.0"

        updBtnW := 158
        offsetX := 150

        This.MainFrame.SetFont("s12 w700 q5")
        arr.Push This.MainFrame.Add("Text", Format("x{} y{}", This.contentGap, This.contentGap), "About EVE-X-Preview")
        This.MainFrame.SetFont("s11 w400")
        arr.Push This.MainFrame.Add("Text", Format("xp yp+{} w{} h2 +0x10", This.lGap, This.sepW))

        arr.Push This.MainFrame.Add("Text", Format("xp yp+{}", This.lGap), "Created by gonzo83")
        arr.Push This.MainFrame.Add("Text", Format("xp yp+{}", This.lGap), "Forked and Maintained by khivus")
        arr.Push This.MainFrame.Add("Text", Format("xp yp+{}", This.lGap), "Credits to: mrmjstc and CJKondur")

        arr.Push This.MainFrame.Add("Text", Format("xp yp+{} Section", This.xlGap), "Current Version:")
        arr.Push This.MainFrame.Add("Text", Format("xp+{} yp", offsetX), "v" This.programVersion)

        arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.lGap), "Latest Release:")
        arr.Push This.MainFrame.Add("Text", Format("xp+{} yp", offsetX) " vlatestReleaseVersion", "unknown          ") ; This spaces is stupid because ahk cuts text after update if initial text is shorter than updated text

        arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.lGap), "Latest Pre-Release:")
        arr.Push This.MainFrame.Add("Text", Format("xp+{} yp", offsetX) " vlatestPreReleaseVersion", "unknown          ")

        arr.Push This.MainFrame.Add("Button", Format("xs ys+{} Section", This.xlGap) " vcheckUpdatesBtn", "Check Updates")

        arr.Push This.MainFrame.Add("Button", Format("xs ys+{} w{} Section", This.xlGap, updBtnW) " vupdateToReleaseBtn", "Update to Release")
        arr.Push This.MainFrame.Add("Button", Format("xs+{} ys w{}", updBtnW + This.baseGrid, updBtnW) " vupdateToPreReleaseBtn", "Update to Pre-Release")

        arr.Push This.MainFrame.Add("Button", Format("xs ys+{} Section", This.xlGap) " vhelpBtn", "Help")
        arr.Push This.MainFrame.Add("Button", Format("xp+{} yp Section", 48 + This.baseGrid) " vreportBugBtn", "Report Bug")

        arr.Push This.MainFrame.Add("Button", Format("x{} y{} w{} Section", This.contentW - 130 - This.contentGap, This.guiHeight - 65 - This.contentGap, 130) " vdebugModeBtn", "Debug Mode: " . (This.DebugMode ? "On" : "Off"))

        arr.Push This.MainFrame.Add("Button", Format("x{} y{} w{} Section", This.contentW - 130 - This.contentGap, This.guiHeight - 30 - This.contentGap, 130) " vfunnyBtn", "Funny" . (This.ThisThat ? "!" : "?"))

        This.MainFrame["checkUpdatesBtn"].OnEvent("Click", (obj, *) => checkForNewUpdate())
        This.MainFrame["updateToReleaseBtn"].OnEvent("Click", (obj, *) => processUpdateApp("false"))
        This.MainFrame["updateToPreReleaseBtn"].OnEvent("Click", (obj, *) => processUpdateApp("true"))
        This.MainFrame["helpBtn"].OnEvent("Click", (obj, *) => helpButtonHandler())
        This.MainFrame["reportBugBtn"].OnEvent("Click", (obj, *) => reportBugButtonHandler())
        This.MainFrame["debugModeBtn"].OnEvent("Click", (obj, *) => debugModeHandler())
        This.MainFrame["funnyBtn"].OnEvent("Click", (obj, *) => funnyHandler())

        checkForNewUpdate() {
            try {
                ; Getting json of latest release
                apiUrl := "https://api.github.com/repos/khivus/EVE-X-Preview/releases"
                whr := ComObject("WinHttp.WinHttpRequest.5.1")
                whr.Open("GET", apiUrl)
                whr.SetRequestHeader("User-Agent", "AHK")
                whr.Send()
                whr.WaitForResponse()
                json_ans := whr.ResponseText
            }
            catch {
                MsgBox("GitHub not available or no internet connection!") ; Probably no internet connection
                return
            }

            ; Finding tag of latest release
            pattern := '"tag_name":\s*"v?V?([^"]+)",[\s\S]*?"prerelease":\s*(true|false),'
            pos := 1
            while matchPos := RegExMatch(json_ans, pattern, &match, pos) {
                tag := match[1]
                preRelease := match[2]

                if preRelease = "false" && This.latestReleaseTag = ""
                    This.latestReleaseTag := tag
                else if preRelease = "true" && This.latestPreReleaseTag = ""
                    This.latestPreReleaseTag := tag

                if This.latestReleaseTag != "" && This.latestPreReleaseTag != ""
                    break

                pos := matchPos + match.Len ; Advance past this match
            }

            if This.latestReleaseTag != "" {
                This.MainFrame["latestReleaseVersion"].Value := "v" This.latestReleaseTag
                if VerCompare(This.latestReleaseTag, This.programVersion) != 0
                    This.MainFrame["updateToReleaseBtn"].Enabled := 1
            }

            if This.latestPreReleaseTag != "" {
                This.MainFrame["latestPreReleaseVersion"].Value := "v" This.latestPreReleaseTag
                if VerCompare(This.latestPreReleaseTag, This.programVersion) != 0
                    This.MainFrame["updateToPreReleaseBtn"].Enabled := 1
            }
        }

        processUpdateApp(preRelease?) {
            if !IsSet(preRelease)
                return

            if preRelease = "false"
                newTag := This.latestReleaseTag
            else if preRelease = "true"
                newTag := This.latestPreReleaseTag
            else
                return

            SetWorkingDir(A_ScriptDir)

            updaterExeName := "EVE-X-Preview-Updater.exe"
            updaterExeUrl := "https://github.com/khivus/EVE-X-Preview/releases/download/v" newTag "/" updaterExeName

            if FileExist(updaterExeName)
                FileDelete(updaterExeName)
            
            Download(updaterExeUrl, updaterExeName) ; Download file from GitHub

            if !FileExist(updaterExeName)
                Throw Error("Could not download EVE-X-Preview-Updater.exe!")
            
            This.First_Start_After_Update := 1 ; For showing update message
            SetTimer(This.Save_Settings_Delay_Timer, -200)
            Sleep 250 ; Waiting for settings to save

            Run(updaterExeName " `"" A_ScriptName "`" `"" newTag "`"")

            ExitApp
        }

        helpButtonHandler() {
            Run("https://github.com/khivus/EVE-X-Preview/blob/main/README.MD")
        }
    
        reportBugButtonHandler() {
            Run("https://github.com/khivus/EVE-X-Preview/issues/new")
        }

        debugModeHandler() {
            This.DebugMode := !This.DebugMode
            This.MainFrame["debugModeBtn"].Text := "Debug Mode: " . (This.DebugMode ? "On" : "Off")
            SetTimer(This.Save_Settings_Delay_Timer, -200)
        }

        funnyHandler() {
            This.ThisThat := !This.ThisThat
            This.MainFrame["funnyBtn"].Text := "Funny" . (This.ThisThat ? "!" : "?")
            SetTimer(This.Save_Settings_Delay_Timer, -200)
        }

        This.MainFrame.Group["About"] := arr
        for k, v in This.MainFrame.Group["About"]
            v.Visible := 0
    }

    On_WM_MOUSEMOVE(wParam, lParam, msg, Hwnd) {
        static PrevHwnd := 0
        if (Hwnd != PrevHwnd) {
            Text := "", ToolTip() ; Turn off any previous tooltip.
            CurrControl := GuiCtrlFromHwnd(Hwnd)
            if CurrControl {
                if !CurrControl.HasProp("ToolTip")
                    return ; No tooltip for this control.
                Text := CurrControl.ToolTip
                SetTimer () => ToolTip(Text), -1000
                SetTimer () => ToolTip(), -4000 ; Remove the tooltip.
            }
            PrevHwnd := Hwnd
        }
    }

    Profiles_to_Array() {
        ll := []
        for k, v in This.Profiles
            ll.Push(k)
        return ll
    }

    Dont_Minimize_List() {
        list := ""
        for k in This.Dont_Minimize_Clients {
            list .= This.CleanTitle(k) "`n"
        }
        return list
    }

    DontCloseList() {
        list := ""
        for k in This.DontCloseClients {
            list .= This.CleanTitle(k) "`n"
        }
        return list
    }

    UpdateOvrText() {
        if This.LastUsedProfile == "Default"
            ovr_explanation_text := "Selected Groups Override Settings Groups from Default Profile"
        else
            ovr_explanation_text := "Selected Groups Disables Override from Default Profile"

        This.MainFrame["OvrLabel"].Text := ovr_explanation_text
    }

    ; ; Capture a hotkey and put the AHK hotkey string into an Edit control.
    ; hotkeyCapture(editField) {
    ;     oldValue := editField.Value
    ;     editField.Value := "Press key..."
    ;     editField.Opt("+ReadOnly")
    ;     Suspend true

    ;     capturedString := ""

    ;     try {
    ;         state := {
    ;             mods: Map("Shift", false, "Ctrl", false, "Alt", false, "Win", false),
    ;             mainKey: "",
    ;             done: false
    ;         }

    ;         ih := InputHook("T5 I")
    ;         ih.KeyOpt("{All}", "N")
    ;         ih.NotifyNonText := true

    ;         ih.OnKeyDown := (ih, vk, sc) => This.CaptureKeyDown(ih, vk, sc, state)
    ;         ih.OnKeyUp   := (ih, vk, sc) => This.CaptureKeyUp(ih, vk, sc, state)

    ;         ih.Start()

    ;         ; Seed modifier state from keys already physically held when capture begins
    ;         if GetKeyState("Shift")
    ;             state.mods["Shift"] := true
    ;         if GetKeyState("Ctrl")
    ;             state.mods["Ctrl"] := true
    ;         if GetKeyState("Alt")
    ;             state.mods["Alt"] := true
    ;         if GetKeyState("LWin") || GetKeyState("RWin")
    ;             state.mods["Win"] := true

    ;         ih.Wait()

    ;         if (state.mainKey != "")
    ;             capturedString := This.BuildHotkeyString(state.mods, state.mainKey)
    ;     } finally {
    ;         Suspend false
    ;         editField.Opt("-ReadOnly")

    ;         if (capturedString != "")
    ;             editField.Value := capturedString
    ;         else
    ;             editField.Value := oldValue
    ;     }
    ; }

    ; CaptureKeyDown(ih, vk, sc, state) {
    ;     if (state.done)
    ;         return

    ;     name := GetKeyName(Format("vk{:x}sc{:x}", vk, sc))

    ;     if (name = "Escape") {
    ;         state.done    := true
    ;         state.mainKey := ""
    ;         ih.Stop()
    ;         return
    ;     }

    ;     switch vk {
    ;         case 0x10, 0xA0, 0xA1:  state.mods["Shift"] := true   ; VK_SHIFT, VK_LSHIFT, VK_RSHIFT
    ;         case 0x11, 0xA2, 0xA3:  state.mods["Ctrl"]  := true   ; VK_CONTROL, VK_LCONTROL, VK_RCONTROL
    ;         case 0x12, 0xA4, 0xA5:  state.mods["Alt"]   := true   ; VK_MENU, VK_LMENU, VK_RMENU
    ;         case 0x5B, 0x5C:         state.mods["Win"]   := true   ; VK_LWIN, VK_RWIN
    ;         default:
    ;             if (state.mainKey = "")
    ;                 state.mainKey := name
    ;     }
    ; }

    ; CaptureKeyUp(ih, vk, sc, state) {
    ;     if (state.done)
    ;         return

    ;     switch vk {
    ;         case 0x10, 0xA0, 0xA1:  return   ; Shift — keep waiting
    ;         case 0x11, 0xA2, 0xA3:  return   ; Ctrl  — keep waiting
    ;         case 0x12, 0xA4, 0xA5:  return   ; Alt   — keep waiting
    ;         case 0x5B, 0x5C:         return   ; Win   — keep waiting
    ;     }

    ;     ; Non-modifier key released — stop only if we captured a main key
    ;     if (state.mainKey != "") {
    ;         state.done := true
    ;         ih.Stop()
    ;     }
    ; }

    ; BuildHotkeyString(mods, key) {
    ;     prefix := ""
    ;     if mods["Ctrl"]   prefix .= "^"
    ;     if mods["Alt"]    prefix .= "!"
    ;     if mods["Shift"]  prefix .= "+"
    ;     if mods["Win"]    prefix .= "#"
    ;     return prefix . key
    ; }

    _Button_Load(obj?,*) {
        if (IsSet(obj))
            This.NeedRestart := 1
        
        This.LastUsedProfile := This.Sidebar["SelectedProfile"].Text
        This.ProfileOverride()
        This.Refresh_ControlValues()
        This.UpdateOvrText()

        if This.LastUsedProfile = "Default" {
            for k, ob in This.MainFrame.Group["Other"] {
                if InStr(ob.name, "Global")
                    ob.Enabled := 1
            }
        }
        else {
            for k, v in This.MainFrame.Group {
                for group, enab in This.ComboGroups {
                    if k != group || !enab
                        continue
                    if group = "Other" {
                        Loop 4
                            v[A_Index].Enabled := 0
                    }
                    else if group = "Hotkeys Settings" {
                        Loop 21
                            v[A_Index].Enabled := 0
                    }
                    else if group = "Hotkey Groups" {
                        Loop 14
                            v[A_Index].Enabled := 0
                    }
                    else {
                        for _, ob in v {
                            ob.Enabled := 0
                        }
                    }
                }
            }
            for k, ob in This.MainFrame.Group["Other"] {
                if InStr(ob.name, "Global") {
                    groupName_ := StrReplace(ob.name, "Global", "")
                    groupName := StrReplace(groupName_, "_", A_Space)
                    if This.ComboGroups[groupName] || This.AntiGlobalGroups[groupName]
                        ob.Enabled := 1
                }
            }
        }

        SetTimer(This.Save_Settings_Delay_Timer, -200)
    }

    Refresh_ControlValues() {

        ;Client Settings
        This.MainFrame["MinimizeInactiveClients"].value := This.MinimizeInactiveClients
        This.MainFrame["AlwaysMaximize"].value := This.AlwaysMaximize
        This.MainFrame["Dont_Minimize_Clients"].value := This.Dont_Minimize_List()
        This.MainFrame["Minimizeclients_Delay"].value := This.Minimizeclients_Delay
        This.MainFrame["DontCloseOnLoginScreen"].value := This.DontCloseOnLoginScreen
        This.MainFrame["dontCloseActiveClient"].value := This.dontCloseActiveClient
        This.MainFrame["DontCloseClients"].value := This.DontCloseList()

        ;Custom Colors
        This.MainFrame["Ccoloractive"].value := This.CustomColorsActive
        This.MainFrame["Cchars"].value := This.CustomColors_AllCharNames
        This.MainFrame["CBorderColor"].value := This.CustomColors_AllBColors
        This.MainFrame["CTextColor"].value := This.CustomColors_AllTColors
        This.MainFrame["IABorderColor"].value := This.CustomColors_IABorder_Colors

        ;Hotkey Groups
        This.MainFrame["PreserveHotkeysOnLogout"].value := This.PreserveHotkeysOnLogout
        This.MainFrame["KeepGroupsPositions"].value := This.KeepGroupsPositions
        This.MainFrame["dynamicGroupsEnabled"].value := This.dynamicGroupsEnabled
        This.MainFrame["dynamicGroupsColor"].value := This.dynamicGroupsColor
        This.RedrawColorPreview(This.MainFrame["dynamicGroupsColor"])
        This.MainFrame["GroupsHoldDelay"].value := This.GroupsHoldDelay
        This.MainFrame["HotkeyGroupDDL"].Delete()
        This.MainFrame["HotkeyGroupDDL"].Add(This.GetGroupList())
        This.MainFrame["ForwardsKey"].value := "", This.MainFrame["ForwardsKey"].Enabled := 0
        This.MainFrame["BackwardsdKey"].value := "", This.MainFrame["BackwardsdKey"].Enabled := 0
        This.MainFrame["HKCharlist"].value := "", This.MainFrame["HKCharlist"].Enabled := 0

        ;Hotkeys
        This.MainFrame["Suspend_Hotkeys_Hotkey"].value := This.Suspend_Hotkeys_Hotkey
        This.MainFrame["HideThumbnailsHotkey"].value := This.HideThumbnailsHotkey
        This.MainFrame["ClickThroughHotkey"].value := This.ClickThroughHotkey
        This.MainFrame["Hotkey_Scoope"].value := (This.Global_Hotkeys ? 1 : 2)
        This.MainFrame["Login_Screen_Cycle_Hotkey"].value := This.Login_Screen_Cycle_Hotkey
        This.MainFrame["LoginScreenCycleDirectionForwards"].value := This.LoginScreenCycleDirection
        This.MainFrame["LoginScreenCycleDirectionBackwards"].value := (This.LoginScreenCycleDirection ? 0 : 1)
        This.MainFrame["Close_Active_EVE_Win_Hotkey"].value := This.Close_Active_EVE_Win_Hotkey
        This.MainFrame["Close_All_EVE_Win_Hotkey"].value := This.Close_All_EVE_Win_Hotkey
        This.MainFrame["Reload_Program_Hotkey"].value := This.Reload_Program_Hotkey

        Charlist := "", Hklist := ""
        for index, value in This._Hotkeys {
            for name, hotkey in value {
                Charlist .= This.CleanTitle(name) "`n"
                Hklist .= hotkey "`n"
            }
        }

        This.MainFrame["HotkeyCharList"].value := Charlist
        This.MainFrame["HotkeyList"].value := Hklist

        ;Thumbnail Settings
        This.MainFrame["ShowThumbnailTextOverlay"].value := This.ShowThumbnailTextOverlay
        This.MainFrame["ThumbnailTextColor"].value := This.ThumbnailTextColor
        This.RedrawColorPreview(This.MainFrame["ThumbnailTextColor"])
        This.MainFrame["ThumbnailTextSize"].value := This.ThumbnailTextSize
        This.MainFrame["ThumbnailTextFont"].value := This.ThumbnailTextFont
        This.MainFrame["ThumbnailTextMarginsx"].value := This.ThumbnailTextMargins["x"]
        This.MainFrame["ThumbnailTextMarginsy"].value := This.ThumbnailTextMargins["y"]
        This.MainFrame["ClientHighligtColor"].value := This.ClientHighligtColor
        This.RedrawColorPreview(This.MainFrame["ClientHighligtColor"])
        This.MainFrame["ClientHighligtBorderthickness"].value := This.ClientHighligtBorderthickness
        This.MainFrame["ShowClientHighlightBorder"].value := This.ShowClientHighlightBorder
        This.MainFrame["AutoSaveThumbnailPositions"].value := This.AutoSaveThumbnailPositions
        This.MainFrame["HideThumbnails"].value := This.HideThumbnails
        This.MainFrame["HideThumbnailsOnLostFocus"].value := This.HideThumbnailsOnLostFocus
        This.MainFrame["ThumbnailOpacity"].value := IntegerToPercentage(This.ThumbnailOpacity)
        This.MainFrame["ShowThumbnailsAlwaysOnTop"].value := This.ShowThumbnailsAlwaysOnTop
        This.MainFrame["ClickThroughActive"].value := This.ClickThroughActive
        This.MainFrame["ShowAllBorders"].value := This.ShowAllColoredBorders
        This.MainFrame["InactiveClientBorderthickness"].value := This.InactiveClientBorderthickness
        This.MainFrame["InactiveClientBorderColor"].value := This.InactiveClientBorderColor
        This.RedrawColorPreview(This.MainFrame["InactiveClientBorderColor"])
        This.MainFrame["ThumbnailBackgroundColor"].value := This.ThumbnailBackgroundColor
        This.RedrawColorPreview(This.MainFrame["ThumbnailBackgroundColor"])
        This.MainFrame["ThumbnailStartLocationx"].value := This.ThumbnailStartLocation["x"]
        This.MainFrame["ThumbnailStartLocationy"].value := This.ThumbnailStartLocation["y"]
        This.MainFrame["ThumbnailStartLocationwidth"].value := This.ThumbnailStartLocation["width"]
        This.MainFrame["ThumbnailStartLocationheight"].value := This.ThumbnailStartLocation["height"]
        This.MainFrame["ThumbnailMinimumSizewidth"].value := This.ThumbnailMinimumSize["width"]
        This.MainFrame["ThumbnailMinimumSizeheight"].value := This.ThumbnailMinimumSize["height"]
        This.MainFrame["ThumbnailSnap"].value := This.ThumbnailSnap
        This.MainFrame["ThumbnailSnap_Distance"].value := This.ThumbnailSnap_Distance
        This.MainFrame["PreserveThumbPosOnLogout"].value := This.PreserveThumbPosOnLogout
        This.MainFrame["PreserveCharNameOnLogout"].value := This.PreserveCharNameOnLogout
        This.MainFrame["HideThumbForActiveWin"].value := This.HideThumbForActiveWin
        This.MainFrame["ShiftThumbsForLoginScreen"].value := This.ShiftThumbsForLoginScreen
        This.MainFrame["ShiftThumbsCollisionCheck"].value := This.ShiftThumbsCollisionCheck
        current_index := This.directions.Get(This.ShiftThumbsDirection, 8)
        This.MainFrame["ShiftThumbsDirection"].Choose(current_index)
        This.MainFrame["ShiftThumbHorizontalStep"].value := This.ShiftThumbHorizontalStep
        This.MainFrame["ShiftThumbVerticalStep"].value := This.ShiftThumbVerticalStep

        ;Thumbnail Visibility
        This.MainFrame["Visibility_List"].Delete()
        for k, v in This.compare_openclients_with_list() {
            if (k != "EVE" || v != "") {
                if This.Thumbnail_visibility.Has(v)
                    This.Tv_LV.Add("Check", v,)
                else
                    This.Tv_LV.Add("", v,)
            }
        }

        ; Game Logs Monitoring
        This.MainFrame["gameLogsMonitoringEnabled"].value := This.gameLogsMonitoringEnabled
        This.MainFrame["monitoringInterval"].value := This.monitoringInterval
        This.MainFrame["gameLogsDirectory"].value := This.gameLogsDirectory
        This.MainFrame["monitorOnlySelectedChars"].value := This.monitorOnlySelectedChars
        This.MainFrame["lastEventPriority"].value := This.lastEventPriority
        This.MainFrame["supressForFocused"].value := This.supressForFocused
        This.MainFrame["showEventText"].value := This.showEventText
        This.MainFrame["flashBorderEnabled"].value := This.flashBorderEnabled
        This.MainFrame["stopDisplayingOnSwitch"].value := This.stopDisplayingOnSwitch
        This.MainFrame["eventDisplayDuration"].value := This.eventDisplayDuration
        This.MainFrame["flashBorderInterval"].value := This.flashBorderInterval

        ; Monitored Events
        for event, v in This.monitoredEvents {
            This.MainFrame["E" event].value := This.monitoredEvents[event]["enabled"]
            This.MainFrame["C" event].value := This.monitoredEvents[event]["color"]
            This.RedrawColorPreview(This.MainFrame["C" event])
        }
        This.MainFrame["shootingInterval"].value := This.shootingInterval

        ; Non-EVE Applications
        This.MainFrame["NonEVEGroupsDDL"].Delete()
        This.MainFrame["NonEVEGroupsDDL"].Add(This.GetNonEVEGroupsList())
        This.MainFrame["NonEVEForwardsKey"].value := ""
        This.MainFrame["NonEVEBackwardsdKey"].value := ""
        This.MainFrame["NonEVEProcessGroups"].value := ""
        This.MainFrame["NonEVETitlesGroups"].value := ""

        t1 := "", t2 := "", t3 := ""
        for i, v in This.NonEVEHotkeys["exe"]
            t1 .= v "`n"
        for i, v in This.NonEVEHotkeys["title"]
            t2 .= v "`n"
        for i, v in This.NonEVEHotkeys["hotkey"]
            t3 .= v "`n"

        This.MainFrame["NonEVEProcessHtks"].value := t1
        This.MainFrame["NonEVETitlesHtks"].value := t2
        This.MainFrame["NonEVEHotkeysHtks"].value := t3

        ; Other
        This.MainFrame["SwitchLangOnErr"].value := This.SwitchLangOnErr

        for group in This.GlobalGroupsOrder {
            group_ := StrReplace(group, A_Space, "_")
            if This.Global_Groups[group] && This.LastUsedProfile != "Default"
                This.MainFrame["Global" group_].Value := This.AntiGlobalGroups[group]
            else
                This.MainFrame["Global" group_].Value := This.Global_Groups[group]
        }

        ; Tray Menu Settings
        for TMI in This.TrayMenuShortcutsOrder {
            TMI_ := StrReplace(TMI, A_Space, "_")
            This.MainFrame["TM" TMI_].Value := This.TrayMenuShortcuts[TMI]
        }

        for k, v in This.MainFrame.Group {
            for _, ob in v {
                ob.Enabled := 1
            }
        }
        This.enableCtrlsInGroupsSettings(0)
        This.enableCtrlsInNonEVEGroupsSettings(0)

        This.MainFrame["updateToReleaseBtn"].Enabled := 0
        This.MainFrame["updateToPreReleaseBtn"].Enabled := 0

        This.MainFrame["InactiveClientBorderthickness"].Enabled := This.ShowAllColoredBorders
        This.MainFrame["InactiveClientBorderColor"].Enabled := This.ShowAllColoredBorders

        for group in This.GlobalGroupsOrder {
            group_ := StrReplace(group, A_Space, "_")
            This.MainFrame["Global" group_].Enabled := 0
        }
    }


    compare_openclients_with_list() {
        EvENameList := []
        for EveHwnd in This.ThumbWindows.OwnProps() {
            try {
                if title := WinGetTitle("Ahk_Id " EveHwnd) = "EVE" {
                    continue
                }
                EvENameList.Push WinGetTitle("Ahk_Id " EveHwnd)
            }
        }
        return EvENameList
    }


    GetGroupList() {
        List := []
        if (IsObject(This.Hotkey_Groups)) {
            for k in This.Hotkey_Groups {
                List.Push(k)
            }
            return List
        }
        else
            return []
    }
}
;Class End


IntegerToPercentage(integerValue) {
    percentage := (integerValue < 0 ? 0 : integerValue > 255 ? 100 : Round(integerValue * 100 / 255))
    return percentage
}


CompareArrays(arr1, arr2) {
    commonValues := {}

    for _, value in arr1 {
        if (IsInArray(value, arr2))
            commonValues.%value% := 1
        else
            commonValues.%value% := 0
    }

    for _, value in arr2 {
        if (!IsInArray(value, arr1))
            commonValues.%value% := 0
    }

    return commonValues
}

IsInArray(value, arr) {
    for _, item in arr {
        if (item = value)
            return true
    }
    return false
}
