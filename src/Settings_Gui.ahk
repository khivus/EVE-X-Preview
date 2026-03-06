Class Settings_Gui {

    MainGui() {
        ;if settings got chnaged which require a restart to apply
        This.NeedRestart := 1

        SetControlDelay(-1)
        ; This.S_Gui := Gui("+OwnDialogs +MinimizeBox -Resize -MaximizeBox SysMenu +MinSize500x250")
        This.S_Gui := Gui("+OwnDialogs -MinimizeBox -Resize -MaximizeBox SysMenu")
        This.S_Gui.Title := "EVE-X-Preview - Settings"

        This.SetState()
        
        This.CreateSidebar()
        This.CreateMainFrame()

        This.ClientSettings_Ctrl()
        This.CustomColors_Ctrl()
        This.HotkeyGroups_Ctrl()
        This.Hotkeys_Ctrl()
        This.ThumbnailsBehavior_Ctrl()
        This.ThumbnailsVisuals_Ctrl()
        This.ThumbnailVisibility_Ctrl()
        This.Other_Ctrl()

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
        This.baseGrid := 8
        This.contentGap := 16

        This.guiWidth := 700
        This.guiHeight := 630

        This.btnH := 32
        This.objH := 20

        This.sidebarW := 240
        This.sidebarInnerW := This.sidebarW - This.contentGap * 2
        This.sidebar3BtnW := (This.sidebarInnerW - This.contentGap) / 3

        This.contentW := This.guiWidth - This.sidebarW

        This.lGap := 24
        This.xlGap := 32
        This.editW := 160
        This.editH := 200
        This.editOffset := 3
        This.offsetX := 250
        This.editExH := 300
        This.sepW := This.contentW - This.contentGap * 2 + 2

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

        for index, btn_text in This._ProfileProps {
            if index == 1 {
                btn := This.Sidebar.Add("Button", Format("xp yp+{} w{} h{}", This.contentGap, This.sidebarInnerW, This.btnH), btn_text)
                This.firstBtn := btn
            }
            else
                btn := This.Sidebar.Add("Button", Format("xp yp+{} w{} h{}", This.btnH + This.baseGrid, This.sidebarInnerW, This.btnH), btn_text)
            btn.OnEvent("Click", (Obj, *) => This.SettingsGroup_Handler(Obj))
        }

        This.Sidebar.Add("Button", Format("xp y{} w{} h{}", This.guiHeight - This.contentGap - This.btnH, This.sidebarInnerW, This.btnH) " vAbout_Button", "About").OnEvent("Click", (*) => This.About_Button_Handler())
        This.Sidebar.Add("Button", Format("xp yp-{} w{} h{}", This.btnH + This.baseGrid, (This.sidebarInnerW - This.baseGrid) / 2, This.btnH) " vHelp_Button", "Help").OnEvent("Click", (*) => This.Help_Button_Handler())
        This.Sidebar.Add("Button", Format("xp+{} yp w{} h{}", (This.sidebarInnerW - This.baseGrid) / 2 + This.baseGrid, (This.sidebarInnerW - This.baseGrid) / 2, This.btnH) " vReportBugBtn", "Report Bug").OnEvent("Click", (*) => This.Report_Bug_Button_Handler())

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

    About_Button_Handler() {
        Version := FileGetVersion(A_ScriptName)
        static text := "EVE-X-Preview v" Version "`n`nCreated by gonzo83`nForked by khivus`n"
        funny := "`nF`nu`nn`nn`ny"
        text_len := StrLen(text)
        funny_len := StrLen(funny)

        if This.ThisThat && SubStr(text, text_len - funny_len + 1, funny_len) != funny
            text := text . funny

        what := MsgBox(text, "EVE-X-Preview - About", "CancelTryAgainContinue Iconi")

        if what == "Cancel" || what == "Continue"
            SetTimer(This.Save_Settings_Delay_Timer, -200)
        else if what == "TryAgain" {
            This.ThisThat := !This.ThisThat

            if not This.ThisThat
                text := SubStr(text, 1, text_len - 1)

            This.About_Button_Handler()
        }
    }

    Help_Button_Handler() {
        Run("https://github.com/khivus/EVE-X-Preview/blob/main/README.MD")
    }

    Report_Bug_Button_Handler() {
        Run("https://github.com/khivus/EVE-X-Preview/issues/new")
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

        ClientSettings.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Don't Minimize Clients:")
        ClientSettings.Push This.MainFrame.Add("Edit", Format("xs yp+{} w{} h{}", This.contentGap, This.editW, This.editExH) " vDont_Minimize_Clients -Wrap", This.Dont_Minimize_List())
        ImpBtn1 := This.MainFrame.Add("Button", Format("xp yp+{} w{}", This.editExH + This.baseGrid, This.editW), "Import from Launched")

        ClientSettings.Push This.MainFrame.Add("Text", Format("xs+{} ys Section", This.offsetX), "Don't Close Clients:")
        ClientSettings.Push This.MainFrame.Add("Edit", Format("xp yp+{} w{} h{}", This.contentGap, This.editW, This.editExH) " vDontCloseClients -Wrap", This.DontCloseList())
        ImpBtn2 := This.MainFrame.Add("Button", Format("xp yp+{} w{}", This.editExH + This.baseGrid, This.editW), "Import from Launched")

        This.MainFrame["MinimizeInactiveClients"].OnEvent("Click", (obj, *) => cSettings_EventHandler(obj))
        This.MainFrame["AlwaysMaximize"].OnEvent("Click", (obj, *) => cSettings_EventHandler(obj))
        This.MainFrame["Minimizeclients_Delay"].OnEvent("Change", (obj, *) => cSettings_EventHandler(obj))
        This.MainFrame["Dont_Minimize_Clients"].OnEvent("Change", (obj, *) => cSettings_EventHandler(obj))
        ClientSettings.Push ImpBtn1
        ImpBtn1.OnEvent("Click", (*) => This.ImportNamesFromThumbs(This.MainFrame["Dont_Minimize_Clients"]))
        This.MainFrame["DontCloseOnLoginScreen"].OnEvent("Click", (obj, *) => cSettings_EventHandler(obj))
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
            if (obj.name = "DontCloseOnLoginScreen") {
                This.DontCloseOnLoginScreen := obj.value
            }
            else if (obj.name = "DontCloseClients") {
                This.DontCloseClients := obj.value
            }
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
        CustomColors.Push This.MainFrame.Add("Edit", Format("xp ys+{} w{} h{}", This.contentGap, This.editW, This.editH) " -Wrap vCchars", This.CustomColors_AllCharNames)

        ImpBtn := This.MainFrame.Add("Button", Format("xp yp+{} w{}", This.editH + This.baseGrid, This.editW), "Import from Launched")

        CustomColors.Push This.MainFrame.Add("Text", Format("xs+{} ys Section", This.offsetX), "Active Border Color:")
        CustomColors.Push This.MainFrame.Add("Edit", Format("xp ys+{} w{} h{}", This.contentGap, This.editW, This.editH) " -Wrap vCBorderColor", This.CustomColors_AllBColors)

        CustomColors.Push This.MainFrame.Add("Text", Format("x{} ys+{} Section", This.contentGap, This.editH + This.xlGap + This.btnH), "Text Color:")
        CustomColors.Push This.MainFrame.Add("Edit", Format("xp ys+{} w{} h{}", This.contentGap, This.editW, This.editH) " -Wrap vCTextColor", This.CustomColors_AllTColors)

        CustomColors.Push This.MainFrame.Add("Text", Format("xs+{} ys Section", This.offsetX), "Inactive Border Color:")
        CustomColors.Push This.MainFrame.Add("Edit", Format("xp ys+{} w{} h{}", This.contentGap, This.editW, This.editH) " -Wrap vIABorderColor", This.CustomColors_IABorder_Colors)

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
                This.NeedRestart := 1
            }
            else if (obj.Name = "CBorderColor") {
                indexOld := This.IndexcBorder
                This.CustomColors_AllBColors := obj.value
                if (indexOld < This.IndexcBorder) {
                    obj.value := This.CustomColors_AllBColors
                    ControlSend("^{End}", obj.Hwnd)
                }
                This.NeedRestart := 1
            }
            else if (obj.Name = "CTextColor") {
                indexOld := This.IndexcText
                This.CustomColors_AllTColors := obj.value
                if (indexOld < This.IndexcText) {
                    obj.value := This.CustomColors_AllTColors
                    ControlSend("^{End}", obj.Hwnd)
                }
                This.NeedRestart := 1
            }            
            else if (obj.Name = "IABorderColor") {
                indexOld := This.IndexcText
                This.CustomColors_IABorder_Colors := obj.value
                if (indexOld < This.IndexcText) {
                    obj.value := This.CustomColors_IABorder_Colors
                    ControlSend("^{End}", obj.Hwnd)
                }
                This.NeedRestart := 1
            }            
            SetTimer(This.Save_Settings_Delay_Timer, -200)
        }
    }

    HotkeyGroups_Ctrl() {
        This.MainFrame.Group["Hotkey Groups"] := [], Hotkey_Groups := []

        btnW := (This.editW - This.baseGrid) / 2
        btnEditH := 26


        This.MainFrame.SetFont("s12 w700 q5")
        Hotkey_Groups.Push This.MainFrame.Add("Text", Format("x{} y{}", This.contentGap, This.contentGap), "Hotkey Groups")
        This.MainFrame.SetFont("s11 w400")
        Hotkey_Groups.Push This.MainFrame.Add("Text", Format("xp yp+{} w{} h2 +0x10", This.lGap, This.sepW))

        Hotkey_Groups.Push This.MainFrame.Add("Text", Format("xp yp+{} Section", This.lGap), "Preserve Hotkeys on Logout:")
        Hotkey_Groups.Push This.MainFrame.Add("CheckBox", Format("xp+{} yp", This.offsetX) " vPreserveHotkeysOnLogout Checked" This.PreserveHotkeysOnLogout, "On/Off")

        Hotkey_Groups.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Keep Groups Positions:")
        Hotkey_Groups.Push This.MainFrame.Add("CheckBox", Format("xp+{} yp", This.offsetX) " vKeepGroupsPositions Checked" This.KeepGroupsPositions, "On/Off")

        Hotkey_Groups.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Add/Delete Groups:")
        addBtn := This.MainFrame.Add("Button", Format("xp+{} yp-{} w{} h{}", This.offsetX, 5, btnW, btnEditH), "Add")
        DeleteButton := This.MainFrame.Add("Button", Format("xp+{} yp w{} h{}", btnW + This.baseGrid, btnW, btnEditH), "Delete")

        Hotkey_Groups.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Select Group:")
        ddl := This.MainFrame.Add("DDL", Format("xp+{} yp-{} w{}", This.offsetX, This.editOffset, This.editW) " vHotkeyGroupDDL", This.GetGroupList())

        Hotkey_Groups.Push This.MainFrame.Add("Text", Format("xs yp+{} Section", This.xlGap), "Forwards Hotkey:")
        HKForwards := This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{}", This.offsetX, This.editOffset, This.editW) " Disabled vForwardsKey")

        Hotkey_Groups.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Backwards Hotkey:")
        HKBackwards := This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{}", This.offsetX, This.editOffset, This.editW) " Disabled vBackwardsdKey")

        Hotkey_Groups.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Characters List:")
        EditBox := This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{} h{}", This.offsetX, This.editOffset, This.editW, This.editExH) " -Wrap Disabled vHKCharlist")

        ImportBtn := This.MainFrame.Add("Button", Format("xp yp+{} w{}", This.editExH + This.baseGrid, This.editW) " Disabled vImpNamesBtn", "Import from Launched")

        Hotkey_Groups.Push ddl
        Hotkey_Groups.Push addBtn
        Hotkey_Groups.Push DeleteButton
        Hotkey_Groups.Push HKForwards
        Hotkey_Groups.Push HKBackwards
        Hotkey_Groups.Push EditBox
        Hotkey_Groups.Push ImportBtn

        This.MainFrame["PreserveHotkeysOnLogout"].OnEvent("Click", (obj, *) => EventHandler(obj))
        This.MainFrame["KeepGroupsPositions"].OnEvent("Click", (obj, *) => EventHandler(obj))
        This.MainFrame["HotkeyGroupDDL"].OnEvent("Change", (*) => SetEditText(ddl, EditBox, HKForwards, HKBackwards, ImportBtn))
        addBtn.OnEvent("Click", (*) => CreateNewGroup(ddl, HKForwards, HKBackwards, EditBox))
        DeleteButton.OnEvent("Click", (*) => Delete_Group(ddl, HKForwards, HKBackwards, EditBox))
        This.MainFrame["ForwardsKey"].OnEvent("Change", (obj, *) => SaveHKGroupList(obj))
        This.MainFrame["BackwardsdKey"].OnEvent("Change", (obj, *) => SaveHKGroupList(obj))
        This.MainFrame["HKCharlist"].OnEvent("Change", (obj, *) => SaveHKGroupList(obj))
        ImportBtn.OnEvent("Click", (*) => This.ImportNamesFromThumbs(EditBox))

        This.MainFrame.Group["Hotkey Groups"] := Hotkey_Groups
        for k, v in This.MainFrame.Group["Hotkey Groups"]
            v.Visible := 0

        EventHandler(obj) {
            if (obj.name = "PreserveHotkeysOnLogout") {
                This.PreserveHotkeysOnLogout := obj.value
                This.NeedRestart := 1
            }
            else if (obj.name = "KeepGroupsPositions") {
                This.KeepGroupsPositions := obj.value
                This.NeedRestart := 1
            }
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
            ForwardHKObj.Enabled := 1, BackwardHKObj.Enabled := 1, EditObj.Enabled := 1
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
            ForwardHKObj.Enabled := 0, BackwardHKObj.Enabled := 0, EditObj.Enabled := 0
            This.NeedRestart := 1

            SetTimer(This.Save_Settings_Delay_Timer, -200)
        }

        SetEditText(ddlObj, EditObj, ForwardHKObj?, BackwardHKObj?, ImpBtn?) {
            text := ""
            if (ddlObj.Text != "" && This.Hotkey_Groups.Has(ddlObj.Text)) {
                for index, Names in This.Hotkey_Groups[ddlObj.Text]["Characters"] {
                    text .= Names "`n"
                }
                EditObj.value := text, EditObj.Enabled := 1
                ForwardHKObj.value := This.Hotkey_Groups[ddlObj.Text]["ForwardsHotkey"], ForwardHKObj.Enabled := 1
                BackwardHKObj.value := This.Hotkey_Groups[ddlObj.Text]["BackwardsHotkey"], BackwardHKObj.Enabled := 1
                ImpBtn.Enabled := 1
            }
        }

        SaveHKGroupList(obj) {
            if (obj.Name = "HKCharlist" && ddl.Text != "") {
                Arr := []
                for k, v in StrSplit(obj.value, "`n") {
                    Chars := Trim(v, "`n ")
                    if (Chars = "")
                        continue
                    Arr.Push(Chars)
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

    Hotkeys_Ctrl() {
        This.MainFrame.Group["Hotkeys"] := [], Hotkeys := []

        Charlist := "", Hklist := ""
        for index, value in This._Hotkeys {
            for name, hotkey in value {
                Charlist .= name "`n"
                Hklist .= hotkey "`n"
            }
        }

        This.MainFrame.SetFont("s12 w700 q5")
        Hotkeys.Push This.MainFrame.Add("Text", Format("x{} y{}", This.contentGap, This.contentGap), "Hotkeys")
        This.MainFrame.SetFont("s11 w400")
        Hotkeys.Push This.MainFrame.Add("Text", Format("xp yp+{} w{} h2 +0x10", This.lGap, This.sepW))

        Hotkeys.Push This.MainFrame.Add("Text", Format("xp yp+{} Section", This.lGap), "Suspend All Hotkeys - Hotkey:")
        Hotkeys.Push This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{}", This.offsetX, This.editOffset, This.editW) " vSuspend_Hotkeys_Hotkey", This.Suspend_Hotkeys_Hotkey)

        Hotkeys.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Hotkey Activation Scope:")
        Hotkeys.Push This.MainFrame.Add("DDL", Format("xp+{} yp-{} w{}", This.offsetX, This.editOffset, This.editW) " vTTT vHotkey_Scoope Choose" (This.Global_Hotkeys ? 1 : 2), ["Global", "If an EVE window is Active"])

        Hotkeys.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Cycle Login Screens - Hotkey:")
        Hotkeys.Push This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{}", This.offsetX, This.editOffset, This.editW) " vLogin_Screen_Cycle_Hotkey", This.Login_Screen_Cycle_Hotkey)

        Hotkeys.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Login Screen Cycle Direction:")
        Hotkeys.Push This.MainFrame.Add("Radio", Format("xp+{} yp", This.offsetX + 1) " vLoginScreenCycleDirectionForwards Checked" This.LoginScreenCycleDirection, "Old->New")
        Hotkeys.Push This.MainFrame.Add("Radio", Format("xp+{} yp", 83) " vLoginScreenCycleDirectionBackwards Checked" (This.LoginScreenCycleDirection ? 0 : 1), "New->Old")

        Hotkeys.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Close Active EVE Window - Hotkey:")
        Hotkeys.Push This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{}", This.offsetX, This.editOffset, This.editW) " vClose_Active_EVE_Win_Hotkey", This.Close_Active_EVE_Win_Hotkey)

        Hotkeys.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Close All EVE Windows - Hotkey:")
        Hotkeys.Push This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{}", This.offsetX, This.editOffset, This.editW) " vClose_All_EVE_Win_Hotkey", This.Close_All_EVE_Win_Hotkey)

        Hotkeys.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Reload EVE-X-Preview - Hotkey:")
        Hotkeys.Push This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{}", This.offsetX, This.editOffset, This.editW) " vReload_Program_Hotkey", This.Reload_Program_Hotkey)

        Hotkeys.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap - 3), "Character Name:")
        HKCharList := This.MainFrame.Add("Edit", Format("xp yp+{} w{} h{}", This.contentGap, This.editW, This.editH) " -Wrap vHotkeyCharList", Charlist)
        Hotkeys.Push HKCharList

        ImpBtn := This.MainFrame.Add("Button", Format("xp yp+{} w{}", This.editH + This.baseGrid, This.editW), "Import from Launched")

        Hotkeys.Push This.MainFrame.Add("Text", Format("xs+{} ys Section", This.offsetX), "Hotkeys:")
        HKKeylist := This.MainFrame.Add("Edit", Format("xp yp+{} w{} h{}", This.contentGap, This.editW, This.editH) " -Wrap vHotkeyList", Hklist)
        Hotkeys.Push HKKeylist

        This.MainFrame["Suspend_Hotkeys_Hotkey"].OnEvent("Change", (obj, *) => cHotkeys_EventHandler(obj))
        This.MainFrame["Hotkey_Scoope"].OnEvent("Change", (obj, *) => cHotkeys_EventHandler(obj))
        This.MainFrame["Login_Screen_Cycle_Hotkey"].OnEvent("Change", (obj, *) => cHotkeys_EventHandler(obj))
        This.MainFrame["LoginScreenCycleDirectionForwards"].OnEvent("Click", (obj, *) => cHotkeys_EventHandler(obj))
        This.MainFrame["LoginScreenCycleDirectionBackwards"].OnEvent("Click", (obj, *) => cHotkeys_EventHandler(obj))
        This.MainFrame["Close_Active_EVE_Win_Hotkey"].OnEvent("Change", (obj, *) => cHotkeys_EventHandler(obj))
        This.MainFrame["Close_All_EVE_Win_Hotkey"].OnEvent("Change", (obj, *) => cHotkeys_EventHandler(obj))
        This.MainFrame["Reload_Program_Hotkey"].OnEvent("Change", (obj, *) => cHotkeys_EventHandler(obj))
        HKCharList.OnEvent("Change", (obj, *) => EventHandler(obj))
        Hotkeys.Push ImpBtn
        ImpBtn.OnEvent("Click", (*) => This.ImportNamesFromThumbs(HKCharList))
        HKKeylist.OnEvent("Change", (obj, *) => EventHandler(obj))

        This.MainFrame.Group["Hotkeys"] := Hotkeys
        for k, v in This.MainFrame.Group["Hotkeys"]
            v.Visible := 0

        cHotkeys_EventHandler(obj) {
            if (obj.name = "Suspend_Hotkeys_Hotkey") {
                This.Suspend_Hotkeys_Hotkey := Trim(obj.value, "`n ")
                This.NeedRestart := 1
            }
            else if (obj.name = "Hotkey_Scoope") {
                This.Global_Hotkeys := (obj.value = 1 ? 1 : 0)
                This.NeedRestart := 1
            }
            
            else if (obj.name = "Login_Screen_Cycle_Hotkey") {
                This.Login_Screen_Cycle_Hotkey := Trim(obj.value, "`n ")
                This.NeedRestart := 1
            }
            else if (obj.name = "LoginScreenCycleDirectionForwards") {
                This.LoginScreenCycleDirection := 1
                This.NeedRestart := 1
            }
            else if (obj.name = "LoginScreenCycleDirectionBackwards") {
                This.LoginScreenCycleDirection := 0
                This.NeedRestart := 1
            }
            else if (obj.name = "Close_Active_EVE_Win_Hotkey") {
                This.Close_Active_EVE_Win_Hotkey := Trim(obj.value, "`n ")
                This.NeedRestart := 1
            }
            else if (obj.name = "Close_All_EVE_Win_Hotkey") {
                This.Close_All_EVE_Win_Hotkey := Trim(obj.value, "`n ")
                This.NeedRestart := 1
            }
            else if (obj.name = "Reload_Program_Hotkey") {
                This.Reload_Program_Hotkey := Trim(obj.value, "`n ")
                This.NeedRestart := 1
            }
            SetTimer(This.Save_Settings_Delay_Timer, -200)
        }

        ;Parse All hotkeys to a Array on value change
        EventHandler(obj) {
            tempvar := []
            ListChars := StrSplit(This.MainFrame["HotkeyCharList"].value, "`n"), ListHotkeys := StrSplit(This.MainFrame["HotkeyList"].value, "`n")
            for k, v in ListChars {
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
                tempvar.Push Map(chars, keys)
            }
            this._Hotkeys := tempvar
            This.NeedRestart := 1
            SetTimer(This.Save_Settings_Delay_Timer, -200)
        }
    }

    ThumbnailsBehavior_Ctrl() {
        This.MainFrame.Group["Thumbnails Behavior"] := [], arr := []

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

        arr.Push This.MainFrame.Add("Text", Format("xp yp+{} Section", This.lGap), "Hide Thumbnails on Lost Focus:")
        arr.Push This.MainFrame.Add("CheckBox", Format("xp+{} yp", This.offsetX) " vHideThumbnailsOnLostFocus Checked" This.HideThumbnailsOnLostFocus, "On/Off")

        arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Show Thumbnails Always on Top:")
        arr.Push This.MainFrame.Add("CheckBox", Format("xp+{} yp", This.offsetX) " vShowThumbnailsAlwaysOnTop Checked" This.ShowThumbnailsAlwaysOnTop, "On/Off")

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

        This.MainFrame["HideThumbnailsOnLostFocus"].OnEvent("Click", (obj, *) => EventHandler(obj))
        This.MainFrame["ShowThumbnailsAlwaysOnTop"].OnEvent("Click", (obj, *) => EventHandler(obj))
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
            if (obj.name = "HideThumbnailsOnLostFocus") {
                This.HideThumbnailsOnLostFocus := obj.value
            }
            else if (obj.name = "ShowThumbnailsAlwaysOnTop") {
                This.ShowThumbnailsAlwaysOnTop := obj.value
                This.NeedRestart := 1
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
                This.NeedRestart := 1
            }
            else if (obj.name = "PreserveCharNameOnLogout") {
                This.PreserveCharNameOnLogout := obj.value
                This.NeedRestart := 1
            }
            else if (obj.name = "HideThumbForActiveWin") {
                This.HideThumbForActiveWin := obj.value
                This.NeedRestart := 1
            }
            else if (obj.name = "ShiftThumbsForLoginScreen") {
                This.ShiftThumbsForLoginScreen := obj.value
                This.NeedRestart := 1
            }
            else if (obj.name = "ShiftThumbsCollisionCheck") {
                This.ShiftThumbsCollisionCheck := obj.value
                This.NeedRestart := 1
            }
            else if (obj.name = "ShiftThumbsDirection") {
                This.ShiftThumbsDirection := obj.Value
                This.NeedRestart := 1
            }
            else if (obj.name = "ShiftThumbHorizontalStep") {
                This.ShiftThumbHorizontalStep := obj.value
                This.NeedRestart := 1
            }
            else if (obj.name = "ShiftThumbVerticalStep") {
                This.ShiftThumbVerticalStep := obj.value
                This.NeedRestart := 1
            }
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
        arr.Push This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{}", This.offsetX, This.editOffset, This.editW) " vThumbnailTextColor -Wrap", This.ThumbnailTextColor)

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
        arr.Push This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{}", This.offsetX, This.editOffset, This.editW) " vClientHighligtColor -Wrap", This.ClientHighligtColor)

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
        arr.Push This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{}", This.offsetX, This.editOffset, This.editW) " vInactiveClientBorderColor -Wrap", This.InactiveClientBorderColor)

        arr.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Thumbnail Background Color (Hex/RGB):")
        arr.Push This.MainFrame.Add("Edit", Format("xp+{} yp-{} w{}", This.offsetX, This.editOffset, This.editW) " vThumbnailBackgroundColor", This.ThumbnailBackgroundColor)

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
                This.NeedRestart := 1
            }
            else if (obj.name = "ThumbnailTextColor") {
                This.ThumbnailTextColor := obj.value
                This.NeedRestart := 1
            }
            else if (obj.name = "ThumbnailTextSize") {
                This.ThumbnailTextSize := obj.value
                This.NeedRestart := 1
            }
            else if (obj.name = "ThumbnailTextFont") {
                This.ThumbnailTextFont := obj.value
                This.NeedRestart := 1
            }
            else if (obj.name = "ThumbnailTextMarginsx") {
                This.ThumbnailTextMargins["x"] := obj.value
                This.NeedRestart := 1
            }
            else if (obj.name = "ThumbnailTextMarginsy") {
                This.ThumbnailTextMargins["y"] := obj.value
                This.NeedRestart := 1
            }
            else if (obj.name = "ClientHighligtColor") {
                This.ClientHighligtColor := obj.value
                This.NeedRestart := 1
            }
            else if (obj.name = "ClientHighligtBorderthickness") {
                This.ClientHighligtBorderthickness := obj.value
                This.NeedRestart := 1
            }
            else if (obj.name = "ShowClientHighlightBorder") {
                This.ShowClientHighlightBorder := obj.value
            }
            else if (obj.name = "ThumbnailOpacity") {
                This.ThumbnailOpacity := obj.value
                This.NeedRestart := 1
            }
            else if (obj.Name = "ShowAllBorders") {
                This.ShowAllColoredBorders := obj.value
                This.MainFrame["InactiveClientBorderthickness"].Enabled := This.ShowAllColoredBorders
                This.MainFrame["InactiveClientBorderColor"].Enabled := This.ShowAllColoredBorders
                This.NeedRestart := 1
            }
            else if (obj.Name = "InactiveClientBorderColor") {
                This.InactiveClientBorderColor := obj.value
                This.NeedRestart := 1
            }
            else if (obj.Name = "InactiveClientBorderthickness") {
                This.InactiveClientBorderthickness := obj.value
                This.NeedRestart := 1
            }
            else if (obj.name = "ThumbnailBackgroundColor") {
                This.ThumbnailBackgroundColor := obj.value
                This.NeedRestart := 1
            }
            else if (obj.name = "ThumbnailMinimumSizewidth") {
                This.ThumbnailMinimumSize["width"] := obj.value
            }
            else if (obj.name = "ThumbnailMinimumSizeheight") {
                This.ThumbnailMinimumSize["height"] := obj.value
            }
            SetTimer(This.Save_Settings_Delay_Timer, -200)
        }
    }

    ThumbnailVisibility_Ctrl() {
        This.MainFrame.Group["Thumbnail Visibility"] := [], Thumbnail_visibility := []

        This.MainFrame.SetFont("s12 w700 q5")
        Thumbnail_visibility.Push This.MainFrame.Add("Text", Format("x{} y{}", This.contentGap, This.contentGap), "Thumbnail Visibility")
        This.MainFrame.SetFont("s11 w400")
        Thumbnail_visibility.Push This.MainFrame.Add("Text", Format("xp yp+{} w{} h2 +0x10", This.lGap, This.sepW))

        Thumbnail_visibility.Push This.MainFrame.Add("Text", Format("xp yp+{} Section", This.lGap), "Select Any Client to Hide The Thumbnail:")
        This.Tv_LV := This.MainFrame.Add("ListView", Format("xp yp+{} w{}", This.contentGap, This.editW) " r20 Checked -LV0x10 -Multi -Sort vVisibility_List", ["Client Name"])
        Thumbnail_visibility.Push This.Tv_LV

        for k, v in This.compare_openclients_with_list() {
            if (k != "EVE" || v != "") {
                if This.Thumbnail_visibility.Has(v)
                    This.Tv_LV.Add("Check", v,)
                else
                    This.Tv_LV.Add("", v,)
            }
        }

        This.Tv_LV.ModifyCol(1, 150), This.Tv_LV.ModifyCol(2, 115)
        This.Tv_LV.OnEvent("ItemCheck", ObjBindMethod(This, "_Tv_LVSelectedRow"))

        This.MainFrame.Group["Thumbnail Visibility"] := Thumbnail_visibility
        for k, v in This.MainFrame.Group["Thumbnail Visibility"]
            v.Visible := 0
    }

    Other_Ctrl() {
        This.MainFrame.Group["Other"] := [], Other := []

        This.MainFrame.SetFont("s12 w700 q5")
        Other.Push This.MainFrame.Add("Text", Format("x{} y{}", This.contentGap, This.contentGap), "Other")
        This.MainFrame.SetFont("s11 w400")
        Other.Push This.MainFrame.Add("Text", Format("xp yp+{} w{} h2 +0x10", This.lGap, This.sepW))

        Other.Push This.MainFrame.Add("Text", Format("xp yp+{} Section", This.lGap), "Switch Language to English on Error:")
        Other.Push This.MainFrame.Add("CheckBox", Format("xp+{} yp", This.offsetX) " vSwitchLangOnErr Checked" This.SwitchLangOnErr, "On/Off")

        Other.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Check for Updates on Startup:")
        Other.Push This.MainFrame.Add("CheckBox", Format("xp+{} yp", This.offsetX) " vCheck_Updates Checked" This.Check_Updates, "On/Off")

        Other.Push This.MainFrame.Add("Text", Format("xs ys+{} w{} h2 +0x10", This.xlGap, This.sepW))
        Other.Push This.MainFrame.Add("Text", Format("xp yp+{} Section", This.contentGap), "Selected Groups Override Defaults, Skipping Per-Profile Setup.")
        
        Other.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Hotkeys:")
        Other.Push This.MainFrame.Add("CheckBox", Format("xp+{} yp", This.offsetX) " vGlobalHotkeys Checked" This.Global_Groups["Hotkeys"], "On/Off")
        
        Other.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Thumbnails Behavior:")
        Other.Push This.MainFrame.Add("CheckBox", Format("xp+{} yp", This.offsetX) " vGlobalThumbnailsBehavior Checked" This.Global_Groups["Thumbnails Behavior"], "On/Off")
        
        Other.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Thumbnails Visuals:")
        Other.Push This.MainFrame.Add("CheckBox", Format("xp+{} yp", This.offsetX) " vGlobalThumbnailsVisuals Checked" This.Global_Groups["Thumbnails Visuals"], "On/Off")
        
        Other.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Thumbnail Visibility:")
        Other.Push This.MainFrame.Add("CheckBox", Format("xp+{} yp", This.offsetX) " vGlobalThumbnailVisibility Checked" This.Global_Groups["Thumbnail Visibility"], "On/Off")
        
        Other.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Client Settings:")
        Other.Push This.MainFrame.Add("CheckBox", Format("xp+{} yp", This.offsetX) " vGlobalClientSettings Checked" This.Global_Groups["Client Settings"], "On/Off")
        
        Other.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Custom Colors:")
        Other.Push This.MainFrame.Add("CheckBox", Format("xp+{} yp", This.offsetX) " vGlobalCustomColors Checked" This.Global_Groups["Custom Colors"], "On/Off")
        
        Other.Push This.MainFrame.Add("Text", Format("xs ys+{} Section", This.xlGap), "Other:")
        Other.Push This.MainFrame.Add("CheckBox", Format("xp+{} yp", This.offsetX) " vGlobalOther Checked" This.Global_Groups["Other"], "On/Off")

        Other.Push This.MainFrame.Add("Button", Format("xs ys+{} Section", This.xlGap) " vUpdateGlobals", "Update")

        Other.Push This.MainFrame.Add("Text", Format("xs ys+{} w{} h2 +0x10", This.xlGap + This.objH, This.sepW))
        Other.Push This.MainFrame.Add("Text", Format("xp yp+{} Section", This.contentGap), "Update All Thumbnails Size to Default Size.")
        Other.Push This.MainFrame.Add("Button", Format("xs ys+{} Section", This.xlGap) " vUpdateThumbnails", "Update All Thumbnails")

        This.MainFrame["SwitchLangOnErr"].OnEvent("Click", (obj, *) => cOther_EventHandler(obj))
        This.MainFrame["Check_Updates"].OnEvent("Click", (obj, *) => cOther_EventHandler(obj))
        This.MainFrame["UpdateGlobals"].OnEvent("Click", (obj, *) => cOther_EventHandler(obj))
        This.MainFrame["UpdateThumbnails"].OnEvent("Click", (obj, *) => cOther_EventHandler(obj))

        cOther_EventHandler(obj) {
            need_reload := 0

            if (obj.name = "SwitchLangOnErr") {
                This.SwitchLangOnErr := obj.value
            }
            else if (obj.name = "Check_Updates") {
                This.Check_Updates := obj.value
            }
            else if (obj.name = "UpdateGlobals") {
                if This.Global_Groups["Hotkeys"] != This.MainFrame["GlobalHotkeys"].value {
                    This.Global_Groups["Hotkeys"] := This.MainFrame["GlobalHotkeys"].value
                    need_reload := 1
                } 
                if This.Global_Groups["Thumbnails Behavior"] != This.MainFrame["GlobalThumbnailsBehavior"].value {
                    This.Global_Groups["Thumbnails Behavior"] := This.MainFrame["GlobalThumbnailsBehavior"].value
                    need_reload := 1
                }
                if This.Global_Groups["Thumbnails Visuals"] != This.MainFrame["GlobalThumbnailsVisuals"].value {
                    This.Global_Groups["Thumbnails Visuals"] := This.MainFrame["GlobalThumbnailsVisuals"].value
                    need_reload := 1
                }
                if This.Global_Groups["Thumbnail Visibility"] != This.MainFrame["GlobalThumbnailVisibility"].value {
                    This.Global_Groups["Thumbnail Visibility"] := This.MainFrame["GlobalThumbnailVisibility"].value
                    need_reload := 1
                }
                if This.Global_Groups["Client Settings"] != This.MainFrame["GlobalClientSettings"].value {
                    This.Global_Groups["Client Settings"] := This.MainFrame["GlobalClientSettings"].value
                    need_reload := 1
                }
                if This.Global_Groups["Custom Colors"] != This.MainFrame["GlobalCustomColors"].value {
                    This.Global_Groups["Custom Colors"] := This.MainFrame["GlobalCustomColors"].value
                    need_reload := 1
                }
                if This.Global_Groups["Other"] != This.MainFrame["GlobalOther"].value {
                    This.Global_Groups["Other"] := This.MainFrame["GlobalOther"].value
                    need_reload := 1
                }

            }
            else if (obj.name = "UpdateThumbnails") {
                This.Update_All_Thumbnails()
                need_reload := 1
            }
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
            list .= k "`n"
        }
        return list
    }

    DontCloseList() {
        list := ""
        for k in This.DontCloseClients {
            list .= k "`n"
        }
        return list
    }

    _Button_Load(obj?,*) {
        if (IsSet(obj))
            This.NeedRestart := 1
        
        This.LastUsedProfile := This.Sidebar["SelectedProfile"].Text
        This.Refresh_ControlValues()
        This.ProfileOverride()

        if (This.Sidebar["SelectedProfile"].Text = "Default") {
            for k, ob in This.MainFrame.Group["Other"] {
                if InStr(ob.name, "Global")
                    ob.Enabled := 1
            }
        }
        else {
            for k, v in This.MainFrame.Group {
                for group, enab in This.Global_Groups {
                    if k != group || !enab
                        continue
                    for _, ob in v {
                        ob.Enabled := 0
                    }
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
        This.MainFrame["HotkeyGroupDDL"].Delete()
        This.MainFrame["HotkeyGroupDDL"].Add(This.GetGroupList())
        This.MainFrame["ForwardsKey"].value := "", This.MainFrame["ForwardsKey"].Enabled := 0
        This.MainFrame["BackwardsdKey"].value := "", This.MainFrame["BackwardsdKey"].Enabled := 0
        This.MainFrame["HKCharlist"].value := "", This.MainFrame["HKCharlist"].Enabled := 0

        ;Hotkeys
        This.MainFrame["Suspend_Hotkeys_Hotkey"].value := This.Suspend_Hotkeys_Hotkey
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
                Charlist .= name "`n"
                Hklist .= hotkey "`n"
            }
        }

        This.MainFrame["HotkeyCharList"].value := Charlist
        This.MainFrame["HotkeyList"].value := Hklist

        ;Thumbnail Settings
        This.MainFrame["ShowThumbnailTextOverlay"].value := This.ShowThumbnailTextOverlay
        This.MainFrame["ThumbnailTextColor"].value := This.ThumbnailTextColor
        This.MainFrame["ThumbnailTextSize"].value := This.ThumbnailTextSize
        This.MainFrame["ThumbnailTextFont"].value := This.ThumbnailTextFont
        This.MainFrame["ThumbnailTextMarginsx"].value := This.ThumbnailTextMargins["x"]
        This.MainFrame["ThumbnailTextMarginsy"].value := This.ThumbnailTextMargins["y"]
        This.MainFrame["ClientHighligtColor"].value := This.ClientHighligtColor
        This.MainFrame["ClientHighligtBorderthickness"].value := This.ClientHighligtBorderthickness
        This.MainFrame["ShowClientHighlightBorder"].value := This.ShowClientHighlightBorder
        This.MainFrame["HideThumbnailsOnLostFocus"].value := This.HideThumbnailsOnLostFocus
        This.MainFrame["ThumbnailOpacity"].value := IntegerToPercentage(This.ThumbnailOpacity)
        This.MainFrame["ShowThumbnailsAlwaysOnTop"].value := This.ShowThumbnailsAlwaysOnTop
        This.MainFrame["ShowAllBorders"].value := This.ShowAllColoredBorders
        This.MainFrame["InactiveClientBorderthickness"].value := This.InactiveClientBorderthickness
        This.MainFrame["InactiveClientBorderColor"].value := This.InactiveClientBorderColor
        This.MainFrame["ThumbnailBackgroundColor"].value := This.ThumbnailBackgroundColor
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

        ; Other
        This.MainFrame["SwitchLangOnErr"].value := This.SwitchLangOnErr
        This.MainFrame["Check_Updates"].value := This.Check_Updates

        This.MainFrame["GlobalHotkeys"].value := This.Global_Groups["Hotkeys"]
        This.MainFrame["GlobalThumbnailsBehavior"].value := This.Global_Groups["Thumbnails Behavior"]
        This.MainFrame["GlobalThumbnailsVisuals"].value := This.Global_Groups["Thumbnails Visuals"]
        This.MainFrame["GlobalThumbnailVisibility"].value := This.Global_Groups["Thumbnail Visibility"]
        This.MainFrame["GlobalClientSettings"].value := This.Global_Groups["Client Settings"]
        This.MainFrame["GlobalCustomColors"].value := This.Global_Groups["Custom Colors"]
        This.MainFrame["GlobalOther"].value := This.Global_Groups["Other"]

        for k, v in This.MainFrame.Group {
            for _, ob in v {
                ob.Enabled := 1
            }
            This.MainFrame["HKCharlist"].Enabled := 0
            This.MainFrame["ForwardsKey"].Enabled := 0
            This.MainFrame["BackwardsdKey"].Enabled := 0
            This.MainFrame["ImpNamesBtn"].Enabled := 0
        }
        This.MainFrame["InactiveClientBorderthickness"].Enabled := This.ShowAllColoredBorders
        This.MainFrame["InactiveClientBorderColor"].Enabled := This.ShowAllColoredBorders

        This.MainFrame["GlobalHotkeys"].Enabled := 0
        This.MainFrame["GlobalThumbnailsBehavior"].Enabled := 0
        This.MainFrame["GlobalThumbnailsVisuals"].Enabled := 0
        This.MainFrame["GlobalThumbnailVisibility"].Enabled := 0
        This.MainFrame["GlobalClientSettings"].Enabled := 0
        This.MainFrame["GlobalCustomColors"].Enabled := 0
        This.MainFrame["GlobalOther"].Enabled := 0
    }


    compare_openclients_with_list() {
        EvENameList := []
        for EveHwnd in This.ThumbWindows.OwnProps() {
            try {
                if title := This.CleanTitle(WinGetTitle("Ahk_Id " EveHwnd) = "") {
                    continue
                }
                EvENameList.Push This.CleanTitle(WinGetTitle("Ahk_Id " EveHwnd))
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
