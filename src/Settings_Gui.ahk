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
        This.ThumbnailSettings_Ctrl()
        This.ThumbnailVisibility_Ctrl()
        This.ExcludeFromClosing_Ctrl()
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

        ; This.SB := ScrollBar(This.MainFrame, This.guiWidth - This.sidebarW, This.guiHeight)
        ; This.SB.Vertical := true
        ; This.SB.Horizontal := false

        ; OnMessage(0x020A, ObjBindMethod(this, "OnMouseWheel"))

    }

    ; OnMouseWheel(wParam, lParam, msg, hwnd) {
    ;     wheelDelta := ((wParam >> 16) << 24) >> 24
    ;     direction := (wheelDelta > 0) ? 0 : 1
        
    ;     shift := GetKeyState("Shift", "P")
    ;     ; scrollType := shift ? 0x114 : 0x115
    ;     scrollType := 0x115
        
    ;     this.SB.ScrollMsg(direction, 0, scrollType, this.MainFrame.Hwnd)
        
    ;     return 0
    ; }

    SetState() {
        This.baseGrid := 8
        This.contentGap := 16

        This.guiWidth := 800
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
        This.offsetX := 240

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

        This.ddl_options := [
            "left -> right, top -> bottom", 
            "left -> right, bottom -> top", 
            "right -> left, top -> bottom", 
            "right -> left, bottom -> top", 
            "top -> bottom, left -> right", 
            "top -> bottom, right -> left", 
            "bottom -> top, left -> right", 
            "bottom -> top, right -> left"
        ]
    }

    CreateSidebar() {
        This.Sidebar := Gui("+Parent" This.S_Gui.Hwnd " -Caption -Border")
        This.Sidebar.BackColor := "0xf0f0f0"
        This.Sidebar.SetFont("s11 w400")

        This.Sidebar.Add("Text", Format("x{} y{}", This.contentGap, This.contentGap), "Select profile")
        
        This.SelectProfile_DDL := This.Sidebar.Add("DDL", Format("xp yp+{} w{}", This.contentGap, This.sidebarInnerW) " Section vSelectedProfile", This.Profiles_to_Array())

        btnAdd := This.Sidebar.Add("Button", Format("xp-1 ys+{} w{} h{}", This.objH + This.baseGrid, This.sidebar3BtnW, This.btnH), "Add")
        btnRename := This.Sidebar.Add("Button", Format("xp+{} ys+{} w{} h{}", This.sidebar3BtnW + This.baseGrid + 1, This.objH + This.baseGrid, This.sidebar3BtnW, This.btnH), "Rename")
        btnDelete := This.Sidebar.Add("Button", Format("xp+{} ys+{} w{} h{}", This.sidebar3BtnW + This.baseGrid + 1, This.objH + This.baseGrid, This.sidebar3BtnW, This.btnH), "Delete")

        searchEdit := This.Sidebar.Add("Edit", Format("xs yp+{} w{}", This.btnH + This.contentGap, This.sidebarInnerW))

        This.Sidebar.Add("Button", Format("xp yp+{} w{} h{}", This.objH + This.baseGrid, This.sidebarInnerW, This.btnH), "Find next")

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
        This.Sidebar.Add("Button", Format("xp yp-{} w{} h{}", This.btnH + This.baseGrid, This.sidebarInnerW, This.btnH) " vHelp_Button", "Help").OnEvent("Click", (*) => This.Help_Button_Handler())

        This.SelectProfile_DDL.Choose(This.LastUsedProfile)
        This.SelectProfile_DDL.OnEvent("Change", (obj,*) => This._Button_Load(Obj))
        btnAdd.OnEvent("Click", ObjBindMethod(This, "Create_Profile"))
        ; btnRename.OnEvent("Click", ObjBindMethod(This, "Rename_Profile")) TODO
        btnDelete.OnEvent("Click", ObjBindMethod(This, "Delete_Profile"))
        DllCall("User32.dll\SendMessageW", "Ptr", searchEdit.Hwnd, "UInt", 0x1500 + 1, "UPtr", true, "WStr", "Search setting...", "Ptr")
        ; Search button handler TODO
    }

    CreateMainFrame() {
        This.MainFrame := Gui("+Parent" This.S_Gui.Hwnd " -Caption -Border +Resize")
        This.MainFrame.BackColor := "0xf7f7f7"
        This.MainFrame.SetFont("s11 w400")

        ; Create map group which holds the GUI objects for the settings groups
        This.MainFrame.Group := Map()

        ; Separator
        ; This.S_Gui.Add("Text", Format("x{} yp+{} w{} h2 +0x10", This.sidebarX + This.baseGrid + 1, This.btnH + This.baseGrid, This.sidebarW - This.contentGap))
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

        ; ;Sets Margins for the following Buttons
        ; This.S_Gui.MarginX := 80, This.S_Gui.MarginY := 20

        ; ;Font options for the Profiles Buttons
        ; This.S_Gui.SetFont("s12 Bold")

        ; ; Primary buttons
        ; This.S_Gui.Add("Button", "x19 y12 w196 h64 vGlobal_Settings", "Global Settings").OnEvent("Click", (obj, *) => Button_Handler(obj))
        ; This.S_Gui.Add("Button", "xp+206 yp wp hp vProfile_Settings", "Profile Settings").OnEvent("Click", (obj, *) => Button_Handler(obj))

        ;Font options for the Utility Buttons
        ; This.S_Gui.SetFont("s10")

        ; Utility buttons
        ; This.S_Gui.Add("Button", "xp+206 y12 w90 h28 vAbout_Button", "About").OnEvent("Click", (*) => This.About_Button_Handler())
        ; This.S_Gui.Add("Button", "xp yp+36 wp hp vHelp_Button", "Help").OnEvent("Click", (*) => This.Help_Button_Handler())

        ; This.S_Gui.Show("hide")

        ;Create the Arrays which hold the GUI objects for the controls
        ; This.S_Gui.Controls := [], This.S_Gui.ClientSettings := []

        ;Sets Margins for the following controls
        ; This.S_Gui.MarginX := 25

        ;Default Font options for the controls
        ; This.S_Gui.SetFont("s9 w400")

        ;Creates the Controls
        ; This.Global_Settings()
        ; This.Profile_Settings()
        ; This.ClientSettings_Ctrl()
        ; This.Custom_ColorsCtrl()
        ; This.Hotkey_GroupsCtrl()
        ; This.HotkeysCtrl()
        ; This.ThumbnailSettings_Ctrl()
        ; This.Thumbnail_visibilityCtrl()
        ; This.ExcludeFromClosing_Ctrl()

    ;     This.S_Gui.Show("AutoSize Center")
    ;     This._Button_Load()

    ;     This.Seetings_DDL.OnEvent("Change", (Obj, *) => SettingsDDL_Handler(Obj))

    ;     This.S_Gui.OnEvent("Close", (*) => GuiDestroy())

    ;     GuiDestroy(*) {
    ;         This.S_Gui.Destroy()
    ;         if (This.NeedRestart)
    ;             Reload()
    ;     }

    ;     SettingsDDL_Handler(Obj) {
    ;         for k, v in This.S_Gui.Controls.Profile_Settings.PsDDL {
    ;             if k = Obj.Text {
    ;                 for _, ob in v
    ;                     ob.Visible := 1
    ;             }
    ;             else {
    ;                 for _, ob in v
    ;                     ob.Visible := 0
    ;             }                
    ;         }
    ;         This.S_Gui.Show("AutoSize")
    ;     }

    ;     Button_Handler(obj) {
    ;         if (obj.Name = "Global_Settings") {
    ;             for ButtonName, Controls in This.S_Gui.Controls.OwnProps() {
    ;                 if ButtonName = obj.Name {
    ;                     for _, Ctrl in Controls {
    ;                         Ctrl.Visible := 1
    ;                     }
    ;                 }
    ;                 else {
    ;                     for _, Ctrl in Controls {
    ;                         Ctrl.Visible := 0
    ;                     }
    ;                     for _, Ctrl in This.S_Gui.Controls.Profile_Settings.PsDDL {
    ;                         for k, v in Ctrl
    ;                             v.Visible := 0
    ;                     }

    ;                 }
    ;             }
    ;         }
    ;         else if (obj.Name = "Profile_Settings") {
    ;             for ButtonName, Controls in This.S_Gui.Controls.OwnProps() {
    ;                 if ButtonName = obj.Name {
    ;                     for _, Ctrl in Controls {
    ;                         Ctrl.Visible := 1
    ;                     }
    ;                     for _, Ctrl in This.S_Gui.Controls.Profile_Settings.PsDDL {
    ;                         if (This.Seetings_DDL.Text = _) {
    ;                             for k, v in Ctrl {
    ;                                 v.Visible := 1
    ;                             }
    ;                         }
    ;                     }                        
    ;                 }
    ;                 else {
    ;                     for _, Ctrl in Controls {
    ;                         Ctrl.Visible := 0
    ;                     }
    ;                 }
    ;             }
    ;             if (This.Profiles.Count = 1 && This.SelectProfile_DDL.Text = "Default")
    ;                 MsgBox("you need create a profile first to change the settings")
    ;         }

    ;         This.S_Gui.Show("AutoSize")
    ;     }
    ; }

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


    ;This Function creates all Settings controls for the Profile Settings Button
    ; Profile_Settings(visible?) {
    ;     This.S_Gui.Controls.Profile_Settings := [], This.S_Gui.Controls.Profile_Settings.PsDDL := Map()

    ;     ;This.S_Gui.Controls.Profile_Settings.Push This.S_Gui.Add("GroupBox", "x20 y80 h200 w500 vPSGroupBox", "")
    ;     This.S_Gui.Controls.Profile_Settings.Push This.S_Gui.Add("Text", "x58 y95", "Select Profile:")

    ;     This.SelectProfile_DDL := This.S_Gui.Add("DDL", "w200 xp-30 yp+18 Section vSelectedProfile", This.Profiles_to_Array())
    ;     This.S_Gui.Controls.Profile_Settings.Push This.SelectProfile_DDL
    ;     This.SelectProfile_DDL.Choose(This.LastUsedProfile)
    ;     This.SelectProfile_DDL.OnEvent("Change", (obj,*) => This._Button_Load(Obj))

    ;     Button_Delete := This.S_Gui.Add("Button", "w60 xs+360 yp-2 ", "Delete")
    ;     This.S_Gui.Controls.Profile_Settings.Push Button_Delete
    ;     Button_Delete.OnEvent("Click", ObjBindMethod(This, "Delete_Profile"))

    ;     Button_New := This.S_Gui.Add("Button", "wp x+5 yp ", "New")
    ;     This.S_Gui.Controls.Profile_Settings.Push Button_New
    ;     Button_New.OnEvent("Click", ObjBindMethod(This, "Create_Profile"))

    ;     ;*Seperator line
    ;     This.Seperator_text := This.S_Gui.Add("Text", "xs+15 y+5 w460 h2 +0x10")
    ;     This.S_Gui.Controls.Profile_Settings.Push This.Seperator_text

    ;     This.S_Gui.Controls.Profile_Settings.Push This.S_Gui.Add("Text", "xp+190 y+5", "Profile Settings:")

    ;     This.Seetings_DDL := This.S_Gui.Add("DDL", "w180 xp-40 y+5 vSeetings_Props", This._ProfileProps)
    ;     This.Seetings_DDL.Choose(1)
    ;     ;This.Seetings_DDL.OnEvent("Change", ObjBindMethod(This, "ProfileSettings_DDL"))
    ;     This.S_Gui.Controls.Profile_Settings.Push This.Seetings_DDL

    ;     ;*Seperator line
    ;     This.S_Gui.Controls.Profile_Settings.Push This.S_Gui.Add("Text", "x150 yp+30 w260 h2 Section +0x10")

    ;     ;Sets all controls invisible at beginning
    ;     for k, v in This.S_Gui.Controls.Profile_Settings
    ;         v.Visible := 0
    ; }

    ClientSettings_Ctrl(visible?) {
        This.MainFrame.Group["Client Settings"] := [], ClientSettings := []
        
        ClientSettings.Push This.MainFrame.Add("Text", Format("x{} y{}", This.contentGap, This.contentGap), "Client Settings")

        ClientSettings.Push This.MainFrame.Add("Text", Format("xp yp+{} Section", This.xlGap), "Minimize Inactive Clients:")
        ClientSettings.Push This.MainFrame.Add("CheckBox", Format("xs+{} yp", This.offsetX) " vMinimizeInactiveClients Checked" This.MinimizeInactiveClients, "On/Off")

        ClientSettings.Push This.MainFrame.Add("Text", Format("xs yp+{} Section", This.xlGap), "Always Maximize Clients:")
        ClientSettings.Push This.MainFrame.Add("CheckBox", Format("xs+{} yp", This.offsetX) " vAlwaysMaximize Checked" This.AlwaysMaximize, "On/Off")

        ClientSettings.Push This.MainFrame.Add("Text", Format("xs yp+{} Section", This.xlGap), "EVE Window Minimize Delay (ms):")
        ClientSettings.Push This.MainFrame.Add("Edit", Format("xs+{} yp-{} w{}", This.offsetX, This.editOffset, This.editW) " vMinimizeclients_Delay", This.Minimizeclients_Delay)

        ClientSettings.Push This.MainFrame.Add("Text", Format("xs yp+{} Section", This.xlGap), "Dont Minimize Clients:")
        ClientSettings.Push This.MainFrame.Add("Edit", Format("xs+{} yp-{} w{} h{}", This.offsetX, This.editOffset, This.editW, 200) " vDont_Minimize_Clients -Wrap", This.Dont_Minimize_List())

        This.MainFrame["MinimizeInactiveClients"].OnEvent("Click", (obj, *) => cSettings_EventHandler(obj))
        This.MainFrame["AlwaysMaximize"].OnEvent("Click", (obj, *) => cSettings_EventHandler(obj))
        This.MainFrame["Minimizeclients_Delay"].OnEvent("Change", (obj, *) => cSettings_EventHandler(obj))
        This.MainFrame["Dont_Minimize_Clients"].OnEvent("Change", (obj, *) => cSettings_EventHandler(obj))

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
            SetTimer(This.Save_Settings_Delay_Timer, -200)
        }
    }

    ; User defined colors per Client
    CustomColors_Ctrl() {
        This.MainFrame.Group["Custom Colors"] := [], CustomColors := []

        CustomColors.Push This.MainFrame.Add("Text", Format("x{} y{}", This.contentGap, This.contentGap), "Custom Colors")

        CustomColors.Push This.MainFrame.Add("Text", Format("xp yp+{} Section", This.xlGap), "Custom Colors Active:")
        CustomColors.Push This.MainFrame.Add("CheckBox", Format("xs+{} yp", This.offsetX) " vCcoloractive Checked" This.CustomColorsActive, " On/Off")

        CustomColors.Push This.MainFrame.Add("Text", Format("xs yp+{} Section", This.xlGap), "Character Name:")
        CustomColors.Push This.MainFrame.Add("Edit", Format("xp ys+{} w{} h{}", This.contentGap, This.editW, This.editH) " -Wrap vCchars", This.CustomColors_AllCharNames)

        CustomColors.Push This.MainFrame.Add("Text", Format("xs+{} ys Section", This.editW + This.xlGap), "Active Border Color:")
        CustomColors.Push This.MainFrame.Add("Edit", Format("xp ys+{} w{} h{}", This.contentGap, This.editW, This.editH) " -Wrap vCBorderColor", This.CustomColors_AllBColors)

        CustomColors.Push This.MainFrame.Add("Text", Format("x{} yp+{} Section", This.contentGap, This.editH + This.contentGap), "Text Color:")
        CustomColors.Push This.MainFrame.Add("Edit", Format("xp ys+{} w{} h{}", This.contentGap, This.editW, This.editH) " -Wrap vCTextColor", This.CustomColors_AllTColors)

        CustomColors.Push This.MainFrame.Add("Text", Format("xs+{} ys Section", This.editW + This.xlGap), "Inactive Border Color:")
        CustomColors.Push This.MainFrame.Add("Edit", Format("xp ys+{} w{} h{}", This.contentGap, This.editW, This.editH) " -Wrap vIABorderColor", This.CustomColors_IABorder_Colors)

        This.MainFrame["Ccoloractive"].OnEvent("Click", (obj, *) => Cclors_Eventhandler(obj))
        This.MainFrame["Cchars"].OnEvent("Change", (obj, *) => Cclors_Eventhandler(obj))
        This.MainFrame["CBorderColor"].OnEvent("Change", (obj, *) => Cclors_Eventhandler(obj))
        This.MainFrame["CTextColor"].OnEvent("Change", (obj, *) => Cclors_Eventhandler(obj))
        This.MainFrame["IABorderColor"].OnEvent("Change", (obj, *) => Cclors_Eventhandler(obj))

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

        Hotkey_Groups.Push This.MainFrame.Add("Text", "x16 y16", "Select Group:")
        ddl := This.MainFrame.Add("DropDownList", " xp-30 yp+18 w180 vHotkeyGroupDDL", This.GetGroupList())
        Hotkey_Groups.Push ddl
        This.MainFrame["HotkeyGroupDDL"].OnEvent("Change", (*) => SetEditText(ddl, EditBox, HKForwards, HKBackwards))

        DeleteButton := This.MainFrame.Add("Button", "xs+370 yp-1 w60", "Delete")
        NewButton := This.MainFrame.Add("Button", "x+5 yp w60", "New")
        DeleteButton.OnEvent("Click", (*) => Delete_Group(ddl, HKForwards, HKBackwards, EditBox))
        NewButton.OnEvent("Click", (*) => CreateNewGroup(ddl, HKForwards, HKBackwards, EditBox))

        Hotkey_Groups.Push DeleteButton
        Hotkey_Groups.Push NewButton

        EditBox := This.MainFrame.Add("Edit", "xs+8 y275 w250 h225 -Wrap +HScroll Disabled vHKCharlist")
        Hotkey_Groups.Push EditBox
        This.MainFrame["HKCharlist"].OnEvent("Change", (obj, *) => SaveHKGroupList(obj))

        Hotkey_Groups.Push This.MainFrame.Add("Text", "xs300 yp20", "Forwards Hotkey:")
        HKForwards := This.MainFrame.Add("Edit", "xp yp+20 w150 Disabled vForwardsKey")
        Hotkey_Groups.Push HKForwards
        This.MainFrame["ForwardsKey"].OnEvent("Change", (obj, *) => SaveHKGroupList(obj))

        Hotkey_Groups.Push This.MainFrame.Add("Text", "xp yp50", "Backwards Hotkey:")
        HKBackwards := This.MainFrame.Add("Edit", "xp yp+20 w150 Disabled vBackwardsdKey")
        Hotkey_Groups.Push HKBackwards
        This.MainFrame["BackwardsdKey"].OnEvent("Change", (obj, *) => SaveHKGroupList(obj))

        This.MainFrame.Group["Hotkey Groups"] := Hotkey_Groups
        for k, v in This.MainFrame.Group["Hotkey Groups"]
            v.Visible := 0


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

        SetEditText(ddlObj, EditObj, ForwardHKObj?, BackwardHKObj?) {
            text := ""
            if (ddlObj.Text != "" && This.Hotkey_Groups.Has(ddlObj.Text)) {
                for index, Names in This.Hotkey_Groups[ddlObj.Text]["Characters"] {
                    text .= Names "`n"
                }
                EditObj.value := text, EditObj.Enabled := 1
                ForwardHKObj.value := This.Hotkey_Groups[ddlObj.Text]["ForwardsHotkey"], ForwardHKObj.Enabled := 1
                BackwardHKObj.value := This.Hotkey_Groups[ddlObj.Text]["BackwardsHotkey"], BackwardHKObj.Enabled := 1
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

        RowSpacing := 30
        SeparatorSpacing := 10
        LabelOffset := 15
        FieldOffset := 289
        DefaultWidth := 180
        DigitWidth := 40

        Hotkeys.Push This.MainFrame.Add("Text", "xp+" . LabelOffset . " yp+" . RowSpacing . " Section", "Suspend All Hotkeys – Hotkey")

        Hotkeys.Push This.MainFrame.Add("Edit", "xp+" . FieldOffset . " ys w" . DefaultWidth . " vSuspend_Hotkeys_Hotkey", This.Suspend_Hotkeys_Hotkey)
        This.MainFrame["Suspend_Hotkeys_Hotkey"].OnEvent("Change", (obj, *) => cHotkeys_EventHandler(obj))

        Hotkeys.Push This.MainFrame.Add("Text", "xs ys+" . RowSpacing . " Section", "Hotkey Activation Scope")

        Hotkeys.Push This.MainFrame.Add("DDL", "xp+" . FieldOffset . " ys w" . DefaultWidth . " vTTT vHotkey_Scoope Choose" (This.Global_Hotkeys ? 1 : 2), ["Global", "If an EVE window is Active"])
        This.MainFrame["Hotkey_Scoope"].OnEvent("Change", (obj, *) => cHotkeys_EventHandler(obj))

        Hotkeys.Push This.MainFrame.Add("Text", "xs ys+" . RowSpacing . " Section", "Cycle Login Screens – Hotkey")

        Hotkeys.Push This.MainFrame.Add("Edit", "xp+" . FieldOffset . " ys w" . DefaultWidth . " vLogin_Screen_Cycle_Hotkey", This.Login_Screen_Cycle_Hotkey)
        This.MainFrame["Login_Screen_Cycle_Hotkey"].OnEvent("Change", (obj, *) => cHotkeys_EventHandler(obj))

        Hotkeys.Push This.MainFrame.Add("Text", "xs ys+" . RowSpacing . " Section", "Login Screen Cycle Direction")

        Hotkeys.Push This.MainFrame.Add("Radio", "xp+" . FieldOffset . " ys w37 vLoginScreenCycleDirectionForwards Checked" This.LoginScreenCycleDirection, "Old->New")
        Hotkeys.Push This.MainFrame.Add("Radio", " xp+90 yp w37 vLoginScreenCycleDirectionBackwards Checked" (This.LoginScreenCycleDirection ? 0 : 1), "New->Old")
        This.MainFrame["LoginScreenCycleDirectionForwards"].OnEvent("Click", (obj, *) => cHotkeys_EventHandler(obj))
        This.MainFrame["LoginScreenCycleDirectionBackwards"].OnEvent("Click", (obj, *) => cHotkeys_EventHandler(obj))

        Hotkeys.Push This.MainFrame.Add("Text", "xs ys+" . RowSpacing . " Section", "Close Active EVE Window – Hotkey")

        Hotkeys.Push This.MainFrame.Add("Edit", "xp+" . FieldOffset . " ys w" . DefaultWidth . " vClose_Active_EVE_Win_Hotkey", This.Close_Active_EVE_Win_Hotkey)
        This.MainFrame["Close_Active_EVE_Win_Hotkey"].OnEvent("Change", (obj, *) => cHotkeys_EventHandler(obj))

        Hotkeys.Push This.MainFrame.Add("Text", "xs ys+" . RowSpacing . " Section", "Close All EVE Windows – Hotkey")

        Hotkeys.Push This.MainFrame.Add("Edit", "xp+" . FieldOffset . " ys w" . DefaultWidth . " vClose_All_EVE_Win_Hotkey", This.Close_All_EVE_Win_Hotkey)
        This.MainFrame["Close_All_EVE_Win_Hotkey"].OnEvent("Change", (obj, *) => cHotkeys_EventHandler(obj))

        Hotkeys.Push This.MainFrame.Add("Text", "xs ys+" . RowSpacing . " Section", "Reload EVE-X-Preview – Hotkey")

        Hotkeys.Push This.MainFrame.Add("Edit", "xp+" . FieldOffset . " ys w" . DefaultWidth . " vReload_Program_Hotkey", This.Reload_Program_Hotkey)
        This.MainFrame["Reload_Program_Hotkey"].OnEvent("Change", (obj, *) => cHotkeys_EventHandler(obj))

        Hotkeys.Push This.MainFrame.Add("Text", "xs ys+" . SeparatorSpacing . " Section", "Preserve Hotkeys on Logout")

        Hotkeys.Push This.MainFrame.Add("CheckBox", "xp+" . FieldOffset . " ys vPreserveHotkeysOnLogout Checked" This.PreserveHotkeysOnLogout, "On/Off")
        This.MainFrame["PreserveHotkeysOnLogout"].OnEvent("Click", (obj, *) => cHotkeys_EventHandler(obj))

        Hotkeys.Push This.MainFrame.Add("Text", " x16 y16 section", "Character Name:")
        HKCharList := This.MainFrame.Add("Edit", " xp-30 yp20 w180 h350 -Wrap vHotkeyCharList", Charlist)
        Hotkeys.Push HKCharList
        HKCharList.OnEvent("Change", (obj, *) => EventHandler(obj))

        Hotkeys.Push This.MainFrame.Add("Text", " xs+210 ys", "Hotkeys:")
        HKKeylist := This.MainFrame.Add("Edit", " xp-50 yp20 w180 h350 -Wrap vHotkeyList", Hklist)
        Hotkeys.Push HKKeylist
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
            else if (obj.name = "PreserveHotkeysOnLogout") {
                This.PreserveHotkeysOnLogout := obj.value
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

    ThumbnailSettings_Ctrl() {
        This.MainFrame.Group["Thumbnail Settings"] := [], ThumbnailSettings := []

        ThumbnailSettings.Push This.MainFrame.Add("Text", "x16 y16 Section", "Show Thumbnail Text Overlay:")
        ThumbnailSettings.Push This.MainFrame.Add("Text", " xs y+15 ", "Thumbnail Text Color:")
        ThumbnailSettings.Push This.MainFrame.Add("Text", " xs y+15 ", "Thumbnail Text Size:")
        ThumbnailSettings.Push This.MainFrame.Add("Text", " xs y+15 ", "Thumbnail Text Font:")
        ThumbnailSettings.Push This.MainFrame.Add("Text", " xs y+15 ", "Thumbnail Text Margins:")
        ThumbnailSettings.Push This.MainFrame.Add("Text", " xs y+15 ", "Client Highligt Color:")
        ThumbnailSettings.Push This.MainFrame.Add("Text", " xs y+15 ", "Client Highligt Border Thickness:")
        ThumbnailSettings.Push This.MainFrame.Add("Text", " xs y+15 ", "Show Client Highlight Border:")
        ThumbnailSettings.Push This.MainFrame.Add("Text", " xs y+15 ", "Hide Thumbnails On Lost Focus:")
        ThumbnailSettings.Push This.MainFrame.Add("Text", " xs y+15 ", "Thumbnail Opacity:")
        ThumbnailSettings.Push This.MainFrame.Add("Text", " xs y+15 ", "Show Thumbnails AlwaysOnTop:")
        ThumbnailSettings.Push This.MainFrame.Add("Text", " xs y+15", "Show All Borders:")
        ThumbnailSettings.Push This.MainFrame.Add("Text", " xs y+15", "Inactive Client Border Thickness:")
        ThumbnailSettings.Push This.MainFrame.Add("Text", " xs y+15", "Inactive Client Border Color:")

        ThumbnailSettings.Push This.MainFrame.Add("CheckBox", "xs+230 ys Section vShowThumbnailTextOverlay Checked" This.ShowThumbnailTextOverlay, "On/Off")
        This.MainFrame["ShowThumbnailTextOverlay"].OnEvent("Click", (obj, *) => ThumbnailSettings_EventHandler(obj))

        ThumbnailSettings.Push This.MainFrame.Add("Edit", "xs y+11  w120 vThumbnailTextColor -Wrap", This.ThumbnailTextColor)
        ThumbnailSettings.Push This.MainFrame.Add("Text", " x+5 yp+3 ", "Hex or RGB")
        This.MainFrame["ThumbnailTextColor"].OnEvent("Change", (obj, *) => ThumbnailSettings_EventHandler(obj))

        ThumbnailSettings.Push This.MainFrame.Add("Edit", "xs y+10 w30 vThumbnailTextSize -Wrap", This.ThumbnailTextSize)
        This.MainFrame["ThumbnailTextSize"].OnEvent("Change", (obj, *) => ThumbnailSettings_EventHandler(obj))

        ThumbnailSettings.Push This.MainFrame.Add("Edit", "xs y+8 w120 vThumbnailTextFont -Wrap", This.ThumbnailTextFont)
        This.MainFrame["ThumbnailTextFont"].OnEvent("Change", (obj, *) => ThumbnailSettings_EventHandler(obj))

        ThumbnailSettings.Push This.MainFrame.Add("Text", " xs y+12", "width px:")
        ThumbnailSettings.Push This.MainFrame.Add("Edit", "x+5 yp-4  w40 vThumbnailTextMarginsx -Wrap", This.ThumbnailTextMargins["x"])
        This.MainFrame["ThumbnailTextMarginsx"].OnEvent("Change", (obj, *) => ThumbnailSettings_EventHandler(obj))

        ThumbnailSettings.Push This.MainFrame.Add("Text", " xs+100 yp+4 ", "height px:")
        ThumbnailSettings.Push This.MainFrame.Add("Edit", "xp+60 yp-4  w40 vThumbnailTextMarginsy -Wrap", This.ThumbnailTextMargins["y"])
        This.MainFrame["ThumbnailTextMarginsy"].OnEvent("Change", (obj, *) => ThumbnailSettings_EventHandler(obj))

        ThumbnailSettings.Push This.MainFrame.Add("Edit", "xs y+7 w120 vClientHighligtColor -Wrap", This.ClientHighligtColor)
        ThumbnailSettings.Push This.MainFrame.Add("Text", " x+5 yp+3 ", "Hex or RGB")
        This.MainFrame["ClientHighligtColor"].OnEvent("Change", (obj, *) => ThumbnailSettings_EventHandler(obj))

        ThumbnailSettings.Push This.MainFrame.Add("Text", " xs y+15 ", "px:")
        ThumbnailSettings.Push This.MainFrame.Add("Edit", "x+5 yp-3  w30 vClientHighligtBorderthickness -Wrap", This.ClientHighligtBorderthickness)
        This.MainFrame["ClientHighligtBorderthickness"].OnEvent("Change", (obj, *) => ThumbnailSettings_EventHandler(obj))

        ThumbnailSettings.Push This.MainFrame.Add("CheckBox", "xs y+10 vShowClientHighlightBorder Checked" This.ShowClientHighlightBorder, "On/Off")
        This.MainFrame["ShowClientHighlightBorder"].OnEvent("Click", (obj, *) => ThumbnailSettings_EventHandler(obj))

        ThumbnailSettings.Push This.MainFrame.Add("CheckBox", "xs y+16 vHideThumbnailsOnLostFocus Checked" This.HideThumbnailsOnLostFocus, "On/Off")
        This.MainFrame["HideThumbnailsOnLostFocus"].OnEvent("Click", (obj, *) => ThumbnailSettings_EventHandler(obj))

        ThumbnailSettings.Push This.MainFrame.Add("Text", " xs y+15 ", "%")
        ThumbnailSettings.Push This.MainFrame.Add("Edit", "x+4 yp-4  w40 vThumbnailOpacity -Wrap", IntegerToPercentage(This.ThumbnailOpacity))
        This.MainFrame["ThumbnailOpacity"].OnEvent("Change", (obj, *) => ThumbnailSettings_EventHandler(obj))

        ThumbnailSettings.Push This.MainFrame.Add("CheckBox", "xs y+12 vShowThumbnailsAlwaysOnTop Checked" This.ShowThumbnailsAlwaysOnTop, "On/Off")
        This.MainFrame["ShowThumbnailsAlwaysOnTop"].OnEvent("Click", (obj, *) => ThumbnailSettings_EventHandler(obj))

        ThumbnailSettings.Push This.MainFrame.Add("CheckBox", "xs y+15 vShowAllBorders Checked" This.ShowAllColoredBorders, "On/Off")
        This.MainFrame["ShowAllBorders"].OnEvent("Click", (obj, *) => ThumbnailSettings_EventHandler(obj))

        ThumbnailSettings.Push This.MainFrame.Add("Text", " xs y+12 ", "px:")
        ThumbnailSettings.Push This.MainFrame.Add("Edit", "x+5 yp-3  w30 vInactiveClientBorderthickness -Wrap", This.InactiveClientBorderthickness)
        This.MainFrame["InactiveClientBorderthickness"].OnEvent("Change", (obj, *) => ThumbnailSettings_EventHandler(obj))

        ThumbnailSettings.Push This.MainFrame.Add("Edit", "xs y+5 w120 vInactiveClientBorderColor -Wrap", This.InactiveClientBorderColor)
        ThumbnailSettings.Push This.MainFrame.Add("Text", " x+5 yp+3 ", "Hex or RGB")
        This.MainFrame["InactiveClientBorderColor"].OnEvent("Change", (obj, *) => ThumbnailSettings_EventHandler(obj))

        ; /////////////////////////////////
        RowSpacing := 30
        SeparatorSpacing := 10
        FieldOffset := 289
        DefaultWidth := 180
        DigitWidth := 40

        ThumbnailSettings.Push This.MainFrame.Add("Text", "x16 ys+" . RowSpacing . " Section", "Thumbnail Background Color (Hex/RGB)")

        ThumbnailSettings.Push This.MainFrame.Add("Edit", "xp+" . FieldOffset . " ys w" . DefaultWidth . " vThumbnailBackgroundColor", This.ThumbnailBackgroundColor)
        This.MainFrame["ThumbnailBackgroundColor"].OnEvent("Change", (obj, *) => ThumbnailSettings_EventHandler(obj))
        
        ThumbnailSettings.Push This.MainFrame.Add("Text", "xs ys+" . RowSpacing . " Section", "Default Thumbnail Position (px)")
        
        ThumbnailSettings.Push This.MainFrame.Add("Text", "xp+" . FieldOffset . " ys", "x:")
        ThumbnailSettings.Push This.MainFrame.Add("Edit", "x+5 y+-18 w" . DigitWidth . " vThumbnailStartLocationx", This.ThumbnailStartLocation["x"])
        This.MainFrame["ThumbnailStartLocationx"].OnEvent("Change", (obj, *) => ThumbnailSettings_EventHandler(obj))
        
        ThumbnailSettings.Push This.MainFrame.Add("Text", "x+8 ys ", "y:")
        ThumbnailSettings.Push This.MainFrame.Add("Edit", "x+5 y+-18 w" . DigitWidth . " vThumbnailStartLocationy", This.ThumbnailStartLocation["y"])
        This.MainFrame["ThumbnailStartLocationy"].OnEvent("Change", (obj, *) => ThumbnailSettings_EventHandler(obj))

        ThumbnailSettings.Push This.MainFrame.Add("Text", "xs ys+" . RowSpacing . " Section", "Default Thumbnail Size (px)")

        ThumbnailSettings.Push This.MainFrame.Add("Text", "xp+" . FieldOffset . " ys ", "width:")
        ThumbnailSettings.Push This.MainFrame.Add("Edit", "x+5 y+-18 w" . DigitWidth . " vThumbnailStartLocationwidth", This.ThumbnailStartLocation["width"])
        This.MainFrame["ThumbnailStartLocationwidth"].OnEvent("Change", (obj, *) => ThumbnailSettings_EventHandler(obj))

        ThumbnailSettings.Push This.MainFrame.Add("Text", "x+8 ys ", "height:")
        ThumbnailSettings.Push This.MainFrame.Add("Edit", "x+5 y+-18 w" . DigitWidth . " vThumbnailStartLocationheight", This.ThumbnailStartLocation["height"])
        This.MainFrame["ThumbnailStartLocationheight"].OnEvent("Change", (obj, *) => ThumbnailSettings_EventHandler(obj))

        ThumbnailSettings.Push This.MainFrame.Add("Text", "xs ys+" . RowSpacing . " Section", "Minimum Thumbnail Size (px)")

        ThumbnailSettings.Push This.MainFrame.Add("Text", "xp+" . FieldOffset . " ys", "width:")
        ThumbnailSettings.Push This.MainFrame.Add("Edit", "x+5 y+-18 w" . DigitWidth . " vThumbnailMinimumSizewidth", This.ThumbnailMinimumSize["width"])
        This.MainFrame["ThumbnailMinimumSizewidth"].OnEvent("Change", (obj, *) => ThumbnailSettings_EventHandler(obj))

        ThumbnailSettings.Push This.MainFrame.Add("Text", "x+8 ys ", "height:")
        ThumbnailSettings.Push This.MainFrame.Add("Edit", "x+5 y+-18 w" . DigitWidth . " vThumbnailMinimumSizeheight", This.ThumbnailMinimumSize["height"])
        This.MainFrame["ThumbnailMinimumSizeheight"].OnEvent("Change", (obj, *) => ThumbnailSettings_EventHandler(obj))

        ThumbnailSettings.Push This.MainFrame.Add("Text", "xs ys+" . RowSpacing . " Section", "Enable Thumbnail Snap")

        ThumbnailSettings.Push This.MainFrame.Add("Radio", "xp+" . FieldOffset . " ys w37 vThumbnailSnapOn Checked" This.ThumbnailSnap, "On")
        ThumbnailSettings.Push This.MainFrame.Add("Radio", " xp+50 yp w37 vThumbnailSnapOff Checked" (This.ThumbnailSnap ? 0 : 1), "Off")
        This.MainFrame["ThumbnailSnapOn"].OnEvent("Click", (obj, *) => ThumbnailSettings_EventHandler(obj))
        This.MainFrame["ThumbnailSnapOff"].OnEvent("Click", (obj, *) => ThumbnailSettings_EventHandler(obj))

        ThumbnailSettings.Push This.MainFrame.Add("Text", "xs ys+" . RowSpacing . " Section", "Thumbnail Snap Distance (px)")

        ThumbnailSettings.Push This.MainFrame.Add("Edit", "xp+" . FieldOffset . " ys w" . DigitWidth . " vThumbnailSnap_Distance", This.ThumbnailSnap_Distance)
        This.MainFrame["ThumbnailSnap_Distance"].OnEvent("Change", (obj, *) => ThumbnailSettings_EventHandler(obj))

        ThumbnailSettings.Push This.MainFrame.Add("Text", "xs ys+" . RowSpacing . " Section", "Hide Thumbnail for Active Window")

        ThumbnailSettings.Push This.MainFrame.Add("CheckBox", "xp+" . FieldOffset . " ys vHideThumbForActiveWin Checked" This.HideThumbForActiveWin, "On/Off")
        This.MainFrame["HideThumbForActiveWin"].OnEvent("Click", (obj, *) => ThumbnailSettings_EventHandler(obj))

        ; Category Separator
        ThumbnailSettings.Push This.MainFrame.Add("Text", "xs ys+" . RowSpacing . " Section w473 h2 +0x10")
        
        ThumbnailSettings.Push This.MainFrame.Add("Text", "xs ys+" . SeparatorSpacing . " Section", "Shift Thumbnails on Login Screen")

        ThumbnailSettings.Push This.MainFrame.Add("CheckBox", "xp+" . FieldOffset . " ys vShiftThumbsForLoginScreen Checked" This.ShiftThumbsForLoginScreen, "On/Off")
        This.MainFrame["ShiftThumbsForLoginScreen"].OnEvent("Click", (obj, *) => ThumbnailSettings_EventHandler(obj))

        ThumbnailSettings.Push This.MainFrame.Add("Text", "xs ys+" . RowSpacing . " Section", "Enable Thumbnail Collision Avoidance")

        ThumbnailSettings.Push This.MainFrame.Add("CheckBox", "xp+" . FieldOffset . " ys vShiftThumbsCollisionCheck Checked" This.ShiftThumbsCollisionCheck, "On/Off")
        This.MainFrame["ShiftThumbsCollisionCheck"].OnEvent("Click", (obj, *) => ThumbnailSettings_EventHandler(obj))

        ThumbnailSettings.Push This.MainFrame.Add("Text", "xs ys+" . RowSpacing . " Section", "Thumbnail Shift Direction")

        ThumbnailSettings.Push This.MainFrame.Add("DropDownList", "xp+" . FieldOffset . " ys w" . DefaultWidth . " vShiftThumbsDirection Choose" . Integer(This.ShiftThumbsDirection), This.ddl_options)
        This.MainFrame["ShiftThumbsDirection"].OnEvent("Change", (obj, *) => ThumbnailSettings_EventHandler(obj))

        ThumbnailSettings.Push This.MainFrame.Add("Text", "xs ys+" . RowSpacing . " Section", "Thumbnail Shift Horizontal Step (px)")

        ThumbnailSettings.Push This.MainFrame.Add("Edit", "xp+" . FieldOffset . " ys w" . DigitWidth . " vShiftThumbHorizontalStep 0", This.ShiftThumbHorizontalStep)
        ThumbnailSettings.Push This.MainFrame.Add("Text", "x+8 ys+3 ", "0 = Width")
        This.MainFrame["ShiftThumbHorizontalStep"].OnEvent("Change", (obj, *) => ThumbnailSettings_EventHandler(obj))

        ThumbnailSettings.Push This.MainFrame.Add("Text", "xs ys+" . RowSpacing . " Section", "Thumbnail Shift Vertical Step (px)")

        ThumbnailSettings.Push This.MainFrame.Add("Edit", "xp+" . FieldOffset . " ys w" . DigitWidth . " vShiftThumbVerticalStep 0", This.ShiftThumbVerticalStep)
        ThumbnailSettings.Push This.MainFrame.Add("Text", "x+8 ys+3 ", "0 = Height")
        This.MainFrame["ShiftThumbVerticalStep"].OnEvent("Change", (obj, *) => ThumbnailSettings_EventHandler(obj))

        ThumbnailSettings.Push This.MainFrame.Add("Text", "xs ys+" . RowSpacing . " Section", "Preserve Character Name On logout")

        ThumbnailSettings.Push This.MainFrame.Add("CheckBox", "xp+" . FieldOffset . " ys vPreserveCharNameOnLogout Checked" This.PreserveCharNameOnLogout, "On/Off")
        This.MainFrame["PreserveCharNameOnLogout"].OnEvent("Click", (obj, *) => ThumbnailSettings_EventHandler(obj))

        ThumbnailSettings.Push This.MainFrame.Add("Text", "xs ys+" . RowSpacing . " Section", "Preserve Thumbnail Position on Logout")

        ThumbnailSettings.Push This.MainFrame.Add("CheckBox", "xp+" . FieldOffset . " ys vPreserveThumbPosOnLogout Checked" This.PreserveThumbPosOnLogout, "On/Off")
        This.MainFrame["PreserveThumbPosOnLogout"].OnEvent("Click", (obj, *) => ThumbnailSettings_EventHandler(obj))

        This.MainFrame.Group["Thumbnail Settings"] := ThumbnailSettings
        for k, v in This.MainFrame.Group["Thumbnail Settings"] {
            v.Visible := 0
        }

        ;Parse All hotkeys to a Array on value change
        ThumbnailSettings_EventHandler(obj) {
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
            else if (obj.name = "HideThumbnailsOnLostFocus") {
                This.HideThumbnailsOnLostFocus := obj.value
            }
            else if (obj.name = "ThumbnailOpacity") {
                This.ThumbnailOpacity := obj.value
                This.NeedRestart := 1
            }
            else if (obj.name = "ShowThumbnailsAlwaysOnTop") {
                This.ShowThumbnailsAlwaysOnTop := obj.value
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
            else if (obj.name = "ThumbnailMinimumSizewidth") {
                This.ThumbnailMinimumSize["width"] := obj.value
            }
            else if (obj.name = "ThumbnailMinimumSizeheight") {
                This.ThumbnailMinimumSize["height"] := obj.value
            }
            else if (obj.name = "ThumbnailSnapOn") {
                This.ThumbnailSnap := 1
            }
            else if (obj.name = "ThumbnailSnapOff") {
                This.ThumbnailSnap := 0
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

    ThumbnailVisibility_Ctrl() {
        This.MainFrame.Group["Thumbnail Visibility"] := [], Thumbnail_visibility := []

        Thumbnail_visibility.Push This.MainFrame.Add("Text", "x16 y16 w250", "Select any Client to hide the Thumbnail")
        This.Tv_LV := This.MainFrame.Add("ListView", "xp+15 yp+30 w210 Checked -LV0x10 -Multi r20 -Sort vVisibility_List", ["Client Name       "])
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

    ExcludeFromClosing_Ctrl() {
        This.MainFrame.Group["Exclude from Closing"] := [], ExcludeFromClosing := []

        ExcludeFromClosing.Push This.MainFrame.Add("Text", "x16 y16 Section ", "Exclude on Login Screen:")
        ExcludeFromClosing.Push This.MainFrame.Add("Text", "xs y+15 ", "Don't Close Clients:")

        ExcludeFromClosing.Push This.MainFrame.Add("CheckBox", "xs+230 ys Section vExcludeOnLoginScreen Checked" This.ExcludeOnLoginScreen, "On/Off")
        This.MainFrame["ExcludeOnLoginScreen"].OnEvent("Click", (obj, *) => cExcludeFromClosing_EventHandler(obj))

        ExcludeFromClosing.Push This.MainFrame.Add("Edit", "xs y+15 w220 h180 vDontCloseClients -Wrap", This.DontCloseList())
        This.MainFrame["DontCloseClients"].OnEvent("Change", (obj, *) => cExcludeFromClosing_EventHandler(obj))

        cExcludeFromClosing_EventHandler(obj) {
            if (obj.name = "ExcludeOnLoginScreen") {
                This.ExcludeOnLoginScreen := obj.value
            }
            else if (obj.name = "DontCloseClients") {
                This.DontCloseClients := obj.value
            }
            SetTimer(This.Save_Settings_Delay_Timer, -200)
        }

        This.MainFrame.Group["Exclude from Closing"] := ExcludeFromClosing
        for k, v in This.MainFrame.Group["Exclude from Closing"]
            v.Visible := 0
    }

    Other_Ctrl() {
        This.MainFrame.Group["Other"] := [], Other := []

        Other.Push This.MainFrame.Add("Text", "x16 y16 Section", "Switch Language to English on Error (Unstable!)")

        Other.Push This.MainFrame.Add("CheckBox", "xp+50 ys vSwitchLangOnErr Checked" This.SwitchLangOnErr, "On/Off")
        This.MainFrame["SwitchLangOnErr"].OnEvent("Click", (obj, *) => cOther_EventHandler(obj))

        Other.Push This.MainFrame.Add("Text", "xs ys+16 Section", "Check for Updates on Startup")

        Other.Push This.MainFrame.Add("CheckBox", "xp+50 ys vCheck_Updates Checked" This.Check_Updates, "On/Off")
        This.MainFrame["Check_Updates"].OnEvent("Click", (obj, *) => cOther_EventHandler(obj))

        cOther_EventHandler(obj) {
            if (obj.name = "SwitchLangOnErr") {
                This.SwitchLangOnErr := obj.value
            }
            else if (obj.name = "Check_Updates") {
                This.Check_Updates := obj.value
            }
            SetTimer(This.Save_Settings_Delay_Timer, -200)
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

        ; if (This.Sidebar["SelectedProfile"].Text = "Default") {
        ;     for k, v in This.MainFrame.Group {
        ;         for _, ob in v {
        ;             ob.Enabled := 0
        ;         }
        ;     }
        ; }
        SetTimer(This.Save_Settings_Delay_Timer, -200)
    }

    Refresh_ControlValues() {

        ;Client Settings
        This.MainFrame["MinimizeInactiveClients"].value := This.MinimizeInactiveClients
        This.MainFrame["AlwaysMaximize"].value := This.AlwaysMaximize
        This.MainFrame["Dont_Minimize_Clients"].value := This.Dont_Minimize_List()
        This.MainFrame["Minimizeclients_Delay"].value := This.Minimizeclients_Delay

        ;Custom Colors
        This.MainFrame["Ccoloractive"].value := This.CustomColorsActive
        This.MainFrame["Cchars"].value := This.CustomColors_AllCharNames
        This.MainFrame["CBorderColor"].value := This.CustomColors_AllBColors
        This.MainFrame["CTextColor"].value := This.CustomColors_AllTColors
        This.MainFrame["IABorderColor"].value := This.CustomColors_IABorder_Colors

        ;Hotkey Groups
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
        This.MainFrame["PreserveHotkeysOnLogout"].value := This.PreserveHotkeysOnLogout
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
        This.MainFrame["ThumbnailSnapOn"].value := This.ThumbnailSnap
        This.MainFrame["ThumbnailSnapOff"].value := (This.ThumbnailSnap ? 0 : 1)
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

        ;Exclude from Closing
        This.MainFrame["ExcludeOnLoginScreen"].value := This.ExcludeOnLoginScreen
        This.MainFrame["DontCloseClients"].value := This.DontCloseList()

        ; Other
        This.MainFrame["SwitchLangOnErr"].value := This.SwitchLangOnErr
        This.MainFrame["Check_Updates"].value := This.Check_Updates

        for k, v in This.MainFrame.Group {
            for _, ob in v {
                ob.Enabled := 1
            }
            This.MainFrame["HKCharlist"].Enabled := 0
            This.MainFrame["ForwardsKey"].Enabled := 0
            This.MainFrame["BackwardsdKey"].Enabled := 0
        }
        This.MainFrame["InactiveClientBorderthickness"].Enabled := This.ShowAllColoredBorders
        This.MainFrame["InactiveClientBorderColor"].Enabled := This.ShowAllColoredBorders
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
