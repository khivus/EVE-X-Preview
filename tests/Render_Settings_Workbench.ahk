#Requires AutoHotkey v2.0
#SingleInstance Force
; Render isolated fixture windows, without loading or writing user settings.
global __EVE_X_PREVIEW_TESTING__ := true
#Include "..\Main.ahk"

class SettingsRenderOwner extends Main_Class {
    __New() {
        this._JSON := JSON.Load(default_JSON)
        this._JSON["SettingsTheme"] := "light"
        this.ProfileOverride()
        this.Save_Settings_Delay_Timer := (*) => 0
    }
    SaveJsonToFile(*) => 0
}

RenderWorkbenchPng(hwnd, path, app?) {
    previousCritical := A_IsCritical
    Critical(true)
    module := DllCall("LoadLibraryW", "WStr", "gdiplus.dll", "Ptr")
    WinGetPos(&windowX, &windowY, &width, &height, "ahk_id " hwnd)
    dc := DllCall("GetWindowDC", "Ptr", hwnd, "Ptr")
    memory := DllCall("CreateCompatibleDC", "Ptr", dc, "Ptr")
    bitmap := DllCall("CreateCompatibleBitmap", "Ptr", dc, "Int", width, "Int", height, "Ptr")
    previous := DllCall("SelectObject", "Ptr", memory, "Ptr", bitmap, "Ptr")
    token := 0, image := 0
    try {
        if !module || !dc || !memory || !bitmap
            throw OSError()
        if !DllCall("PrintWindow", "Ptr", hwnd, "Ptr", memory, "UInt", 2)
            throw OSError()
        ; WM_PRINT requests native child controls explicitly; some Windows
        ; sessions render only the non-client frame through PrintWindow.
        DllCall("SendMessageW", "Ptr", hwnd, "UInt", 0x317, "Ptr", memory, "Ptr", 0x1E)
        if IsSet(app) {
            ; WM_PRINT does not clip descendants to nested GUI viewports.
            ; Repaint the real root footer after printing the scroll canvas.
            point := Buffer(8, 0), client := Buffer(16, 0)
            DllCall("ClientToScreen", "Ptr", hwnd, "Ptr", point)
            DllCall("GetClientRect", "Ptr", hwnd, "Ptr", client)
            originX := NumGet(point, 0, "Int") - windowX
            originY := NumGet(point, 4, "Int") - windowY
            footerTop := originY + Round(app.bodyHeight * A_ScreenDPI / 96)
            footerLeft := originX + Round(app.sidebarW * A_ScreenDPI / 96)
            footer := Buffer(16, 0)
            NumPut("Int", footerLeft, "Int", footerTop, "Int", originX + NumGet(client, 8, "Int"), "Int", originY + NumGet(client, 12, "Int"), footer)
            brush := DllCall("CreateSolidBrush", "UInt", SettingsTheme.BGR(app.theme.background), "Ptr")
            DllCall("FillRect", "Ptr", memory, "Ptr", footer, "Ptr", brush)
            DllCall("DeleteObject", "Ptr", brush)
            for control in app.S_Gui {
                rect := Buffer(16, 0)
                DllCall("GetWindowRect", "Ptr", control.Hwnd, "Ptr", rect)
                x := NumGet(rect, 0, "Int") - windowX, y := NumGet(rect, 4, "Int") - windowY
                if y < footerTop
                    continue
                savedDc := DllCall("SaveDC", "Ptr", memory, "Int")
                DllCall("IntersectClipRect", "Ptr", memory, "Int", footerLeft, "Int", footerTop,
                    "Int", originX + NumGet(client, 8, "Int"), "Int", originY + NumGet(client, 12, "Int"))
                DllCall("SetViewportOrgEx", "Ptr", memory, "Int", x, "Int", y, "Ptr", 0)
                DllCall("SendMessageW", "Ptr", control.Hwnd, "UInt", 0x317, "Ptr", memory, "Ptr", 0x1E)
                DllCall("RestoreDC", "Ptr", memory, "Int", savedDc)
            }
        }
        startup := Buffer(A_PtrSize = 8 ? 24 : 16, 0)
        NumPut("UInt", 1, startup)
        if DllCall("gdiplus\GdiplusStartup", "Ptr*", &token, "Ptr", startup, "Ptr", 0)
            throw Error("GDI+ startup failed.")
        if DllCall("gdiplus\GdipCreateBitmapFromHBITMAP", "Ptr", bitmap, "Ptr", 0, "Ptr*", &image)
            throw Error("Bitmap conversion failed.")
        encoder := Buffer(16, 0)
        DllCall("ole32\CLSIDFromString", "WStr", "{557CF406-1A04-11D3-9A73-0000F81EF32E}", "Ptr", encoder)
        if DllCall("gdiplus\GdipSaveImageToFile", "Ptr", image, "WStr", path, "Ptr", encoder, "Ptr", 0)
            throw Error("PNG export failed.")
    } finally {
        if image
            DllCall("gdiplus\GdipDisposeImage", "Ptr", image)
        if token
            DllCall("gdiplus\GdiplusShutdown", "Ptr", token)
        DllCall("SelectObject", "Ptr", memory, "Ptr", previous)
        DllCall("DeleteObject", "Ptr", bitmap)
        DllCall("DeleteDC", "Ptr", memory)
        DllCall("ReleaseDC", "Ptr", hwnd, "Ptr", dc)
        if module
            DllCall("FreeLibrary", "Ptr", module)
        Critical(previousCritical)
    }
}

try {
    TraySetIcon(A_ScriptDir "\..\res\Icon.ico")
    output := A_Args.Length ? A_Args[1] : A_Temp "\EVE-X-Preview-settings-render"
    DirCreate(output)
    owner := SettingsRenderOwner()
    app := SettingsWorkbench(owner, false)
    app.model.SaveRows("character-shortcuts", [["Sample Scout", "^!F1"], ["Sample Logistics", "^!F2"]])
    app.model.SaveRows("custom-names", [["Sample Scout", "Scout"], ["Sample Logistics", "Support"]])
    app.model.SaveRows("character-colors", [["Sample Scout", "FF8800", "808080", "FFFFFF"], ["Sample Logistics", "42C8F5", "111111", "96E6A1"]])
    app.model.AddGroup("characters", "Sample fleet")
    app.model.SaveRows("group-members", [["Sample Scout"], ["Sample Logistics"]], "Sample fleet")
    app.Changed()
    app.width := 1000, app.height := 740
    app.S_Gui.Show("NoActivate x-3000 y-3000 w1000 h740")
    app.Layout()
    for section in ["appearance", "character-appearance", "labels", "character-groups", "events", "sounds", "dps", "close-protection", "quick-group", "profiles", "performance"] {
        app.Navigate(section)
        Sleep(80)
        RenderWorkbenchPng(app.S_Gui.Hwnd, output "\" section ".png", app)
    }
    app.model.Set("incomingDPSEnabled", 1)
    app.width := 800
    app.S_Gui.Show("NoActivate x-3000 y-3000 w800 h740")
    app.Layout()
    app.Navigate("dps")
    RenderWorkbenchPng(app.S_Gui.Hwnd, output "\dps-compact.png", app)
    app.width := 1000
    app.S_Gui.Show("NoActivate x-3000 y-3000 w1000 h740")
    app.Layout()
    app.Navigate("character-shortcuts")
    grid := app.tableEditors[1].grid
    grid.Select(1, 2)
    grid.Select(2, 2, true)
    RenderWorkbenchPng(app.S_Gui.Hwnd, output "\grid-selection.png", app)
    grid.Select(1, 1)
    grid.BeginEdit()
    RenderWorkbenchPng(app.S_Gui.Hwnd, output "\grid-editor.png", app)
    grid.CancelEditor()
    app.Navigate("appearance")
    app.settingsScroller.ScrollTo(100000)
    RenderWorkbenchPng(app.S_Gui.Hwnd, output "\appearance-borders.png", app)
    app.Navigate("character-shortcuts")
    app.tableEditors[1].Edit(1)
    WinMove(-3000, -3000,,, app.activeDialog.Hwnd)
    RenderWorkbenchPng(app.activeDialog.Hwnd, output "\shortcut-editor.png")
    app.CloseChild(app.activeDialog)
    app.Navigate("interface")
    for mode in ["light", "dark", "light"] {
        for binding in app.bindings
            if binding.field.id = "SettingsTheme" {
                binding.control.Choose(mode = "light" ? 2 : 3)
                app.ChoiceChanged(binding, binding.control)
            }
        RenderWorkbenchPng(app.S_Gui.Hwnd, output "\theme-switch-" mode ".png", app)
    }
    app.Navigate("appearance")
    app.search.Value := "border"
    app.RunSearch()
    RenderWorkbenchPng(app.S_Gui.Hwnd, output "\search-highlight-light.png", app)
    app.model.Set("SettingsTheme", "dark")
    app.Changed()
    RenderWorkbenchPng(app.S_Gui.Hwnd, output "\search-highlight-dark.png", app)
    app.search.Value := ""
    app.RunSearch()
    app.Navigate("shortcuts")
    for binding in app.bindings
        if binding.field.id = "HideThumbnailsHotkey" {
            app.AdvancedShortcut(binding.helper)
            binding.control.Value := "^!F23"
            app.CommitInlineShortcut(binding.helper)
        }
    RenderWorkbenchPng(app.S_Gui.Hwnd, output "\shortcut-inline.png", app)
    app.CreateProfile("Sample inherited")
    app.model.SetInheritance("appearance", true)
    app.Navigate("appearance")
    RenderWorkbenchPng(app.S_Gui.Hwnd, output "\inherited-appearance.png", app)
    app.Navigate("inheritance")
    grid := app.grids[1]
    grid.Select(1, 2)
    grid.Select(3, 2, true)
    RenderWorkbenchPng(app.S_Gui.Hwnd, output "\inheritance-selection.png", app)
    app.SelectProfile("Default")
    app.Navigate("events")
    event := SettingsEventEditor(app.eventTable, "underAttackByPlayer", 2)
    event.Show()
    WinMove(-3000, -3000,,, app.activeDialog.Hwnd)
    RenderWorkbenchPng(app.activeDialog.Hwnd, output "\event-editor.png")
    app.activeDialog.root.Show("NoActivate x-3000 y-3000 w500 h400")
    app.activeDialog.Layout(500, 400)
    app.activeDialog.scroll.ScrollTo(100000)
    RenderWorkbenchPng(app.activeDialog.Hwnd, output "\event-editor-compact.png")
    app.CloseChild(app.activeDialog)
    for width in [780, 1200, 800, 1000] {
        app.width := width
        app.S_Gui.Show("NoActivate x-3000 y-3000 w" width " h740")
        app.Layout()
    }
    FileAppend("Footer after resize: dirty=" app.model.Dirty "; errors=" app.errors.Count "; tableErrors=" app.tableErrors.Count "; relations=" app.relationIssues.Length "; applyEnabled=" app.applyButton.Enabled "`n", "*")
    RenderWorkbenchPng(app.S_Gui.Hwnd, output "\resize-footer.png", app)
    app.width := 800, app.height := 500
    app.S_Gui.Show("NoActivate x-3000 y-3000 w800 h500")
    app.Layout()
    app.Navigate("character-groups")
    Sleep(80)
    RenderWorkbenchPng(app.S_Gui.Hwnd, output "\compact.png", app)
    app.width := 1000, app.height := 740
    app.S_Gui.Show("NoActivate x-3000 y-3000 w1000 h740")
    app.Layout()
    app.model.Set("SettingsTheme", "dark")
    app.Changed()
    for section in ["appearance", "character-appearance", "events", "dps", "sounds", "character-groups", "about", "inheritance"] {
        app.Navigate(section)
        RenderWorkbenchPng(app.S_Gui.Hwnd, output "\" section "-dark.png", app)
    }
    app.SelectProfile("Default")
    app.Navigate("inheritance")
    RenderWorkbenchPng(app.S_Gui.Hwnd, output "\inheritance-template-dark.png", app)
    app.AskName("Create profile", "", ObjBindMethod(app, "CreateProfile"), true)
    WinMove(-3000, -3000,,, app.activeDialog.Hwnd)
    app.activeDialog.profileCreation.name.Value := "Sample profile"
    RenderWorkbenchPng(app.activeDialog.Hwnd, output "\profile-creation-dark.png")
    app.CloseChild(app.activeDialog)
    app.Navigate("interface")
    for binding in app.bindings
        if binding.field.id = "SettingsTheme" {
            SendMessage(0x14F, 1, 0, binding.control)
            info := app.theme.chrome.ComboInfo(binding.control.Hwnd)
            popup := NumGet(info, A_PtrSize = 8 ? 56 : 48, "Ptr")
            RenderWorkbenchPng(popup, output "\dropdown-popup-dark.png")
            SendMessage(0x14F, 0, 0, binding.control)
        }
    app.theme.ShowTooltip("Choose a light or dark interface, or follow the Windows app theme.", -2800, -2800)
    RenderWorkbenchPng(app.theme.tooltip.Hwnd, output "\tooltip-dark.png")
    app.theme.HideTooltip()
    app.SelectProfile("Sample inherited")
    app.Navigate("appearance")
    app.EditDefaultHere("appearance")
    RenderWorkbenchPng(app.S_Gui.Hwnd, output "\default-edit-dark.png", app)
    app.EndDefaultEditing()
    app.updateChecker.releaseTag := "9.9.0"
    app.UpdateVersionControls()
    RenderWorkbenchPng(app.S_Gui.Hwnd, output "\update-notice-dark.png", app)
    app.SelectProfile("Default")
    app.Navigate("events")
    event := SettingsEventEditor(app.eventTable, "underAttackByPlayer", 2)
    event.Show()
    WinMove(-3000, -3000,,, app.activeDialog.Hwnd)
    RenderWorkbenchPng(app.activeDialog.Hwnd, output "\event-editor-dark.png")
    app.CloseChild(app.activeDialog)
    app.width := 800, app.height := 500
    app.S_Gui.Show("NoActivate x-3000 y-3000 w800 h500")
    app.Layout()
    app.ResizeContent()
    app.Navigate("appearance")
    RenderWorkbenchPng(app.S_Gui.Hwnd, output "\resize-compact-dark.png", app)
    app.SelectProfile("Sample inherited")
    app.model.ActivateProfile()
    app.Changed()
    app.Navigate("profiles")
    SendMessage(0xF4, 1, 1, app.applyButton)
    RenderWorkbenchPng(app.S_Gui.Hwnd, output "\profile-activation-dark.png", app)
    app.Dispose()
    FileAppend("Settings snapshots: " output "`n", "*")
    ExitApp(0)
} catch as err {
    if IsSet(app)
        app.Dispose()
    FileAppend(err.Message "`n" err.Stack "`n", "*")
    ExitApp(1)
}
