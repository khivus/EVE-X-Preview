class SettingsWorkbenchPages extends Main_Class {
    QuickGroupGesture() {
        this.Subheading("Quick Group thumbnail gesture")
        this.Note("Use this gesture on a thumbnail to add or remove its character. This is also editable under Thumbnails / Mouse actions.")
        options := [], selected := 1
        rows := this.RowsForEdit("gestures")
        for index, option in SettingsTables.GestureOptions {
            options.Push(option.label)
            if option.value = rows[-1][2]
                selected := index
        }
        input := this.MainFrame.AddDropDownList("x20 y" this.y " w" (this.contentWidth - 40), options)
        input.Choose(selected)
        input.Enabled := !this.model.Inherited("Thumbnails Interactions")
        input.ToolTip := "Toggle Quick Group membership using this gesture on a thumbnail."
        input.SettingsHelp := input.ToolTip
        this.y += 34
        if this.HasOwnProp("quickGestureError") && this.quickGestureError != ""
            this.Note(this.quickGestureError, "B42318")
        this.Action("Open all thumbnail mouse actions", (*) => this.Navigate("mouse-actions"), 320)
        if this.model.Get("ClickThroughActive")
            this.Note("Click-through is on. Turn it off in Mouse actions to use thumbnail gestures.", "9A5700")
        input.OnEvent("Change", Change)
        Change(*) {
            rows := this.RowsForEdit("gestures")
            previous := rows[-1][2]
            rows[-1][2] := SettingsTables.GestureOptions[input.Value].value
            try {
                this.model.SaveRows("gestures", rows)
                key := this.model.profile "|gestures|"
                if this.tableStates.Has(key)
                    this.tableStates.Delete(key)
                if this.tableErrors.Has(key)
                    this.tableErrors.Delete(key)
                this.quickGestureError := ""
                this.Changed()
                this.Render()
            } catch as err {
                this.quickGestureError := err.Message
                for index, option in SettingsTables.GestureOptions
                    if option.value = previous
                        input.Choose(index)
                this.Render()
            }
        }
    }

    RenderSpecial(id) {
        switch id {
            case "appearance": this.BorderDiagram()
            case "labels":
                this.reshuffleButton := this.Action("Reshuffle preview labels", (*) => this.ReshufflePreview())
                this.Note("Anonymous labels take priority. Custom character labels are restored when anonymization is turned off.", "555555", 38)
                this.Subheading("Custom character labels")
                this.AddRowTable("custom-names")
            case "character-appearance":
                if !this.model.Get("Ccoloractive")
                    this.Note("Character-specific colors are off. You can prepare rows now, then enable the option above to use them.", "9A5700", 40)
                this.AddRowTable("character-colors")
            case "layout":
                this.recoverThumbnailsButton := this.Action("Recover off-screen thumbnails", (*) => this.RecoverRuntimeThumbnails())
                this.Note("Moves current thumbnails back on screen immediately, even when positions are locked. Thumbnails fully visible across connected monitors stay in place.", "555555", 54)
                this.Action("Reset this profile's thumbnail sizes on Apply", (*) => this.StageSizeReset(), 340)
            case "hidden-characters": this.AddRowTable("hidden-characters")
            case "mouse-actions":
                this.mouseActionNotice := this.Note("", "9A5700", 40)
                this.AddRowTable("gestures")
            case "shortcuts": this.ConflictNotice()
            case "character-shortcuts":
                this.AddRowTable("character-shortcuts")
                this.ConflictNotice()
            case "character-groups": this.GroupEditor("characters")
            case "app-groups": this.GroupEditor("apps")
            case "app-shortcuts":
                this.AddRowTable("app-shortcuts")
                this.ConflictNotice()
            case "quick-group", "switching-behavior": this.ConflictNotice()
            case "close-protection":
                this.Subheading("Character exceptions")
                this.AddRowTable("exceptions")
            case "window-behavior": this.Action("Capture open EVE window placement to draft", (*) => this.StageClientPlacement(), 380)
            case "monitored-characters":
                this.monitoringNotice := this.Note(this.MonitoringStatus(), "9A5700", 54)
                this.AddRowTable("monitored-characters")
            case "sources": this.SourceStatus()
            case "events": SettingsEventTable(this).Build()
            case "sounds": SettingsSoundLibrary(this).Build()
            case "profiles": this.ProfileManager()
            case "inheritance": this.InheritanceEditor()
            case "tray": this.TrayCommands()
            case "diagnostics": this.Diagnostics()
            case "about": this.AboutUpdates()
        }
    }

    AddRowTable(kind, groupName := "", ordered := false) {
        editor := SettingsRowTable(this, kind, groupName, ordered)
        this.tableEditors.Push(editor)
        editor.Build()
        return editor
    }

    AddRecorder(gui, field, y, width, onComplete, value := "", context?) {
        stacked := width < 480
        if stacked
            field.Move(,, width)
        field.GetPos(&x,, &fieldWidth, &fieldHeight)
        buttonY := stacked ? y + 34 : y
        recordX := stacked ? x : x + fieldWidth + 8
        recordW := stacked ? Max(60, Floor((width - 16) * 0.34)) : 76
        clearW := stacked ? Max(48, Floor((width - 16) * 0.26)) : 62
        advancedW := stacked ? Max(74, width - recordW - clearW - 16) : 92
        record := gui.AddButton("x" recordX " y" buttonY " w" recordW " h26", "Record")
        clear := gui.AddButton("x" (recordX + recordW + 8) " y" buttonY " w" clearW " h26", "Clear")
        advanced := gui.AddButton("x" (recordX + recordW + clearW + 16) " y" buttonY " w" advancedW " h26", "Advanced")
        waiting := gui.AddText("x" x " y" y " w" fieldWidth " h26 Hidden Border Center +0x300 BackgroundE0E0E0", "")
        record.ToolTip := "Record a shortcut. Press Esc or Record again to cancel."
        clear.ToolTip := "Remove this shortcut assignment."
        advanced.ToolTip := "Unlock this field to edit AutoHotkey syntax inline. Done returns to readable key names."
        waiting.ToolTip := "Press and release a shortcut. Mouse buttons and the wheel can be recorded over this area."
        meta := IsSet(context) ? context : {key: "dialog|" gui.Hwnd "|" field.Hwnd, group: ""}
        rawKey := this.model.profile "|" meta.key
        helper := {field: field, button: record, clear: clear, advanced: advanced, waiting: waiting,
            multiline: false, gui: gui, key: value, complete: onComplete, extraHeight: stacked ? 34 : 0,
            rawMode: false, updating: false, rawKey: rawKey, meta: meta, profile: this.model.profile, section: this.currentSection}
        field.OnEvent("Change", ObjBindMethod(this, "InlineShortcutChanged", helper))
        if this.shortcutDrafts.Has(rawKey) {
            helper.rawMode := true
            helper.updating := true
            field.Opt("-ReadOnly")
            field.Value := this.shortcutDrafts[rawKey].text
            advanced.Text := "Done"
            helper.updating := false
        }
        record.OnEvent("Click", ObjBindMethod(this, "RecordShortcut", helper))
        clear.OnEvent("Click", (*) => this.FinishShortcut(helper, ""))
        advanced.OnEvent("Click", ObjBindMethod(this, "AdvancedShortcut", helper))
        this.recorders.Push(helper)
        return helper
    }

    RecordShortcut(helper, *) {
        if HotkeyCapture.Active {
            same := HotkeyCapture.Current.helper = helper
            HotkeyCapture.Current.Cancel()
            if same
                return
        }
        if helper.field.Enabled
            HotkeyCapture().Start(helper.gui, helper, ObjBindMethod(this, "RecordedShortcut", helper))
    }

    RecordedShortcut(helper, value) {
        if value != ""
            this.FinishShortcut(helper, value)
    }

    FinishShortcut(helper, value) {
        helper.updating := true
        try {
            SettingsValue.Hotkey(value)
            helper.complete.Call(value)
            helper.key := value
            helper.field.Value := helper.rawMode ? value : SettingsValue.HotkeyText(value)
            if this.shortcutErrors.Has(helper.rawKey)
                this.shortcutErrors.Delete(helper.rawKey)
            if helper.rawMode
                this.StoreShortcutDraft(helper, value)
            if helper.meta.HasOwnProp("binding")
                this.PaintError(helper.meta.binding, "")
            this.UpdateStatus()
        } catch as err
            this.ShortcutProblem(helper, err.Message)
        finally helper.updating := false
    }
    StoreShortcutDraft(helper, text) {
        state := {text: text, profile: helper.profile, section: helper.section, group: helper.meta.group}
        if helper.meta.HasOwnProp("groupKind")
            state.groupKind := helper.meta.groupKind, state.groupName := helper.meta.groupName
        this.shortcutDrafts[helper.rawKey] := state
    }
    ShortcutProblem(helper, message) {
        this.StoreShortcutDraft(helper, helper.field.Value)
        state := this.shortcutDrafts[helper.rawKey]
        this.shortcutErrors[helper.rawKey] := {profile: state.profile, section: state.section, message: message}
        if helper.meta.HasOwnProp("binding")
            this.PaintError(helper.meta.binding, message)
        helper.field.ToolTip := message
        this.UpdateStatus(message)
    }
    InlineShortcutChanged(helper, *) {
        if this.closed || !helper.rawMode || helper.updating
            return
        this.CommitInlineShortcut(helper)
    }
    CommitInlineShortcut(helper) {
        if this.closed || helper.profile != this.model.profile || !DllCall("IsWindow", "Ptr", helper.field.Hwnd)
            return true
        if !helper.rawMode
            return true
        value := Trim(helper.field.Value)
        helper.updating := true
        try {
            SettingsValue.Hotkey(value)
            helper.complete.Call(value)
            helper.key := value
            this.StoreShortcutDraft(helper, value)
            if this.shortcutErrors.Has(helper.rawKey)
                this.shortcutErrors.Delete(helper.rawKey)
            if helper.meta.HasOwnProp("binding")
                this.PaintError(helper.meta.binding, "")
            this.UpdateStatus()
            return true
        } catch as err {
            this.ShortcutProblem(helper, err.Message)
            return false
        } finally helper.updating := false
    }
    CommitInlineShortcuts(gui?) {
        valid := true
        for helper in this.recorders
            if (!IsSet(gui) || helper.gui = gui) && helper.rawMode
                valid := this.CommitInlineShortcut(helper) && valid
        return valid
    }
    AdvancedShortcut(helper, *) {
        if !helper.field.Enabled
            return
        this.CancelHotkeyRecording()
        if helper.rawMode {
            if !this.CommitInlineShortcut(helper)
                return
            helper.rawMode := false
            helper.updating := true
            helper.field.Opt("+ReadOnly")
            helper.field.Value := SettingsValue.HotkeyText(helper.key)
            helper.advanced.Text := "Advanced"
            helper.updating := false
            this.shortcutDrafts.Delete(helper.rawKey)
        } else {
            helper.rawMode := true
            helper.updating := true
            helper.field.Opt("-ReadOnly")
            helper.field.Value := helper.key
            helper.advanced.Text := "Done"
            helper.updating := false
            this.StoreShortcutDraft(helper, helper.key)
            helper.field.Focus()
        }
    }

    ConflictNotice() {
        this.conflictLabel := this.Note("", "9A5700", 42)
        this.Action("Review shortcut conflicts", (*) => this.ShowShortcutConflicts())
        this.RefreshConflictNotice()
    }

    RefreshConflictNotice() {
        try conflicts := HotkeyBindings.Conflicts(HotkeyBindings.Collect(this))
        catch
            return
        this.shortcutConflicts := conflicts
        if this.HasOwnProp("conflictLabel") && DllCall("IsWindow", "Ptr", this.conflictLabel.Hwnd)
            this.conflictLabel.Value := conflicts.Count ? conflicts.Count " shortcut assignments overlap. Review conflicts; intentional overlaps remain allowed."
                : "No overlapping shortcuts in the selected profile. Other profile activation shortcuts are included."
    }

    ShowShortcutConflicts() {
        this.RefreshConflictNotice()
        groups := HotkeyBindings.ConflictGroups(HotkeyBindings.Collect(this))
        dialog := this.ChildDialog("Shortcut conflicts")
        dialog.conflictGroups := groups
        dialog.AddButton("x20 y16 w100 h28", "Close").OnEvent("Click", (*) => this.CloseChild(dialog))
        dialog.AddText("x136 y18 w580 h26", "Conflicting assignments are grouped by shortcut. Intentional overlaps remain allowed.")
        y := 58
        if !groups.Length
            dialog.AddText("x20 y58 w680 h32", "No overlapping shortcuts were found.")
        for group in groups {
            title := ""
            for key in group.keys
                title .= (title = "" ? "" : " / ") key
            dialog.SetFont("s11 Bold")
            heading := dialog.AddText("x20 y" y " w680 h26", title)
            height := Max(26, SettingsLayout.TextHeight(heading, title, 680))
            heading.Move(,,, height)
            y += height + 6
            dialog.SetFont("s10 Norm")
            for entry in group.entries {
                text := entry.label " (" SettingsValue.HotkeyText(entry.key) ")"
                control := dialog.AddText("x32 y" y " w668 h24", text)
                height := Max(24, SettingsLayout.TextHeight(control, text, 668))
                control.Move(,,, height)
                y += height + 2
            }
            y += 18
        }
        dialog.Show("w740 h540")
    }

    BorderDiagram() {
        this.Subheading("Border layout preview")
        mode := this.model.Get("EventActiveBorderMode")
        ; Saved colors can use #hex, 0xhex or RGB notation. GUI options need bare hex.
        active := "E36A0D"
        try active := SettingsValue.Color(this.model.Get("ClientHighligtColor"))
        event := "E9A23B"
        outer := mode = 1 ? active : event
        inner := mode = 1 ? event : active
        if mode = 3
            outer := active, inner := active
        this.MainFrame.AddText("x20 y" this.y " w210 h76 ThemePreserve Background" outer)
        this.MainFrame.AddText("x28 y" (this.y + 8) " w194 h60 ThemePreserve Background" inner)
        this.MainFrame.AddText("x36 y" (this.y + 16) " w178 h44 Center +0x200 BackgroundF7F7F7", "Client preview")
        this.y += 88
        this.Note(mode = 1 ? "Outer: active client. Inner: event alert. Orange illustrates an event color."
            : mode = 2 ? "Outer: event alert. Inner: active client. Orange illustrates an event color."
            : "Combined: the active state and event alert share the border.", "555555", 34)
    }

    MonitoringStatus() {
        if !this.model.Get("gameLogsMonitoringEnabled")
            return "Game-log monitoring is off. You can prepare the character list before enabling it."
        return this.model.Get("monitorOnlySelectedChars")
            ? "Only the listed characters are monitored. These switches also appear under Log sources."
            : "All detected characters are monitored. Enable Monitor only selected characters to use this list."
    }

    ReshufflePreview() {
        this.StopTestThumbnail()
        this.ToggleTestThumbnail()
        this.UpdateStatus("Preview labels reshuffled. Applying restarts the app and assigns a fresh runtime set.")
    }

    RecoverRuntimeThumbnails() {
        count := this.owner.RecoverOffscreenThumbnails()
        this.UpdateStatus(count ? "Recovered " count " off-screen thumbnail(s)." : "No off-screen thumbnails need recovery.")
    }

    StageSizeReset() {
        width := this.model.Get("ThumbnailStartLocationwidth"), height := this.model.Get("ThumbnailStartLocationheight")
        this.model.ResetThumbnailSizes(this.owner)
        this.Changed()
        this.UpdateStatus("Thumbnail sizes will reset to " width " × " height " px after Apply & restart.")
    }

    SourceStatus() {
        this.Subheading("Detected sources")
        game := this.ResolveGameLogsDirectory(), chat := this.ResolveChatLogsDirectory()
        this.Note("Game logs: " game "`n" (DirExist(game) ? "Folder found." : "Folder not found. Enable EVE game logging or choose a folder."), "555555", 60)
        this.Note("Chat logs: " chat "`n" (DirExist(chat) ? "Folder found." : "Folder not found. Enable EVE chat logging or choose a folder."), "555555", 60)
        count := 0
        for hwnd in WinGetList("ahk_exe exefile.exe")
            count++
        this.Note("Open EVE client windows: " count ". Log readers attach after Apply & restart."
            "`nThis tool reads local log files and is not endorsed by CCP.", "555555", 42)
        this.Action("Refresh detected sources", (*) => this.Render())
    }

    StageClientPlacement() {
        positions := this.model.Profiles[this.model.profile]["Client Possitions"]
        count := 0, minimized := 0
        for hwnd in WinGetList("ahk_exe exefile.exe") {
            try {
                title := WinGetTitle(hwnd)
                if title = "EVE"
                    continue
                state := WinGetMinMax(hwnd)
                if state = -1 {
                    minimized++
                    continue
                }
                WinGetPos(&x, &y, &width, &height, hwnd)
                positions[title] := Map("x", x, "y", y, "width", width, "height", height, "IsMaximized", state = 1 ? 1 : 0)
                count++
            }
        }
        this.Changed()
        this.UpdateStatus(count " window placements captured to the draft. " minimized " minimized clients were left unchanged.")
    }

    GroupEditor(kind) {
        groups := this.model.Groups(kind)
        sourceGroup := kind = "characters" ? "Character groups" : "Non-EVE Applications"
        inherited := this.model.Inherited(sourceGroup)
        if !groups.Count {
            this.Action("Create your first group...", (*) => this.AskName("Create group", "", ObjBindMethod(this, "CreateGroup", kind))).Enabled := !inherited
            this.Note("No groups yet. Create a group, then add members and record its switching shortcuts.", "555555", 42)
            return
        }
        names := []
        for name in groups
            names.Push(name)
        selectionKey := this.model.profile "|" kind
        selected := this.groupSelection.Get(selectionKey, names[1])
        if !groups.Has(selected)
            selected := names[1]
        this.groupSelection[selectionKey] := selected
        this.MainFrame.AddText("x20 y" this.y " w400 h22", "Group to edit")
        ddl := this.MainFrame.AddDropDownList("x20 y" (this.y + 26) " w" Min(500, this.contentWidth - 40), names)
        ddl.Choose(selected)
        ddl.OnEvent("Change", ObjBindMethod(this, "SelectGroup", kind))
        this.y += 64
        buttons := this.ButtonRow([
            {label: "New group...", width: 112, callback: (*) => this.AskName("Create group", "", ObjBindMethod(this, "CreateGroup", kind))},
            {label: "Rename...", width: 100, callback: (*) => this.AskName("Rename group", selected, ObjBindMethod(this, "RenameGroup", kind, selected))},
            {label: "Remove group", width: 120, callback: (*) => this.RemoveGroup(kind, selected)}])
        for button in buttons
            button.Enabled := !inherited
        for pair in kind = "characters"
            ? [["ForwardsHotkey", "Next character"], ["BackwardsHotkey", "Previous character"], ["FirstCharHotkey", "First available character"]]
            : [["fkey", "Next application"], ["bkey", "Previous application"]] {
            this.GroupShortcut(kind, selected, pair[1], pair[2], inherited)
        }
        this.Subheading(kind = "characters" ? "Character switching order" : "Application switching order")
        this.AddRowTable(kind = "characters" ? "group-members" : "app-group-members", selected, true)
        this.ConflictNotice()
    }

    GroupShortcut(kind, name, key, label, inherited) {
        raw := this.model.Groups(kind)[name][key]
        width := this.contentWidth - 40
        heading := this.MainFrame.AddText("x20 y" this.y " w" width " h22", label)
        heading.Enabled := !inherited
        this.theme.Font(heading, "Bold " (inherited ? "c707070" : "c202020"))
        this.y += 26
        field := this.MainFrame.AddEdit("x20 y" this.y " w" Max(120, width - 262) " h26 ReadOnly", SettingsValue.HotkeyText(raw))
        helper := this.AddRecorder(this.MainFrame, field, this.y, width, ObjBindMethod(this, "SetGroupShortcut", kind, name, key), raw,
            {key: "group|" kind "|" name "|" key, group: kind = "characters" ? "Character groups" : "Non-EVE Applications", groupKind: kind, groupName: name})
        field.ToolTip := "Switch within this group. Use Record or Advanced to set its shortcut."
        field.SettingsHelp := field.ToolTip
        for control in [field, helper.button, helper.clear, helper.advanced]
            control.Enabled := !inherited
        this.y += 42 + helper.extraHeight
    }

    SetGroupShortcut(kind, name, key, value) {
        this.model.SetGroupShortcut(kind, name, key, value)
        this.Changed()
        this.RefreshConflictNotice()
    }

    SelectGroup(kind, control, *) {
        this.groupSelection[this.model.profile "|" kind] := control.Text
        this.Render()
    }

    CreateGroup(kind, name, *) {
        created := this.model.AddGroup(kind, name)
        this.groupSelection[this.model.profile "|" kind] := created
        this.Changed()
        this.Render()
    }

    RenameGroup(kind, old, name, *) {
        renamed := this.model.RenameGroup(kind, old, name)
        this.groupSelection[this.model.profile "|" kind] := renamed
        this.Changed()
        this.Render()
    }

    RemoveGroup(kind, name) {
        this.model.DeleteGroup(kind, name)
        this.Changed()
        this.Render()
        this.UpdateStatus("Group removed from the draft. Undo restores it and its assignments.")
    }

    ProfileManager() {
        running := this.owner.LastUsedProfile
        staged := this.model.config["LastUsedProfile"]
        this.Note("Running profile: " running "`nProfile to activate on Apply: " staged, "526778", 44)
        this.Action("Create profile...", (*) => this.AskName("Create profile", "", ObjBindMethod(this, "CreateProfile"), true))
        this.Action("Rename profile...", (*) => this.AskName("Rename profile", this.model.profile, ObjBindMethod(this, "RenameProfile"))).Enabled := this.model.profile != "Default"
        this.Action("Remove profile (Undo available)", (*) => this.RemoveProfile()).Enabled := this.model.profile != "Default"
        this.Action("Activate this profile on Apply", (*) => this.StageProfileActivation()).Enabled := staged != this.model.profile
        if this.model.profile = "Default"
            this.Note("Default is built in and cannot be renamed or removed. Other profiles can inherit its sections.", "555555", 40)
        this.ConflictNotice()
    }

    RefreshProfileSelector() {
        this.profileSelector.Delete()
        this.profileSelector.Add(this.ProfileNames())
        this.profileSelector.Choose(this.SelectedProfile)
    }

    CreateProfile(name, copyCurrent := true, inheritance?) {
        this.FlushTableEdits()
        if IsSet(inheritance)
            this.model.AddProfile(name, copyCurrent, inheritance)
        else
            this.model.AddProfile(name, copyCurrent)
        this.RefreshProfileSelector()
        this.Changed()
        this.Render()
    }

    RenameProfile(name, *) {
        this.FlushTableEdits()
        previous := this.model.profile
        this.model.RenameProfile(name)
        this.RekeyProfileDrafts(previous, this.model.profile)
        this.RefreshProfileSelector()
        this.Changed()
        this.Render()
    }

    RemoveProfile() {
        this.FlushTableEdits()
        previous := this.model.profile
        this.model.DeleteProfile()
        this.model.undo[-1].editorDrafts := this.RekeyProfileDrafts(previous)
        this.RefreshProfileSelector()
        this.Changed()
        this.Render()
        this.UpdateStatus("Profile removed from the draft. Undo restores it. Running settings are unchanged until Apply.")
    }

    StageProfileActivation() {
        this.model.ActivateProfile()
        this.Changed()
        this.Render()
    }

    AskName(title, value, callback, copyOptions := false) {
        if copyOptions
            return this.AskProfileCreation(title, callback)
        dialog := this.ChildDialog(title)
        dialog.AddText("x20 y16 w450 h20", "Name")
        name := dialog.AddEdit("x20 y42 w450 h26", value)
        dialog.Help(name, "Choose a unique name. Existing profiles or groups cannot be overwritten.")
        error := dialog.AddText("x20 y150 w450 h44 cB42318", "")
        dialog.AddButton("x250 y204 w104 h30 Default", "Save to draft").OnEvent("Click", Save)
        dialog.AddButton("x366 y204 w104 h30", "Cancel").OnEvent("Click", (*) => this.CloseChild(dialog))
        dialog.Show("w490 h254")
        Save(*) {
            try {
                callback.Call(name.Value, true)
                this.CloseChild(dialog)
            } catch as err
                dialog.SetError(error, err.Message)
        }
    }

    AskProfileCreation(title, callback) {
        dialog := this.ChildDialog(title)
        dialog.AddText("x20 y16 w600 h20", "Profile name")
        name := dialog.AddEdit("x20 y42 w600 h26")
        dialog.Help(name, "Choose a unique profile name. Creating a profile stages it in this draft.")
        dialog.AddText("x20 y80 w600 h20", "Starting values")
        choice := dialog.AddDropDownList("x20 y106 w600 Choose1", ["Copy this profile's effective settings", "Start with application defaults"])
        dialog.Help(choice, "Starting values for custom subsections. Inherited subsections use Default instead.")
        dialog.AddText("x20 y146 w600 h20", "Sections to inherit from Default")
        flags := this.model.CreationInheritance()
        presets := []
        for index, preset in [["Use inheritance defaults", "template"], ["Disable all", "none"], ["Enable all", "all"]] {
            button := dialog.AddButton("x" (20 + (index - 1) * 202) " y174 w196 h28", preset[1])
            button.OnEvent("Click", ApplyPreset.Bind(preset[2]))
            dialog.Help(button, preset[2] = "template" ? "Load the new-profile inheritance template configured in Default / Section inheritance."
                : preset[2] = "none" ? "Give every subsection its own values in this profile." : "Make every subsection use the Default profile's settings.")
            presets.Push(button)
        }
        list := dialog.AddListView("x20 y214 w600 h300 NoSortHdr", ["Settings subsection", "Inherit from Default"])
        grid := SettingsGrid(this, "new-profile-inheritance", list, [{label: "Settings subsection", kind: "readonly"},
            {label: "Inherit", kind: "bool"}], ReadRows(), SaveRows, {append: false, parent: dialog.content, scroller: dialog.scroll, shared: true})
        list.ModifyCol(1, 420), list.ModifyCol(2, 160)
        error := dialog.AddText("x20 y526 w600 h40 cB42318", "")
        create := dialog.AddButton("x380 y584 w116 h30 Default", "Create profile")
        create.OnEvent("Click", Save)
        dialog.AddButton("x504 y584 w116 h30", "Cancel").OnEvent("Click", (*) => this.CloseChild(dialog))
        dialog.profileCreation := {name: name, choice: choice, grid: grid, presets: presets, create: create}
        dialog.Show("w640 h634")
        ReadRows() {
            rows := []
            for group in SettingsModel.Groups
                rows.Push([SettingsWorkbench.GroupLabel(group), flags[group]])
            return rows
        }
        SaveRows(rows) {
            for index, row in rows
                flags[SettingsModel.Groups[index]] := !!row[2]
        }
        ApplyPreset(mode, *) {
            grid.CommitEditor()
            flags := mode = "template" ? this.model.CreationInheritance() : Map()
            if mode != "template"
                for group in SettingsModel.Groups
                    flags[group] := mode = "all"
            grid.Reset(ReadRows())
        }
        Save(*) {
            try {
                grid.CommitEditor()
                if this.tableErrors.Has(grid.key)
                    throw ValueError(this.tableErrors[grid.key].message)
                callback.Call(name.Value, choice.Value = 1, flags.Clone())
                this.CloseChild(dialog)
            } catch as err
                dialog.SetError(error, err.Message)
        }
    }

    InheritanceEditor() {
        template := this.model.profile = "Default"
        this.Note(template ? "Choose which subsections new profiles inherit when you use inheritance defaults. Existing profiles keep their source choices."
            : "Choose sections to inherit from Default. Bulk presets can be adjusted one section at a time.")
        presets := []
        if !template
            presets.Push({label: "Use inheritance defaults", width: 190, callback: ApplyPreset.Bind("template")})
        presets.Push({label: "Disable all", width: 190, callback: ApplyPreset.Bind("none")},
            {label: "Enable all", width: 190, callback: ApplyPreset.Bind("all")})
        this.ButtonRow(presets)
        list := this.MainFrame.AddListView("x20 y" this.y " w" (this.contentWidth - 40) " h330 NoSortHdr", ["Settings subsection", template ? "New profiles inherit" : "Source"])
        groups := SettingsModel.Groups
        rows := ReadRows()
        sourceColumn := template ? {label: "Inherit", kind: "bool"}
            : {label: "Source", kind: "choice", choices: [{label: "Default", value: "Default"}, {label: this.model.profile, value: this.model.profile}]}
        grid := SettingsGrid(this, "inheritance", list, [{label: "Settings subsection", kind: "readonly"}, sourceColumn],
            rows, Save, {append: false, read: ReadRows})
        list.ModifyCol(1, Round((this.contentWidth - 62) * 0.68))
        list.ModifyCol(2, Round((this.contentWidth - 62) * 0.32))
        this.y += 338
        this.Note(template ? "Yes means new profiles inherit that subsection from Default. Change individual rows or use a bulk preset."
            : "Click a Source cell to change it. Select several source cells and type 1 for Default or 2 for this profile. Ctrl+C/V also works.")
        ReadRows() {
            rows := []
            defaults := this.model.CreationInheritance()
            for group in groups
                rows.Push([SettingsWorkbench.GroupLabel(group), template ? defaults[group] : this.model.Source(group)])
            return rows
        }
        Save(rows) {
            if template {
                flags := Map()
                for index, row in rows
                    flags[groups[index]] := !!row[2]
                this.model.SetCreationInheritance(flags)
                return
            }
            for index, row in rows {
                inherit := row[2] = "Default"
                if inherit != this.model.Inherited(groups[index])
                    this.SetSectionInheritance(groups[index], inherit)
            }
        }
        ApplyPreset(mode, *) {
            grid.CommitEditor()
            defaults := this.model.CreationInheritance(), changes := []
            for index, group in groups {
                inherit := mode = "all" || mode = "template" && defaults[group]
                changes.Push([index, 2, template ? (inherit ? 1 : 0) : inherit ? "Default" : this.model.profile])
            }
            grid.SetCells(changes)
        }
    }

    ChildDialog(title) {
        this.FlushTableEdits()
        this.CancelHotkeyRecording()
        this.ClearSettingsTooltip()
        if this.HasOwnProp("activeDialog")
            throw Error("Finish the current editor first.")
        dialog := SettingsChildDialog(this, title)
        this.activeDialog := dialog
        this.S_Gui.Opt("+Disabled")
        return dialog
    }

    ButtonRow(actions) {
        buttons := []
        x := 20, y := this.y
        commonWidth := 0
        for action in actions
            commonWidth := Max(commonWidth, action.width)
        for index, action in actions {
            width := Min(commonWidth, this.contentWidth - 40)
            required := width
            if action.HasOwnProp("keepWithNext") && index < actions.Length
                required += 8 + width
            if x > 20 && x + required > this.contentWidth - 20
                x := 20, y += 38
            button := this.MainFrame.AddButton("x" x " y" y " w" width " h28", action.label)
            button.OnEvent("Click", action.callback)
            buttons.Push(button)
            x += width + 8
        }
        this.y := y + 40
        return buttons
    }

    CloseChild(dialog) {
        index := this.grids.Length
        while index > 0 {
            grid := this.grids[index]
            if grid.parent.Hwnd = dialog.content.Hwnd {
                grid.Dispose()
                this.grids.RemoveAt(index)
                if this.tableStates.Has(grid.key)
                    this.tableStates.Delete(grid.key)
                if this.tableErrors.Has(grid.key)
                    this.tableErrors.Delete(grid.key)
            }
            index--
        }
        this.CancelHotkeyRecording()
        this.ClearSettingsTooltip()
        index := this.recorders.Length
        while index > 0 {
            if this.recorders[index].gui = dialog {
                helper := this.recorders.RemoveAt(index)
                helper.rawMode := false
                if this.shortcutDrafts.Has(helper.rawKey)
                    this.shortcutDrafts.Delete(helper.rawKey)
                if this.shortcutErrors.Has(helper.rawKey)
                    this.shortcutErrors.Delete(helper.rawKey)
            }
            index--
        }
        if this.HasOwnProp("settingsSoundPlayer")
            this.settingsSoundPlayer.Stop()
        this.S_Gui.Opt("-Disabled")
        dialog.Destroy()
        this.theme.Prune()
        if this.HasOwnProp("activeDialog")
            this.DeleteProp("activeDialog")
        if this.HasOwnProp("renderAfterChild") && this.renderAfterChild {
            this.renderAfterChild := false
            this.Render()
        }
        this.UpdateStatus()
    }

    PickRunningWindows(applications, callback) {
        windows := this.RunningWindows(applications)
        dialog := this.ChildDialog(applications ? "Choose running applications" : "Choose running EVE characters")
        dialog.AddText("x20 y16 w660 h42", "Select one or more windows, then add the selected rows. Existing assignments are kept.")
        list := dialog.AddListView("x20 y64 w660 h300 NoSortHdr", applications ? ["Executable", "Window title"] : ["Character", "Window title"])
        rows := []
        for item in windows
            rows.Push([applications ? item.exe : item.character, item.title])
        grid := SettingsGrid(this, "running-window-picker", list,
            [{label: applications ? "Executable" : "Character", kind: "readonly"}, {label: "Window title", kind: "readonly"}],
            rows, (*) => 0, {append: false, readonly: true, local: true, parent: dialog.content, scroller: dialog.scroll})
        list.ModifyCol(1, 240), list.ModifyCol(2, 400)
        status := dialog.AddText("x20 y376 w660 h42", windows.Length ? "Shift/Ctrl+click selects windows; Ctrl+A selects all. Ctrl+C copies selected cells." : "No matching windows found. Type into an empty table row to enter a target manually.")
        add := dialog.AddButton("x412 y430 w144 h30 Default", "Add selected")
        add.Enabled := windows.Length > 0
        add.OnEvent("Click", AddSelected)
        dialog.AddButton("x568 y430 w112 h30", "Cancel").OnEvent("Click", (*) => this.CloseChild(dialog))
        dialog.Show("w700 h480")
        AddSelected(*) {
            selected := [], indexes := Map()
            for key, cell in grid.selected
                indexes[cell[1]] := true
            for index in indexes
                selected.Push(windows[index])
            if !selected.Length {
                status.Value := "Select at least one window."
                return
            }
            try {
                callback.Call(selected)
                this.CloseChild(dialog)
            } catch as err
                status.Value := err.Message
        }
    }

    RunningWindows(applications) {
        result := []
        ownPid := DllCall("GetCurrentProcessId")
        for hwnd in WinGetList(applications ? "" : "ahk_exe exefile.exe") {
            try {
                title := WinGetTitle(hwnd)
                if title = "" || WinGetPID(hwnd) = ownPid
                    continue
                exe := WinGetProcessName(hwnd)
                character := exe = "exefile.exe" ? this.CleanTitle(title) : ""
                if !applications && character = ""
                    continue
                result.Push({exe: exe, title: title, character: character})
            }
        }
        return result
    }

    TrayCommands() {
        commands := this.model.Group("Tray Menu Settings")["TrayMenuShortcuts"]
        for name, visible in commands {
            checkbox := this.MainFrame.AddCheckbox("x20 y" this.y " w" (this.contentWidth - 40) " h28 Checked" (!!visible), name)
            checkbox.Enabled := !this.model.Inherited("Tray Menu Settings")
            checkbox.SettingsHelp := "Show the '" name "' command in the tray menu. This does not activate the feature."
            checkbox.ToolTip := checkbox.SettingsHelp
            checkbox.OnEvent("Click", ObjBindMethod(this, "SetTrayCommand", name))
            this.y += 36
        }
    }

    SetTrayCommand(name, control, *) {
        this.model.Group("Tray Menu Settings")["TrayMenuShortcuts"][name] := control.Value
        this.Changed()
    }

    Diagnostics() {
        this.Subheading("Immediate actions")
        this.Action("Show application log", (*) => ProgramLog.ShowOwned(this.S_Gui.Hwnd))
        this.Action("Export diagnostic snapshot...", (*) => ProgramLog.ExportDialog())
        this.runtimePreviewButton := this.Action(this.owner.livePreviewsPaused ? "Resume live previews now" : "Pause live previews now", (*) => this.ToggleRuntimePreviews())
        this.runtimePreviewStatus := this.Note(this.owner.LivePreviewStatus(), "526778", 40)
        this.Note("Logs are written under %LOCALAPPDATA%\EVE-X-Preview\Logs. Debugging changes above remain staged until Apply.", "555555", 44)
    }

    ToggleRuntimePreviews() {
        this.owner.ToggleLivePreviews()
        this.Render()
    }

    AboutUpdates() {
        this.Subheading("EVE-X-Preview")
        this.installedVersion := this.VersionRow("Installed version", "v" this.programVersion)
        this.stableVersion := this.VersionRow("Latest stable", "Not checked")
        this.preVersion := this.VersionRow("Latest pre-release", "Not checked")
        this.Note("Created by gonzo83. Maintained by khivus.`nContributions by mrmjstc and CJKondur.", "555555", 42)
        this.versionStatus := this.Note(this.updateChecker.status, "526778", 42)
        this.Action("Check for updates", (*) => this.StartUpdateCheck())
        this.Action("Download latest settings editor", (*) => this.OpenSettingsDownload())
        this.releaseInstall := this.Action("Install stable version...", (*) => this.InstallVersion(false), 360)
        this.preInstall := this.Action("Install pre-release version...", (*) => this.InstallVersion(true), 360)
        this.Action("Read included changelog", (*) => ChangelogWindow.Show(false, this.S_Gui.Hwnd))
        this.Action("Open documentation", (*) => Run("https://github.com/khivus/EVE-X-Preview/blob/dev/README.MD"))
        this.Action("Report a problem on GitHub", (*) => Run("https://github.com/khivus/EVE-X-Preview/issues/new"))
        this.UpdateVersionControls()
    }
    VersionRow(label, value) {
        this.MainFrame.AddText("x20 y" this.y " w170 h24", label)
        control := this.MainFrame.AddText("x198 y" this.y " w" Max(100, this.contentWidth - 218) " h24", value)
        this.y += 30
        return control
    }

    OpenSettingsDownload() {
        url := this.updateChecker.settingsUrl
        Run(url != "" ? url : "https://github.com/khivus/EVE-X-Preview/releases")
        this.UpdateStatus(url != "" ? "Settings editor download opened in your browser." : "Opened releases. Look for EVE-X-Preview-Settings.exe; an asset may not be published yet.")
    }

    StartUpdateCheck(force := true) {
        this.updateChecker.Start(force)
        SetTimer(this.updateTimer, 200)
        this.UpdateVersionControls()
    }

    UpdateVersionControls() {
        if this.closed
            return
        available := UpdateChecker.IsNewer(this.updateChecker.releaseTag, this.programVersion)
        wasVisible := this.updateNotice.Visible
        this.updateNotice.Visible := available
        this.updateNotice.Text := available ? "Update available: v" this.updateChecker.releaseTag : ""
        this.updateNotice.ToolTip := "A newer stable version is available. Open About to review and install it."
        if available != wasVisible
            this.Layout()
        if this.currentSection = "about" && this.HasOwnProp("versionStatus") {
            checker := this.updateChecker
            this.versionStatus.Value := checker.status
            missing := checker.busy ? "checking..." : checker.status = "Not checked yet." ? "not checked" : "unknown / not found"
            SettingsTheme.Assign(this.stableVersion, "Value", checker.releaseTag != "" ? "v" checker.releaseTag : missing)
            SettingsTheme.Assign(this.preVersion, "Value", checker.preReleaseTag != "" ? "v" checker.preReleaseTag : missing)
            this.releaseInstall.Text := this.InstallLabel(checker.releaseTag, false)
            this.preInstall.Text := this.InstallLabel(checker.preReleaseTag, true)
            this.releaseInstall.Enabled := checker.releaseTag != "" && VerCompare(checker.releaseTag, this.programVersion) != 0 && !checker.busy
            this.preInstall.Enabled := checker.preReleaseTag != "" && VerCompare(checker.preReleaseTag, this.programVersion) != 0 && !checker.busy
        }
        if !this.updateChecker.busy
            SetTimer(this.updateTimer, 0)
    }

    InstallLabel(tag, pre) {
        if tag = ""
            return pre ? "Install pre-release (check first)" : "Install stable version (check first)"
        direction := VerCompare(tag, this.programVersion) < 0 ? "Downgrade to " : "Install "
        return direction (pre ? "pre-release v" : "stable v") tag "..."
    }

    InstallVersion(pre) {
        tag := pre ? this.updateChecker.preReleaseTag : this.updateChecker.releaseTag
        if tag = ""
            return
        if this.HasDraftChanges() {
            this.ShowError("Apply or discard the settings draft before installing a different version.")
            return
        }
        downgrade := VerCompare(tag, this.programVersion) < 0
        prompt := (downgrade ? "Install the older " : "Install ") (pre ? "pre-release " : "stable release ") "v" tag "? EVE-X-Preview will close for the installer."
        if MsgBox(prompt, downgrade ? "Confirm downgrade" : "Install version", "YesNo Icon? Owner" this.S_Gui.Hwnd) != "Yes"
            return
        updater := A_Temp "\EVE-X-Preview-Updater.exe"
        try {
            Download("https://github.com/khivus/EVE-X-Preview/releases/download/v" tag "/EVE-X-Preview-Updater.exe", updater)
            this.owner.First_Start_After_Update := 1
            this.owner.SaveJsonToFile()
            Run '*RunAs "' updater '" "' A_ScriptFullPath '" "' tag '"'
            ExitApp()
        } catch as err
            this.ShowError("The installer could not be started. " err.Message)
    }
}
