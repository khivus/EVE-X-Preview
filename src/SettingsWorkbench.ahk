; A task-oriented editor over an isolated SettingsModel draft. The legacy page
; builders remain available to old integrations; MainGui opens this workbench.
class SettingsWorkbench extends SettingsWorkbenchPages {
    static NumericFieldWidth := 128
    static Active := 0
    static ResumeVariable := "EVE_X_PREVIEW_SETTINGS_RESUME"
    static HelpCondition := (*) => SettingsWorkbench.HelpAllowed()
    static HelpHandler := (*) => SettingsWorkbench.HelpRequested()

    static HelpAllowed() {
        app := this.Active
        return IsObject(app) && !app.closed && (WinActive("ahk_id " app.S_Gui.Hwnd)
            || (app.HasOwnProp("activeDialog") && WinActive("ahk_id " app.activeDialog.Hwnd)))
    }

    static HelpRequested() {
        if IsObject(this.Active)
            this.Active.FocusedHelp()
    }

    __New(owner, show := true) {
        this.owner := owner
        this.model := owner.HasOwnProp("settingsRecoveredModel") ? owner.DeleteProp("settingsRecoveredModel") : SettingsModel(owner._JSON)
        this._JSON := this.model.config
        this.default_JSON := JSON.Load(default_JSON)
        this.shortcutDrafts := Map(), this.shortcutErrors := Map()
        this.pending := Map()
        this.errors := Map()
        this.relationIssues := this.model.Relations()
        this.scrollPositions := Map()
        this.groupSelection := Map()
        this.bindings := []
        this.recorders := []
        this.tableEditors := []
        this.grids := [], this.tableStates := Map(), this.tableErrors := Map()
        this.expandedPages := Map("Thumbnails", true)
        this.closed := false
        this.building := true
        this.currentSection := "appearance"
        this.Save_Settings_Delay_Timer := (*) => this.Changed()
        this.ProfileOverride()
        this.programVersion := UpdateChecker.CurrentVersion()
        this.updateChecker := this.CreateUpdateChecker()
        this.theme := SettingsTheme(this, this.model.Get("SettingsTheme"))
        this.S_Gui := this.theme.Surface(Gui("+Resize +0x02000000", "EVE-X-Preview - Settings"))
        try {
            this.S_Gui.Opt("+OwnDialogs")
            this.S_Gui.SetFont("s10", "Segoe UI")
            this.S_Gui.BackColor := this.theme.background
            this.S_Gui.OnEvent("Close", (*) => this.RequestClose())
            this.S_Gui.OnEvent("Escape", (*) => this.RequestClose())
            this.S_Gui.OnEvent("Size", ObjBindMethod(this, "OnResize"))
            this.Sidebar := this.theme.Surface(Gui("+Parent" this.S_Gui.Hwnd " -Caption -Border +0x02000000"), true)
            this.Sidebar.BackColor := this.theme.sidebar
            this.Sidebar.SetFont("s10", "Segoe UI")
            this.Sidebar.AddText("x16 y16 w208 h20", "Selected profile")
            this.profileSelector := this.Sidebar.AddDropDownList("x16 y40 w208 vWorkbenchProfile", this.ProfileNames())
            this.profileSelector.Choose(this.model.profile)
            this.profileSelector.OnEvent("Change", (*) => this.SelectProfile(this.profileSelector.Text))
            this.manageProfiles := this.Sidebar.AddButton("x16 y72 w208 h28", "Manage profiles")
            this.manageProfiles.OnEvent("Click", (*) => this.Navigate("profiles"))
            this.updateNotice := this.Sidebar.AddButton("x16 y0 w208 h30 Hidden", "")
            this.updateNotice.OnEvent("Click", (*) => this.Navigate("about"))
            this.Sidebar.AddText("x16 y114 w208 h20", "Search settings")
            this.search := this.Sidebar.AddEdit("x16 y138 w176 h26 vSettingsSearch")
            this.clearSearchButton := this.Sidebar.AddButton("x196 y138 w28 h26", "×")
            this.clearSearchButton.ToolTip := "Clear search"
            this.clearSearchButton.OnEvent("Click", (*) => this.ClearSearchEntry())
            this.search.ToolTip := "Search setting names, explanations and categories. Clear this field to show all sections."
            this.search.OnEvent("Change", (*) => this.ScheduleSearch())
            this.nav := this.Sidebar.AddTreeView("x12 y178 w216 h360 +0x20 -HScroll") ; SHOWSELALWAYS
            this.nav.OnEvent("ItemSelect", ObjBindMethod(this, "NavigationSelected"))
            this.nav.OnEvent("ItemExpand", ObjBindMethod(this, "NavigationExpanded"))
            this.testThumbnailButton := this.Sidebar.AddButton("x16 y550 w208 h30", "Test thumbnail")
            this.testThumbnailButton.ToolTip := "Preview this draft's appearance, labels, borders and sample monitoring data."
            this.testThumbnailButton.OnEvent("Click", (*) => this.ToggleTestThumbnail())
            this.helpNote := this.Sidebar.AddText("x16 y592 w208 h36 c555555", "Help is shown with each setting.`nPress F1 for focused-control help.")
            this.status := this.S_Gui.AddText("x252 y690 w370 h42 c444444 +0x80", "")
            this.closeButton := this.S_Gui.AddButton("x542 y695 w80 h30", "Close")
            this.closeButton.OnEvent("Click", (*) => this.RequestClose())
            this.undoButton := this.S_Gui.AddButton("x630 y695 w70 h30", "Undo")
            this.undoButton.OnEvent("Click", (*) => this.UndoOrLocate())
            this.discardButton := this.S_Gui.AddButton("x708 y695 w118 h30", "Discard changes")
            this.discardButton.OnEvent("Click", (*) => this.Discard())
            this.applyButton := this.S_Gui.AddButton("x834 y695 w146 h30 Default", "Apply && restart")
            this.applyButton.OnEvent("Click", (*) => this.Apply())
            this.searchTimer := ObjBindMethod(this, "RunSearch")
            this.resizeTimer := ObjBindMethod(this, "ResizeContent")
            this.updateTimer := ObjBindMethod(this, "UpdateVersionControls")
            this.tooltipCallback := ObjBindMethod(this, "On_WM_MOUSEMOVE")
            OnMessage(0x200, this.tooltipCallback)
            this.sizeMoveCallback := ObjBindMethod(this, "SizeMove")
            OnMessage(0x231, this.sizeMoveCallback)
            OnMessage(0x232, this.sizeMoveCallback)
            this.sizing := false
            SettingsWorkbench.Active := this
            this.helpCondition := SettingsWorkbench.HelpCondition
            HotIf(this.helpCondition)
            try Hotkey("F1", SettingsWorkbench.HelpHandler, "On")
            finally HotIf()
            this.BuildNavigation()
            this.building := false
            this.width := 1000, this.height := 740
            this.Layout()
            this.Render()
            this.UpdateStatus()
            this.theme.Repaint()
            if show {
                this.Show()
                this.StartUpdateCheck(false)
            }
        } catch as err {
            this.Dispose()
            throw err
        }
    }

    HasDraftChanges() => this.model.Dirty || this.pending.Count || this.tableErrors.Count || this.shortcutErrors.Count || this.GridEditing()
    CreateUpdateChecker() => UpdateChecker()
    GridEditing() {
        for grid in this.grids
            if IsObject(grid.editor)
                return true
        return false
    }
    FlushTableEdits() {
        if this.HasOwnProp("flushingGrids") && this.flushingGrids
            return
        this.flushingGrids := true
        try {
            this.CommitInlineShortcuts()
            for grid in this.grids
                grid.CommitEditor()
        } finally this.flushingGrids := false
    }
    ClearTableGroup(group) {
        for key, state in this.shortcutDrafts.Clone()
            if state.profile = this.model.profile && SettingsModel.Canonical(state.group) = SettingsModel.Canonical(group) {
                this.shortcutDrafts.Delete(key)
                if this.shortcutErrors.Has(key)
                    this.shortcutErrors.Delete(key)
                for helper in this.recorders
                    if helper.rawKey = key
                        helper.rawMode := false
            }

        for key, state in this.tableStates.Clone()
            if SettingsModel.Canonical(state.group) = SettingsModel.Canonical(group) && SubStr(key, 1, StrLen(this.model.profile) + 1) = this.model.profile "|" {
                this.tableStates.Delete(key)
                if this.tableErrors.Has(key)
                    this.tableErrors.Delete(key)
            }
    }
    RowsForEdit(kind, group := "") {
        key := this.model.profile "|" kind "|" group
        return this.tableStates.Has(key) && this.tableErrors.Has(key)
            ? SettingsModel.Clone(this.tableStates[key].rows) : this.model.Rows(kind, group)
    }
    RekeyProfileDrafts(previous, name := "") {
        removed := Map()
        prefix := previous "|"
        for collection in ["pending", "errors", "tableStates", "tableErrors", "shortcutDrafts", "shortcutErrors"] {
            removed[collection] := Map()
            for key, value in this.%collection%.Clone()
                if SubStr(key, 1, StrLen(prefix)) = prefix {
                    this.%collection%.Delete(key)
                    removed[collection][key] := value
                    if name != "" {
                        newKey := name SubStr(key, StrLen(previous) + 1)
                        if collection = "tableErrors" || collection = "shortcutErrors" || collection = "shortcutDrafts"
                            value.profile := name
                        this.%collection%[newKey] := value
                    }
                }
        }
        return removed
    }
    LocateTableIssue() {
        for key, issue in this.tableErrors {
            if this.model.profile != issue.profile || this.currentSection != issue.section
                this.ShowIssueProfile(issue.profile, issue.section)
            this.UpdateStatus(issue.message)
            return
        }
    }

    LastUsedProfile {
        get => this.model.profile
        set => this.model.profile := value
    }

    SelectedProfile => this.HasOwnProp("defaultEditContext") ? this.defaultEditContext.profile : this.model.profile

    EditDefaultHere(group, *) {
        group := SettingsModel.Canonical(group)
        if this.HasOwnProp("defaultEditContext") || !this.model.Inherited(group)
            return
        if this.currentSection != group
            this.Navigate(group)
        this.FlushTableEdits()
        this.RememberScroll()
        this.defaultEditContext := {profile: this.model.profile, section: group}
        this.model.SelectProfile("Default")
        this.ProfileOverride()
        this.Render()
    }

    EndDefaultEditing(render := true, *) {
        if !this.HasOwnProp("defaultEditContext")
            return
        this.FlushTableEdits()
        this.RememberScroll()
        context := this.DeleteProp("defaultEditContext")
        this.model.SelectProfile(context.profile)
        this.ProfileOverride()
        if render
            this.Render()
    }

    ShowIssueProfile(profile, section) {
        selected := this.SelectedProfile
        if profile = "Default" && selected != "Default" && this.model.Source(section, selected) = "Default" {
            this.Navigate(section)
            if !this.HasOwnProp("defaultEditContext")
                this.EditDefaultHere(section)
        } else {
            this.SelectProfile(profile)
            this.Navigate(section)
        }
    }

    ProfileNames() {
        names := []
        for name in this.model.Profiles
            names.Push(name)
        return names
    }

    Show() {
        MonitorGetWorkArea(MonitorGetPrimary(), &left, &top, &right, &bottom)
        bounds := SettingsWindowSizing.Fit(left, top, right, bottom, A_ScreenDPI)
        this.S_Gui.Opt("+MinSize" bounds.minWidth "x" bounds.minHeight)
        this.S_Gui.Show("x" bounds.x " y" bounds.y " w" bounds.width " h" bounds.height)
        this.width := bounds.width, this.height := bounds.height
        this.Layout()
    }

    OnResize(gui, state, width, height) {
        if state = -1 || this.closed || this.building
            return
        if this.HasOwnProp("width") && this.width = width && this.height = height
            return
        this.width := width, this.height := height
        this.Layout()
        if !this.sizing
            SetTimer(this.resizeTimer, -120)
    }

    SizeMove(wParam, lParam, message, hwnd) {
        if this.closed || hwnd != this.S_Gui.Hwnd
            return
        this.sizing := message = 0x231
        SetTimer(this.resizeTimer, 0)
        if !this.sizing
            this.ResizeContent()
    }

    ResizeContent() {
        if !this.closed && !this.sizing && (!this.HasOwnProp("renderedContentWidth") || this.renderedContentWidth != this.contentWidth)
            this.Render()
    }

    Layout() {
        this.theme.BeginUpdate(this.S_Gui)
        try {
            this.sidebarW := Min(240, Max(170, Round(this.width * 0.25)))
            this.contentWidth := Min(1000, Max(220, this.width - this.sidebarW - 18))
            compact := this.width < 960
            bodyHeight := Max(180, this.height - (compact ? 104 : 76))
            this.bodyHeight := bodyHeight
            this.Sidebar.Show("NoActivate x0 y0 w" this.sidebarW " h" this.height)
            this.profileSelector.Move(,, this.sidebarW - 32)
            this.manageProfiles.Move(,, this.sidebarW - 32)
            this.search.Move(,, this.sidebarW - 64)
            this.clearSearchButton.Move(this.sidebarW - 44,, 28)
            noticeHeight := this.updateNotice.Visible ? 38 : 0
            this.nav.Move(,, this.sidebarW - 24, Max(28, this.height - (this.height < 440 ? 226 : 286) - noticeHeight))
            this.updateNotice.Move(, this.height - (this.height < 440 ? 74 : 134), this.sidebarW - 32)
            this.testThumbnailButton.Move(, this.height - (this.height < 440 ? 36 : 96), this.sidebarW - 32)
            this.helpNote.Visible := this.height >= 440
            this.helpNote.Move(, this.height - 54, this.sidebarW - 32)
            this.status.Move(this.sidebarW + 20, this.height - (compact ? 100 : 60), compact ? this.width - this.sidebarW - 40 : Max(100, this.width - this.sidebarW - 480), compact ? 38 : 48)
            this.closeButton.Move(this.width - 458, this.height - 52)
            this.undoButton.Move(this.width - 368, this.height - 52)
            this.discardButton.Move(this.width - 290, this.height - 52)
            this.applyButton.Move(this.width - 164, this.height - 52)
            if this.HasOwnProp("settingsScroller")
                this.settingsScroller.Show(Round(this.sidebarW * A_ScreenDPI / 96), 0, this.contentWidth, bodyHeight)
        } finally this.theme.EndUpdate(this.S_Gui)
    }

    BuildNavigation(query := "") {
        this.navBuilding := true
        try {
            this.nav.Delete()
            this.navItems := Map()
            this.navSections := Map()
            this.navRoots := []
            this.navPages := Map()
            sections := SettingsSchema.Search(query)
            for page in SettingsSchema.Pages {
                matching := []
                for section in sections
                    if section.page = page
                        matching.Push(section)
                if !matching.Length
                    continue
                root := this.nav.Add(page, 0, "Bold")
                this.navRoots.Push(root)
                this.navPages[root] := page
                this.navItems[root] := matching[1].id
                expand := query != "" || this.expandedPages.Get(page, false)
                for section in matching {
                    item := this.nav.Add(section.navTitle, root)
                    this.navItems[item] := section.id
                    this.navSections[section.id] := item
                }
                if expand
                    this.nav.Modify(root, "Expand")
            }
            if this.navSections.Has(this.currentSection) {
                item := this.navSections[this.currentSection]
                parent := this.nav.GetParent(item)
                if query != "" || this.expandedPages.Get(this.navPages[parent], false)
                    this.nav.Modify(item, "Select Vis")
                else
                    this.nav.Modify(parent, "Select")
            }
            else if sections.Length {
                this.currentSection := sections[1].id
                this.nav.Modify(this.navSections[this.currentSection], "Select Vis")
            } else
                this.nav.Add("No matching settings")
        } finally this.navBuilding := false
    }

    NavigationExpanded(control, item, expanded) {
        if !this.navBuilding && this.search.Value = "" && this.navPages.Has(item)
            this.expandedPages[this.navPages[item]] := expanded
    }

    NavigationSelected(control, item) {
        if this.navBuilding || !this.navItems.Has(item) || this.navPages.Has(item) || control.GetSelection() != item
            return
        this.Navigate(this.navItems[item])
    }

    ScheduleSearch() {
        if !this.closed
            SetTimer(this.searchTimer, -180)
    }
    ClearSearchEntry() {
        this.search.Value := ""
        this.RunSearch()
        this.search.Focus()
    }

    RunSearch() {
        SetTimer(this.searchTimer, 0)
        this.EndDefaultEditing(false)
        this.BuildNavigation(this.search.Value)
        this.Render()
    }

    Navigate(id) {
        if this.HasOwnProp("defaultEditContext") && this.defaultEditContext.section != id
            this.EndDefaultEditing(false)
        this.RememberScroll()
        this.CancelHotkeyRecording()
        this.ClearSettingsTooltip()
        this.currentSection := id
        if !this.navSections.Has(id) {
            this.search.Value := ""
            this.BuildNavigation()
        }
        this.navBuilding := true
        try {
            if this.search.Value = "" {
                parent := this.nav.GetParent(this.navSections[id])
                this.nav.Modify(parent, "Expand")
                this.expandedPages[this.navPages[parent]] := true
            }
            this.nav.Modify(this.navSections[id], "Select Vis")
        }
        finally this.navBuilding := false
        this.Render()
    }

    RememberScroll() {
        if this.HasOwnProp("settingsScroller") && this.HasOwnProp("renderedScrollKey")
            this.scrollPositions[this.renderedScrollKey] := this.settingsScroller.offset
    }

    SelectProfile(name) {
        this.EndDefaultEditing(false)
        this.FlushTableEdits()
        this.RememberScroll()
        this.model.SelectProfile(name)
        this.quickGestureError := ""
        this.ProfileOverride()
        this.profileSelector.Choose(name)
        this.Render()
        this.UpdateStatus()
    }

    NewContent() {
        this.FlushTableEdits()
        ; Queued edit notifications can outlive controls during a page rebuild.
        for helper in this.recorders
            helper.rawMode := false
        for name in ["conflictLabel", "monitoringNotice", "reshuffleButton", "mouseActionNotice", "eventTable", "soundLibrary", "runtimePreviewButton", "runtimePreviewStatus", "versionStatus"]
            if this.HasOwnProp(name)
                this.DeleteProp(name)
        for grid in this.grids
            grid.Dispose()
        this.grids := []
        this.ClearSettingsTooltip()
        if this.HasOwnProp("settingsScroller") {
            old := this.settingsScroller
            old.Stop()
            old.content.Destroy()
            old.viewport.Destroy()
            old.header.Destroy()
        }
        this.settingsScroller := SettingsScrollPane(this.S_Gui)
        this.theme.Scroll(this.settingsScroller)
        this.MainFrame := this.theme.Surface(this.settingsScroller.content)
        this.MainFrame.SetFont("s10", "Segoe UI")
        this.MainFrame.Group := Map()
        this.bindings := []
        this.recorders := []
        this.tableEditors := []
        this.pageControls := []
        this.y := 66
    }

    Render() {
        if this.closed || this.building
            return
        if this.HasOwnProp("activeDialog") {
            this.renderAfterChild := true
            return
        }
        this.CancelHotkeyRecording()
        this.building := true
        this.theme.BeginUpdate(this.S_Gui)
        try {
            this.RememberScroll()
            this.NewContent()
            section := SettingsSchema.FindSection(this.currentSection)
            this.MainFrame.SetFont("s13 Bold")
            title := (section.id = "inheritance" && this.model.profile = "Default" ? "New-profile inheritance defaults" : section.title)
            if this.HasOwnProp("defaultEditContext")
                title .= " — editing DEFAULT"
            heading := this.MainFrame.AddText("x20 y14 w" (this.contentWidth - 40) " h26", title)
            this.MainFrame.SetFont("s10 Norm")
            rule := this.MainFrame.AddText("x20 y45 w" (this.contentWidth - 40) " h2 +0x10")
            this.pageControls.Push(heading, rule)
            if this.search.Value != "" && !this.navSections.Count {
                this.Note("No settings match this search. Try a shorter word, or clear the search field.")
            } else {
                this.RenderSources(section)
                if section.help != "" && section.id != "inheritance"
                    this.Note(section.help)
                this.RenderPrerequisites(section)
                if section.id = "quick-group"
                    this.QuickGroupGesture()
                for id in section.fields
                    if section.id != "events" || id != "shootingInterval"
                        this.RenderField(SettingsSchema.Fields[id])
                if section.id != "tray" ; Tray options now use the shared scalar schema.
                    this.RenderSpecial(section.id)
            }
            for control in this.MainFrame
                if control != heading && control != rule
                    this.pageControls.Push(control)
            this.MainFrame.Group[section.title] := this.pageControls
            this.Layout()
            this.settingsScroller.Select(this.pageControls)
            this.renderedScrollKey := JSON.Dump([this.SelectedProfile, this.currentSection, this.HasOwnProp("defaultEditContext") ? "default" : "profile"])
            this.settingsScroller.ScrollTo(this.scrollPositions.Get(this.renderedScrollKey, 0))
            this.renderedProfile := this.model.profile
            this.renderedSection := this.currentSection
            this.renderedContentWidth := this.contentWidth
            this.RefreshDependencies()
            this.RevealSearchMatch()
            this.theme.Prune()
            this.UpdateStatus()
            this.theme.Repaint()
        } finally {
            this.building := false
            this.theme.EndUpdate(this.S_Gui)
        }
    }

    Note(text, color := "555555", height := 42) {
        control := this.MainFrame.AddText("x20 y" this.y " w" (this.contentWidth - 40) " h" height " c" color " +0x80", text)
        height := Max(22, SettingsLayout.TextHeight(control, text, this.contentWidth - 40))
        control.Move(,,, height)
        this.y += height + 8
        return control
    }
    RevealSearchMatch() {
        query := Trim(this.search.Value)
        if query = ""
            return
        for binding in this.bindings
            for control in [binding.label, binding.hint]
                if SettingsSearchText.Matches(control.Text, query).Length {
                    this.settingsScroller.EnsureVisible(control)
                    return
                }
    }

    Subheading(text) {
        this.MainFrame.SetFont("s11 Bold")
        control := this.MainFrame.AddText("x20 y" this.y " w" (this.contentWidth - 40) " h24 +0x80", text)
        this.MainFrame.SetFont("s10 Norm")
        this.y += 30
        return control
    }

    RenderSources(section) {
        if this.HasOwnProp("defaultEditContext") {
            this.Note("Editing DEFAULT profile settings. Changes affect every profile that inherits this subsection. Selected profile: " this.defaultEditContext.profile ".", "9A5700")
            this.Action("Done editing Default", (*) => this.EndDefaultEditing(), 260)
            return
        }
        groups := Map()
        if section.group != ""
            groups[section.group] := true
        if section.id = "quick-group"
            groups["mouse-actions"] := true
        for id in section.fields {
            field := SettingsSchema.Fields[id]
            if !field.global && field.group != ""
                groups[field.group] := true
        }
        for group in groups {
            label := SettingsWorkbench.GroupLabel(group)
            if this.model.profile = "Default" {
                this.Note("Source: Default profile — " label ".", "526778", 22)
                continue
            }
            this.MainFrame.AddText("x20 y" this.y " w" (this.contentWidth - 40) " h20", "Settings source: " label)
            select := this.MainFrame.AddDropDownList("x20 y" (this.y + 24) " w" Min(430, this.contentWidth - 40), ["Use Default profile settings", "Customize for this profile"])
            select.Choose(this.model.Inherited(group) ? 1 : 2)
            select.ToolTip := "Choose the source for " label ". Other subsections keep their own sources. Stored custom values are retained."
            select.OnEvent("Change", ObjBindMethod(this, "ChangeSource", group))
            this.y += 62
            if this.model.Inherited(group)
                this.Action(group = section.id ? "Edit Default here" : "Edit Default " label, ObjBindMethod(this, "EditDefaultHere", group), 320)
        }
    }

    static GroupLabel(group) {
        for section in SettingsSchema.Sections
            if section.id = group
                return section.title
        return Map("Thumbnails Behavior", "thumbnail layout, placement and visibility", "Thumbnails Visuals", "appearance, labels and minimum sizes",
            "Thumbnails Interactions", "mouse actions and Quick Group", "Thumbnail Visibility", "hidden characters", "Custom Colors", "character colors",
            "Hotkeys Settings", "keyboard shortcuts", "Hotkey Groups", "switching behavior and focus retries",
            "Character shortcuts", "character shortcuts", "Character groups", "character groups",
            "Client Settings", "EVE window behavior and character exceptions", "Game Logs Monitoring", "log sources, systems and alert behavior",
            "Focused-client sounds", "focused-client sound muting", "Focused-client DPS", "focused-client DPS visibility",
            "Monitored Events", "event alerts", "DPS Monitoring", "DPS", "Non-EVE Applications", "other applications",
            "Tray Menu Settings", "tray menu commands", "Other", "performance and compatibility")[group]
    }

    SetSectionInheritance(group, inherit) {
        this.FlushTableEdits()
        this.ClearTableGroup(group)
        this.model.SetInheritance(group, inherit)
        if inherit {
            for id, field in SettingsSchema.Fields {
                key := this.model.profile "|" id
                if field.group = group {
                    if this.pending.Has(key)
                        this.pending.Delete(key)
                    if this.errors.Has(key)
                        this.errors.Delete(key)
                }
            }
        }
    }

    ChangeSource(group, control, *) {
        this.SetSectionInheritance(group, control.Value = 1)
        this.Changed()
        this.Render()
    }

    RenderPrerequisites(section) {
        if section.id = "dps" || section.id = "events" || section.id = "alert-behavior" {
            enabled := this.model.Get("gameLogsMonitoringEnabled")
            this.Note(enabled ? "Game-log reading is enabled in this draft. Changes take effect after Apply & restart."
                : "Game-log reading is off. Enable it in Log sources to use alerts and DPS.", enabled ? "526778" : "9A5700", 34)
            if !enabled
                this.Action("Open log sources", (*) => this.Navigate("sources"))
        }
    }

    RenderField(field) {
        value := this.model.Get(field.id)
        key := this.model.profile "|" field.id
        display := this.pending.Get(key, value)
        width := this.contentWidth - 40
        this.y += 4
        binding := {field: field, key: key, inputs: [], y: this.y}
        if field.kind = "bool" {
            control := this.MainFrame.AddCheckbox("x20 y" this.y " w" width " h28 +0x2000 Checked" (!!value), field.label)
            control.OnEvent("Click", ObjBindMethod(this, "BooleanChanged", binding))
            binding.label := control
            control.SetFont("Bold")
            labelHeight := Max(26, SettingsLayout.TextHeight(control, field.label, width - 26))
            control.Move(,,, labelHeight)
            this.y += labelHeight + 4
        } else {
            inline := width >= 620 && field.kind != "hotkey" && field.kind != "slider" && field.kind != "multiline" && field.kind != "directory"
            inputX := inline ? 320 : 20
            inputWidth := width - inputX + 20
            binding.label := this.MainFrame.AddText("x20 y" this.y " w" (inline ? 288 : width) " h26 +0x80", field.label)
            binding.label.SetFont("Bold")
            labelHeight := Max(24, SettingsLayout.TextHeight(binding.label, field.label, inline ? 288 : width))
            binding.label.Move(,,, labelHeight)
            if !inline
                this.y += labelHeight + 4
            switch field.kind {
                case "choice":
                    options := [], selected := 0
                    for index, option in field.choices {
                        options.Push(option.label)
                        if option.value = value
                            selected := index
                    }
                    control := this.MainFrame.AddDropDownList("x" inputX " y" this.y " w" Min(580, inputWidth), options)
                    if selected
                        control.Choose(selected)
                    SendMessage(0x160, Round(Min(580, inputWidth) * A_ScreenDPI / 96), 0, control) ; CB_SETDROPPEDWIDTH
                    control.OnEvent("Change", ObjBindMethod(this, "ChoiceChanged", binding))
                case "hotkey":
                    control := this.MainFrame.AddEdit("x20 y" this.y " w" Max(120, width - 262) " h26 ReadOnly", SettingsValue.HotkeyText(value))
                    helper := this.AddRecorder(this.MainFrame, control, this.y, width, ObjBindMethod(this, "ShortcutSet", binding), value, {key: field.id, group: field.group, binding: binding})
                    binding.helper := helper
                    for button in [helper.button, helper.clear, helper.advanced]
                        binding.inputs.Push(button)
                    this.y += helper.extraHeight
                case "slider":
                    sliderWidth := Min(260, Max(100, width - SettingsWorkbench.NumericFieldWidth - 8))
                    slider := this.MainFrame.AddSlider("x20 y" (this.y - 3) " w" sliderWidth " h32 Range" field.min "-" field.max " AltSubmit TickInterval" Max(1, Round((field.max - field.min) / 10)) " Line5 Page" Max(5, Round((field.max - field.min) / 10)), value)
                    control := this.MainFrame.AddEdit("x" (28 + sliderWidth) " y" this.y " w" SettingsWorkbench.NumericFieldWidth " h26", display)
                    slider.OnEvent("Change", ObjBindMethod(this, "SliderChanged", binding, slider, control))
                    SettingsLayout.SliderTicks(this.MainFrame, slider)
                    binding.slider := slider
                    binding.inputs.Push(slider)
                    control.OnEvent("Change", ObjBindMethod(this, "TextChanged", binding))
                    control.OnEvent("LoseFocus", ObjBindMethod(this, "CommitField", binding))
                case "font":
                    control := this.MainFrame.AddComboBox("x" inputX " y" this.y " w" Min(430, inputWidth) " h260", SettingsFonts.Names())
                    control.Text := display
                    control.OnEvent("Change", ObjBindMethod(this, "TextChanged", binding))
                    control.OnEvent("LoseFocus", ObjBindMethod(this, "CommitField", binding))
                case "multiline":
                    control := this.MainFrame.AddEdit("x20 y" this.y " w" width " h150 Multi -Wrap", display)
                    control.OnEvent("Change", ObjBindMethod(this, "TextChanged", binding))
                    this.y += 124
                default:
                    editWidth := field.kind = "directory" ? Max(100, inputWidth - 108) : field.kind = "color" ? Min(156, Max(72, inputWidth - 126)) : field.kind = "integer" || field.kind = "number" ? Min(SettingsWorkbench.NumericFieldWidth, inputWidth) : Min(240, inputWidth)
                    control := this.MainFrame.AddEdit("x" inputX " y" this.y " w" editWidth " h26", display)
                    control.OnEvent("Change", ObjBindMethod(this, "TextChanged", binding))
                    control.OnEvent("LoseFocus", ObjBindMethod(this, "CommitField", binding))
                    if field.kind = "color" {
                        color := "F0F0F0"
                        try color := SettingsValue.Color(value)
                        swatch := this.MainFrame.AddText("x" (inputX + editWidth + 8) " y" this.y " w26 h26 Border +0x100 ThemePreserve Background" color)
                        swatch.ToolTip := "Preview of the selected color. Use Choose to change it."
                        binding.swatch := swatch
                        picker := this.MainFrame.AddButton("x" (inputX + editWidth + 42) " y" this.y " w84 h26", "Choose...")
                        picker.OnEvent("Click", ObjBindMethod(this, "PickColor", binding, control))
                        binding.inputs.Push(picker)
                    } else if field.kind = "directory" {
                        browse := this.MainFrame.AddButton("x" (width - 78) " y" this.y " w98 h26", "Browse...")
                        browse.OnEvent("Click", ObjBindMethod(this, "PickDirectory", binding, control))
                        binding.inputs.Push(browse)
                    } else if field.kind = "integer" {
                        spin := this.MainFrame.AddUpDown("Range" field.min "-" field.max, value)
                        binding.inputs.Push(spin)
                        control.Value := display
                    }
            }
            this.y += Max(28, inline ? labelHeight : 28)
        }
        binding.control := control
        binding.inputs.Push(control)
        control.SettingsHelp := field.label ": " this.FieldHint(field)
        control.ToolTip := control.SettingsHelp
        for input in binding.inputs
            if !input.HasProp("ToolTip") {
                input.ToolTip := control.SettingsHelp
                input.SettingsHelp := control.SettingsHelp
            }
        binding.label.ToolTip := control.SettingsHelp
        if binding.label.Type = "Text"
            binding.label.Opt("+0x100")
        binding.hint := this.MainFrame.AddText("x20 y" this.y " w" width " h20 c606A76 +0x80", this.FieldHint(field))
        binding.hint.SetFont("s9 Norm")
        hints := [this.FieldHint(field), "Inherited from Default. Choose Edit Default here or Customize above."]
        if field.dependency != ""
            hints.Push("Enable '" SettingsSchema.Fields[field.dependency].label "' to use this setting.")
        hintHeight := 18
        for hint in hints
            hintHeight := Max(hintHeight, SettingsLayout.TextHeight(binding.hint, hint, width))
        binding.hint.Move(,,, hintHeight)
        this.y += hintHeight + 5
        this.MainFrame.AddText("x20 y" this.y " w" width " h1 BackgroundDDE1E7")
        this.y += 5
        this.bindings.Push(binding)
        if this.errors.Has(key)
            this.PaintError(binding, this.errors[key])
    }

    FieldHint(field) {
        range := field.kind = "integer" || field.kind = "number" || field.kind = "slider"
            ? field.min "–" field.max (field.unit = "%" ? "%" : field.unit != "" ? " " field.unit : "") ". " : ""
        return (field.global ? "Shared by all profiles. " : "") range field.help
    }

    BooleanChanged(binding, control, *) {
        this.model.Set(binding.field.id, control.Value)
        this.Changed()
        this.RefreshDependencies()
    }

    ChoiceChanged(binding, control, *) {
        if control.Value
            this.model.Set(binding.field.id, binding.field.choices[control.Value].value)
        this.Changed()
        if binding.field.id = "EventActiveBorderMode"
            this.Render()
    }

    TextChanged(binding, control, *) {
        if this.building || (binding.HasOwnProp("updating") && binding.updating)
            return
        raw := control.Type = "ComboBox" ? control.Text : control.Value
        this.pending[binding.key] := raw
        try {
            value := this.model.Set(binding.field.id, raw)
            if this.errors.Has(binding.key)
                this.errors.Delete(binding.key)
            if binding.HasOwnProp("slider") && (!binding.HasOwnProp("sliderTracking") || !binding.sliderTracking)
                SettingsTheme.Assign(binding.slider, "Value", value)
            if binding.HasOwnProp("swatch")
                this.theme.SetSwatch(binding.swatch, value)
        } catch as err {
            this.errors[binding.key] := err.Message
        }
        this.Changed()
    }

    CommitField(binding, control, *) {
        this.TextChanged(binding, control)
        if this.errors.Has(binding.key)
            return
        binding.updating := true
        try {
            if control.Type = "ComboBox"
                control.Text := this.model.Get(binding.field.id)
            else
                control.Value := this.model.Get(binding.field.id)
            if this.pending.Has(binding.key)
                this.pending.Delete(binding.key)
        } finally binding.updating := false
    }

    SliderChanged(binding, slider, control, eventControl := 0, info := 8) {
        binding.sliderTracking := info = 5
        value := Max(binding.field.min, Min(binding.field.max, Round(slider.Value / 5) * 5))
        if !binding.sliderTracking
            SettingsTheme.Assign(slider, "Value", value)
        SettingsTheme.Assign(control, "Value", value)
        this.TextChanged(binding, control)
    }

    ShortcutSet(binding, value) {
        try {
            this.model.Set(binding.field.id, value)
            display := binding.HasOwnProp("helper") && binding.helper.rawMode ? value : SettingsValue.HotkeyText(value)
            if binding.control.Value != display
                binding.control.Value := display
            this.Changed()
            this.RefreshConflictNotice()
        } catch as err
            this.ShowError(err.Message)
    }

    PickColor(binding, control, *) {
        color := this.ChooseColor(this.model.Get(binding.field.id))
        if color != "" {
            control.Value := color
            this.CommitField(binding, control)
        }
    }
    ChooseColor(value) => ColorPicker.Show(this.S_Gui.Hwnd, value)

    PickDirectory(binding, control, *) {
        directory := DirSelect("*" this.model.Get(binding.field.id), 3, "Choose " binding.field.label)
        if directory != "" {
            control.Value := directory
            this.CommitField(binding, control)
        }
    }

    PaintError(binding, message) {
        if !binding.HasOwnProp("hint")
            return
        this.theme.Font(binding.hint, message != "" ? "cB42318" : "c555555")
        SettingsTheme.Assign(binding.hint, "Value", message != "" ? message : this.FieldHint(binding.field))
    }

    RefreshDependencies() {
        for binding in this.bindings {
            field := binding.field
            inherited := !field.global && this.model.Inherited(field.group)
            parentOff := field.dependency != "" && !this.model.Get(field.dependency)
            for input in binding.inputs
                SettingsTheme.Assign(input, "Enabled", !inherited && !parentOff)
            SettingsTheme.Assign(binding.label, "Enabled", !inherited && !parentOff)
            this.theme.Font(binding.label, "Bold " (inherited || parentOff ? "c707070" : "c202020"))
            message := this.ValidationMessage(binding.key)
            if message = "" {
                reason := inherited ? "Inherited from Default. Choose Edit Default here or Customize above."
                    : parentOff ? "Enable '" SettingsSchema.Fields[field.dependency].label "' to use this setting." : this.FieldHint(field)
                this.theme.Font(binding.hint, inherited || parentOff ? "c765B00" : "c555555")
                SettingsTheme.Assign(binding.hint, "Value", reason)
            } else
                this.PaintError(binding, message)
        }
        if this.HasOwnProp("monitoringNotice")
            SettingsTheme.Assign(this.monitoringNotice, "Value", this.MonitoringStatus())
        if this.currentSection = "labels" && this.HasOwnProp("reshuffleButton")
            SettingsTheme.Assign(this.reshuffleButton, "Enabled", !!this.model.Get("AnonymizeThumbnails"))
        if this.currentSection = "mouse-actions" && this.HasOwnProp("mouseActionNotice")
            SettingsTheme.Assign(this.mouseActionNotice, "Value", this.model.Get("ClickThroughActive")
                ? "Click-through is on. Thumbnail mouse actions become available after you turn it off and apply."
                : this.model.Get("LockThumbnailPositions") ? "Layout locking blocks move and resize gestures. Unlock thumbnails in Layout to use those actions."
                : "Move and resize actions use dragging; the other actions use clicks.")
    }

    Changed() {
        this.theme.SetMode(this.model.Get("SettingsTheme"))
        this.ProfileOverride()
        this.relationIssues := this.model.Relations()
        this.UpdateStatus()
        this.RefreshDependencies()
        SettingsSoundLibrary.RefreshPreviewGain(this)
    }

    ValidationMessage(key) {
        if this.shortcutErrors.Has(key)
            return this.shortcutErrors[key].message
        if this.errors.Has(key)
            return this.errors[key]
        for issue in this.relationIssues
            if key = issue.profile "|" issue.id || (issue.HasOwnProp("relatedId") && key = issue.profile "|" issue.relatedId)
                return issue.message
        return ""
    }

    UpdateStatus(message := "") {
        if !this.HasOwnProp("status") || this.closed
            return
        invalid := this.errors.Count + this.relationIssues.Length + this.tableErrors.Count + this.shortcutErrors.Count
        dirty := this.model.Dirty, editing := this.GridEditing()
        status := message != "" ? message : invalid ? "Fix " invalid " invalid field(s) before applying."
            : dirty || editing ? this.model.ChangeCount " settings changed, apply to save." : "No pending changes. Running profile: " this.owner.LastUsedProfile
        if this.HasOwnProp("defaultEditContext")
            status := "Editing DEFAULT. " status
        this.theme.Font(this.status, invalid ? "cB42318" : "c444444")
        SettingsTheme.Assign(this.status, "Value", status)
        this.status.ToolTip := this.status.Value
        SettingsTheme.Assign(this.applyButton, "Enabled", (dirty || editing) && !invalid)
        SettingsTheme.Assign(this.discardButton, "Enabled", !!(dirty || editing || this.pending.Count || this.tableErrors.Count || this.shortcutErrors.Count))
        SettingsTheme.Assign(this.undoButton, "Text", invalid ? "Fix fields" : "Undo")
        this.undoButton.ToolTip := invalid ? "Go to the first invalid setting, including unfinished edits in another profile." : "Restore the most recently removed row, group, profile or library reference."
        SettingsTheme.Assign(this.undoButton, "Enabled", invalid || this.model.undo.Length > 0)
    }

    ShowError(message) {
        this.UpdateStatus(message)
        MsgBox(message, "Settings", "Icon! Owner" this.S_Gui.Hwnd)
    }

    Undo() {
        this.FlushTableEdits()
        if this.tableErrors.Count {
            this.LocateTableIssue()
            return
        }
        try {
            entry := this.model.undo.Length ? this.model.undo[-1] : 0
            if IsObject(entry) && entry.kind = "grid-batch" {
                this.ShowIssueProfile(entry.profile, entry.section)
                for grid in this.grids
                    if grid.key = entry.key {
                        rows := SettingsModel.Clone(grid.rows)
                        for removed in entry.entries
                            rows.InsertAt(Min(removed.index, rows.Length + 1), SettingsModel.Clone(removed.row))
                        this.model.undo.Pop()
                        grid.SaveCandidate(rows)
                        return
                    }
                throw ValueError("Open the original table before restoring its removed rows.")
            }
            if IsObject(entry) && entry.kind = "group" && entry.profile = "Default" && this.SelectedProfile != "Default"
                this.ShowIssueProfile("Default", entry.groupKind = "characters" ? "character-groups" : "app-groups")
            if this.model.UndoLast() {
                if IsObject(entry) && entry.HasOwnProp("editorDrafts")
                    for collection, values in entry.editorDrafts
                        for key, value in values
                            this.%collection%[key] := value
                this.RefreshProfileSelector()
                this.Changed()
                this.Render()
            }
        } catch as err
            this.ShowError(err.Message)
    }

    UndoOrLocate() {
        if this.errors.Count || this.relationIssues.Length || this.tableErrors.Count || this.shortcutErrors.Count
            this.LocateIssue()
        else
            this.Undo()
    }

    LocateIssue() {
        for key, issue in this.shortcutErrors {
            this.ShowIssueProfile(issue.profile, issue.section)
            if this.shortcutDrafts.Has(key) {
                state := this.shortcutDrafts[key]
                if state.HasOwnProp("groupKind")
                    this.groupSelection[issue.profile "|" state.groupKind] := state.groupName
            }
            this.Navigate(issue.section)
            this.UpdateStatus(issue.message)
            return
        }
        if this.tableErrors.Count {
            this.LocateTableIssue()
            return
        }
        for key, message in this.errors {
            separator := InStr(key, "|", false, -1)
            profile := SubStr(key, 1, separator - 1)
            id := SubStr(key, separator + 1)
            this.ShowIssueProfile(profile, SettingsSchema.Fields[id].section)
            this.UpdateStatus(message)
            return
        }
        if this.relationIssues.Length {
            issue := this.relationIssues[1]
            this.ShowIssueProfile(issue.profile, SettingsSchema.Fields[issue.id].section)
            this.UpdateStatus(issue.message)
        }
    }

    Apply() {
        this.FlushTableEdits()
        if this.shortcutErrors.Count {
            this.LocateIssue()
            return false
        }
        if this.tableErrors.Count {
            this.LocateTableIssue()
            return false
        }
        if this.HasOwnProp("activeDialog") {
            this.UpdateStatus("Save or cancel the open row/event editor before applying the draft.")
            return false
        }
        this.relationIssues := this.model.Relations()
        if this.errors.Count {
            this.LocateIssue()
            return false
        }
        if this.relationIssues.Length {
            issue := this.relationIssues[1]
            this.ShowIssueProfile(issue.profile, SettingsSchema.Fields[issue.id].section)
            this.UpdateStatus(issue.message)
            return false
        }
        if !this.model.Dirty
            return true
        return this.CommitDraft()
    }

    CommitDraft() {
        failure := ""
        previousCritical := A_IsCritical
        Critical(true)
        try {
            SettingsSession.Commit(this.owner, this.model)
        } catch as err {
            failure := err.Message
        } finally Critical(previousCritical)
        if failure != "" {
            this.ShowError("Settings could not be saved. Your draft is still open.`n" failure)
            return false
        }
        this.resumeState := this.CaptureSession()
        this.Dispose()
        this.RestartAfterApply()
        return true
    }

    CaptureSession() {
        this.RememberScroll()
        WinGetPos(&x, &y,,, this.S_Gui.Hwnd)
        return Map("script", A_ScriptFullPath, "profile", this.SelectedProfile, "section", this.currentSection,
            "search", this.search.Value, "scroll", this.scrollPositions.Clone(), "expanded", this.expandedPages.Clone(),
            "x", x, "y", y, "width", this.width, "height", this.height,
            "maximized", WinGetMinMax(this.S_Gui.Hwnd) = 1)
    }

    RestartAfterApply() {
        EnvSet(SettingsWorkbench.ResumeVariable, JSON.Dump(this.resumeState))
        Reload()
    }

    static TakeResume() {
        value := EnvGet(this.ResumeVariable)
        EnvSet(this.ResumeVariable, "") ; One restart only; normal launches stay closed.
        if value = ""
            return 0
        try {
            state := JSON.Load(value)
            if !(state is Map) || state.Get("script", "") != A_ScriptFullPath
                return 0
            for key in ["profile", "section", "search", "scroll", "expanded", "x", "y", "width", "height", "maximized"]
                if !state.Has(key)
                    return 0
            if !(state["scroll"] is Map) || !(state["expanded"] is Map)
                return 0
            for key in ["x", "y", "width", "height"]
                if !IsNumber(state[key])
                    return 0
            SettingsSchema.FindSection(state["section"])
            return state
        }
        return 0
    }

    RestoreSession(state, show := true) {
        this.S_Gui.Show("Hide x" state["x"] " y" state["y"] " w" state["width"] " h" state["height"])
        work := SettingsWindowSizing.WorkArea(this.S_Gui.Hwnd)
        fit := SettingsWindowSizing.Fit(work.left, work.top, work.right, work.bottom, A_ScreenDPI)
        this.width := Max(fit.minWidth, Min(state["width"], Floor((work.right - work.left - 24) * 96 / A_ScreenDPI)))
        this.height := Max(fit.minHeight, Min(state["height"], Floor((work.bottom - work.top - 40) * 96 / A_ScreenDPI)))
        this.S_Gui.Opt("+MinSize" fit.minWidth "x" fit.minHeight)
        this.S_Gui.Show("Hide w" this.width " h" this.height)
        WinGetPos(,, &outerWidth, &outerHeight, this.S_Gui.Hwnd)
        x := Max(work.left, Min(state["x"], work.right - outerWidth))
        y := Max(work.top, Min(state["y"], work.bottom - outerHeight))
        WinMove(x, y,,, this.S_Gui.Hwnd)
        if this.model.Profiles.Has(state["profile"])
            this.model.SelectProfile(state["profile"])
        this.ProfileOverride()
        this.profileSelector.Choose(this.model.profile)
        if this.HasOwnProp("renderedScrollKey")
            this.DeleteProp("renderedScrollKey")
        this.scrollPositions := state["scroll"]
        this.expandedPages := state["expanded"]
        this.currentSection := state["section"]
        this.search.Value := state["search"]
        this.BuildNavigation(this.search.Value)
        this.Layout()
        this.Navigate(this.currentSection)
        if show
            this.S_Gui.Show(state["maximized"] ? "Maximize" : "")
    }

    Discard() {
        if !this.HasDraftChanges()
            return
        if MsgBox("Discard every unapplied change in this settings window?", "Discard settings draft", "YesNo Icon? Owner" this.S_Gui.Hwnd) != "Yes"
            return
        this.DiscardDraft()
    }

    DiscardDraft() {
        if this.HasOwnProp("defaultEditContext")
            this.DeleteProp("defaultEditContext")
        for helper in this.recorders
            helper.rawMode := false
        for grid in this.grids
            grid.Dispose()
        this.grids := [], this.tableStates := Map(), this.tableErrors := Map()
        this.shortcutDrafts := Map(), this.shortcutErrors := Map()
        this.quickGestureError := ""
        this.model := SettingsModel(this.owner._JSON)
        this._JSON := this.model.config
        this.pending := Map(), this.errors := Map()
        this.profileSelector.Delete()
        this.profileSelector.Add(this.ProfileNames())
        this.profileSelector.Choose(this.model.profile)
        this.Changed()
        this.Render()
    }

    RequestClose() {
        this.FlushTableEdits()
        if this.HasOwnProp("activeDialog") {
            this.UpdateStatus("Save or cancel the open settings dialog before closing the workbench.")
            return
        }
        if this.HasDraftChanges() {
            this.CloseDecision()
            return
        }
        this.Dispose()
    }

    CloseDecision() {
        if this.HasOwnProp("closePrompt")
            return
        dialog := this.theme.Surface(Gui("+Owner" this.S_Gui.Hwnd " -MinimizeBox -MaximizeBox", "Unapplied settings"))
        this.closePrompt := dialog
        dialog.SetFont("s10", "Segoe UI")
        dialog.AddText("x20 y18 w490 h48", "Your draft has unapplied changes. Apply them and restart EVE-X-Preview, discard them, or keep editing.")
        dialog.AddButton("x20 y82 w154 h32", "Apply && restart").OnEvent("Click", (*) => this.FinishClose("apply"))
        dialog.AddButton("x184 y82 w154 h32", "Discard draft").OnEvent("Click", (*) => this.FinishClose("discard"))
        dialog.AddButton("x348 y82 w154 h32 Default", "Keep editing").OnEvent("Click", (*) => this.FinishClose("keep"))
        dialog.OnEvent("Close", (*) => this.FinishClose("keep"))
        dialog.OnEvent("Escape", (*) => this.FinishClose("keep"))
        this.S_Gui.Opt("+Disabled")
        SettingsWindowSizing.ShowCentered(dialog, this.S_Gui.Hwnd, 522, 136)
    }

    FinishClose(action) {
        if this.HasOwnProp("closePrompt") {
            this.S_Gui.Opt("-Disabled")
            this.closePrompt.Destroy()
            this.DeleteProp("closePrompt")
        }
        if action = "apply" {
            if this.Apply() && !this.closed
                this.Dispose()
        }
        else if action = "discard"
            this.Dispose()
    }

    FocusedHelp(*) {
        handle := DllCall("GetFocus", "Ptr")
        control := GuiCtrlFromHwnd(handle)
        message := IsObject(control) && control.HasProp("SettingsHelp") ? control.SettingsHelp
            : IsObject(control) && control.HasProp("ToolTip") ? control.ToolTip : "Select a setting to see its description. Help is also shown directly below each setting."
        MsgBox(message, "Setting help", "Owner" this.S_Gui.Hwnd)
    }
    ShowSettingsTooltip(hwnd) {
        MouseGetPos(,,, &hovered, 2)
        if hovered != hwnd || !DllCall("IsWindowVisible", "Ptr", hwnd)
            return
        control := GuiCtrlFromHwnd(hwnd)
        if !control || !control.HasProp("ToolTip") || control.ToolTip = ""
            return
        this.theme.ShowTooltip(control.ToolTip)
        this.settingsTooltipHideTimer := ObjBindMethod(this, "ClearSettingsTooltip")
        SetTimer(this.settingsTooltipHideTimer, -5000)
    }
    ClearSettingsTooltip() {
        for name in ["settingsTooltipShowTimer", "settingsTooltipHideTimer"]
            if this.HasOwnProp(name)
                SetTimer(this.%name%, 0)
        this.settingsTooltipHwnd := 0
        if this.HasOwnProp("theme")
            this.theme.HideTooltip()
    }

    Action(label, callback, width := 260) {
        button := this.MainFrame.AddButton("x20 y" this.y " w" Min(380, this.contentWidth - 40) " h30", label)
        button.OnEvent("Click", callback)
        this.y += 42
        return button
    }

    Dispose() {
        if this.closed
            return
        this.closed := true
        if this.updateChecker.busy
            this.updateChecker.Finish("Settings closed.")
        for grid in this.grids
            grid.Dispose()
        this.CancelHotkeyRecording()
        this.ClearSettingsTooltip()
        this.StopTestThumbnail()
        for name in ["searchTimer", "resizeTimer", "updateTimer", "Save_Settings_Delay_Timer"]
            if this.HasOwnProp(name)
                SetTimer(this.%name%, 0)
        if this.HasOwnProp("tooltipCallback")
            OnMessage(0x200, this.tooltipCallback, 0)
        if this.HasOwnProp("sizeMoveCallback") {
            OnMessage(0x231, this.sizeMoveCallback, 0)
            OnMessage(0x232, this.sizeMoveCallback, 0)
        }
        if SettingsWorkbench.Active = this
            SettingsWorkbench.Active := 0
        if this.HasOwnProp("helpCondition") {
            HotIf(this.helpCondition)
            try Hotkey("F1", "Off")
            finally HotIf()
        }
        if this.HasOwnProp("settingsSoundPlayer")
            this.settingsSoundPlayer.Dispose()
        if this.HasOwnProp("activeDialog")
            this.activeDialog.Destroy()
        if this.HasOwnProp("closePrompt")
            this.closePrompt.Destroy()
        if this.HasOwnProp("settingsScroller")
            this.settingsScroller.Stop()
        this.S_Gui.Destroy()
        this.theme.Dispose()
        if this.owner.HasOwnProp("settingsWorkbench") && this.owner.settingsWorkbench = this
            this.owner.DeleteProp("settingsWorkbench")
    }
}

class SettingsFonts {
    static Available(value) {
        for name in this.Names()
            if StrLower(name) = StrLower(value)
                return true
        return false
    }
    static Names() {
        static names := []
        if names.Length
            return names
        found := Map()
        found.CaseSense := "Off"
        dc := DllCall("GetDC", "Ptr", 0, "Ptr")
        logfont := Buffer(92, 0)
        NumPut("UChar", 1, logfont, 23) ; DEFAULT_CHARSET
        callback := CallbackCreate(EnumFont, "F", 4)
        try DllCall("EnumFontFamiliesExW", "Ptr", dc, "Ptr", logfont, "Ptr", callback, "Ptr", 0, "UInt", 0)
        finally {
            CallbackFree(callback)
            DllCall("ReleaseDC", "Ptr", 0, "Ptr", dc)
        }
        text := ""
        for name in found
            text .= name "`n"
        for name in StrSplit(RTrim(Sort(text), "`n"), "`n")
            if name != ""
                names.Push(name)
        return names
        EnumFont(font, metric, type, data) {
            name := StrGet(font + 28, 32, "UTF-16")
            if SubStr(name, 1, 1) != "@"
                found[name] := true
            return 1
        }
    }
}

class SettingsLayout {
    static SliderTicks(gui, slider) {
        slider.Opt("+0x1 -0x10")
        ; Ticks belong to the slider's own native/custom paint, not overlapping siblings.
    }
    static TextHeight(control, text, width) {
        rect := Buffer(16, 0)
        NumPut("Int", Round(width * A_ScreenDPI / 96), rect, 8)
        dc := DllCall("GetDC", "Ptr", control.Hwnd, "Ptr")
        font := SendMessage(0x31, 0, 0, control)
        previous := DllCall("SelectObject", "Ptr", dc, "Ptr", font, "Ptr")
        try {
            DllCall("DrawTextW", "Ptr", dc, "WStr", text, "Int", -1, "Ptr", rect, "UInt", 0xC50)
            return Max(18, Ceil(NumGet(rect, 12, "Int") * 96 / A_ScreenDPI) + 2)
        } finally {
            DllCall("SelectObject", "Ptr", dc, "Ptr", previous)
            DllCall("ReleaseDC", "Ptr", control.Hwnd, "Ptr", dc)
        }
    }
}
