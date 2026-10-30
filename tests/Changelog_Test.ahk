TestRunner.Register("Changelog shows the full local history in one resizable modeless window", ChangelogWindowTest)
ChangelogWindowTest() {
    expected := FileRead(A_ScriptDir "\..\Changelog.md", "UTF-8")
    AssertEqual(expected, ChangelogWindow.ReadText())
    AssertTrue(RegExMatch(UpdateChecker.CurrentVersion(), "^\d+(\.\d+){1,3}$"))
    heartbeat := {count: 0}
    tick := (*) => heartbeat.count++
    try {
        SetTimer(tick, 10)
        window := ChangelogWindow.Show()
        hwnd := window.Hwnd
        rendered := RegExReplace(expected, "m)^#{1,6} +", "")
        rendered := RegExReplace(rendered, "m)^[ \t]*\* +", Chr(0x2022) "`t")
        rendered := StrReplace(rendered, "``")
        AssertEqual(RTrim(StrReplace(rendered, "`r"), "`n"), RTrim(StrReplace(window["Contents"].Text, "`r"), "`n"), "Every release must be present, including the end of the file.")
        AssertEqual(hwnd, ChangelogWindow.Show().Hwnd, "Repeated actions must reuse the open window.")
        AssertTrue(DllCall("IsWindowEnabled", "Ptr", hwnd))
        Sleep(80)
        AssertTrue(heartbeat.count > 0, "Timers must keep running while the changelog stays open.")
        AssertEqual(0, SendMessage(0xCE, 0, 0, window["Contents"]), "The first release must be visible after the window is shown and laid out.") ; EM_GETFIRSTVISIBLELINE
        SendMessage(0x115, 7, 0, window["Contents"]) ; WM_VSCROLL / SB_BOTTOM
        AssertTrue(SendMessage(0xCE, 0, 0, window["Contents"]) > 0, "Reproduce a reader previously scrolled to the end.")
        ChangelogWindow.Show()
        AssertEqual(0, SendMessage(0xCE, 0, 0, window["Contents"]), "Opening an existing changelog must return to the beginning.")
        ChangelogWindow.Resize(window, 0, 900, 650)
        window["Contents"].GetPos(, , &width, &height)
        AssertEqual(868, width)
        AssertEqual(514, height)
        ChangelogWindow.Show(true)
        AssertTrue(InStr(window["Heading"].Value, "Successfully updated to EVE-X-Preview v" UpdateChecker.CurrentVersion()))
        ChangelogWindow.Close()
        AssertFalse(DllCall("IsWindow", "Ptr", hwnd))
        AssertTrue(IsObject(ChangelogWindow.Show()), "The changelog can be reopened after closing.")
        Sleep(30)
        AssertEqual(0, SendMessage(0xCE, 0, 0, ChangelogWindow.window["Contents"]), "A reopened window must also start at the beginning.")
    } finally {
        SetTimer(tick, 0)
        ChangelogWindow.Close()
    }
}

ChangelogCharacterFormat(control, start, length) {
    selection := Buffer(8, 0)
    NumPut("Int", start, "Int", start + length, selection)
    SendMessage(0x437, 0, selection.Ptr, control) ; EM_EXSETSEL
    format := Buffer(116, 0) ; CHARFORMAT2W
    NumPut("UInt", format.Size, format)
    SendMessage(0x43A, 1, format.Ptr, control) ; EM_GETCHARFORMAT / SCF_SELECTION
    return {size: NumGet(format, 12, "Int"), bold: !!(NumGet(format, 8, "UInt") & 1), font: StrGet(format.Ptr + 26, 32, "UTF-16")}
}

TestRunner.Register("Markdown renders sized headings, bullets, inline formatting and literal Unicode safely", ChangelogMarkdownTest)
ChangelogMarkdownTest() {
    window := Gui()
    try {
        control := SimpleMarkdown.Add(window, "w600 h400")
        SimpleMarkdown.SetText(control, "# Release`n### Added`nBody **bold** and ``code``.`n* Item {safe} \path – Привет 😀`nEnd")
        expected := "Release`nAdded`nBody bold and code.`n" Chr(0x2022) "`tItem {safe} \path – Привет 😀`nEnd"
        text := StrReplace(control.Text, "`r")
        AssertEqual(expected, text, "Only supported Markdown markers should be consumed; Unicode and RTF syntax must remain literal.")
        AssertTrue(ChangelogCharacterFormat(control, 0, 7).bold)
        AssertEqual(400, ChangelogCharacterFormat(control, 0, 7).size)
        AssertEqual(280, ChangelogCharacterFormat(control, 8, 5).size)
        AssertEqual(220, ChangelogCharacterFormat(control, 14, 4).size)
        AssertFalse(ChangelogCharacterFormat(control, 14, 4).bold, "Heading bold must not leak into body text.")
        AssertTrue(ChangelogCharacterFormat(control, 19, 4).bold)
        AssertEqual("Consolas", ChangelogCharacterFormat(control, 28, 4).font)
        AssertTrue(WinGetStyle(control.Hwnd) & 0x800, "The reader must remain read-only.")
        SimpleMarkdown.SetText(control, "# C#`n## Closing heading ##`n``**literal**`` {\rtf1} ")
        AssertEqual("C#`nClosing heading`n**literal** {\rtf1} ", StrReplace(control.Text, "`r"))
    } finally window.Destroy()
}

TestRunner.Register("About changelog button uses the standard adjacent-control gap", ChangelogAboutSpacingTest)
ChangelogAboutSpacingTest() {
    app := ColorSettingsFixture()
    try {
        app.About_Ctrl()
        app.MainFrame["checkUpdatesBtn"].GetPos(&x, &y, &width)
        app.MainFrame["showChangelogBtn"].GetPos(&nextX, &nextY)
        AssertEqual(app.baseGrid, nextX - x - width)
        AssertEqual(y, nextY)
    } finally app.Dispose()
}

class ChangelogStartupFixture extends Main_Class {
    __New() {
        this._JSON := Map("First_Start_After_Update", 1)
        this.saved := 0
        this.Save_Settings_Delay_Timer := (*) => this.saved++
    }
    CleanupUpdater(*) {
        ; Never touch a user's updater from this test.
    }
}

TestRunner.Register("Post-update notice returns immediately and only opens on the first startup", PostUpdateChangelogTest)
PostUpdateChangelogTest() {
    app := ChangelogStartupFixture()
    try {
        started := A_TickCount
        app.ShowPostUpdateChangelog()
        AssertTrue(A_TickCount - started < 1000, "Post-update startup must not wait for updater cleanup or a dialog.")
        AssertEqual(0, app.First_Start_After_Update)
        Sleep(300)
        AssertTrue(app.saved > 0, "The cleared first-start flag must be saved.")
        AssertTrue(IsObject(ChangelogWindow.window))
        AssertTrue(InStr(ChangelogWindow.window["Heading"].Value, "Successfully updated"))
        AssertEqual(0, SendMessage(0xCE, 0, 0, ChangelogWindow.window["Contents"]), "The automatic post-update notice must start at the first release.")
        ChangelogWindow.Close()
        app.ShowPostUpdateChangelog()
        Sleep(30)
        AssertEqual(0, ChangelogWindow.window, "Normal startups must not reopen the notice.")
    } finally {
        SetTimer(app.Save_Settings_Delay_Timer, 0)
        ChangelogWindow.Close()
    }
}
