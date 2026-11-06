TestRunner.Register("Updates exclude drafts and pre-releases from stable notifications", UpdateFilteringTest)
UpdateFilteringTest() {
    AssertEqual("1.6.0.14", UpdateChecker.ReleaseTag(JSON.Load('{"tag_name":"v1.6.0.14","draft":false,"prerelease":false}')))
    AssertEqual("", UpdateChecker.ReleaseTag(JSON.Load('{"tag_name":"v9.0","draft":false,"prerelease":true}')))
    AssertEqual("", UpdateChecker.ReleaseTag(JSON.Load('{"tag_name":"v9.0","draft":true,"prerelease":false}')))
    AssertEqual("", UpdateChecker.ReleaseTag(JSON.Load('{"tag_name":"v9.0-beta","draft":false,"prerelease":false}')))
    AssertFalse(UpdateChecker.IsNewer("1.6.0.12", "1.6.0.13"))
    AssertFalse(UpdateChecker.IsNewer("1.6.0.13", "1.6.0.13"))
    AssertTrue(UpdateChecker.IsNewer("1.6.0.14", "1.6.0.13"))
    AssertTrue(UpdateChecker.IsNewer("1.10.0", "1.9.0"))
}

class UpdateRequestFixture {
    Status := 200
    done := false
    aborted := false
    ResponseText := '{"tag_name":"v1.6.0.14","draft":false,"prerelease":false}'
    WaitForResponse(timeout) {
        AssertEqual(0, timeout, "Polling must never wait for network I/O")
        return this.done
    }
    Abort() {
        this.aborted := true
    }
}

class UpdateCheckerFixture extends UpdateChecker {
    requests := []
    Request(url) {
        request := UpdateRequestFixture()
        if InStr(url, "per_page")
            request.ResponseText := '[{"tag_name":"v2.0","draft":false,"prerelease":true}]'
        this.requests.Push(request)
        return request
    }
}

TestRunner.Register("Update checks complete asynchronously and cache with manual retry", UpdateLifecycleTest)
UpdateLifecycleTest() {
    checker := UpdateCheckerFixture()
    try {
        checker.Start()
        checker.Start(true)
        AssertTrue(InStr(ProgramLog.Text, "Update check started; manual=0"))
        AssertTrue(InStr(ProgramLog.Text, "already running"))
        AssertEqual(2, checker.requests.Length, "Duplicate requests must be suppressed")
        checker.Poll()
        AssertTrue(checker.busy)
        for request in checker.requests
            request.done := true
        checker.Poll()
        AssertFalse(checker.busy)
        AssertEqual("1.6.0.14", checker.releaseTag)
        AssertEqual("2.0", checker.preReleaseTag)
        AssertTrue(InStr(ProgramLog.Text, "Update response stable; HTTP 200"))
        AssertTrue(InStr(ProgramLog.Text, "stable=1.6.0.14; pre-release=2.0"))
        checker.Start()
        AssertEqual(2, checker.requests.Length)
        checker.Start(true)
        AssertEqual(4, checker.requests.Length)
        checker.startedAt := A_TickCount - 16000
        checker.Poll()
        AssertFalse(checker.busy)
        AssertTrue(checker.requests[3].aborted)
        AssertEqual("Check timed out. Try again.", checker.status)
        AssertTrue(InStr(ProgramLog.Text, "Update check timed out"))
    } finally {
        checker.Finish("Done")
    }
}

TestRunner.Register("Update HTTP errors recover without clearing a known stable update", UpdateFailureTest)
UpdateFailureTest() {
    checker := UpdateCheckerFixture()
    try {
        checker.releaseTag := "1.6.0.14"
        checker.Start()
        checker.requests[1].done := true
        checker.requests[1].Status := 403
        checker.Poll()
        AssertFalse(checker.busy)
        AssertEqual("Could not check. Try again.", checker.status)
        AssertEqual("1.6.0.14", checker.releaseTag)
        AssertTrue(InStr(ProgramLog.Text, "HTTP 403"))
    } finally {
        checker.Finish("Done")
    }
}

TestRunner.Register("Shared sidebar status follows the latest save or asynchronous update action", UpdateSidebarStatusTest)
UpdateSidebarStatusTest() {
    app := SettingsLayoutFixture()
    checker := UpdateCheckerFixture()
    path := A_Temp "\EVE-X-Preview-update-status-" A_TickCount ".json"
    try {
        app.About_Ctrl()
        app.programVersion := "1.6.0.13"
        app.updateChecker := checker
        app.updateUiTimer := ObjBindMethod(app, "RefreshUpdateStatus")
        app.sidebarStatusBlock.GetPos(&x, &y, &width, &height)
        app.RefreshUpdateStatus()
        AssertEqual("", app.updateNotice.Value)
        app.CheckForUpdates()
        AssertEqual("Checking for updates...", app.updateNotice.Value)
        for request in checker.requests
            request.done := true
        checker.Poll()
        app.RefreshUpdateStatus()
        AssertTrue(InStr(app.updateNotice.Value, "Update available: v1.6.0.14"))
        AssertTrue(InStr(app.updateNotice.Value, "See About to install."))
        app.TrackClientPossitions := !app.TrackClientPossitions
        app.SaveJsonToFile(path)
        AssertEqual("Changes saved.`nClose Settings to apply.", app.updateNotice.Value)
        app.RefreshUpdateStatus()
        AssertEqual("Changes saved.`nClose Settings to apply.", app.updateNotice.Value, "A repeated refresh must not restore an older update result.")
        app.TrackClientPossitions := !app.TrackClientPossitions
        app.CheckForUpdates()
        app.SaveJsonToFile(path)
        AssertEqual("Checking for updates...", app.updateNotice.Value, "The delayed save of an earlier edit must not replace a newer update check.")
        checker.requests[3].ResponseText := '{"tag_name":"v1.6.0.13","draft":false,"prerelease":false}'
        for request in checker.requests
            request.done := true
        checker.Poll()
        app.RefreshUpdateStatus()
        AssertEqual("No stable updates found.", app.updateNotice.Value, "Checking again replaces the save message with the result.")
        for release in [app.programVersion, "1.6.0.12", ""] {
            checker.releaseTag := release
            checker.preReleaseTag := "99.0"
            app.RefreshUpdateStatus()
            AssertEqual("No stable updates found.", app.updateNotice.Value, "Pre-releases never trigger stable notifications.")
        }
        for status in ["Could not check. Try again.", "Check timed out. Try again."] {
            checker.status := status
            app.RefreshUpdateStatus()
            AssertTrue(InStr(app.updateNotice.Value, "Update check unavailable."), "A failure must not claim there are no updates.")
        }
        checker.releaseTag := "1.6.0.14"
        app.RefreshUpdateStatus()
        AssertTrue(InStr(app.updateNotice.Value, "Update available: v1.6.0.14"), "A failed retry retains a known available update.")
        app.sidebarStatusBlock.GetPos(&newX, &newY, &newWidth, &newHeight)
        AssertEqual([x, y, width, height], [newX, newY, newWidth, newHeight], "All statuses use the same fixed footer space.")
    } finally {
        SetTimer(app.updateUiTimer, 0)
        checker.Finish("Done")
        app.Dispose()
        for file in [path, path ".save.tmp"]
            if FileExist(file)
                FileDelete(file)
    }
}

TestRunner.Register("Updates select the latest published single-file settings asset without accepting foreign downloads", UpdateSettingsAssetTest)
UpdateSettingsAssetTest() {
    Asset(tag, url := "", draft := false) => Map("tag_name", "v" tag, "draft", draft, "assets", [Map("name", "EVE-X-Preview-Settings.exe", "browser_download_url", url != "" ? url : "https://github.com/khivus/EVE-X-Preview/releases/download/v" tag "/EVE-X-Preview-Settings.exe")])
    releases := [Asset("1.7.0.3"), Asset("9.0",, true), Map("tag_name", "v10.0", "draft", false, "assets", []), Asset("1.6.2"), Asset("11.0", "https://example.com/untrusted.exe")]
    result := UpdateChecker.SettingsAsset(releases)
    AssertEqual("1.7.0.3", result["tag"])
    AssertEqual("https://github.com/khivus/EVE-X-Preview/releases/download/v1.7.0.3/EVE-X-Preview-Settings.exe", result["url"])
    AssertEqual("", UpdateChecker.SettingsAsset([])["url"])
    app := SettingsWorkbenchFixture()
    try {
        app.Navigate("about")
        found := false
        for control in app.MainFrame
            if control.Type = "Button" && control.Text = "Download latest settings editor"
                found := true
        AssertTrue(found, "Internal About must offer the companion download.")
        app.preInstall.GetPos(, &preY)
        app.settingsInstall.GetPos(, &settingsY)
        AssertTrue(settingsY > preY, "Companion download follows pre-release installation.")
        AssertFalse(app.settingsInstall.Enabled, "Downloading waits for a known published asset.")
    } finally app.Dispose()
}

TestRunner.Register("Updates stage all installed components and preserve the optional editor choice", ReleaseInstallerComponentsTest)
ReleaseInstallerComponentsTest() {
    folder := A_Temp "\EveX-update-test-" A_TickCount
    DirCreate(folder)
    main := folder "\EVE-X-Preview.exe", editor := folder "\EVE-X-Preview-Settings.exe"
    status := {UpdateStatus: (*) => 0}
    try {
        FileAppend("old-main", main)
        installer := ReleaseInstallerFixture(main, "1.7.0.4")
        AssertEqual(1, installer.files.Length)
        installer.Install(status)
        AssertEqual("new-EVE-X-Preview.exe", FileRead(main))
        AssertFalse(FileExist(editor), "Optional settings must not be installed implicitly.")
        FileDelete(main)
        FileAppend("installed-main-before-update", main)
        FileAppend("old-editor", editor)
        installer := ReleaseInstallerFixture(main, "1.7.0.4")
        AssertEqual(2, installer.files.Length)
        installer.rejectEditor := true
        failed := false
        try installer.Install(status)
        catch
            failed := true
        AssertTrue(failed)
        AssertEqual("installed-main-before-update", FileRead(main), "Both downloads must finish before either installed file changes.")
        AssertEqual("old-editor", FileRead(editor))
        installer.rejectEditor := false
        lock := DllCall("CreateFileW", "Str", editor, "UInt", 0x80000000, "UInt", 0, "Ptr", 0, "UInt", 3, "UInt", 0, "Ptr", 0, "Ptr")
        try {
            failed := false
            try installer.Install(status)
            catch
                failed := true
            AssertTrue(failed, "A locked companion must fail the transaction.")
            AssertEqual("installed-main-before-update", FileRead(main), "Replacement failure restores the original main executable.")
        } finally DllCall("CloseHandle", "Ptr", lock)
        installer.Install(status)
        AssertEqual("new-EVE-X-Preview-Settings.exe", FileRead(editor))
        AssertFalse(FileExist(main ".update"))
        AssertFalse(FileExist(editor ".previous"))
    } finally DirDelete(folder, true)
}
class ReleaseInstallerFixture extends ReleaseInstaller {
    rejectEditor := false
    DownloadAsset(url, path) {
        SplitPath(RegExReplace(url, "^.*/"), &asset)
        FileAppend("new-" asset, path)
    }
    Verify(path) {
        if this.rejectEditor && InStr(path, "Settings.exe")
            throw Error("Simulated incomplete editor download")
    }
}

TestRunner.Register("Updates download settings into the installation verify completion and launch without a browser", SettingsDirectDownloadTest)
SettingsDirectDownloadTest() {
    folder := A_Temp "\EveX-download-test-" A_TickCount
    DirCreate(folder)
    downloader := SettingsDownloadFixture({settingsBridge: {installation: folder, HasExternal: false}})
    try {
        AssertTrue(SettingsDownload.ShouldSuggest(Map(), downloader.Path, true))
        AssertFalse(SettingsDownload.ShouldSuggest(Map("SettingsDownloadSuggested", 1), downloader.Path, true))
        AssertFalse(SettingsDownload.ShouldSuggest(Map(), downloader.Path, false))
        stream := ComObject("ADODB.Stream")
        try {
            stream.Type := 1, stream.Open(), stream.LoadFromFile(A_AhkPath)
            downloader.request := {Status: 200, ResponseBody: stream.Read(), WaitForResponse: (*) => true, Abort: (*) => 0}
        } finally stream.Close()
        downloader.tag := FileGetVersion(A_AhkPath), downloader.startedAt := A_TickCount
        downloader.stagedReady := false, downloader.closeRequestedAt := 0, downloader.busy := true
        downloader.Poll()
        AssertFalse(downloader.busy)
        AssertTrue(FileExist(downloader.Path))
        AssertEqual(FileGetVersion(A_AhkPath), FileGetVersion(downloader.Path))
        AssertTrue(downloader.Current(downloader.tag))
        AssertFalse(SettingsDownload.ShouldSuggest(Map(), downloader.Path, true))
        Sleep(350)
        AssertEqual(1, downloader.launches)
        AssertFalse(FileExist(downloader.Path ".download"))
    } finally DirDelete(folder, true)
}
class SettingsDownloadFixture extends SettingsDownload {
    launches := 0
    Launch() => this.launches++
}
