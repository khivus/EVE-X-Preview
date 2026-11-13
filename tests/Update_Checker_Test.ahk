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

TestRunner.Register("Updates download separate EXEs and preserve the optional editor choice", ReleaseInstallerComponentsTest)
ReleaseInstallerComponentsTest() {
    folder := A_Temp "\EveX-update-test-" A_TickCount
    DirCreate(folder)
    main := folder "\EVE-X-Preview.exe", editor := folder "\EVE-X-Preview-Settings.exe"
    status := {UpdateStatus: (*) => 0}
    version := FileGetVersion(A_AhkPath)
    try {
        probeFixture := folder "\settings-probe.exe"
        helper := A_ScriptDir "\Exe_Update_Fixtures.ps1"
        code := RunUpdaterTestProcess('powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "' helper '" -Mode probe -Path "' probeFixture '" -FixtureVersion "' version '" -BaseExecutable "' A_AhkPath '"', "compile settings probe")
        AssertEqual(0, code)
        FileAppend("old-main", main)
        FileAppend("personal settings", folder "\EVE-X-Preview.json")
        installer := ReleaseInstallerFixture(main, version)
        installer.settingsFixture := probeFixture
        AssertFalse(installer.includeSettings)
        installer.Install(status)
        AssertEqual(version, FileGetVersion(main))
        AssertEqual(["https://github.com/khivus/EVE-X-Preview/releases/download/v" version "/EVE-X-Preview.exe"], installer.downloads)
        AssertFalse(FileExist(editor), "Optional settings must not be installed implicitly.")
        FileDelete(main)
        FileAppend("installed-main-before-update", main)
        FileAppend("old-editor", editor)
        installer := ReleaseInstallerFixture(main, version)
        installer.settingsFixture := probeFixture
        AssertTrue(installer.includeSettings)
        code := RunUpdaterTestProcess('powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "' helper '" -Mode interruption -Path "' main '"', "create interrupted update")
        AssertEqual(0, code)
        installer.failMain := true
        AssertThrows(() => installer.Install(status))
        AssertEqual("installed-main-before-update", FileRead(main), "Pending recovery must finish before a failing download.")
        AssertFalse(DirExist(folder "\.EVE-X-Preview-update"))
        installer.failMain := false
        installer.failSettings := true
        AssertThrows(() => installer.Install(status))
        AssertEqual("installed-main-before-update", FileRead(main), "A failed settings download must leave main unchanged.")
        AssertEqual("old-editor", FileRead(editor))
        installer.failSettings := false
        lock := DllCall("CreateFileW", "Str", editor, "UInt", 0x80000000, "UInt", 1, "Ptr", 0, "UInt", 3, "UInt", 0, "Ptr", 0, "Ptr")
        try {
            failed := false
            try installer.Install(status)
            catch
                failed := true
            AssertTrue(failed, "A locked companion must fail the transaction.")
            AssertEqual("installed-main-before-update", FileRead(main), "Replacement failure restores the original main executable.")
        } finally DllCall("CloseHandle", "Ptr", lock)
        installer.downloads := []
        installer.Install(status)
        AssertEqual(version, FileGetVersion(editor))
        AssertEqual(["https://github.com/khivus/EVE-X-Preview/releases/download/v" version "/EVE-X-Preview.exe",
            "https://github.com/khivus/EVE-X-Preview/releases/download/v" version "/EVE-X-Preview-Settings.exe"], installer.downloads)
        AssertEqual("personal settings", FileRead(folder "\EVE-X-Preview.json"))
        AssertFalse(FileExist(folder "\EVE-X-Preview.package.json"))
        AssertFalse(DirExist(folder "\.EVE-X-Preview-update"))
    } finally DirDelete(folder, true)
}
class ReleaseInstallerFixture extends ReleaseInstaller {
    downloads := []
    failSettings := false
    failMain := false
    settingsFixture := ""
    RunTransaction(engine, request) {
        SplitPath(request, &phase)
        return RunUpdaterTestProcess('"' A_WinDir '\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "' engine '" -RequestPath "' request '"', "transaction " phase)
    }
    DownloadAsset(url, path) {
        if !RegExMatch(url, "^https://github\.com/khivus/EVE-X-Preview/releases/download/v[\d.]+/EVE-X-Preview(?:-Settings)?\.exe$")
            throw Error("Updater must request separate EXEs, without archive or sidecar assets: " url)
        this.downloads.Push(url)
        if this.failMain
            throw Error("Simulated network failure")
        if this.failSettings && InStr(url, "-Settings.exe")
            throw Error("Simulated settings download failure")
        FileCopy(InStr(url, "-Settings.exe") ? this.settingsFixture : A_AhkPath, path)
    }
}

RunUpdaterTestProcess(command, phase, timeoutMs := 120000) {
    TestRunner.Log("RUN updater: " phase)
    started := A_TickCount
    Run(command,, "Hide", &pid)
    process := DllCall("OpenProcess", "UInt", 0x100000 | 0x1000 | 1, "Int", false, "UInt", pid, "Ptr")
    if !process {
        failure := OSError()
        if ProcessExist(pid) {
            ProcessClose(pid)
            if ProcessWaitClose(pid, 5)
                throw Error("Could not stop updater test phase after handle failure: " phase)
        }
        throw failure
    }
    try {
        loop {
            state := DllCall("WaitForSingleObject", "Ptr", process, "UInt", 0, "UInt")
            if state = 0
                break
            if state != 258
                throw OSError()
            if A_TickCount - started >= timeoutMs
                throw Error("Updater test phase timed out: " phase " (" timeoutMs " ms)")
            Sleep(25) ; Keep the test interpreter responsive while the child runs.
        }
        code := 0
        if !DllCall("GetExitCodeProcess", "Ptr", process, "UInt*", &code)
            throw OSError()
        TestRunner.Log("DONE updater: " phase "; exit=" code "; elapsed=" (A_TickCount - started) " ms")
        return code
    } finally {
        try {
            if DllCall("WaitForSingleObject", "Ptr", process, "UInt", 0, "UInt") = 258 {
                if !DllCall("TerminateProcess", "Ptr", process, "UInt", 1)
                    throw OSError()
                if DllCall("WaitForSingleObject", "Ptr", process, "UInt", 5000, "UInt") != 0
                    throw Error("Could not stop updater test phase: " phase)
            }
        } finally DllCall("CloseHandle", "Ptr", process)
    }
}

TestRunner.Register("Updater subprocess waits preserve exit codes and terminate a timed-out fixture", UpdaterProcessTimeoutTest)
UpdaterProcessTimeoutTest() {
    folder := A_Temp "\EveX-updater-timeout-" A_TickCount
    DirCreate(folder)
    script := folder "\child.ahk", marker := folder "\pid.txt"
    FileAppend('#Requires AutoHotkey v2.0`n#NoTrayIcon`nFileAppend(ProcessExist(), A_Args[1])`nSleep(A_Args[2])`nExitApp(7)', script, "UTF-8")
    command := '"' A_AhkPath '" /ErrorStdOut "' script '" "' marker '" '
    try {
        AssertEqual(7, RunUpdaterTestProcess(command "100", "exit-code fixture"))
        FileDelete(marker)
        AssertThrows(() => RunUpdaterTestProcess(command "60000", "timeout fixture", 2000), "Updater test phase timed out: timeout fixture")
        AssertTrue(FileExist(marker), "The fixture must start before the timeout assertion.")
        AssertFalse(ProcessExist(FileRead(marker) + 0), "The timed-out child must be terminated.")
    } finally DirDelete(folder, true)
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
