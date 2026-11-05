; WinHTTP runs the network work asynchronously. The timer only polls completion.
class UpdateChecker {
    busy := false
    checkedAt := 0
    releaseTag := ""
    preReleaseTag := ""
    status := "Not checked yet."
    settingsTag := ""
    settingsUrl := ""

    Start(force := false) {
        if this.busy || (!force && this.checkedAt && A_TickCount - this.checkedAt < 300000) {
            ProgramLog.Add("Update check skipped: " (this.busy ? "already running" : "cached result"))
            return
        }
        ProgramLog.Add("Update check started; manual=" force)
        this.reportedResponses := Map()
        this.busy := true
        this.status := "Checking GitHub..."
        this.startedAt := A_TickCount
        try {
            this.releaseRequest := this.Request("https://api.github.com/repos/khivus/EVE-X-Preview/releases/latest")
            ; Pre-releases remain an explicit option in About, never a notification.
            this.preRequest := this.Request("https://api.github.com/repos/khivus/EVE-X-Preview/releases?per_page=30")
            this.pollTimer := ObjBindMethod(this, "Poll")
            SetTimer(this.pollTimer, 100)
        } catch as err {
            ProgramLog.Error(err, "Starting update check")
            this.Finish("Could not check. Try again.")
        }
    }

    Request(url) {
        ProgramLog.Add("Update request GET " url)
        request := ComObject("WinHttp.WinHttpRequest.5.1")
        request.SetTimeouts(3000, 3000, 5000, 5000)
        request.Open("GET", url, true)
        request.SetRequestHeader("User-Agent", "EVE-X-Preview")
        request.SetRequestHeader("Accept", "application/vnd.github+json")
        request.Send()
        return request
    }

    Poll() {
        try {
            if A_TickCount - this.startedAt > 15000 {
                ProgramLog.Add("Update check timed out after " (A_TickCount - this.startedAt) "ms", "WARN")
                this.Finish("Check timed out. Try again.")
                return
            }
            if !this.releaseRequest.WaitForResponse(0)
                return
            this.LogResponse("stable", this.releaseRequest.Status)
            if this.releaseRequest.Status = 404
                tag := ""
            else if this.releaseRequest.Status = 200
                tag := UpdateChecker.ReleaseTag(JSON.Load(this.releaseRequest.ResponseText))
            else
                throw Error("GitHub stable request failed; HTTP " this.releaseRequest.Status)
            this.releaseTag := tag
            ; Failure of the optional pre-release lookup must not hide a stable update.
            if !this.preRequest.WaitForResponse(0)
                return
            this.LogResponse("pre-release", this.preRequest.Status)
            this.preReleaseTag := ""
            if this.preRequest.Status = 200 {
                releases := JSON.Load(this.preRequest.ResponseText)
                if !(releases is Array)
                    throw Error("Invalid releases response")
                for release in releases {
                    if tag := UpdateChecker.ReleaseTag(release, true) {
                        this.preReleaseTag := tag
                        break
                    }
                }
                if this.releaseRequest.Status = 200
                    releases.Push(JSON.Load(this.releaseRequest.ResponseText))
                companion := UpdateChecker.SettingsAsset(releases)
                this.settingsTag := companion["tag"], this.settingsUrl := companion["url"]
            }
            this.Finish("Updates checked.")
        } catch as err {
            ProgramLog.Error(err, "Polling update check")
            this.Finish("Could not check. Try again.")
        }
    }

    LogResponse(channel, status) {
        if !this.reportedResponses.Has(channel) {
            this.reportedResponses[channel] := true
            ProgramLog.Add("Update response " channel "; HTTP " status, status = 200 || status = 404 ? "INFO" : "WARN")
        }
    }

    Finish(status) {
        ProgramLog.Add("Update check finished: " status "; stable=" this.releaseTag "; pre-release=" this.preReleaseTag)
        if this.HasOwnProp("pollTimer")
            SetTimer(this.pollTimer, 0)
        for name in ["releaseRequest", "preRequest"] {
            if this.HasOwnProp(name) {
                try this.%name%.Abort()
                this.DeleteProp(name)
            }
        }
        this.busy := false
        this.checkedAt := A_TickCount
        this.status := status
    }

    static ReleaseTag(release, preRelease := false) {
        if !(release is Map) || !release.Has("draft") || !release.Has("prerelease") || !release.Has("tag_name")
            throw Error("Invalid release response")
        if release["draft"] || !!release["prerelease"] != preRelease
            return ""
        tag := RegExReplace(release["tag_name"], "i)^v")
        ; Match the numeric version scheme used by the application and updater.
        return RegExMatch(tag, "^\d+(\.\d+){1,3}$") ? tag : ""
    }

    static IsNewer(tag, current) {
        return tag != "" && VerCompare(tag, current) > 0
    }

    static SettingsAsset(releases) {
        result := Map("tag", "", "url", "")
        for release in releases {
            if !(release is Map) || release.Get("draft", true)
                continue
            tag := RegExReplace(release.Get("tag_name", ""), "i)^v")
            if !RegExMatch(tag, "^\d+(\.\d+){1,3}$") || (result["tag"] != "" && VerCompare(tag, result["tag"]) <= 0)
                continue
            for asset in release.Get("assets", []) {
                url := asset.Get("browser_download_url", "")
                if asset.Get("name", "") = "EVE-X-Preview-Settings.exe" && RegExMatch(url, "^https://github\.com/khivus/EVE-X-Preview/releases/download/[^/]+/EVE-X-Preview-Settings\.exe$") {
                    result["tag"] := tag, result["url"] := url
                    break
                }
            }
        }
        return result
    }

    static CurrentVersion() {
        try {
            if A_IsCompiled
                return FileGetVersion(A_ScriptFullPath)
            SplitPath(A_LineFile, , &sourceDirectory)
            if RegExMatch(FileRead(sourceDirectory "\..\Main.ahk"), "U_version = ([\d.]+)", &versionMatch)
                return versionMatch[1]
        }
        return "1.0.0.0"
    }
}
