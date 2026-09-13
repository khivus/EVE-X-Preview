; WinHTTP runs the network work asynchronously. The timer only polls completion.
class UpdateChecker {
    busy := false
    checkedAt := 0
    releaseTag := ""
    preReleaseTag := ""
    status := "Not checked yet."

    Start(force := false) {
        if this.busy || (!force && this.checkedAt && A_TickCount - this.checkedAt < 300000)
            return
        this.busy := true
        this.status := "Checking GitHub..."
        this.startedAt := A_TickCount
        try {
            this.releaseRequest := this.Request("https://api.github.com/repos/khivus/EVE-X-Preview/releases/latest")
            ; Pre-releases remain an explicit option in About, never a notification.
            this.preRequest := this.Request("https://api.github.com/repos/khivus/EVE-X-Preview/releases?per_page=30")
            this.pollTimer := ObjBindMethod(this, "Poll")
            SetTimer(this.pollTimer, 100)
        } catch {
            this.Finish("Could not check. Try again.")
        }
    }

    Request(url) {
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
                this.Finish("Check timed out. Try again.")
                return
            }
            if !this.releaseRequest.WaitForResponse(0)
                return
            if this.releaseRequest.Status = 404
                tag := ""
            else if this.releaseRequest.Status = 200
                tag := UpdateChecker.ReleaseTag(JSON.Load(this.releaseRequest.ResponseText))
            else
                throw Error("GitHub request failed")
            this.releaseTag := tag
            ; Failure of the optional pre-release lookup must not hide a stable update.
            if !this.preRequest.WaitForResponse(0)
                return
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
            }
            this.Finish("Updates checked.")
        } catch {
            this.Finish("Could not check. Try again.")
        }
    }

    Finish(status) {
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
}
