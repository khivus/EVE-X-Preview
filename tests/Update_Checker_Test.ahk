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
        AssertEqual(2, checker.requests.Length, "Duplicate requests must be suppressed")
        checker.Poll()
        AssertTrue(checker.busy)
        for request in checker.requests
            request.done := true
        checker.Poll()
        AssertFalse(checker.busy)
        AssertEqual("1.6.0.14", checker.releaseTag)
        AssertEqual("2.0", checker.preReleaseTag)
        checker.Start()
        AssertEqual(2, checker.requests.Length)
        checker.Start(true)
        AssertEqual(4, checker.requests.Length)
        checker.startedAt := A_TickCount - 16000
        checker.Poll()
        AssertFalse(checker.busy)
        AssertTrue(checker.requests[3].aborted)
        AssertEqual("Check timed out. Try again.", checker.status)
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
    } finally {
        checker.Finish("Done")
    }
}
